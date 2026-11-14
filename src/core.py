"""
Busca Databook - motor de leitura, indice e busca.

Le o conteudo dos documentos de uma pasta (e subpastas), guarda o texto em
um indice SQLite FTS5 e faz buscas rapidas, sem diferenciar maiusculas,
minusculas ou acentos ("trafo" encontra "TRAFO", "Trafô", "trafos"...).
"""

import hashlib
import hmac
import html
import io
import json
import logging
import os
import re
import secrets
import shutil
import sqlite3
import struct
import sys
import threading
import time
import unicodedata
import zipfile
from pathlib import Path

log = logging.getLogger(__name__)
# Registro de auditoria (LGPD, art. 37): quem fez o que e quando. Nunca grava o texto dos
# documentos nem os termos pesquisados. O app configura o arquivo (auditoria.log).
log_auditoria = logging.getLogger("buscadatabook.auditoria")
log_auditoria.propagate = False


def auditar(usuario, evento, detalhe=""):
    """Grava um evento no registro de auditoria (uma linha, sem quebras)."""
    limpo = lambda t: re.sub(r"[\r\n|]+", " ", str(t or "")).strip()[:500]
    log_auditoria.info("%s | %s | %s", limpo(usuario) or "-", limpo(evento), limpo(detalhe))

# --------------------------------------------------------------------------
# Configuracao
# --------------------------------------------------------------------------

# Formatos cujo CONTEUDO e lido
EXT_CONTEUDO = {
    ".docx", ".docm", ".doc",
    ".xlsx", ".xlsm", ".xls",
    ".pptx", ".pptm",
    ".pdf",
    ".txt", ".csv", ".log", ".xml", ".json", ".md", ".htm", ".html", ".rtf",
    ".zip",
}
# Todos os demais (png, jpg, mpp, dwg, dwf, ...) sao buscados so pelo nome.

MAX_TAMANHO_MB = 300          # arquivos maiores sao indexados so pelo nome
MAX_PROFUNDIDADE_ZIP = 3      # zip dentro de zip dentro de zip...
OCR_DPI = 200
OCR_MAX_PAGINAS = 200         # limite de paginas escaneadas por PDF
SEP_ZIP = " >> "              # separador entre o zip e o arquivo interno

PREFIXO_EMPRESA = "DB_"       # pastas "DB_NOME" na pasta principal viram empresas
SENHA_MIN = 8                 # senhas novas; senhas antigas fora da regra sao trocadas no proximo login
REGRAS_SENHA = (f"Mínimo de {SENHA_MIN} caracteres, com letra maiúscula,\n"
                "letra minúscula, número e caractere especial (ex.: ! @ # $ % & *).")


def problemas_senha(senha, login=""):
    """Lista o que falta para a senha atender a regra (lista vazia = senha aceita)."""
    faltam = []
    if len(senha) < SENHA_MIN:
        faltam.append(f"ter pelo menos {SENHA_MIN} caracteres")
    if not any(c.isupper() for c in senha):
        faltam.append("ter uma letra maiúscula")
    if not any(c.islower() for c in senha):
        faltam.append("ter uma letra minúscula")
    if not any(c.isdigit() for c in senha):
        faltam.append("ter um número")
    if not any(not c.isalnum() and not c.isspace() for c in senha):
        faltam.append("ter um caractere especial (ex.: ! @ # $ % & *)")
    if login and len(login.strip()) >= 3 and login.strip().lower() in senha.lower():
        faltam.append("não conter o nome de usuário")
    return faltam


def senha_forte(senha, login=""):
    return not problemas_senha(senha, login)

FORMATO_EXPORTACAO = "buscadatabook-indice/1"
EXT_EXPORTACAO = ".bdx"
FORMATO_PACOTE = "buscadatabook-pacote/1"
EXT_PACOTE = ".bdpack"

# Atributos de arquivo do Windows usados pelo OneDrive (Arquivos sob demanda)
_ATR_OFFLINE = 0x1000
_ATR_RECALL_ON_OPEN = 0x40000
_ATR_PINNED = 0x80000                  # "Sempre manter neste dispositivo"
_ATR_UNPINNED = 0x100000               # "Liberar espaco"
_ATR_RECALL_ON_DATA_ACCESS = 0x400000
_ATR_NUVEM = _ATR_OFFLINE | _ATR_RECALL_ON_OPEN | _ATR_RECALL_ON_DATA_ACCESS
_ATR_GRAVAVEIS = 0x1 | 0x2 | 0x4 | 0x20 | 0x100 | 0x2000   # somente leitura, oculto, sistema...


def caminho_longo(p) -> str:
    """
    Caminho que o Windows aceita mesmo com mais de 260 caracteres (prefixo \\\\?\\).
    Databooks em rede costumam passar desse limite; sem isso, o arquivo e listado mas
    "nao encontrado" ao abrir. Fora do Windows, devolve o caminho como esta.
    """
    p = str(p)
    if os.name != "nt" or p.startswith("\\\\?\\"):
        return p
    p = os.path.abspath(p)
    if p.startswith("\\\\"):                       # \\servidor\pasta -> \\?\UNC\servidor\pasta
        return "\\\\?\\UNC\\" + p[2:]
    return "\\\\?\\" + p


def caminho_para_abrir(p) -> str:
    """Caminho para entregar ao Windows ao abrir um arquivo com o programa padrao.
    Acima de ~250 caracteres, usa o nome curto (8.3) quando o disco oferece."""
    p = str(p)
    if os.name != "nt" or len(p) < 250:
        return p
    try:
        import ctypes
        buf = ctypes.create_unicode_buffer(32768)
        if ctypes.windll.kernel32.GetShortPathNameW(caminho_longo(p), buf, len(buf)):
            curto = buf.value
            if curto.startswith("\\\\?\\UNC\\"):
                curto = "\\\\" + curto[8:]
            elif curto.startswith("\\\\?\\"):
                curto = curto[4:]
            return curto
    except Exception:
        log.debug("Nome curto indisponível para %s", p, exc_info=True)
    return caminho_longo(p)


def pasta_dados_app() -> Path:
    """Dados do usuario (indices, usuarios, config, logs) - separados do programa."""
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    p = Path(base) / "BuscaDatabook"
    p.mkdir(parents=True, exist_ok=True)
    return p


# --------------------------------------------------------------------------
# OCR (PDFs escaneados) - usa o Tesseract embutido no instalador
# --------------------------------------------------------------------------

_ocr_estado = {"ok": None, "idioma": "eng", "exe": None}
_ocr_lock = threading.Lock()


def pasta_tesseract_embutido() -> Path:
    if getattr(sys, "frozen", False):                       # programa instalado
        return Path(sys.executable).resolve().parent / "tesseract"
    return Path(__file__).resolve().parent.parent / "vendor" / "tesseract"  # codigo-fonte


def configurar_ocr():
    """Detecta o Tesseract (primeiro o embutido). Retorna True se o OCR estiver disponivel."""
    with _ocr_lock:
        if _ocr_estado["ok"] is not None:
            return _ocr_estado["ok"]
        try:
            import pytesseract
            embutido = pasta_tesseract_embutido()
            exe_embutido = embutido / "tesseract.exe"
            candidatos = [
                exe_embutido,
                Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
                Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
                Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Tesseract-OCR" / "tesseract.exe",
            ]
            for c in candidatos:
                if c.is_file():
                    pytesseract.pytesseract.tesseract_cmd = str(c)
                    if c == exe_embutido:
                        os.environ["TESSDATA_PREFIX"] = str(embutido / "tessdata")
                    _ocr_estado["exe"] = str(c)
                    break
            idiomas = pytesseract.get_languages(config="")
            _ocr_estado["idioma"] = "por+eng" if "por" in idiomas else "eng"
            _ocr_estado["ok"] = True
            log.info("OCR disponível: %s (%s)", _ocr_estado["exe"] or "PATH", _ocr_estado["idioma"])
        except Exception:
            log.warning("OCR indisponível", exc_info=True)
            _ocr_estado["ok"] = False
        return _ocr_estado["ok"]


def _ocr_imagem(img) -> str:
    import pytesseract
    return pytesseract.image_to_string(img, lang=_ocr_estado["idioma"])


# --------------------------------------------------------------------------
# Extratores de texto (recebem bytes ou caminho; devolvem str)
# --------------------------------------------------------------------------

def _abrir(fonte):
    """
    Entrega o CONTEUDO do arquivo as bibliotecas (nao o nome): assim um .xls que na
    verdade e .xlsx e aceito, e caminhos longos do Windows funcionam.
    """
    if isinstance(fonte, (bytes, bytearray)):
        return io.BytesIO(fonte)
    with open(caminho_longo(fonte), "rb") as f:
        return io.BytesIO(f.read())


def ler_docx(fonte) -> str:
    import docx
    d = docx.Document(_abrir(fonte))
    partes = [p.text for p in d.paragraphs]
    for t in d.tables:
        for linha in t.rows:
            for cel in linha.cells:
                partes.append(cel.text)
    for s in d.sections:
        for hf in (s.header, s.footer):
            try:
                partes.extend(p.text for p in hf.paragraphs)
            except Exception:
                pass
    return "\n".join(partes)


def ler_xlsx(fonte) -> str:
    import openpyxl
    wb = openpyxl.load_workbook(_abrir(fonte), read_only=True, data_only=True)
    partes = []
    try:
        for ws in wb.worksheets:
            partes.append(f"[Aba: {ws.title}]")
            for linha in ws.iter_rows(values_only=True):
                vals = [str(v) for v in linha if v is not None and str(v).strip()]
                if vals:
                    partes.append(" | ".join(vals))
    finally:
        wb.close()
    return "\n".join(partes)


def ler_xls(fonte) -> str:
    try:
        import xlrd
    except ImportError:
        return ler_binario_legado(fonte)
    wb = xlrd.open_workbook(file_contents=_ler_bytes(fonte))
    partes = []
    for sh in wb.sheets():
        partes.append(f"[Aba: {sh.name}]")
        for r in range(sh.nrows):
            vals = [str(v) for v in sh.row_values(r) if str(v).strip()]
            if vals:
                partes.append(" | ".join(vals))
    return "\n".join(partes)


def ler_pptx(fonte) -> str:
    from pptx import Presentation
    pres = Presentation(_abrir(fonte))
    partes = []
    for i, slide in enumerate(pres.slides, 1):
        partes.append(f"[Slide {i}]")
        for shape in slide.shapes:
            if shape.has_text_frame:
                partes.append(shape.text_frame.text)
            if getattr(shape, "has_table", False) and shape.has_table:
                for linha in shape.table.rows:
                    for cel in linha.cells:
                        partes.append(cel.text)
        if slide.has_notes_slide:
            partes.append(slide.notes_slide.notes_text_frame.text)
    return "\n".join(partes)


ASSINATURAS_IMAGEM = (b"\xff\xd8\xff", b"\x89PNG", b"II*\x00", b"MM\x00*", b"BM", b"GIF8")


def _ler_bytes(fonte) -> bytes:
    if isinstance(fonte, (bytes, bytearray)):
        return bytes(fonte)
    with open(caminho_longo(fonte), "rb") as f:   # Python le caminhos longos; o PDFium nao
        return f.read()


def ler_imagem(fonte, usar_ocr=True, progresso=None) -> str:
    """Imagem (inclusive salva com extensao .pdf): texto so por OCR."""
    if not (usar_ocr and configurar_ocr()):
        return ""
    from PIL import Image, ImageSequence
    partes = []
    with Image.open(io.BytesIO(_ler_bytes(fonte))) as img:
        for i, quadro in enumerate(ImageSequence.Iterator(img)):     # TIFF com varias paginas
            if i >= OCR_MAX_PAGINAS:
                break
            if progresso:
                progresso(f"OCR imagem {i + 1}")
            partes.append(_ocr_imagem(quadro.convert("RGB")))
    return "\n".join(partes)


def ler_pdf(fonte, usar_ocr=True, progresso=None) -> str:
    """
    Le o texto do PDF. Paginas sem texto (escaneadas) passam por OCR.
    Tolerante a defeitos comuns: lixo antes do cabecalho %PDF, imagem salva como .pdf,
    paginas que nao abrem ou OCR que falha numa pagina (pula so aquela pagina).
    """
    import pypdfium2 as pdfium
    dados = _ler_bytes(fonte)
    if not dados.startswith(b"%PDF"):
        if dados.startswith(ASSINATURAS_IMAGEM):
            return ler_imagem(dados, usar_ocr, progresso)
        inicio = dados.find(b"%PDF", 0, 1024 * 1024)
        if inicio > 0:
            dados = dados[inicio:]
        elif not dados.strip():
            raise ValueError("arquivo vazio")
    pdf = pdfium.PdfDocument(dados)
    partes = []
    ocr_feitas = 0
    try:
        for i in range(len(pdf)):
            try:
                pag = pdf[i]
            except Exception as e:
                log.info("Página %d ignorada: %s", i + 1, e)
                partes.append(f"[Página {i + 1}]")
                continue
            try:
                texto = ""
                try:
                    tp = pag.get_textpage()
                    texto = tp.get_text_range() or ""
                    tp.close()
                except Exception:
                    pass
                if len(texto.strip()) < 20 and usar_ocr and ocr_feitas < OCR_MAX_PAGINAS and configurar_ocr():
                    if progresso:
                        progresso(f"OCR página {i + 1}")
                    try:
                        img = pag.render(scale=OCR_DPI / 72).to_pil()
                        texto = _ocr_imagem(img)
                    except Exception as e:
                        log.info("OCR falhou na página %d: %s", i + 1, e)
                    ocr_feitas += 1
            finally:
                pag.close()
            partes.append(f"[Página {i + 1}]\n{texto}")
    finally:
        pdf.close()
    return "\n".join(partes)


def ler_texto(fonte) -> str:
    dados = _ler_bytes(fonte)
    txt = ""
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            txt = dados.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if txt.lstrip().startswith("{\\rtf"):
        txt = re.sub(r"\\[a-z]+-?\d* ?|[{}]", " ", txt)
    elif "<" in txt[:2000] and ">" in txt[:2000]:
        txt = re.sub(r"<[^>]+>", " ", txt)
    return txt


def ler_doc(fonte) -> str:
    """.doc antigo (Word 97-2003): tenta o Word via COM; senao, leitura bruta."""
    if not isinstance(fonte, (bytes, bytearray)) and os.name == "nt":
        try:
            import win32com.client  # pywin32 (opcional)
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            try:
                alvo = str(fonte)
                if alvo.startswith("\\\\?\\") and len(alvo) < 264:   # Word nao aceita o prefixo \\?\\
                    alvo = alvo[4:]
                d = word.Documents.Open(alvo, ReadOnly=True, AddToRecentFiles=False)
                txt = d.Content.Text
                d.Close(False)
                return txt
            finally:
                word.Quit()
        except Exception:
            pass
    return ler_binario_legado(fonte)


def ler_binario_legado(fonte) -> str:
    """Extrai trechos de texto legiveis de arquivos binarios antigos (.doc/.xls)."""
    dados = _ler_bytes(fonte)
    trechos = []
    for m in re.finditer(rb"(?:[\x20-\x7e\xc0-\xff]\x00){4,}", dados):   # UTF-16LE
        trechos.append(m.group().decode("utf-16le", "ignore"))
    for m in re.finditer(rb"[\x20-\x7e\xc0-\xff]{4,}", dados):            # cp1252
        trechos.append(m.group().decode("cp1252", "ignore"))
    return "\n".join(trechos)


LEITORES = {
    ".docx": ler_docx, ".docm": ler_docx,
    ".xlsx": ler_xlsx, ".xlsm": ler_xlsx,
    ".xls": ler_xls,
    ".doc": ler_doc,
    ".pptx": ler_pptx, ".pptm": ler_pptx,
    ".pdf": ler_pdf,
}


def ler_ooxml_bruto(fonte) -> str:
    """
    Reserva para Word/Excel/PowerPoint (docx/xlsx/pptx) que a biblioteca normal recusa
    (arquivo com pequenos defeitos ou recursos incomuns): le o texto direto dos XML do
    pacote, ignorando as partes danificadas.
    """
    dados = _ler_bytes(fonte)
    partes = []
    with zipfile.ZipFile(io.BytesIO(dados)) as z:
        for info in z.infolist():
            n = info.filename.lower()
            if not n.endswith(".xml") or not n.startswith(("word/", "xl/", "ppt/")):
                continue
            if "/theme" in n or "/styles" in n or "/_rels" in n or n.endswith("settings.xml"):
                continue
            try:
                xml = z.read(info).decode("utf-8", "ignore")
            except Exception:
                continue                       # parte danificada: segue com as demais
            texto = html.unescape(re.sub(r"<[^>]+>", " ", xml))
            texto = re.sub(r"\s+", " ", texto).strip()
            if texto:
                partes.append(texto)
    if not partes:
        raise ValueError("nenhum texto legível no pacote do Office")
    return "\n".join(partes)


def _cabecalho(fonte, n=8) -> bytes:
    if isinstance(fonte, (bytes, bytearray)):
        return bytes(fonte[:n])
    with open(caminho_longo(fonte), "rb") as f:
        return f.read(n)


def _formato_real(ext, cab):
    """Corrige a extensao pelo conteudo: ex. .xls que na verdade e .xlsx, .docx que e .doc."""
    zip_ = cab.startswith(b"PK\x03\x04")
    ole = cab.startswith(b"\xd0\xcf\x11\xe0")
    familias = ((".xls", ".xlsx", ".xlsm"), (".doc", ".docx", ".docm"), (".pptx", ".pptm"))
    for fam in familias:
        if ext in fam:
            if zip_:
                return {".xls": ".xlsx", ".doc": ".docx"}.get(ext, ext)
            if ole:
                return {".xlsx": ".xls", ".xlsm": ".xls", ".docx": ".doc", ".docm": ".doc"}.get(ext, ext)
    if ext == ".pdf" and cab.startswith(ASSINATURAS_IMAGEM):
        return ".imagem"
    return ext


def extrair(nome: str, fonte, usar_ocr=True, progresso=None, profundidade=0):
    """
    Devolve lista de tuplas (subcaminho_no_zip_ou_vazio, texto).
    Para arquivos normais e uma lista com um item; para zip, um por membro.
    """
    ext = os.path.splitext(nome)[1].lower()
    if ext not in EXT_CONTEUDO:
        return [("", "")]
    ext = _formato_real(ext, _cabecalho(fonte))
    if ext == ".zip":
        return _extrair_zip(fonte, usar_ocr, progresso, profundidade)
    if ext == ".pdf":
        return [("", ler_pdf(fonte, usar_ocr, progresso))]
    if ext == ".imagem":
        return [("", ler_imagem(fonte, usar_ocr, progresso))]
    leitor = LEITORES.get(ext, ler_texto)
    try:
        return [("", leitor(fonte))]
    except Exception as erro:
        # reservas: Office moderno danificado -> XML bruto; Excel/Word antigo -> texto bruto
        if ext in (".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm"):
            try:
                return [("", ler_ooxml_bruto(fonte))]
            except Exception:
                raise erro
        if ext in (".xls", ".doc"):
            texto = ler_binario_legado(_ler_bytes(fonte))
            if texto.strip():
                return [("", texto)]
        raise


def _extrair_zip(fonte, usar_ocr, progresso, profundidade):
    saida = []
    if profundidade >= MAX_PROFUNDIDADE_ZIP:
        return [("", "")]
    with zipfile.ZipFile(_abrir(fonte)) as z:
        for info in z.infolist():
            if info.is_dir():
                continue
            nome_int = _corrigir_nome_zip(info)
            ext = os.path.splitext(nome_int)[1].lower()
            texto = ""
            if ext in EXT_CONTEUDO and info.file_size <= MAX_TAMANHO_MB * 1024 * 1024:
                try:
                    dados = z.read(info)
                    sub = extrair(nome_int, dados, usar_ocr, progresso, profundidade + 1)
                    for subcam, t in sub:
                        caminho = nome_int + (SEP_ZIP + subcam if subcam else "")
                        saida.append((caminho, caminho + "\n" + t))
                    continue
                except Exception as e:
                    texto = f"(erro ao ler: {e})"
            # o nome do arquivo interno tambem e pesquisavel
            saida.append((nome_int, nome_int + "\n" + texto))
    return saida or [("", "")]


def _corrigir_nome_zip(info):
    """Zips criados no Windows costumam ter nomes em cp437/cp850."""
    nome = info.filename
    if not (info.flag_bits & 0x800):
        try:
            nome = nome.encode("cp437").decode("cp850")
        except Exception:
            pass
    return nome


# --------------------------------------------------------------------------
# OneDrive: economia de espaco
# --------------------------------------------------------------------------

def so_na_nuvem(st) -> bool:
    """True se o arquivo esta so na nuvem (OneDrive sob demanda): le-lo dispara o download."""
    atr = getattr(st, "st_file_attributes", 0)
    return bool(atr & _ATR_NUVEM) and not (atr & _ATR_PINNED)


def liberar_espaco(caminho) -> bool:
    """Devolve o arquivo a nuvem, como o "Liberar espaco" do Explorador (attrib +U -P)."""
    if os.name != "nt":
        return False
    try:
        import ctypes
        k32 = ctypes.windll.kernel32
        k32.GetFileAttributesW.argtypes = [ctypes.c_wchar_p]
        k32.GetFileAttributesW.restype = ctypes.c_uint32
        k32.SetFileAttributesW.argtypes = [ctypes.c_wchar_p, ctypes.c_uint32]
        k32.SetFileAttributesW.restype = ctypes.c_int
        atr = k32.GetFileAttributesW(str(caminho))
        if atr == 0xFFFFFFFF:
            return False
        return bool(k32.SetFileAttributesW(str(caminho), (atr & _ATR_GRAVAVEIS) | _ATR_UNPINNED))
    except Exception:
        log.debug("Não foi possível liberar espaço de %s", caminho, exc_info=True)
        return False


# --------------------------------------------------------------------------
# Empresas (pastas DB_NOME)
# --------------------------------------------------------------------------

def listar_empresas(raiz: str):
    """
    Pastas logo abaixo da raiz cujo nome comeca com "DB_".
    Devolve lista ordenada de (nome_exibido, caminho_completo). "DB_ENEL" -> "ENEL".
    """
    empresas = []
    try:
        with os.scandir(raiz) as it:
            for e in it:
                try:
                    if not e.is_dir():
                        continue
                except OSError:
                    continue
                if e.name.upper().startswith(PREFIXO_EMPRESA) and len(e.name) > len(PREFIXO_EMPRESA):
                    nome = e.name[len(PREFIXO_EMPRESA):].strip()
                    if nome:
                        empresas.append((nome, e.path))
    except OSError:
        log.warning("Não foi possível listar empresas em %s", raiz, exc_info=True)
    empresas.sort(key=lambda x: x[0].casefold())
    return empresas


# --------------------------------------------------------------------------
# Criptografia dos arquivos que saem do computador (.bdx e indice do .bdpack)
# --------------------------------------------------------------------------
#
# Formato (AES-256-GCM em blocos, chave derivada da senha com scrypt):
#   cabecalho: MAGICA(8) | versao(1) | logN(1) | r(1) | p(1) | sal(16) | prefixo_nonce(7)
#   registros: tipo(1) | tamanho(4) | dados cifrados (com etiqueta de 16 bytes)
#     tipo 0 = metadados (JSON) - sempre o primeiro; permite conferir a senha na hora
#     tipo 1 = bloco de dados;  tipo 2 = ultimo bloco (protege contra arquivo truncado)
#   nonce = prefixo(7) | contador(4) | tipo(1);  o cabecalho inteiro e dado autenticado.

MAGICA_CIFRA = b"BDCRYPT\x01"
SENHA_ARQUIVO_MIN = 12
_BLOCO_CIFRA = 4 << 20
_CAB_CIFRA = struct.Struct(">8sBBBB16s7s")


def esta_cifrado(arquivo) -> bool:
    try:
        with open(arquivo, "rb") as f:
            return f.read(8) == MAGICA_CIFRA
    except OSError:
        return False


def checar_senha_arquivo(senha):
    if len(senha or "") < SENHA_ARQUIVO_MIN:
        raise ValueError(f"A senha do arquivo precisa ter pelo menos {SENHA_ARQUIVO_MIN} caracteres "
                         "(use “Gerar senha” para criar uma senha forte).")


def gerar_senha_arquivo():
    """Senha aleatoria forte e facil de ditar: XXXXX-XXXXX-XXXXX-XXXXX (~100 bits)."""
    bruto = "".join(secrets.choice(ALFABETO_CODIGO) for _ in range(20))
    return "-".join(bruto[i:i + 5] for i in range(0, 20, 5))


def _chave_arquivo(senha, sal, log_n, r, p):
    return hashlib.scrypt(senha.encode("utf-8"), salt=sal, n=1 << log_n, r=r, p=p,
                          maxmem=1 << 28, dklen=32)


def cifrar_arquivo(origem, destino, senha, meta=None):
    """Cifra o arquivo `origem` em `destino` (gravacao atomica). meta: dict curto, tambem cifrado."""
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    checar_senha_arquivo(senha)
    sal, prefixo, log_n, r, p = os.urandom(16), os.urandom(7), 16, 8, 1
    cab = _CAB_CIFRA.pack(MAGICA_CIFRA, 1, log_n, r, p, sal, prefixo)
    aes = AESGCM(_chave_arquivo(senha, sal, log_n, r, p))
    destino = Path(destino)
    parcial = destino.with_name(destino.name + ".tmp")
    cont = 0

    def gravar(f, tipo, dados):
        nonlocal cont
        ct = aes.encrypt(prefixo + cont.to_bytes(4, "big") + bytes([tipo]), dados, cab)
        f.write(bytes([tipo]) + len(ct).to_bytes(4, "big") + ct)
        cont += 1

    try:
        with open(origem, "rb") as src, open(parcial, "wb") as f:
            f.write(cab)
            gravar(f, 0, json.dumps(meta or {}, ensure_ascii=False).encode("utf-8"))
            atual = src.read(_BLOCO_CIFRA)
            while True:
                prox = src.read(_BLOCO_CIFRA)
                gravar(f, 1 if prox else 2, atual)
                if not prox:
                    break
                atual = prox
        os.replace(parcial, destino)
    except Exception:
        try:
            parcial.unlink()
        except OSError:
            pass
        raise


class _LeitorCifrado:
    """Le um arquivo cifrado registro a registro (aceita qualquer objeto de arquivo binario)."""

    def __init__(self, f, senha):
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        self.f = f
        cab = f.read(_CAB_CIFRA.size)
        if len(cab) != _CAB_CIFRA.size or cab[:8] != MAGICA_CIFRA:
            raise ValueError("Arquivo cifrado inválido.")
        _, versao, log_n, r, p, sal, self.prefixo = _CAB_CIFRA.unpack(cab)
        if versao != 1 or not (14 <= log_n <= 20) or not (1 <= r <= 16) or not (1 <= p <= 4):
            raise ValueError("Arquivo cifrado em formato desconhecido (atualize o programa).")
        self.cab = cab
        self.aes = AESGCM(_chave_arquivo(senha or "", sal, log_n, r, p))
        self.cont = 0

    def proximo(self):
        """(tipo, dados) do proximo registro, ou None no fim do arquivo."""
        from cryptography.exceptions import InvalidTag
        t = self.f.read(5)
        if not t:
            return None
        if len(t) != 5:
            raise ValueError("Arquivo cifrado incompleto ou danificado.")
        tipo, tam = t[0], int.from_bytes(t[1:], "big")
        if tipo > 2 or tam > _BLOCO_CIFRA + 16:
            raise ValueError("Arquivo cifrado danificado.")
        ct = self.f.read(tam)
        if len(ct) != tam:
            raise ValueError("Arquivo cifrado incompleto ou danificado.")
        try:
            dados = self.aes.decrypt(self.prefixo + self.cont.to_bytes(4, "big") + bytes([tipo]), ct, self.cab)
        except InvalidTag:
            if self.cont == 0:
                raise ValueError("Senha incorreta.")
            raise ValueError("Arquivo cifrado danificado ou alterado.")
        self.cont += 1
        return tipo, dados

    def meta(self):
        r = self.proximo()
        if not r or r[0] != 0:
            raise ValueError("Arquivo cifrado danificado.")
        try:
            return json.loads(r[1].decode("utf-8"))
        except Exception:
            raise ValueError("Arquivo cifrado danificado.")


def ler_meta_cifrada(fonte, senha):
    """Confere a senha e devolve os metadados, sem decifrar o restante. fonte: caminho ou arquivo."""
    if hasattr(fonte, "read"):
        return _LeitorCifrado(fonte, senha).meta()
    with open(fonte, "rb") as f:
        return _LeitorCifrado(f, senha).meta()


def decifrar_arquivo(fonte, destino, senha):
    """Decifra para `destino` (removido se algo falhar). Devolve os metadados."""
    def processar(f):
        leitor = _LeitorCifrado(f, senha)
        meta = leitor.meta()
        terminou = False
        with open(destino, "wb") as out:
            while True:
                reg = leitor.proximo()
                if reg is None:
                    break
                if terminou or reg[0] == 0:
                    raise ValueError("Arquivo cifrado danificado.")
                out.write(reg[1])
                terminou = reg[0] == 2
        if not terminou:
            raise ValueError("Arquivo cifrado incompleto (o download ou a cópia pode ter sido interrompido).")
        return meta
    try:
        if hasattr(fonte, "read"):
            return processar(fonte)
        with open(fonte, "rb") as f:
            return processar(f)
    except Exception:
        try:
            Path(destino).unlink()
        except OSError:
            pass
        raise


def pasta_temporaria() -> Path:
    """Arquivos temporarios do programa (dentro da pasta de dados do usuario, nunca no OneDrive)."""
    p = pasta_dados_app() / "tmp"
    p.mkdir(exist_ok=True)
    return p


def _tmp(nome):
    return pasta_temporaria() / f"{nome}-{os.getpid()}-{secrets.token_hex(4)}"


def limpar_temporarios():
    """Apaga temporarios esquecidos (ex.: programa fechado no meio de uma importacao)."""
    try:
        for f in pasta_temporaria().iterdir():
            try:
                f.unlink()
            except OSError:
                pass
    except OSError:
        pass


# --------------------------------------------------------------------------
# Protecao de dados pessoais (LGPD): pastas e arquivos protegidos
# --------------------------------------------------------------------------
#
# Regras guardadas DENTRO do indice (tabela meta), para valerem tambem no computador do
# cliente quando o indice viaja num pacote. Cada regra: {"termo": "ASO", "modo": ...}
#   "so_nome": o arquivo aparece na busca so pelo nome; o conteudo nao e lido
#   "ignorar": o arquivo nao entra no indice
# O termo casa com uma palavra (ou sequencia de palavras) inteira do caminho relativo:
# "ASO" casa com "ASO", "01-ASO", "ASO_Joao.pdf", mas nao com "CASO".

MODO_SO_NOME = "so_nome"
MODO_IGNORAR = "ignorar"
MODOS_PROTECAO = {MODO_SO_NOME: "Só o nome (conteúdo não é lido)",
                  MODO_IGNORAR: "Ignorar (não entra no índice)"}
SUGESTOES_PROTECAO = ["ASO", "Atestado", "Exame", "Exames", "Prontuário", "Laudo médico", "Saúde",
                      "Documentos pessoais", "RG", "CPF", "CNH", "Holerite", "Contracheque"]
STATUS_PROTEGIDO = "protegido (só nome)"
MAX_REGRAS = 200


def _normalizar_palavras(t):
    t = unicodedata.normalize("NFKD", str(t or ""))
    t = "".join(c for c in t if not unicodedata.combining(c)).lower()
    return " ".join(re.findall(r"[a-z0-9]+", t))


def validar_regras(regras):
    """Limpa e valida a lista de regras. Devolve a lista normalizada (sem duplicatas)."""
    saida, vistos = [], {}
    for r in regras or []:
        if not isinstance(r, dict):
            continue
        termo = str(r.get("termo") or "").strip()[:60]
        modo = r.get("modo")
        if modo not in MODOS_PROTECAO:
            raise ValueError(f"Modo de proteção inválido: {modo!r}")
        chave = _normalizar_palavras(termo)
        if len(chave) < 2:
            raise ValueError(f"Termo de proteção curto demais: {termo!r} (mínimo 2 letras ou números).")
        if chave in vistos:                       # repetido: "ignorar" vence "so_nome"
            if modo == MODO_IGNORAR:
                saida[vistos[chave]]["modo"] = MODO_IGNORAR
            continue
        vistos[chave] = len(saida)
        saida.append({"termo": termo, "modo": modo})
    if len(saida) > MAX_REGRAS:
        raise ValueError(f"Máximo de {MAX_REGRAS} regras de proteção.")
    return saida


class Protecao:
    """Aplica as regras a caminhos relativos (pastas + nome do arquivo)."""

    def __init__(self, regras):
        self.regras = validar_regras(regras)
        self._chaves = [(" " + _normalizar_palavras(r["termo"]) + " ", r["modo"]) for r in self.regras]

    def __bool__(self):
        return bool(self._chaves)

    def modo(self, caminho_relativo):
        if not self._chaves:
            return None
        alvo = " " + _normalizar_palavras(caminho_relativo) + " "
        achou = None
        for chave, modo in self._chaves:
            if chave in alvo:
                if modo == MODO_IGNORAR:
                    return MODO_IGNORAR
                achou = MODO_SO_NOME
        return achou


def _juntar_regras(a, b):
    return validar_regras(list(a or []) + list(b or []))


# --------------------------------------------------------------------------
# Indice
# --------------------------------------------------------------------------

def caminho_indice(pasta_raiz: str) -> Path:
    h = hashlib.sha1(os.path.normcase(os.path.abspath(pasta_raiz)).encode("utf-8")).hexdigest()[:16]
    d = pasta_dados_app() / "indices"
    d.mkdir(exist_ok=True)
    return d / f"{h}.db"


ESQUEMA = """
CREATE TABLE IF NOT EXISTS arquivos (
    caminho  TEXT PRIMARY KEY,
    mtime    REAL,
    tamanho  INTEGER,
    status   TEXT
);
CREATE VIRTUAL TABLE IF NOT EXISTS conteudo USING fts5(
    caminho UNINDEXED,
    interno UNINDEXED,
    nome,
    texto,
    tokenize = "unicode61 remove_diacritics 2"
);
CREATE TABLE IF NOT EXISTS meta (chave TEXT PRIMARY KEY, valor TEXT);
"""


class Indice:
    def __init__(self, pasta_raiz: str):
        raiz = os.path.abspath(pasta_raiz)
        self.db_path = caminho_indice(raiz)
        self.con = sqlite3.connect(str(self.db_path), check_same_thread=False, timeout=30)
        self.con.execute("PRAGMA journal_mode=WAL")  # permite buscar enquanto indexa
        self.con.executescript(ESQUEMA)
        # Reaproveita a grafia da raiz gravada no indice ("C:\Dados" x "c:\dados"),
        # para que os caminhos guardados e o filtro por empresa usem o mesmo prefixo.
        r = self.con.execute("SELECT valor FROM meta WHERE chave='raiz'").fetchone()
        if r and os.path.normcase(r[0]) == os.path.normcase(raiz):
            raiz = r[0]
        else:
            self.con.execute("INSERT OR REPLACE INTO meta VALUES ('raiz', ?)", (raiz,))
            self.con.commit()
        self.raiz = raiz

    def fechar(self):
        self.con.close()

    # ---- protecao de dados pessoais ---------------------------------------

    def regras_protecao(self):
        r = self.con.execute("SELECT valor FROM meta WHERE chave='protecao'").fetchone()
        try:
            return validar_regras(json.loads(r[0])) if r and r[0] else []
        except (ValueError, TypeError):
            log.warning("Regras de proteção inválidas no índice; ignoradas")
            return []

    def _rel(self, caminho):
        return caminho[len(self.raiz):].lstrip("\\/") if caminho.startswith(self.raiz) else caminho

    def definir_regras_protecao(self, regras):
        """Grava as regras e ja as aplica ao que esta indexado (sem precisar reindexar).
        Devolve (removidos, so_nome): quantos arquivos sairam do indice / ficaram so com o nome."""
        regras = validar_regras(regras)
        with self.con:
            self.con.execute("INSERT OR REPLACE INTO meta VALUES ('protecao', ?)",
                             (json.dumps(regras, ensure_ascii=False),))
        return self.aplicar_protecao()

    def aplicar_protecao(self):
        prot = Protecao(self.regras_protecao())
        removidos = so_nome = 0
        if not prot:
            return 0, 0
        with self.con:
            for caminho, status in self.con.execute("SELECT caminho, status FROM arquivos").fetchall():
                modo = prot.modo(self._rel(caminho))
                if modo == MODO_IGNORAR:
                    self._remover(caminho)
                    removidos += 1
                elif modo == MODO_SO_NOME and status != STATUS_PROTEGIDO:
                    self.con.execute("DELETE FROM conteudo WHERE caminho = ?", (caminho,))
                    self.con.execute("INSERT INTO conteudo (caminho, interno, nome, texto) VALUES (?,?,?,?)",
                                     (caminho, "", _nome_pesquisavel(caminho, self.raiz), ""))
                    self.con.execute("UPDATE arquivos SET status = ? WHERE caminho = ?", (STATUS_PROTEGIDO, caminho))
                    so_nome += 1
            # documentos dentro de .zip cujo nome interno e protegido
            for rowid, caminho, interno in self.con.execute(
                    "SELECT rowid, caminho, interno FROM conteudo WHERE interno <> ''").fetchall():
                modo = prot.modo(interno)
                if modo == MODO_IGNORAR:
                    self.con.execute("DELETE FROM conteudo WHERE rowid = ?", (rowid,))
                elif modo == MODO_SO_NOME:
                    self.con.execute("UPDATE conteudo SET texto = '' WHERE rowid = ?", (rowid,))
        return removidos, so_nome

    def empresas(self):
        return listar_empresas(self.raiz)

    # ---- atualizacao incremental ----------------------------------------

    def _varrer(self, cancelado, avisar):
        """
        Lista os arquivos da pasta principal (e subpastas) sem abrir nenhum deles:
        nao dispara download do OneDrive. Devolve (no_disco, nuvem, falhas):
          no_disco: {caminho: (mtime, tamanho)}
          nuvem:    caminhos que estao so na nuvem
          falhas:   quantas pastas/arquivos nao puderam ser listados (rede, permissao...)
        """
        no_disco, nuvem, falhas = {}, set(), 0
        pilha = [self.raiz]
        while pilha:
            if cancelado():
                break
            pasta = pilha.pop()
            try:
                with os.scandir(caminho_longo(pasta)) as it:   # pastas profundas (> 260 caracteres)
                    entradas = list(it)
            except OSError:
                falhas += 1
                log.warning("Pasta inacessível durante a listagem: %s", pasta)
                continue
            subpastas = []
            for e in entradas:
                nome = e.name
                try:
                    caminho = os.path.join(pasta, nome)              # guardado sem o prefixo \\?\\
                    if e.is_dir(follow_symlinks=False):
                        if not nome.startswith(("~", "$", ".")):   # ignora pastas ocultas/sistema
                            subpastas.append(caminho)
                        continue
                    if nome.startswith("~$"):                       # temporario do Office
                        continue
                    st = e.stat(follow_symlinks=False)
                except OSError:
                    falhas += 1
                    continue
                no_disco[caminho] = (st.st_mtime, st.st_size)
                if so_na_nuvem(st):
                    nuvem.add(caminho)
            pilha.extend(sorted(subpastas, key=str.casefold, reverse=True))  # ordem alfabetica
            avisar(f"Listando arquivos... {len(no_disco)}")
        return no_disco, nuvem, falhas

    def atualizar(self, usar_ocr=True, progresso=None, cancelado=lambda: False, liberar_nuvem=True,
                  tentar_erros=False):
        """
        Varre a pasta e (re)indexa so o que mudou.
        progresso(msg, atual, total) e chamado periodicamente.
        liberar_nuvem: arquivos que estavam so no OneDrive voltam para a nuvem depois de lidos.
        tentar_erros: le de novo tambem os arquivos que deram erro antes (mesmo sem mudanca).
        Devolve dict com contagens.
        """
        def avisar(msg, i=0, n=0):
            if progresso:
                progresso(msg, i, n)

        avisar("Listando arquivos...")
        no_disco, nuvem, falhas = self._varrer(cancelado, avisar)

        if cancelado():
            # varredura incompleta: nao remove nada do indice com base nela
            return {"total": len(no_disco), "lidos": 0, "removidos": 0, "erros": 0,
                    "liberados": 0, "falhas": falhas, "cancelado": True}

        # protecao LGPD: arquivos "ignorar" saem da lista (e do indice, mesmo com pastas inacessiveis)
        prot = Protecao(self.regras_protecao())
        modos = {}
        if prot:
            for p in list(no_disco):
                m = prot.modo(self._rel(p))
                if m == MODO_IGNORAR:
                    del no_disco[p]
                elif m:
                    modos[p] = m
            for (p,) in self.con.execute("SELECT caminho FROM arquivos").fetchall():
                if prot.modo(self._rel(p)) == MODO_IGNORAR:
                    self._remover(p)
            self.con.commit()

        no_indice = {r[0]: (r[1], r[2]) for r in
                     self.con.execute("SELECT caminho, mtime, tamanho FROM arquivos")}
        status_ant = dict(self.con.execute("SELECT caminho, status FROM arquivos").fetchall())

        # Se alguma pasta nao pode ser lida (rede/VPN instavel), nao remove nada:
        # um arquivo "sumido" pode estar apenas inacessivel naquele momento.
        removidos = [] if falhas else [p for p in no_indice if p not in no_disco]
        novos = [p for p, mt in no_disco.items() if _mudou(no_indice.get(p), mt)
                 # mudou a regra: protegido -> normal (reler) ou normal -> protegido (tirar o conteudo)
                 or (status_ant.get(p) == STATUS_PROTEGIDO) != (modos.get(p) == MODO_SO_NOME)]
        if tentar_erros:
            ja = set(novos)
            novos += [p for (p,) in self.con.execute("SELECT caminho FROM arquivos WHERE status LIKE 'erro%'")
                      if p in no_disco and p not in ja]

        for p in removidos:
            self._remover(p)
        self.con.commit()

        erros = liberados = lidos = 0
        total = len(novos)
        t_commit = time.time()
        for i, p in enumerate(novos, 1):
            if cancelado():
                break
            nome = os.path.basename(p)
            avisar(f"Lendo {nome}", i, total)
            mtime, tam = no_disco[p]
            status = "ok"
            leu_conteudo = (tam <= MAX_TAMANHO_MB * 1024 * 1024
                            and os.path.splitext(nome)[1].lower() in EXT_CONTEUDO)
            try:
                if modos.get(p) == MODO_SO_NOME:
                    partes = [("", "")]                     # protegido: nem abre o arquivo
                    status = STATUS_PROTEGIDO
                    leu_conteudo = False
                elif tam > MAX_TAMANHO_MB * 1024 * 1024:
                    partes = [("", "")]
                    status = "grande"
                else:
                    partes = extrair(nome, caminho_longo(p), usar_ocr,
                                     lambda m, _n=nome, _i=i: avisar(f"{_n}: {m}", _i, total))
            except Exception as e:
                partes = [("", "")]
                status = f"erro: {type(e).__name__}: {e}"[:300]
                erros += 1
            # so devolve a nuvem o que estava na nuvem antes e foi baixado para leitura
            if liberar_nuvem and leu_conteudo and p in nuvem and liberar_espaco(caminho_longo(p)):
                liberados += 1
            self._remover(p)
            nome_busca = _nome_pesquisavel(p, self.raiz)
            for interno, texto in partes:
                if interno and prot:                        # documento dentro de .zip
                    m = prot.modo(interno)
                    if m == MODO_IGNORAR:
                        continue
                    if m == MODO_SO_NOME:
                        texto = ""
                self.con.execute(
                    "INSERT INTO conteudo (caminho, interno, nome, texto) VALUES (?,?,?,?)",
                    (p, interno, nome_busca, texto))
            self.con.execute("INSERT OR REPLACE INTO arquivos VALUES (?,?,?,?)",
                             (p, mtime, tam, status))
            lidos += 1
            if time.time() - t_commit > 5:
                self.con.commit()
                t_commit = time.time()
        self.con.commit()
        self.con.execute("INSERT OR REPLACE INTO meta VALUES ('atualizado', ?)",
                         (time.strftime("%d/%m/%Y %H:%M"),))
        self.con.commit()
        # lidos = efetivamente lidos (num cancelamento, menos que a_ler); o que ja foi lido fica salvo
        return {"total": len(no_disco), "lidos": lidos, "a_ler": total, "removidos": len(removidos),
                "erros": erros, "liberados": liberados, "falhas": falhas, "cancelado": cancelado()}

    def contem(self, caminho) -> bool:
        """True se o caminho esta dentro da pasta principal deste indice."""
        raiz = os.path.normcase(os.path.join(self.raiz, ""))
        return os.path.normcase(os.path.abspath(caminho)).startswith(raiz)

    def _remover(self, p):
        self.con.execute("DELETE FROM conteudo WHERE caminho = ?", (p,))
        self.con.execute("DELETE FROM arquivos WHERE caminho = ?", (p,))

    def info(self):
        n = self.con.execute("SELECT COUNT(*) FROM arquivos").fetchone()[0]
        r = self.con.execute("SELECT valor FROM meta WHERE chave='atualizado'").fetchone()
        return n, (r[0] if r else None)

    def arquivos_com_erro(self):
        return self.con.execute(
            "SELECT caminho, status FROM arquivos WHERE status LIKE 'erro%' ORDER BY caminho").fetchall()

    # ---- busca ------------------------------------------------------------

    def _prefixo_seguro(self, pasta):
        """Prefixo da pasta (com barra final), garantindo que ela esta dentro da raiz."""
        prefixo = os.path.join(os.path.abspath(pasta), "")
        raiz = os.path.join(self.raiz, "")
        if not os.path.normcase(prefixo).startswith(os.path.normcase(raiz)):
            raise ValueError("A pasta escolhida está fora da pasta principal.")
        # usa a grafia da raiz gravada + o restante do caminho
        return raiz + prefixo[len(raiz):]

    def buscar(self, termo: str, pasta: str = None, limite=2000):
        """
        Regras:
          trafo            -> palavras que comecam com "trafo" (trafo, trafos...)
          "trafo 500 kva"  -> frase exata
          trafo bucha      -> documentos com as duas palavras
          trafo OU bucha   -> qualquer uma das palavras
        pasta: se informada, restringe a busca a essa pasta e subpastas (ex.: uma empresa).
        Devolve lista de dicts: caminho, interno, nome, pasta, ext, trecho, onde.
        """
        consulta = montar_consulta(termo)
        if not consulta:
            return []
        filtro, params = "", [consulta]
        if pasta:
            prefixo = self._prefixo_seguro(pasta)
            # comparacao exata de prefixo (LIKE trataria "_" de "DB_" como curinga)
            filtro = " AND substr(caminho, 1, ?) = ?"
            params += [len(prefixo), prefixo]
        params.append(limite)
        sql = f"""
            SELECT caminho, interno,
                   snippet(conteudo, 3, '[', ']', ' … ', 14) AS trecho,
                   snippet(conteudo, 2, '[', ']', ' … ', 14) AS trecho_nome,
                   bm25(conteudo, 0, 0, 5.0, 1.0) AS rank
            FROM conteudo
            WHERE conteudo MATCH ?{filtro}
            ORDER BY rank
            LIMIT ?
        """
        resultados = []
        vistos = set()
        for caminho, interno, trecho, trecho_nome, _ in self.con.execute(sql, params):
            chave = (caminho, interno)
            if chave in vistos:
                continue
            vistos.add(chave)
            achou_no_texto = "[" in (trecho or "")
            resultados.append({
                "caminho": caminho,
                "interno": interno,
                "nome": os.path.basename(caminho),
                "pasta": os.path.dirname(caminho),
                "ext": os.path.splitext(interno or caminho)[1].lower().lstrip("."),
                "trecho": _limpar(trecho if achou_no_texto else trecho_nome),
                "onde": "conteúdo" if achou_no_texto else "nome",
            })
        meta = self._meta_arquivos(r["caminho"] for r in resultados)
        for r in resultados:
            r["mtime"], r["tamanho"] = meta.get(r["caminho"], (0.0, 0))
        return resultados

    def buscar_nome(self, termo: str, pasta: str = None, limite=2000):
        """
        Busca SO pelo nome do arquivo (com extensao), ignorando o conteudo e as pastas.
        Vale para arquivos comuns e para arquivos dentro de .zip.
        Mesmas regras de digitacao da busca normal (prefixo, "frase exata", OU).
        """
        grupos = _grupos_nome(termo)
        if not grupos:
            return []
        filtro, params = "", []
        if pasta:
            prefixo = self._prefixo_seguro(pasta)
            filtro = " WHERE substr(caminho, 1, ?) = ?"
            params = [len(prefixo), prefixo]
        achados = {}

        # 1) arquivos comuns e o proprio .zip: tabela pequena (um registro por arquivo)
        for caminho, mtime, tam in self.con.execute(
                f"SELECT caminho, mtime, tamanho FROM arquivos{filtro}", params):
            if _nome_casa(os.path.basename(caminho), grupos):
                achados[(caminho, "")] = (mtime or 0.0, tam or 0)

        # 2) arquivos dentro de .zip: o indice pre-seleciona, o nome confirma
        consulta = montar_consulta(termo)
        if consulta:
            sql_zip = ("SELECT caminho, interno FROM conteudo WHERE conteudo MATCH ? AND interno <> ''"
                       + (" AND substr(caminho, 1, ?) = ?" if pasta else "") + " LIMIT 20000")
            zips = [(c, i) for c, i in self.con.execute(sql_zip, [consulta] + params)
                    if _nome_casa(_nome_documento({"interno": i, "nome": ""}), grupos)]
            meta = self._meta_arquivos(c for c, _ in zips)
            for c, i in zips:
                achados[(c, i)] = meta.get(c, (0.0, 0))

        resultados = []
        for (caminho, interno), (mtime, tam) in achados.items():
            resultados.append({
                "caminho": caminho,
                "interno": interno,
                "nome": os.path.basename(caminho),
                "pasta": os.path.dirname(caminho),
                "ext": os.path.splitext(interno or caminho)[1].lower().lstrip("."),
                "trecho": "",
                "onde": "nome",
                "mtime": mtime,
                "tamanho": tam,
            })
        resultados.sort(key=lambda r: (_nome_documento(r).casefold(), r["caminho"].casefold()))
        return resultados[:limite]

    def _meta_arquivos(self, caminhos):
        """{caminho: (mtime, tamanho)} - consulta em lotes (limite de parametros do SQLite)."""
        meta, lista = {}, list(set(caminhos))
        for i in range(0, len(lista), 500):
            parte = lista[i:i + 500]
            marcas = ",".join("?" * len(parte))
            for c, mt, tam in self.con.execute(
                    f"SELECT caminho, mtime, tamanho FROM arquivos WHERE caminho IN ({marcas})", parte):
                meta[c] = (mt or 0.0, tam or 0)
        return meta

    # ---- exportar / importar ----------------------------------------------

    def _filtro_exportacao(self, pastas, incluir_fora):
        """WHERE para exportar so algumas empresas. pastas=None exporta o indice inteiro."""
        if pastas is None:
            return "1", []
        conds, params = [], []
        for p in pastas:
            pref = self._prefixo_seguro(p)
            conds.append("substr(caminho, 1, ?) = ?")
            params += [len(pref), pref]
        if incluir_fora:   # arquivos que nao estao em nenhuma pasta DB_
            todas = [self._prefixo_seguro(c) for _, c in self.empresas()]
            if todas:
                conds.append("(" + " AND ".join("substr(caminho, 1, ?) <> ?" for _ in todas) + ")")
                for t in todas:
                    params += [len(t), t]
            else:
                conds.append("1")
        if not conds:
            raise ValueError("Escolha ao menos uma empresa para exportar.")
        return "(" + " OR ".join(conds) + ")", params

    def exportar(self, destino, pastas=None, incluir_fora=True, senha=None):
        """
        Gera um arquivo .bdx com o indice inteiro ou so das empresas escolhidas, CIFRADO com a
        senha (obrigatoria: o indice contem o texto dos documentos). Devolve quantos arquivos.
        """
        checar_senha_arquivo(senha)
        plano = _tmp("exportando") 
        try:
            n = self._exportar_plano(plano, pastas, incluir_fora)
            cifrar_arquivo(plano, destino, senha, {"formato": FORMATO_EXPORTACAO, "raiz": self.raiz,
                                                   "arquivos": n, "exportado_em": time.strftime("%d/%m/%Y %H:%M")})
        finally:
            try:
                plano.unlink()
            except OSError:
                pass
        log.info("Índice exportado (cifrado): %s (%d arquivos)", destino, n)
        return n

    def _exportar_plano(self, destino, pastas=None, incluir_fora=True):
        """Gera o .bdx SEM cifra, so para uso interno (sempre numa pasta temporaria do programa)."""
        destino = Path(destino)
        tmp = destino.with_name(destino.name + ".tmp")
        if tmp.exists():
            tmp.unlink()
        filtro, params = self._filtro_exportacao(pastas, incluir_fora)
        con = sqlite3.connect(str(tmp))
        try:
            con.executescript(ESQUEMA)
            con.execute("ATTACH DATABASE ? AS origem", (str(self.db_path),))
            con.execute(f"INSERT INTO arquivos SELECT caminho, mtime, tamanho, status "
                        f"FROM origem.arquivos WHERE {filtro}", params)
            con.execute(f"INSERT INTO conteudo (caminho, interno, nome, texto) "
                        f"SELECT caminho, interno, nome, texto FROM origem.conteudo WHERE {filtro}", params)
            r = con.execute("SELECT valor FROM origem.meta WHERE chave='atualizado'").fetchall()
            con.executemany("INSERT OR REPLACE INTO meta VALUES (?, ?)", [
                ("raiz", self.raiz),
                ("formato", FORMATO_EXPORTACAO),
                ("exportado_em", time.strftime("%d/%m/%Y %H:%M")),
                ("atualizado", r[0][0] if r else ""),
                ("protecao", json.dumps(self.regras_protecao(), ensure_ascii=False)),
            ])
            con.commit()
            n = con.execute("SELECT COUNT(*) FROM arquivos").fetchall()[0][0]
            con.execute("DETACH DATABASE origem")
        except Exception:
            con.close()
            tmp.unlink(missing_ok=True)
            raise
        con.close()
        os.replace(tmp, destino)
        return n

    def mesclar(self, arquivo):
        """Junta um indice (ja com os caminhos desta pasta) ao indice atual.
        Arquivos que ja existiam sao substituidos; os demais sao mantidos."""
        self.con.commit()
        self.con.execute("ATTACH DATABASE ? AS imp", (str(arquivo),))
        try:
            with self.con:
                self.con.execute("DELETE FROM conteudo WHERE caminho IN (SELECT caminho FROM imp.arquivos)")
                self.con.execute("DELETE FROM arquivos WHERE caminho IN (SELECT caminho FROM imp.arquivos)")
                self.con.execute("INSERT INTO arquivos SELECT caminho, mtime, tamanho, status FROM imp.arquivos")
                self.con.execute("INSERT INTO conteudo (caminho, interno, nome, texto) "
                                 "SELECT caminho, interno, nome, texto FROM imp.conteudo")
                r = self.con.execute("SELECT valor FROM imp.meta WHERE chave='atualizado'").fetchall()
                if r and r[0][0]:
                    self.con.execute("INSERT OR REPLACE INTO meta VALUES ('atualizado', ?)", (r[0][0],))
                r = self.con.execute("SELECT valor FROM imp.meta WHERE chave='protecao'").fetchall()
                vindas = json.loads(r[0][0]) if r and r[0][0] else []
            n = self.con.execute("SELECT COUNT(*) FROM imp.arquivos").fetchall()[0][0]
        finally:
            self.con.execute("DETACH DATABASE imp")
        if vindas:   # regras de protecao viajam com o indice e passam a valer aqui tambem
            self.definir_regras_protecao(_juntar_regras(self.regras_protecao(), vindas))
        return n


def ler_exportacao(arquivo, senha=None):
    """Le e valida o cabecalho de um indice exportado. Devolve dict(raiz, arquivos, exportado_em, cifrado).
    Arquivo cifrado: confere a senha lendo so os metadados (rapido)."""
    if esta_cifrado(arquivo):
        if not senha:
            raise ValueError("Este índice é protegido por senha.")
        m = ler_meta_cifrada(arquivo, senha)
        if m.get("formato") != FORMATO_EXPORTACAO:
            raise ValueError("Este arquivo não é um índice exportado pelo Busca Databook.")
        return {"raiz": str(m.get("raiz") or ""), "arquivos": int(m.get("arquivos") or 0),
                "exportado_em": str(m.get("exportado_em") or ""), "cifrado": True}
    try:
        con = sqlite3.connect(str(arquivo))
    except sqlite3.Error as e:
        raise ValueError(f"Não foi possível abrir o arquivo: {e}")
    try:
        con.execute("PRAGMA query_only = 1")
        meta = dict(con.execute("SELECT chave, valor FROM meta").fetchall())
        if meta.get("formato") != FORMATO_EXPORTACAO:
            raise ValueError("Este arquivo não é um índice exportado pelo Busca Databook.")
        n = con.execute("SELECT COUNT(*) FROM arquivos").fetchone()[0]
    except sqlite3.DatabaseError as e:
        raise ValueError(f"Arquivo de índice inválido ou corrompido: {e}")
    finally:
        con.close()
    return {"raiz": meta.get("raiz", ""), "arquivos": n, "exportado_em": meta.get("exportado_em", ""),
            "cifrado": False}


def importar_indice(arquivo, nova_raiz, senha=None):
    """
    Importa um .bdx para a pasta principal nova_raiz (pode ser outro caminho que o de origem:
    os caminhos sao ajustados). Junta ao indice existente dessa pasta. Devolve quantos arquivos.
    Indices cifrados sao decifrados numa pasta temporaria do programa (apagada no fim).
    """
    tmp = _tmp("importando")
    idx = None
    try:
        if esta_cifrado(arquivo):
            decifrar_arquivo(arquivo, tmp, senha)
        else:
            shutil.copyfile(arquivo, tmp)   # trabalha numa copia: o arquivo recebido fica intacto
        info = ler_exportacao(tmp)
        idx = Indice(nova_raiz)
        con = sqlite3.connect(str(tmp))
        try:
            if con.execute("PRAGMA quick_check").fetchone()[0] != "ok":
                raise ValueError("O arquivo de índice está corrompido.")
            p_ant = os.path.join(info["raiz"], "")
            p_nova = os.path.join(idx.raiz, "")
            if p_ant != p_nova:
                for tabela in ("arquivos", "conteudo"):
                    con.execute(f"UPDATE {tabela} SET caminho = ? || substr(caminho, ?) "
                                f"WHERE substr(caminho, 1, ?) = ?",
                                (p_nova, len(p_ant) + 1, len(p_ant), p_ant))
            # seguranca: nenhum caminho pode apontar para fora da pasta escolhida
            fora = sum(con.execute(f"SELECT COUNT(*) FROM {t} WHERE substr(caminho, 1, ?) <> ?",
                                   (len(p_nova), p_nova)).fetchone()[0] for t in ("arquivos", "conteudo"))
            if fora:
                raise ValueError("O índice contém caminhos fora da pasta escolhida. Importação cancelada.")
            con.commit()
        finally:
            con.close()
        n = idx.mesclar(tmp)
        log.info("Índice importado de %s para %s (%d arquivos)", arquivo, idx.raiz, n)
        return n
    finally:
        if idx:
            idx.fechar()
        for f in (tmp, Path(str(tmp) + "-journal")):
            try:
                f.unlink()
            except OSError:
                pass


# --------------------------------------------------------------------------
# Pacote do cliente (.bdpack): identidade visual + pasta principal + indice
# --------------------------------------------------------------------------
#
# Um .bdpack e um zip com no maximo 3 itens:
#   manifest.json        - nome do cliente, cor do menu, pasta de origem, resumo do indice
#   logo.png             - logomarca do cliente (opcional)
#   indice.bdx.cifrado   - indice exportado e CIFRADO com senha (opcional; desde a 3.6)
#   indice.bdx           - indice sem cifra (so pacotes antigos, 3.3 a 3.5: ainda aceitos)
# Nada e extraido para caminhos vindos do arquivo: so esses nomes fixos sao lidos.

RE_COR = re.compile(r"^#[0-9A-Fa-f]{6}$")
MAX_LOGO_BYTES = 5 * 1024 * 1024
MAX_MANIFEST_BYTES = 64 * 1024


def pasta_marca() -> Path:
    p = pasta_dados_app() / "marca"
    p.mkdir(exist_ok=True)
    return p


def preparar_logo(dados: bytes) -> bytes:
    """Valida e re-grava a logomarca como PNG limpo (corta bordas transparentes, max. 600x200)."""
    if len(dados) > MAX_LOGO_BYTES:
        raise ValueError("A logomarca é grande demais (máximo 5 MB).")
    from PIL import Image
    try:
        with Image.open(io.BytesIO(dados)) as teste:
            teste.verify()
        img = Image.open(io.BytesIO(dados)).convert("RGBA")
    except Exception:
        raise ValueError("O arquivo da logomarca não é uma imagem válida.")
    caixa = img.getchannel("A").getbbox()
    if caixa:
        img = img.crop(caixa)
    img.thumbnail((600, 200), getattr(Image, "Resampling", Image).LANCZOS)
    saida = io.BytesIO()
    img.save(saida, "PNG")
    return saida.getvalue()


def _validar_manifest(m):
    if not isinstance(m, dict) or m.get("formato") != FORMATO_PACOTE:
        raise ValueError("Este arquivo não é um pacote do Busca Databook.")
    cliente = str(m.get("cliente") or "").strip()
    if not cliente or len(cliente) > 80:
        raise ValueError("Pacote sem nome de cliente válido.")
    cor = m.get("cor_menu")
    if cor is not None and not RE_COR.match(str(cor)):
        raise ValueError("Cor do menu inválida no pacote.")
    return {"cliente": cliente, "cor_menu": cor, "pasta_origem": str(m.get("pasta_origem") or ""),
            "indice": m.get("indice") if isinstance(m.get("indice"), dict) else None,
            "criado_em": str(m.get("criado_em") or "")}


def criar_pacote(destino, cliente, cor_menu=None, logo_bytes=None, indice=None, pastas=None, incluir_fora=True,
                 senha=None):
    """
    Gera o .bdpack. indice: um Indice aberto (ou None para pacote so com a identidade visual);
    pastas/incluir_fora: quais empresas exportar (mesma regra de Indice.exportar).
    Com indice, a senha e obrigatoria: o indice vai cifrado dentro do pacote.
    """
    if indice is not None:
        checar_senha_arquivo(senha)
    cliente = (cliente or "").strip()
    if not cliente:
        raise ValueError("Informe o nome do cliente.")
    if cor_menu and not RE_COR.match(cor_menu):
        raise ValueError("Cor inválida: use o formato #RRGGBB (ex.: #FA7C13).")
    logo = preparar_logo(logo_bytes) if logo_bytes else None
    manifest = {"formato": FORMATO_PACOTE, "cliente": cliente, "cor_menu": cor_menu or None,
                "criado_em": time.strftime("%d/%m/%Y %H:%M"), "pasta_origem": "", "indice": None}
    destino = Path(destino)
    tmp_bdx = tmp_cif = None
    try:
        if indice is not None:
            tmp_bdx, tmp_cif = _tmp("pacote"), _tmp("pacote-cifrado")
            n = indice._exportar_plano(tmp_bdx, pastas, incluir_fora)
            cifrar_arquivo(tmp_bdx, tmp_cif, senha, {"formato": FORMATO_EXPORTACAO, "raiz": indice.raiz,
                                                     "arquivos": n, "exportado_em": time.strftime("%d/%m/%Y %H:%M")})
            manifest["pasta_origem"] = indice.raiz
            manifest["indice"] = {"arquivos": n, "atualizado": (indice.info()[1] or ""), "cifrado": True,
                                  "protecao": len(indice.regras_protecao())}
        parcial = destino.with_name(destino.name + ".tmp")
        with zipfile.ZipFile(parcial, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
            if logo:
                z.writestr("logo.png", logo)
            if tmp_cif:
                z.write(tmp_cif, "indice.bdx.cifrado", compress_type=zipfile.ZIP_STORED)
        os.replace(parcial, destino)
    finally:
        for f in (tmp_bdx, tmp_cif):
            if f:
                try:
                    f.unlink()
                except OSError:
                    pass
    log.info("Pacote do cliente criado: %s (%s)", destino, cliente)
    return manifest


def ler_pacote(arquivo):
    """Le e valida o pacote sem extrair o indice. Devolve o manifest validado + tem_logo/tem_indice."""
    try:
        z = zipfile.ZipFile(arquivo)
    except (zipfile.BadZipFile, OSError) as e:
        raise ValueError(f"Não foi possível abrir o pacote: {e}")
    with z:
        nomes = {i.filename: i for i in z.infolist()}
        if (set(nomes) - {"manifest.json", "logo.png", "indice.bdx", "indice.bdx.cifrado"}
                or "manifest.json" not in nomes or {"indice.bdx", "indice.bdx.cifrado"} <= set(nomes)):
            raise ValueError("Pacote com conteúdo inesperado. Por segurança, ele não será usado.")
        if nomes["manifest.json"].file_size > MAX_MANIFEST_BYTES:
            raise ValueError("Pacote inválido (manifesto grande demais).")
        if "logo.png" in nomes and nomes["logo.png"].file_size > MAX_LOGO_BYTES:
            raise ValueError("Pacote inválido (logomarca grande demais).")
        try:
            m = json.loads(z.read("manifest.json").decode("utf-8"))
        except Exception:
            raise ValueError("Pacote inválido (manifesto ilegível).")
        info = _validar_manifest(m)
        info["tem_logo"] = "logo.png" in nomes
        info["indice_cifrado"] = "indice.bdx.cifrado" in nomes
        info["tem_indice"] = info["indice_cifrado"] or "indice.bdx" in nomes
        return info


def conferir_senha_pacote(arquivo, senha):
    """True se a senha abre o indice do pacote (le so os metadados cifrados: rapido)."""
    with zipfile.ZipFile(arquivo) as z, z.open("indice.bdx.cifrado") as f:
        try:
            ler_meta_cifrada(f, senha)
            return True
        except ValueError as e:
            if str(e) == "Senha incorreta.":
                return False
            raise


def aplicar_marca(cliente, cor_menu=None, logo_png=None):
    """Grava a identidade visual deste computador (vale para todos os usuarios do app neste Windows)."""
    d = pasta_marca()
    logo_arq = d / "logo.png"
    if logo_png:
        tmp = d / "logo.tmp"
        tmp.write_bytes(preparar_logo(logo_png))
        os.replace(tmp, logo_arq)
    elif logo_arq.exists():
        logo_arq.unlink()
    tmp = d / "marca.tmp"
    tmp.write_text(json.dumps({"cliente": cliente, "cor_menu": cor_menu}, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, d / "marca.json")


def carregar_marca():
    """{"cliente", "cor_menu", "logo"} da personalizacao instalada, ou None."""
    d = pasta_dados_app() / "marca"
    try:
        m = json.loads((d / "marca.json").read_text(encoding="utf-8"))
        cliente = str(m.get("cliente") or "").strip()[:80]
        cor = m.get("cor_menu") if m.get("cor_menu") and RE_COR.match(str(m.get("cor_menu"))) else None
        if not cliente:
            return None
        logo = d / "logo.png"
        return {"cliente": cliente, "cor_menu": cor, "logo": str(logo) if logo.is_file() else None}
    except Exception:
        return None


def remover_marca():
    d = pasta_dados_app() / "marca"
    for n in ("marca.json", "logo.png"):
        try:
            (d / n).unlink()
        except OSError:
            pass


def importar_pacote(arquivo, nova_raiz=None, senha=None):
    """
    Aplica o pacote: identidade visual e, se houver indice, importa-o para nova_raiz.
    Indice cifrado: decifra ANTES de aplicar qualquer coisa (senha errada nao muda nada).
    Devolve (info do pacote, arquivos importados).
    """
    info = ler_pacote(arquivo)
    n = 0
    tmp = None
    try:
        with zipfile.ZipFile(arquivo) as z:
            logo = z.read("logo.png") if info["tem_logo"] else None
            if info["tem_indice"]:
                if not nova_raiz:
                    raise ValueError("Escolha a pasta principal para importar o índice.")
                tmp = _tmp("pacote-imp")
                if info["indice_cifrado"]:
                    with z.open("indice.bdx.cifrado") as src:
                        decifrar_arquivo(src, tmp, senha)
                else:                                       # pacote antigo (3.3 a 3.5)
                    with z.open("indice.bdx") as src, open(tmp, "wb") as dst:
                        shutil.copyfileobj(src, dst, 1 << 20)
        aplicar_marca(info["cliente"], info["cor_menu"], logo)
        if tmp:
            n = importar_indice(tmp, nova_raiz)
    finally:
        if tmp:
            try:
                tmp.unlink()
            except OSError:
                pass
    log.info("Pacote do cliente aplicado: %s (%d arquivos)", info["cliente"], n)
    return info, n


# --------------------------------------------------------------------------
# Eliminacao de dados (LGPD, arts. 15 e 16)
# --------------------------------------------------------------------------

def apagar_indice(pasta_raiz):
    """Apaga o indice de uma pasta principal (o banco e seus arquivos auxiliares).
    Os documentos originais nao sao tocados. O indice precisa estar fechado."""
    base = caminho_indice(os.path.abspath(pasta_raiz))
    apagados = 0
    for f in (base, Path(str(base) + "-wal"), Path(str(base) + "-shm"), Path(str(base) + "-journal")):
        if f.exists():
            f.unlink()
            apagados += 1
    return apagados > 0


def apagar_todos_os_dados():
    """Apaga TUDO o que o programa guarda neste usuario do Windows: indices, usuarios, configuracao,
    personalizacao, registros, temporarios e atualizacoes baixadas. Os documentos nao sao tocados.
    Antes de chamar, feche os indices e os arquivos de log. Devolve a lista do que nao foi apagado."""
    falhas = []
    raiz = pasta_dados_app()

    def erro(func, caminho, exc):
        falhas.append(str(caminho))
    for item in raiz.iterdir():
        try:
            if item.is_dir():
                shutil.rmtree(item, onerror=erro)
            else:
                item.unlink()
        except OSError:
            falhas.append(str(item))
    return falhas


# --------------------------------------------------------------------------
# Copias e versoes repetidas nos resultados
# --------------------------------------------------------------------------

def _mudou(no_indice, no_disco) -> bool:
    """Compara (mtime, tamanho) com tolerancia de 1 ms na data."""
    if no_indice is None:
        return True
    return no_indice[1] != no_disco[1] or abs((no_indice[0] or 0) - no_disco[0]) > 0.001


def _nome_documento(r) -> str:
    """Nome do documento (para zip: nome do arquivo interno)."""
    if r["interno"]:
        ultimo = r["interno"].split(SEP_ZIP)[-1]
        return ultimo.replace("\\", "/").rsplit("/", 1)[-1]
    return r["nome"]


def _empresa_do_caminho(caminho, prefixo_raiz) -> str:
    """Pasta DB_ de primeiro nivel onde o arquivo esta ("" se nao estiver em nenhuma)."""
    if not os.path.normcase(caminho).startswith(os.path.normcase(prefixo_raiz)):
        return ""
    primeira = re.split(r"[\\/]", caminho[len(prefixo_raiz):], maxsplit=1)[0]
    return primeira.casefold() if primeira.upper().startswith(PREFIXO_EMPRESA) else ""


def agrupar_repetidos(resultados, raiz):
    """
    Junta documentos com o mesmo nome dentro da mesma empresa e mostra so um:
      - versoes (tamanhos diferentes): fica a mais recente;
      - copias identicas (mesmo tamanho): fica a do caminho mais curto.
    Cada item devolvido traz em "copias" todos os documentos do grupo (mais recente primeiro).
    A ordem de relevancia da busca e mantida.
    """
    prefixo = os.path.join(raiz, "")
    grupos, ordem = {}, []
    for r in resultados:
        chave = (_empresa_do_caminho(r["caminho"], prefixo), _nome_documento(r).casefold())
        if chave not in grupos:
            grupos[chave] = []
            ordem.append(chave)
        grupos[chave].append(r)
    saida = []
    for chave in ordem:
        membros = grupos[chave]
        versoes = {}
        for r in membros:
            versoes.setdefault(r.get("tamanho", 0), []).append(r)
        recente = max(versoes.values(), key=lambda v: max(x.get("mtime", 0) for x in v))
        principal = dict(min(recente, key=lambda x: (len(x["caminho"]) + len(x["interno"]), x["caminho"])))
        principal["copias"] = sorted(membros, key=lambda x: (-x.get("mtime", 0), len(x["caminho"])))
        saida.append(principal)
    return saida


def _nome_pesquisavel(caminho, raiz):
    """Nome do arquivo + subpastas relativas, para achar tambem por pasta."""
    rel = os.path.relpath(caminho, raiz)
    base = os.path.splitext(rel)[0]
    # separa palavras grudadas por _ - . para o tokenizador
    return re.sub(r"[_\-.\\/]+", " ", base)


def _limpar(t):
    return re.sub(r"\s+", " ", t or "").strip()


def analisar_termo(termo: str):
    """
    Interpreta o texto digitado. Devolve grupos ligados por OU; cada grupo e uma lista de
    itens ligados por E: ("prefixo", [palavra]) ou ("frase", [palavras...]).
    """
    grupos, atual = [], []
    for p in re.findall(r'"[^"]+"|\S+', (termo or "").strip()):
        if p.upper() in ("OU", "OR"):
            if atual:
                grupos.append(atual)
            atual = []
            continue
        if p.startswith('"') and p.endswith('"') and len(p) > 2:
            palavras = re.findall(r"\w+", p[1:-1], flags=re.UNICODE)
            if palavras:
                atual.append(("frase", palavras))
        else:
            for tok in re.findall(r"\w+", p, flags=re.UNICODE):
                atual.append(("prefixo", [tok]))
    if atual:
        grupos.append(atual)
    return [g for g in grupos if g]


def montar_consulta(termo: str) -> str:
    """Converte o texto digitado em consulta FTS5 segura (sem operadores do usuario)."""
    grupos = analisar_termo(termo)
    if not grupos:
        return ""
    def item(tipo, palavras):
        return '"' + " ".join(palavras) + '"' + ("*" if tipo == "prefixo" else "")
    return " OR ".join("(" + " AND ".join(item(t, p) for t, p in g) + ")" for g in grupos)


def _palavras_nome(texto):
    """Palavras de um nome, sem acento e sem diferenciar maiusculas ("_", "-", "." separam)."""
    sem_acento = "".join(c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c))
    return re.findall(r"[^\W_]+", sem_acento.casefold())


def _grupos_nome(termo):
    """analisar_termo() ja normalizado para comparar com nomes de arquivo."""
    saida = []
    for g in analisar_termo(termo):
        itens = []
        for tipo, palavras in g:
            norm = [w for p in palavras for w in _palavras_nome(p)]
            if norm:
                itens.append((tipo, norm))
        if itens:
            saida.append(itens)
    return saida


def _nome_casa(nome, grupos) -> bool:
    """True se o nome do arquivo atende a busca (palavra por prefixo; frase exata e em sequencia)."""
    palavras = _palavras_nome(nome)
    if not palavras:
        return False

    def casa(tipo, termos):
        if tipo == "prefixo":
            return all(any(w.startswith(t) for w in palavras) for t in termos)
        n = len(termos)
        return any(palavras[i:i + n] == termos for i in range(len(palavras) - n + 1))

    return any(all(casa(t, ws) for t, ws in g) for g in grupos)


# --------------------------------------------------------------------------
# Usuarios (login)
# --------------------------------------------------------------------------

ALFABETO_CODIGO = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"   # sem I, O, 0, 1 (evita confusao ao digitar)
TAM_CODIGO = 12                                       # ~60 bits: inviavel adivinhar


def _normalizar_codigo(codigo):
    return re.sub(r"[^A-Z0-9]", "", (codigo or "").upper())


class Usuarios:
    """
    Usuarios locais deste computador. Senhas e codigos de recuperacao sao guardados so
    como hash PBKDF2 (com sal): nem o programa nem o administrador conseguem le-los.
    """
    ITERACOES = 600_000          # OWASP (PBKDF2-HMAC-SHA256); hashes antigos (200 mil) sao refeitos no login
    ITERACOES_ANTIGAS = 200_000

    def __init__(self):
        self.con = sqlite3.connect(str(pasta_dados_app() / "usuarios.db"))
        self.con.execute("""CREATE TABLE IF NOT EXISTS usuarios (
            login TEXT PRIMARY KEY, sal BLOB, hash BLOB, admin INTEGER)""")
        # colunas novas (3.2): senha provisoria e codigo de recuperacao
        cols = {r[1] for r in self.con.execute("PRAGMA table_info(usuarios)")}
        for nome, tipo in (("provisoria", "INTEGER DEFAULT 0"), ("rec_sal", "BLOB"), ("rec_hash", "BLOB"),
                           ("iter", f"INTEGER DEFAULT {self.ITERACOES_ANTIGAS}"),
                           ("rec_iter", f"INTEGER DEFAULT {self.ITERACOES_ANTIGAS}")):
            if nome not in cols:
                self.con.execute(f"ALTER TABLE usuarios ADD COLUMN {nome} {tipo}")
        self.con.commit()

    @classmethod
    def _hash(cls, senha, sal, iteracoes=None):
        return hashlib.pbkdf2_hmac("sha256", senha.encode("utf-8"), sal, iteracoes or cls.ITERACOES)

    @staticmethod
    def _checar_senha(senha, login=""):
        faltam = problemas_senha(senha, login)
        if faltam:
            raise ValueError("A senha precisa:\n- " + "\n- ".join(faltam))

    @staticmethod
    def _aleatorio(n):
        return "".join(secrets.choice(ALFABETO_CODIGO) for _ in range(n))

    def vazio(self):
        return self.con.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0] == 0

    def criar(self, login, senha=None, admin=False):
        """Cria o usuario. Sem senha informada, gera uma senha provisoria e a devolve."""
        login = login.strip().lower()
        if not login:
            raise ValueError("Informe um nome de usuário.")
        provisoria = senha is None
        if provisoria:
            senha = self._aleatorio(10)           # aleatoria e de uso unico: troca obrigatoria no 1o acesso
        else:
            self._checar_senha(senha, login)
        if self.con.execute("SELECT 1 FROM usuarios WHERE login=?", (login,)).fetchone():
            raise ValueError("Já existe um usuário com esse nome.")
        sal = os.urandom(16)
        self.con.execute("INSERT INTO usuarios (login, sal, hash, admin, provisoria, iter) VALUES (?,?,?,?,?,?)",
                         (login, sal, self._hash(senha, sal), int(admin), int(provisoria), self.ITERACOES))
        self.con.commit()
        return senha if provisoria else None

    def validar(self, login, senha):
        """Devolve {login, admin, provisoria, tem_codigo} se a senha estiver certa; senao None."""
        login = login.strip().lower()
        r = self.con.execute("SELECT sal, hash, admin, provisoria, rec_hash, iter FROM usuarios WHERE login=?",
                             (login,)).fetchone()
        if r:
            if hmac.compare_digest(self._hash(senha, r[0], r[5]), r[1]):
                if (r[5] or 0) < self.ITERACOES:          # hash antigo: refaz com o custo atual
                    sal = os.urandom(16)
                    self.con.execute("UPDATE usuarios SET sal=?, hash=?, iter=? WHERE login=?",
                                     (sal, self._hash(senha, sal), self.ITERACOES, login))
                    self.con.commit()
                return {"login": login, "admin": bool(r[2]), "provisoria": bool(r[3]), "tem_codigo": bool(r[4])}
        else:
            # mesmo custo de tempo para usuario inexistente (nao revela quais logins existem)
            self._hash(senha, b"\0" * 16)
        return None

    def listar(self):
        """[(login, admin, provisoria)]"""
        return self.con.execute(
            "SELECT login, admin, COALESCE(provisoria, 0) FROM usuarios ORDER BY login").fetchall()

    def remover(self, login):
        self.con.execute("DELETE FROM usuarios WHERE login=?", (login,))
        self.con.commit()

    def fechar(self):
        self.con.close()

    def trocar_senha(self, login, nova, provisoria=False):
        if not provisoria:                         # provisoria: gerada pelo programa (redefinir)
            self._checar_senha(nova, login)
        sal = os.urandom(16)
        self.con.execute("UPDATE usuarios SET sal=?, hash=?, provisoria=?, iter=? WHERE login=?",
                         (sal, self._hash(nova, sal), int(provisoria), self.ITERACOES, login))
        self.con.commit()

    def redefinir(self, login):
        """Administrador: gera uma senha provisoria. O usuario sera obrigado a troca-la no proximo acesso."""
        senha = self._aleatorio(10)
        self.trocar_senha(login, senha, provisoria=True)
        return senha

    # ---- codigo de recuperacao ---------------------------------------------

    def gerar_codigo(self, login):
        """Cria um codigo de recuperacao novo (o anterior deixa de valer). Exibir uma unica vez."""
        bruto = self._aleatorio(TAM_CODIGO)
        sal = os.urandom(16)
        self.con.execute("UPDATE usuarios SET rec_sal=?, rec_hash=?, rec_iter=? WHERE login=?",
                         (sal, self._hash(bruto, sal), self.ITERACOES, login))
        self.con.commit()
        return "-".join(bruto[i:i + 4] for i in range(0, TAM_CODIGO, 4))

    def recuperar(self, login, codigo, nova):
        """
        "Esqueci minha senha": com o codigo certo, define a nova senha e devolve um NOVO
        codigo (cada codigo vale uma vez). Codigo ou usuario errado: devolve None.
        """
        login = login.strip().lower()
        self._checar_senha(nova, login)
        cod = _normalizar_codigo(codigo)
        r = self.con.execute("SELECT rec_sal, rec_hash, rec_iter FROM usuarios WHERE login=?", (login,)).fetchone()
        if not r or not r[0] or not r[1]:
            self._hash(cod, b"\0" * 16)          # mesmo custo de tempo: nao revela se o usuario existe
            return None
        if len(cod) != TAM_CODIGO or not hmac.compare_digest(self._hash(cod, r[0], r[2]), r[1]):
            return None
        self.trocar_senha(login, nova, provisoria=False)
        return self.gerar_codigo(login)
