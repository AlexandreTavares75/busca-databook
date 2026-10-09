"""
Busca Databook - motor de leitura, indice e busca.

Le o conteudo dos documentos de uma pasta (e subpastas), guarda o texto em
um indice SQLite FTS5 e faz buscas rapidas, sem diferenciar maiusculas,
minusculas ou acentos ("trafo" encontra "TRAFO", "Trafô", "trafos"...).
"""

import hashlib
import hmac
import io
import logging
import os
import re
import sqlite3
import sys
import threading
import time
import zipfile
from pathlib import Path

log = logging.getLogger(__name__)

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
SENHA_MIN = 6                 # vale para senhas novas; senhas antigas continuam aceitas


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
    """Aceita caminho (str/Path) ou bytes; devolve algo que as libs leem."""
    return io.BytesIO(fonte) if isinstance(fonte, (bytes, bytearray)) else fonte


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
    if isinstance(fonte, (bytes, bytearray)):
        wb = xlrd.open_workbook(file_contents=bytes(fonte))
    else:
        wb = xlrd.open_workbook(str(fonte))
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


def ler_pdf(fonte, usar_ocr=True, progresso=None) -> str:
    """Le o texto do PDF. Paginas sem texto (escaneadas) passam por OCR."""
    import pypdfium2 as pdfium
    dados = bytes(fonte) if isinstance(fonte, (bytes, bytearray)) else str(fonte)
    pdf = pdfium.PdfDocument(dados)
    partes = []
    ocr_feitas = 0
    try:
        for i in range(len(pdf)):
            pag = pdf[i]
            tp = pag.get_textpage()
            texto = tp.get_text_range() or ""
            tp.close()
            if len(texto.strip()) < 20 and usar_ocr and ocr_feitas < OCR_MAX_PAGINAS and configurar_ocr():
                if progresso:
                    progresso(f"OCR página {i + 1}")
                img = pag.render(scale=OCR_DPI / 72).to_pil()
                texto = _ocr_imagem(img)
                ocr_feitas += 1
            pag.close()
            partes.append(f"[Página {i + 1}]\n{texto}")
    finally:
        pdf.close()
    return "\n".join(partes)


def ler_texto(fonte) -> str:
    dados = fonte if isinstance(fonte, (bytes, bytearray)) else Path(fonte).read_bytes()
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
                d = word.Documents.Open(str(fonte), ReadOnly=True, AddToRecentFiles=False)
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
    dados = fonte if isinstance(fonte, (bytes, bytearray)) else Path(fonte).read_bytes()
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


def extrair(nome: str, fonte, usar_ocr=True, progresso=None, profundidade=0):
    """
    Devolve lista de tuplas (subcaminho_no_zip_ou_vazio, texto).
    Para arquivos normais e uma lista com um item; para zip, um por membro.
    """
    ext = os.path.splitext(nome)[1].lower()
    if ext == ".zip":
        return _extrair_zip(fonte, usar_ocr, progresso, profundidade)
    if ext not in EXT_CONTEUDO:
        return [("", "")]
    if ext == ".pdf":
        return [("", ler_pdf(fonte, usar_ocr, progresso))]
    leitor = LEITORES.get(ext, ler_texto)
    return [("", leitor(fonte))]


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

    def empresas(self):
        return listar_empresas(self.raiz)

    # ---- atualizacao incremental ----------------------------------------

    def atualizar(self, usar_ocr=True, progresso=None, cancelado=lambda: False):
        """
        Varre a pasta e (re)indexa so o que mudou.
        progresso(msg, atual, total) e chamado periodicamente.
        Devolve dict com contagens.
        """
        def avisar(msg, i=0, n=0):
            if progresso:
                progresso(msg, i, n)

        avisar("Listando arquivos...")
        no_disco = {}
        for dirpath, dirnames, filenames in os.walk(self.raiz, onerror=lambda e: None):
            if cancelado():
                break
            # ignora pastas ocultas/sistema
            dirnames[:] = [d for d in dirnames if not d.startswith(("~", "$", "."))]
            for f in filenames:
                if f.startswith("~$"):  # arquivo temporario do Office
                    continue
                p = os.path.join(dirpath, f)
                try:
                    st = os.stat(p)
                    no_disco[p] = (st.st_mtime, st.st_size)
                except OSError:
                    pass
            avisar(f"Listando arquivos... {len(no_disco)}")

        if cancelado():
            # varredura incompleta: nao remove nada do indice com base nela
            return {"total": len(no_disco), "lidos": 0, "removidos": 0, "erros": 0, "cancelado": True}

        no_indice = {r[0]: (r[1], r[2]) for r in
                     self.con.execute("SELECT caminho, mtime, tamanho FROM arquivos")}

        removidos = [p for p in no_indice if p not in no_disco]
        novos = [p for p, mt in no_disco.items() if no_indice.get(p) != mt]

        for p in removidos:
            self._remover(p)
        self.con.commit()

        erros = 0
        total = len(novos)
        t_commit = time.time()
        for i, p in enumerate(novos, 1):
            if cancelado():
                break
            nome = os.path.basename(p)
            avisar(f"Lendo {nome}", i, total)
            mtime, tam = no_disco[p]
            status = "ok"
            try:
                if tam > MAX_TAMANHO_MB * 1024 * 1024:
                    partes = [("", "")]
                    status = "grande"
                else:
                    partes = extrair(nome, p, usar_ocr,
                                     lambda m, _n=nome, _i=i: avisar(f"{_n}: {m}", _i, total))
            except Exception as e:
                partes = [("", "")]
                status = f"erro: {type(e).__name__}: {e}"[:300]
                erros += 1
            self._remover(p)
            nome_busca = _nome_pesquisavel(p, self.raiz)
            for interno, texto in partes:
                self.con.execute(
                    "INSERT INTO conteudo (caminho, interno, nome, texto) VALUES (?,?,?,?)",
                    (p, interno, nome_busca, texto))
            self.con.execute("INSERT OR REPLACE INTO arquivos VALUES (?,?,?,?)",
                             (p, mtime, tam, status))
            if time.time() - t_commit > 5:
                self.con.commit()
                t_commit = time.time()
        self.con.commit()
        self.con.execute("INSERT OR REPLACE INTO meta VALUES ('atualizado', ?)",
                         (time.strftime("%d/%m/%Y %H:%M"),))
        self.con.commit()
        return {"total": len(no_disco), "lidos": total, "removidos": len(removidos),
                "erros": erros, "cancelado": cancelado()}

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
        return resultados


def _nome_pesquisavel(caminho, raiz):
    """Nome do arquivo + subpastas relativas, para achar tambem por pasta."""
    rel = os.path.relpath(caminho, raiz)
    base = os.path.splitext(rel)[0]
    # separa palavras grudadas por _ - . para o tokenizador
    return re.sub(r"[_\-.\\/]+", " ", base)


def _limpar(t):
    return re.sub(r"\s+", " ", t or "").strip()


def montar_consulta(termo: str) -> str:
    """Converte o texto digitado em consulta FTS5 segura (sem operadores do usuario)."""
    termo = (termo or "").strip()
    if not termo:
        return ""
    partes = re.findall(r'"[^"]+"|\S+', termo)
    grupos, atual = [], []
    for p in partes:
        if p.upper() in ("OU", "OR"):
            if atual:
                grupos.append(atual)
            atual = []
            continue
        if p.startswith('"') and p.endswith('"') and len(p) > 2:
            frase = " ".join(re.findall(r"\w+", p[1:-1], flags=re.UNICODE))
            if frase:
                atual.append('"' + frase + '"')
        else:
            for tok in re.findall(r"\w+", p, flags=re.UNICODE):
                atual.append('"' + tok + '"*')
    if atual:
        grupos.append(atual)
    grupos = [g for g in grupos if g]
    if not grupos:
        return ""
    return " OR ".join("(" + " AND ".join(g) + ")" for g in grupos)


# --------------------------------------------------------------------------
# Usuarios (login)
# --------------------------------------------------------------------------

class Usuarios:
    ITERACOES = 200_000

    def __init__(self):
        self.con = sqlite3.connect(str(pasta_dados_app() / "usuarios.db"))
        self.con.execute("""CREATE TABLE IF NOT EXISTS usuarios (
            login TEXT PRIMARY KEY, sal BLOB, hash BLOB, admin INTEGER)""")
        self.con.commit()

    @classmethod
    def _hash(cls, senha, sal):
        return hashlib.pbkdf2_hmac("sha256", senha.encode("utf-8"), sal, cls.ITERACOES)

    @staticmethod
    def _checar_senha(senha):
        if len(senha) < SENHA_MIN:
            raise ValueError(f"A senha precisa ter pelo menos {SENHA_MIN} caracteres.")

    def vazio(self):
        return self.con.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0] == 0

    def criar(self, login, senha, admin=False):
        login = login.strip().lower()
        if not login:
            raise ValueError("Informe um nome de usuário.")
        self._checar_senha(senha)
        if self.con.execute("SELECT 1 FROM usuarios WHERE login=?", (login,)).fetchone():
            raise ValueError("Já existe um usuário com esse nome.")
        sal = os.urandom(16)
        self.con.execute("INSERT INTO usuarios VALUES (?,?,?,?)",
                         (login, sal, self._hash(senha, sal), int(admin)))
        self.con.commit()

    def validar(self, login, senha):
        login = login.strip().lower()
        r = self.con.execute("SELECT sal, hash, admin FROM usuarios WHERE login=?", (login,)).fetchone()
        if r:
            if hmac.compare_digest(self._hash(senha, r[0]), r[1]):
                return {"login": login, "admin": bool(r[2])}
        else:
            # mesmo custo de tempo para usuario inexistente (nao revela quais logins existem)
            self._hash(senha, b"\0" * 16)
        return None

    def listar(self):
        return self.con.execute("SELECT login, admin FROM usuarios ORDER BY login").fetchall()

    def remover(self, login):
        self.con.execute("DELETE FROM usuarios WHERE login=?", (login,))
        self.con.commit()

    def trocar_senha(self, login, nova):
        self._checar_senha(nova)
        sal = os.urandom(16)
        self.con.execute("UPDATE usuarios SET sal=?, hash=? WHERE login=?",
                         (sal, self._hash(nova, sal), login))
        self.con.commit()