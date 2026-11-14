"""
Busca Databook - interface para Windows (Tkinter).

Executar a partir do codigo-fonte:  python src/app.py
Instalador: gerado pelo GitHub Actions (ver .github/workflows/release.yml).
"""

import ctypes
import io
import json
import logging
import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from pathlib import Path
import tkinter as tk
from logging.handlers import RotatingFileHandler
from tkinter import ttk, filedialog, messagebox, simpledialog
import tkinter.font as tkfont

import core
import updater
from versao import APP_NOME, REPO_GITHUB, VERSAO

log = logging.getLogger("app")

CONFIG = core.pasta_dados_app() / "config.json"
ARQUIVO_LOG = core.pasta_dados_app() / "erro.log"
ARQUIVO_AUDITORIA = core.pasta_dados_app() / "auditoria.log"
MUTEX_NOME = "BuscaDatabook_Instancia"   # o instalador espera o programa fechar por este nome
INTERVALO_VERIFICACAO_S = 24 * 3600      # verifica atualizacoes no maximo 1x por dia
TENTATIVAS_LOGIN = 5
BLOQUEIO_LOGIN_S = 30
BLOQUEIO_INATIVIDADE_MIN = 15            # padrao; o usuario escolhe 5 a 60 em Opcoes
OPCOES_INATIVIDADE_MIN = (5, 10, 15, 30, 60)

# Menu lateral
LARGURA_MENU = 220
AZUL = "#002060"                 # fundo do menu lateral (muda com a cor do cliente)
AZUL_TEXTO_SUAVE = "#AFC3E8"
BOTAO_COR = "#0B3A8C"            # botao em relevo, um tom acima do fundo
BOTAO_HOVER = "#1A4FA8"
DESTAQUE = "#002060"             # texto do botao selecionado e do "Buscando em"
MARCA = None                     # personalizacao do cliente (core.carregar_marca)


def _misturar(cor, alvo, fator):
    c = [int(cor[i:i + 2], 16) for i in (1, 3, 5)]
    a = [int(alvo[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02X%02X%02X" % tuple(round(x + (y - x) * fator) for x, y in zip(c, a))


def _luminancia(cor):
    def canal(v):
        v /= 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = (int(cor[i:i + 2], 16) for i in (1, 3, 5))
    return 0.2126 * canal(r) + 0.7152 * canal(g) + 0.0722 * canal(b)


def rolagem_lateral(tv):
    """Shift + roda do mouse rola a lista para os lados (a barra horizontal tambem funciona)."""
    def rolar(e):
        tv.xview_scroll(-3 if e.delta > 0 else 3, "units")
        return "break"
    tv.bind("<Shift-MouseWheel>", rolar)


def ajustar_ultima_coluna(tv, col, textos, minimo=200, maximo=4000):
    """Alarga a coluna `col` ate caber o texto mais longo (ex.: o caminho da pasta), para a
    barra de rolagem horizontal mostrar o endereco inteiro. Sem texto longo, a coluna ocupa
    o espaco que sobra na janela."""
    fonte = tkfont.nametofont("TkDefaultFont")
    maior = max((fonte.measure(t) for t in textos[:3000]), default=0) + 24
    outras = sum(tv.column(c, "width") for c in ("#0",) + tuple(tv["columns"]) if c != col)
    sobra = tv.winfo_width() - outras - 4
    tv.column(col, width=max(minimo, min(maximo, max(maior, sobra))), stretch=False)
    tv._largura_conteudo = maior            # usado ao redimensionar a janela


def manter_ultima_coluna(tv, col, minimo=200, maximo=4000):
    """Ao redimensionar a janela, recalcula a coluna `col` (preenche o espaco ou cabe o texto)."""
    def ao_mudar(_e=None):
        outras = sum(tv.column(c, "width") for c in ("#0",) + tuple(tv["columns"]) if c != col)
        sobra = tv.winfo_width() - outras - 4
        tv.column(col, width=max(minimo, min(maximo, max(getattr(tv, "_largura_conteudo", 0), sobra))))
    tv.bind("<Configure>", ao_mudar, add="+")


def aplicar_tema(marca):
    """Cores do menu lateral a partir da cor do cliente. Cores claras (ex. laranja) ganham botoes
    mais escuros para o texto branco continuar legivel; cores escuras, botoes mais claros."""
    global AZUL, AZUL_TEXTO_SUAVE, BOTAO_COR, BOTAO_HOVER, DESTAQUE, MARCA
    MARCA = marca
    cor = (marca or {}).get("cor_menu")
    if not cor:
        AZUL, AZUL_TEXTO_SUAVE, BOTAO_COR, BOTAO_HOVER, DESTAQUE = (
            "#002060", "#AFC3E8", "#0B3A8C", "#1A4FA8", "#002060")
        return
    AZUL = cor
    if _luminancia(cor) > 0.12:                       # cor clara/viva
        BOTAO_COR = _misturar(cor, "#000000", 0.30)
        BOTAO_HOVER = _misturar(cor, "#000000", 0.18)
        AZUL_TEXTO_SUAVE = "#FFFFFF"
        DESTAQUE = BOTAO_COR
    else:                                             # cor escura
        BOTAO_COR = _misturar(cor, "#FFFFFF", 0.12)
        BOTAO_HOVER = _misturar(cor, "#FFFFFF", 0.22)
        AZUL_TEXTO_SUAVE = _misturar(cor, "#FFFFFF", 0.65)
        DESTAQUE = cor
FONTE_MENU = ("Segoe UI", 10)
FONTE_MENU_SEL = ("Segoe UI", 10, "bold")
EMPRESA_DEV = "Omni Solutions Ltda"

FILTROS = {
    "Todos os tipos": None,
    "PDF": {"pdf"},
    "Word": {"doc", "docx", "docm"},
    "Excel": {"xls", "xlsx", "xlsm", "csv"},
    "PowerPoint": {"ppt", "pptx", "pptm"},
    "Imagens": {"png", "jpg", "jpeg", "bmp", "gif", "tif", "tiff"},
    "Desenhos / Projeto": {"dwg", "dwf", "dxf", "mpp"},
}

DICAS = """Como buscar

trafo
    Encontra trafo, Trafo, TRAFO, trafô, trafos...

"trafo 500 kva"
    Entre aspas: busca a frase exata.

trafo bucha
    Documentos que tenham as DUAS palavras.

trafo OU bucha
    Documentos que tenham QUALQUER uma delas.

Procurar no
• "Nome e conteúdo": procura dentro dos documentos e nos nomes.
• "Somente nome do arquivo": procura só no nome do arquivo (com a
  extensão), sem olhar o conteúdo nem as pastas. Ex.: TR 1098830287
  ou "Elera comissionamento" ou relatorio pdf.

Empresas
• O menu azul à esquerda mostra as pastas "DB_NOME" da pasta principal.
• Clique numa empresa para buscar só na pasta dela (e subpastas).
• "Todas" busca na pasta principal inteira.

Cópias e versões
• Documentos com o mesmo nome na mesma empresa aparecem uma vez só:
  a versão mais recente, no caminho mais curto.
• A coluna "Cópias" mostra quantos existem. Botão direito >
  "Ver todas as cópias e versões" lista todos.
• Desmarque "Ocultar cópias e versões antigas" para ver tudo.

O que é lido
• Conteúdo: Word, Excel, PowerPoint, PDF (inclusive escaneado, por OCR),
  textos, e arquivos dentro de .zip.
• Só o nome: imagens (PNG, JPEG...), MPP, DWG, DWF e demais formatos.

Atalhos
• Enter: busca.  Esc: limpa a busca.
• Duplo clique ou Enter na lista: abre o arquivo.
• Botão direito: abrir pasta, salvar cópia, copiar caminho, ver cópias.
• Ctrl+clique / Shift+clique: seleciona vários para salvar cópia de uma vez.
"""


# --------------------------------------------------------------------------
# Log de erros, recursos e permissoes
# --------------------------------------------------------------------------

def configurar_log():
    """Grava erros em %LOCALAPPDATA%\\BuscaDatabook\\erro.log (com rodizio de arquivos)."""
    h = RotatingFileHandler(ARQUIVO_LOG, maxBytes=1_000_000, backupCount=2, encoding="utf-8")
    h.setFormatter(logging.Formatter("%(asctime)s %(levelname)s [%(name)s] %(message)s"))
    logging.basicConfig(level=logging.INFO, handlers=[h], force=True)
    # auditoria (LGPD): arquivo proprio, guardado por mais tempo (~20 MB)
    ha = RotatingFileHandler(ARQUIVO_AUDITORIA, maxBytes=1_000_000, backupCount=20, encoding="utf-8")
    ha.setFormatter(logging.Formatter("%(asctime)s | %(message)s", "%Y-%m-%d %H:%M:%S"))
    core.log_auditoria.handlers[:] = [ha]
    core.log_auditoria.setLevel(logging.INFO)

    def gancho(tipo, valor, tb):
        logging.critical("Erro não tratado", exc_info=(tipo, valor, tb))

    sys.excepthook = gancho
    threading.excepthook = lambda a: logging.critical(
        "Erro em tarefa de fundo", exc_info=(a.exc_type, a.exc_value, a.exc_traceback))


def caminho_recurso(*partes) -> Path:
    """Arquivos embutidos (ex.: logo): no .exe ficam em _MEIPASS; no codigo-fonte, na raiz do projeto."""
    base = getattr(sys, "_MEIPASS", None) or Path(__file__).resolve().parent.parent
    return Path(base).joinpath(*partes)


def carregar_logo(largura_max, altura_max):
    """Carrega assets/logo.png, corta as bordas transparentes e reduz. None se falhar."""
    try:
        from PIL import Image, ImageTk
        img = Image.open(caminho_recurso("assets", "logo.png")).convert("RGBA")
        caixa = img.getchannel("A").getbbox()
        if caixa:
            img = img.crop(caixa)
        img.thumbnail((largura_max, altura_max), getattr(Image, "Resampling", Image).LANCZOS)
        return ImageTk.PhotoImage(img)
    except Exception:
        log.warning("Logomarca não carregada", exc_info=True)
        return None


def aplicar_icone(janela):
    """Icone da barra de titulo e da barra de tarefas (vale para todas as janelas do app)."""
    ico = caminho_recurso("assets", "icone.ico")
    if sys.platform.startswith("win") and ico.is_file():
        try:
            janela.iconbitmap(default=str(ico))
        except tk.TclError:
            log.debug("Ícone não aplicado", exc_info=True)


def empresas_permitidas(usuario, empresas):
    """
    Ponto unico de controle de acesso por empresa.
    Hoje todos os usuarios veem todas as empresas. Para restringir no futuro,
    filtre aqui a lista [(nome, caminho), ...] conforme o usuario.
    """
    return empresas


# --------------------------------------------------------------------------
# Utilitarios
# --------------------------------------------------------------------------

def carregar_config():
    try:
        return json.loads(CONFIG.read_text(encoding="utf-8"))
    except Exception:
        return {}


def salvar_config(cfg):
    """Gravacao atomica: nunca deixa o config.json pela metade."""
    try:
        tmp = CONFIG.with_suffix(".tmp")
        tmp.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, CONFIG)
    except Exception:
        log.warning("Não foi possível salvar a configuração", exc_info=True)


def abrir_no_sistema(caminho):
    if sys.platform.startswith("win"):
        os.startfile(caminho)  # noqa
    elif sys.platform == "darwin":
        subprocess.Popen(["open", caminho])
    else:
        subprocess.Popen(["xdg-open", caminho])


def mostrar_na_pasta(caminho):
    if sys.platform.startswith("win"):
        subprocess.Popen(["explorer", "/select,", os.path.normpath(caminho)])
    else:
        abrir_no_sistema(os.path.dirname(caminho))


def extrair_do_zip(caminho_zip, interno, destino_dir):
    """Extrai um arquivo de dentro do zip (inclusive zip dentro de zip).
    Usa so o nome-base do membro: um zip malicioso nao consegue gravar fora de destino_dir."""
    partes = interno.split(core.SEP_ZIP)
    dados = None
    zf = zipfile.ZipFile(core.caminho_longo(caminho_zip))
    try:
        for i, parte in enumerate(partes):
            membro = next((m for m in zf.infolist() if core._corrigir_nome_zip(m) == parte), None)
            if membro is None:
                raise FileNotFoundError(parte)
            dados = zf.read(membro)
            if i < len(partes) - 1:
                zf.close()
                zf = zipfile.ZipFile(io.BytesIO(dados))
    finally:
        zf.close()
    nome = os.path.basename(partes[-1].replace("\\", "/")) or "arquivo"
    destino = os.path.join(destino_dir, nome)
    with open(destino, "wb") as f:
        f.write(dados)
    return destino


def formatar_data(mtime):
    return time.strftime("%d/%m/%Y %H:%M", time.localtime(mtime)) if mtime else ""


def formatar_tamanho(n):
    if n >= 1024 * 1024:
        return f"{n / 1024 / 1024:.1f} MB".replace(".", ",")
    return f"{max(1, round(n / 1024))} KB"


# --------------------------------------------------------------------------
# Login
# --------------------------------------------------------------------------

class TelaLogin(tk.Toplevel):
    def __init__(self, master, usuarios):
        super().__init__(master)
        self.usuarios = usuarios
        self.resultado = None
        self.falhas = 0
        self.bloqueado_ate = 0.0
        self.primeiro_acesso = usuarios.vazio()
        self.title(APP_NOME + " - Acesso")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.cancelar)

        f = ttk.Frame(self, padding=20)
        f.pack(fill="both", expand=True)
        titulo = "Primeiro acesso: crie o usuário administrador" if self.primeiro_acesso else "Entre com seu usuário"
        ttk.Label(f, text=APP_NOME, font=("Segoe UI", 14, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 4))
        ttk.Label(f, text=titulo).grid(row=1, column=0, columnspan=2, pady=(0, 12))

        ttk.Label(f, text="Usuário:").grid(row=2, column=0, sticky="e", padx=4, pady=4)
        self.e_user = ttk.Entry(f, width=28)
        self.e_user.grid(row=2, column=1, pady=4)
        ttk.Label(f, text="Senha:").grid(row=3, column=0, sticky="e", padx=4, pady=4)
        self.e_senha = ttk.Entry(f, width=28, show="•")
        self.e_senha.grid(row=3, column=1, pady=4)
        linha = 4
        if self.primeiro_acesso:
            ttk.Label(f, text="Confirmar senha:").grid(row=4, column=0, sticky="e", padx=4, pady=4)
            self.e_senha2 = ttk.Entry(f, width=28, show="•")
            self.e_senha2.grid(row=4, column=1, pady=4)
            ttk.Label(f, text=core.REGRAS_SENHA, foreground="#666", justify="left").grid(
                row=5, column=0, columnspan=2, sticky="w", pady=(4, 0))
            linha = 6

        b = ttk.Frame(f)
        b.grid(row=linha, column=0, columnspan=2, pady=(12, 0))
        ttk.Button(b, text="Entrar", command=self.entrar).pack(side="left", padx=4)
        ttk.Button(b, text="Sair", command=self.cancelar).pack(side="left", padx=4)
        if not self.primeiro_acesso:
            esqueci = ttk.Label(f, text="Esqueci minha senha", foreground="#0645AD", cursor="hand2",
                                font=("Segoe UI", 9, "underline"))
            esqueci.grid(row=linha + 1, column=0, columnspan=2, pady=(10, 0))
            esqueci.bind("<Button-1>", lambda e: self.esqueci())
            linha += 1
        ttk.Label(f, text=f"versão {VERSAO}", foreground="#888").grid(row=linha + 1, column=0, columnspan=2, pady=(10, 0))

        self.bind("<Return>", lambda e: self.entrar())
        # Sem transient(master): a janela principal fica oculta ate o login, e uma
        # janela "transient" de uma janela oculta nao aparece no Windows.
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 3
        self.geometry(f"+{x}+{y}")
        self.lift()
        self.attributes("-topmost", True)
        self.after(400, lambda: self.attributes("-topmost", False))
        self.focus_force()
        self.e_user.focus_set()
        try:
            self.grab_set()
        except tk.TclError:          # janela ainda nao visivel: tenta de novo em seguida
            self.after(100, self.grab_set)

    def bloqueado(self, parent=None):
        """True (e avisa) se ainda esta no tempo de espera apos muitas tentativas erradas."""
        restante = self.bloqueado_ate - time.monotonic()
        if restante > 0:
            messagebox.showwarning(APP_NOME, f"Muitas tentativas. Aguarde {int(restante) + 1} segundos.",
                                   parent=parent or self)
            return True
        return False

    def registrar_falha(self):
        """Conta tentativas erradas (senha ou codigo). Devolve True se acabou de bloquear."""
        self.falhas += 1
        if self.falhas >= TENTATIVAS_LOGIN:
            self.falhas = 0
            self.bloqueado_ate = time.monotonic() + BLOQUEIO_LOGIN_S
            return True
        return False

    def entrar(self):
        if self.bloqueado():
            return
        user, senha = self.e_user.get(), self.e_senha.get()
        if self.primeiro_acesso:
            if senha != self.e_senha2.get():
                messagebox.showerror(APP_NOME, "As senhas não conferem.", parent=self)
                return
            try:
                self.usuarios.criar(user, senha, admin=True)
                log.info("Administrador inicial criado: %s", user.strip().lower())
                core.auditar(user.strip().lower(), "usuário criado", "administrador inicial")
            except Exception as e:
                messagebox.showerror(APP_NOME, str(e), parent=self)
                return
        r = self.usuarios.validar(user, senha)
        if not r:
            log.warning("Falha de login (usuário informado: %r)", user.strip().lower()[:64])
            core.auditar(user.strip().lower()[:64], "falha de login")
            self.e_senha.delete(0, "end")
            extra = (f"\n\nMuitas tentativas: aguarde {BLOQUEIO_LOGIN_S} segundos."
                     if self.registrar_falha() else "")
            messagebox.showerror(APP_NOME, "Usuário ou senha incorretos." + extra, parent=self)
            return
        self.falhas = 0

        if r["provisoria"]:
            # senha provisoria (definida pelo administrador): troca obrigatoria antes de entrar
            tela = TelaNovaSenha(self, self.usuarios, r["login"], obrigatoria=True)
            self.wait_window(tela)
            if not tela.ok:
                self.e_senha.delete(0, "end")
                return
            r["provisoria"] = False
        elif not core.senha_forte(senha, r["login"]):
            # senha criada antes da regra atual (3.4): troca obrigatoria antes de entrar
            tela = TelaNovaSenha(self, self.usuarios, r["login"], obrigatoria=True,
                                 motivo="Sua senha não atende às novas regras de segurança.")
            self.wait_window(tela)
            if not tela.ok:
                self.e_senha.delete(0, "end")
                return
            log.info("Senha fora da regra trocada no login: %s", r["login"])
            core.auditar(r["login"], "senha trocada", "senha antiga fora da regra")
        if not r["tem_codigo"]:
            # todo usuario recebe um codigo de recuperacao (inclusive os que ja existiam)
            mostrar_codigo(self, self.usuarios.gerar_codigo(r["login"]))
            log.info("Código de recuperação gerado para %s", r["login"])

        log.info("Login: %s", r["login"])
        core.auditar(r["login"], "login")
        self.resultado = r
        self.destroy()

    def esqueci(self):
        if self.bloqueado():
            return
        TelaRecuperar(self, self.usuarios, self.e_user.get())

    def cancelar(self):
        self.resultado = None
        self.destroy()


class TelaNovaSenha(tk.Toplevel):
    """Cria uma senha nova. obrigatoria=True: primeiro acesso com senha provisoria."""

    def __init__(self, master, usuarios, login, obrigatoria=False, motivo=None):
        super().__init__(master)
        self.usuarios, self.login, self.ok = usuarios, login, False
        self.title("Criar nova senha")
        self.resizable(False, False)
        f = ttk.Frame(self, padding=18)
        f.pack(fill="both", expand=True)
        if motivo:
            msg = f"Olá, {login}. {motivo}\nCrie agora uma nova senha para continuar."
        elif obrigatoria:
            msg = f"Olá, {login}. Sua senha atual é provisória.\nCrie agora a sua senha pessoal para continuar."
        else:
            msg = f"Nova senha para {login}."
        ttk.Label(f, text=msg, justify="left").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(f, text="Nova senha:").grid(row=1, column=0, sticky="e", padx=4, pady=4)
        self.e1 = ttk.Entry(f, width=28, show="•")
        self.e1.grid(row=1, column=1, pady=4)
        ttk.Label(f, text="Repita a senha:").grid(row=2, column=0, sticky="e", padx=4, pady=4)
        self.e2 = ttk.Entry(f, width=28, show="•")
        self.e2.grid(row=2, column=1, pady=4)
        ttk.Label(f, text=core.REGRAS_SENHA, foreground="#666", justify="left").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(4, 0))
        b = ttk.Frame(f)
        b.grid(row=4, column=0, columnspan=2, pady=(12, 0))
        ttk.Button(b, text="Salvar", command=self.salvar).pack(side="left", padx=4)
        ttk.Button(b, text="Cancelar", command=self.destroy).pack(side="left", padx=4)
        self.bind("<Return>", lambda e: self.salvar())
        self.transient(master)
        self.e1.focus_set()
        self.grab_set()

    def salvar(self):
        nova = self.e1.get()
        if nova != self.e2.get():
            messagebox.showerror(APP_NOME, "As senhas não conferem.", parent=self)
            return
        if self.usuarios.validar(self.login, nova):
            messagebox.showerror(APP_NOME, "A nova senha precisa ser diferente da atual.", parent=self)
            return
        try:
            self.usuarios.trocar_senha(self.login, nova)
        except ValueError as e:
            messagebox.showerror(APP_NOME, str(e), parent=self)
            return
        log.info("Usuário %s definiu nova senha", self.login)
        core.auditar(self.login, "senha alterada")
        self.ok = True
        self.destroy()


class TelaSegredo(tk.Toplevel):
    """Mostra uma unica vez um codigo de recuperacao ou uma senha provisoria."""

    def __init__(self, master, titulo, valor, texto, exigir_confirmacao):
        super().__init__(master)
        self.title(titulo)
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.fechar)
        f = ttk.Frame(self, padding=18)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text=texto, justify="left", wraplength=420).pack(anchor="w")
        caixa = tk.Entry(f, font=("Consolas", 18, "bold"), justify="center", width=18,
                         relief="solid", bd=1, readonlybackground="#F2F5FB", fg=DESTAQUE)
        caixa.insert(0, valor)
        caixa.config(state="readonly")
        caixa.pack(pady=14, ipady=6)
        ttk.Button(f, text="Copiar", command=lambda: (self.clipboard_clear(), self.clipboard_append(valor))
                   ).pack()
        self.var_ok = tk.BooleanVar(value=not exigir_confirmacao)
        if exigir_confirmacao:
            ttk.Checkbutton(f, text="Guardei o código em local seguro", variable=self.var_ok,
                            command=lambda: self.bt.config(state="normal" if self.var_ok.get() else "disabled")
                            ).pack(pady=(14, 0))
        self.bt = ttk.Button(f, text="Fechar", command=self.fechar,
                             state="normal" if not exigir_confirmacao else "disabled")
        self.bt.pack(pady=(10, 0))
        self.transient(master)
        self.grab_set()
        self.focus_force()

    def fechar(self):
        if self.var_ok.get():
            try:
                self.clipboard_clear()   # nao deixa o segredo esquecido na area de transferencia
            except tk.TclError:
                pass
            self.destroy()


def mostrar_codigo(master, codigo):
    tela = TelaSegredo(
        master, "Código de recuperação", codigo,
        "Este é o seu código de recuperação. Se você esquecer a senha, use-o em "
        "“Esqueci minha senha” na tela de acesso para criar uma nova.\n\n"
        "• Anote em papel ou guarde no seu gerenciador de senhas (não em um arquivo neste computador).\n"
        "• Ele aparece só agora e vale uma única vez: depois de usado, você recebe outro.\n"
        "• Não compartilhe com ninguém.", exigir_confirmacao=True)
    master.wait_window(tela)


class TelaRecuperar(tk.Toplevel):
    """'Esqueci minha senha': usuario + codigo de recuperacao + nova senha."""

    def __init__(self, login_tela, usuarios, login_sugerido=""):
        super().__init__(login_tela)
        self.login_tela, self.usuarios = login_tela, usuarios
        self.title("Esqueci minha senha")
        self.resizable(False, False)
        f = ttk.Frame(self, padding=18)
        f.pack(fill="both", expand=True)
        ttk.Label(f, justify="left", wraplength=360,
                  text="Informe o seu código de recuperação para criar uma nova senha.\n"
                       "Não tem o código? Peça ao administrador para redefinir a sua senha."
                  ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        campos = (("Usuário:", ""), ("Código:", ""), ("Nova senha:", "•"), ("Repita a senha:", "•"))
        self.entradas = []
        for i, (rotulo, mascara) in enumerate(campos, 1):
            ttk.Label(f, text=rotulo).grid(row=i, column=0, sticky="e", padx=4, pady=4)
            e = ttk.Entry(f, width=28, show=mascara)
            e.grid(row=i, column=1, pady=4)
            self.entradas.append(e)
        self.entradas[0].insert(0, login_sugerido.strip())
        ttk.Label(f, text="Código no formato XXXX-XXXX-XXXX.\n\nNova senha: " + core.REGRAS_SENHA,
                  foreground="#666", justify="left").grid(row=5, column=0, columnspan=2, sticky="w", pady=(4, 0))
        b = ttk.Frame(f)
        b.grid(row=6, column=0, columnspan=2, pady=(12, 0))
        ttk.Button(b, text="Criar nova senha", command=self.confirmar).pack(side="left", padx=4)
        ttk.Button(b, text="Cancelar", command=self.destroy).pack(side="left", padx=4)
        self.bind("<Return>", lambda e: self.confirmar())
        self.transient(login_tela)
        (self.entradas[1] if login_sugerido.strip() else self.entradas[0]).focus_set()
        self.grab_set()

    def confirmar(self):
        if self.login_tela.bloqueado(parent=self):
            return
        login, codigo, nova, nova2 = (e.get() for e in self.entradas)
        if nova != nova2:
            messagebox.showerror(APP_NOME, "As senhas não conferem.", parent=self)
            return
        try:
            novo_codigo = self.usuarios.recuperar(login, codigo, nova)
        except ValueError as e:
            messagebox.showerror(APP_NOME, str(e), parent=self)
            return
        if not novo_codigo:
            log.warning("Recuperação de senha recusada (usuário informado: %r)", login.strip().lower()[:64])
            core.auditar(login.strip().lower()[:64], "recuperação de senha recusada")
            extra = (f"\n\nMuitas tentativas: aguarde {BLOQUEIO_LOGIN_S} segundos."
                     if self.login_tela.registrar_falha() else "")
            messagebox.showerror(APP_NOME, "Usuário ou código de recuperação incorretos.\n"
                                           "Se não tiver o código, peça ao administrador para "
                                           "redefinir a sua senha." + extra, parent=self)
            return
        log.info("Senha recuperada com código: %s", login.strip().lower())
        core.auditar(login.strip().lower(), "senha recuperada com código")
        messagebox.showinfo(APP_NOME, "Senha alterada!\n\nO código usado deixou de valer. "
                                      "Na próxima tela está o seu NOVO código de recuperação.", parent=self)
        mostrar_codigo(self, novo_codigo)
        self.login_tela.e_user.delete(0, "end")
        self.login_tela.e_user.insert(0, login.strip())
        self.login_tela.e_senha.delete(0, "end")
        self.login_tela.e_senha.focus_set()
        self.destroy()


class TelaUsuarios(tk.Toplevel):
    def __init__(self, master, usuarios, usuario_atual):
        super().__init__(master)
        self.usuarios = usuarios
        self.atual = usuario_atual
        self.title("Usuários")
        self.resizable(False, False)
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        self.lista = ttk.Treeview(f, columns=("perfil", "situacao"), height=8)
        self.lista.heading("#0", text="Usuário")
        self.lista.heading("perfil", text="Perfil")
        self.lista.heading("situacao", text="Senha")
        self.lista.column("#0", width=170)
        self.lista.column("perfil", width=110)
        self.lista.column("situacao", width=150)
        self.lista.grid(row=0, column=0, rowspan=4, padx=(0, 8))
        ttk.Button(f, text="Novo usuário", command=self.novo).grid(row=0, column=1, sticky="ew", pady=2)
        ttk.Button(f, text="Redefinir senha", command=self.redefinir).grid(row=1, column=1, sticky="ew", pady=2)
        ttk.Button(f, text="Excluir", command=self.excluir).grid(row=2, column=1, sticky="ew", pady=2)
        ttk.Button(f, text="Fechar", command=self.destroy).grid(row=3, column=1, sticky="sew", pady=2)
        self.recarregar()
        self.transient(master)
        self.grab_set()

    def recarregar(self):
        self.lista.delete(*self.lista.get_children())
        for login, admin, provisoria in self.usuarios.listar():
            self.lista.insert("", "end", iid=login, text=login,
                              values=("Administrador" if admin else "Usuário",
                                      "provisória (trocar no acesso)" if provisoria else "definida pelo usuário"))

    def selecionado(self):
        s = self.lista.selection()
        return s[0] if s else None

    def _mostrar_provisoria(self, login, senha):
        tela = TelaSegredo(
            self, "Senha provisória", senha,
            f"Senha provisória de “{login}”.\n\nInforme-a ao usuário pessoalmente ou por um canal "
            "seguro. No próximo acesso ele será obrigado a criar a própria senha, que você não "
            "conhecerá.", exigir_confirmacao=False)
        self.wait_window(tela)

    def novo(self):
        login = simpledialog.askstring("Novo usuário", "Nome de usuário:", parent=self)
        if not login:
            return
        admin = messagebox.askyesno("Novo usuário", "Este usuário será administrador?", parent=self)
        try:
            senha = self.usuarios.criar(login, None, admin)       # gera senha provisoria
            log.info("Usuário criado por %s: %s (admin=%s)", self.atual, login.strip().lower(), admin)
            core.auditar(self.atual, "usuário criado", f"{login.strip().lower()} (administrador: {'sim' if admin else 'não'})")
        except Exception as e:
            messagebox.showerror("Usuários", f"Não foi possível criar: {e}", parent=self)
            return
        self.recarregar()
        self._mostrar_provisoria(login.strip().lower(), senha)

    def redefinir(self):
        s = self.selecionado()
        if not s:
            messagebox.showinfo("Usuários", "Selecione um usuário na lista.", parent=self)
            return
        if s == self.atual:
            messagebox.showinfo("Usuários", "Para a sua própria senha use Minha conta > Trocar minha senha.",
                                parent=self)
            return
        if not messagebox.askyesno("Usuários", f"Redefinir a senha de {s}?\n\nA senha atual deixará de "
                                               "funcionar e será gerada uma senha provisória.", parent=self):
            return
        senha = self.usuarios.redefinir(s)
        log.info("Senha de %s redefinida (provisória) por %s", s, self.atual)
        core.auditar(self.atual, "senha redefinida (provisória)", s)
        self.recarregar()
        self._mostrar_provisoria(s, senha)

    def excluir(self):
        s = self.selecionado()
        if not s:
            return
        if s == self.atual:
            messagebox.showwarning("Usuários", "Você não pode excluir o próprio usuário.", parent=self)
            return
        if messagebox.askyesno("Usuários", f"Excluir o usuário {s}?", parent=self):
            self.usuarios.remover(s)
            log.info("Usuário %s excluído por %s", s, self.atual)
            core.auditar(self.atual, "usuário excluído", s)
            self.recarregar()


# --------------------------------------------------------------------------
# Exportar indice / cópias de um documento
# --------------------------------------------------------------------------

class TelaExportar(tk.Toplevel):
    """Escolha das empresas a exportar. resultado = (pastas ou None para tudo, incluir_fora)."""

    def __init__(self, master, empresas):
        super().__init__(master)
        self.empresas = empresas
        self.resultado = None
        self.title("Exportar índice")
        self.resizable(False, False)
        f = ttk.Frame(self, padding=14)
        f.pack(fill="both", expand=True)

        ttk.Label(f, text="Empresas a exportar:", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        caixa = ttk.Frame(f)
        caixa.grid(row=1, column=0, sticky="nsew", pady=(4, 6))
        self.lista = tk.Listbox(caixa, selectmode="multiple", height=min(14, max(4, len(empresas))),
                                width=46, exportselection=False, activestyle="none")
        sb = ttk.Scrollbar(caixa, orient="vertical", command=self.lista.yview)
        self.lista.configure(yscrollcommand=sb.set)
        self.lista.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        for nome, _ in empresas:
            self.lista.insert("end", nome)
        self.lista.select_set(0, "end")

        bts = ttk.Frame(f)
        bts.grid(row=2, column=0, sticky="w")
        ttk.Button(bts, text="Marcar todas", command=lambda: self.lista.select_set(0, "end")).pack(side="left")
        ttk.Button(bts, text="Desmarcar todas",
                   command=lambda: self.lista.select_clear(0, "end")).pack(side="left", padx=6)

        self.var_fora = tk.BooleanVar(value=True)
        ttk.Checkbutton(f, text='Incluir arquivos fora das pastas "DB_"',
                        variable=self.var_fora).grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Label(f, foreground="#a0522d", wraplength=380, justify="left",
                  text="⚠ O arquivo exportado contém o texto dos documentos. Envie só por um "
                       "canal privado e só para quem tem direito a esse conteúdo.").grid(
            row=4, column=0, sticky="w", pady=(10, 0))

        fim = ttk.Frame(f)
        fim.grid(row=5, column=0, sticky="e", pady=(12, 0))
        ttk.Button(fim, text="Exportar...", command=self.confirmar).pack(side="left", padx=4)
        ttk.Button(fim, text="Cancelar", command=self.destroy).pack(side="left")
        self.transient(master)
        self.grab_set()

    def confirmar(self):
        sel = list(self.lista.curselection())
        fora = self.var_fora.get()
        if not sel and not fora:
            messagebox.showwarning("Exportar índice", "Escolha ao menos uma empresa.", parent=self)
            return
        if len(sel) == len(self.empresas) and fora:
            self.resultado = (None, True)                       # indice inteiro
        else:
            self.resultado = ([self.empresas[i][1] for i in sel], fora)
        self.destroy()


class TelaSenhaArquivo(tk.Toplevel):
    """Senha de um indice exportado / pacote. criar=True: definir (com confirmacao e "Gerar senha");
    criar=False: digitar para abrir. resultado: a senha ou None."""

    def __init__(self, master, criar, titulo, texto=""):
        super().__init__(master)
        self.criar = criar
        self.resultado = None
        self.title(titulo)
        self.resizable(False, False)
        f = ttk.Frame(self, padding=16)
        f.pack(fill="both", expand=True)
        if texto:
            ttk.Label(f, text=texto, wraplength=420, justify="left").grid(row=0, column=0, columnspan=3,
                                                                         sticky="w", pady=(0, 10))
        self.var_ver = tk.BooleanVar(value=False)
        ttk.Label(f, text="Senha:").grid(row=1, column=0, sticky="e", padx=4, pady=3)
        self.e1 = ttk.Entry(f, width=34, show="•")
        self.e1.grid(row=1, column=1, pady=3, sticky="w")
        if criar:
            ttk.Button(f, text="Gerar senha", command=self._gerar).grid(row=1, column=2, padx=(6, 0))
            ttk.Label(f, text="Confirmar:").grid(row=2, column=0, sticky="e", padx=4, pady=3)
            self.e2 = ttk.Entry(f, width=34, show="•")
            self.e2.grid(row=2, column=1, pady=3, sticky="w")
            ttk.Button(f, text="Copiar", command=self._copiar).grid(row=2, column=2, padx=(6, 0))
            ttk.Label(f, foreground="#a0522d", wraplength=420, justify="left",
                      text=f"Mínimo de {core.SENHA_ARQUIVO_MIN} caracteres. Anote a senha: sem ela o arquivo "
                           "não pode ser aberto. Envie-a por um canal DIFERENTE do arquivo (ex.: telefone "
                           "ou mensagem), nunca no mesmo e-mail.").grid(row=3, column=0, columnspan=3,
                                                                        sticky="w", pady=(8, 0))
        ttk.Checkbutton(f, text="Mostrar senha", variable=self.var_ver, command=self._ver).grid(
            row=4, column=1, sticky="w", pady=(6, 0))
        b = ttk.Frame(f)
        b.grid(row=5, column=0, columnspan=3, sticky="e", pady=(12, 0))
        ttk.Button(b, text="OK", command=self._ok).pack(side="left", padx=4)
        ttk.Button(b, text="Cancelar", command=self.destroy).pack(side="left")
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        try:
            if master.winfo_viewable():
                self.transient(master)
        except tk.TclError:
            pass
        self.lift()
        self.e1.focus_set()
        self.grab_set()

    def _ver(self):
        mostrar = "" if self.var_ver.get() else "•"
        self.e1.config(show=mostrar)
        if self.criar:
            self.e2.config(show=mostrar)

    def _gerar(self):
        s = core.gerar_senha_arquivo()
        for e in (self.e1, self.e2):
            e.delete(0, "end")
            e.insert(0, s)
        self.var_ver.set(True)
        self._ver()

    def _copiar(self):
        self.clipboard_clear()
        self.clipboard_append(self.e1.get())

    def _ok(self):
        s = self.e1.get()
        if self.criar:
            if s != self.e2.get():
                messagebox.showerror(self.title(), "As senhas não conferem.", parent=self)
                return
            try:
                core.checar_senha_arquivo(s)
            except ValueError as e:
                messagebox.showerror(self.title(), str(e), parent=self)
                return
        elif not s:
            return
        self.resultado = s
        self.destroy()


class TelaProtecao(tk.Toplevel):
    """Regras de protecao de dados pessoais (LGPD) da pasta principal. resultado: lista ou None."""

    def __init__(self, master, regras):
        super().__init__(master)
        self.resultado = None
        self.regras = [dict(r) for r in regras]
        self.title("Proteção de dados pessoais (LGPD)")
        self.resizable(False, False)
        f = ttk.Frame(self, padding=14)
        f.pack(fill="both", expand=True)
        ttk.Label(f, wraplength=560, justify="left", text=(
            "Pastas ou arquivos cujo caminho contém um destes termos (palavra inteira, sem diferenciar "
            "acentos e maiúsculas) recebem proteção especial. Use para documentos com dados pessoais "
            "sensíveis, como atestados de saúde (ASO) e exames.\n\n"
            "• Só o nome: o arquivo aparece na busca pelo nome, mas o conteúdo não é lido.\n"
            "• Ignorar: o arquivo não entra no índice.\n\n"
            "As regras valem para esta pasta principal e viajam junto com o índice exportado ou o "
            "pacote do cliente.")).grid(row=0, column=0, columnspan=4, sticky="w")
        self.tv = ttk.Treeview(f, columns=("modo",), height=9, selectmode="extended")
        self.tv.heading("#0", text="Termo")
        self.tv.heading("modo", text="Proteção")
        self.tv.column("#0", width=230)
        self.tv.column("modo", width=300)
        self.tv.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(10, 4))
        self.var_termo = tk.StringVar()
        ttk.Entry(f, textvariable=self.var_termo, width=28).grid(row=2, column=0, sticky="w")
        self.var_modo = tk.StringVar(value=core.MODOS_PROTECAO[core.MODO_SO_NOME])
        ttk.Combobox(f, textvariable=self.var_modo, state="readonly", width=32,
                     values=list(core.MODOS_PROTECAO.values())).grid(row=2, column=1, sticky="w", padx=6)
        ttk.Button(f, text="Adicionar", command=self._adicionar).grid(row=2, column=2, sticky="w")
        ttk.Button(f, text="Remover selecionadas", command=self._remover).grid(row=2, column=3, sticky="e")
        ttk.Label(f, text="Sugestões (clique para adicionar como “Só o nome”):").grid(
            row=3, column=0, columnspan=4, sticky="w", pady=(10, 2))
        sug = ttk.Frame(f)
        sug.grid(row=4, column=0, columnspan=4, sticky="w")
        for i, t in enumerate(core.SUGESTOES_PROTECAO):
            ttk.Button(sug, text=t, width=len(t) + 2,
                       command=lambda t=t: self._incluir(t, core.MODO_SO_NOME)).grid(row=i // 7, column=i % 7,
                                                                                    padx=2, pady=2, sticky="w")
        b = ttk.Frame(f)
        b.grid(row=5, column=0, columnspan=4, sticky="e", pady=(12, 0))
        ttk.Button(b, text="Salvar e aplicar", command=self._salvar).pack(side="left", padx=4)
        ttk.Button(b, text="Cancelar", command=self.destroy).pack(side="left")
        self._recarregar()
        self.transient(master)
        self.grab_set()

    def _recarregar(self):
        self.tv.delete(*self.tv.get_children())
        for i, r in enumerate(self.regras):
            self.tv.insert("", "end", iid=str(i), text=r["termo"], values=(core.MODOS_PROTECAO[r["modo"]],))

    def _incluir(self, termo, modo):
        try:
            self.regras = core.validar_regras(self.regras + [{"termo": termo, "modo": modo}])
        except ValueError as e:
            messagebox.showerror(self.title(), str(e), parent=self)
            return False
        self._recarregar()
        return True

    def _adicionar(self):
        modo = next(k for k, v in core.MODOS_PROTECAO.items() if v == self.var_modo.get())
        if self._incluir(self.var_termo.get(), modo):
            self.var_termo.set("")

    def _remover(self):
        fora = {int(i) for i in self.tv.selection()}
        self.regras = [r for i, r in enumerate(self.regras) if i not in fora]
        self._recarregar()

    def _salvar(self):
        self.resultado = self.regras
        self.destroy()


class TelaTexto(tk.Toplevel):
    """Janela de leitura (aviso de privacidade, registro de auditoria), com "Salvar cópia" opcional."""

    def __init__(self, master, titulo, texto, salvar_como=None):
        super().__init__(master)
        self.title(titulo)
        self.geometry("760x560")
        f = ttk.Frame(self, padding=10)
        f.pack(fill="both", expand=True)
        t = tk.Text(f, wrap="word", font=("Segoe UI", 10), padx=8, pady=8)
        sb = ttk.Scrollbar(f, orient="vertical", command=t.yview)
        t.configure(yscrollcommand=sb.set)
        t.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        t.insert("1.0", texto)
        t.config(state="disabled")
        b = ttk.Frame(self, padding=(10, 0, 10, 10))
        b.pack(fill="x")
        if salvar_como:
            ttk.Button(b, text="Salvar cópia...", command=lambda: self._salvar(texto, salvar_como)).pack(side="left")
        ttk.Button(b, text="Fechar", command=self.destroy).pack(side="right")
        self.transient(master)

    def _salvar(self, texto, nome):
        destino = filedialog.asksaveasfilename(parent=self, initialfile=nome, defaultextension=".txt",
                                               filetypes=[("Texto", "*.txt")])
        if destino:
            Path(destino).write_text(texto, encoding="utf-8")


class TelaBloqueio(tk.Toplevel):
    """Sessao bloqueada por inatividade: so a senha do proprio usuario desbloqueia.
    As tarefas em andamento (ex.: indexacao) continuam."""

    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.falhas = 0
        self.bloqueado_ate = 0.0
        self.title(APP_NOME + " - Sessão bloqueada")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", lambda: None)
        f = ttk.Frame(self, padding=20)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Sessão bloqueada", font=("Segoe UI", 12, "bold")).grid(
            row=0, column=0, columnspan=2, pady=(0, 6))
        ttk.Label(f, text=f"Usuário: {app.usuario['login']}").grid(row=1, column=0, columnspan=2, pady=(0, 10))
        ttk.Label(f, text="Senha:").grid(row=2, column=0, sticky="e", padx=4)
        self.e = ttk.Entry(f, width=28, show="•")
        self.e.grid(row=2, column=1, pady=4)
        b = ttk.Frame(f)
        b.grid(row=3, column=0, columnspan=2, pady=(12, 0))
        ttk.Button(b, text="Desbloquear", command=self._desbloquear).pack(side="left", padx=4)
        ttk.Button(b, text="Sair do programa", command=self._sair).pack(side="left", padx=4)
        self.bind("<Return>", lambda e: self._desbloquear())
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 3
        self.geometry(f"+{x}+{y}")
        self.lift()
        self.e.focus_force()

    def _desbloquear(self):
        restante = self.bloqueado_ate - time.monotonic()
        if restante > 0:
            messagebox.showwarning(APP_NOME, f"Muitas tentativas. Aguarde {int(restante) + 1} segundos.", parent=self)
            return
        if self.app.usuarios.validar(self.app.usuario["login"], self.e.get()):
            self.destroy()
            self.app.desbloquear()
            return
        self.e.delete(0, "end")
        core.auditar(self.app.usuario["login"], "falha ao desbloquear a sessão")
        self.falhas += 1
        extra = ""
        if self.falhas >= TENTATIVAS_LOGIN:
            self.falhas = 0
            self.bloqueado_ate = time.monotonic() + BLOQUEIO_LOGIN_S
            extra = f"\n\nMuitas tentativas: aguarde {BLOQUEIO_LOGIN_S} segundos."
        messagebox.showerror(APP_NOME, "Senha incorreta." + extra, parent=self)

    def _sair(self):
        if self.app.trabalhando and not messagebox.askyesno(
                APP_NOME, "Há uma tarefa em andamento. Sair mesmo assim?\n(O que já foi lido fica salvo.)",
                parent=self):
            return
        self.destroy()
        self.app.sair(forcar=True)


AVISO_PRIVACIDADE = """PRIVACIDADE E PROTEÇÃO DE DADOS PESSOAIS (LGPD)

O QUE O PROGRAMA GUARDA
• Um índice com o texto extraído dos documentos da pasta principal. Esse texto pode conter dados pessoais (nomes, números de documentos, assinaturas) e até dados pessoais sensíveis (ex.: atestados de saúde).
• Os usuários do programa (as senhas ficam guardadas só como hash, impossível de ler), as configurações, o registro de uso (auditoria.log) e o registro de erros (erro.log). O registro de uso não guarda o conteúdo dos documentos nem os termos pesquisados.
Tudo fica somente neste computador, na pasta:
{pasta}

O QUE SAI DO COMPUTADOR
• Nenhum conteúdo de documento é enviado pela internet: a leitura e o OCR são feitos neste computador.
• O programa consulta o GitHub para verificar atualizações. Como em qualquer acesso a um site, o GitHub recebe o endereço de rede (IP) do computador. A verificação automática pode ser desligada em Opções.
• Índices exportados e pacotes de cliente saem cifrados com senha (AES-256). Envie a senha por um canal diferente do arquivo.

QUEM RESPONDE PELOS DADOS
• O controlador dos dados é a organização dona dos documentos{controlador}. Pedidos e dúvidas dos titulares (confirmação, acesso, correção, eliminação) devem ser encaminhados ao encarregado de dados (DPO) dessa organização.

BOAS PRÁTICAS
• Cada pessoa com a sua própria conta do Windows e do programa. Nunca compartilhe senhas.
• Mantenha a criptografia de disco do Windows (BitLocker) ligada: o login do programa protege as telas, mas o índice fica gravado na pasta acima.
• Em Opções > Proteção de dados pessoais, marque pastas como ASO e exames para que o conteúdo não seja lido.
• A sessão é bloqueada após um tempo sem uso (Opções > Bloquear após inatividade).
• Ao fim do uso ou do contrato: Arquivo > Eliminar dados.

INCIDENTES
Em caso de perda, roubo ou vazamento (ex.: notebook extraviado, pacote enviado à pessoa errada), avise IMEDIATAMENTE o responsável pela organização dona dos documentos: a comunicação à ANPD tem prazo de 3 dias úteis.
"""


class TelaPacote(tk.Toplevel):
    """Dados do pacote do cliente. resultado = dict ou None."""

    def __init__(self, master, empresas, tem_indice, inicial):
        super().__init__(master)
        self.empresas = empresas
        self.resultado = None
        self._prev = None
        self.title("Criar pacote do cliente")
        self.resizable(False, False)
        f = ttk.Frame(self, padding=14)
        f.pack(fill="both", expand=True)
        f.columnconfigure(1, weight=1)

        ttk.Label(f, text="Nome do cliente:").grid(row=0, column=0, sticky="w", pady=3)
        self.var_cliente = tk.StringVar(value=inicial.get("cliente", ""))
        ttk.Entry(f, textvariable=self.var_cliente, width=40).grid(row=0, column=1, columnspan=2, sticky="ew", pady=3)

        ttk.Label(f, text="Cor do menu:").grid(row=1, column=0, sticky="w", pady=3)
        cores = ttk.Frame(f)
        cores.grid(row=1, column=1, columnspan=2, sticky="w", pady=3)
        self.var_cor = tk.StringVar(value=inicial.get("cor", ""))
        ttk.Entry(cores, textvariable=self.var_cor, width=10).pack(side="left")
        self.amostra = tk.Label(cores, width=4, relief="solid", bd=1)
        self.amostra.pack(side="left", padx=6)
        ttk.Button(cores, text="Escolher...", command=self._escolher_cor).pack(side="left")
        ttk.Label(cores, text="(vazio = azul padrão)", foreground="#666").pack(side="left", padx=6)
        self.var_cor.trace_add("write", lambda *a: self._amostra())

        ttk.Label(f, text="Logomarca:").grid(row=2, column=0, sticky="nw", pady=3)
        self.var_logo = tk.StringVar(value=inicial.get("logo", ""))
        ttk.Entry(f, textvariable=self.var_logo, width=40).grid(row=2, column=1, sticky="ew", pady=3)
        ttk.Button(f, text="Escolher...", command=self._escolher_logo).grid(row=2, column=2, padx=(4, 0), pady=3)
        self.previa = ttk.Label(f)
        self.previa.grid(row=3, column=1, columnspan=2, sticky="w", pady=(2, 8))
        self.var_logo.trace_add("write", lambda *a: self._previa_logo())

        self.var_indice = tk.BooleanVar(value=tem_indice)
        ttk.Checkbutton(f, text="Incluir o índice da pasta principal atual", variable=self.var_indice,
                        state="normal" if tem_indice else "disabled",
                        command=self._indice_mudou).grid(row=4, column=0, columnspan=3, sticky="w")
        caixa = ttk.Frame(f)
        caixa.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(4, 2))
        self.lista = tk.Listbox(caixa, selectmode="multiple", height=min(10, max(3, len(empresas))),
                                exportselection=False, activestyle="none")
        sb = ttk.Scrollbar(caixa, orient="vertical", command=self.lista.yview)
        self.lista.configure(yscrollcommand=sb.set)
        self.lista.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        for nome, _ in empresas:
            self.lista.insert("end", nome)
        self.lista.select_set(0, "end")
        self.var_fora = tk.BooleanVar(value=True)
        self.chk_fora = ttk.Checkbutton(f, text='Incluir arquivos fora das pastas "DB_"', variable=self.var_fora)
        self.chk_fora.grid(row=6, column=0, columnspan=3, sticky="w")
        ttk.Label(f, foreground="#a0522d", wraplength=440, justify="left",
                  text="⚠ Com o índice, o pacote contém o texto dos documentos. Envie só por canal "
                       "privado e só a quem tem direito a esse conteúdo. Nunca publique no GitHub."
                  ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(10, 0))
        fim = ttk.Frame(f)
        fim.grid(row=8, column=0, columnspan=3, sticky="e", pady=(12, 0))
        ttk.Button(fim, text="Criar pacote...", command=self.confirmar).pack(side="left", padx=4)
        ttk.Button(fim, text="Cancelar", command=self.destroy).pack(side="left")
        self._amostra()
        self._previa_logo()
        self._indice_mudou()
        self.transient(master)
        self.grab_set()

    def _escolher_cor(self):
        from tkinter import colorchooser
        atual = self.var_cor.get().strip()
        _, hexa = colorchooser.askcolor(color=atual if core.RE_COR.match(atual) else None, parent=self)
        if hexa:
            self.var_cor.set(hexa.upper())

    def _amostra(self):
        c = self.var_cor.get().strip()
        self.amostra.config(bg=c if core.RE_COR.match(c) else "#002060")

    def _escolher_logo(self):
        p = filedialog.askopenfilename(parent=self, title="Logomarca do cliente",
                                       filetypes=[("Imagens", "*.png *.jpg *.jpeg *.bmp *.gif"), ("Todos", "*.*")])
        if p:
            self.var_logo.set(os.path.normpath(p))

    def _previa_logo(self):
        p = self.var_logo.get().strip()
        self._prev = None
        if p and os.path.isfile(p):
            try:
                from PIL import Image, ImageTk
                img = Image.open(p)
                img.thumbnail((240, 60))
                self._prev = ImageTk.PhotoImage(img)
            except Exception:
                self._prev = None
        self.previa.config(image=self._prev or "", text="" if self._prev else ("(imagem inválida)" if p else ""))

    def _indice_mudou(self):
        estado = "normal" if self.var_indice.get() else "disabled"
        self.lista.config(state=estado)
        self.chk_fora.config(state=estado)

    def confirmar(self):
        cliente = self.var_cliente.get().strip()
        cor = self.var_cor.get().strip().upper()
        logo = self.var_logo.get().strip()
        if not cliente:
            messagebox.showwarning("Pacote do cliente", "Informe o nome do cliente.", parent=self)
            return
        if cor and not core.RE_COR.match(cor):
            messagebox.showwarning("Pacote do cliente", "Cor inválida. Use o formato #RRGGBB, ex.: #FA7C13.", parent=self)
            return
        if logo and not os.path.isfile(logo):
            messagebox.showwarning("Pacote do cliente", "Arquivo da logomarca não encontrado.", parent=self)
            return
        pastas, fora = None, True
        if self.var_indice.get():
            sel = list(self.lista.curselection())
            fora = self.var_fora.get()
            if not sel and not fora:
                messagebox.showwarning("Pacote do cliente", "Escolha ao menos uma empresa.", parent=self)
                return
            if not (len(sel) == len(self.empresas) and fora):
                pastas = [self.empresas[i][1] for i in sel]
        self.resultado = {"cliente": cliente, "cor": cor or None, "logo": logo or None,
                          "incluir_indice": self.var_indice.get(), "pastas": pastas, "fora": fora}
        self.destroy()


class TelaCopias(tk.Toplevel):
    """Todas as copias e versoes de um documento que apareceu uma vez so na busca."""

    def __init__(self, master, principal, ao_abrir, ao_abrir_pasta):
        super().__init__(master)
        self.itens = principal["copias"]
        self.ao_abrir = ao_abrir
        self.ao_abrir_pasta = ao_abrir_pasta
        self.title(f"Cópias e versões — {core._nome_documento(principal)}")
        self.geometry("980x320")
        f = ttk.Frame(self, padding=10)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="★ = exibida na busca (versão mais recente, no caminho mais curto). "
                          "Mesmo tamanho = cópia idêntica; tamanho diferente = outra versão.",
                  foreground="#555").pack(anchor="w", pady=(0, 6))
        lista = ttk.Frame(f)
        lista.pack(fill="both", expand=True)
        self.tv = ttk.Treeview(lista, columns=("data", "tamanho", "pasta"), selectmode="extended")
        self.tv.heading("#0", text="Arquivo")
        self.tv.heading("data", text="Modificado em")
        self.tv.heading("tamanho", text="Tamanho")
        self.tv.heading("pasta", text="Pasta")
        self.tv.column("#0", width=280, stretch=False)
        self.tv.column("data", width=120, anchor="center", stretch=False)
        self.tv.column("tamanho", width=80, anchor="e", stretch=False)
        self.tv.column("pasta", width=460, stretch=False)
        sy = ttk.Scrollbar(lista, orient="vertical", command=self.tv.yview)
        sx = ttk.Scrollbar(lista, orient="horizontal", command=self.tv.xview)
        self.tv.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tv.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        lista.rowconfigure(0, weight=1)
        lista.columnconfigure(0, weight=1)
        rolagem_lateral(self.tv)
        manter_ultima_coluna(self.tv, "pasta", minimo=460)
        for i, r in enumerate(self.itens):
            marca = "★ " if (r["caminho"], r["interno"]) == (principal["caminho"], principal["interno"]) else "    "
            nome = r["nome"] + (f"  ›  {r['interno']}" if r["interno"] else "")
            self.tv.insert("", "end", iid=str(i), text=marca + nome,
                           values=(formatar_data(r.get("mtime")), formatar_tamanho(r.get("tamanho", 0)), r["pasta"]))
        ajustar_ultima_coluna(self.tv, "pasta", [r["pasta"] for r in self.itens], minimo=460)
        self.tv.bind("<Double-1>", lambda e: self._abrir())
        b = ttk.Frame(f)
        b.pack(fill="x", pady=(8, 0))
        ttk.Button(b, text="Abrir", command=self._abrir).pack(side="left")
        ttk.Button(b, text="Abrir pasta", command=self._pasta).pack(side="left", padx=6)
        ttk.Button(b, text="Fechar", command=self.destroy).pack(side="right")
        self.transient(master)

    def _sel(self):
        return [self.itens[int(k)] for k in self.tv.selection()]

    def _abrir(self):
        self.ao_abrir(self._sel())

    def _pasta(self):
        self.ao_abrir_pasta(self._sel())


# --------------------------------------------------------------------------
# Menu lateral de empresas
# --------------------------------------------------------------------------

class MenuEmpresas(tk.Frame):
    """Barra azul a esquerda: "Todas" + um botao por pasta DB_NOME + logomarca no rodape."""

    def __init__(self, master, ao_escolher):
        super().__init__(master, bg=AZUL, width=LARGURA_MENU)
        self.pack_propagate(False)
        self.ao_escolher = ao_escolher
        self.botoes = {}            # caminho (ou None para "Todas") -> Label
        self.selecionado = None

        tk.Label(self, text="EMPRESAS", bg=AZUL, fg=AZUL_TEXTO_SUAVE, anchor="w",
                 font=("Segoe UI", 9, "bold"), padx=16, pady=14).pack(side="top", fill="x")
        self._montar_rodape()       # antes da lista: side="bottom" mantem o rodape fixo

        area = tk.Frame(self, bg=AZUL)
        area.pack(side="top", fill="both", expand=True)
        self.canvas = tk.Canvas(area, bg=AZUL, highlightthickness=0, bd=0)
        self.barra = ttk.Scrollbar(area, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.barra.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.lista = tk.Frame(self.canvas, bg=AZUL)
        self._janela = self.canvas.create_window((0, 0), window=self.lista, anchor="nw")
        self.lista.bind("<Configure>", lambda e: self._ajustar_rolagem())
        self.canvas.bind("<Configure>", lambda e: (self.canvas.itemconfigure(self._janela, width=e.width),
                                                   self._ajustar_rolagem()))
        self.bind_all("<MouseWheel>", self._roda_mouse, add="+")

    def _montar_rodape(self):
        rodape = tk.Frame(self, bg=AZUL)
        rodape.pack(side="bottom", fill="x", pady=(6, 14))
        tk.Frame(rodape, bg=BOTAO_COR, height=1).pack(fill="x", padx=12, pady=(0, 10))
        tk.Label(rodape, text="Desenvolvido por:", bg=AZUL, fg=AZUL_TEXTO_SUAVE,
                 font=("Segoe UI", 8)).pack()
        self._logo = carregar_logo(LARGURA_MENU - 40, 90)   # referencia mantida (senao o Tk descarta)
        if self._logo:
            tk.Label(rodape, image=self._logo, bg=AZUL).pack(pady=(6, 0))
        else:
            tk.Label(rodape, text=EMPRESA_DEV, bg=AZUL, fg="white",
                     font=("Segoe UI", 9, "bold")).pack(pady=(4, 0))

    def _ajustar_rolagem(self):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        precisa = self.lista.winfo_reqheight() > self.canvas.winfo_height()
        if precisa and not self.barra.winfo_ismapped():
            self.barra.pack(side="right", fill="y")
        elif not precisa and self.barra.winfo_ismapped():
            self.barra.pack_forget()
            self.canvas.yview_moveto(0)

    def _roda_mouse(self, ev):
        """Rola a lista so quando o mouse esta sobre o menu lateral."""
        try:
            alvo = self.winfo_containing(ev.x_root, ev.y_root)
        except (KeyError, tk.TclError):
            return
        caminho = str(alvo) if alvo is not None else ""
        if caminho != str(self) and not caminho.startswith(str(self) + "."):
            return
        if self.lista.winfo_reqheight() > self.canvas.winfo_height():
            self.canvas.yview_scroll(int(-ev.delta / 120), "units")

    def carregar(self, empresas):
        for w in self.lista.winfo_children():
            w.destroy()
        self.botoes.clear()
        self._botao("Todas", None, None)
        tk.Frame(self.lista, bg=BOTAO_COR, height=1).pack(fill="x", padx=12, pady=6)
        if not empresas:
            tk.Label(self.lista, text='Nenhuma pasta "DB_..."\nna pasta principal.', bg=AZUL,
                     fg=AZUL_TEXTO_SUAVE, justify="left", anchor="w", padx=16,
                     font=("Segoe UI", 9)).pack(fill="x", pady=4)
        for nome, caminho in empresas:
            self._botao(nome, caminho, nome)
        self.marcar(None)

    def _botao(self, texto, chave, nome):
        b = tk.Label(self.lista, text=texto, bg=BOTAO_COR, fg="white", anchor="w", justify="left",
                     padx=12, pady=6, font=FONTE_MENU, cursor="hand2", takefocus=1,
                     relief="raised", bd=2, wraplength=LARGURA_MENU - 50,
                     highlightthickness=1, highlightbackground=AZUL, highlightcolor="white")
        b.pack(fill="x", padx=10, pady=3)

        def soltar(e):
            # como um botao de verdade: so aciona se soltar o clique em cima dele
            if 0 <= e.x < b.winfo_width() and 0 <= e.y < b.winfo_height():
                self._clicar(chave, nome)
            else:
                self._pintar(chave, b)

        b.bind("<ButtonPress-1>", lambda e: b.config(relief="sunken"))
        b.bind("<ButtonRelease-1>", soltar)
        b.bind("<Return>", lambda e: self._clicar(chave, nome))
        b.bind("<space>", lambda e: self._clicar(chave, nome))
        b.bind("<Enter>", lambda e: chave != self.selecionado and b.config(bg=BOTAO_HOVER))
        b.bind("<Leave>", lambda e: self._pintar(chave, b))
        self.botoes[chave] = b

    def _pintar(self, chave, b):
        if chave == self.selecionado:
            b.config(bg="white", fg=DESTAQUE, font=FONTE_MENU_SEL, relief="sunken")
        else:
            b.config(bg=BOTAO_COR, fg="white", font=FONTE_MENU, relief="raised")

    def _clicar(self, chave, nome):
        self.marcar(chave)
        self.ao_escolher(chave, nome)

    def marcar(self, chave):
        """Selecionado: branco e 'afundado'. Demais: azul em relevo."""
        if chave not in self.botoes:
            chave = None
        self.selecionado = chave
        for k, b in self.botoes.items():
            self._pintar(k, b)


# --------------------------------------------------------------------------
# Janela principal
# --------------------------------------------------------------------------

class App:
    def __init__(self, root, usuario, usuarios):
        self.root = root
        self.usuario = usuario
        self.usuarios = usuarios
        self.cfg = carregar_config()
        self.fila = queue.Queue()
        self.trabalhando = False     # alguma tarefa longa em andamento (indexar/exportar/importar)
        self.tarefa = None           # "indexar", "indexar_auto", "exportar", "importar"
        self.atualizando = False
        self.cancelar_flag = threading.Event()
        self.indice_busca = None
        self.resultados = []         # tudo o que a busca encontrou
        self.exibidos = []           # o que esta na lista (filtro de tipo + copias agrupadas)
        self.empresa = None          # caminho da pasta da empresa; None = todas
        self.empresa_nome = None
        self.temp_dir = tempfile.mkdtemp(prefix="buscadatabook_")
        self.ordem = {}
        self.sessao_bloqueada = False
        self._janelas_ocultas = []
        self._msgs_adiadas = []          # mensagens de tarefas que terminaram com a sessao bloqueada
        self._ultima_atividade = time.monotonic()
        self._dados_apagados = False

        self._titulo()
        root.geometry(self.cfg.get("geometria", "1280x720"))
        root.minsize(980, 520)
        root.protocol("WM_DELETE_WINDOW", self.sair)
        root.report_callback_exception = self._erro_tk

        self._estilo()
        self._menu()
        self._montar()
        self._carregar_pasta_inicial()
        self.root.after(100, self._processar_fila)
        self.root.after(3000, self._verificacao_automatica)
        for ev in ("<Any-KeyPress>", "<Any-ButtonPress>", "<Motion>", "<MouseWheel>"):
            root.bind_all(ev, self._atividade, add="+")
        self.root.after(30_000, self._checar_inatividade)

    def _titulo(self):
        cliente = f" — {MARCA['cliente']}" if MARCA else ""
        self.root.title(f"{APP_NOME}{cliente}  ·  {VERSAO}  ·  {self.usuario['login']}")

    def _erro_tk(self, tipo, valor, tb):
        log.error("Erro na interface", exc_info=(tipo, valor, tb))
        messagebox.showerror(APP_NOME, f"Ocorreu um erro inesperado:\n{valor}\n\n"
                                       f"Detalhes gravados em:\n{ARQUIVO_LOG}")

    # ---- layout -----------------------------------------------------------

    def _estilo(self):
        st = ttk.Style()
        if "vista" in st.theme_names():
            st.theme_use("vista")
        st.configure("Treeview", rowheight=24)
        st.configure("Titulo.TLabel", font=("Segoe UI", 10, "bold"))
        st.configure("Escopo.TLabel", font=("Segoe UI", 10, "bold"), foreground=DESTAQUE)

    def _menu(self):
        m = tk.Menu(self.root)
        arq = tk.Menu(m, tearoff=0)
        arq.add_command(label="Escolher pasta...", command=self.escolher_pasta)
        arq.add_command(label="Atualizar índice", command=self.atualizar_indice)
        arq.add_command(label="Refazer índice do zero", command=self.refazer_indice)
        arq.add_command(label="Recarregar lista de empresas", command=lambda: self._carregar_empresas(manter=True))
        arq.add_command(label="Arquivos que não puderam ser lidos", command=self.ver_erros)
        arq.add_separator()
        if self.usuario["admin"]:    # o arquivo exportado leva o texto dos documentos
            arq.add_command(label="Exportar índice (com senha)...", command=self.exportar_indice)
        arq.add_command(label="Importar índice...", command=self.importar_indice)
        arq.add_separator()
        arq.add_command(label="Importar pacote do cliente...", command=lambda: self.importar_pacote())
        if self.usuario["admin"]:
            arq.add_command(label="Criar pacote do cliente...", command=self.criar_pacote)
            arq.add_command(label="Remover personalização do cliente", command=self.remover_personalizacao)
            arq.add_separator()
            elim = tk.Menu(arq, tearoff=0)
            elim.add_command(label="Apagar o índice desta pasta...", command=self.apagar_indice_pasta)
            elim.add_command(label="Apagar todos os dados do programa neste computador...",
                             command=self.apagar_todos_os_dados)
            arq.add_cascade(label="Eliminar dados (LGPD)", menu=elim)
            arq.add_command(label="Registro de auditoria...", command=self.ver_auditoria)
        arq.add_separator()
        arq.add_command(label="Bloquear sessão agora", command=self.bloquear)
        arq.add_command(label="Sair", command=self.sair)
        m.add_cascade(label="Arquivo", menu=arq)

        op = tk.Menu(m, tearoff=0)
        self.var_ocr = tk.BooleanVar(value=self.cfg.get("ocr", True))
        self.var_auto_idx = tk.BooleanVar(value=self.cfg.get("atualizar_ao_abrir", True))
        self.var_nuvem = tk.BooleanVar(value=self.cfg.get("economizar_onedrive", True))
        self.var_auto_upd = tk.BooleanVar(value=self.cfg.get("verificar_atualizacoes", True))
        op.add_checkbutton(label="Ler PDFs escaneados (OCR) ao indexar", variable=self.var_ocr,
                           command=self._salvar_cfg)
        op.add_checkbutton(label="Atualizar o índice ao abrir o programa", variable=self.var_auto_idx,
                           command=self._salvar_cfg)
        op.add_checkbutton(label="Economizar espaço do OneDrive (devolver à nuvem após ler)",
                           variable=self.var_nuvem, command=self._salvar_cfg)
        op.add_checkbutton(label="Verificar atualizações do programa ao iniciar", variable=self.var_auto_upd,
                           command=self._salvar_cfg)
        op.add_separator()
        self.var_inativ = tk.IntVar(value=self._minutos_bloqueio())
        sub = tk.Menu(op, tearoff=0)
        for mnt in OPCOES_INATIVIDADE_MIN:
            sub.add_radiobutton(label=f"{mnt} minutos", value=mnt, variable=self.var_inativ, command=self._salvar_cfg)
        op.add_cascade(label="Bloquear após inatividade", menu=sub)
        if self.usuario["admin"]:
            op.add_command(label="Proteção de dados pessoais (LGPD)...", command=self.editar_protecao)
        m.add_cascade(label="Opções", menu=op)

        conta = tk.Menu(m, tearoff=0)
        conta.add_command(label="Trocar minha senha", command=self.trocar_minha_senha)
        conta.add_command(label="Gerar novo código de recuperação", command=self.novo_codigo)
        m.add_cascade(label="Minha conta", menu=conta)
        if self.usuario["admin"]:
            m.add_command(label="Usuários", command=lambda: TelaUsuarios(self.root, self.usuarios, self.usuario["login"]))
        aj = tk.Menu(m, tearoff=0)
        aj.add_command(label="Como buscar", command=lambda: messagebox.showinfo("Como buscar", DICAS))
        aj.add_command(label="Verificar atualizações", command=lambda: self._verificar_atualizacao(manual=True))
        aj.add_command(label="Privacidade e proteção de dados (LGPD)", command=self.ver_privacidade)
        aj.add_separator()
        aj.add_command(label="Sobre", command=self._sobre)
        m.add_cascade(label="Ajuda", menu=aj)
        self.root.config(menu=m)

    def _sobre(self):
        messagebox.showinfo("Sobre", f"{APP_NOME} {VERSAO}\n"
                                     "Busca de palavras em pastas locais, de rede e de nuvem sincronizada.\n\n"
                                     f"Dados e log de erros:\n{core.pasta_dados_app()}\n\n"
                                     f"github.com/{REPO_GITHUB}")

    def _montar(self):
        self.menu_emp = MenuEmpresas(self.root, self._empresa_escolhida)
        self.menu_emp.pack(side="left", fill="y")
        principal = ttk.Frame(self.root)
        principal.pack(side="left", fill="both", expand=True)
        self.principal = principal

        topo = ttk.Frame(principal, padding=(10, 10, 10, 4))
        topo.pack(fill="x")
        topo.columnconfigure(1, weight=1)

        # logomarca do cliente: canto superior direito, acima de "Atualizar indice"
        self.lbl_marca = ttk.Label(topo)
        self.lbl_marca.grid(row=0, column=2, columnspan=3, sticky="e", pady=(0, 6))
        self._img_marca = None
        self._mostrar_logo_cliente()

        ttk.Label(topo, text="Pasta principal:", style="Titulo.TLabel").grid(row=1, column=0, sticky="w")
        self.var_pasta = tk.StringVar()
        self.cb_pasta = ttk.Combobox(topo, textvariable=self.var_pasta, values=self.cfg.get("recentes", []))
        self.cb_pasta.grid(row=1, column=1, sticky="ew", padx=6)
        self.cb_pasta.bind("<<ComboboxSelected>>", lambda e: self._pasta_mudou())
        self.cb_pasta.bind("<Return>", lambda e: self._pasta_mudou())
        ttk.Button(topo, text="Procurar...", command=self.escolher_pasta).grid(row=1, column=2, padx=2)
        self.bt_indice = ttk.Button(topo, text="Atualizar índice", command=self.atualizar_indice)
        self.bt_indice.grid(row=1, column=3, columnspan=2, sticky="ew", padx=2)

        self.lbl_info = ttk.Label(topo, text="", foreground="#555")
        self.lbl_info.grid(row=2, column=1, sticky="w", padx=6, pady=(2, 8))

        ttk.Label(topo, text="Buscar:", style="Titulo.TLabel").grid(row=3, column=0, sticky="w")
        self.var_termo = tk.StringVar()
        self.e_termo = ttk.Entry(topo, textvariable=self.var_termo, font=("Segoe UI", 11))
        self.e_termo.grid(row=3, column=1, sticky="ew", padx=6)
        self.e_termo.bind("<Return>", lambda e: self.buscar())
        self.root.bind("<Escape>", lambda e: self.limpar_busca())
        ttk.Button(topo, text="Buscar", command=self.buscar).grid(row=3, column=2, padx=2)
        ttk.Button(topo, text="Limpar", command=self.limpar_busca).grid(row=3, column=3, padx=2)
        self.var_filtro = tk.StringVar(value="Todos os tipos")
        cb = ttk.Combobox(topo, textvariable=self.var_filtro, values=list(FILTROS), state="readonly", width=18)
        cb.grid(row=3, column=4, padx=2)
        cb.bind("<<ComboboxSelected>>", lambda e: self._mostrar_resultados())

        linha3 = ttk.Frame(topo)
        linha3.grid(row=4, column=1, sticky="w", padx=6, pady=(6, 0))
        ttk.Label(linha3, text="Procurar no:").pack(side="left")
        self.var_modo = tk.StringVar(value=self.cfg.get("modo_busca", "tudo"))
        ttk.Radiobutton(linha3, text="Nome e conteúdo", value="tudo", variable=self.var_modo,
                        command=self._modo_mudou).pack(side="left", padx=(6, 0))
        ttk.Radiobutton(linha3, text="Somente nome do arquivo", value="nome", variable=self.var_modo,
                        command=self._modo_mudou).pack(side="left", padx=(8, 0))
        self.lbl_escopo = ttk.Label(linha3, text="", style="Escopo.TLabel")
        self.lbl_escopo.pack(side="left", padx=(18, 0))
        self.var_ocultar = tk.BooleanVar(value=self.cfg.get("ocultar_repetidos", True))
        ttk.Checkbutton(topo, text="Ocultar cópias e versões antigas", variable=self.var_ocultar,
                        command=self._ocultar_mudou).grid(row=4, column=2, columnspan=3, sticky="e", pady=(6, 0))

        # resultados
        meio = ttk.Frame(principal, padding=(10, 4))
        meio.pack(fill="both", expand=True)
        cols = ("tipo", "onde", "copias", "trecho", "pasta")
        self.tv = ttk.Treeview(meio, columns=cols, selectmode="extended")
        self.tv.heading("#0", text="Arquivo", command=lambda: self._ordenar("#0"))
        self.tv.heading("tipo", text="Tipo", command=lambda: self._ordenar("tipo"))
        self.tv.heading("onde", text="Achado no", command=lambda: self._ordenar("onde"))
        self.tv.heading("copias", text="Cópias")
        self.tv.heading("trecho", text="Trecho")
        self.tv.heading("pasta", text="Pasta", command=lambda: self._ordenar("pasta"))
        self.tv.column("#0", width=280, stretch=False)
        self.tv.column("tipo", width=55, stretch=False, anchor="center")
        self.tv.column("onde", width=80, stretch=False, anchor="center")
        self.tv.column("copias", width=60, stretch=False, anchor="center")
        self.tv.column("trecho", width=410, stretch=False)
        self.tv.column("pasta", width=300, stretch=False)   # alarga ate caber o caminho (ver _mostrar_resultados)
        sy = ttk.Scrollbar(meio, orient="vertical", command=self.tv.yview)
        sx = ttk.Scrollbar(meio, orient="horizontal", command=self.tv.xview)
        self.tv.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tv.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        meio.rowconfigure(0, weight=1)
        meio.columnconfigure(0, weight=1)
        self.tv.bind("<Double-1>", lambda e: self.abrir())
        self.tv.bind("<Return>", lambda e: self.abrir())
        self.tv.bind("<Button-3>", self._menu_contexto)
        rolagem_lateral(self.tv)
        manter_ultima_coluna(self.tv, "pasta")

        self.menu_ctx = tk.Menu(self.root, tearoff=0)
        self.menu_ctx.add_command(label="Abrir", command=self.abrir)
        self.menu_ctx.add_command(label="Abrir pasta do arquivo", command=self.abrir_pasta)
        self.menu_ctx.add_command(label="Salvar cópia...", command=self.salvar_copia)
        self.menu_ctx.add_command(label="Copiar caminho", command=self.copiar_caminho)
        self.menu_ctx.add_separator()
        self.menu_ctx.add_command(label="Ver todas as cópias e versões", command=self.ver_copias)

        # rodape
        base = ttk.Frame(principal, padding=(10, 4, 10, 10))
        base.pack(fill="x")
        ttk.Button(base, text="Abrir", command=self.abrir).pack(side="left", padx=(0, 4))
        ttk.Button(base, text="Abrir pasta", command=self.abrir_pasta).pack(side="left", padx=4)
        ttk.Button(base, text="Salvar cópia...", command=self.salvar_copia).pack(side="left", padx=4)
        self.lbl_status = ttk.Label(base, text="Pronto.")
        self.lbl_status.pack(side="left", padx=12, fill="x", expand=True)
        self.bt_cancelar = ttk.Button(base, text="Cancelar", command=self.cancelar_indexacao)
        self.prog = ttk.Progressbar(base, length=220, mode="determinate")

        self.menu_emp.carregar([])
        self._atualizar_escopo()

    # ---- empresas ---------------------------------------------------------

    def _carregar_empresas(self, manter=False):
        anterior = self.empresa if manter else None
        empresas = empresas_permitidas(self.usuario, self.indice_busca.empresas()) if self.indice_busca else []
        self.menu_emp.carregar(empresas)
        nomes = {c: n for n, c in empresas}
        self.empresa = anterior if anterior in nomes else None
        self.empresa_nome = nomes.get(self.empresa)
        self.menu_emp.marcar(self.empresa)
        self._atualizar_escopo()

    def _empresa_escolhida(self, caminho, nome):
        self.empresa, self.empresa_nome = caminho, nome
        self._atualizar_escopo()
        if self.var_termo.get().strip():
            self.buscar()
        else:
            self.resultados = []
            self._mostrar_resultados()
        self.e_termo.focus_set()

    def _texto_escopo(self):
        return f"empresa {self.empresa_nome}" if self.empresa else "todas as empresas"

    def _atualizar_escopo(self):
        self.lbl_escopo.config(text=f"Buscando em: {self._texto_escopo()}")

    # ---- pasta / indice ---------------------------------------------------

    def _carregar_pasta_inicial(self):
        p = self.cfg.get("ultima_pasta")
        if p:
            self.var_pasta.set(p)
            self._pasta_mudou(silencioso=True)
            # atualizacao automatica em segundo plano (so se a pasta ja foi indexada antes)
            if self.indice_busca and self.var_auto_idx.get() and self.indice_busca.info()[0] > 0:
                self.root.after(1500, lambda: self.atualizar_indice(automatico=True))
        self.e_termo.focus_set()

    def escolher_pasta(self):
        p = filedialog.askdirectory(title="Escolha a pasta principal", initialdir=self.var_pasta.get() or None)
        if p:
            self.var_pasta.set(os.path.normpath(p))
            self._pasta_mudou()

    def _pasta_mudou(self, silencioso=False):
        if self.trabalhando:
            messagebox.showinfo(APP_NOME, "Aguarde o fim da tarefa em andamento (ou cancele) "
                                          "antes de trocar de pasta.")
            return
        pasta = self.var_pasta.get().strip().strip('"')
        if not pasta:
            return
        if not os.path.isdir(pasta):
            self.lbl_info.config(text="⚠ Pasta não encontrada ou sem acesso (verifique a rede / VPN).")
            return
        rec = [r for r in self.cfg.get("recentes", []) if os.path.normcase(r) != os.path.normcase(pasta)]
        self.cfg["recentes"] = [pasta] + rec[:9]
        self.cfg["ultima_pasta"] = pasta
        self.cb_pasta["values"] = self.cfg["recentes"]
        self._salvar_cfg()
        if self.indice_busca:
            self.indice_busca.fechar()
        self.indice_busca = core.Indice(pasta)
        self.resultados = []
        self._mostrar_resultados()
        self._carregar_empresas()
        self._atualizar_info()
        n, _ = self.indice_busca.info()
        if n == 0 and not silencioso:
            if messagebox.askyesno(APP_NOME, "Esta pasta ainda não foi indexada.\n\n"
                                   "Deseja ler os arquivos agora? Na primeira vez pode demorar, "
                                   "dependendo da quantidade de documentos."):
                self.atualizar_indice()

    def _atualizar_info(self):
        if not self.indice_busca:
            return
        n, quando = self.indice_busca.info()
        if n:
            self.lbl_info.config(text=f"{n:,} arquivos no índice · atualizado em {quando}".replace(",", "."))
        else:
            self.lbl_info.config(text="Pasta ainda não indexada. Clique em “Atualizar índice”.")

    def _iniciar_tarefa(self, tarefa, cancelavel):
        self.trabalhando = True
        self.tarefa = tarefa
        self.cancelar_flag.clear()
        self.bt_indice.config(state="disabled")
        self.prog.pack(side="right", padx=4)
        if cancelavel:
            self.bt_cancelar.pack(side="right", padx=4)
        else:
            self.prog.config(mode="indeterminate")
            self.prog.start(15)

    def _terminar_tarefa(self):
        self.trabalhando = False
        self.tarefa = None
        self.prog.stop()
        self.prog.config(mode="determinate", value=0)
        self.bt_indice.config(state="normal")
        self.prog.pack_forget()
        self.bt_cancelar.pack_forget()

    def atualizar_indice(self, automatico=False):
        if self.trabalhando:
            return
        pasta = self.var_pasta.get().strip()
        if not pasta or not os.path.isdir(pasta):
            if not automatico:
                messagebox.showwarning(APP_NOME, "Escolha uma pasta válida primeiro.")
            return
        if self.var_ocr.get() and not core.configurar_ocr() and not automatico:
            messagebox.showinfo(APP_NOME, "O OCR (Tesseract) não está disponível.\n\n"
                                "Os PDFs escaneados serão indexados apenas pelo nome. "
                                "Reinstale o programa para restaurar o OCR.")
        self._iniciar_tarefa("indexar_auto" if automatico else "indexar", cancelavel=True)
        if automatico:
            self.lbl_status.config(text="Verificando arquivos novos em segundo plano (a busca continua liberada)...")
        # manual: tenta de novo os arquivos que deram erro; automatico: so o que mudou
        args = (pasta, self.var_ocr.get(), self.var_nuvem.get(), not automatico)
        threading.Thread(target=self._tarefa_indexar, args=args, daemon=True).start()

    def refazer_indice(self):
        pasta = self.var_pasta.get().strip()
        if not pasta or self.trabalhando:
            return
        if not messagebox.askyesno(APP_NOME, "Apagar o índice desta pasta e ler todos os arquivos novamente?"):
            return
        regras = []
        if self.indice_busca:
            regras = self.indice_busca.regras_protecao()      # as regras de protecao continuam valendo
            self.indice_busca.fechar()
            self.indice_busca = None
        db = core.caminho_indice(pasta)
        for suf in ("", "-wal", "-shm"):
            try:
                os.remove(str(db) + suf)
            except OSError:
                pass
        self.indice_busca = core.Indice(pasta)
        if regras:
            self.indice_busca.definir_regras_protecao(regras)
        core.auditar(self.usuario["login"], "índice refeito do zero", pasta)
        self.atualizar_indice()

    def _tarefa_indexar(self, pasta, usar_ocr, liberar_nuvem, tentar_erros=False):
        idx = core.Indice(pasta)
        try:
            r = idx.atualizar(usar_ocr=usar_ocr,
                              progresso=lambda m, i, n: self.fila.put(("prog", m, i, n)),
                              cancelado=self.cancelar_flag.is_set,
                              liberar_nuvem=liberar_nuvem,
                              tentar_erros=tentar_erros)
            self.fila.put(("fim", r))
        except Exception as e:
            log.exception("Erro ao indexar %s", pasta)
            self.fila.put(("erro", str(e)))
        finally:
            idx.fechar()

    def cancelar_indexacao(self):
        self.cancelar_flag.set()
        self.lbl_status.config(text="Cancelando...")

    def _processar_fila(self):
        ultimo_prog = None
        try:
            while True:
                msg = self.fila.get_nowait()
                tipo = msg[0]
                if self.sessao_bloqueada and tipo != "prog":
                    self._msgs_adiadas.append(msg)    # nada aparece na tela com a sessao bloqueada
                    continue
                if tipo == "cb":                      # ("cb", funcao, args): roda na interface
                    msg[1](*msg[2])
                elif tipo == "prog":
                    ultimo_prog = msg
                elif tipo == "fim":
                    self._fim_indexacao(msg[1])
                elif tipo == "erro":
                    self._fim_indexacao(None, msg[1])
                elif tipo.startswith(("exp_", "imp_", "pac_", "pim_")):
                    self._fim_exportar_importar(msg)
                elif tipo.startswith("upd_"):
                    self._tratar_atualizacao(msg)
        except queue.Empty:
            pass
        except Exception:
            log.exception("Erro ao processar mensagens da interface")
        if ultimo_prog and self.trabalhando and self.tarefa in ("indexar", "indexar_auto"):
            _, m, i, n = ultimo_prog
            txt = f"{i}/{n} · {m}" if n else m
            if self.tarefa == "indexar_auto":
                txt = "Atualização automática: " + txt
            self.lbl_status.config(text=txt[:140])
            if n:
                self.prog.config(mode="determinate", maximum=n, value=i)
            else:
                self.prog.config(mode="indeterminate")
                self.prog.step(5)
        self.root.after(150, self._processar_fila)

    def _fim_indexacao(self, r, erro=None):
        automatico = self.tarefa == "indexar_auto"
        self._terminar_tarefa()
        self._atualizar_info()
        self._carregar_empresas(manter=True)   # pode ter surgido uma pasta DB_ nova
        if erro:
            self.lbl_status.config(text=f"Erro ao indexar (detalhes em {ARQUIVO_LOG.name}).")
            if not automatico:
                messagebox.showerror(APP_NOME, f"Erro ao indexar:\n{erro}\n\nDetalhes em:\n{ARQUIVO_LOG}")
            return
        if automatico and not (r["lidos"] or r["removidos"]) and not r["falhas"]:
            txt = "Índice em dia: nenhum arquivo novo ou alterado"
        else:
            txt = f"Índice atualizado: {r['lidos']} lidos/relidos, {r['removidos']} removidos"
        if r.get("liberados"):
            txt += f", {r['liberados']} devolvidos à nuvem do OneDrive"
        if r["erros"]:
            txt += f", {r['erros']} não puderam ser lidos (Arquivo > Arquivos que não puderam ser lidos)"
        if r.get("falhas"):
            txt += (f". ⚠ {r['falhas']} pasta(s)/arquivo(s) inacessíveis (rede/VPN?): "
                    "por segurança nada foi removido do índice")
        if r["cancelado"]:
            faltam = max(0, r.get("a_ler", r["lidos"]) - r["lidos"])
            txt = ("Indexação cancelada. " + txt
                   + (f". Os já lidos ficaram salvos; faltam {faltam}, que serão lidos na próxima atualização"
                      if faltam else ""))
        self.lbl_status.config(text=txt + ".")
        if self.var_termo.get().strip() and (not automatico or r["lidos"] or r["removidos"]):
            self.buscar()

    def ver_erros(self):
        if not self.indice_busca:
            return
        erros = self.indice_busca.arquivos_com_erro()
        if not erros:
            messagebox.showinfo(APP_NOME, "Nenhum arquivo com erro de leitura.")
            return
        w = tk.Toplevel(self.root)
        w.title(f"Arquivos que não puderam ser lidos ({len(erros)})")
        w.geometry("900x440")
        barra = ttk.Frame(w, padding=8)
        barra.pack(side="bottom", fill="x")

        def tentar():
            w.destroy()
            self.atualizar_indice()      # a atualizacao manual rele os arquivos com erro

        ttk.Button(barra, text="Tentar ler novamente", command=tentar,
                   state="disabled" if self.trabalhando else "normal").pack(side="left")
        ttk.Label(barra, foreground="#555",
                  text="Lê de novo só estes arquivos (e os novos ou alterados).").pack(side="left", padx=10)
        ttk.Button(barra, text="Fechar", command=w.destroy).pack(side="right")
        t = tk.Text(w, wrap="none")
        t.pack(fill="both", expand=True)
        for c, s in erros:
            t.insert("end", f"{c}\n    {s}\n")
        t.config(state="disabled")

    # ---- exportar / importar indice --------------------------------------

    def exportar_indice(self):
        if not self.usuario["admin"]:
            return
        if self.trabalhando:
            messagebox.showinfo(APP_NOME, "Aguarde o fim da tarefa em andamento (ou cancele).")
            return
        if not self.indice_busca or self.indice_busca.info()[0] == 0:
            messagebox.showinfo(APP_NOME, "Escolha uma pasta principal já indexada para exportar.")
            return
        tela = TelaExportar(self.root, self.indice_busca.empresas())
        self.root.wait_window(tela)
        if tela.resultado is None:
            return
        pastas, fora = tela.resultado
        ts = TelaSenhaArquivo(self.root, True, "Senha do índice exportado",
                              "O índice exportado contém o texto dos documentos e será cifrado com esta senha.")
        self.root.wait_window(ts)
        senha = ts.resultado
        if not senha:
            return
        base = os.path.basename(self.indice_busca.raiz.rstrip("\\/")) or "Indice"
        destino = filedialog.asksaveasfilename(
            title="Salvar índice exportado", defaultextension=core.EXT_EXPORTACAO,
            initialfile=f"Indice_{base}_{time.strftime('%Y-%m-%d')}{core.EXT_EXPORTACAO}",
            filetypes=[("Índice do Busca Databook", f"*{core.EXT_EXPORTACAO}")])
        if not destino:
            return
        raiz = self.indice_busca.raiz
        self._iniciar_tarefa("exportar", cancelavel=False)
        self.lbl_status.config(text="Exportando índice...")

        def tarefa():
            idx = core.Indice(raiz)
            try:
                n = idx.exportar(destino, pastas, fora, senha=senha)
                core.auditar(self.usuario["login"], "índice exportado (cifrado)", f"{n} arquivos -> {destino}")
                self.fila.put(("exp_fim", n, destino))
            except Exception as e:
                log.exception("Erro ao exportar índice")
                self.fila.put(("exp_erro", str(e)))
            finally:
                idx.fechar()

        threading.Thread(target=tarefa, daemon=True).start()

    def importar_indice(self):
        if self.trabalhando:
            messagebox.showinfo(APP_NOME, "Aguarde o fim da tarefa em andamento (ou cancele).")
            return
        arq = filedialog.askopenfilename(
            title="Importar índice",
            filetypes=[("Índice do Busca Databook", f"*{core.EXT_EXPORTACAO}"), ("Todos os arquivos", "*.*")])
        if not arq:
            return
        senha = None
        if core.esta_cifrado(arq):
            r = self._pedir_senha_arquivo("Importar índice", lambda s: core.ler_exportacao(arq, s))
            if not r:
                return
            senha, info = r
        else:
            try:
                info = core.ler_exportacao(arq)
            except ValueError as e:
                messagebox.showerror(APP_NOME, str(e))
                return
        raiz = info["raiz"]
        resumo = (f"Índice com {info['arquivos']:,} arquivos".replace(",", ".")
                  + (f", exportado em {info['exportado_em']}" if info["exportado_em"] else "")
                  + f".\n\nPasta de origem:\n{raiz}\n\n")
        if raiz and os.path.isdir(raiz):
            resp = messagebox.askyesnocancel(
                APP_NOME, resumo + "Essa pasta existe neste computador.\n\n"
                                   "Sim: usar essa pasta\nNão: escolher outra pasta")
            if resp is None:
                return
            nova = raiz if resp else filedialog.askdirectory(title="Onde está a pasta principal neste computador?")
        else:
            messagebox.showinfo(APP_NOME, resumo + "Essa pasta não foi encontrada neste computador.\n"
                                                   "Na próxima tela, escolha onde estão os databooks.")
            nova = filedialog.askdirectory(title="Onde está a pasta principal neste computador?")
        if not nova:
            return
        nova = os.path.normpath(nova)
        self._iniciar_tarefa("importar", cancelavel=False)
        self.lbl_status.config(text="Importando índice...")

        def tarefa():
            try:
                n = core.importar_indice(arq, nova, senha)
                core.auditar(self.usuario["login"], "índice importado", f"{n} arquivos <- {arq}")
                self.fila.put(("imp_fim", n, nova))
            except Exception as e:
                log.exception("Erro ao importar índice")
                self.fila.put(("imp_erro", str(e)))

        threading.Thread(target=tarefa, daemon=True).start()

    # ---- personalizacao do cliente (pacote .bdpack) -------------------------

    def _mostrar_logo_cliente(self):
        self._img_marca = None
        logo = (MARCA or {}).get("logo")
        if logo:
            try:
                from PIL import Image, ImageTk
                img = Image.open(logo)
                img.thumbnail((260, 56), getattr(Image, "Resampling", Image).LANCZOS)
                self._img_marca = ImageTk.PhotoImage(img)
            except Exception:
                log.warning("Logomarca do cliente não carregada", exc_info=True)
        if self._img_marca:
            self.lbl_marca.config(image=self._img_marca, text="")
            self.lbl_marca.grid()
        else:
            self.lbl_marca.config(image="", text="")
            self.lbl_marca.grid_remove()

    def _reaplicar_tema(self):
        """Aplica na hora a personalizacao instalada (cores, titulo, logomarca)."""
        aplicar_tema(core.carregar_marca())
        self._titulo()
        ttk.Style().configure("Escopo.TLabel", foreground=DESTAQUE)
        self.menu_emp.destroy()
        self.menu_emp = MenuEmpresas(self.root, self._empresa_escolhida)
        self.menu_emp.pack(side="left", fill="y", before=self.principal)
        self._carregar_empresas(manter=True)
        self._mostrar_logo_cliente()

    def criar_pacote(self):
        if self.trabalhando:
            messagebox.showinfo(APP_NOME, "Aguarde o fim da tarefa em andamento (ou cancele).")
            return
        tem_indice = bool(self.indice_busca and self.indice_busca.info()[0] > 0)
        empresas = self.indice_busca.empresas() if tem_indice else []
        tela = TelaPacote(self.root, empresas, tem_indice, self.cfg.get("ultimo_pacote", {}))
        self.root.wait_window(tela)
        r = tela.resultado
        if not r:
            return
        self.cfg["ultimo_pacote"] = {"cliente": r["cliente"], "cor": r["cor"] or "", "logo": r["logo"] or ""}
        self._salvar_cfg()
        senha = None
        if r["incluir_indice"]:
            ts = TelaSenhaArquivo(self.root, True, "Senha do pacote",
                                  "O índice vai dentro do pacote cifrado com esta senha. O cliente vai precisar "
                                  "dela para importar o pacote.")
            self.root.wait_window(ts)
            senha = ts.resultado
            if not senha:
                return
        nome = "".join(c if c.isalnum() or c in " -_" else "_" for c in r["cliente"]).strip() or "Cliente"
        destino = filedialog.asksaveasfilename(
            title="Salvar pacote do cliente", defaultextension=core.EXT_PACOTE,
            initialfile=f"BuscaDatabook_{nome}_{time.strftime('%Y-%m-%d')}{core.EXT_PACOTE}",
            filetypes=[("Pacote do Busca Databook", f"*{core.EXT_PACOTE}")])
        if not destino:
            return
        raiz = self.indice_busca.raiz if (r["incluir_indice"] and self.indice_busca) else None
        self._iniciar_tarefa("exportar", cancelavel=False)
        self.lbl_status.config(text="Criando o pacote do cliente...")

        def tarefa():
            idx = core.Indice(raiz) if raiz else None
            try:
                logo = open(r["logo"], "rb").read() if r["logo"] else None
                m = core.criar_pacote(destino, r["cliente"], r["cor"], logo, idx, r["pastas"], r["fora"], senha=senha)
                core.auditar(self.usuario["login"], "pacote do cliente criado",
                             f"{r['cliente']}: {(m['indice'] or {}).get('arquivos', 0)} arquivos -> {destino}")
                self.fila.put(("pac_fim", m, destino))
            except Exception as e:
                log.exception("Erro ao criar pacote")
                self.fila.put(("pac_erro", str(e)))
            finally:
                if idx:
                    idx.fechar()

        threading.Thread(target=tarefa, daemon=True).start()

    def importar_pacote(self, arquivo=None):
        if self.trabalhando:
            messagebox.showinfo(APP_NOME, "Aguarde o fim da tarefa em andamento (ou cancele).")
            return
        if not arquivo:
            arquivo = filedialog.askopenfilename(
                title="Importar pacote do cliente",
                filetypes=[("Pacote do Busca Databook", f"*{core.EXT_PACOTE}"), ("Todos os arquivos", "*.*")])
            if not arquivo:
                return
        try:
            info = core.ler_pacote(arquivo)
        except ValueError as e:
            messagebox.showerror(APP_NOME, str(e))
            return
        senha = None
        if info["indice_cifrado"]:
            def testar(s):
                if not core.conferir_senha_pacote(arquivo, s):
                    raise ValueError("Senha incorreta.")
                return True
            r = self._pedir_senha_arquivo(f"Pacote de {info['cliente']}", testar)
            if not r:
                return
            senha = r[0]
        resumo = f"Pacote: {info['cliente']}" + (f" (criado em {info['criado_em']})" if info["criado_em"] else "")
        if info["tem_indice"] and info["indice"]:
            resumo += f"\nÍndice com {info['indice'].get('arquivos', 0):,} arquivos".replace(",", ".")
        nova = None
        if info["tem_indice"]:
            origem = info["pasta_origem"]
            if origem and os.path.isdir(origem):
                resp = messagebox.askyesnocancel(
                    APP_NOME, resumo + f"\n\nPasta dos databooks:\n{origem}\n\nEssa pasta existe neste computador."
                                       "\n\nSim: usar essa pasta\nNão: escolher outra pasta")
                if resp is None:
                    return
                nova = origem if resp else filedialog.askdirectory(title="Onde está a pasta dos databooks neste computador?")
            else:
                messagebox.showinfo(APP_NOME, resumo + "\n\nNa próxima tela, escolha onde estão os databooks "
                                                       "neste computador.")
                nova = filedialog.askdirectory(title="Onde está a pasta dos databooks neste computador?")
            if not nova:
                return
            nova = os.path.normpath(nova)
        elif not messagebox.askyesno(APP_NOME, resumo + "\n\nAplicar a personalização deste cliente?"):
            return
        self._iniciar_tarefa("importar", cancelavel=False)
        self.lbl_status.config(text="Aplicando o pacote do cliente...")

        def tarefa():
            try:
                info2, n = core.importar_pacote(arquivo, nova, senha)
                core.auditar(self.usuario["login"], "pacote do cliente importado", f"{info2['cliente']}: {n} arquivos")
                self.fila.put(("pim_fim", info2, n, nova))
            except Exception as e:
                log.exception("Erro ao importar pacote")
                self.fila.put(("pim_erro", str(e)))

        threading.Thread(target=tarefa, daemon=True).start()

    def remover_personalizacao(self):
        if not MARCA:
            messagebox.showinfo(APP_NOME, "Nenhuma personalização de cliente instalada.")
            return
        if messagebox.askyesno(APP_NOME, f"Remover a personalização de “{MARCA['cliente']}” "
                                         "(nome, cores e logomarca)?\n\nO índice e as buscas não são afetados."):
            core.remover_marca()
            log.info("Personalização do cliente removida")
            core.auditar(self.usuario["login"], "personalização do cliente removida", MARCA["cliente"])
            self._reaplicar_tema()

    def _fim_exportar_importar(self, msg):
        self._terminar_tarefa()
        tipo = msg[0]
        if tipo == "pac_fim":
            _, m, destino = msg
            self.lbl_status.config(text="Pacote do cliente criado.")
            det = (f"Índice: {m['indice']['arquivos']:,} arquivos".replace(",", ".") if m["indice"]
                   else "Sem índice (só a identidade visual).")
            aviso = ("\n\nO índice está cifrado: envie a SENHA por um canal diferente do pacote "
                     "(ex.: telefone ou mensagem)." if m["indice"] else "")
            messagebox.showinfo(APP_NOME, f"Pacote de “{m['cliente']}” criado:\n{destino}\n\n{det}\n\n"
                                          "No computador do cliente: instale o programa (versão 3.6 ou mais nova) "
                                          "e use Arquivo > Importar pacote do cliente (ou dê dois cliques no "
                                          "arquivo)." + aviso)
            return
        if tipo == "pim_fim":
            _, info, n, nova = msg
            self._reaplicar_tema()
            if nova:
                self.var_pasta.set(nova)
                self._pasta_mudou(silencioso=True)
            self.lbl_status.config(text=f"Pacote de {info['cliente']} aplicado.")
            messagebox.showinfo(APP_NOME, f"Pacote de “{info['cliente']}” aplicado."
                                + (f"\n\n{n:,} arquivos importados para:\n{nova}".replace(",", ".") if nova else ""))
            return
        if tipo in ("pac_erro", "pim_erro"):
            acao = "criar" if tipo == "pac_erro" else "aplicar"
            self.lbl_status.config(text=f"Não foi possível {acao} o pacote.")
            messagebox.showerror(APP_NOME, f"Não foi possível {acao} o pacote do cliente:\n{msg[1]}")
            return
        if tipo == "exp_fim":
            _, n, destino = msg
            self.lbl_status.config(text=f"Índice exportado: {n} arquivos.")
            messagebox.showinfo(APP_NOME, f"{n:,} arquivos exportados para:\n{destino}".replace(",", ".")
                                + "\n\nO arquivo está cifrado: envie a SENHA por um canal diferente do arquivo."
                                + "\n\nNo outro computador: Arquivo > Importar índice.")
        elif tipo == "imp_fim":
            _, n, nova = msg
            self.var_pasta.set(nova)          # a pasta importada passa a ser a pasta principal
            self._pasta_mudou(silencioso=True)
            self.lbl_status.config(text=f"Índice importado: {n} arquivos.")
            messagebox.showinfo(APP_NOME, f"{n:,} arquivos importados.".replace(",", ".")
                                + f"\n\nPasta principal configurada:\n{nova}\n\n"
                                  "A busca já pode ser usada. Para incluir arquivos novos, "
                                  "clique em “Atualizar índice”.")
        else:
            acao = "exportar" if tipo == "exp_erro" else "importar"
            self.lbl_status.config(text=f"Não foi possível {acao} o índice.")
            messagebox.showerror(APP_NOME, f"Não foi possível {acao} o índice:\n{msg[1]}")

    def _pedir_senha_arquivo(self, titulo, testar):
        """Pede a senha ate acertar (no maximo TENTATIVAS_LOGIN vezes). testar(senha) devolve um resultado
        ou levanta ValueError("Senha incorreta."). Devolve (senha, resultado) ou None."""
        for tentativa in range(TENTATIVAS_LOGIN):
            ts = TelaSenhaArquivo(self.root, False, titulo, "Este arquivo é protegido por senha. "
                                  "Digite a senha informada por quem enviou o arquivo.")
            self.root.wait_window(ts)
            if not ts.resultado:
                return None
            try:
                self.root.config(cursor="watch")
                self.root.update_idletasks()
                return ts.resultado, testar(ts.resultado)
            except ValueError as e:
                if str(e) != "Senha incorreta.":
                    messagebox.showerror(APP_NOME, str(e))
                    return None
                core.auditar(self.usuario["login"], "senha de arquivo incorreta", titulo)
                messagebox.showerror(APP_NOME, "Senha incorreta.")
            finally:
                self.root.config(cursor="")
        messagebox.showerror(APP_NOME, "Muitas tentativas erradas. Confirme a senha com quem enviou o arquivo.")
        return None

    # ---- protecao de dados pessoais e eliminacao (LGPD) -----------------------

    def editar_protecao(self):
        if self.trabalhando:
            messagebox.showinfo(APP_NOME, "Aguarde o fim da tarefa em andamento (ou cancele).")
            return
        if not self.indice_busca:
            messagebox.showinfo(APP_NOME, "Escolha a pasta principal primeiro.")
            return
        tela = TelaProtecao(self.root, self.indice_busca.regras_protecao())
        self.root.wait_window(tela)
        if tela.resultado is None:
            return
        regras, raiz = tela.resultado, self.indice_busca.raiz
        self._iniciar_tarefa("protecao", cancelavel=False)
        self.lbl_status.config(text="Aplicando a proteção de dados ao índice...")

        def tarefa():
            idx = core.Indice(raiz)
            try:
                rem, so_nome = idx.definir_regras_protecao(regras)
                core.auditar(self.usuario["login"], "regras de proteção alteradas",
                             "; ".join(f"{r['termo']}={r['modo']}" for r in regras) or "(nenhuma)")
                self.fila.put(("cb", self._fim_protecao, (rem, so_nome, None)))
            except Exception as e:
                log.exception("Erro ao aplicar a proteção")
                self.fila.put(("cb", self._fim_protecao, (0, 0, str(e))))
            finally:
                idx.fechar()

        threading.Thread(target=tarefa, daemon=True).start()

    def _fim_protecao(self, rem, so_nome, erro):
        self._terminar_tarefa()
        self._atualizar_info()
        if erro:
            messagebox.showerror(APP_NOME, f"Não foi possível aplicar a proteção:\n{erro}")
            return
        self.resultados = []
        self._mostrar_resultados()
        self.lbl_status.config(text="Proteção de dados aplicada.")
        messagebox.showinfo(APP_NOME, "Proteção de dados aplicada.\n\n"
                                      f"{rem} arquivo(s) retirado(s) do índice.\n"
                                      f"{so_nome} arquivo(s) passaram a ser buscados só pelo nome.\n\n"
                                      "Ao remover uma regra, o conteúdo volta a ser lido na próxima "
                                      "atualização do índice.")

    def apagar_indice_pasta(self):
        if self.trabalhando or not self.indice_busca:
            messagebox.showinfo(APP_NOME, "Escolha a pasta e aguarde o fim de tarefas em andamento.")
            return
        raiz = self.indice_busca.raiz
        conf = simpledialog.askstring(
            "Apagar índice", f"Isto apaga o índice (e as regras de proteção) de:\n{raiz}\n\n"
                             "Os documentos originais NÃO são apagados.\n\nPara confirmar, digite APAGAR:",
            parent=self.root)
        if (conf or "").strip().upper() != "APAGAR" or not self._confirmar_senha_atual("Apagar índice"):
            return
        self.indice_busca.fechar()
        self.indice_busca = None
        try:
            core.apagar_indice(raiz)
        except OSError as e:
            messagebox.showerror(APP_NOME, f"Não foi possível apagar o índice:\n{e}")
        core.auditar(self.usuario["login"], "índice apagado", raiz)
        self.indice_busca = core.Indice(raiz)
        self.resultados = []
        self._mostrar_resultados()
        self._atualizar_info()
        messagebox.showinfo(APP_NOME, "Índice apagado.")

    def apagar_todos_os_dados(self):
        if self.trabalhando:
            messagebox.showinfo(APP_NOME, "Aguarde o fim da tarefa em andamento (ou cancele).")
            return
        conf = simpledialog.askstring(
            "Apagar todos os dados",
            "Isto apaga TUDO o que o programa guarda neste usuário do Windows: índices, usuários e senhas, "
            "configurações, personalização do cliente e registros.\n\nOs documentos originais NÃO são "
            "apagados. O programa será fechado e, ao abrir de novo, começará do zero.\n\n"
            "Para confirmar, digite APAGAR TUDO:", parent=self.root)
        if (conf or "").strip().upper() != "APAGAR TUDO" or not self._confirmar_senha_atual("Apagar todos os dados"):
            return
        core.auditar(self.usuario["login"], "todos os dados do programa apagados")
        if self.indice_busca:
            self.indice_busca.fechar()
            self.indice_busca = None
        self.usuarios.fechar()
        for lg in (logging.getLogger(), core.log_auditoria):
            for h in list(lg.handlers):
                h.close()
                lg.removeHandler(h)
        falhas = core.apagar_todos_os_dados()
        self._dados_apagados = True
        shutil.rmtree(self.temp_dir, ignore_errors=True)
        if falhas:
            messagebox.showwarning(APP_NOME, "Alguns itens não puderam ser apagados (podem estar em uso):\n\n"
                                   + "\n".join(falhas[:15]) + f"\n\nApague manualmente a pasta:\n{core.pasta_dados_app()}")
        else:
            messagebox.showinfo(APP_NOME, "Todos os dados do programa foram apagados. O programa será fechado.")
        self.root.destroy()

    def ver_auditoria(self):
        arquivos = sorted(core.pasta_dados_app().glob("auditoria.log*"),
                          key=lambda p: -int(p.suffix[1:]) if p.suffix[1:].isdigit() else 0)
        linhas = []
        for a in arquivos:                 # do mais antigo (auditoria.log.20) ao atual
            try:
                linhas += a.read_text(encoding="utf-8", errors="replace").splitlines()
            except OSError:
                pass
        texto = ("Registro de auditoria (data | usuário | evento | detalhe)\n\n"
                 + ("\n".join(linhas[-20000:]) if linhas else "(vazio)"))
        TelaTexto(self.root, "Registro de auditoria", texto,
                  salvar_como=f"auditoria_busca_databook_{time.strftime('%Y-%m-%d')}.txt")

    def ver_privacidade(self):
        controlador = f" ({MARCA['cliente']})" if MARCA else ""
        TelaTexto(self.root, "Privacidade e proteção de dados (LGPD)",
                  AVISO_PRIVACIDADE.format(pasta=core.pasta_dados_app(), controlador=controlador))

    # ---- bloqueio por inatividade -------------------------------------------

    def _minutos_bloqueio(self):
        m = self.cfg.get("bloqueio_min", BLOQUEIO_INATIVIDADE_MIN)
        return m if m in OPCOES_INATIVIDADE_MIN else BLOQUEIO_INATIVIDADE_MIN

    def _atividade(self, _e=None):
        self._ultima_atividade = time.monotonic()

    def _checar_inatividade(self):
        try:
            if (not self.sessao_bloqueada
                    and time.monotonic() - self._ultima_atividade > self._minutos_bloqueio() * 60):
                self.bloquear(automatico=True)
        finally:
            self.root.after(30_000, self._checar_inatividade)

    def bloquear(self, automatico=False):
        if self.sessao_bloqueada:
            return
        self.sessao_bloqueada = True
        # janela modal aberta (grab): solta o grab, senao a tela de bloqueio nao recebe o teclado
        self._grab_anterior = self.root.grab_current()
        if self._grab_anterior:
            self._grab_anterior.grab_release()
        self._janelas_ocultas = [w for w in self.root.winfo_children()
                                 if isinstance(w, tk.Toplevel) and w.winfo_viewable()]
        for w in self._janelas_ocultas:
            w.withdraw()
        self.root.withdraw()
        core.auditar(self.usuario["login"], "sessão bloqueada", "inatividade" if automatico else "pelo usuário")
        TelaBloqueio(self)

    def desbloquear(self):
        self.root.deiconify()
        for w in self._janelas_ocultas:
            try:
                w.deiconify()
            except tk.TclError:
                pass
        self._janelas_ocultas = []
        if getattr(self, "_grab_anterior", None):
            try:
                self._grab_anterior.grab_set()
            except tk.TclError:
                pass
            self._grab_anterior = None
        self.sessao_bloqueada = False
        self._ultima_atividade = time.monotonic()
        core.auditar(self.usuario["login"], "sessão desbloqueada")
        for msg in self._msgs_adiadas:          # devolve a fila o que chegou durante o bloqueio
            self.fila.put(msg)
        self._msgs_adiadas = []

    # ---- busca ------------------------------------------------------------

    def buscar(self):
        termo = self.var_termo.get().strip()
        if not termo:
            return
        if not self.indice_busca:
            self._pasta_mudou()
            if not self.indice_busca:
                return
        n, _ = self.indice_busca.info()
        if n == 0 and not self.trabalhando:
            if messagebox.askyesno(APP_NOME, "A pasta ainda não foi indexada. Indexar agora?"):
                self.atualizar_indice()
            return
        try:
            if self.var_modo.get() == "nome":
                self.resultados = self.indice_busca.buscar_nome(termo, pasta=self.empresa)
            else:
                self.resultados = self.indice_busca.buscar(termo, pasta=self.empresa)
        except Exception as e:
            log.exception("Erro na busca")
            messagebox.showerror(APP_NOME, f"Não foi possível buscar: {e}")
            return
        self._mostrar_resultados()

    def limpar_busca(self):
        self.var_termo.set("")
        self.resultados = []
        self._mostrar_resultados()
        if not self.trabalhando:
            self.lbl_status.config(text="Pronto.")
        self.e_termo.focus_set()

    def _ocultar_mudou(self):
        self._salvar_cfg()
        self._mostrar_resultados()

    def _modo_mudou(self):
        self.cfg["modo_busca"] = self.var_modo.get()
        self._salvar_cfg()
        if self.var_termo.get().strip():
            self.buscar()
        self.e_termo.focus_set()

    def _mostrar_resultados(self):
        self.tv.delete(*self.tv.get_children())
        filtro = FILTROS.get(self.var_filtro.get())
        base = [r for r in self.resultados if not filtro or r["ext"] in filtro]
        if self.var_ocultar.get() and self.indice_busca:
            self.exibidos = core.agrupar_repetidos(base, self.indice_busca.raiz)
        else:
            self.exibidos = base
        for i, r in enumerate(self.exibidos):
            nome = r["nome"] + (f"  ›  {r['interno']}" if r["interno"] else "")
            n = len(r.get("copias", ()))
            self.tv.insert("", "end", iid=str(i), text=nome,
                           values=(r["ext"].upper(), r["onde"], str(n) if n > 1 else "", r["trecho"], r["pasta"]))
        ajustar_ultima_coluna(self.tv, "pasta", [r["pasta"] for r in self.exibidos])
        self.tv.xview_moveto(0)
        termo = self.var_termo.get().strip()
        if not self.trabalhando and termo:
            onde = self._texto_escopo()
            if self.var_modo.get() == "nome":
                onde = "nomes de arquivo, " + onde
            visiveis = len(self.exibidos)
            txt = (f"{visiveis} resultado(s) para “{termo}” em {onde}." if visiveis
                   else f"Nada encontrado para “{termo}” em {onde}.")
            ocultos = len(base) - visiveis
            if ocultos:
                txt += f" {ocultos} cópia(s)/versão(ões) antiga(s) ocultas."
            self.lbl_status.config(text=txt)

    def _ordenar(self, col):
        inverso = self.ordem.get(col, False)
        itens = [(self.tv.item(k, "text") if col == "#0" else self.tv.set(k, col), k)
                 for k in self.tv.get_children("")]
        itens.sort(key=lambda x: x[0].lower(), reverse=inverso)
        for pos, (_, k) in enumerate(itens):
            self.tv.move(k, "", pos)
        self.ordem[col] = not inverso

    # ---- acoes sobre resultados ------------------------------------------

    def _selecionados(self):
        return [self.exibidos[int(k)] for k in self.tv.selection()]

    def _menu_contexto(self, ev):
        linha = self.tv.identify_row(ev.y)
        if linha:
            if linha not in self.tv.selection():
                self.tv.selection_set(linha)
            self.menu_ctx.tk_popup(ev.x_root, ev.y_root)

    def _permitido(self, r):
        """So abre arquivos de dentro da pasta principal (protege contra indices adulterados)."""
        if self.indice_busca and self.indice_busca.contem(r["caminho"]):
            return True
        messagebox.showwarning(APP_NOME, f"Por segurança, só são abertos arquivos da pasta principal:\n{r['caminho']}")
        return False

    def abrir(self, itens=None):
        for r in (self._selecionados() if itens is None else itens)[:10]:
            if not self._permitido(r):
                continue
            try:
                if not os.path.exists(core.caminho_longo(r["caminho"])):
                    messagebox.showwarning(APP_NOME, f"Arquivo não encontrado (movido ou apagado?):\n{r['caminho']}\n\n"
                                           "Atualize o índice.")
                    continue
                if r["interno"]:
                    d = tempfile.mkdtemp(dir=self.temp_dir)
                    abrir_no_sistema(extrair_do_zip(r["caminho"], r["interno"], d))
                else:
                    abrir_no_sistema(core.caminho_para_abrir(r["caminho"]))
            except Exception as e:
                log.warning("Falha ao abrir %s", r["caminho"], exc_info=True)
                messagebox.showerror(APP_NOME, f"Não foi possível abrir:\n{r['caminho']}\n\n{e}")

    def abrir_pasta(self, itens=None):
        sel = self._selecionados() if itens is None else itens
        if sel and self._permitido(sel[0]):
            mostrar_na_pasta(sel[0]["caminho"])

    def ver_copias(self):
        sel = self._selecionados()
        if not sel:
            return
        if len(sel[0].get("copias", ())) < 2:
            messagebox.showinfo(APP_NOME, "Este documento não tem outras cópias ou versões nos resultados.")
            return
        TelaCopias(self.root, sel[0], self.abrir, self.abrir_pasta)

    def salvar_copia(self):
        sel = self._selecionados()
        if not sel:
            messagebox.showinfo(APP_NOME, "Selecione um ou mais arquivos na lista.")
            return
        destino = filedialog.askdirectory(title="Salvar cópia em...",
                                          initialdir=self.cfg.get("pasta_copias") or os.path.expanduser("~\\Downloads"))
        if not destino:
            return
        self.cfg["pasta_copias"] = destino
        self._salvar_cfg()
        ok, falhas = 0, []
        for r in sel:
            if not self.indice_busca or not self.indice_busca.contem(r["caminho"]):
                falhas.append(f"{r['nome']}: fora da pasta principal")
                continue
            try:
                if r["interno"]:
                    extrair_do_zip(r["caminho"], r["interno"], destino)
                else:
                    alvo = os.path.join(destino, r["nome"])
                    base, ext = os.path.splitext(alvo)
                    n = 1
                    while os.path.exists(alvo):
                        alvo = f"{base} ({n}){ext}"
                        n += 1
                    shutil.copy2(core.caminho_longo(r["caminho"]), alvo)
                ok += 1
            except Exception as e:
                falhas.append(f"{r['nome']}: {e}")
        msg = f"{ok} arquivo(s) copiado(s) para:\n{destino}"
        if falhas:
            msg += "\n\nFalharam:\n" + "\n".join(falhas[:10])
        if messagebox.askyesno(APP_NOME, msg + "\n\nAbrir a pasta de destino?"):
            abrir_no_sistema(destino)

    def copiar_caminho(self):
        sel = self._selecionados()
        if sel:
            txt = "\n".join(r["caminho"] + (f" ({r['interno']})" if r["interno"] else "") for r in sel)
            self.root.clipboard_clear()
            self.root.clipboard_append(txt)
            self.lbl_status.config(text="Caminho copiado.")

    def _confirmar_senha_atual(self, titulo):
        atual = simpledialog.askstring(titulo, "Para confirmar, digite a sua senha atual:", show="•",
                                       parent=self.root)
        if atual is None:
            return False
        if not self.usuarios.validar(self.usuario["login"], atual):
            log.warning("Senha atual incorreta em '%s' (%s)", titulo, self.usuario["login"])
            core.auditar(self.usuario["login"], "senha atual incorreta", titulo)
            messagebox.showerror(APP_NOME, "Senha atual incorreta.")
            return False
        return True

    def trocar_minha_senha(self):
        if not self._confirmar_senha_atual("Trocar senha"):
            return
        tela = TelaNovaSenha(self.root, self.usuarios, self.usuario["login"])
        self.root.wait_window(tela)
        if tela.ok:
            messagebox.showinfo(APP_NOME, "Senha alterada.")

    def novo_codigo(self):
        if not self._confirmar_senha_atual("Novo código de recuperação"):
            return
        if not messagebox.askyesno(APP_NOME, "Gerar um novo código de recuperação?\n\n"
                                             "O código anterior deixará de valer."):
            return
        codigo = self.usuarios.gerar_codigo(self.usuario["login"])
        log.info("Novo código de recuperação gerado por %s", self.usuario["login"])
        core.auditar(self.usuario["login"], "código de recuperação gerado")
        mostrar_codigo(self.root, codigo)

    # ---- atualizacao automatica do programa ------------------------------

    def _verificacao_automatica(self):
        if not (updater.rodando_instalado() and self.var_auto_upd.get()):
            return
        if time.time() - float(self.cfg.get("ultima_verificacao", 0)) < INTERVALO_VERIFICACAO_S:
            return
        self._verificar_atualizacao(manual=False)

    def _verificar_atualizacao(self, manual):
        if self.atualizando:
            return
        if not updater.rodando_instalado():
            if manual:
                messagebox.showinfo(APP_NOME, "A atualização automática só funciona no programa instalado.")
            return
        self.atualizando = True

        def tarefa():
            try:
                self.fila.put(("upd_info", updater.buscar_atualizacao(), manual))
            except Exception as e:
                log.warning("Falha ao verificar atualização", exc_info=True)
                self.fila.put(("upd_erro", str(e), manual))

        threading.Thread(target=tarefa, daemon=True).start()

    def _baixar_atualizacao(self, info):
        self.atualizando = True
        self.lbl_status.config(text=f"Baixando a versão {info['versao']}...")

        def tarefa():
            try:
                caminho, sha = updater.baixar_e_verificar(
                    info, progresso=lambda b, t: self.fila.put(("upd_prog", b, t)))
                self.fila.put(("upd_pronto", caminho, sha, info["versao"]))
            except Exception as e:
                log.error("Falha ao baixar atualização", exc_info=True)
                self.fila.put(("upd_erro", str(e), True))

        threading.Thread(target=tarefa, daemon=True).start()

    def _tratar_atualizacao(self, msg):
        tipo = msg[0]
        if tipo == "upd_prog":
            _, b, t = msg
            if t and not self.trabalhando:
                self.lbl_status.config(text=f"Baixando atualização... {b * 100 // t}%")
            return
        if tipo == "upd_erro":
            self.atualizando = False
            if msg[2]:
                messagebox.showerror(APP_NOME, f"Não foi possível atualizar:\n{msg[1]}")
            return
        if tipo == "upd_info":
            _, info, manual = msg
            self.atualizando = False
            self.cfg["ultima_verificacao"] = time.time()
            self._salvar_cfg()
            if info is None:
                if manual:
                    messagebox.showinfo(APP_NOME, f"Você já está na versão mais recente ({VERSAO}).")
                return
            notas = f"\n\nNovidades:\n{info['notas'][:800]}" if info["notas"] else ""
            if messagebox.askyesno(APP_NOME, f"Nova versão disponível: {info['versao']} (atual: {VERSAO})."
                                             f"{notas}\n\nBaixar e instalar agora?\n"
                                             "O programa será fechado e reaberto automaticamente."):
                self._baixar_atualizacao(info)
            return
        if tipo == "upd_pronto":
            _, caminho, sha, versao = msg
            self.atualizando = False
            self.lbl_status.config(text=f"Versão {versao} verificada (SHA-256 ok). Instalando...")
            try:
                updater.instalar(caminho, sha)
            except Exception as e:
                log.error("Falha ao iniciar instalador", exc_info=True)
                messagebox.showerror(APP_NOME, f"Não foi possível instalar:\n{e}")
                return
            self.sair(forcar=True)

    # ---- encerrar ----------------------------------------------------------

    def _salvar_cfg(self):
        for chave, var in (("ocr", "var_ocr"), ("atualizar_ao_abrir", "var_auto_idx"),
                           ("economizar_onedrive", "var_nuvem"), ("verificar_atualizacoes", "var_auto_upd"),
                           ("ocultar_repetidos", "var_ocultar")):
            if hasattr(self, var):
                self.cfg[chave] = bool(getattr(self, var).get())
        if hasattr(self, "var_inativ"):
            self.cfg["bloqueio_min"] = int(self.var_inativ.get())
        salvar_config(self.cfg)

    def sair(self, forcar=False):
        # a atualizacao automatica e interrompida sem perguntar; as demais tarefas pedem confirmacao
        if self.trabalhando and not forcar and self.tarefa != "indexar_auto" and not messagebox.askyesno(
                APP_NOME, "Há uma tarefa em andamento. Sair mesmo assim?\n"
                          "(O que já foi lido fica salvo.)"):
            return
        self.cancelar_flag.set()
        self.cfg["geometria"] = self.root.geometry()
        self._salvar_cfg()
        core.auditar(self.usuario["login"], "saída do programa")
        if self.indice_busca:
            try:
                self.indice_busca.fechar()
            except Exception:
                pass
        shutil.rmtree(self.temp_dir, ignore_errors=True)
        self.root.destroy()


# --------------------------------------------------------------------------
# Inicio
# --------------------------------------------------------------------------

_mutex = None


def criar_mutex():
    """Sinaliza ao instalador que o programa esta aberto (ele espera fechar antes de atualizar)."""
    global _mutex
    if sys.platform.startswith("win"):
        try:
            _mutex = ctypes.windll.kernel32.CreateMutexW(None, False, MUTEX_NOME)
        except Exception:
            log.warning("Não foi possível criar o mutex", exc_info=True)


def main():
    configurar_log()
    log.info("Iniciando %s %s", APP_NOME, VERSAO)
    criar_mutex()
    updater.limpar_baixados()
    core.limpar_temporarios()
    if sys.platform.startswith("win"):
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # nitidez em telas de alta resolução
        except Exception:
            pass
    root = tk.Tk()
    root.withdraw()
    aplicar_tema(core.carregar_marca())
    aplicar_icone(root)
    # sem console no .exe: erros de callback vao para o erro.log (a App substitui depois)
    root.report_callback_exception = lambda t, v, tb: log.error("Erro na interface", exc_info=(t, v, tb))
    try:
        usuarios = core.Usuarios()
        login = TelaLogin(root, usuarios)
        root.wait_window(login)
        if not login.resultado:
            root.destroy()
            return
        root.deiconify()
        app = App(root, login.resultado, usuarios)
        pacote = next((a for a in sys.argv[1:] if a.lower().endswith(core.EXT_PACOTE)), None)
        if pacote and os.path.isfile(pacote):          # duplo clique no .bdpack
            root.after(800, lambda: app.importar_pacote(pacote))
        root.mainloop()
    except Exception as e:
        log.critical("Falha fatal", exc_info=True)
        try:
            messagebox.showerror(APP_NOME, f"Erro ao iniciar:\n{e}\n\nDetalhes em:\n{ARQUIVO_LOG}")
        finally:
            try:
                root.destroy()
            except Exception:
                pass


if __name__ == "__main__":
    main()
