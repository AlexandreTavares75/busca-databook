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

import core
import updater
from versao import APP_NOME, REPO_GITHUB, VERSAO

log = logging.getLogger("app")

CONFIG = core.pasta_dados_app() / "config.json"
ARQUIVO_LOG = core.pasta_dados_app() / "erro.log"
MUTEX_NOME = "BuscaDatabook_Instancia"   # o instalador espera o programa fechar por este nome
INTERVALO_VERIFICACAO_S = 24 * 3600      # verifica atualizacoes no maximo 1x por dia
TENTATIVAS_LOGIN = 5
BLOQUEIO_LOGIN_S = 30

# Menu lateral
LARGURA_MENU = 220
AZUL = "#002060"
AZUL_TEXTO_SUAVE = "#AFC3E8"
BOTAO_COR = "#0B3A8C"            # botao em relevo, um tom acima do fundo
BOTAO_HOVER = "#1A4FA8"
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
    zf = zipfile.ZipFile(caminho_zip)
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
            linha = 5

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
            except Exception as e:
                messagebox.showerror(APP_NOME, str(e), parent=self)
                return
        r = self.usuarios.validar(user, senha)
        if not r:
            log.warning("Falha de login (usuário informado: %r)", user.strip().lower()[:64])
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
        if not r["tem_codigo"]:
            # todo usuario recebe um codigo de recuperacao (inclusive os que ja existiam)
            mostrar_codigo(self, self.usuarios.gerar_codigo(r["login"]))
            log.info("Código de recuperação gerado para %s", r["login"])

        log.info("Login: %s", r["login"])
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

    def __init__(self, master, usuarios, login, obrigatoria=False):
        super().__init__(master)
        self.usuarios, self.login, self.ok = usuarios, login, False
        self.title("Criar nova senha")
        self.resizable(False, False)
        f = ttk.Frame(self, padding=18)
        f.pack(fill="both", expand=True)
        msg = (f"Olá, {login}. Sua senha atual é provisória.\nCrie agora a sua senha pessoal para continuar."
               if obrigatoria else f"Nova senha para {login}.")
        ttk.Label(f, text=msg, justify="left").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(f, text="Nova senha:").grid(row=1, column=0, sticky="e", padx=4, pady=4)
        self.e1 = ttk.Entry(f, width=28, show="•")
        self.e1.grid(row=1, column=1, pady=4)
        ttk.Label(f, text="Repita a senha:").grid(row=2, column=0, sticky="e", padx=4, pady=4)
        self.e2 = ttk.Entry(f, width=28, show="•")
        self.e2.grid(row=2, column=1, pady=4)
        ttk.Label(f, text=f"Mínimo de {core.SENHA_MIN} caracteres.", foreground="#666").grid(
            row=3, column=1, sticky="w")
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
                         relief="solid", bd=1, readonlybackground="#F2F5FB", fg=AZUL)
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
        ttk.Label(f, text="Código no formato XXXX-XXXX-XXXX", foreground="#666").grid(row=5, column=1, sticky="w")
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
            extra = (f"\n\nMuitas tentativas: aguarde {BLOQUEIO_LOGIN_S} segundos."
                     if self.login_tela.registrar_falha() else "")
            messagebox.showerror(APP_NOME, "Usuário ou código de recuperação incorretos.\n"
                                           "Se não tiver o código, peça ao administrador para "
                                           "redefinir a sua senha." + extra, parent=self)
            return
        log.info("Senha recuperada com código: %s", login.strip().lower())
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
        self.tv = ttk.Treeview(f, columns=("data", "tamanho", "pasta"), selectmode="extended")
        self.tv.heading("#0", text="Arquivo")
        self.tv.heading("data", text="Modificado em")
        self.tv.heading("tamanho", text="Tamanho")
        self.tv.heading("pasta", text="Pasta")
        self.tv.column("#0", width=280)
        self.tv.column("data", width=120, anchor="center", stretch=False)
        self.tv.column("tamanho", width=80, anchor="e", stretch=False)
        self.tv.column("pasta", width=460)
        self.tv.pack(fill="both", expand=True)
        for i, r in enumerate(self.itens):
            marca = "★ " if (r["caminho"], r["interno"]) == (principal["caminho"], principal["interno"]) else "    "
            nome = r["nome"] + (f"  ›  {r['interno']}" if r["interno"] else "")
            self.tv.insert("", "end", iid=str(i), text=marca + nome,
                           values=(formatar_data(r.get("mtime")), formatar_tamanho(r.get("tamanho", 0)), r["pasta"]))
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
            b.config(bg="white", fg=AZUL, font=FONTE_MENU_SEL, relief="sunken")
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

        root.title(f"{APP_NOME} {VERSAO}  —  {usuario['login']}")
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
        st.configure("Escopo.TLabel", font=("Segoe UI", 10, "bold"), foreground=AZUL)

    def _menu(self):
        m = tk.Menu(self.root)
        arq = tk.Menu(m, tearoff=0)
        arq.add_command(label="Escolher pasta...", command=self.escolher_pasta)
        arq.add_command(label="Atualizar índice", command=self.atualizar_indice)
        arq.add_command(label="Refazer índice do zero", command=self.refazer_indice)
        arq.add_command(label="Recarregar lista de empresas", command=lambda: self._carregar_empresas(manter=True))
        arq.add_command(label="Arquivos que não puderam ser lidos", command=self.ver_erros)
        arq.add_separator()
        arq.add_command(label="Exportar índice...", command=self.exportar_indice)
        arq.add_command(label="Importar índice...", command=self.importar_indice)
        arq.add_separator()
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

        topo = ttk.Frame(principal, padding=(10, 10, 10, 4))
        topo.pack(fill="x")
        topo.columnconfigure(1, weight=1)

        ttk.Label(topo, text="Pasta principal:", style="Titulo.TLabel").grid(row=0, column=0, sticky="w")
        self.var_pasta = tk.StringVar()
        self.cb_pasta = ttk.Combobox(topo, textvariable=self.var_pasta, values=self.cfg.get("recentes", []))
        self.cb_pasta.grid(row=0, column=1, sticky="ew", padx=6)
        self.cb_pasta.bind("<<ComboboxSelected>>", lambda e: self._pasta_mudou())
        self.cb_pasta.bind("<Return>", lambda e: self._pasta_mudou())
        ttk.Button(topo, text="Procurar...", command=self.escolher_pasta).grid(row=0, column=2, padx=2)
        self.bt_indice = ttk.Button(topo, text="Atualizar índice", command=self.atualizar_indice)
        self.bt_indice.grid(row=0, column=3, columnspan=2, sticky="ew", padx=2)

        self.lbl_info = ttk.Label(topo, text="", foreground="#555")
        self.lbl_info.grid(row=1, column=1, sticky="w", padx=6, pady=(2, 8))

        ttk.Label(topo, text="Buscar:", style="Titulo.TLabel").grid(row=2, column=0, sticky="w")
        self.var_termo = tk.StringVar()
        self.e_termo = ttk.Entry(topo, textvariable=self.var_termo, font=("Segoe UI", 11))
        self.e_termo.grid(row=2, column=1, sticky="ew", padx=6)
        self.e_termo.bind("<Return>", lambda e: self.buscar())
        self.root.bind("<Escape>", lambda e: self.limpar_busca())
        ttk.Button(topo, text="Buscar", command=self.buscar).grid(row=2, column=2, padx=2)
        ttk.Button(topo, text="Limpar", command=self.limpar_busca).grid(row=2, column=3, padx=2)
        self.var_filtro = tk.StringVar(value="Todos os tipos")
        cb = ttk.Combobox(topo, textvariable=self.var_filtro, values=list(FILTROS), state="readonly", width=18)
        cb.grid(row=2, column=4, padx=2)
        cb.bind("<<ComboboxSelected>>", lambda e: self._mostrar_resultados())

        linha3 = ttk.Frame(topo)
        linha3.grid(row=3, column=1, sticky="w", padx=6, pady=(6, 0))
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
                        command=self._ocultar_mudou).grid(row=3, column=2, columnspan=3, sticky="e", pady=(6, 0))

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
        self.tv.column("trecho", width=410)
        self.tv.column("pasta", width=300)
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
        args = (pasta, self.var_ocr.get(), self.var_nuvem.get())
        threading.Thread(target=self._tarefa_indexar, args=args, daemon=True).start()

    def refazer_indice(self):
        pasta = self.var_pasta.get().strip()
        if not pasta or self.trabalhando:
            return
        if not messagebox.askyesno(APP_NOME, "Apagar o índice desta pasta e ler todos os arquivos novamente?"):
            return
        if self.indice_busca:
            self.indice_busca.fechar()
            self.indice_busca = None
        db = core.caminho_indice(pasta)
        for suf in ("", "-wal", "-shm"):
            try:
                os.remove(str(db) + suf)
            except OSError:
                pass
        self.indice_busca = core.Indice(pasta)
        self.atualizar_indice()

    def _tarefa_indexar(self, pasta, usar_ocr, liberar_nuvem):
        idx = core.Indice(pasta)
        try:
            r = idx.atualizar(usar_ocr=usar_ocr,
                              progresso=lambda m, i, n: self.fila.put(("prog", m, i, n)),
                              cancelado=self.cancelar_flag.is_set,
                              liberar_nuvem=liberar_nuvem)
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
                if tipo == "prog":
                    ultimo_prog = msg
                elif tipo == "fim":
                    self._fim_indexacao(msg[1])
                elif tipo == "erro":
                    self._fim_indexacao(None, msg[1])
                elif tipo.startswith("exp_") or tipo.startswith("imp_"):
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
            txt = "Indexação cancelada. " + txt
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
        w.title("Arquivos que não puderam ser lidos")
        w.geometry("900x400")
        t = tk.Text(w, wrap="none")
        t.pack(fill="both", expand=True)
        for c, s in erros:
            t.insert("end", f"{c}\n    {s}\n")
        t.config(state="disabled")

    # ---- exportar / importar indice --------------------------------------

    def exportar_indice(self):
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
                self.fila.put(("exp_fim", idx.exportar(destino, pastas, fora), destino))
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
                self.fila.put(("imp_fim", core.importar_indice(arq, nova), nova))
            except Exception as e:
                log.exception("Erro ao importar índice")
                self.fila.put(("imp_erro", str(e)))

        threading.Thread(target=tarefa, daemon=True).start()

    def _fim_exportar_importar(self, msg):
        self._terminar_tarefa()
        tipo = msg[0]
        if tipo == "exp_fim":
            _, n, destino = msg
            self.lbl_status.config(text=f"Índice exportado: {n} arquivos.")
            messagebox.showinfo(APP_NOME, f"{n:,} arquivos exportados para:\n{destino}".replace(",", ".")
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
                if not os.path.exists(r["caminho"]):
                    messagebox.showwarning(APP_NOME, f"Arquivo não encontrado (movido ou apagado?):\n{r['caminho']}\n\n"
                                           "Atualize o índice.")
                    continue
                if r["interno"]:
                    d = tempfile.mkdtemp(dir=self.temp_dir)
                    abrir_no_sistema(extrair_do_zip(r["caminho"], r["interno"], d))
                else:
                    abrir_no_sistema(r["caminho"])
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
                    shutil.copy2(r["caminho"], alvo)
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
    if sys.platform.startswith("win"):
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # nitidez em telas de alta resolução
        except Exception:
            pass
    root = tk.Tk()
    root.withdraw()
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
        App(root, login.resultado, usuarios)
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
