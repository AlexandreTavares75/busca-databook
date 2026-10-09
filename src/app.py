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
AZUL = "#002060"
AZUL_HOVER = "#0E3A8A"
AZUL_TEXTO_SUAVE = "#AFC3E8"
FONTE_MENU = ("Segoe UI", 10)
FONTE_MENU_SEL = ("Segoe UI", 10, "bold")

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

Empresas
- O menu azul à esquerda mostra as pastas "DB_NOME" da pasta principal.
- Clique numa empresa para buscar só na pasta dela (e subpastas).
- "Todas" busca na pasta principal inteira.

O que é lido
- Conteúdo: Word, Excel, PowerPoint, PDF (inclusive escaneado, por OCR),
  textos, e arquivos dentro de .zip.
- Só o nome: imagens (PNG, JPEG...), MPP, DWG, DWF e demais formatos.

Atalhos
- Duplo clique ou Enter: abre o arquivo.
- Botão direito: abrir pasta, salvar cópia, copiar caminho.
- Ctrl+clique / Shift+clique: seleciona vários para salvar cópia de uma vez.
"""


# --------------------------------------------------------------------------
# Log de erros e permissoes
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

    def entrar(self):
        restante = self.bloqueado_ate - time.monotonic()
        if restante > 0:
            messagebox.showwarning(APP_NOME, f"Muitas tentativas. Aguarde {int(restante) + 1} segundos.", parent=self)
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
            self.falhas += 1
            log.warning("Falha de login (usuário informado: %r)", user.strip().lower()[:64])
            self.e_senha.delete(0, "end")
            if self.falhas >= TENTATIVAS_LOGIN:
                self.falhas = 0
                self.bloqueado_ate = time.monotonic() + BLOQUEIO_LOGIN_S
                messagebox.showerror(APP_NOME, f"Usuário ou senha incorretos.\n\nMuitas tentativas: "
                                     f"aguarde {BLOQUEIO_LOGIN_S} segundos.", parent=self)
            else:
                messagebox.showerror(APP_NOME, "Usuário ou senha incorretos.", parent=self)
            return
        log.info("Login: %s", r["login"])
        self.resultado = r
        self.destroy()

    def cancelar(self):
        self.resultado = None
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
        self.lista = ttk.Treeview(f, columns=("perfil",), height=8)
        self.lista.heading("#0", text="Usuário")
        self.lista.heading("perfil", text="Perfil")
        self.lista.column("#0", width=180)
        self.lista.column("perfil", width=120)
        self.lista.grid(row=0, column=0, rowspan=4, padx=(0, 8))
        ttk.Button(f, text="Novo usuário", command=self.novo).grid(row=0, column=1, sticky="ew", pady=2)
        ttk.Button(f, text="Trocar senha", command=self.senha).grid(row=1, column=1, sticky="ew", pady=2)
        ttk.Button(f, text="Excluir", command=self.excluir).grid(row=2, column=1, sticky="ew", pady=2)
        ttk.Button(f, text="Fechar", command=self.destroy).grid(row=3, column=1, sticky="sew", pady=2)
        self.recarregar()
        self.transient(master)
        self.grab_set()

    def recarregar(self):
        self.lista.delete(*self.lista.get_children())
        for login, admin in self.usuarios.listar():
            self.lista.insert("", "end", iid=login, text=login, values=("Administrador" if admin else "Usuário",))

    def selecionado(self):
        s = self.lista.selection()
        return s[0] if s else None

    def novo(self):
        login = simpledialog.askstring("Novo usuário", "Nome de usuário:", parent=self)
        if not login:
            return
        senha = simpledialog.askstring("Novo usuário", f"Senha (mínimo {core.SENHA_MIN} caracteres):",
                                       show="•", parent=self)
        if senha is None:
            return
        admin = messagebox.askyesno("Novo usuário", "Este usuário será administrador?", parent=self)
        try:
            self.usuarios.criar(login, senha, admin)
            log.info("Usuário criado por %s: %s (admin=%s)", self.atual, login.strip().lower(), admin)
        except Exception as e:
            messagebox.showerror("Usuários", f"Não foi possível criar: {e}", parent=self)
        self.recarregar()

    def senha(self):
        s = self.selecionado()
        if not s:
            return
        nova = simpledialog.askstring("Trocar senha", f"Nova senha para {s}:", show="•", parent=self)
        if nova:
            try:
                self.usuarios.trocar_senha(s, nova)
                log.info("Senha de %s alterada por %s", s, self.atual)
                messagebox.showinfo("Usuários", "Senha alterada.", parent=self)
            except Exception as e:
                messagebox.showerror("Usuários", str(e), parent=self)

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
# Menu lateral de empresas
# --------------------------------------------------------------------------

class MenuEmpresas(tk.Frame):
    """Barra azul a esquerda: "Todas" + um botao por pasta DB_NOME."""

    def __init__(self, master, ao_escolher):
        super().__init__(master, bg=AZUL, width=210)
        self.pack_propagate(False)
        self.ao_escolher = ao_escolher
        self.botoes = {}            # caminho (ou None para "Todas") -> Label
        self.selecionado = None

        tk.Label(self, text="EMPRESAS", bg=AZUL, fg=AZUL_TEXTO_SUAVE, anchor="w",
                 font=("Segoe UI", 9, "bold"), padx=16, pady=14).pack(fill="x")

        area = tk.Frame(self, bg=AZUL)
        area.pack(fill="both", expand=True)
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
        tk.Frame(self.lista, bg=AZUL_HOVER, height=1).pack(fill="x", padx=12, pady=6)
        if not empresas:
            tk.Label(self.lista, text='Nenhuma pasta "DB_..."\nna pasta principal.', bg=AZUL,
                     fg=AZUL_TEXTO_SUAVE, justify="left", anchor="w", padx=16,
                     font=("Segoe UI", 9)).pack(fill="x", pady=4)
        for nome, caminho in empresas:
            self._botao(nome, caminho, nome)
        self.marcar(None)

    def _botao(self, texto, chave, nome):
        b = tk.Label(self.lista, text=texto, bg=AZUL, fg="white", anchor="w", padx=16, pady=7,
                     font=FONTE_MENU, cursor="hand2", takefocus=1,
                     highlightthickness=1, highlightbackground=AZUL, highlightcolor="white")
        b.pack(fill="x")
        acionar = lambda e=None: self._clicar(chave, nome)
        b.bind("<Button-1>", acionar)
        b.bind("<Return>", acionar)
        b.bind("<space>", acionar)
        b.bind("<Enter>", lambda e: chave != self.selecionado and b.config(bg=AZUL_HOVER))
        b.bind("<Leave>", lambda e: chave != self.selecionado and b.config(bg=AZUL))
        self.botoes[chave] = b

    def _clicar(self, chave, nome):
        self.marcar(chave)
        self.ao_escolher(chave, nome)

    def marcar(self, chave):
        """Destaca o botao selecionado (fundo branco, texto azul)."""
        if chave not in self.botoes:
            chave = None
        self.selecionado = chave
        for k, b in self.botoes.items():
            if k == chave:
                b.config(bg="white", fg=AZUL, font=FONTE_MENU_SEL)
            else:
                b.config(bg=AZUL, fg="white", font=FONTE_MENU)


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
        self.trabalhando = False
        self.atualizando = False
        self.cancelar_flag = threading.Event()
        self.indice_busca = None
        self.resultados = []
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
        arq.add_command(label="Sair", command=self.sair)
        m.add_cascade(label="Arquivo", menu=arq)
        if self.usuario["admin"]:
            m.add_command(label="Usuários", command=lambda: TelaUsuarios(self.root, self.usuarios, self.usuario["login"]))
        else:
            m.add_command(label="Trocar minha senha", command=self.trocar_minha_senha)
        aj = tk.Menu(m, tearoff=0)
        aj.add_command(label="Como buscar", command=lambda: messagebox.showinfo("Como buscar", DICAS))
        aj.add_command(label="Verificar atualizações", command=lambda: self._verificar_atualizacao(manual=True))
        self.var_auto_upd = tk.BooleanVar(value=self.cfg.get("verificar_atualizacoes", True))
        aj.add_checkbutton(label="Verificar atualizações ao iniciar", variable=self.var_auto_upd,
                           command=self._salvar_cfg)
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
        self.bt_indice.grid(row=0, column=3, padx=2)

        self.lbl_info = ttk.Label(topo, text="", foreground="#555")
        self.lbl_info.grid(row=1, column=1, sticky="w", padx=6, pady=(2, 8))

        ttk.Label(topo, text="Buscar:", style="Titulo.TLabel").grid(row=2, column=0, sticky="w")
        self.var_termo = tk.StringVar()
        self.e_termo = ttk.Entry(topo, textvariable=self.var_termo, font=("Segoe UI", 11))
        self.e_termo.grid(row=2, column=1, sticky="ew", padx=6)
        self.e_termo.bind("<Return>", lambda e: self.buscar())
        ttk.Button(topo, text="Buscar", command=self.buscar).grid(row=2, column=2, padx=2)
        self.var_filtro = tk.StringVar(value="Todos os tipos")
        cb = ttk.Combobox(topo, textvariable=self.var_filtro, values=list(FILTROS), state="readonly", width=18)
        cb.grid(row=2, column=3, padx=2)
        cb.bind("<<ComboboxSelected>>", lambda e: self._mostrar_resultados())

        self.lbl_escopo = ttk.Label(topo, text="", style="Escopo.TLabel")
        self.lbl_escopo.grid(row=3, column=1, sticky="w", padx=6, pady=(6, 0))
        self.var_ocr = tk.BooleanVar(value=self.cfg.get("ocr", True))
        ttk.Checkbutton(topo, text="Ler PDFs escaneados (OCR) ao indexar", variable=self.var_ocr,
                        command=self._salvar_cfg).grid(row=3, column=2, columnspan=2, sticky="e", pady=(6, 0))

        # resultados
        meio = ttk.Frame(principal, padding=(10, 4))
        meio.pack(fill="both", expand=True)
        cols = ("tipo", "onde", "trecho", "pasta")
        self.tv = ttk.Treeview(meio, columns=cols, selectmode="extended")
        self.tv.heading("#0", text="Arquivo", command=lambda: self._ordenar("#0"))
        self.tv.heading("tipo", text="Tipo", command=lambda: self._ordenar("tipo"))
        self.tv.heading("onde", text="Achado no", command=lambda: self._ordenar("onde"))
        self.tv.heading("trecho", text="Trecho")
        self.tv.heading("pasta", text="Pasta", command=lambda: self._ordenar("pasta"))
        self.tv.column("#0", width=280, stretch=False)
        self.tv.column("tipo", width=55, stretch=False, anchor="center")
        self.tv.column("onde", width=80, stretch=False, anchor="center")
        self.tv.column("trecho", width=430)
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
            self.tv.delete(*self.tv.get_children())
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
        self.e_termo.focus_set()

    def escolher_pasta(self):
        p = filedialog.askdirectory(title="Escolha a pasta principal", initialdir=self.var_pasta.get() or None)
        if p:
            self.var_pasta.set(os.path.normpath(p))
            self._pasta_mudou()

    def _pasta_mudou(self, silencioso=False):
        if self.trabalhando:
            messagebox.showinfo(APP_NOME, "Aguarde o fim da indexação (ou cancele) antes de trocar de pasta.")
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
        self.tv.delete(*self.tv.get_children())
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

    def atualizar_indice(self):
        if self.trabalhando:
            return
        pasta = self.var_pasta.get().strip()
        if not pasta or not os.path.isdir(pasta):
            messagebox.showwarning(APP_NOME, "Escolha uma pasta válida primeiro.")
            return
        if self.var_ocr.get() and not core.configurar_ocr():
            messagebox.showinfo(APP_NOME, "O OCR (Tesseract) não está disponível.\n\n"
                                "Os PDFs escaneados serão indexados apenas pelo nome. "
                                "Reinstale o programa para restaurar o OCR.")
        self.trabalhando = True
        self.cancelar_flag.clear()
        self.bt_indice.config(state="disabled")
        self.prog.pack(side="right", padx=4)
        self.bt_cancelar.pack(side="right", padx=4)
        usar_ocr = self.var_ocr.get()
        threading.Thread(target=self._tarefa_indexar, args=(pasta, usar_ocr), daemon=True).start()

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

    def _tarefa_indexar(self, pasta, usar_ocr):
        idx = core.Indice(pasta)
        try:
            r = idx.atualizar(usar_ocr=usar_ocr,
                              progresso=lambda m, i, n: self.fila.put(("prog", m, i, n)),
                              cancelado=self.cancelar_flag.is_set)
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
                elif tipo.startswith("upd_"):
                    self._tratar_atualizacao(msg)
        except queue.Empty:
            pass
        except Exception:
            log.exception("Erro ao processar mensagens da interface")
        if ultimo_prog and self.trabalhando:
            _, m, i, n = ultimo_prog
            txt = f"{i}/{n} · {m}" if n else m
            self.lbl_status.config(text=txt[:120])
            if n:
                self.prog.config(mode="determinate", maximum=n, value=i)
            else:
                self.prog.config(mode="indeterminate")
                self.prog.step(5)
        self.root.after(150, self._processar_fila)

    def _fim_indexacao(self, r, erro=None):
        self.trabalhando = False
        self.bt_indice.config(state="normal")
        self.prog.pack_forget()
        self.bt_cancelar.pack_forget()
        self._atualizar_info()
        self._carregar_empresas(manter=True)   # pode ter surgido uma pasta DB_ nova
        if erro:
            self.lbl_status.config(text="Erro ao indexar.")
            messagebox.showerror(APP_NOME, f"Erro ao indexar:\n{erro}\n\nDetalhes em:\n{ARQUIVO_LOG}")
            return
        txt = f"Índice atualizado: {r['lidos']} lidos/relidos, {r['removidos']} removidos"
        if r["erros"]:
            txt += f", {r['erros']} não puderam ser lidos (Arquivo > Arquivos que não puderam ser lidos)"
        if r["cancelado"]:
            txt = "Indexação cancelada. " + txt
        self.lbl_status.config(text=txt + ".")
        if self.var_termo.get().strip():
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
            self.resultados = self.indice_busca.buscar(termo, pasta=self.empresa)
        except Exception as e:
            log.exception("Erro na busca")
            messagebox.showerror(APP_NOME, f"Não foi possível buscar: {e}")
            return
        self._mostrar_resultados()

    def _mostrar_resultados(self):
        self.tv.delete(*self.tv.get_children())
        filtro = FILTROS.get(self.var_filtro.get())
        visiveis = 0
        for i, r in enumerate(self.resultados):
            if filtro and r["ext"] not in filtro:
                continue
            nome = r["nome"] + (f"  ›  {r['interno']}" if r["interno"] else "")
            self.tv.insert("", "end", iid=str(i), text=nome,
                           values=(r["ext"].upper(), r["onde"], r["trecho"], r["pasta"]))
            visiveis += 1
        termo = self.var_termo.get().strip()
        if not self.trabalhando and termo:
            onde = self._texto_escopo()
            self.lbl_status.config(text=f"{visiveis} resultado(s) para “{termo}” em {onde}." if visiveis
                                   else f"Nada encontrado para “{termo}” em {onde}.")

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
        return [self.resultados[int(k)] for k in self.tv.selection()]

    def _menu_contexto(self, ev):
        linha = self.tv.identify_row(ev.y)
        if linha:
            if linha not in self.tv.selection():
                self.tv.selection_set(linha)
            self.menu_ctx.tk_popup(ev.x_root, ev.y_root)

    def abrir(self):
        for r in self._selecionados()[:10]:
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

    def abrir_pasta(self):
        sel = self._selecionados()
        if sel:
            mostrar_na_pasta(sel[0]["caminho"])

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

    def trocar_minha_senha(self):
        atual = simpledialog.askstring("Trocar senha", "Senha atual:", show="•", parent=self.root)
        if atual is None:
            return
        if not self.usuarios.validar(self.usuario["login"], atual):
            messagebox.showerror(APP_NOME, "Senha atual incorreta.")
            return
        nova = simpledialog.askstring("Trocar senha", f"Nova senha (mínimo {core.SENHA_MIN} caracteres):",
                                      show="•", parent=self.root)
        if not nova:
            return
        if simpledialog.askstring("Trocar senha", "Repita a nova senha:", show="•", parent=self.root) != nova:
            messagebox.showerror(APP_NOME, "As senhas não conferem.")
            return
        try:
            self.usuarios.trocar_senha(self.usuario["login"], nova)
            log.info("Usuário %s trocou a própria senha", self.usuario["login"])
            messagebox.showinfo(APP_NOME, "Senha alterada.")
        except Exception as e:
            messagebox.showerror(APP_NOME, str(e))

    # ---- atualizacao automatica ------------------------------------------

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
        self.cfg["ocr"] = bool(self.var_ocr.get()) if hasattr(self, "var_ocr") else True
        if hasattr(self, "var_auto_upd"):
            self.cfg["verificar_atualizacoes"] = bool(self.var_auto_upd.get())
        salvar_config(self.cfg)

    def sair(self, forcar=False):
        if self.trabalhando and not forcar and not messagebox.askyesno(
                APP_NOME, "A indexação está em andamento. Sair mesmo assim?\n"
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