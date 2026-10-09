"""
Atualizacao automatica via GitHub Releases.

Fluxo:
  1. consulta a ultima release publicada (api.github.com);
  2. se a versao for maior que a atual, localiza o instalador
     "BuscaDatabook-Setup-X.Y.Z.exe" e o arquivo "...exe.sha256";
  3. baixa o instalador para %LOCALAPPDATA%\\BuscaDatabook\\atualizacoes,
     calculando o SHA-256 durante o download;
  4. so instala se o hash bater com o .sha256 publicado (e com o digest que o
     proprio GitHub calcula no upload, quando disponivel);
  5. confere o hash de novo imediatamente antes de executar.

Somente HTTPS e somente hosts do GitHub, inclusive nos redirecionamentos.
"""

import hashlib
import json
import logging
import os
import re
import ssl
import subprocess
import sys
import urllib.request
from pathlib import Path
from urllib.parse import urlparse

import core
from versao import REPO_GITHUB, VERSAO

log = logging.getLogger(__name__)

URL_API = f"https://api.github.com/repos/{REPO_GITHUB}/releases/latest"
TIMEOUT_S = 30
TAMANHO_MAX = 800 * 1024 * 1024        # trava contra downloads absurdos
RE_VERSAO = re.compile(r"^v?(\d+)\.(\d+)\.(\d+)$")
RE_SHA256 = re.compile(r"\b([0-9a-fA-F]{64})\b")


class ErroAtualizacao(Exception):
    pass


# --------------------------------------------------------------------------
# HTTP restrito ao GitHub
# --------------------------------------------------------------------------

def _host_permitido(host: str) -> bool:
    host = (host or "").lower()
    return host in ("api.github.com", "github.com") or host.endswith(".githubusercontent.com")


def _checar_url(url: str):
    u = urlparse(url)
    if u.scheme != "https" or not _host_permitido(u.hostname):
        raise ErroAtualizacao(f"Endereço de download não permitido: {u.scheme}://{u.hostname}")


class _RedirecionamentoSeguro(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        _checar_url(newurl)  # bloqueia antes de seguir o redirecionamento
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def _abrir(url: str, aceitar="application/octet-stream"):
    _checar_url(url)
    contexto = ssl.create_default_context()   # valida certificado (repositorio do Windows)
    opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=contexto),
                                         _RedirecionamentoSeguro())
    req = urllib.request.Request(url, headers={
        "User-Agent": f"BuscaDatabook/{VERSAO}",
        "Accept": aceitar,
    })
    resp = opener.open(req, timeout=TIMEOUT_S)
    _checar_url(resp.geturl())
    return resp


# --------------------------------------------------------------------------
# Versoes
# --------------------------------------------------------------------------

def versao_tupla(txt: str):
    m = RE_VERSAO.match((txt or "").strip())
    if not m:
        raise ErroAtualizacao(f"Versão em formato inesperado: {txt!r}")
    return tuple(int(x) for x in m.groups())


def rodando_instalado() -> bool:
    """Atualizacao so faz sentido no programa instalado (.exe), nao no codigo-fonte."""
    return bool(getattr(sys, "frozen", False)) and sys.platform.startswith("win")


def pasta_atualizacoes() -> Path:
    p = core.pasta_dados_app() / "atualizacoes"
    p.mkdir(exist_ok=True)
    return p


def limpar_baixados():
    """Remove instaladores de atualizacoes anteriores."""
    try:
        for f in pasta_atualizacoes().iterdir():
            if f.is_file():
                f.unlink()
    except OSError:
        log.debug("Não foi possível limpar a pasta de atualizações", exc_info=True)


# --------------------------------------------------------------------------
# Consulta / download / instalacao
# --------------------------------------------------------------------------

def buscar_atualizacao():
    """Devolve dict com a nova versao, ou None se ja estiver atualizado."""
    with _abrir(URL_API, "application/vnd.github+json") as r:
        dados = json.loads(r.read(2_000_000).decode("utf-8"))

    nova = versao_tupla(dados.get("tag_name", ""))
    if nova <= versao_tupla(VERSAO):
        return None

    versao_txt = ".".join(map(str, nova))
    nome = f"BuscaDatabook-Setup-{versao_txt}.exe"
    assets = {a.get("name"): a for a in dados.get("assets", [])}
    exe, sha = assets.get(nome), assets.get(nome + ".sha256")
    if not exe or not sha:
        raise ErroAtualizacao(
            f"A versão {versao_txt} foi publicada sem o instalador ou sem o arquivo de "
            "verificação SHA-256. Por segurança, a atualização não será feita.")
    return {
        "versao": versao_txt,
        "nome": nome,
        "notas": (dados.get("body") or "").strip()[:1500],
        "url": exe["browser_download_url"],
        "tamanho": int(exe.get("size") or 0),
        "url_sha256": sha["browser_download_url"],
        "digest_github": (exe.get("digest") or "").lower(),
    }


def _hash_publicado(info) -> str:
    with _abrir(info["url_sha256"]) as r:
        texto = r.read(4096).decode("ascii", "replace")
    m = RE_SHA256.search(texto)
    if not m:
        raise ErroAtualizacao("O arquivo de verificação SHA-256 está vazio ou inválido.")
    esperado = m.group(1).lower()
    dg = info.get("digest_github", "")
    if dg.startswith("sha256:") and dg[7:] != esperado:
        raise ErroAtualizacao("O SHA-256 publicado não confere com o calculado pelo GitHub.")
    return esperado


def sha256_arquivo(caminho: Path) -> str:
    h = hashlib.sha256()
    with open(caminho, "rb") as f:
        for bloco in iter(lambda: f.read(1 << 20), b""):
            h.update(bloco)
    return h.hexdigest()


def baixar_e_verificar(info, progresso=None):
    """Baixa o instalador e confere o SHA-256. Devolve (caminho, sha256)."""
    if info["tamanho"] > TAMANHO_MAX:
        raise ErroAtualizacao("Instalador maior que o limite permitido.")
    esperado = _hash_publicado(info)

    destino = pasta_atualizacoes() / info["nome"]
    parcial = destino.with_name(destino.name + ".parcial")
    h = hashlib.sha256()
    baixado = 0
    try:
        with _abrir(info["url"]) as r, open(parcial, "wb") as f:
            total = int(r.headers.get("Content-Length") or info["tamanho"] or 0)
            while True:
                bloco = r.read(1 << 20)
                if not bloco:
                    break
                baixado += len(bloco)
                if baixado > TAMANHO_MAX:
                    raise ErroAtualizacao("Download excedeu o tamanho máximo permitido.")
                h.update(bloco)
                f.write(bloco)
                if progresso:
                    progresso(baixado, total)
        obtido = h.hexdigest()
        if obtido != esperado:
            log.error("SHA-256 divergente: esperado %s, obtido %s", esperado, obtido)
            raise ErroAtualizacao(
                "O arquivo baixado não passou na verificação de segurança (SHA-256 diferente). "
                "A atualização foi cancelada.")
        os.replace(parcial, destino)
    finally:
        if parcial.exists():
            try:
                parcial.unlink()
            except OSError:
                pass
    log.info("Atualização %s baixada e verificada (%s)", info["versao"], esperado)
    return destino, esperado


def instalar(caminho: Path, sha256_esperado: str):
    """Confere o hash novamente e inicia o instalador em modo silencioso.
    O chamador deve encerrar o programa logo em seguida."""
    if sha256_arquivo(caminho) != sha256_esperado:
        raise ErroAtualizacao("O instalador foi alterado depois do download. Atualização cancelada.")
    args = [str(caminho), "/SILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/RELAUNCH"]
    flags = 0
    if sys.platform.startswith("win"):
        flags = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
    subprocess.Popen(args, creationflags=flags, close_fds=True)
    log.info("Instalador iniciado: %s", caminho.name)