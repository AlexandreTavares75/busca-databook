"""
Assinatura das atualizacoes do Busca Databook (Ed25519).

Roda no computador do desenvolvedor, NUNCA no GitHub. Requer:  pip install cryptography

  1) Uma unica vez - criar o par de chaves:
       python tools/assinar_release.py gerar-chave
     Pede uma senha (protege a chave privada) e mostra a chave PUBLICA, que deve
     ser copiada para CHAVE_PUBLICA_ATUALIZACAO em src/versao.py.
     A chave privada fica em %USERPROFILE%\\BuscaDatabook-chave\\chave_privada.pem.
     Faca uma copia dela num pendrive guardado: sem ela nao ha como publicar
     atualizacoes que os programas instalados aceitem.

  2) A cada versao - assinar o .sha256 baixado do rascunho da release:
       python tools/assinar_release.py assinar BuscaDatabook-Setup-X.Y.Z.exe.sha256
     Gera BuscaDatabook-Setup-X.Y.Z.exe.sha256.sig ao lado. Anexe o .sig ao
     rascunho da release e publique.

  3) Opcional - conferir uma assinatura com a chave de src/versao.py:
       python tools/assinar_release.py verificar BuscaDatabook-Setup-X.Y.Z.exe.sha256
"""

import argparse
import base64
import getpass
import os
import re
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
CHAVE_PADRAO = Path(os.environ.get("USERPROFILE") or Path.home()) / "BuscaDatabook-chave" / "chave_privada.pem"
RE_LINHA = re.compile(rb"^[0-9a-f]{64}  BuscaDatabook-Setup-\d+\.\d+\.\d+\.exe\s*$")
SENHA_MIN = 12


def _crypto():
    try:
        from cryptography.hazmat.primitives import serialization
        from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
        return serialization, Ed25519PrivateKey
    except ImportError:
        sys.exit("Instale o componente antes:  pip install cryptography")


def _publica_b64(chave_privada) -> str:
    serialization, _ = _crypto()
    bruta = chave_privada.public_key().public_bytes(serialization.Encoding.Raw,
                                                    serialization.PublicFormat.Raw)
    return base64.b64encode(bruta).decode("ascii")


def _chave_do_programa() -> str:
    sys.path.insert(0, str(RAIZ / "src"))
    try:
        import versao
        return (getattr(versao, "CHAVE_PUBLICA_ATUALIZACAO", "") or "").strip()
    except Exception:
        return ""


def gerar_chave(args):
    serialization, Ed25519PrivateKey = _crypto()
    destino = Path(args.chave)
    if destino.exists():
        sys.exit(f"Já existe uma chave em {destino}. Por segurança, ela não será substituída.\n"
                 "Se quiser mesmo trocar de chave, mova o arquivo antigo para outro lugar antes.")
    try:
        destino.resolve().relative_to(RAIZ)
        sys.exit("Não grave a chave privada dentro da pasta do repositório. Escolha outra pasta.")
    except ValueError:
        pass
    senha = getpass.getpass(f"Crie uma senha para a chave privada (mínimo {SENHA_MIN} caracteres): ")
    if len(senha) < SENHA_MIN:
        sys.exit(f"Senha curta demais (mínimo {SENHA_MIN} caracteres).")
    if getpass.getpass("Repita a senha: ") != senha:
        sys.exit("As senhas não conferem.")
    chave = Ed25519PrivateKey.generate()
    pem = chave.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8,
                              serialization.BestAvailableEncryption(senha.encode("utf-8")))
    destino.parent.mkdir(parents=True, exist_ok=True)
    with open(destino, "xb") as f:
        f.write(pem)
    print(f"\nChave privada criada em: {destino}")
    print("  -> Faça AGORA uma cópia desse arquivo num pendrive e guarde a senha em lugar seguro.")
    print("  -> Nunca envie esse arquivo por e-mail nem o coloque no GitHub.\n")
    print("Chave PÚBLICA (copie para CHAVE_PUBLICA_ATUALIZACAO em src/versao.py):\n")
    print(f'CHAVE_PUBLICA_ATUALIZACAO = "{_publica_b64(chave)}"\n')


def _carregar_privada(caminho):
    serialization, _ = _crypto()
    caminho = Path(caminho)
    if not caminho.is_file():
        sys.exit(f"Chave privada não encontrada: {caminho}")
    senha = getpass.getpass("Senha da chave privada: ")
    try:
        return serialization.load_pem_private_key(caminho.read_bytes(), senha.encode("utf-8"))
    except (ValueError, TypeError):
        sys.exit("Senha incorreta ou arquivo de chave inválido.")


def _ler_sha256(caminho) -> bytes:
    caminho = Path(caminho)
    if not caminho.is_file():
        sys.exit(f"Arquivo não encontrado: {caminho}")
    conteudo = caminho.read_bytes()
    if not RE_LINHA.match(conteudo):
        sys.exit("Este não parece o .sha256 gerado pelo GitHub Actions "
                 "(esperado: '<hash>  BuscaDatabook-Setup-X.Y.Z.exe').")
    nome_no_arquivo = conteudo.split()[1].decode("ascii")
    if caminho.name != nome_no_arquivo + ".sha256":
        sys.exit(f"O nome do arquivo ({caminho.name}) não corresponde ao instalador "
                 f"citado dentro dele ({nome_no_arquivo}).")
    return conteudo


def assinar(args):
    conteudo = _ler_sha256(args.arquivo)
    chave = _carregar_privada(args.chave)
    publica = _publica_b64(chave)
    do_programa = _chave_do_programa()
    if do_programa and do_programa != publica:
        sys.exit("ATENÇÃO: esta chave privada NÃO corresponde à chave pública de src/versao.py.\n"
                 "Os programas instalados recusariam esta assinatura. Nada foi gravado.")
    if not do_programa:
        print("Aviso: src/versao.py ainda está sem CHAVE_PUBLICA_ATUALIZACAO.")
    assinatura = chave.sign(conteudo)
    chave.public_key().verify(assinatura, conteudo)          # autoteste
    saida = Path(str(args.arquivo) + ".sig")
    saida.write_text(base64.b64encode(assinatura).decode("ascii") + "\n", encoding="ascii")
    print(f"Assinatura gravada em: {saida}")
    print("Agora anexe esse .sig ao rascunho da release no GitHub e publique.")


def verificar(args):
    conteudo = _ler_sha256(args.arquivo)
    sig = Path(str(args.arquivo) + ".sig")
    if not sig.is_file():
        sys.exit(f"Assinatura não encontrada: {sig}")
    chave = args.publica or _chave_do_programa()
    if not chave:
        sys.exit("Informe a chave pública (--publica) ou preencha src/versao.py.")
    sys.path.insert(0, str(RAIZ / "src"))
    import updater
    try:
        updater.verificar_assinatura(conteudo, sig.read_text(encoding="ascii"), chave)
    except updater.ErroAtualizacao as e:
        sys.exit(f"FALHOU: {e}")
    print("OK: assinatura válida para esta chave pública.")


def main():
    ap = argparse.ArgumentParser(description="Assinatura das atualizações do Busca Databook")
    sub = ap.add_subparsers(dest="comando", required=True)
    g = sub.add_parser("gerar-chave", help="cria o par de chaves (uma única vez)")
    g.add_argument("--chave", default=str(CHAVE_PADRAO), help="onde gravar a chave privada")
    g.set_defaults(func=gerar_chave)
    a = sub.add_parser("assinar", help="assina o .sha256 de uma release")
    a.add_argument("arquivo", help="BuscaDatabook-Setup-X.Y.Z.exe.sha256")
    a.add_argument("--chave", default=str(CHAVE_PADRAO), help="arquivo da chave privada")
    a.set_defaults(func=assinar)
    v = sub.add_parser("verificar", help="confere a assinatura de um .sha256")
    v.add_argument("arquivo", help="BuscaDatabook-Setup-X.Y.Z.exe.sha256")
    v.add_argument("--publica", help="chave pública em base64 (padrão: a de src/versao.py)")
    v.set_defaults(func=verificar)
    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
