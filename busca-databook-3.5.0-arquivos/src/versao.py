"""Versao do aplicativo - fonte unica.

Para publicar uma nova versao:
  1. altere VERSAO abaixo (formato X.Y.Z);
  2. faca commit e crie a tag correspondente:  git tag vX.Y.Z && git push --tags
  3. o GitHub Actions gera o instalador e cria a release como RASCUNHO;
  4. baixe o .sha256 do rascunho, assine no seu computador
     (python tools/assinar_release.py assinar <arquivo.sha256>),
     anexe o .sig ao rascunho e so entao publique.
O GitHub Actions confere se a tag bate com este arquivo antes de gerar o instalador.
"""

APP_NOME = "Busca Databook"
VERSAO = "3.5.0"
REPO_GITHUB = "AlexandreTavares75/busca-databook"

# Chave PUBLICA (Ed25519, base64) que confere a assinatura das atualizacoes.
# Gerada com:  python tools/assinar_release.py gerar-chave
# A chave PRIVADA correspondente nunca entra no repositorio.
CHAVE_PUBLICA_ATUALIZACAO = ""
