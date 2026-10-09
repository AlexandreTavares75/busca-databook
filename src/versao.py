"""Versao do aplicativo - fonte unica.

Para publicar uma nova versao:
  1. altere VERSAO abaixo (formato X.Y.Z);
  2. faca commit e crie a tag correspondente:  git tag vX.Y.Z && git push --tags
O GitHub Actions confere se a tag bate com este arquivo antes de gerar o instalador.
"""

APP_NOME = "Busca Databook"
VERSAO = "3.0.1"
REPO_GITHUB = "AlexandreTavares75/busca-databook"