"""
Gera o icone do app (.ico) a partir da logomarca.

Extrai automaticamente o simbolo (a figura a esquerda do "OSL"), pinta de branco
e centraliza num quadrado azul #002060 de cantos arredondados.
A logo deve ter tracos escuros sobre fundo branco ou transparente.

Uso:  python tools/gerar_icone.py assets/logo_quadrada.png assets/icone.ico
"""

import sys
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw

AZUL = (0x00, 0x20, 0x60, 255)
TAMANHOS = [16, 20, 24, 32, 40, 48, 64, 128, 256]
BASE = 512
LANCZOS = getattr(Image, "Resampling", Image).LANCZOS


def _tinta(img):
    """Mascara suave (0-255) dos tracos escuros e opacos (bordas sem serrilhado)."""
    rgba = img.convert("RGBA")
    escuro = rgba.convert("L").point(lambda v: max(0, min(255, (230 - v) * 255 // 180)))
    return ImageChops.multiply(escuro, rgba.getchannel("A"))


def _faixas(tem_tinta, folga):
    """Agrupa posicoes com desenho em faixas continuas (tolerando 'folga' vazios)."""
    faixas, ini, fim, vazio = [], None, None, 0
    for i, tem in enumerate(tem_tinta):
        if tem:
            if ini is None:
                ini = i
            fim, vazio = i, 0
        elif ini is not None:
            vazio += 1
            if vazio > folga:
                faixas.append((ini, fim))
                ini, vazio = None, 0
    if ini is not None:
        faixas.append((ini, fim))
    return faixas


def extrair_simbolo(caminho):
    m = _tinta(Image.open(caminho))
    b = m.point(lambda v: 255 if v > 100 else 0)
    w, h = b.size
    linhas = _faixas([b.crop((0, y, w, y + 1)).getbbox() is not None for y in range(h)], h // 50)
    if not linhas:
        raise SystemExit("Nenhum desenho encontrado na imagem.")
    y0, y1 = max(linhas, key=lambda f: f[1] - f[0])          # bloco mais alto: simbolo | OSL
    colunas = _faixas([b.crop((x, y0, x + 1, y1 + 1)).getbbox() is not None for x in range(w)], w // 100)
    if len(colunas) < 2:
        raise SystemExit("Não consegui separar o símbolo do restante da logo.")
    x0, x1 = colunas[0]                                        # primeira figura a esquerda
    recorte = m.crop((x0, y0, x1 + 1, y1 + 1))
    return recorte.crop(recorte.getbbox())


def desenhar(simbolo, ocupacao, raio):
    tela = Image.new("RGBA", (BASE, BASE), (0, 0, 0, 0))
    ImageDraw.Draw(tela).rounded_rectangle((0, 0, BASE - 1, BASE - 1), radius=int(BASE * raio), fill=AZUL)
    escala = BASE * ocupacao / max(simbolo.size)
    s = simbolo.resize((max(1, round(simbolo.width * escala)), max(1, round(simbolo.height * escala))), LANCZOS)
    branco = Image.new("RGBA", s.size, (255, 255, 255, 255))
    tela.paste(branco, ((BASE - s.width) // 2, (BASE - s.height) // 2), s)
    return tela


def main():
    if len(sys.argv) != 3:
        raise SystemExit("Uso: python tools/gerar_icone.py <logo.png> <saida.ico>")
    origem, destino = Path(sys.argv[1]), Path(sys.argv[2])
    simbolo = extrair_simbolo(origem)
    grande = desenhar(simbolo, 0.70, 0.18)
    pequeno = desenhar(simbolo, 0.86, 0.12)        # tamanhos minusculos: simbolo maior
    imagens = [(pequeno if t <= 32 else grande).resize((t, t), LANCZOS) for t in TAMANHOS]
    destino.parent.mkdir(parents=True, exist_ok=True)
    imagens[-1].save(destino, format="ICO", sizes=[(t, t) for t in TAMANHOS], append_images=imagens[:-1])
    previa = destino.with_name("icone_preview.png")
    grande.resize((256, 256), LANCZOS).save(previa)
    print(f"Ícone gerado: {destino}\nPrévia: {previa}")


if __name__ == "__main__":
    main()