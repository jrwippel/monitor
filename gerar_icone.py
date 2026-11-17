"""Gera o arquivo de ícone (timesheet.ico) a partir do logo oficial do relógio.
Usado pelo PyInstaller (--icon) e pelo instalador. Rode uma vez:  python gerar_icone.py
"""
from PIL import Image, ImageDraw


def _hex_para_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def gerar_logo_relogio(tamanho=256):
    esc = 4
    s = tamanho * esc
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    def p(v):
        return v / 100 * s

    raio_fundo = p(48)
    cx = cy = s / 2
    grad = Image.new("RGB", (s, s))
    gpx = grad.load()
    r1, g1, b1 = _hex_para_rgb("#667eea")
    r2, g2, b2 = _hex_para_rgb("#764ba2")
    for y in range(s):
        for x in range(s):
            t = ((x / s) + (y / s)) / 2
            gpx[x, y] = (int(r1 + (r2 - r1) * t), int(g1 + (g2 - g1) * t), int(b1 + (b2 - b1) * t))
    mask = Image.new("L", (s, s), 0)
    ImageDraw.Draw(mask).ellipse([cx - raio_fundo, cy - raio_fundo, cx + raio_fundo, cy + raio_fundo], fill=255)
    img.paste(grad, (0, 0), mask)

    rf = p(35)
    d.ellipse([cx - rf, cy - rf, cx + rf, cy + rf], fill=(255, 255, 255, 242))

    roxo1 = _hex_para_rgb("#667eea")
    roxo2 = _hex_para_rgb("#764ba2")
    rm = p(2)
    for (mx, my) in [(50, 20), (80, 50), (50, 80), (20, 50)]:
        d.ellipse([p(mx) - rm, p(my) - rm, p(mx) + rm, p(my) + rm], fill=roxo1)
    d.line([p(50), p(50), p(50), p(35)], fill=roxo2, width=int(p(3)))
    d.line([p(50), p(50), p(65), p(50)], fill=roxo1, width=int(p(2.5)))
    rc = p(4)
    d.ellipse([cx - rc, cy - rc, cx + rc, cy + rc], fill=roxo2)
    d.polygon([(p(72), p(72)), (p(82), p(77)), (p(72), p(82))], fill=roxo1)
    return img.resize((tamanho, tamanho), Image.LANCZOS)


if __name__ == "__main__":
    logo = gerar_logo_relogio(256)
    logo.save("timesheet.ico", format="ICO",
             sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    logo.save("timesheet.png")  # versão PNG, útil para a página de download
    print("Gerado: timesheet.ico e timesheet.png")
