#!/usr/bin/env python3
"""Genera le icone di reMNote (PNG e ICO) in windows/assets/. Richiede Pillow."""
import os
from PIL import Image, ImageDraw

OUT = os.path.join(os.path.dirname(__file__), "..", "windows", "assets")
S = 1024
TEAL = (43, 131, 146, 255)
DARK = (20, 67, 77, 255)
ORANGE = (242, 106, 33, 255)
WHITE = (253, 253, 253, 255)


def draw_icon():
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    # fondo
    d.rounded_rectangle((28, 28, S - 28, S - 28), radius=220, fill=TEAL, outline=DARK, width=26)
    # scheda "note"
    card = (230, 190, 730, 850)
    d.rounded_rectangle(card, radius=70, fill=WHITE, outline=DARK, width=22)
    # testata arancione
    d.rounded_rectangle((card[0] + 11, card[1] + 11, card[2] - 11, card[1] + 160), radius=60,
                        fill=ORANGE, corners=(True, True, False, False))
    d.rounded_rectangle((card[0] + 70, card[1] + 62, card[0] + 290, card[1] + 108), radius=23, fill=WHITE)
    # righe di testo
    y = card[1] + 240
    for w in (360, 300, 360, 250):
        d.rounded_rectangle((card[0] + 70, y, card[0] + 70 + w, y + 42), radius=21, fill=DARK)
        y += 110
    # onde verso destra (trasmissione)
    cx, cy = 730, 520
    for r in (95, 170, 245):
        d.arc((cx - r, cy - r, cx + r, cy + r), start=-42, end=42, fill=ORANGE, width=34)
    return im


def main():
    os.makedirs(OUT, exist_ok=True)
    big = draw_icon()
    for size in (512, 192, 64, 32):
        big.resize((size, size), Image.LANCZOS).save(os.path.join(OUT, "remnote-%d.png" % size))
    big.resize((180, 180), Image.LANCZOS).save(os.path.join(OUT, "apple-touch-icon.png"))
    sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    big.save(os.path.join(OUT, "remnote.ico"), format="ICO", sizes=sizes)
    print("icone create in", os.path.abspath(OUT))


if __name__ == "__main__":
    main()
