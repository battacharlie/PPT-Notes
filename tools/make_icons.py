#!/usr/bin/env python3
"""Genera le icone di reMNote (PNG e ICO) in windows/assets/ a partire dal logo
windows/assets/remnote-logo.png (logo tondo, sfondo trasparente). Richiede Pillow."""
import os
from PIL import Image

ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "windows", "assets")


def main():
    logo = Image.open(os.path.join(ASSETS, "remnote-logo.png")).convert("RGBA")
    w, h = logo.size
    m = int(round(w * 0.018))          # toglie il margine vuoto attorno al cerchio
    base = logo.crop((m, m, w - m, h - m))
    for size in (512, 192, 64, 32):
        base.resize((size, size), Image.LANCZOS).save(os.path.join(ASSETS, "remnote-%d.png" % size))
    # iOS riempie la trasparenza di nero: icona su fondo bianco
    touch = Image.new("RGBA", (180, 180), (255, 255, 255, 255))
    small = base.resize((164, 164), Image.LANCZOS)
    touch.alpha_composite(small, (8, 8))
    touch.convert("RGB").save(os.path.join(ASSETS, "apple-touch-icon.png"))
    big = base.resize((256, 256), Image.LANCZOS)
    big.save(os.path.join(ASSETS, "remnote.ico"), format="ICO",
             sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print("icone create in", os.path.abspath(ASSETS))


if __name__ == "__main__":
    main()
