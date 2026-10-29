#!/usr/bin/env python3
"""Scrive le informazioni di versione usate dalla compilazione di DeckWhisperer.

Uso:  python tools/buildinfo.py                 (legge il file VERSION)
      python tools/buildinfo.py <visibile> <numero>   (usato dalla compilazione automatica)

Crea windows/_buildinfo.py (versione mostrata nel programma) e
windows/version_info.txt (proprietà del file mefistofele.exe).
"""
import os
import re
import sys

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
WIN = os.path.join(ROOT, "windows")

TEMPLATE = """VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(%(a)d, %(b)d, %(c)d, %(d)d),
    prodvers=(%(a)d, %(b)d, %(c)d, %(d)d),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable('040904B0', [
        StringStruct('CompanyName', 'www.andreabattaglioli.it'),
        StringStruct('FileDescription', 'MEFISTOFELE DeckWhisperer - Remote notes & control for PowerPoint'),
        StringStruct('FileVersion', '%(dotted)s'),
        StringStruct('InternalName', 'mefistofele'),
        StringStruct('LegalCopyright', 'Andrea Battaglioli - www.andreabattaglioli.it - Powered by www.mywebtvlive.com'),
        StringStruct('OriginalFilename', 'mefistofele.exe'),
        StringStruct('ProductName', 'MEFISTOFELE DeckWhisperer'),
        StringStruct('ProductVersion', '%(dotted)s')
      ])
    ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"""


def main():
    if len(sys.argv) >= 3:
        shown, num = sys.argv[1].lstrip("v"), sys.argv[2].lstrip("v")
    else:
        with open(os.path.join(ROOT, "VERSION"), encoding="utf-8") as f:
            shown = num = f.read().strip().lstrip("v")
    parts = [int(x) for x in re.findall(r"\d+", num)[:4]]
    parts += [0] * (4 - len(parts))
    with open(os.path.join(WIN, "_buildinfo.py"), "w", encoding="utf-8") as f:
        f.write('APP_VERSION = "%s"\n' % shown)
    with open(os.path.join(WIN, "version_info.txt"), "w", encoding="utf-8") as f:
        f.write(TEMPLATE % {"a": parts[0], "b": parts[1], "c": parts[2], "d": parts[3],
                            "dotted": ".".join(str(p) for p in parts)})
    print("versione mostrata: %s | numero file: %s" % (shown, ".".join(str(p) for p in parts)))


if __name__ == "__main__":
    main()
