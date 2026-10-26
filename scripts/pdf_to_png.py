# /// script
# requires-python = ">=3.10"
# dependencies = ["pymupdf"]
# ///
"""Convierte un PDF en PNG (una imagen por página). Uso: pdf_to_png.py <pdf> <carpeta_salida> [dpi]"""
import sys
from pathlib import Path

import pymupdf

pdf, out = Path(sys.argv[1]), Path(sys.argv[2])
dpi = int(sys.argv[3]) if len(sys.argv) > 3 else 80
out.mkdir(parents=True, exist_ok=True)
with pymupdf.open(pdf) as doc:
    for i, page in enumerate(doc, 1):
        f = out / f"{pdf.stem}-p{i}.png"
        page.get_pixmap(dpi=dpi).save(f)
        print(f)
