# /// script
# requires-python = ">=3.10"
# dependencies = ["pymupdf"]
# ///
"""Crea hojas de contacto (miniaturas numeradas) de un PDF para elegir layouts de un vistazo.

Uso: contact_sheet.py <entrada.pdf> <carpeta_salida> [columnas=4] [filas=4]
Genera sheet-1.png, sheet-2.png... Cada miniatura lleva el número de página (= nº de diapositiva).
"""
import sys
from pathlib import Path

import pymupdf as fitz


def main() -> int:
    pdf, out = Path(sys.argv[1]), Path(sys.argv[2])
    cols = int(sys.argv[3]) if len(sys.argv) > 3 else 4
    rows = int(sys.argv[4]) if len(sys.argv) > 4 else 4
    out.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(pdf)
    first = doc[0].rect
    tw = 560
    th = int(tw * first.height / first.width)
    pad = 12
    per = cols * rows
    for s, start in enumerate(range(0, len(doc), per), 1):
        sheet = fitz.open()
        page = sheet.new_page(width=cols * (tw + pad) + pad, height=rows * (th + pad + 18) + pad)
        for k, n in enumerate(range(start, min(start + per, len(doc)))):
            r, c = divmod(k, cols)
            x = pad + c * (tw + pad)
            y = pad + r * (th + pad + 18)
            page.insert_text((x, y + 12), f"#{n + 1}", fontsize=12)
            rect = fitz.Rect(x, y + 16, x + tw, y + 16 + th)
            page.show_pdf_page(rect, doc, n)
            page.draw_rect(rect, color=(0.6, 0.6, 0.6), width=0.5)
        pix = page.get_pixmap(dpi=72)
        pix.save(out / f"sheet-{s}.png")
    print(f"{(len(doc) + per - 1) // per} hojas en {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
