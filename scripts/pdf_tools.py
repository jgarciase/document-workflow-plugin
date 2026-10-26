# /// script
# requires-python = ">=3.10"
# dependencies = ["pypdf", "pymupdf"]
# ///
"""Operaciones habituales con PDF. Sin dependencias del sistema.

  pdf_tools.py info    <pdf>
  pdf_tools.py text    <pdf> [--layout]
  pdf_tools.py merge   <salida.pdf> <a.pdf> <b.pdf> ...
  pdf_tools.py split   <pdf> <carpeta>             (una página por archivo)
  pdf_tools.py extract <pdf> <salida.pdf> <rango>  (rango: 1-3,5)
  pdf_tools.py rotate  <pdf> <salida.pdf> <grados> [rango]
  pdf_tools.py fields  <pdf>                       (campos de formulario rellenables)
  pdf_tools.py fill    <pdf> <salida.pdf> <valores.json>
"""
import json
import sys
from pathlib import Path

from pypdf import PdfReader, PdfWriter


def pages_from_range(spec: str, total: int) -> list[int]:
    out: list[int] = []
    for part in spec.split(","):
        if "-" in part:
            a, b = part.split("-")
            out += range(int(a), int(b) + 1)
        else:
            out.append(int(part))
    return [p - 1 for p in out if 1 <= p <= total]


def main() -> int:
    sys.stdout.reconfigure(encoding="utf8")
    cmd, *a = sys.argv[1:]
    if cmd == "info":
        r = PdfReader(a[0])
        print(f"páginas: {len(r.pages)}  cifrado: {r.is_encrypted}  metadatos: {dict(r.metadata or {})}")
        print(f"tamaño p1: {r.pages[0].mediabox.width} x {r.pages[0].mediabox.height} pt")
    elif cmd == "text":
        import pymupdf
        with pymupdf.open(a[0]) as doc:
            for i, p in enumerate(doc, 1):
                print(f"\n## Página {i}")
                if "--layout" in a:  # por bloques, respeta mejor columnas y tablas
                    print("\n".join(b[4].rstrip() for b in p.get_text("blocks")))
                else:
                    print(p.get_text().rstrip())
    elif cmd == "merge":
        w = PdfWriter()
        for f in a[1:]:
            w.append(f)
        w.write(a[0])
    elif cmd == "split":
        r = PdfReader(a[0])
        out = Path(a[1])
        out.mkdir(parents=True, exist_ok=True)
        for i, p in enumerate(r.pages, 1):
            w = PdfWriter()
            w.add_page(p)
            w.write(out / f"{Path(a[0]).stem}-p{i}.pdf")
    elif cmd == "extract":
        r = PdfReader(a[0])
        w = PdfWriter()
        for i in pages_from_range(a[2], len(r.pages)):
            w.add_page(r.pages[i])
        w.write(a[1])
    elif cmd == "rotate":
        r = PdfReader(a[0])
        w = PdfWriter()
        sel = set(pages_from_range(a[3], len(r.pages))) if len(a) > 3 else set(range(len(r.pages)))
        for i, p in enumerate(r.pages):
            w.add_page(p.rotate(int(a[2])) if i in sel else p)
        w.write(a[1])
    elif cmd == "fields":
        for k, v in (PdfReader(a[0]).get_fields() or {}).items():
            print(f"{k}: tipo={v.get('/FT')} valor={v.get('/V')!r} opciones={v.get('/_States_', '')}")
    elif cmd == "fill":
        r = PdfReader(a[0])
        w = PdfWriter()
        w.append(r)
        vals = json.loads(Path(a[2]).read_text(encoding="utf8"))
        for page in w.pages:
            w.update_page_form_field_values(page, vals, auto_regenerate=False)
        w.set_need_appearances_writer(True)
        w.write(a[1])
    else:
        print(__doc__)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
