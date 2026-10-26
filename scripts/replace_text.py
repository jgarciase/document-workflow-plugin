# /// script
# requires-python = ">=3.10"
# dependencies = ["python-docx", "python-pptx"]
# ///
"""Busca y reemplaza texto en docx/pptx conservando el formato del primer fragmento afectado.

Funciona aunque el texto esté partido en varios runs, y alcanza tablas, cabeceras, pies,
controles de contenido (docx) y formas/tablas/notas (pptx).

Uso: replace_text.py <entrada> <salida> <mapa.json>
  mapa.json: {"{{cliente}}": "ACME", "2025": "2026"}
"""
import json
import sys
from pathlib import Path

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def replace_in_paragraph(p, run_tag: str, text_tag: str, mapping: dict) -> int:
    runs = [r for r in p.iter(run_tag)]
    spans = []  # (run, t_element)
    for r in runs:
        for t in r.iter(text_tag):
            spans.append(t)
    if not spans:
        return 0
    count = 0
    for old, new in mapping.items():
        while True:
            full = "".join(t.text or "" for t in spans)
            i = full.find(old)
            if i < 0:
                break
            j = i + len(old)
            pos = 0
            first = True
            for t in spans:
                s = t.text or ""
                a, b = pos, pos + len(s)
                pos = b
                if b <= i or a >= j:
                    continue
                lo, hi = max(i, a) - a, min(j, b) - a
                t.text = s[:lo] + (new if first else "") + s[hi:]
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                first = False
            count += 1
    return count


def do_docx(src, dst, mapping):
    from docx import Document

    d = Document(src)
    roots = [d.element.body]
    for s in d.sections:
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer,
                     s.even_page_header, s.even_page_footer):
            roots.append(part._element)
    n = sum(replace_in_paragraph(p, f"{W}r", f"{W}t", mapping) for root in roots for p in root.iter(f"{W}p"))
    d.save(dst)
    return n


def do_pptx(src, dst, mapping):
    from pptx import Presentation

    prs = Presentation(src)
    n = 0
    for s in prs.slides:
        trees = [s._element]
        if s.has_notes_slide:
            trees.append(s.notes_slide._element)
        for tree in trees:
            for p in tree.iter(f"{A}p"):
                n += replace_in_paragraph(p, f"{A}r", f"{A}t", mapping)
    prs.save(dst)
    return n


def main() -> int:
    src, dst, mp = sys.argv[1:4]
    mapping = json.loads(Path(mp).read_text(encoding="utf8"))
    ext = Path(src).suffix.lower()
    fn = {".docx": do_docx, ".pptx": do_pptx}.get(ext)
    if not fn:
        print("Solo .docx y .pptx", file=sys.stderr)
        return 2
    print(f"{fn(src, dst, mapping)} reemplazos -> {dst}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
