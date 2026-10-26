# /// script
# requires-python = ">=3.10"
# dependencies = ["python-docx"]
# ///
"""Volcado completo de texto: párrafos, filas de tablas (incluye celdas dentro de controles
de contenido sdt, que python-docx omite) y sdt a nivel de cuerpo.

Uso: dump_docx_full.py <archivo.docx>
"""
import sys

from docx import Document
from docx.text.paragraph import Paragraph

sys.stdout.reconfigure(encoding="utf8")
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def text(el) -> str:
    return "".join(t.text or "" for t in el.iter(f"{W}t"))


d = Document(sys.argv[1])
ti = 0
for i, el in enumerate(d.element.body.iterchildren()):
    tag = el.tag.split("}")[1]
    if tag == "p":
        p = Paragraph(el, d)
        print(f"[{i}] P[{p.style.name}] {p.text}")
    elif tag == "tbl":
        print(f"[{i}] TABLE#{ti}")
        for r, tr in enumerate(el.iter(f"{W}tr")):
            cells = [text(tc) for tc in tr.iter(f"{W}tc")]
            sdts = [(a.get(W + "val")) for a in tr.iter(f"{W}alias")]
            print(f"    r{r} {cells} sdt={sdts}")
        ti += 1
    elif tag == "sdt":
        alias = el.find(f".//{W}alias")
        print(f"[{i}] SDT alias={alias.get(W+'val') if alias is not None else None} len={len(text(el))}")
