# /// script
# requires-python = ">=3.10"
# dependencies = ["python-docx"]
# ///
"""Genera un .docx nuevo partiendo de una plantilla: conserva cabecera, pie, márgenes y los
bloques que indiques, sustituye marcadores y añade contenido con los estilos de la plantilla.

Uso: build_docx_from_template.py <plantilla.docx> <salida.docx> <spec.json>

spec.json:
{
  "keep":    [0, 1, 2, 3, 8, 9],            // índices del cuerpo a conservar (ver dump_docx_full.py)
  "fill":    {"<Título>": "Mi título"},     // marcadores a sustituir (aunque estén partidos en runs)
  "content": [["EJIE-Titulo1", "Texto"],    // [estilo, texto] en orden
              ["Normal", "Párrafo"]]
}
Después, si la plantilla tiene índice: powershell -File scripts/refresh_docx.ps1 <salida.docx>
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from docx import Document  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from replace_text import W, replace_in_paragraph  # noqa: E402


def main() -> int:
    src, dst, spec_path = sys.argv[1:4]
    spec = json.loads(Path(spec_path).read_text(encoding="utf-8-sig"))
    d = Document(src)
    body = d.element.body
    keep = set(spec.get("keep", []))
    for i, el in enumerate(list(body.iterchildren())):
        if el.tag != qn("w:sectPr") and i not in keep:
            body.remove(el)

    sect = body.find(qn("w:sectPr"))
    names = {s.name for s in d.styles}
    for style, text in spec.get("content", []):
        if style not in names:
            raise SystemExit(f"Estilo inexistente en la plantilla: {style!r}. Disponibles: {sorted(names)}")
        p = d.add_paragraph(text, style=style)
        body.remove(p._p)
        sect.addprevious(p._p)

    fill = spec.get("fill", {})
    n = 0
    if fill:
        roots = [body] + [part._element for s in d.sections for part in (
            s.header, s.footer, s.first_page_header, s.first_page_footer)]
        n = sum(replace_in_paragraph(p, f"{W}r", f"{W}t", fill) for r in roots for p in r.iter(f"{W}p"))
    d.save(dst)
    print(f"OK -> {dst} ({n} marcadores sustituidos)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
