---
name: docx
description: Crear, leer y editar documentos Word (.docx/.dotx) con Python portátil (python-docx), sin Node ni LibreOffice. Úsala ante cualquier mención de Word, .docx, informe, memoria, carta o ficha, o para reemplazar texto, insertar tablas, imágenes, cabeceras y pies. Si hay una plantilla, combínala con la skill template-adapt.
---

# Word (.docx)

## Reglas de portabilidad

- Nunca `python`, `python3` ni `pip`. Todo se lanza con el lanzador del plugin (instala `uv` si falta):
  - Windows: `powershell -File bootstrap/run.ps1 <script> <args>`
  - macOS: `sh bootstrap/run.sh <script> <args>`
- Tus propios scripts llevan cabecera PEP 723 (`# /// script` + `dependencies = ["python-docx"]`).
- Rutas de usuario siempre absolutas. Trabaja sobre una copia: nunca sobrescribas el original.
- Temporales (scripts de un solo uso, vistas previas, pruebas) solo en la carpeta temporal del sistema, nunca en el proyecto; ver `template-adapt`.

## Qué usar

| Tarea | Herramienta |
|---|---|
| Leer (con estilos y tablas) | `scripts/read_office.py archivo.docx` |
| Reemplazar texto o rellenar `{{marcadores}}` conservando formato | `scripts/replace_text.py in.docx out.docx mapa.json` |
| Crear o ampliar | `python-docx` en tu propio script |
| Analizar plantilla | `scripts/extract_brand.py` y `scripts/dump_table_xml.py` |
| Verificar visualmente | `scripts/render_preview.ps1` o `.sh` y luego `scripts/pdf_to_png.py` |

## Crear con calidad

- **Parte de una plantilla** siempre que exista (`Document("plantilla.docx")`); si no, de `Document()` y define estilos, no formato manual.
- Aplica **estilos** (`style="Heading 1"`, `"List Bullet"`, estilos propios) en lugar de fuente, tamaño o color por párrafo. Así cambiar la marca es cambiar un estilo.
- Títulos con `Heading N` reales: permiten índice y navegación.
- Tablas: estilo de tabla de la plantilla, `autofit = False`, anchos explícitos (`cell.width`) que sumen el ancho útil, fila de cabecera con `w:tblHeader` y filas con `w:cantSplit`.
- Imágenes: `run.add_picture(ruta, width=Cm(x))`; calcula el ancho máximo con ancho de página menos márgenes.
- Saltos de página: `doc.add_page_break()`; sección apaisada: nueva sección con `orientation` y tamaños intercambiados.
- Número de página, índice y campos: inserta `w:fldSimple` o `w:fldChar` por XML. Word los actualiza al abrir (F9); avisa al usuario.
- Idioma: fija `w:lang` en el estilo Normal si el documento no es inglés, para el corrector.

## Trampas

- `python-docx` **no ve** celdas ni filas dentro de controles de contenido (`w:sdt`). Usa `read_office.py` o recorre `iter(qn("w:tc"))`.
- Un párrafo se compone de varios *runs*: buscar texto en `run.text` falla si está partido. Usa `replace_text.py`.
- Para borrar un elemento: `el.getparent().remove(el)`. Para insertar en un punto: `ancla.addprevious(nuevo)`.
- Si copias una fila con controles de contenido, cambia los `w:id` para que no se dupliquen.
- No inventes datos (fechas, códigos, URLs). Si faltan, escribe "Pendiente" y avísalo.

## Verificación obligatoria

1. Reabre el archivo con `read_office.py` (comprueba que no está corrupto y que el texto es el esperado).
2. `render_preview` y `pdf_to_png`, y revisa las imágenes: desbordes, tablas partidas, logos, paginación.
3. Corrige y repite hasta que quede limpio.
