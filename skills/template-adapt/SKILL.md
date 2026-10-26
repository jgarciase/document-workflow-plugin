---
name: template-adapt
description: Crea o completa documentos Word (.docx), PowerPoint (.pptx) o Excel (.xlsx) partiendo de una plantilla existente para respetar colores, fuentes, logos, tamaños, estilos y layouts. Úsala cuando el usuario aporte una plantilla o pida un documento con la identidad corporativa.
---

# Adaptar contenido a una plantilla

## Reglas de portabilidad (obligatorias)

- Nunca uses `python`, `python3` ni `pip`. Ejecuta todo con el lanzador del plugin, que instala `uv` si falta:
  - Windows: `powershell -File bootstrap/run.ps1 scripts/<script>.py <args>`
  - macOS: `sh bootstrap/run.sh scripts/<script>.py <args>`
- Para tus propios scripts de generación, usa cabecera PEP 723 (`# /// script` + `dependencies = [...]`) y lánzalos con el mismo lanzador.
- No uses Node ni LibreOffice. Para renderizar usa `scripts/render_preview.ps1` (Windows) o `scripts/render_preview.sh` (macOS), que usan Microsoft Office instalado.

Esta skill orquesta el flujo con plantilla. Para los detalles de cada formato usa las skills hermanas `docx`, `pptx`, `xlsx` y `pdf`.

## Scripts de apoyo (`scripts/`)

- `read_office.py`, `replace_text.py`, `pdf_tools.py`, `xlsx_check.py`, `recalc_xlsx.ps1`/`.sh`: ver cada skill de formato.
- `extract_brand.py`: perfil de marca (tema, fuentes, estilos, layouts).
- `dump_docx_full.py`: texto en orden, incluyendo celdas y filas dentro de controles de contenido (`sdt`), que `python-docx` no muestra.
- `build_docx_from_template.py <plantilla> <salida> <spec.json>`: atajo para docx. Conserva los bloques de `keep` (índices de `dump_docx_full.py`, p. ej. portada e índice), sustituye marcadores `fill` (aunque Word los parta en runs) y añade `content` como pares `[estilo, texto]`. Falla con la lista de estilos si uno no existe.
- `list_styles.py <plantilla>`: docx → estilos disponibles (fuente, tamaño, color); pptx → layouts y sus placeholders (idx, tipo). Ejecútalo antes de generar.
- `list_styles.py <pptx> --slides`: diapositivas de ejemplo con sus formas de texto numeradas.
- `contact_sheet.py <pdf> <carpeta> [cols] [filas]`: hojas de miniaturas numeradas. Para plantillas pptx que son una **galería de diseños**: renderiza con `render_preview`, genera la hoja de contacto, míra la imagen y elige qué diapositivas clonar.
- `build_pptx_from_template.py <plantilla> <salida> <spec.json>`: cada diapositiva es `from_slide` (clona una de ejemplo, con imágenes y decoración; rellena `texts` por nº de forma, `delete` formas sobrantes, `table`) o `layout` (diapositiva nueva con `placeholders` por idx). Admite `notes`. Elimina las diapositivas de ejemplo originales.
- `xlsx_features.py <xlsx>`: avisa de lo que `openpyxl` destruiría (slicers, pivots, gráficos...). `edit_xlsx_com.ps1 <in> <out> <spec.json>`: edita xlsx con Excel conservando todo.
- `refresh_docx.ps1 <docx>`: actualiza índice y campos con Word y guarda. Úsalo siempre tras añadir títulos si la plantilla tiene índice.
- `dump_table_xml.py`: XML de una tabla, para clonar su formato.
- `render_preview.ps1` / `.sh`: docx/pptx/xlsx a PDF con Office. `pdf_to_png.py`: PDF a imágenes para revisarlas.

## Archivos temporales (obligatorio)

El proyecto del usuario no debe llenarse de basura. Solo el entregable final va a la carpeta que pida el usuario.

- Scripts de generación de un solo uso, perfiles JSON, PDF de vista previa e imágenes: en una subcarpeta de la carpeta temporal del sistema (`$env:TEMP\document-workflow\` en Windows, `$TMPDIR/document-workflow/` en macOS).
- Escribe el script allí, no en el proyecto, y lánzalo con el lanzador pasando su ruta completa.
- `render_preview` ya escribe en esa carpeta por defecto. Para `pdf_to_png.py` pasa una subcarpeta temporal.
- Al terminar, borra lo que hayas creado en esa carpeta.
- Nunca dejes `.pyc`, `*.preview.pdf`, capturas ni copias de prueba junto a los documentos del usuario.

## Trampas conocidas

- Las plantillas corporativas usan `sdt` (controles de contenido) y secciones repetibles. Edita el texto de dentro, no los borres salvo que sean cuadros de ayuda.
- Nunca inventes datos que no estén en el documento (fechas, códigos, URLs, resultados). Déjalos como "Pendiente" y avisa.
- Si añades tablas, usa el estilo de tabla de la plantilla y alinea a la izquierda, con anchos que sumen el `tblGrid` original.

- Los marcadores tipo `<Título>` suelen estar partidos en varios runs: sustitúyelos con `fill` / `replace_text.py`, no con `run.text`.
- Los scripts en la carpeta temporal se escriben en UTF-8 **sin BOM** (en PowerShell 5, `Set-Content -Encoding UTF8` añade BOM y rompe el bloque `# /// script`). Usa la herramienta de escritura de archivos o `[IO.File]::WriteAllText` con `UTF8Encoding($false)`.
- Word puede lanzar `RPC_E_DISCONNECTED` al cerrar; los scripts del plugin lo toleran. Si ves Word colgado, cierra `WINWORD` y reintenta.
- La numeración de los estilos de título la define la plantilla; revísala en la imagen y avisa si el resultado no es el esperado.

## Flujo

1. **Analizar la plantilla**: `extract_brand.py <plantilla> -o brand-profile.json`. Lee el perfil: tema, fuentes, estilos, layouts, márgenes y logos.
2. **Generar partiendo del archivo de la plantilla**, nunca desde un documento vacío:
   - **docx** (caso simple): `build_docx_from_template.py` + `refresh_docx.ps1`. Caso complejo: `Document("plantilla.docx")`, vacía el cuerpo de ejemplo conservando `sectPr` (cabecera, pie, márgenes) y añade contenido con `style="<nombre del perfil>"`. No fijes fuente, color ni tamaño a mano si existe un estilo equivalente.
   - **pptx**: `Presentation("plantilla.pptx")`, usa `slide_layouts[i]` del perfil y rellena los placeholders por `idx`. Elimina las diapositivas de ejemplo al final. No dibujes textos en cajas sueltas si hay placeholder.
   - **xlsx**: abre la plantilla con `openpyxl`, copia estilos de las celdas modelo (`copy(cell._style)`) y conserva anchos, formatos y paneles inmovilizados.
3. **Verificar**: renderiza a PDF, conviértelo a imágenes y compáralo con la plantilla. Revisa desbordes de texto, logos, numeración de página y contraste.
4. **Entregar** el archivo y resume qué estilos y layouts usaste.

## Calidad

- Los colores salen del tema (`theme.colors`); no inventes tonos.
- Si falta un estilo para lo que pide el usuario (por ejemplo, una tabla), elige el más cercano del perfil y avisa.
- Conserva el idioma del usuario y no mezcles fuentes fuera de las del perfil.
