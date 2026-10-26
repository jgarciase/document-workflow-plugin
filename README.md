# document-workflow

Plugin de Cursor para crear y editar documentos Office y PDF adaptándolos a una plantilla corporativa (colores, fuentes, logos, cabeceras, layouts).

## Requisitos

- Windows o macOS con Microsoft Office instalado (para vista previa y recálculo; no hace falta LibreOffice).
- Nada más: no necesitas Python, Node ni `uv`. El primer uso instala `uv` en tu perfil (sin permisos de administrador) y descarga las dependencias de cada script automáticamente.

## Instalación

Instala el plugin desde el marketplace de Cursor de tu organización o copia esta carpeta como plugin local.

## Uso

Dile al agente, por ejemplo:

- «Crea un informe sobre X usando esta plantilla: `C:\ruta\plantilla.docx`»
- «Haz una presentación de 6 diapositivas con `plantilla.pptx`»
- «Rellena este Excel manteniendo el formato»

El flujo recomendado: analizar la plantilla (`list_styles.py`, `extract_brand.py`), generar (`build_docx_from_template.py`, `build_pptx_from_template.py`), actualizar índice (`refresh_docx.ps1`), renderizar (`render_preview.ps1/.sh`) y revisar las imágenes.

## Scripts

| Script | Función |
| --- | --- |
| `bootstrap/run.ps1`, `run.sh` | Lanzador: instala `uv` si falta y ejecuta el script |
| `list_styles.py` | Estilos (docx) y layouts/placeholders (pptx) |
| `extract_brand.py` | Perfil de marca: tema, fuentes, estilos, márgenes |
| `build_docx_from_template.py` | docx nuevo desde plantilla |
| `build_pptx_from_template.py` | pptx nuevo desde plantilla |
| `contact_sheet.py` | Miniaturas numeradas de un PDF (elegir diseños de una galería) |
| `xlsx_features.py` | Detecta slicers/pivots/gráficos que openpyxl destruiría |
| `edit_xlsx_com.ps1` | Edita xlsx con Excel conservando todo |
| `replace_text.py` | Buscar y reemplazar conservando formato |
| `refresh_docx.ps1` | Actualiza índice y campos con Word |
| `render_preview.ps1/.sh` | docx/pptx/xlsx a PDF con Office |
| `pdf_to_png.py`, `pdf_tools.py` | PDF a imágenes; unir, dividir, formularios |
| `recalc_xlsx.ps1/.sh`, `xlsx_check.py` | Recalcular y auditar fórmulas |
| `read_office.py`, `dump_docx_full.py`, `dump_table_xml.py` | Inspección |

## Archivos temporales

Los scripts de un solo uso, JSON intermedios y vistas previas van a `%TEMP%\document-workflow\` (Windows) o `$TMPDIR/document-workflow/` (macOS) y se borran al terminar. Tu proyecto solo recibe el entregable final.

## Problemas frecuentes

- **Word/PowerPoint ocupado o `RPC_E_CALL_REJECTED`**: los scripts reintentan solos. Si persiste, cierra Office y repite.
- **Primera ejecución lenta**: está instalando `uv` y dependencias; las siguientes son rápidas.
- **Sin acceso a internet** en la primera ejecución: `uv` y los paquetes se descargan una vez; pide a TI acceso a `astral.sh` y `pypi.org`.
