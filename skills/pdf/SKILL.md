---
name: pdf
description: Leer, extraer texto, unir, dividir, rotar, rellenar formularios y crear archivos PDF con Python portátil (pypdf, pymupdf, reportlab), sin herramientas del sistema. Úsala ante cualquier mención de PDF, extraer texto de un PDF, juntar o separar páginas, formularios PDF o generar un PDF.
---

# PDF

## Reglas de portabilidad

- Nunca `python`, `python3` ni `pip`, ni `qpdf`, `pdftotext`, `pdftoppm` (no están en los equipos). Lanza con `bootstrap/run.ps1` (Windows) o `bootstrap/run.sh` (macOS).
- Rutas absolutas, y siempre una copia nueva de salida.
- Temporales (scripts de un solo uso, vistas previas, pruebas) solo en la carpeta temporal del sistema, nunca en el proyecto; ver `template-adapt`.

## Qué usar

`scripts/pdf_tools.py` cubre lo habitual:

| Orden | Resultado |
|---|---|
| `info f.pdf` | páginas, cifrado, metadatos, tamaño |
| `text f.pdf [--layout]` | texto por página (`--layout` por bloques, mejor para columnas) |
| `merge out.pdf a.pdf b.pdf` | unir |
| `split f.pdf carpeta` | una página por archivo |
| `extract f.pdf out.pdf 1-3,5` | seleccionar páginas |
| `rotate f.pdf out.pdf 90 [rango]` | rotar |
| `fields f.pdf` | listar campos de formulario |
| `fill f.pdf out.pdf valores.json` | rellenar campos |

Para ver páginas: `scripts/pdf_to_png.py f.pdf carpeta [dpi]` y revisa las imágenes.

## Crear un PDF

- **Desde un documento Office** (lo mejor para respetar la marca): crea el docx/pptx/xlsx con su skill y expórtalo con `scripts/render_preview.ps1`/`.sh <archivo> <salida.pdf>`, que usa Microsoft Office.
- **Desde cero** con `reportlab` (`dependencies = ["reportlab"]`): usa `platypus` (`SimpleDocTemplate`, `Paragraph`, `Table`, `Image`) con estilos definidos, no coordenadas manuales. Registra la fuente corporativa con `TTFont` si dispones del `.ttf`.

## Formularios

1. `fields` para ver nombres y tipos (`/Tx` texto, `/Btn` casilla o radio, `/Ch` lista).
2. `valores.json` con `{"nombre_campo": "valor"}`; casillas con el valor de estado que muestre `fields`.
3. `fill` y verifica abriendo el PDF renderizado. Si el PDF **no tiene campos**, hay que superponer texto con `pymupdf` (`page.insert_text((x, y), texto)`), calculando coordenadas sobre una imagen de la página.

## Trampas

- PDF escaneado sin texto: `text` devuelve vacío. El OCR necesita Tesseract, que no está instalado: avisa al usuario en lugar de inventar el contenido.
- PDF cifrado: `PdfReader(f).decrypt("clave")` solo con la clave que dé el usuario.
- Extraer tablas de un PDF es heurístico: `page.find_tables()` de `pymupdf` suele bastar; verifica el resultado contra la imagen de la página.
- No modifiques el PDF original; no inventes contenido que no se pueda leer.
