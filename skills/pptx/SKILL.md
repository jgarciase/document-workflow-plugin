---
name: pptx
description: Crear, leer y editar presentaciones PowerPoint (.pptx/.potx) con Python portátil (python-pptx), sin Node ni LibreOffice. Úsala ante cualquier mención de presentación, diapositivas, deck, slides o .pptx, o para rellenar una plantilla corporativa, reemplazar texto, añadir tablas, gráficos o imágenes. Con plantilla, combínala con template-adapt.
---

# PowerPoint (.pptx)

## Reglas de portabilidad

- Nunca `python`, `python3` ni `pip`. Lanza con `bootstrap/run.ps1` (Windows) o `bootstrap/run.sh` (macOS).
- Scripts propios con cabecera PEP 723 y `dependencies = ["python-pptx"]`.
- Trabaja sobre una copia y con rutas absolutas.
- Temporales (scripts de un solo uso, vistas previas, pruebas) solo en la carpeta temporal del sistema, nunca en el proyecto; ver `template-adapt`.

## Qué usar

| Tarea | Herramienta |
|---|---|
| Leer (layouts, placeholders `idx`, notas) | `scripts/read_office.py deck.pptx` |
| Layouts disponibles de una plantilla | `scripts/extract_brand.py plantilla.pptx` |
| Reemplazar texto o `{{marcadores}}` | `scripts/replace_text.py in.pptx out.pptx mapa.json` |
| Crear / modificar | `python-pptx` |
| Verificar | `scripts/render_preview.ps1`/`.sh` + `scripts/pdf_to_png.py` |

## Crear con calidad

- **Usa siempre los layouts de la plantilla** (`prs.slide_layouts[i]`) y rellena los placeholders por `idx` (`slide.placeholders[idx]`). No pongas cajas de texto sueltas si hay placeholder: se pierden fuente, tamaño y posición de marca.
- Una idea por diapositiva. Título con la conclusión, no solo el tema. Máximo ~6 líneas por cuerpo; si hay más, divide.
- Colores y fuentes **del tema**: `MSO_THEME_COLOR.ACCENT_1` en lugar de RGB fijos, para que respete la marca.
- Tablas: `shapes.add_table`, ajusta anchos de columna y tamaño de fuente (≥ 12 pt). Gráficos nativos con `shapes.add_chart` y `CategoryChartData`: son editables en PowerPoint.
- Imágenes: conserva la proporción (indica solo `width` o solo `height`).
- Notas del orador: `slide.notes_slide.notes_text_frame.text`.

## Eliminar o reordenar diapositivas

`python-pptx` no lo expone. Hazlo por XML sobre `prs.slides._sldIdLst`:

```python
ids = prs.slides._sldIdLst
sld = ids[i]
prs.part.drop_rel(sld.rId)   # libera la relación
ids.remove(sld)               # quita la diapositiva
```

Reordenar: `ids.remove(sld); ids.insert(nueva_pos, sld)`. Al terminar, borra las diapositivas de ejemplo de la plantilla.

## Trampas

- Los layouts pueden tener placeholders que no quieres: los vacíos se ven como "Haga clic para agregar texto" en edición. Elimínalos (`ph._element.getparent().remove(ph._element)`).
- Si copias una diapositiva, copia también las relaciones (imágenes). Más seguro: crea una nueva con el mismo layout y rellénala.
- No inventes cifras, nombres ni fechas.

## Verificación obligatoria

Renderiza a PDF, convierte a PNG y revisa cada diapositiva: desbordes de texto, solapes, contraste, logos, numeración. Itera hasta que quede limpio.
