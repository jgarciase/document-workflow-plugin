---
name: xlsx
description: Crear, leer, analizar y editar hojas de cálculo Excel (.xlsx/.xlsm) con Python portátil (openpyxl), con fórmulas reales y recálculo mediante Microsoft Excel instalado, sin LibreOffice. Úsala ante cualquier mención de Excel, hoja de cálculo, .xlsx, tabla dinámica de datos, informe con fórmulas o formato condicional.
---

# Excel (.xlsx)

## Reglas de portabilidad

- Nunca `python`, `python3` ni `pip`. Lanza con `bootstrap/run.ps1` (Windows) o `bootstrap/run.sh` (macOS).
- Scripts propios con cabecera PEP 723 y `dependencies = ["openpyxl"]` (añade `pandas` solo si hace falta).
- Trabaja sobre una copia y con rutas absolutas.
- Temporales (scripts de un solo uso, vistas previas, pruebas) solo en la carpeta temporal del sistema, nunca en el proyecto; ver `template-adapt`.

## Qué usar

| Tarea | Herramienta |
|---|---|
| Leer con coordenadas, fórmulas y valores | `scripts/read_office.py libro.xlsx` |
| Estilos y anchos de una plantilla | `scripts/extract_brand.py plantilla.xlsx` |
| **Recalcular** (valores en caché) | `scripts/recalc_xlsx.ps1` (Windows) / `recalc_xlsx.sh` (macOS) |
| Auditar errores y fórmulas | `scripts/xlsx_check.py libro.xlsx` |
| Crear / editar | `openpyxl` |
| Verificar visualmente | `scripts/render_preview.ps1`/`.sh` + `scripts/pdf_to_png.py` |

## Reglas de modelado

- **Fórmulas, no valores calculados en Python.** Si el usuario cambia un dato, el resultado debe actualizarse: `ws["C2"] = "=A2*B2"`.
- Supuestos y constantes en celdas propias (hoja "Parámetros"), referenciadas por fórmulas; nunca números mágicos dentro de la fórmula.
- Convención de color si no hay plantilla: azul para entradas, negro para fórmulas.
- Usa nombres de hoja sin espacios o entre comillas simples: `='Ventas 2026'!B2`.
- Formatos de número explícitos (`'#,##0.00'`, `'0.0%'`, `'dd/mm/yyyy'`); no guardes fechas como texto.
- Cabecera con estilo, `ws.freeze_panes = "A2"`, anchos de columna, filtro (`ws.auto_filter.ref`). Tablas con `openpyxl.worksheet.table.Table` para rangos con nombre.
- Formato condicional con `openpyxl.formatting.rule`; validación con `DataValidation`.
- Gráficos nativos con `openpyxl.chart`.

## Compatibilidad de funciones

Excel recalcula todo, pero `openpyxl` escribe las fórmulas tal cual. Las funciones nuevas (`XLOOKUP`, `FILTER`, `SORT`, `UNIQUE`, `LET`) exigen el prefijo `_xlfn.` al escribirlas (`=_xlfn.XLOOKUP(...)`) y las matrices dinámicas pueden mostrar `@`. Si el destinatario puede tener Excel antiguo, usa `INDEX/MATCH` y `SUMIFS`.

## Libros existentes con segmentaciones, tablas dinámicas o gráficos (importante)

`openpyxl` **elimina** al guardar: segmentaciones (slicers), parte de las tablas dinámicas, gráficos, formas, formato condicional y validaciones avanzadas. Antes de editar un libro que no hayas creado tú:

1. Ejecuta `scripts/xlsx_features.py <libro.xlsx>`. Código de salida 0 = seguro para `openpyxl`; 2 = riesgo.
2. Si hay riesgo, edita con Excel (COM) usando `scripts/edit_xlsx_com.ps1 <in> <out.xlsx> <spec.json>` (Windows). Permite fijar celdas/fórmulas, añadir filas copiando el formato de la última y refrescar tablas dinámicas, conservando todo lo demás. Trabaja siempre sobre una copia.
3. Tras editar, vuelve a pasar `xlsx_features.py` sobre la salida para comprobar que se conserva todo.
## Flujo

1. Edita con `openpyxl` (conserva estilos si abres la plantilla; `load_workbook(..., keep_vba=True)` para `.xlsm`).
2. **Recalcula** con `recalc_xlsx` (Excel calcula y guarda los valores).
3. **Audita** con `xlsx_check.py`: debe terminar en `OK` sin `#REF!`, `#NAME?`, `#DIV/0!`.
4. Renderiza y revisa el aspecto.

## Trampas

- `openpyxl` descarta valores en caché al guardar: sin recalcular, otros lectores verán celdas vacías.
- No soporta tablas dinámicas, segmentaciones ni macros nuevas: si la plantilla las tiene, no pases el archivo por `openpyxl`; avisa al usuario.
- Con `load_workbook(data_only=True)` y guardar, **pierdes las fórmulas**. Nunca guardes ese objeto.
- No inventes cifras: si faltan datos, deja la celda vacía o marcada "Pendiente".
