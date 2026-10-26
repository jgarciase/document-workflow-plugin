# /// script
# requires-python = ">=3.10"
# dependencies = []
# ///
"""Detecta funciones de un .xlsx que openpyxl NO conserva al guardar (segmentaciones, tablas dinámicas,
formato condicional/validaciones avanzadas, macros, gráficos, imágenes...). Úsalo ANTES de editar.

Uso: xlsx_features.py <libro.xlsx>
Código de salida: 0 = seguro para openpyxl, 2 = riesgo (usa edición por COM o por XML).
"""
import re
import sys
import zipfile

RISK = [
    (r"^xl/slicers/|^xl/slicerCaches/", "segmentaciones (slicers)", "se pierden con openpyxl"),
    (r"^xl/pivotTables/|^xl/pivotCache/", "tablas dinámicas", "openpyxl las conserva a medias; las caché pueden quedar rotas"),
    (r"^xl/charts/", "gráficos", "openpyxl los descarta al leer"),
    (r"^xl/drawings/.*\.xml$", "dibujos/formas/imágenes", "openpyxl solo conserva imágenes simples"),
    (r"^xl/vbaProject\.bin$", "macros VBA", "guardar como .xlsx las elimina; usa .xlsm y COM"),
    (r"^xl/externalLinks/", "vínculos externos", "se mantienen pero sin caché fiable"),
    (r"^xl/timelines/|^xl/timelineCaches/", "escalas de tiempo", "se pierden"),
    (r"^xl/queryTables/|^xl/connections\.xml$", "conexiones de datos", "se pierden"),
]


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    z = zipfile.ZipFile(sys.argv[1])
    names = z.namelist()
    found = []
    for pat, label, note in RISK:
        hits = [n for n in names if re.search(pat, n)]
        if hits:
            found.append((label, len(hits), note))
    ext = 0
    for n in names:
        if re.match(r"^xl/worksheets/sheet\d+\.xml$", n):
            x = z.read(n).decode("utf8", "ignore")
            if "<extLst" in x:
                ext += 1
    if ext:
        found.append(("extensiones de hoja (formato condicional/validación avanzada, minigráficos)", ext,
                      "openpyxl las elimina"))
    wb = z.read("xl/workbook.xml").decode("utf8", "ignore")
    sheets = re.findall(r'<sheet [^>]*name="([^"]+)"', wb)
    print(f"Hojas: {sheets}")
    if not found:
        print("Sin elementos de riesgo: se puede editar con openpyxl.")
        return 0
    print("ELEMENTOS QUE OPENPYXL NO CONSERVA:")
    for label, n, note in found:
        print(f"- {label} ({n}): {note}")
    print("Recomendación: edita con Excel COM (recalc/refresh) o toca solo el XML de las celdas; "
          "si usas openpyxl, entrega SIEMPRE una copia y avisa de lo que se pierde.")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
