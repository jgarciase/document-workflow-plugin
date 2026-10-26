# Edita un .xlsx con Microsoft Excel (COM) conservando TODO: segmentaciones, tablas dinámicas, gráficos,
# formato condicional, validaciones y macros. Úsalo cuando xlsx_features.py avise de riesgo con openpyxl.
# Uso: ./scripts/edit_xlsx_com.ps1 <entrada.xlsx> <salida.xlsx> <spec.json>
#
# spec.json:
# { "sheets": {
#     "Hoja1": { "cells": { "A1": "Texto", "B2": 5, "C3": "=SUM(B2:B10)" },
#                "append_rows": [["a", 1], ["b", 2]],      // tras la última fila usada, desde la columna "append_col"
#                "append_col": "B",                         // por defecto A
#                "copy_format_from_last": true }            // copia el formato de la última fila existente
#   },
#   "refresh_pivots": true }                                   // actualiza tablas dinámicas y conexiones
param([Parameter(Mandatory)][string]$In, [Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$SpecPath)
$ErrorActionPreference = 'Stop'
$In = (Resolve-Path $In).Path
$Out = [IO.Path]::GetFullPath($Out)
if ($In -eq $Out) { throw 'La salida debe ser una copia distinta de la entrada.' }
$spec = Get-Content $SpecPath -Raw -Encoding UTF8 | ConvertFrom-Json

function Retry([scriptblock]$Action) {
  for ($i = 1; $i -le 8; $i++) {
    try { return & $Action } catch {
      if ($i -eq 8 -or $_.Exception.HResult -ne -2147418111) { throw }
      Start-Sleep -Milliseconds 1500
    }
  }
}
function ColNum([string]$c) { $n = 0; foreach ($ch in $c.ToUpper().ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }; $n }

$app = New-Object -ComObject Excel.Application
$app.Visible = $false; $app.DisplayAlerts = $false
try {
  $wb = Retry { $app.Workbooks.Open($In) }
  foreach ($prop in $spec.sheets.PSObject.Properties) {
    $ws = $wb.Worksheets.Item($prop.Name)
    $s = $prop.Value
    if ($s.cells) {
      foreach ($c in $s.cells.PSObject.Properties) {
        $cell = $ws.Range($c.Name)
        if ($c.Value -is [string] -and $c.Value.StartsWith('=')) { $cell.Formula = $c.Value } else { $cell.Value2 = $c.Value }
      }
    }
    if ($s.append_rows) {
      $last = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
      $col0 = if ($s.append_col) { ColNum $s.append_col } else { 1 }
      $r = $last
      foreach ($row in $s.append_rows) {
        $r++
        if ($s.copy_format_from_last) {
          [void]$ws.Rows.Item($last).Copy($ws.Rows.Item($r))   # copia formato (y contenido)
          [void]$ws.Rows.Item($r).ClearContents()               # deja solo el formato
        }
        for ($k = 0; $k -lt @($row).Count; $k++) { $ws.Cells.Item($r, $col0 + $k).Value2 = @($row)[$k] }
      }
    }
  }
  if ($spec.refresh_pivots) { Retry { $wb.RefreshAll() }; Retry { $app.CalculateUntilAsyncQueriesDone() } }
  Retry { $app.CalculateFull() }
  New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
  Retry { $wb.SaveAs($Out) }   # conserva el formato del original por extensión
  try { $wb.Close($false) } catch {}
} finally {
  try { $app.Quit() } catch {}
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
Write-Output $Out
