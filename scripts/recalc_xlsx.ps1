# Recalcula un .xlsx con Microsoft Excel (COM) y lo guarda en el sitio, con valores en caché.
# Uso: ./scripts/recalc_xlsx.ps1 <archivo.xlsx>
param([Parameter(Mandatory)][string]$File)
$ErrorActionPreference = 'Stop'
$File = (Resolve-Path $File).Path
$app = $null
try {
  $app = New-Object -ComObject Excel.Application
  $app.Visible = $false; $app.DisplayAlerts = $false
  for ($i = 1; $i -le 8; $i++) {
    try { $wb = $app.Workbooks.Open($File); break } catch {
      if ($i -eq 8 -or $_.Exception.HResult -ne -2147418111) { throw }
      Start-Sleep -Milliseconds 1500
    }
  }
  $app.CalculateFull()
  $wb.Save(); $wb.Close($false)
} finally {
  if ($app) { try { $app.Quit() } catch {}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
}
Write-Output "Recalculado: $File"
