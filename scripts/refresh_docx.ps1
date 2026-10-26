# Actualiza índices (TOC) y campos de un .docx con Word y guarda. Sin LibreOffice.
# Uso: ./scripts/refresh_docx.ps1 <archivo.docx>
param([Parameter(Mandatory)][string]$File)
$ErrorActionPreference = 'Stop'
$File = (Resolve-Path $File).Path

function Retry([scriptblock]$Action) {
  for ($i = 1; $i -le 8; $i++) {
    try { return & $Action } catch {
      if ($i -eq 8 -or $_.Exception.HResult -ne -2147418111) { throw }
      Start-Sleep -Milliseconds 1500
    }
  }
}

$app = New-Object -ComObject Word.Application
$app.Visible = $false; $app.DisplayAlerts = 0
try {
  $doc = Retry { $app.Documents.Open($File) }
  Retry { $doc.TablesOfContents | ForEach-Object { $_.Update() } }
  Retry { [void]$doc.Fields.Update() }
  Retry { $doc.Save() }
  Retry { $doc.Close($false) }
} finally {
  # Word puede desconectarse al cerrar (RPC_E_DISCONNECTED): es inocuo, el archivo ya está guardado.
  try { $app.Quit() } catch {}
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
Write-Output "Actualizado: $File"
