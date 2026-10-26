# Renderiza docx/pptx/xlsx a PDF usando Microsoft Office instalado (COM). Sin LibreOffice.
# Uso: ./scripts/render_preview.ps1 <archivo> [salida.pdf]
# Robusto: reintenta si Office está ocupado (RPC_E_CALL_REJECTED) y cierra siempre su instancia.
param([Parameter(Mandatory)][string]$File, [string]$Out)
$ErrorActionPreference = 'Stop'
$File = (Resolve-Path $File).Path
# Por defecto la salida va a la carpeta temporal del sistema, nunca junto al archivo de origen.
if (-not $Out) {
  $dir = Join-Path ([IO.Path]::GetTempPath()) 'document-workflow'
  New-Item -ItemType Directory -Force $dir | Out-Null
  $Out = Join-Path $dir ([IO.Path]::GetFileNameWithoutExtension($File) + '.pdf')
}
$Out = [IO.Path]::GetFullPath($Out)

function Retry([scriptblock]$Action) {
  for ($i = 1; $i -le 8; $i++) {
    try { return & $Action } catch {
      if ($i -eq 8 -or $_.Exception.HResult -ne -2147418111) { throw }
      Start-Sleep -Milliseconds 1500
    }
  }
}

$app = $null
try {
  switch ([IO.Path]::GetExtension($File).ToLower()) {
    { $_ -in '.docx', '.dotx' } {
      $app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
      $doc = Retry { $app.Documents.Open($File, $false, $true) }
      Retry { $doc.ExportAsFixedFormat($Out, 17) }
      try { $doc.Close($false) } catch {}  # Word puede desconectarse al cerrar; el PDF ya existe
    }
    { $_ -in '.pptx', '.potx' } {
      $app = New-Object -ComObject PowerPoint.Application
      $p = Retry { $app.Presentations.Open($File, $true, $false, $false) }
      Retry { $p.SaveAs($Out, 32) }; $p.Close()
    }
    '.xlsx' {
      $app = New-Object -ComObject Excel.Application; $app.Visible = $false; $app.DisplayAlerts = $false
      $wb = Retry { $app.Workbooks.Open($File, 0, $true) }
      Retry { $wb.ExportAsFixedFormat(0, $Out) }; $wb.Close($false)
    }
    default { throw "Formato no soportado: $File" }
  }
} finally {
  if ($app) { try { $app.Quit() } catch {}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
}
if (-not (Test-Path $Out)) { throw "No se generó el PDF: $Out" }
Write-Output $Out
