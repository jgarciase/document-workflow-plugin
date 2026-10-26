# Lanzador portable (Windows). Uso: ./bootstrap/run.ps1 scripts/extract_brand.py <args>
# Si uv no existe lo instala en el perfil del usuario (sin admin, sin Python previo).
$ErrorActionPreference = 'Stop'
$uv = Get-Command uv -ErrorAction SilentlyContinue
if (-not $uv) {
    $local = Join-Path $env:USERPROFILE '.local\bin\uv.exe'
    if (-not (Test-Path $local)) {
        Write-Host '[document-workflow] Instalando uv (una sola vez)...'
        Invoke-RestMethod https://astral.sh/uv/install.ps1 | Invoke-Expression
    }
    $env:Path = "$(Split-Path $local);$env:Path"
}
$root = Split-Path $PSScriptRoot -Parent
$script = $args[0]
$rest = @($args | Select-Object -Skip 1)
# Ruta absoluta (p. ej. un script temporal) se usa tal cual; relativa, respecto a la raíz del plugin.
$target = if ([IO.Path]::IsPathRooted($script)) { $script } else { Join-Path $root $script }
& uv run --quiet $target @rest
exit $LASTEXITCODE
