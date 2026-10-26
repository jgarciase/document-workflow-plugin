#!/usr/bin/env sh
# Lanzador portable (macOS). Uso: ./bootstrap/run.sh scripts/extract_brand.py <args>
set -e
if ! command -v uv >/dev/null 2>&1; then
  if [ ! -x "$HOME/.local/bin/uv" ]; then
    echo "[document-workflow] Instalando uv (una sola vez)..." >&2
    curl -LsSf https://astral.sh/uv/install.sh | sh >&2
  fi
  export PATH="$HOME/.local/bin:$PATH"
fi
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
script="$1"; shift
case "$script" in
  /*) target="$script" ;;       # ruta absoluta (p. ej. script temporal): tal cual
  *)  target="$ROOT/$script" ;; # relativa: respecto a la raíz del plugin
esac
exec uv run --quiet "$target" "$@"
