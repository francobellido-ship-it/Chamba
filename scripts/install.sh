#!/bin/sh
set -eu
cd "$(dirname "$0")/.."
if ! command -v soffice >/dev/null 2>&1 && ! command -v libreoffice >/dev/null 2>&1; then
  printf '%s\n' 'Falta LibreOffice Calc. Instálalo con las fuentes Liberation, DejaVu y Carlito; el Dockerfile ya incluye estos paquetes.' >&2
  exit 1
fi
export UV_CACHE_DIR="${UV_CACHE_DIR:-/workspace/.cache/uv}"
uv sync --locked
