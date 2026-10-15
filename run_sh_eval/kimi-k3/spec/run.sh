#!/usr/bin/env bash
set +x
set -euo pipefail
umask 077
SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
if [[ -f "$SCRIPT_DIR/config.env" ]]; then
  set -a
  source "$SCRIPT_DIR/config.env"
  set +a
fi
exec "${PYTHON_BIN:-python3}" "$SCRIPT_DIR/launch.py" "$@"
