#!/usr/bin/env bash
set +x
set -euo pipefail
SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
PRESET="$(basename -- "$(dirname -- "$SCRIPT_DIR")")/$(basename -- "$SCRIPT_DIR")"
exec bash "$SCRIPT_DIR/../../_common/launch.sh" "$PRESET" "$@"
