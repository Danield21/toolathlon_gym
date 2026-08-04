#!/usr/bin/env bash
# Stop local SGLang started by start_sglang.sh
set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck disable=SC1091
source "${SCRIPT_DIR}/config.env"
PID_FILE="${SCRIPT_DIR}/logs/sglang_${SGLANG_PORT}.pid"
if [[ -f "$PID_FILE" ]]; then
  pid=$(cat "$PID_FILE")
  if kill -0 "$pid" 2>/dev/null; then
    echo "[info] killing sglang pid=$pid"
    kill "$pid" 2>/dev/null || true
    sleep 2
    kill -9 "$pid" 2>/dev/null || true
  fi
  rm -f "$PID_FILE"
fi
# Also match launch_server on this port
pkill -f "sglang.launch_server.*--port ${SGLANG_PORT}" 2>/dev/null || true
echo "[info] stopped"
