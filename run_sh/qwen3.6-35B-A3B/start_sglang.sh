#!/usr/bin/env bash
# Start SGLang OpenAI server for Qwen3.6-35B-A3B on gnho019 (or SGLANG_NODE).
#
# Usage (on GPU node, or via ssh/sbatch):
#   bash run_sh/qwen3.6-35B-A3B/start_sglang.sh
#
# Health:
#   curl http://127.0.0.1:${SGLANG_PORT}/v1/models

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck disable=SC1091
source "${SCRIPT_DIR}/config.env"

MODEL_PATH="${MODEL_PATH:?}"
SGLANG_PORT="${SGLANG_PORT:-30002}"
SGLANG_HOST="${SGLANG_HOST:-0.0.0.0}"
SGLANG_TP_SIZE="${SGLANG_TP_SIZE:-8}"
SGLANG_MEM_FRACTION="${SGLANG_MEM_FRACTION:-0.80}"
SGLANG_CONTEXT_LENGTH="${SGLANG_CONTEXT_LENGTH:-131072}"
SGLANG_CUDA_VISIBLE_DEVICES="${SGLANG_CUDA_VISIBLE_DEVICES:-0,1,2,3,4,5,6,7}"
SGLANG_CONDA_ENV="${SGLANG_CONDA_ENV:-/storage/lintaoLab/bowending/miniconda3/envs/bowen_verl2}"
SERVED_MODEL_NAME="${MODEL_NAME:-Qwen3.6-35B-A3B}"
LOG_DIR="${SCRIPT_DIR}/logs"
mkdir -p "$LOG_DIR"
LOG_FILE="${LOG_DIR}/sglang_${SGLANG_PORT}.log"
PID_FILE="${LOG_DIR}/sglang_${SGLANG_PORT}.pid"

if [[ ! -d "$MODEL_PATH" ]]; then
  echo "[error] MODEL_PATH missing: $MODEL_PATH" >&2
  exit 1
fi
if [[ ! -x "${SGLANG_CONDA_ENV}/bin/python" ]]; then
  echo "[error] conda env python missing: ${SGLANG_CONDA_ENV}/bin/python" >&2
  exit 1
fi

_is_port_listening() {
  local p="$1"
  ss -tln 2>/dev/null | grep -qE ":${p}[[:space:]]" && return 0
  return 1
}

if _is_port_listening "$SGLANG_PORT"; then
  if curl -sf --connect-timeout 3 "http://127.0.0.1:${SGLANG_PORT}/v1/models" >/dev/null; then
    echo "[info] SGLang already healthy on :${SGLANG_PORT} — reuse"
    curl -sS "http://127.0.0.1:${SGLANG_PORT}/v1/models" | head -c 400; echo
    exit 0
  fi
  echo "[error] port ${SGLANG_PORT} busy but /v1/models unhealthy" >&2
  exit 1
fi

export PATH="${SGLANG_CONDA_ENV}/bin:${PATH}"
export CUDA_VISIBLE_DEVICES="${SGLANG_CUDA_VISIBLE_DEVICES}"
export NVCC_PREPEND_FLAGS="${NVCC_PREPEND_FLAGS:-} -std=c++17"
# gnho019 driver 555 / CUDA 12.5 — torch is cu128; some sglang-kernel builds still need cudart/nvrtc 13.
export SGLANG_ENABLE_JIT_DEEPGEMM="${SGLANG_ENABLE_JIT_DEEPGEMM:-0}"
_SP="${SGLANG_CONDA_ENV}/lib/python3.12/site-packages"
# CUDA 12 only — do NOT put nvidia/cu13 on the path (driver 555 / CUDA 12.5).
_LIB_DIRS=(
  "${_SP}/nvidia/cuda_nvrtc/lib"
  "${_SP}/nvidia/cuda_runtime/lib"
  "${_SP}/torch/lib"
)
for d in "${_LIB_DIRS[@]}"; do
  if [[ -d "$d" ]]; then
    export LD_LIBRARY_PATH="${d}${LD_LIBRARY_PATH:+:$LD_LIBRARY_PATH}"
  fi
done
echo "[info] LD_LIBRARY_PATH head: ${LD_LIBRARY_PATH:0:220}..."

# Qwen3.6 requires sglang>=0.5.10 (official README).
# gnho019 has NO outbound network — never pip-install here. Install on login01 only.
VER="$("${SGLANG_CONDA_ENV}/bin/python" -c 'import sglang; print(sglang.__version__)' 2>/dev/null || echo 0)"
echo "[info] sglang version=${VER} (need >= ${SGLANG_MIN_VERSION:-0.5.10})"
NEED_UPGRADE="$("${SGLANG_CONDA_ENV}/bin/python" -c '
import sys
cur, need = sys.argv[1], sys.argv[2]
def parse(v):
    parts=[]
    for x in v.split("."):
        n=""
        for c in x:
            if c.isdigit(): n+=c
            else: break
        parts.append(int(n or 0))
    return tuple(parts+[0,0,0])[:3]
print("1" if parse(cur) < parse(need) else "0")
' "$VER" "${SGLANG_MIN_VERSION:-0.5.10}")"

if [[ "$NEED_UPGRADE" == "1" ]]; then
  echo "[error] sglang ${VER} < ${SGLANG_MIN_VERSION}." >&2
  echo "        Compute nodes are offline. On login01 install into:" >&2
  echo "          ${SGLANG_CONDA_ENV}" >&2
  echo "        e.g. pip install -U 'sglang[all]>=${SGLANG_MIN_VERSION}'" >&2
  exit 1
fi

echo "[info] MODEL_PATH=${MODEL_PATH}"
echo "[info] listen ${SGLANG_HOST}:${SGLANG_PORT} tp=${SGLANG_TP_SIZE} ctx=${SGLANG_CONTEXT_LENGTH}"
echo "[info] CUDA_VISIBLE_DEVICES=${CUDA_VISIBLE_DEVICES}"
echo "[info] log -> ${LOG_FILE}"

# shellcheck disable=SC2086
nohup "${SGLANG_CONDA_ENV}/bin/python" -m sglang.launch_server \
  --model-path "${MODEL_PATH}" \
  --served-model-name "${SERVED_MODEL_NAME}" \
  --host "${SGLANG_HOST}" \
  --port "${SGLANG_PORT}" \
  --tp-size "${SGLANG_TP_SIZE}" \
  --mem-fraction-static "${SGLANG_MEM_FRACTION}" \
  --context-length "${SGLANG_CONTEXT_LENGTH}" \
  --trust-remote-code \
  --reasoning-parser qwen3 \
  --tool-call-parser qwen3_coder \
  --allow-auto-truncate \
  ${SGLANG_EXTRA_ARGS:-} \
  >"${LOG_FILE}" 2>&1 &
echo $! >"${PID_FILE}"
echo "[info] started PID=$(cat "$PID_FILE") extra=${SGLANG_EXTRA_ARGS:-}"

# Wait until healthy
for i in $(seq 1 180); do
  if curl -sf --connect-timeout 2 "http://127.0.0.1:${SGLANG_PORT}/v1/models" >/dev/null; then
    echo "[info] SGLang READY after ${i}0s"
    curl -sS "http://127.0.0.1:${SGLANG_PORT}/v1/models"
    echo
    exit 0
  fi
  if ! kill -0 "$(cat "$PID_FILE")" 2>/dev/null; then
    echo "[error] SGLang process exited early — see ${LOG_FILE}" >&2
    tail -80 "$LOG_FILE" >&2 || true
    exit 1
  fi
  sleep 10
  echo "[info] waiting for /v1/models ... (${i}/180)"
done

echo "[error] timeout waiting for SGLang" >&2
tail -80 "$LOG_FILE" >&2 || true
exit 1
