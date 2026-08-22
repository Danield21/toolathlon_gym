#!/bin/bash
set -euo pipefail

# MiniMax-M3 parallel-47 on gnho014, concurrent=2, no reasoning_effort field.
# Isolated from deepseek job 89977 (PG 37432+, local :30000, 20-way mocks).

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
DS_SCRIPT_DIR="${PROJECT_ROOT}/run_sh/kimi-code-deepseek-v4-flash"
GLM_SCRIPT_DIR="${PROJECT_ROOT}/run_sh/kimi-code-glm-5-3"
PYTHON_BIN="${PYTHON_BIN:-/storage/lintaoLab/bowending/miniconda3/envs/dbw_dev/bin/python3}"

export KIMI_CONFIG_ENV="${KIMI_CONFIG_ENV:-${SCRIPT_DIR}/config.noeffort.env}"
# shellcheck disable=SC1091
source "$KIMI_CONFIG_ENV"

export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
unset KIMI_MODEL_THINKING_EFFORT || true
export RELAY_API_KEY="${RELAY_API_KEY:-$MODEL_API_KEY}"
export RELAY_PORT="${RELAY_PORT:-19317}"
export RELAY_SKIP_AUTOSTART=1
export RELAY_KIND="${RELAY_KIND:-tcp}"
export RELAY_UPSTREAM_HOST="${RELAY_UPSTREAM_HOST:-104.168.43.47}"
export RELAY_UPSTREAM_PORT="${RELAY_UPSTREAM_PORT:-8317}"
export PG_PORT_BASE="${PG_PORT_BASE:-32000}"
export MOCK_PORT_WAIT_LOOPS="${MOCK_PORT_WAIT_LOOPS:-720}"

unset KIMI_SUBAGENTS || true

export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_MiniMax-M3-par47-noeffort}"
export RUN_ID="${RUN_ID:-mm3-c2-$(date +%Y%m%d-%H%M%S)}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/toolathlon_pg_port_leases_mm3c2_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho014}"
# gnho014: 89977 uses 112 CPU, 89707 uses 2 → 14 idle. Request 12 so we start now.
export SLURM_MEM="${SLURM_MEM:-96G}"
export SLURM_CPUS="${SLURM_CPUS:-12}"
export MAX_CONCURRENT="${MAX_CONCURRENT:-2}"
export SLURM_TIME="${SLURM_TIME:-36:00:00}"
export SLURM_JOB_NAME="${SLURM_JOB_NAME:-mm3-c2}"

mkdir -p "$DUMP_ROOT"

if curl -sS -m 5 -o /dev/null \
     -H "Authorization: Bearer ${MODEL_API_KEY}" \
     "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
  echo "[mm3-c2] relay healthy on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND}"
else
  nohup "$PYTHON_BIN" "${DS_SCRIPT_DIR}/api_relay.py" \
    "$RELAY_PORT" "$RELAY_UPSTREAM_HOST" "$RELAY_UPSTREAM_PORT" \
    >/dev/shm/api_relay_mm3_tcp.log 2>&1 &
  sleep 1
  if curl -sS -m 5 -o /dev/null \
       -H "Authorization: Bearer ${MODEL_API_KEY}" \
       "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
    echo "[mm3-c2] relay started on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND} pid=$!"
  else
    echo "[mm3-c2] FATAL: relay failed on :${RELAY_PORT}" >&2
    tail -n 40 /dev/shm/api_relay_mm3_tcp.log 2>/dev/null || true
    exit 1
  fi
fi

echo "=== ping MiniMax-M3 through login-node relay (no reasoning_effort field) ==="
curl -sS -m 30 -w "\nHTTP %{http_code} time=%{time_total}\n" \
  -H "Authorization: Bearer ${MODEL_API_KEY}" \
  -H "Content-Type: application/json" \
  -d "{\"model\":\"${MODEL_NAME}\",\"messages\":[{\"role\":\"user\",\"content\":\"ping\"}],\"max_tokens\":8}" \
  "http://127.0.0.1:${RELAY_PORT}/v1/chat/completions" | tail -n 8

mapfile -t TASKS < "${TASK_LIST:-${GLM_SCRIPT_DIR}/parallel_47.txt}"

echo "=== MiniMax-M3 parallel-47 no-effort c=2 ($(date)) ==="
echo "Model:    $MODEL_NAME @ $MODEL_API_URL  effort=<omitted, backend default adaptive>"
echo "Upstream: ${RELAY_UPSTREAM_HOST}:${RELAY_UPSTREAM_PORT}"
echo "Dump:     $DUMP_ROOT  run_id=$RUN_ID"
echo "Tasks:    ${#TASKS[@]}  node=$SLURM_NODELIST  concurrent=$MAX_CONCURRENT  cpus=$SLURM_CPUS mem=$SLURM_MEM pg_base=$PG_PORT_BASE"
echo "Isolation: PG ${PG_PORT_BASE}+ / WC $((PG_PORT_BASE + 10000))+ vs deepseek 89977 PG 37432+ / WC 47432+ / local :30000"
echo "Subagents: default coder/explore/plan (KIMI_SUBAGENTS unset)"
echo ""

cd "$PROJECT_ROOT"
exec bash "${DS_SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
