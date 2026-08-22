#!/bin/bash
set -euo pipefail

# k3 auto-subagent arm. Same 47 parallel-topology cases as the glm-5.3 eval.
# Default sub-agents (coder/explore/plan) via unset KIMI_SUBAGENTS.
#
# Isolation vs glm53 job 89935 (gnho019, PG 29700, relay 19321):
#   - dedicated login relay :19322 (same upstream 172.16.55.136:8317)
#   - PG_PORT_BASE=29800 + dedicated lease dir
#   - separate DUMP_ROOT / RUN_ID
#   - WooCommerce REST = PGPORT+10000, so 39800+ vs glm 39700+
#   - mock 30xxx locks stay shared per UID+node (serialize, do not double-bind)

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
DS_SCRIPT_DIR="${PROJECT_ROOT}/run_sh/kimi-code-deepseek-v4-flash"
GLM_SCRIPT_DIR="${PROJECT_ROOT}/run_sh/kimi-code-glm-5-3"
PYTHON_BIN="${PYTHON_BIN:-/storage/lintaoLab/bowending/miniconda3/envs/dbw_dev/bin/python3}"

export KIMI_CONFIG_ENV="${KIMI_CONFIG_ENV:-${SCRIPT_DIR}/config.env}"
# shellcheck disable=SC1091
source "$KIMI_CONFIG_ENV"

export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
export KIMI_MODEL_THINKING_EFFORT="${KIMI_MODEL_THINKING_EFFORT:-high}"
export RELAY_API_KEY="${RELAY_API_KEY:-$MODEL_API_KEY}"
export RELAY_PORT="${RELAY_PORT:-19322}"
export RELAY_SKIP_AUTOSTART=1
export RELAY_KIND="${RELAY_KIND:-tcp}"
export RELAY_UPSTREAM_HOST="${RELAY_UPSTREAM_HOST:-172.16.55.136}"
export RELAY_UPSTREAM_PORT="${RELAY_UPSTREAM_PORT:-8317}"
export PG_PORT_BASE="${PG_PORT_BASE:-29800}"
export MOCK_PORT_WAIT_LOOPS="${MOCK_PORT_WAIT_LOOPS:-360}"

unset KIMI_SUBAGENTS || true

export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_k3-subagent-parallel47}"
export RUN_ID="${RUN_ID:-k3-par47-$(date +%Y%m%d-%H%M%S)}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/toolathlon_pg_port_leases_k3_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho014}"
export SLURM_MEM="${SLURM_MEM:-256G}"
export SLURM_CPUS="${SLURM_CPUS:-12}"
export MAX_CONCURRENT="${MAX_CONCURRENT:-6}"
export SLURM_TIME="${SLURM_TIME:-20:00:00}"
export SLURM_JOB_NAME="${SLURM_JOB_NAME:-k3-par47}"

mkdir -p "$DUMP_ROOT"

if curl -sS -m 5 -o /dev/null \
     -H "Authorization: Bearer ${MODEL_API_KEY}" \
     "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
  echo "[k3] relay healthy on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND}"
else
  nohup "$PYTHON_BIN" "${DS_SCRIPT_DIR}/api_relay.py" \
    "$RELAY_PORT" "$RELAY_UPSTREAM_HOST" "$RELAY_UPSTREAM_PORT" \
    >/dev/shm/api_relay_k3_tcp.log 2>&1 &
  sleep 1
  if curl -sS -m 5 -o /dev/null \
       -H "Authorization: Bearer ${MODEL_API_KEY}" \
       "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
    echo "[k3] relay started on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND} pid=$!"
  else
    echo "[k3] FATAL: relay failed on :${RELAY_PORT}" >&2
    tail -n 40 /dev/shm/api_relay_k3_tcp.log 2>/dev/null || true
    exit 1
  fi
fi

echo "=== ping k3 through login-node relay ==="
curl -sS -m 30 -w "\nHTTP %{http_code} time=%{time_total}\n" \
  -H "Authorization: Bearer ${MODEL_API_KEY}" \
  -H "Content-Type: application/json" \
  -d "{\"model\":\"${MODEL_NAME}\",\"messages\":[{\"role\":\"user\",\"content\":\"ping\"}],\"max_tokens\":8,\"reasoning_effort\":\"${KIMI_MODEL_THINKING_EFFORT}\"}" \
  "http://127.0.0.1:${RELAY_PORT}/v1/chat/completions" | tail -n 8

mapfile -t TASKS < "${TASK_LIST:-${GLM_SCRIPT_DIR}/parallel_47.txt}"

echo "=== k3 auto-subagent parallel ($(date)) ==="
echo "Model:    $MODEL_NAME @ $MODEL_API_URL  effort=$KIMI_MODEL_THINKING_EFFORT"
echo "Relay:    login 0.0.0.0:${RELAY_PORT} -> ${RELAY_UPSTREAM_HOST}:${RELAY_UPSTREAM_PORT}"
echo "Dump:     $DUMP_ROOT  run_id=$RUN_ID"
echo "Tasks:    ${#TASKS[@]}  node=$SLURM_NODELIST  cpus=$SLURM_CPUS  concurrent=$MAX_CONCURRENT  pg_base=$PG_PORT_BASE"
echo "Subagents: default coder/explore/plan (KIMI_SUBAGENTS unset)"
echo ""

cd "$PROJECT_ROOT"
exec bash "${DS_SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
