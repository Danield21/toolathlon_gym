#!/bin/bash
set -euo pipefail

# MiniMax-M3 auto-subagent rerun of 13 unique cases on gnho019, concurrent=1.
# 3 were hit by explore/plan tool-name allowlist miss; 12 by API/timeout
# (wc-inventory-alert-gcal and yf-financial-health-excel-notion appear in both
# lists and are run once). Uses the live kimi_harness copy.
# Isolated from k3-c1-rerun (job 90363, PG 31000, relay :19322) and
# gpt56-c1 (job 90121, PG 34000, relay :19324) on the same node.

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
DS_SCRIPT_DIR="${PROJECT_ROOT}/run_sh/kimi-code-deepseek-v4-flash"
PYTHON_BIN="${PYTHON_BIN:-/storage/lintaoLab/bowending/miniconda3/envs/dbw_dev/bin/python3}"

export KIMI_CONFIG_ENV="${KIMI_CONFIG_ENV:-${SCRIPT_DIR}/config.linslab8317.env}"
# shellcheck disable=SC1091
source "$KIMI_CONFIG_ENV"

export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
unset KIMI_MODEL_THINKING_EFFORT || true
export RELAY_API_KEY="${RELAY_API_KEY:-$MODEL_API_KEY}"
# Unconditional overrides so sourced defaults cannot keep old MiniMax :19317 / PG 32000.
export RELAY_PORT=19326
export RELAY_SKIP_AUTOSTART=1
export RELAY_KIND=tcp
export RELAY_UPSTREAM_HOST=172.16.55.136
export RELAY_UPSTREAM_PORT=8317
export MODEL_API_URL="http://192.168.180.240:${RELAY_PORT}"
export PG_PORT_BASE=36000
export MOCK_PORT_WAIT_LOOPS=360

unset KIMI_SUBAGENTS || true

export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_MiniMax-M3-par47-noeffort}"
export RUN_ID="${RUN_ID:-mm3-c1-rerun-$(date +%Y%m%d-%H%M%S)}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/toolathlon_pg_port_leases_mm3c1rerun_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho019}"
# Serial worker. gnho019 already holds gpt56-c1 + k3-c1-rerun (16 CPU / 128G);
# 112 idle cores remain. 8 CPU / 64G matches the other c=1 jobs on this node.
export SLURM_MEM="${SLURM_MEM:-64G}"
export SLURM_CPUS="${SLURM_CPUS:-8}"
export MAX_CONCURRENT=1
# 13 * 2h timeout worst-case + enqueue slack.
export SLURM_TIME="${SLURM_TIME:-36:00:00}"
export SLURM_JOB_NAME="${SLURM_JOB_NAME:-mm3-c1-rerun}"

mkdir -p "$DUMP_ROOT"

if curl -sS -m 8 -o /dev/null \
     -H "Authorization: Bearer ${MODEL_API_KEY}" \
     "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
  echo "[mm3-c1-rerun] relay healthy on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND}"
else
  nohup "$PYTHON_BIN" "${DS_SCRIPT_DIR}/api_relay.py" \
    "$RELAY_PORT" "$RELAY_UPSTREAM_HOST" "$RELAY_UPSTREAM_PORT" \
    >/dev/shm/api_relay_mm3_linslab_tcp.log 2>&1 &
  sleep 1
  if curl -sS -m 8 -o /dev/null \
       -H "Authorization: Bearer ${MODEL_API_KEY}" \
       "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
    echo "[mm3-c1-rerun] relay started on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND} pid=$!"
  else
    echo "[mm3-c1-rerun] FATAL: relay failed on :${RELAY_PORT}" >&2
    tail -n 40 /dev/shm/api_relay_mm3_linslab_tcp.log 2>/dev/null || true
    exit 1
  fi
fi

echo "=== ping MiniMax-M3 through login-node relay (no reasoning_effort field) ==="
curl -sS -m 30 -w "\nHTTP %{http_code} time=%{time_total}\n" \
  -H "Authorization: Bearer ${MODEL_API_KEY}" \
  -H "Content-Type: application/json" \
  -d "{\"model\":\"${MODEL_NAME}\",\"messages\":[{\"role\":\"user\",\"content\":\"ping\"}],\"max_tokens\":8}" \
  "http://127.0.0.1:${RELAY_PORT}/v1/chat/completions" | tail -n 8

mapfile -t TASKS < "${TASK_LIST:-${SCRIPT_DIR}/rerun_allowfix_13.txt}"

echo "=== MiniMax-M3 allowfix/API rerun c=1 ($(date)) ==="
echo "Model:    $MODEL_NAME @ $MODEL_API_URL  effort=<omitted, backend default>"
echo "Relay:    login 0.0.0.0:${RELAY_PORT} -> ${RELAY_UPSTREAM_HOST}:${RELAY_UPSTREAM_PORT}"
echo "Dump:     $DUMP_ROOT  run_id=$RUN_ID"
echo "Tasks:    ${#TASKS[@]}  node=$SLURM_NODELIST  concurrent=$MAX_CONCURRENT  cpus=$SLURM_CPUS mem=$SLURM_MEM pg_base=$PG_PORT_BASE"
echo "Isolation: PG ${PG_PORT_BASE}+ / WC $((PG_PORT_BASE + 10000))+ vs k3-c1-rerun PG 31000+ / WC 41000+ / relay 19322 and gpt56-c1 PG 34000+ / WC 44000+ / relay 19324"
echo "Subagents: default coder/explore/plan (KIMI_SUBAGENTS unset)"
echo "Harness:  live kimi_harness (readonly infix glob + corrected orchestration prompt)"
echo ""

cd "$PROJECT_ROOT"
exec bash "${DS_SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
