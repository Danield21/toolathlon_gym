#!/bin/bash
set -euo pipefail

# gpt-5.6-sol(xhigh) remaining 46 cases on gnho019, concurrent=1 (serial).
# Skips arxiv-conference-prep (already success in par47-rerun).
# Thinking via model suffix; no reasoning_effort body field.

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
DS_SCRIPT_DIR="${PROJECT_ROOT}/run_sh/kimi-code-deepseek-v4-flash"
PYTHON_BIN="${PYTHON_BIN:-/storage/lintaoLab/bowending/miniconda3/envs/dbw_dev/bin/python3}"

export KIMI_CONFIG_ENV="${KIMI_CONFIG_ENV:-${SCRIPT_DIR}/config.xhigh.env}"
# shellcheck disable=SC1091
source "$KIMI_CONFIG_ENV"

export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
unset KIMI_MODEL_THINKING_EFFORT || true
export RELAY_API_KEY="${RELAY_API_KEY:-$MODEL_API_KEY}"
export RELAY_PORT="${RELAY_PORT:-19324}"
export RELAY_SKIP_AUTOSTART=1
export RELAY_KIND="${RELAY_KIND:-tcp}"
export RELAY_UPSTREAM_HOST="${RELAY_UPSTREAM_HOST:-104.168.43.47}"
export RELAY_UPSTREAM_PORT="${RELAY_UPSTREAM_PORT:-8317}"
export PG_PORT_BASE="${PG_PORT_BASE:-34000}"
export MOCK_PORT_WAIT_LOOPS="${MOCK_PORT_WAIT_LOOPS:-360}"

unset KIMI_SUBAGENTS || true

export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_gpt-5.6-sol-xhigh-par47-serial}"
export RUN_ID="${RUN_ID:-gpt56-c1-$(date +%Y%m%d-%H%M%S)}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/toolathlon_pg_port_leases_gpt56c1_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho019}"
# One worker: 8 CPU / 64G is enough; glm/k3 already hold 16 each on gnho019.
export SLURM_MEM="${SLURM_MEM:-64G}"
export SLURM_CPUS="${SLURM_CPUS:-8}"
# Force serial even if config.xhigh.env still has MAX_CONCURRENT=4.
export MAX_CONCURRENT=1
# 46 * 2h timeout worst-case; xhigh is slow so give 4 days.
export SLURM_TIME="${SLURM_TIME:-96:00:00}"
export SLURM_JOB_NAME="${SLURM_JOB_NAME:-gpt56-c1}"

mkdir -p "$DUMP_ROOT"

# Ensure the login TCP relay exists, then wait until /v1/models actually
# answers. After a cancelled xhigh batch the upstream can accept TCP but
# return no HTTP for a few minutes; launching earlier dies in run_on_slurm.sh.
if ! ss -lptn 2>/dev/null | grep -q ":${RELAY_PORT} "; then
  nohup "$PYTHON_BIN" "${DS_SCRIPT_DIR}/api_relay.py" \
    "$RELAY_PORT" "$RELAY_UPSTREAM_HOST" "$RELAY_UPSTREAM_PORT" \
    >/dev/shm/api_relay_gpt56_tcp.log 2>&1 &
  sleep 1
fi
relay_ok=0
for i in $(seq 1 36); do
  if curl -sS -m 8 -o /dev/null \
       -H "Authorization: Bearer ${MODEL_API_KEY}" \
       "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
    relay_ok=1
    echo "[gpt56-c1] relay healthy on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND} (try $i)"
    break
  fi
  echo "[gpt56-c1] waiting for gpt gateway /v1/models ($i/36) ..."
  sleep 10
done
if (( relay_ok == 0 )); then
  echo "[gpt56-c1] FATAL: ${RELAY_UPSTREAM_HOST}:${RELAY_UPSTREAM_PORT} via :${RELAY_PORT} still not serving /v1/models" >&2
  exit 1
fi
# Keep the existing :19324 relay; do not let run_on_slurm.sh spawn/own one.
export RELAY_SKIP_AUTOSTART=1

mapfile -t TASKS < "${TASK_LIST:-${SCRIPT_DIR}/parallel_46_skip_arxiv_conference_prep.txt}"

echo "=== gpt-5.6-sol(xhigh) remaining-46 c=1 ($(date)) ==="
echo "Model:    $MODEL_NAME @ $MODEL_API_URL  thinking=suffix(xhigh)  effort-body=<omitted>"
echo "Upstream: ${RELAY_UPSTREAM_HOST}:${RELAY_UPSTREAM_PORT}"
echo "Dump:     $DUMP_ROOT  run_id=$RUN_ID"
echo "Skip:     arxiv-conference-prep (kept in kimi-code_gpt-5.6-sol-xhigh-par47-rerun)"
echo "Tasks:    ${#TASKS[@]}  node=$SLURM_NODELIST  concurrent=$MAX_CONCURRENT  cpus=$SLURM_CPUS mem=$SLURM_MEM pg_base=$PG_PORT_BASE"
echo "Isolation: PG ${PG_PORT_BASE}+ / WC $((PG_PORT_BASE + 10000))+ vs glm 29900+ / k3 31000+"
echo "Subagents: default coder/explore/plan (KIMI_SUBAGENTS unset)"
echo ""

cd "$PROJECT_ROOT"
exec bash "${DS_SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
