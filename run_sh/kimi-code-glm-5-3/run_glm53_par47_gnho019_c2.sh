#!/bin/bash
set -euo pipefail

# GLM-5.3 parallel-47 on gnho019, concurrent=2, no reasoning_effort field.
# Default sub-agents (coder/explore/plan). Dedicated dump root (do not reuse
# kimi-code_glm-5-3-260801-parallel47).

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
DS_SCRIPT_DIR="${PROJECT_ROOT}/run_sh/kimi-code-deepseek-v4-flash"
PYTHON_BIN="${PYTHON_BIN:-/storage/lintaoLab/bowending/miniconda3/envs/dbw_dev/bin/python3}"

export KIMI_CONFIG_ENV="${KIMI_CONFIG_ENV:-${SCRIPT_DIR}/config.noeffort.env}"
# shellcheck disable=SC1091
source "$KIMI_CONFIG_ENV"

export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
# Explicitly drop any inherited/default effort so the runner does not inject
# KIMI_MODEL_THINKING_EFFORT into the container (run_eval_parallel.sh only
# passes -e when the var is non-empty).
unset KIMI_MODEL_THINKING_EFFORT || true
export RELAY_API_KEY="${RELAY_API_KEY:-$MODEL_API_KEY}"
export RELAY_PORT="${RELAY_PORT:-19321}"
export RELAY_SKIP_AUTOSTART=1
export RELAY_KIND="${RELAY_KIND:-tcp}"
export RELAY_UPSTREAM_HOST="${RELAY_UPSTREAM_HOST:-172.16.55.136}"
export RELAY_UPSTREAM_PORT="${RELAY_UPSTREAM_PORT:-8317}"
export HTTPS_RELAY_PROXY="${HTTPS_RELAY_PROXY:-http://127.0.0.1:7893}"

unset KIMI_SUBAGENTS || true

export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_glm-5-3-par47-noeffort}"
export RUN_ID="${RUN_ID:-glm53-c2-$(date +%Y%m%d-%H%M%S)}"
export PG_PORT_BASE="${PG_PORT_BASE:-29900}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/toolathlon_pg_port_leases_glm53c2_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho019}"
# 2 workers: previous 6-worker/64-CPU job sat at ~9 load. 16 CPU / 96G is
# enough for 2 enroot+pg+node processes without grabbing idle 64 on gnho019.
export SLURM_MEM="${SLURM_MEM:-96G}"
export SLURM_CPUS="${SLURM_CPUS:-16}"
export MAX_CONCURRENT="${MAX_CONCURRENT:-2}"
# 47 cases at c=2; some hit the 2h task timeout. 36h covers a long tail.
export SLURM_TIME="${SLURM_TIME:-36:00:00}"
export SLURM_JOB_NAME="${SLURM_JOB_NAME:-glm53-c2}"

mkdir -p "$DUMP_ROOT"

if curl -sS -m 5 -o /dev/null \
     -H "Authorization: Bearer ${MODEL_API_KEY}" \
     "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
  echo "[glm53-c2] relay healthy on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND}"
else
  if [[ "$RELAY_KIND" == "tcp" ]]; then
    nohup "$PYTHON_BIN" "${DS_SCRIPT_DIR}/api_relay.py" \
      "$RELAY_PORT" "$RELAY_UPSTREAM_HOST" "$RELAY_UPSTREAM_PORT" \
      >/dev/shm/api_relay_glm53_tcp.log 2>&1 &
  else
    nohup "$PYTHON_BIN" "${SCRIPT_DIR}/https_api_relay.py" "$RELAY_PORT" "$HTTPS_RELAY_PROXY" \
      >/dev/shm/api_relay_glm53.log 2>&1 &
  fi
  sleep 1
  if curl -sS -m 5 -o /dev/null \
       -H "Authorization: Bearer ${MODEL_API_KEY}" \
       "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
    echo "[glm53-c2] relay started on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND} pid=$!"
  else
    echo "[glm53-c2] FATAL: relay failed on :${RELAY_PORT}" >&2
    tail -n 40 /dev/shm/api_relay_glm53_tcp.log /dev/shm/api_relay_glm53.log 2>/dev/null || true
    exit 1
  fi
fi

echo "=== ping through login-node relay (no reasoning_effort field) ==="
curl -sS -m 30 -w "\nHTTP %{http_code} time=%{time_total}\n" \
  -H "Authorization: Bearer ${MODEL_API_KEY}" \
  -H "Content-Type: application/json" \
  -d "{\"model\":\"${MODEL_NAME}\",\"messages\":[{\"role\":\"user\",\"content\":\"ping\"}],\"max_tokens\":8}" \
  "http://127.0.0.1:${RELAY_PORT}/v1/chat/completions" | tail -n 8

mapfile -t TASKS < "${TASK_LIST:-${SCRIPT_DIR}/parallel_47.txt}"

echo "=== GLM-5.3 parallel-47 no-effort c=2 ($(date)) ==="
echo "Model:    $MODEL_NAME @ $MODEL_API_URL  effort=<omitted, backend default>"
if [[ "$RELAY_KIND" == "tcp" ]]; then
  echo "Relay:    login 0.0.0.0:${RELAY_PORT} -> ${RELAY_UPSTREAM_HOST}:${RELAY_UPSTREAM_PORT}"
else
  echo "Relay:    login 0.0.0.0:${RELAY_PORT} -> ark /api/plan/v3 via ${HTTPS_RELAY_PROXY}"
fi
echo "Dump:     $DUMP_ROOT  run_id=$RUN_ID"
echo "Tasks:    ${#TASKS[@]}  node=$SLURM_NODELIST  concurrent=$MAX_CONCURRENT  cpus=$SLURM_CPUS mem=$SLURM_MEM pg_base=$PG_PORT_BASE"
echo "Subagents: default coder/explore/plan (KIMI_SUBAGENTS unset)"
echo ""

cd "$PROJECT_ROOT"
exec bash "${DS_SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
