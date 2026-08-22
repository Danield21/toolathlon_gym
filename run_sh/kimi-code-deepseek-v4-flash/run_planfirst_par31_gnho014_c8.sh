#!/bin/bash
set -euo pipefail

# Plan-first remaining 31 of parallel-47 on gnho019, concurrent=8.
# Official DeepSeek (api.deepseek.com) via login HTTPS relay :19325.
# Dumps stay in dumps/kimi-code_deepseek-v4-flash-plan-first (new run_id).
#
# Isolation vs current gnho019 tenants:
#   glm53-c2 89986  PG 29900+ / WC 39900+ / relay :19321
#   k3-c2    89993  PG 31000+ / WC 41000+ / relay :19322
#   gpt56-c1 90121  PG 34000+ / WC 44000+ / relay :19324
#   this job        PG 35000+ / WC 45000+ / relay :19325

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
PYTHON_BIN="${PYTHON_BIN:-/storage/lintaoLab/bowending/miniconda3/envs/dbw_dev/bin/python3}"

export KIMI_CONFIG_ENV="${KIMI_CONFIG_ENV:-${SCRIPT_DIR}/config.official.env}"
# shellcheck disable=SC1091
source "$KIMI_CONFIG_ENV"

export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
export KIMI_PLAN_FIRST="${KIMI_PLAN_FIRST:-1}"
export KIMI_MODEL_THINKING_EFFORT="${KIMI_MODEL_THINKING_EFFORT:-high}"
unset KIMI_SUBAGENTS || true

export RELAY_API_KEY="${RELAY_API_KEY:-$MODEL_API_KEY}"
export RELAY_PORT="${RELAY_PORT:-19325}"
export RELAY_SKIP_AUTOSTART=1
export RELAY_KIND="${RELAY_KIND:-https}"
export HTTPS_RELAY_PROXY="${HTTPS_RELAY_PROXY:-http://127.0.0.1:7893}"
export PG_PORT_BASE="${PG_PORT_BASE:-35000}"
export MOCK_PORT_WAIT_LOOPS="${MOCK_PORT_WAIT_LOOPS:-720}"

export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_deepseek-v4-flash-plan-first}"
export RUN_ID="${RUN_ID:-planfirst-c8-$(date +%Y%m%d-%H%M%S)}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/toolathlon_pg_port_leases_planfirst_c8_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho019}"
# c=8 ≈ 6 CPU/worker. gnho019 has 88 idle cores beside glm/k3/gpt56-c1.
export SLURM_MEM="${SLURM_MEM:-192G}"
export SLURM_CPUS="${SLURM_CPUS:-48}"
# Force 8 even if an inherited env or a later source tries to lower it.
export MAX_CONCURRENT=8
export SLURM_TIME="${SLURM_TIME:-36:00:00}"
export SLURM_JOB_NAME="${SLURM_JOB_NAME:-ds-pf-c8}"

mkdir -p "$DUMP_ROOT"

if curl -sS -m 8 -o /dev/null \
     -H "Authorization: Bearer ${MODEL_API_KEY}" \
     "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
  echo "[ds-pf-c8] relay healthy on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND}"
else
  nohup "$PYTHON_BIN" "${SCRIPT_DIR}/https_deepseek_relay.py" \
    "$RELAY_PORT" "$HTTPS_RELAY_PROXY" \
    >/dev/shm/api_relay_ds_official.log 2>&1 &
  sleep 1
  if curl -sS -m 8 -o /dev/null \
       -H "Authorization: Bearer ${MODEL_API_KEY}" \
       "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
    echo "[ds-pf-c8] relay started on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND} pid=$!"
  else
    echo "[ds-pf-c8] FATAL: HTTPS relay failed on :${RELAY_PORT}" >&2
    tail -n 40 /dev/shm/api_relay_ds_official.log 2>/dev/null || true
    exit 1
  fi
fi

echo "=== ping official deepseek-v4-flash through login HTTPS relay ==="
curl -sS -m 30 -w "\nHTTP %{http_code} time=%{time_total}\n" \
  -H "Authorization: Bearer ${MODEL_API_KEY}" \
  -H "Content-Type: application/json" \
  -d "{\"model\":\"${MODEL_NAME}\",\"messages\":[{\"role\":\"user\",\"content\":\"ping\"}],\"max_tokens\":8,\"reasoning_effort\":\"high\"}" \
  "http://127.0.0.1:${RELAY_PORT}/v1/chat/completions" | tail -n 8

mapfile -t TASKS < "${TASK_LIST:-${SCRIPT_DIR}/plan_first_par31_remaining.txt}"
if (( ${#TASKS[@]} != 31 )); then
  echo "[ds-pf-c8] FATAL: expected 31 tasks, got ${#TASKS[@]}" >&2
  exit 1
fi

echo "=== Plan-first remaining-31 c=8 ($(date)) ==="
echo "Model:    $MODEL_NAME @ $MODEL_API_URL  KIMI_PLAN_FIRST=$KIMI_PLAN_FIRST  effort=${KIMI_MODEL_THINKING_EFFORT}"
echo "Relay:    login 0.0.0.0:${RELAY_PORT} -> https://api.deepseek.com via ${HTTPS_RELAY_PROXY}"
echo "Dump:     $DUMP_ROOT  run_id=$RUN_ID"
echo "Tasks:    ${#TASKS[@]}  node=$SLURM_NODELIST  concurrent=$MAX_CONCURRENT  cpus=$SLURM_CPUS mem=$SLURM_MEM pg_base=$PG_PORT_BASE"
echo "Isolation: PG ${PG_PORT_BASE}+ / WC $((PG_PORT_BASE + 10000))+ vs glm 29900+ / k3 31000+ / gpt56-c1 34000+"
echo "Subagents: default coder/explore/plan (KIMI_SUBAGENTS unset)"
echo ""

cd "$PROJECT_ROOT"
exec bash "${SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
