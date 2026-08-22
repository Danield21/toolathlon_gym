#!/bin/bash
set -euo pipefail

# k3 auto-subagent eval of the 87 need_subagent cases not yet run in
# dumps/kimi-code_k3-par47-noeffort (the 47 parallel-topology cases).
# Same stack as par47-noeffort: config.noeffort.env, relay :19322, no effort,
# default coder/explore/plan. Concurrent=1 on gnho019.
# New dump: dumps/kimi-code_k3-needsub-noeffort
#
# Isolation vs current tenants:
#   mm3-r134 90666      gnho014  120 CPU / 512G / PG 41000+ / relay :19330
#   interactive 89707   gnho014  2 CPU / 12G
#   leftover k3 par47   gnho019  PG 31000+ / 35100+ (jobs finished)
#   this job            gnho019  PG 33000+ / WC 43000+ / relay :19322 / 8 CPU

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
DS_SCRIPT_DIR="${PROJECT_ROOT}/run_sh/kimi-code-deepseek-v4-flash"
PYTHON_BIN="${PYTHON_BIN:-/storage/lintaoLab/bowending/miniconda3/envs/dbw_dev/bin/python3}"

export KIMI_CONFIG_ENV="${KIMI_CONFIG_ENV:-${SCRIPT_DIR}/config.noeffort.env}"
# shellcheck disable=SC1091
source "$KIMI_CONFIG_ENV"

export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
unset KIMI_MODEL_THINKING_EFFORT || true
unset KIMI_SUBAGENTS || true
unset KIMI_PLAN_FIRST || true

export RELAY_API_KEY="${RELAY_API_KEY:-$MODEL_API_KEY}"
export RELAY_PORT=19322
export RELAY_SKIP_AUTOSTART=1
export RELAY_KIND=tcp
export RELAY_UPSTREAM_HOST="${RELAY_UPSTREAM_HOST:-172.16.55.136}"
export RELAY_UPSTREAM_PORT="${RELAY_UPSTREAM_PORT:-8317}"
export MODEL_API_URL="http://192.168.180.240:${RELAY_PORT}"
export PG_PORT_BASE=33000
export MOCK_PORT_WAIT_LOOPS=360
export MAX_CONCURRENT=1
export MAX_CONCURRENT_CAP=1

export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_k3-needsub-noeffort}"
export RUN_ID="${RUN_ID:-k3ns-$(date +%Y%m%d-%H%M%S)}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/tpl_k3ns_${UID}}"
export PG_RUNTIME_ROOT="${PG_RUNTIME_ROOT:-/dev/shm/tpg_k3ns_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho019}"
# Serial worker. gnho019 is idle; 8 CPU / 64G matches prior k3 c=1.
export SLURM_MEM="${SLURM_MEM:-64G}"
export SLURM_CPUS="${SLURM_CPUS:-8}"
# 87 * 2h timeout worst-case; linlab MaxTime is infinite.
export SLURM_TIME="${SLURM_TIME:-7-00:00:00}"
export SLURM_JOB_NAME="${SLURM_JOB_NAME:-k3-ns87}"

mkdir -p "$DUMP_ROOT"

if curl -sS -m 8 -o /dev/null \
     -H "Authorization: Bearer ${MODEL_API_KEY}" \
     "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
  echo "[k3-ns87] relay healthy on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND}"
else
  nohup "$PYTHON_BIN" "${DS_SCRIPT_DIR}/api_relay.py" \
    "$RELAY_PORT" "$RELAY_UPSTREAM_HOST" "$RELAY_UPSTREAM_PORT" \
    >/dev/shm/api_relay_k3_tcp.log 2>&1 &
  sleep 1
  if curl -sS -m 8 -o /dev/null \
       -H "Authorization: Bearer ${MODEL_API_KEY}" \
       "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
    echo "[k3-ns87] relay started on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND} pid=$!"
  else
    echo "[k3-ns87] FATAL: relay failed on :${RELAY_PORT}" >&2
    tail -n 40 /dev/shm/api_relay_k3_tcp.log 2>/dev/null || true
    exit 1
  fi
fi

echo "=== ping k3 through login-node relay (no reasoning_effort field) ==="
ping_ok=0
for ping_try in 1 2 3; do
  if curl -sS -m 45 -w "\nHTTP %{http_code} time=%{time_total}\n" \
       -H "Authorization: Bearer ${MODEL_API_KEY}" \
       -H "Content-Type: application/json" \
       -d "{\"model\":\"${MODEL_NAME}\",\"messages\":[{\"role\":\"user\",\"content\":\"ping\"}],\"max_tokens\":8}" \
       "http://127.0.0.1:${RELAY_PORT}/v1/chat/completions" | tail -n 8; then
    ping_ok=1
    break
  fi
  echo "[k3-ns87] ping try ${ping_try} failed; retrying..."
  sleep 3
done
if [[ "$ping_ok" != 1 ]]; then
  echo "[k3-ns87] FATAL: k3 ping via :${RELAY_PORT} failed" >&2
  exit 1
fi

mapfile -t TASKS < "${TASK_LIST:-${SCRIPT_DIR}/need_subagent_uneval_87.txt}"
if (( ${#TASKS[@]} != 87 )); then
  echo "[k3-ns87] FATAL: expected 87 tasks, got ${#TASKS[@]}" >&2
  exit 1
fi

echo "=== k3 need-subagent leftover c=1 ($(date)) ==="
echo "Model:    $MODEL_NAME @ $MODEL_API_URL  effort=<omitted, backend default>"
echo "Relay:    login 0.0.0.0:${RELAY_PORT} -> ${RELAY_UPSTREAM_HOST}:${RELAY_UPSTREAM_PORT}"
echo "Dump:     $DUMP_ROOT  run_id=$RUN_ID"
echo "Tasks:    ${#TASKS[@]}  node=$SLURM_NODELIST  concurrent=$MAX_CONCURRENT  cap=$MAX_CONCURRENT_CAP  cpus=$SLURM_CPUS mem=$SLURM_MEM pg_base=$PG_PORT_BASE"
echo "Isolation: PG ${PG_PORT_BASE}+ / WC $((PG_PORT_BASE + 10000))+ / relay :${RELAY_PORT} on gnho019 vs mm3-r134 gnho014 PG 41000+"
echo "Skip:     47 already in dumps/kimi-code_k3-par47-noeffort"
echo "Subagents: default coder/explore/plan (KIMI_SUBAGENTS unset)"
echo "Harness:  live kimi_harness"
echo ""

cd "$PROJECT_ROOT"
exec bash "${DS_SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
