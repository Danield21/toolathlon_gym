#!/bin/bash
set -euo pipefail

# GLM-5.3 auto-subagent rerun of 13 cases on gnho014, concurrent=1.
# Uses the live kimi_harness copy (readonly infix glob + orchestration prompt).
# Isolated from gnho014 job 90361 (gpt56sol-xhigh-single5-c5, 112CPU/800G,
# PG 37432+, WC 47432+, mock 30xxx, c=5).

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
DS_SCRIPT_DIR="${PROJECT_ROOT}/run_sh/kimi-code-deepseek-v4-flash"
PYTHON_BIN="${PYTHON_BIN:-/storage/lintaoLab/bowending/miniconda3/envs/dbw_dev/bin/python3}"

export KIMI_CONFIG_ENV="${KIMI_CONFIG_ENV:-${SCRIPT_DIR}/config.rerun-allowfix.env}"
# shellcheck disable=SC1091
source "$KIMI_CONFIG_ENV"

export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
unset KIMI_MODEL_THINKING_EFFORT || true
export RELAY_API_KEY="${RELAY_API_KEY:-$MODEL_API_KEY}"
# Unconditional overrides so sourced defaults cannot keep PG 29900 / c=2.
export RELAY_PORT=19321
export RELAY_SKIP_AUTOSTART=1
export RELAY_KIND=tcp
export RELAY_UPSTREAM_HOST=172.16.55.136
export RELAY_UPSTREAM_PORT=8317
export MODEL_API_URL="http://192.168.180.240:${RELAY_PORT}"
export PG_PORT_BASE=38000
export MOCK_PORT_WAIT_LOOPS=720

unset KIMI_SUBAGENTS || true

export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_glm-5-3-par47-noeffort}"
# Short names: Postgres Unix sockets cap at 107 bytes; a long runtime root +
# run_id made the first attempt pg_fail ("socket path is too long").
export RUN_ID="${RUN_ID:-glm53r-$(date +%Y%m%d-%H%M%S)}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/tpl_g53r_${UID}}"
export PG_RUNTIME_ROOT="${PG_RUNTIME_ROOT:-/dev/shm/tpg_g53r_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho014}"
# Serial worker. gnho014 already holds 90361 (112 CPU) + interactive 89707 (2 CPU);
# 14 idle cores remain. 8 CPU / 64G matches the other c=1 reruns.
export SLURM_MEM="${SLURM_MEM:-64G}"
export SLURM_CPUS="${SLURM_CPUS:-8}"
export MAX_CONCURRENT=1
# 13 * 2h timeout worst-case + enqueue slack.
export SLURM_TIME="${SLURM_TIME:-36:00:00}"
export SLURM_JOB_NAME="${SLURM_JOB_NAME:-glm53-c1-rerun}"

mkdir -p "$DUMP_ROOT"

if curl -sS -m 8 -o /dev/null \
     -H "Authorization: Bearer ${MODEL_API_KEY}" \
     "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
  echo "[glm53-c1-rerun] relay healthy on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND}"
else
  nohup "$PYTHON_BIN" "${DS_SCRIPT_DIR}/api_relay.py" \
    "$RELAY_PORT" "$RELAY_UPSTREAM_HOST" "$RELAY_UPSTREAM_PORT" \
    >/dev/shm/api_relay_glm53_tcp.log 2>&1 &
  sleep 1
  if curl -sS -m 8 -o /dev/null \
       -H "Authorization: Bearer ${MODEL_API_KEY}" \
       "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
    echo "[glm53-c1-rerun] relay started on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND} pid=$!"
  else
    echo "[glm53-c1-rerun] FATAL: relay failed on :${RELAY_PORT}" >&2
    tail -n 40 /dev/shm/api_relay_glm53_tcp.log 2>/dev/null || true
    exit 1
  fi
fi

echo "=== ping glm-5.3 through login-node relay (no reasoning_effort field) ==="
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
  echo "[glm53-c1-rerun] ping try ${ping_try} failed; retrying..."
  sleep 3
done
if [[ "$ping_ok" != 1 ]]; then
  echo "[glm53-c1-rerun] FATAL: glm-5.3 ping via :${RELAY_PORT} failed" >&2
  exit 1
fi

mapfile -t TASKS < "${TASK_LIST:-${SCRIPT_DIR}/rerun_allowfix_13.txt}"

echo "=== GLM-5.3 allowfix/API rerun c=1 ($(date)) ==="
echo "Model:    $MODEL_NAME @ $MODEL_API_URL  effort=<omitted, backend default>"
echo "Relay:    login 0.0.0.0:${RELAY_PORT} -> ${RELAY_UPSTREAM_HOST}:${RELAY_UPSTREAM_PORT}"
echo "Dump:     $DUMP_ROOT  run_id=$RUN_ID"
echo "Tasks:    ${#TASKS[@]}  node=$SLURM_NODELIST  concurrent=$MAX_CONCURRENT  cpus=$SLURM_CPUS mem=$SLURM_MEM pg_base=$PG_PORT_BASE"
echo "Isolation: PG ${PG_PORT_BASE}+ / WC $((PG_PORT_BASE + 10000))+ vs gnho014 job 90361 PG 37432+ / WC 47432+"
echo "Subagents: default coder/explore/plan (KIMI_SUBAGENTS unset)"
echo "Harness:  live kimi_harness (readonly infix glob + corrected orchestration prompt)"
echo ""

cd "$PROJECT_ROOT"
exec bash "${DS_SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
