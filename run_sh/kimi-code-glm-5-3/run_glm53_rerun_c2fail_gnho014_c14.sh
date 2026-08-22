#!/bin/bash
set -euo pipefail

# GLM-5.3 auto-subagent rerun of leftover glm53-c2-20260819 failures.
# Same BLSC stack as g53b (config.blsc-c15.env, relay :19323, no effort).
# Skip the 13 tasks already rerun today as g53b-20260820-194347.
# Dumps stay in dumps/kimi-code_glm-5-3-par47-noeffort (new run_id).
# Live kimi_harness (readonly infix glob).
#
# Isolation vs current gnho014 tenant:
#   interactive 89707  2 CPU / 12G (no PG 37xxx)
#   leftover g53b      PG 38000+ / WC 48000+ / relay :19323 (reuse relay)
#   this job           PG 37000+ / WC 47000+ / relay :19323

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
DS_SCRIPT_DIR="${PROJECT_ROOT}/run_sh/kimi-code-deepseek-v4-flash"
PYTHON_BIN="${PYTHON_BIN:-/storage/lintaoLab/bowending/miniconda3/envs/dbw_dev/bin/python3}"

export KIMI_CONFIG_ENV="${KIMI_CONFIG_ENV:-${SCRIPT_DIR}/config.blsc-c15.env}"
# shellcheck disable=SC1091
source "$KIMI_CONFIG_ENV"

export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
unset KIMI_MODEL_THINKING_EFFORT || true
unset KIMI_SUBAGENTS || true
unset KIMI_PLAN_FIRST || true

export RELAY_API_KEY="${RELAY_API_KEY:-$MODEL_API_KEY}"
export RELAY_PORT=19323
export RELAY_SKIP_AUTOSTART=1
export RELAY_KIND=https
export HTTPS_RELAY_PROXY="${HTTPS_RELAY_PROXY:-http://127.0.0.1:7893}"
export RELAY_BACKLOG=256
export DEEPSEEK_UPSTREAM_HOST=llmapi.blsc.cn
export DEEPSEEK_UPSTREAM_PORT=443
export MODEL_API_URL="http://192.168.180.240:${RELAY_PORT}"
# Keep away from leftover g53b 38000+ and historic 38583 on gnho014.
export PG_PORT_BASE=37000
export MOCK_PORT_WAIT_LOOPS=720
export MAX_CONCURRENT=14
export MAX_CONCURRENT_CAP=14

export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_glm-5-3-par47-noeffort}"
export RUN_ID="${RUN_ID:-g53c-$(date +%Y%m%d-%H%M%S)}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/tpl_g53c_${UID}}"
export PG_RUNTIME_ROOT="${PG_RUNTIME_ROOT:-/dev/shm/tpg_g53c_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho014}"
# c=14 ≈ 6 CPU/worker. gnho014 has 126 idle cores beside interactive 89707.
export SLURM_MEM="${SLURM_MEM:-384G}"
export SLURM_CPUS="${SLURM_CPUS:-84}"
export SLURM_TIME="${SLURM_TIME:-16:00:00}"
export SLURM_JOB_NAME="${SLURM_JOB_NAME:-glm53-c14}"

mkdir -p "$DUMP_ROOT"

if curl -sS -m 8 -o /dev/null \
     -H "Authorization: Bearer ${MODEL_API_KEY}" \
     "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
  echo "[glm53-c14] relay healthy on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND}"
else
  nohup env RELAY_BACKLOG=256 MAX_OUTPUT_TOKENS=128000 \
    DEEPSEEK_UPSTREAM_HOST=llmapi.blsc.cn \
    DEEPSEEK_UPSTREAM_PORT=443 \
    "$PYTHON_BIN" "${DS_SCRIPT_DIR}/https_deepseek_relay.py" \
    "$RELAY_PORT" "$HTTPS_RELAY_PROXY" \
    >/dev/shm/api_relay_glm53_blsc.log 2>&1 &
  sleep 1
  if curl -sS -m 8 -o /dev/null \
       -H "Authorization: Bearer ${MODEL_API_KEY}" \
       "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
    echo "[glm53-c14] relay started on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND} pid=$!"
  else
    echo "[glm53-c14] FATAL: HTTPS relay failed on :${RELAY_PORT}" >&2
    tail -n 40 /dev/shm/api_relay_glm53_blsc.log 2>/dev/null || true
    exit 1
  fi
fi

echo "=== ping GLM-5.3 through login HTTPS relay :${RELAY_PORT} (no reasoning_effort) ==="
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
  echo "[glm53-c14] ping try ${ping_try} failed; retrying..."
  sleep 3
done
if [[ "$ping_ok" != 1 ]]; then
  echo "[glm53-c14] FATAL: GLM-5.3 ping via :${RELAY_PORT} failed" >&2
  exit 1
fi

mapfile -t TASKS < "${TASK_LIST:-${SCRIPT_DIR}/rerun_c2fail_14.txt}"
if (( ${#TASKS[@]} != 14 )); then
  echo "[glm53-c14] FATAL: expected 14 tasks, got ${#TASKS[@]}" >&2
  exit 1
fi

SKIP_LIST="${SCRIPT_DIR}/rerun_allowfix_13.txt"
overlap=0
for t in "${TASKS[@]}"; do
  if grep -qxF "$t" "$SKIP_LIST"; then
    echo "[glm53-c14] FATAL: $t was already rerun today in g53b ($SKIP_LIST)" >&2
    overlap=1
  fi
done
if (( overlap != 0 )); then
  exit 1
fi

echo "=== GLM-5.3 c2-fail BLSC rerun c=14 ($(date)) ==="
echo "Model:    $MODEL_NAME @ $MODEL_API_URL  effort=<omitted>"
echo "Relay:    login 0.0.0.0:${RELAY_PORT} -> https://llmapi.blsc.cn via ${HTTPS_RELAY_PROXY}"
echo "Dump:     $DUMP_ROOT  run_id=$RUN_ID"
echo "Tasks:    ${#TASKS[@]}  node=$SLURM_NODELIST  concurrent=$MAX_CONCURRENT  cap=$MAX_CONCURRENT_CAP  cpus=$SLURM_CPUS mem=$SLURM_MEM pg_base=$PG_PORT_BASE"
echo "Skip:     13 g53b tasks in $SKIP_LIST"
echo "Isolation: PG ${PG_PORT_BASE}+ / WC $((PG_PORT_BASE + 10000))+ vs leftover g53b 38000+; interactive 89707"
echo "Subagents: default coder/explore/plan (KIMI_SUBAGENTS unset)"
echo "Harness:  live kimi_harness (readonly infix glob + corrected orchestration prompt)"
echo ""

cd "$PROJECT_ROOT"
exec bash "${DS_SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
