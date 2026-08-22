#!/bin/bash
set -euo pipefail

# GLM-5.3 auto-subagent rerun of 13 allowlist/API cases on gnho014, concurrent=15.
# Official-compatible BLSC gateway via dedicated login HTTPS relay :19323.
# Dumps stay in dumps/kimi-code_glm-5-3-par47-noeffort (new run_id).
# Live kimi_harness (readonly infix glob).
#
# Isolation vs current gnho014 tenant:
#   interactive 89707  2 CPU / 12G (no PG 38xxx)
#   leftover listeners 30216 / 34263 / 34473 / 38583 / 43387 (not our 38000-38014)
#   this job           PG 38000+ / WC 48000+ / relay :19323
# Login :19321 is the old 8317 glm relay; do not reuse.

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
export PG_PORT_BASE=38000
export MOCK_PORT_WAIT_LOOPS=720
export MAX_CONCURRENT=15
export MAX_CONCURRENT_CAP=15

export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_glm-5-3-par47-noeffort}"
export RUN_ID="${RUN_ID:-g53b-$(date +%Y%m%d-%H%M%S)}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/tpl_g53b_${UID}}"
export PG_RUNTIME_ROOT="${PG_RUNTIME_ROOT:-/dev/shm/tpg_g53b_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho014}"
# c=15 ≈ 6 CPU/worker. gnho014 has 126 idle cores beside interactive 89707 (2 CPU).
export SLURM_MEM="${SLURM_MEM:-384G}"
export SLURM_CPUS="${SLURM_CPUS:-90}"
export SLURM_TIME="${SLURM_TIME:-16:00:00}"
export SLURM_JOB_NAME="${SLURM_JOB_NAME:-glm53-c15}"

mkdir -p "$DUMP_ROOT"

if curl -sS -m 8 -o /dev/null \
     -H "Authorization: Bearer ${MODEL_API_KEY}" \
     "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
  echo "[glm53-c15] relay healthy on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND}"
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
    echo "[glm53-c15] relay started on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND} pid=$!"
  else
    echo "[glm53-c15] FATAL: HTTPS relay failed on :${RELAY_PORT}" >&2
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
  echo "[glm53-c15] ping try ${ping_try} failed; retrying..."
  sleep 3
done
if [[ "$ping_ok" != 1 ]]; then
  echo "[glm53-c15] FATAL: GLM-5.3 ping via :${RELAY_PORT} failed" >&2
  exit 1
fi

mapfile -t TASKS < "${TASK_LIST:-${SCRIPT_DIR}/rerun_allowfix_13.txt}"
if (( ${#TASKS[@]} != 13 )); then
  echo "[glm53-c15] FATAL: expected 13 tasks, got ${#TASKS[@]}" >&2
  exit 1
fi

echo "=== GLM-5.3 allowfix/API rerun c=15 ($(date)) ==="
echo "Model:    $MODEL_NAME @ $MODEL_API_URL  effort=<omitted>"
echo "Relay:    login 0.0.0.0:${RELAY_PORT} -> https://llmapi.blsc.cn via ${HTTPS_RELAY_PROXY}"
echo "Dump:     $DUMP_ROOT  run_id=$RUN_ID"
echo "Tasks:    ${#TASKS[@]}  node=$SLURM_NODELIST  concurrent=$MAX_CONCURRENT  cap=$MAX_CONCURRENT_CAP  cpus=$SLURM_CPUS mem=$SLURM_MEM pg_base=$PG_PORT_BASE"
echo "Isolation: PG ${PG_PORT_BASE}+ / WC $((PG_PORT_BASE + 10000))+ vs gnho014 leftover 34xxx/38583; interactive 89707 has no PG 38xxx"
echo "Subagents: default coder/explore/plan (KIMI_SUBAGENTS unset)"
echo "Harness:  live kimi_harness (readonly infix glob + corrected orchestration prompt)"
echo ""

cd "$PROJECT_ROOT"
exec bash "${DS_SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
