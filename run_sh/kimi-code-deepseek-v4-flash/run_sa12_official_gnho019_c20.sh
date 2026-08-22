#!/bin/bash
set -euo pipefail

# Official DeepSeek auto-subagent rerun of the 12 allowlist-hit cases from
# dev_docs/2026-08-20-explore-plan-readonly-allowlist-hits.md §4.5.
# Default coder/explore/plan (KIMI_SUBAGENTS unset). Not plan-first.
# Dumps stay in dumps/kimi-code_deepseek-v4-flash-subagent (new run_id).
# Live kimi_harness (readonly infix glob).
#
# Isolation vs current gnho019 tenants:
#   mm3-c1-rerun 90369     PG 36000+ / WC 46000+ / relay :19326 / 8 CPU
#   quxingyu kv-m3-fit     4 CPU / 4 GPU (do not touch; this job takes no GPU)
#   leftover pfr20 relay   :19327
#   leftover sa105 relay   :19328
#   this job               PG 33000+ / WC 43000+ / relay :19329 / 96 CPU
# Official API: do not allocate gnho019 H800s.

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
PYTHON_BIN="${PYTHON_BIN:-/storage/lintaoLab/bowending/miniconda3/envs/dbw_dev/bin/python3}"

export KIMI_CONFIG_ENV="${KIMI_CONFIG_ENV:-${SCRIPT_DIR}/config.official-sa12.env}"
# shellcheck disable=SC1091
source "$KIMI_CONFIG_ENV"

export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
export KIMI_MODEL_THINKING_EFFORT="${KIMI_MODEL_THINKING_EFFORT:-high}"
unset KIMI_SUBAGENTS || true
unset KIMI_PLAN_FIRST || true

export RELAY_API_KEY="${RELAY_API_KEY:-$MODEL_API_KEY}"
export RELAY_PORT=19329
export RELAY_SKIP_AUTOSTART=1
export RELAY_KIND=https
export HTTPS_RELAY_PROXY="${HTTPS_RELAY_PROXY:-http://127.0.0.1:7893}"
export RELAY_BACKLOG=256
export DEEPSEEK_UPSTREAM_HOST=api.deepseek.com
export DEEPSEEK_UPSTREAM_PORT=443
export MODEL_API_URL="http://192.168.180.240:${RELAY_PORT}"
export PG_PORT_BASE=33000
export MOCK_PORT_WAIT_LOOPS=720
export MAX_CONCURRENT=20
export MAX_CONCURRENT_CAP=20

export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_deepseek-v4-flash-subagent}"
export RUN_ID="${RUN_ID:-saf12-$(date +%Y%m%d-%H%M%S)}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/tpl_sa12_${UID}}"
export PG_RUNTIME_ROOT="${PG_RUNTIME_ROOT:-/dev/shm/tpg_sa12_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho019}"
# 12 cases at c=20 all start together. ~6 CPU/worker. gnho019 has ~116 idle
# cores beside mm3 (8) and kv-m3-fit (4). No GPU: official API, 4/8 H800 busy.
export SLURM_MEM="${SLURM_MEM:-384G}"
export SLURM_CPUS="${SLURM_CPUS:-96}"
export SLURM_TIME="${SLURM_TIME:-16:00:00}"
export SLURM_JOB_NAME="${SLURM_JOB_NAME:-ds-sa12}"

mkdir -p "$DUMP_ROOT"

if curl -sS -m 8 -o /dev/null \
     "http://127.0.0.1:${RELAY_PORT}/healthz"; then
  echo "[ds-sa12] relay healthy on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND}"
else
  nohup env RELAY_BACKLOG=256 \
    DEEPSEEK_UPSTREAM_HOST=api.deepseek.com \
    DEEPSEEK_UPSTREAM_PORT=443 \
    "$PYTHON_BIN" "${SCRIPT_DIR}/https_deepseek_relay.py" \
    "$RELAY_PORT" "$HTTPS_RELAY_PROXY" \
    >/dev/shm/api_relay_ds_sa12.log 2>&1 &
  sleep 1
  if curl -sS -m 8 -o /dev/null \
       "http://127.0.0.1:${RELAY_PORT}/healthz"; then
    echo "[ds-sa12] relay started on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND} pid=$!"
  else
    echo "[ds-sa12] FATAL: HTTPS relay failed on :${RELAY_PORT}" >&2
    tail -n 40 /dev/shm/api_relay_ds_sa12.log 2>/dev/null || true
    exit 1
  fi
fi

echo "=== ping official deepseek-v4-flash through login HTTPS relay :${RELAY_PORT} ==="
ping_ok=0
for ping_try in 1 2 3; do
  if curl -sS -m 45 -w "\nHTTP %{http_code} time=%{time_total}\n" \
       -H "Authorization: Bearer ${MODEL_API_KEY}" \
       -H "Content-Type: application/json" \
       -d "{\"model\":\"${MODEL_NAME}\",\"messages\":[{\"role\":\"user\",\"content\":\"ping\"}],\"max_tokens\":8,\"reasoning_effort\":\"high\"}" \
       "http://127.0.0.1:${RELAY_PORT}/v1/chat/completions" | tail -n 8; then
    ping_ok=1
    break
  fi
  echo "[ds-sa12] ping try ${ping_try} failed; retrying..."
  sleep 3
done
if [[ "$ping_ok" != 1 ]]; then
  echo "[ds-sa12] FATAL: deepseek-v4-flash ping via :${RELAY_PORT} failed" >&2
  exit 1
fi

mapfile -t TASKS < "${TASK_LIST:-${SCRIPT_DIR}/rerun_subagent_allowfix_12.txt}"
if (( ${#TASKS[@]} != 12 )); then
  echo "[ds-sa12] FATAL: expected 12 tasks, got ${#TASKS[@]}" >&2
  exit 1
fi

echo "=== Official DeepSeek subagent allowfix 12 c=20 ($(date)) ==="
echo "Model:    $MODEL_NAME @ $MODEL_API_URL  KIMI_SUBAGENTS=unset  effort=${KIMI_MODEL_THINKING_EFFORT}"
echo "Relay:    login 0.0.0.0:${RELAY_PORT} -> https://api.deepseek.com via ${HTTPS_RELAY_PROXY}"
echo "Dump:     $DUMP_ROOT  run_id=$RUN_ID"
echo "Tasks:    ${#TASKS[@]}  node=$SLURM_NODELIST  concurrent=$MAX_CONCURRENT  cap=$MAX_CONCURRENT_CAP  cpus=$SLURM_CPUS mem=$SLURM_MEM pg_base=$PG_PORT_BASE"
echo "Isolation: PG ${PG_PORT_BASE}+ / WC $((PG_PORT_BASE + 10000))+ vs mm3 PG 36000+ / leftover pfr20 39000+ / leftover sa105 32000+"
echo "Subagents: default coder/explore/plan (KIMI_SUBAGENTS unset), plan-first off"
echo "Harness:  live kimi_harness (readonly infix glob + corrected orchestration prompt)"
echo ""

cd "$PROJECT_ROOT"
exec bash "${SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
