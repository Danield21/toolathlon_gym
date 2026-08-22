#!/bin/bash
set -euo pipefail

# Official DeepSeek auto-subagent probe: will the new three-mode main prompt
# (Lumenport `plan` example, NOT plan-first) actually dispatch `plan`?
# 20 typical cases from dumps/kimi-code_deepseek-v4-flash-subagent; that dump
# had 0 plan calls. Roster pinned with KIMI_SUBAGENTS=three (alias 3).
#
# Isolation vs current tenants:
#   mm3-af43 91406 gnho014   PG 44000+ / relay :19330 / 120 CPU
#   interactive 89707 gnho014
#   quxingyu m3_rope_layout gnho019 32 CPU (this job takes no GPU)
#   leftover sa12            PG 33000+ / relay :19329
#   this job                 gnho019 PG 35000+ / WC 45000+ / relay :19331 / 96 CPU

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
PYTHON_BIN="${PYTHON_BIN:-/storage/lintaoLab/bowending/miniconda3/envs/dbw_dev/bin/python3}"

export KIMI_CONFIG_ENV="${KIMI_CONFIG_ENV:-${SCRIPT_DIR}/config.official-sa12.env}"
# shellcheck disable=SC1091
source "$KIMI_CONFIG_ENV"

export MODEL_NAME MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
export KIMI_MODEL_THINKING_EFFORT="${KIMI_MODEL_THINKING_EFFORT:-high}"
export KIMI_SUBAGENTS=3
unset KIMI_PLAN_FIRST || true

export RELAY_API_KEY="${RELAY_API_KEY:-$MODEL_API_KEY}"
export RELAY_PORT=19331
export RELAY_SKIP_AUTOSTART=1
export RELAY_KIND=https
export HTTPS_RELAY_PROXY="${HTTPS_RELAY_PROXY:-http://127.0.0.1:7893}"
export RELAY_BACKLOG=256
export DEEPSEEK_UPSTREAM_HOST=api.deepseek.com
export DEEPSEEK_UPSTREAM_PORT=443
export MODEL_API_URL="http://192.168.180.240:${RELAY_PORT}"
export PG_PORT_BASE=35000
export MOCK_PORT_WAIT_LOOPS=720
export MAX_CONCURRENT=20
export MAX_CONCURRENT_CAP=20

export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_deepseek-v4-flash-three-planex20}"
export RUN_ID="${RUN_ID:-px20-$(date +%Y%m%d-%H%M%S)}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/tpl_px20_${UID}}"
export PG_RUNTIME_ROOT="${PG_RUNTIME_ROOT:-/dev/shm/tpg_px20_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho019}"
export SLURM_MEM="${SLURM_MEM:-384G}"
export SLURM_CPUS="${SLURM_CPUS:-96}"
export SLURM_TIME="${SLURM_TIME:-16:00:00}"
export SLURM_JOB_NAME="${SLURM_JOB_NAME:-ds-px20}"

TASK_LIST_FILE="${TASK_LIST:-${SCRIPT_DIR}/planex20_cases.txt}"
mkdir -p "$DUMP_ROOT"

if curl -sS -m 8 -o /dev/null \
     "http://127.0.0.1:${RELAY_PORT}/healthz"; then
  echo "[ds-px20] relay healthy on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND}"
else
  nohup env RELAY_BACKLOG=256 \
    DEEPSEEK_UPSTREAM_HOST=api.deepseek.com \
    DEEPSEEK_UPSTREAM_PORT=443 \
    "$PYTHON_BIN" "${SCRIPT_DIR}/https_deepseek_relay.py" \
    "$RELAY_PORT" "$HTTPS_RELAY_PROXY" \
    >/dev/shm/api_relay_ds_px20.log 2>&1 &
  sleep 1
  if curl -sS -m 8 -o /dev/null \
       "http://127.0.0.1:${RELAY_PORT}/healthz"; then
    echo "[ds-px20] relay started on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND} pid=$!"
  else
    echo "[ds-px20] FATAL: HTTPS relay failed on :${RELAY_PORT}" >&2
    tail -n 40 /dev/shm/api_relay_ds_px20.log 2>/dev/null || true
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
  echo "[ds-px20] ping try ${ping_try} failed; retrying..."
  sleep 3
done
if [[ "$ping_ok" != 1 ]]; then
  echo "[ds-px20] FATAL: deepseek-v4-flash ping via :${RELAY_PORT} failed" >&2
  exit 1
fi

mapfile -t TASKS < "$TASK_LIST_FILE"
if (( ${#TASKS[@]} != 20 )); then
  echo "[ds-px20] FATAL: expected 20 tasks, got ${#TASKS[@]}" >&2
  exit 1
fi

echo "=== Official DeepSeek three-mode plan-example probe c=20 ($(date)) ==="
echo "Model:    $MODEL_NAME @ $MODEL_API_URL  KIMI_SUBAGENTS=$KIMI_SUBAGENTS  effort=${KIMI_MODEL_THINKING_EFFORT}"
echo "Relay:    login 0.0.0.0:${RELAY_PORT} -> https://api.deepseek.com via ${HTTPS_RELAY_PROXY}"
echo "Dump:     $DUMP_ROOT  run_id=$RUN_ID"
echo "Tasks:    ${#TASKS[@]}  node=$SLURM_NODELIST  concurrent=$MAX_CONCURRENT  cap=$MAX_CONCURRENT_CAP  cpus=$SLURM_CPUS mem=$SLURM_MEM pg_base=$PG_PORT_BASE"
echo "Subagents: coder/explore/plan (KIMI_SUBAGENTS=3 via run_on_slurm_three.sh), plan-first off"
echo "Harness:  live kimi_harness (examples_legacy.md includes Lumenport plan example)"
echo ""

cd "$PROJECT_ROOT"
exec bash "${SCRIPT_DIR}/run_on_slurm_three.sh" "${TASKS[@]}"
