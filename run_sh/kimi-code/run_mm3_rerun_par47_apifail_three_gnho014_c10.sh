#!/bin/bash
set -euo pipefail

# MiniMax-M3 official-API auto-subagent rerun of 89 API-invalid cases from
# dumps/kimi-code_MiniMax-M3-par47-noeffort (latest slot per case):
#   57 provider_invalid (insufficient balance, ~13-20s, no eval)
#   32 case_failed whose last log hit is an API/provider error
#      (insufficient balance / RateLimitError / 502 / APIConnectionError)
# Genuine model case_failed (13) are NOT in this list.
#
# KIMI_SUBAGENTS=three via run_on_slurm_three.sh. Concurrent=10 on gnho014.
# New dump: dumps/kimi-code_MiniMax-M3-three-apifail
#
# Isolation vs current tenants:
#   interactive 89707   gnho014  2 CPU / 12G
#   leftover mm3t       PG 42000+ (job finished; shm cleaned)
#   this job            gnho014  PG 43000+ / WC 53000+ / relay :19330 / 80 CPU

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
DS_SCRIPT_DIR="${PROJECT_ROOT}/run_sh/kimi-code-deepseek-v4-flash"
PYTHON_BIN="${PYTHON_BIN:-/storage/lintaoLab/bowending/miniconda3/envs/dbw_dev/bin/python3}"

export KIMI_CONFIG_ENV="${KIMI_CONFIG_ENV:-${SCRIPT_DIR}/config.official-sa167.env}"
# shellcheck disable=SC1091
source "$KIMI_CONFIG_ENV"

export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
unset KIMI_MODEL_THINKING_EFFORT || true
export KIMI_SUBAGENTS=three
unset KIMI_PLAN_FIRST || true

export RELAY_API_KEY="${RELAY_API_KEY:-$MODEL_API_KEY}"
export RELAY_PORT=19330
export RELAY_SKIP_AUTOSTART=1
export RELAY_KIND=https
export HTTPS_RELAY_PROXY="${HTTPS_RELAY_PROXY:-http://127.0.0.1:7893}"
export RELAY_BACKLOG=256
export DEEPSEEK_UPSTREAM_HOST=api.minimaxi.com
export DEEPSEEK_UPSTREAM_PORT=443
export MODEL_API_URL="http://192.168.180.240:${RELAY_PORT}"
export PG_PORT_BASE=43000
export MOCK_PORT_WAIT_LOOPS=720
export MAX_CONCURRENT=10
export MAX_CONCURRENT_CAP=10

export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_MiniMax-M3-three-apifail}"
export RUN_ID="${RUN_ID:-mm3af-$(date +%Y%m%d-%H%M%S)}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/tpl_mm3af_${UID}}"
export PG_RUNTIME_ROOT="${PG_RUNTIME_ROOT:-/dev/shm/tpg_mm3af_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho014}"
export SLURM_MEM="${SLURM_MEM:-256G}"
export SLURM_CPUS="${SLURM_CPUS:-80}"
export SLURM_TIME="${SLURM_TIME:-36:00:00}"
export SLURM_JOB_NAME="${SLURM_JOB_NAME:-mm3-af89}"

mkdir -p "$DUMP_ROOT"

if curl -sS -m 8 -o /dev/null \
     "http://127.0.0.1:${RELAY_PORT}/healthz"; then
  echo "[mm3-af89] relay healthy on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND}"
else
  nohup env RELAY_BACKLOG=256 \
    DEEPSEEK_UPSTREAM_HOST=api.minimaxi.com \
    DEEPSEEK_UPSTREAM_PORT=443 \
    "$PYTHON_BIN" "${DS_SCRIPT_DIR}/https_deepseek_relay.py" \
    "$RELAY_PORT" "$HTTPS_RELAY_PROXY" \
    >/dev/shm/api_relay_mm3_official_sa167.log 2>&1 &
  sleep 1
  if curl -sS -m 8 -o /dev/null \
       "http://127.0.0.1:${RELAY_PORT}/healthz"; then
    echo "[mm3-af89] relay started on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND} pid=$!"
  else
    echo "[mm3-af89] FATAL: HTTPS relay failed on :${RELAY_PORT}" >&2
    tail -n 40 /dev/shm/api_relay_mm3_official_sa167.log 2>/dev/null || true
    exit 1
  fi
fi

echo "=== ping MiniMax-M3 through login HTTPS relay :${RELAY_PORT} (no reasoning_effort) ==="
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
  echo "[mm3-af89] ping try ${ping_try} failed; retrying..."
  sleep 3
done
if [[ "$ping_ok" != 1 ]]; then
  echo "[mm3-af89] FATAL: MiniMax-M3 ping via :${RELAY_PORT} failed" >&2
  exit 1
fi

mapfile -t TASKS < "${TASK_LIST:-${SCRIPT_DIR}/rerun_par47_apifail_89.txt}"
if (( ${#TASKS[@]} != 89 )); then
  echo "[mm3-af89] FATAL: expected 89 tasks, got ${#TASKS[@]}" >&2
  exit 1
fi

echo "=== MiniMax-M3 official par47 API-fail KIMI_SUBAGENTS=three c=10 ($(date)) ==="
echo "Model:    $MODEL_NAME @ $MODEL_API_URL  KIMI_SUBAGENTS=$KIMI_SUBAGENTS  effort=<omitted>"
echo "Relay:    login 0.0.0.0:${RELAY_PORT} -> https://api.minimaxi.com via ${HTTPS_RELAY_PROXY}"
echo "Dump:     $DUMP_ROOT  run_id=$RUN_ID"
echo "Tasks:    ${#TASKS[@]}  node=$SLURM_NODELIST  concurrent=$MAX_CONCURRENT  cap=$MAX_CONCURRENT_CAP  cpus=$SLURM_CPUS mem=$SLURM_MEM pg_base=$PG_PORT_BASE"
echo "Source:   dumps/kimi-code_MiniMax-M3-par47-noeffort latest-slot API-invalid"
echo "Subagents: plan/coder/explore (KIMI_SUBAGENTS=three via run_on_slurm_three.sh)"
echo "Harness:  live kimi_harness"
echo ""

cd "$PROJECT_ROOT"
exec bash "${DS_SCRIPT_DIR}/run_on_slurm_three.sh" "${TASKS[@]}"
