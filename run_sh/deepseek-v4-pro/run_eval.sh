#!/bin/bash
# Evaluate 3 toolathlon_gym tasks with deepseek-v4-pro via OpenAI-compatible API.
#
# IMPORTANT — concurrency:
#   These tasks share a single PostgreSQL and their preprocess scripts clear
#   overlapping schemas (especially `email`, also notion/gcal/gsheet). Running
#   them in parallel WOULD cross-contaminate DB state and invalidate eval.
#   Filesystem workspaces ARE isolated (each task gets its own ephemeral enroot
#   rootfs), but DB is not — so this script runs tasks SEQUENTIALLY.
#
# Usage:
#   bash run_sh/deepseek-v4-pro/run_eval.sh
#   bash run_sh/deepseek-v4-pro/run_eval.sh arxiv-lit-review-gsheet   # one task
#
# Results:
#   dumps/deepseek-v4-pro/<task>/<timestamp>/

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
RUNTIME_ROOT="${TOOLATHLON_EVAL_DOCKER_ROOT:-/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers}"
CONFIG_ENV="${SCRIPT_DIR}/config.env"

# shellcheck disable=SC1091
source "$RUNTIME_ROOT/env.sh"
# shellcheck disable=SC1091
source "$CONFIG_ENV"

# Internal API must bypass HTTP proxy
export no_proxy="${no_proxy:+$no_proxy,}127.0.0.1,localhost,172.16.55.136"
export NO_PROXY="$no_proxy"

export MODEL_NAME MODEL_PLATFORM MODEL_PROVIDER MODEL_API_KEY MODEL_API_URL
export MODEL_GREEDY MODEL_TEMPERATURE MODEL_TOP_P MODEL_N
export DUMP_ROOT
export LOCK_FILE="${DUMP_ROOT}/.run.lock"
MAX_STEPS="${MAX_STEPS:-100}"

mkdir -p "$DUMP_ROOT"
cd "$PROJECT_ROOT"

TASKS=(
  arxiv-lit-review-gsheet
  canvas-assignment-effectiveness-ppt-notion-email
  fetch-howtocook-catering-excel-gcal-email
)

# Optional: override task list from CLI args
if [[ $# -gt 0 ]]; then
  TASKS=("$@")
fi

log() { echo "[$(date +%H:%M:%S)] $*"; }

log "=============================================="
log "  Model:     $MODEL_NAME"
log "  Endpoint:  $MODEL_API_URL"
log "  Sampling:  greedy=${MODEL_GREEDY:-?} temp=${MODEL_TEMPERATURE:-?} top_p=${MODEL_TOP_P:-?}"
log "  Dump root: $DUMP_ROOT"
log "  Tasks:     ${#TASKS[@]} (sequential — shared PG)"
log "  Max steps: $MAX_STEPS"
log "=============================================="

# Ensure postgres is up
if ! PGPASSWORD="${PGPASSWORD}" pg_isready -h "${PGHOST}" -p "${PGPORT}" -U "${PGUSER}" >/dev/null 2>&1; then
  log "Starting postgres ..."
  bash "$PROJECT_ROOT/scripts/enroot_postgres.sh start"
fi

# Quick API sanity check
if command -v curl >/dev/null 2>&1; then
  code="$(curl -sS --connect-timeout 8 -o /dev/null -w '%{http_code}' \
    -H "Authorization: Bearer ${MODEL_API_KEY}" \
    "${MODEL_API_URL%/}/v1/models" || true)"
  if [[ "$code" != "200" ]]; then
    log "[warn] API /v1/models returned HTTP $code — continuing anyway"
  else
    log "API reachable (HTTP 200)."
  fi
fi

SUMMARY="$DUMP_ROOT/summary_$(date +%Y%m%d-%H%M%S).csv"
echo "task,status,exit_code,output_dir,duration_s" > "$SUMMARY"

FAILED=0
for TASK in "${TASKS[@]}"; do
  log "---------- START $TASK ----------"
  START_TS=$(date +%s)
  set +e
  bash "$PROJECT_ROOT/scripts/enroot_run_task.sh" "$TASK" "$MAX_STEPS"
  RC=$?
  set -e
  END_TS=$(date +%s)
  DUR=$((END_TS - START_TS))

  # Latest output dir for this task
  OUT="$(ls -1dt "$DUMP_ROOT/$TASK"/*/ 2>/dev/null | head -1 || true)"
  if [[ $RC -eq 0 ]]; then
    STATUS=success
  else
    STATUS=failed
    FAILED=$((FAILED + 1))
  fi
  echo "${TASK},${STATUS},${RC},${OUT%,},${DUR}" >> "$SUMMARY"
  log "---------- DONE $TASK status=$STATUS exit=$RC (${DUR}s) ----------"
done

log "=============================================="
log "  Finished. Failed: $FAILED / ${#TASKS[@]}"
log "  Summary: $SUMMARY"
log "  Results under: $DUMP_ROOT"
log "=============================================="

exit "$FAILED"
