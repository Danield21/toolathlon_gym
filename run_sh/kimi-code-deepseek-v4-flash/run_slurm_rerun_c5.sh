#!/bin/bash
# Slurm wrapper for rerun-c5 (2026-08-15): 68 cases = c4 batch (77) minus its
# 9 successes. Runs on the image rebuilt after the c4 analysis fixes:
#   P0-A evaluator launch_time weekday strip
#   P0-B AGENT_SQSH fixed to build-produced toolathlon-pack.sqsh
#     (never the contaminated rootfs directory)
#   P1 servers: gcal toInstantString+recurrence INSERT, word abs-path
#     validation, notion openapi oneOf + normalizeParent
#   P1 evals: 11 evaluator/task/GT fixes (PPT table cells, TIMESTAMPTZ
#     literals, dynamic GT, regular_price COALESCE, block_data, base-name
#     aggregation, flexible title/parent extraction)
#   P2 task.md header-row requirements (sf-sales-discount-analysis,
#     sf-sales-customer-loyalty)
#   GT remade: afrobeat tracklist + music-schedule playlist (transcript-
#     derived, artist-consistency eval)
#
# Walltime: 68 cases × ~35min avg / 8 slots ≈ 5.6h + margin → 12h.
#
# Usage (from toolathlon_gym/):
#   RELAY_API_KEY=sk-remote-... bash run_sh/kimi-code-deepseek-v4-flash/run_slurm_rerun_c5.sh
#   MAX_CONCURRENT=8 RELAY_API_KEY=sk-remote-... bash .../run_slurm_rerun_c5.sh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"

export KIMI_CONFIG_ENV="${SCRIPT_DIR}/config.c2.env"

# Source the config on the login node so MODEL_NAME / MODEL_API_URL / MODEL_API_KEY
# are explicitly exported into the environment srun inherits.
# shellcheck disable=SC1090
source "$KIMI_CONFIG_ENV"
export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S

export DUMP_ROOT="${PROJECT_ROOT}/dumps/kimi-code_deepseek-v4-flash-linslab-rerun-c5_$(date +%Y%m%d-%H%M%S)"
export PG_PORT_BASE="${PG_PORT_BASE:-27052}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/toolathlon_pg_port_leases_deepseek_rerun_c5_${UID}}"

TASK_FILE="${SCRIPT_DIR}/rerun_c5/tasks_c5.txt"
[[ -f "$TASK_FILE" ]] || { echo "[error] missing $TASK_FILE" >&2; exit 1; }
mapfile -t TASKS < <(grep -E '^[a-z0-9][a-z0-9-]+$' "$TASK_FILE")

echo "[rerun-c5-slurm] ${#TASKS[@]} tasks"
echo "[rerun-c5-slurm] KIMI_CONFIG_ENV=$KIMI_CONFIG_ENV"
echo "[rerun-c5-slurm] MODEL_NAME=$MODEL_NAME"
echo "[rerun-c5-slurm] DUMP_ROOT=$DUMP_ROOT"
echo "[rerun-c5-slurm] PG_PORT_BASE=$PG_PORT_BASE"
echo "[rerun-c5-slurm] MAX_CONCURRENT=${MAX_CONCURRENT:-8}"

exec bash "${SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
