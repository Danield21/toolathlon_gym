#!/bin/bash
# Slurm wrapper for rerun-c4 (2026-08-15): 77 cases on the REBUILT image
# (data/toolathlon-pack, P0-P6 fixes baked in).
#   §3.1 A 类 22 (MCP venv 污染受害) + §3.2 B 类 1 (30216 冲突, P6 互斥锁兜底)
# + §3.3 C 类 6 (walltime 截断) + §4.3 新 case 48 (同事 49 剔除 howtoc-meal-plan-calendar)
#
# Port-conflict note: the only 30216 collision pair member in this batch is
# canvas-at-risk-intervention; its conflict partner playwright-sf-competitor-
# analysis-notion-excel is NOT in this batch, so the P6 mock-port mutex adds no
# extra serialization here. All 22 other mock ports in this batch are unique.
#
# Walltime: 77 cases × ~35min avg / 8 slots ≈ 6.4h + margin → 12h. The P1
# walltime-aware dispatch will stop scheduling in the final budget window and
# exit cleanly instead of being SIGKILLed (the c3b lesson).
#
# Usage (from toolathlon_gym/):
#   RELAY_API_KEY=sk-remote-... bash run_sh/kimi-code-deepseek-v4-flash/run_slurm_rerun_c4.sh
#   MAX_CONCURRENT=8 RELAY_API_KEY=sk-remote-... bash .../run_slurm_rerun_c4.sh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"

export KIMI_CONFIG_ENV="${SCRIPT_DIR}/config.c2.env"

# Source the config on the login node so MODEL_NAME / MODEL_API_URL / MODEL_API_KEY
# are explicitly exported into the environment srun inherits.
# shellcheck disable=SC1090
source "$KIMI_CONFIG_ENV"
export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S

export DUMP_ROOT="${PROJECT_ROOT}/dumps/kimi-code_deepseek-v4-flash-linslab-rerun-c4_$(date +%Y%m%d-%H%M%S)"
export PG_PORT_BASE="${PG_PORT_BASE:-27052}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/toolathlon_pg_port_leases_deepseek_rerun_c4_${UID}}"

TASK_FILE="${SCRIPT_DIR}/rerun_c4/task_order.txt"
[[ -f "$TASK_FILE" ]] || { echo "[error] missing $TASK_FILE" >&2; exit 1; }
mapfile -t TASKS < <(grep -E '^[a-z0-9][a-z0-9-]+$' "$TASK_FILE")

echo "[rerun-c4-slurm] ${#TASKS[@]} tasks"
echo "[rerun-c4-slurm] KIMI_CONFIG_ENV=$KIMI_CONFIG_ENV"
echo "[rerun-c4-slurm] MODEL_NAME=$MODEL_NAME"
echo "[rerun-c4-slurm] DUMP_ROOT=$DUMP_ROOT"
echo "[rerun-c4-slurm] PG_PORT_BASE=$PG_PORT_BASE"
echo "[rerun-c4-slurm] MAX_CONCURRENT=${MAX_CONCURRENT:-8}"

exec bash "${SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
