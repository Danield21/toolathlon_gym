#!/bin/bash
# Slurm wrapper for §C.2 rerun: 77 non-PASS cases, deepseek-v4-flash-linslab.
#
# This wraps run_on_slurm.sh (which handles the API relay, slurm allocation,
# enroot setup, and reconciliation) but overrides:
#   - KIMI_CONFIG_ENV -> config.c2.env (model_name=deepseek-v4-flash-linslab)
#   - DUMP_ROOT       -> dedicated rerun-c2 dump root
#   - PG_PORT_BASE    -> 26832 (dedicated range, avoids collision)
#
# Usage (from toolathlon_gym/):
#   RELAY_API_KEY=sk-remote-... bash run_sh/kimi-code-deepseek-v4-flash/run_slurm_rerun_c2.sh
#   MAX_CONCURRENT=6 RELAY_API_KEY=sk-remote-... bash .../run_slurm_rerun_c2.sh
#
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"

export KIMI_CONFIG_ENV="${SCRIPT_DIR}/config.c2.env"

# Source the C2 config HERE (on the login node, before srun dispatches) so that
# MODEL_NAME / MODEL_API_URL / MODEL_API_KEY are explicitly exported into the
# environment that srun inherits. Relying solely on KIMI_CONFIG_ENV being
# propagated through the srun → INNER → run_eval_parallel.sh chain proved
# fragile: when KIMI_CONFIG_ENV was dropped, run_eval_parallel.sh fell back to
# the default config.env (MODEL_NAME=deepseek-v4-flash, whose quota is
# exhausted), causing every case to 503 auth_unavailable. Exporting the
# resolved values here is belt-and-suspenders.
# shellcheck disable=SC1090
source "$KIMI_CONFIG_ENV"
export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S

export DUMP_ROOT="${PROJECT_ROOT}/dumps/kimi-code_deepseek-v4-flash-linslab-rerun-c2_$(date +%Y%m%d-%H%M%S)"
export PG_PORT_BASE="${PG_PORT_BASE:-26832}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/toolathlon_pg_port_leases_deepseek_rerun_c2_${UID}}"

TASK_FILE="${SCRIPT_DIR}/rerun_c2/task_order.txt"
[[ -f "$TASK_FILE" ]] || { echo "[error] missing $TASK_FILE" >&2; exit 1; }
mapfile -t TASKS < <(grep -E '^[a-z0-9][a-z0-9-]+$' "$TASK_FILE")

echo "[rerun-c2-slurm] ${#TASKS[@]} tasks"
echo "[rerun-c2-slurm] KIMI_CONFIG_ENV=$KIMI_CONFIG_ENV"
echo "[rerun-c2-slurm] DUMP_ROOT=$DUMP_ROOT"
echo "[rerun-c2-slurm] PG_PORT_BASE=$PG_PORT_BASE"

exec bash "${SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
