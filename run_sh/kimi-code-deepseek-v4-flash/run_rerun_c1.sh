#!/bin/bash
# §C.1 rerun: 91 audit-flagged cases, deepseek-v4-flash + slurm + enroot.
#
# This wraps run_eval_parallel.sh with the §C.1 task list, a dedicated dump
# directory, and a separate port range so it can run alongside other batches.
#
# All audit fixes are applied automatically via runtime hot-staging in
# run_eval_parallel.sh (Kimi dist, notion cli.mjs, canvas build, terminal.yaml,
# kimi_harness, tasks/finalpool). The only image-level requirement is the OS
# packages (poppler/qpdf/bubblewrap/fonts/playwright) from the rebuilt image.
#
# Usage (from toolathlon_gym/):
#   bash run_sh/kimi-code-deepseek-v4-flash/run_rerun_c1.sh
#   MAX_CONCURRENT=8 bash run_sh/kimi-code-deepseek-v4-flash/run_rerun_c1.sh
#
# To run under Slurm, wrap with sbatch (see run_slurm_rerun_c1.sh).

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"

# §C.1 task list (91 cases from the audit handoff doc)
TASK_FILE="${SCRIPT_DIR}/rerun_c1/task_order.txt"
[[ -f "$TASK_FILE" ]] || { echo "[error] missing $TASK_FILE" >&2; exit 1; }

mapfile -t TASKS < <(grep -E '^[a-z0-9][a-z0-9-]+$' "$TASK_FILE")
echo "[rerun-c1] loaded ${#TASKS[@]} tasks from $TASK_FILE"

# Dedicated dump root + port range for this rerun batch (avoids collision with
# the single-agent smoke run and any parallel MiniMax batch).
export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_deepseek-v4-flash-rerun-c1_$(date +%Y%m%d-%H%M%S)}"
export PG_PORT_BASE="${PG_PORT_BASE:-26632}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/toolathlon_pg_port_leases_deepseek_rerun_c1_${UID}}"

# tmpfs for enroot (same as the main launcher)
export ENROOT_DATA_PATH="${ENROOT_DATA_PATH:-/dev/shm/enroot_data}"
export ENROOT_TEMP_PATH="${ENROOT_TEMP_PATH:-/dev/shm/enroot_tmp}"
export ENROOT_RUNTIME_PATH="${ENROOT_RUNTIME_PATH:-/dev/shm/enroot_runtime}"
export ENROOT_CACHE_PATH="${ENROOT_CACHE_PATH:-/dev/shm/enroot_cache}"
mkdir -p "$ENROOT_DATA_PATH" "$ENROOT_TEMP_PATH" "$ENROOT_RUNTIME_PATH" "$ENROOT_CACHE_PATH"

export PATH="/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin:${PATH}"

echo "[rerun-c1] DUMP_ROOT=$DUMP_ROOT"
echo "[rerun-c1] PG_PORT_BASE=$PG_PORT_BASE"
echo "[rerun-c1] MAX_CONCURRENT=${MAX_CONCURRENT:-4}"

# Delegate to the shared parallel launcher (which has all the hot-staging,
# runner-lifecycle, and credential-isolation fixes).
exec bash "${SCRIPT_DIR}/run_eval_parallel.sh" "${TASKS[@]}"
