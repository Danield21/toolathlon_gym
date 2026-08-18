#!/bin/bash
# §C.2 rerun: 77 non-PASS cases from §C.1, deepseek-v4-flash-linslab + slurm + enroot.
#
# This is the post-fix rerun for the 91-case §C.1 batch. It runs ONLY the 77
# cases that were NOT success in the §C.1 run
# (the 14 PASS cases are skipped). The model_name is switched to
# deepseek-v4-flash-linslab (same URL/api_key); all audit fixes from the
# 2026-08-12 case study are applied via runtime hot-staging + the rebuilt image.
#
# Usage (from toolathlon_gym/):
#   bash run_sh/kimi-code-deepseek-v4-flash/run_rerun_c2.sh
#   MAX_CONCURRENT=8 bash run_sh/kimi-code-deepseek-v4-flash/run_rerun_c2.sh
#
# To run under Slurm, wrap with run_slurm_rerun_c2.sh.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"

# §C.2 task list (77 non-PASS cases from §C.1 summary CSV)
TASK_FILE="${SCRIPT_DIR}/rerun_c2/task_order.txt"
[[ -f "$TASK_FILE" ]] || { echo "[error] missing $TASK_FILE" >&2; exit 1; }

mapfile -t TASKS < <(grep -E '^[a-z0-9][a-z0-9-]+$' "$TASK_FILE")
echo "[rerun-c2] loaded ${#TASKS[@]} tasks from $TASK_FILE"

# Dedicated dump root + port range for this rerun batch (avoids collision with
# §C.1 and any parallel MiniMax batch).
export KIMI_CONFIG_ENV="${SCRIPT_DIR}/config.c2.env"
export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_deepseek-v4-flash-linslab-rerun-c2_$(date +%Y%m%d-%H%M%S)}"
export PG_PORT_BASE="${PG_PORT_BASE:-26832}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/toolathlon_pg_port_leases_deepseek_rerun_c2_${UID}}"

# tmpfs for enroot (same as the main launcher)
export ENROOT_DATA_PATH="${ENROOT_DATA_PATH:-/dev/shm/enroot_data}"
export ENROOT_TEMP_PATH="${ENROOT_TEMP_PATH:-/dev/shm/enroot_tmp}"
export ENROOT_RUNTIME_PATH="${ENROOT_RUNTIME_PATH:-/dev/shm/enroot_runtime}"
export ENROOT_CACHE_PATH="${ENROOT_CACHE_PATH:-/dev/shm/enroot_cache}"
mkdir -p "$ENROOT_DATA_PATH" "$ENROOT_TEMP_PATH" "$ENROOT_RUNTIME_PATH" "$ENROOT_CACHE_PATH"

export PATH="/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin:${PATH}"

echo "[rerun-c2] KIMI_CONFIG_ENV=$KIMI_CONFIG_ENV"
echo "[rerun-c2] DUMP_ROOT=$DUMP_ROOT"
echo "[rerun-c2] PG_PORT_BASE=$PG_PORT_BASE"
echo "[rerun-c2] MAX_CONCURRENT=${MAX_CONCURRENT:-4}"

# Delegate to the shared parallel launcher (which has all the hot-staging,
# runner-lifecycle, and credential-isolation fixes).
exec bash "${SCRIPT_DIR}/run_eval_parallel.sh" "${TASKS[@]}"
