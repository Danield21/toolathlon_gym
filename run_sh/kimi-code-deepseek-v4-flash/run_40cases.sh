#!/bin/bash
# Launch 40-case deepseek-v4-flash eval on slurm (T1/T2/T3/T4 x10 each).
#
# This wrapper reads tasks_40.txt and dispatches them to run_on_slurm.sh with
# MAX_CONCURRENT=8 (the launcher's hard cap). Run AFTER the agent image has
# been rebuilt and the NFS rootfs snapshot refreshed.
#
# Prereqs:
#   1. enroot_build_agent.sh finished (new toolathlon-pack.sqsh + shm rootfs)
#   2. NFS rootfs refreshed:   sync_new_rootfs_to_nfs.sh (or manual rsync)
#   3. api_relay.py running on login node (run_on_slurm starts it idempotently)
#
# Usage (from toolathlon_gym/):
#   bash run_sh/kimi-code-deepseek-v4-flash/run_40cases.sh
#
# Env knobs:
#   MAX_CONCURRENT  (default 8, hard-cap 8 in launcher)
#   SLURM_TIME      (default 06:00:00 — 40 cases / 8 concurrency ~ 5 rounds)

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
TASKS_FILE="${SCRIPT_DIR}/tasks_40.txt"

cd "$PROJECT_ROOT"

# Read non-empty, non-comment lines as the task list.
mapfile -t TASKS < <(grep -vE '^\s*(#|$)' "$TASKS_FILE" | tr -d '\r')
if [[ ${#TASKS[@]} -eq 0 ]]; then
  echo "[run_40cases] ERROR: no tasks in $TASKS_FILE" >&2
  exit 1
fi
echo "[run_40cases] ${#TASKS[@]} tasks from $TASKS_FILE"

export MAX_CONCURRENT="${MAX_CONCURRENT:-8}"
# 40 cases at concurrency 8 ~ 5 rounds; budget generous wall time.
export SLURM_TIME="${SLURM_TIME:-06:00:00}"
export SLURM_MEM="${SLURM_MEM:-128G}"
export SLURM_CPUS="${SLURM_CPUS:-64}"

echo "[run_40cases] MAX_CONCURRENT=$MAX_CONCURRENT  SLURM_TIME=$SLURM_TIME"
echo "[run_40cases] dispatching to slurm..."

# run_on_slurm.sh handles: api_relay start, env setup, srun, exec to launcher.
exec bash "${SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
