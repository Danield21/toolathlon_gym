#!/bin/bash
# Launch the 28-task bugfix regression batch for deepseek-v4-flash.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
TASKS_FILE="${SCRIPT_DIR}/tasks_bugfix_regression_28.txt"

cd "$PROJECT_ROOT"

set -a
# shellcheck disable=SC1091
source "${SCRIPT_DIR}/config.env"
set +a

export RELAY_API_KEY="${RELAY_API_KEY:-${MODEL_API_KEY:?MODEL_API_KEY missing}}"
export MAX_CONCURRENT="${MAX_CONCURRENT:-8}"
export SLURM_MEM="${SLURM_MEM:-256G}"
export SLURM_CPUS="${SLURM_CPUS:-96}"
export SLURM_TIME="${SLURM_TIME:-08:00:00}"

mapfile -t TASKS < <(grep -vE '^\s*(#|$)' "$TASKS_FILE" | tr -d '\r')
if [[ ${#TASKS[@]} -eq 0 ]]; then
  echo "[run_bugfix_regression_28] ERROR: no tasks in $TASKS_FILE" >&2
  exit 1
fi

echo "[run_bugfix_regression_28] tasks=${#TASKS[@]} MAX_CONCURRENT=${MAX_CONCURRENT} SLURM_MEM=${SLURM_MEM} SLURM_CPUS=${SLURM_CPUS} SLURM_TIME=${SLURM_TIME}"
exec bash "${SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
