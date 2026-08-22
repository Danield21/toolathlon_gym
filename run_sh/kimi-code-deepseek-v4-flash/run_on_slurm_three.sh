#!/bin/bash
# Launch the classic 3-subagent roster through the shared Slurm runner.
#
# Roster: plan + coder + explore (KIMI_SUBAGENTS=three / 3).
# Mirrors run_on_slurm_fine_grained.sh, which pins ten; this one pins three.
#
# Usage is identical to run_on_slurm.sh:
#   bash run_on_slurm_three.sh <task> [<task> ...]
#   bash run_on_slurm_three.sh --smoke

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BASE_LAUNCHER="${SCRIPT_DIR}/run_on_slurm.sh"

if [[ ! -f "$BASE_LAUNCHER" ]]; then
  echo "[three] FATAL: shared launcher not found: $BASE_LAUNCHER" >&2
  exit 1
fi

export KIMI_SUBAGENTS=three

echo "[three] roster=three (plan/coder/explore) plan_first=${KIMI_PLAN_FIRST:-0}"
exec bash "$BASE_LAUNCHER" "$@"
