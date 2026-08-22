#!/bin/bash
# Launch the fine-grained 10-subagent roster through the shared Slurm runner.
#
# Roster: 1 plan agent + 7 domain specialists + evidence-integrator +
# deliverable-auditor (9 specialized/cross-cutting agents + plan).
# KIMI_PLAN_FIRST remains caller-controlled: the plan agent is available in
# every run, but is mandatory only when KIMI_PLAN_FIRST=1 is explicitly set.
#
# Usage is identical to run_on_slurm.sh:
#   bash run_on_slurm_fine_grained.sh <task> [<task> ...]
#   KIMI_PLAN_FIRST=1 bash run_on_slurm_fine_grained.sh --smoke

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BASE_LAUNCHER="${SCRIPT_DIR}/run_on_slurm.sh"

if [[ ! -f "$BASE_LAUNCHER" ]]; then
  echo "[fine-grained] FATAL: shared launcher not found: $BASE_LAUNCHER" >&2
  exit 1
fi

# Pin the new roster even if the caller inherited KIMI_SUBAGENTS=three or an
# empty value from a previous experiment.
export KIMI_SUBAGENTS=ten

echo "[fine-grained] roster=ten (plan + 7 domain + integrator + auditor) plan_first=${KIMI_PLAN_FIRST:-0}"
exec bash "$BASE_LAUNCHER" "$@"
