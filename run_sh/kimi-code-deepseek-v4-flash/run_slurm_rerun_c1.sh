#!/bin/bash
# §C.1 rerun on Slurm: 91 audit-flagged cases via deepseek-v4-flash + enroot.
#
# Wraps run_on_slurm.sh with the §C.1 task list. The inner Slurm job gets:
#   - API relay (login node 19317 -> API server)
#   - compute-node enroot on /dev/shm (1TB tmpfs)
#   - all runtime hot-staging fixes (Kimi dist, notion, canvas, terminal, harness)
#   - runner-lifecycle fixes (CASE_FAILED vs INFRA_FAILED separation)
#   - post-run reconciliation (missing summary / bad status / missing audit)
#
# Prerequisites:
#   1. Enroot image rebuilt with OS packages (poppler/qpdf/bubblewrap/fonts/
#      playwright). Run scripts/rebuild_image_after_fixes.sh once under a working
#      proxy OR (now) in direct-mirror mode:
#        bash toolathlon_gym/scripts/enroot_build_agent.sh
#   2. API relay reachable (RELAY_API_KEY exported).
#
# Usage (from toolathlon_gym/):
#   export RELAY_API_KEY=...
#   bash run_sh/kimi-code-deepseek-v4-flash/run_slurm_rerun_c1.sh
#   MAX_CONCURRENT=8 SLURM_NODELIST=gnho019 bash .../run_slurm_rerun_c1.sh
#
# Env knobs (all optional):
#   MAX_CONCURRENT   parallel workers per node      (default 4)
#   SLURM_MEM        memory per node                (default 256G for 91 cases)
#   SLURM_CPUS       cpus per node                  (default 64)
#   SLURM_TIME       wall time                      (default 08:00:00)
#   SLURM_NODELIST   fixed compute node             (default auto)
#   DUMP_ROOT        override dump dir              (default auto-timestamped)

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

TASK_FILE="${SCRIPT_DIR}/rerun_c1/task_order.txt"
[[ -f "$TASK_FILE" ]] || { echo "[error] missing $TASK_FILE" >&2; exit 1; }

mapfile -t TASKS < <(grep -E '^[a-z0-9][a-z0-9-]+$' "$TASK_FILE")
echo "[rerun-c1/slurm] dispatching ${#TASKS[@]} tasks to Slurm"

# 91 cases @ ~10-15 min each / MAX_CONCURRENT slots. Default generous resources.
export SLURM_MEM="${SLURM_MEM:-256G}"
export SLURM_CPUS="${SLURM_CPUS:-64}"
export SLURM_TIME="${SLURM_TIME:-08:00:00}"
export MAX_CONCURRENT="${MAX_CONCURRENT:-6}"
export DUMP_ROOT="${DUMP_ROOT:-/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/dumps/kimi-code_deepseek-v4-flash-rerun-c1_$(date +%Y%m%d-%H%M%S)}"
export AUTO_AUDIT_HTML="${AUTO_AUDIT_HTML:-1}"

echo "[rerun-c1/slurm] DUMP_ROOT=$DUMP_ROOT"
echo "[rerun-c1/slurm] MAX_CONCURRENT=$MAX_CONCURRENT  mem=$SLURM_MEM  cpus=$SLURM_CPUS  time=$SLURM_TIME"

# Delegate to the shared Slurm launcher (handles relay, enroot, reconciliation).
exec bash "${SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
