#!/bin/bash
set -euo pipefail

# Auto-subagent experiment (2026-08-17): same 167 tasks as the single-agent
# runs (fix6's 147 rerun cases + the 20 that succeeded on 08-16), but with the
# DEFAULT sub-agent set enabled (coder/explore/plan) by NOT setting
# KIMI_SUBAGENTS — kimi_harness falls back to DEFAULT_SUBAGENTS when the var
# is unset. This is the "subagent" arm of the single-vs-subagent comparison.
#
# Model alias deepseek-v4-flash (config.env, non-linslab quota), same relay
# URL/key as the other runs.
#
# Runs on gnho014 while job 89185 occupies gnho019. Isolation analysis:
#   - PG binds 127.0.0.1 per node; /dev/shm/toolathlon_pg_<UID> is node-local
#     tmpfs, so the two experiments cannot see each other's databases or
#     sockets. Port numbers can even overlap without cross-talk.
#   - Dedicated PG port base 29200 + dedicated lease dir anyway, so the two
#     runs don't collide even if they later land on the same node.
#   - Mock-port locks (/dev/shm/toolathlon_mock_ports_<UID>) are also
#     node-local; stale locks from 20260813 were cleaned before launch.
#   - Separate DUMP_ROOT so results don't mix into the full-run root.

SCRIPT_DIR=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/run_sh/kimi-code-deepseek-v4-flash
PROJECT_ROOT=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym

export KIMI_CONFIG_ENV="${SCRIPT_DIR}/config.env"
source "$KIMI_CONFIG_ENV"
export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
export RELAY_API_KEY="${RELAY_API_KEY:-$MODEL_API_KEY}"

# CRITICAL: do NOT set KIMI_SUBAGENTS — unset means default coder/explore/plan.
unset KIMI_SUBAGENTS || true

export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_deepseek-v4-flash-subagent}"
export RUN_ID="subagent-$(date +%Y%m%d-%H%M%S)"
export PG_PORT_BASE="${PG_PORT_BASE:-29200}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/toolathlon_pg_port_leases_subagent_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho014}"
export SLURM_MEM="${SLURM_MEM:-256G}"
export SLURM_CPUS="${SLURM_CPUS:-96}"
export MAX_CONCURRENT="${MAX_CONCURRENT:-8}"
export SLURM_TIME="${SLURM_TIME:-16:00:00}"

mapfile -t TASKS < "${SCRIPT_DIR}/subagent_cases.txt"

echo "=== Auto-subagent experiment ($(date)) ==="
echo "Model:    $MODEL_NAME (default subagents: coder/explore/plan)"
echo "Dump root: $DUMP_ROOT (run_id=$RUN_ID)"
echo "Tasks:    ${#TASKS[@]} cases; node=$SLURM_NODELIST; pg_base=$PG_PORT_BASE; MAX_CONCURRENT=$MAX_CONCURRENT"
echo ""

cd "$PROJECT_ROOT"
exec bash "${SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
