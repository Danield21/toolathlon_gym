#!/bin/bash
set -euo pipefail

# Rerun fix6 (2026-08-17): 147 cases = 7 user-named (conference-prep,
# tracker-notion, at-risk, grade-equity, late-submission, wc-competitor,
# yt-fireship) + all 141 provider_invalid victims of the 02:16 balance event.
#
# Results land back in the ORIGINAL full-run dump root so each task gains a new
# <RUN_ID>_slot<N> dir; run_on_slurm.sh then regenerates audit.html per case
# and rebuilds audit_index.html/json over ALL slots in that root.
#
# Isolation: gnho019 (idle, 2TB), PG base 29100, dedicated lease dir. Nothing
# else is running right now (squeue empty).

SCRIPT_DIR=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/run_sh/kimi-code-deepseek-v4-flash
PROJECT_ROOT=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym

export KIMI_CONFIG_ENV="${SCRIPT_DIR}/config.c2.env"
source "$KIMI_CONFIG_ENV"
export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
export RELAY_API_KEY="${RELAY_API_KEY:-$MODEL_API_KEY}"
export KIMI_SUBAGENTS=""

# Land results inside the existing full-run dump root.
export DUMP_ROOT="${DUMP_ROOT:-/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/dumps/kimi-code_deepseek-v4-flash-linslab-single-agent-full_20260817-002741}"
export RUN_ID="rerun-fix6-$(date +%Y%m%d-%H%M%S)"
export PG_PORT_BASE="${PG_PORT_BASE:-29100}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/toolathlon_pg_port_leases_fix6_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho019}"
export SLURM_MEM="${SLURM_MEM:-256G}"
export SLURM_CPUS="${SLURM_CPUS:-96}"
# 147 cases: high concurrency. Runner caps at 8; API + NFS are the bottleneck.
export MAX_CONCURRENT="${MAX_CONCURRENT:-8}"
# Long tail: give the allocation room for ~19 waves x ~30min.
export SLURM_TIME="${SLURM_TIME:-16:00:00}"

mapfile -t TASKS < "${SCRIPT_DIR}/rerun_fix6_cases.txt"

echo "=== Rerun fix6 ($(date)) ==="
echo "Dump root: $DUMP_ROOT (results appended as ${RUN_ID}_slot<N>)"
echo "Tasks: ${#TASKS[@]} cases; node=$SLURM_NODELIST; pg_base=$PG_PORT_BASE; MAX_CONCURRENT=$MAX_CONCURRENT"
echo ""

cd "$PROJECT_ROOT"
exec bash "${SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
