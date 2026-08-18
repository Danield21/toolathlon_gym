#!/bin/bash
set -euo pipefail

# Plan-first arm (2026-08-17): tests the hypothesis that mandatory planning
# before execution improves sub-agent dispatch efficiency.
#
#   Arm design:
#   - Same base as the subagent arm (default sub-agents coder/explore/plan,
#     KIMI_SUBAGENTS unset).
#   - KIMI_PLAN_FIRST=1 injects the plan-first mandate section into the main
#     agent prompt (assets/sections/plan_first_default.md, loaded by
#     kimi_main.py::_plan_first_section). One env var — that's the whole switch.
#   - Compliance is MONITORED, not enforced: audit_html_gen.py computes
#     plan_first_ok / pre_plan_actions / n_plan_dispatches per case and shows
#     them in audit.html + audit_index.html. Violating runs stay in the stats.
#
#   Comparison arm: dumps/kimi-code_deepseek-v4-flash-subagent (same model,
#     same pool, same effort; differs only by the plan-first prompt section).
#
#   Isolation: gnho014 (wl rerun occupies gnho019 via job 89317, PG base 29300),
#   own PG base 29400 + dedicated lease dir, separate DUMP_ROOT.

SCRIPT_DIR=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/run_sh/kimi-code-deepseek-v4-flash
PROJECT_ROOT=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym

export KIMI_CONFIG_ENV="${SCRIPT_DIR}/config.env"
source "$KIMI_CONFIG_ENV"
export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
export RELAY_API_KEY="${RELAY_API_KEY:-$MODEL_API_KEY}"

# Default sub-agents (coder/explore/plan) — do NOT set KIMI_SUBAGENTS.
unset KIMI_SUBAGENTS || true

# ── THE plan-first switch ────────────────────────────────────────────────────
export KIMI_PLAN_FIRST="${KIMI_PLAN_FIRST:-1}"

# reasoning_effort=high for this arm (request-body reasoning_effort via the
# CLI's KIMI_MODEL_THINKING_EFFORT override; runner passes it into the container).
export KIMI_MODEL_THINKING_EFFORT="${KIMI_MODEL_THINKING_EFFORT:-high}"

export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_deepseek-v4-flash-plan-first}"
export RUN_ID="planfirst-$(date +%Y%m%d-%H%M%S)"
export PG_PORT_BASE="${PG_PORT_BASE:-29400}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/toolathlon_pg_port_leases_planfirst_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho014}"
export SLURM_MEM="${SLURM_MEM:-256G}"
export SLURM_CPUS="${SLURM_CPUS:-96}"
export MAX_CONCURRENT="${MAX_CONCURRENT:-8}"
export SLURM_TIME="${SLURM_TIME:-16:00:00}"

mapfile -t TASKS < "${TASK_LIST:-${SCRIPT_DIR}/plan_first_cases.txt}"

echo "=== Plan-first arm ($(date)) ==="
echo "Model:    $MODEL_NAME | KIMI_PLAN_FIRST=$KIMI_PLAN_FIRST | effort=${KIMI_MODEL_THINKING_EFFORT:-default}"
echo "Dump root: $DUMP_ROOT (run_id=$RUN_ID)"
echo "Tasks:    ${#TASKS[@]} cases (list=${TASK_LIST:-default 15-case pilot}); node=$SLURM_NODELIST; pg_base=$PG_PORT_BASE"
echo ""

cd "$PROJECT_ROOT"
exec bash "${SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
