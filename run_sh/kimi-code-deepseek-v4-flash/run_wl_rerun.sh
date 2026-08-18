#!/bin/bash
set -euo pipefail

# WL rerun (2026-08-17): re-run the 124 API-tainted cases from the single-agent
# full run on the vLLM backend deepseek-v4-flash-wl with reasoning_effort=high.
#
# Tainted = 101 provider_invalid + 9 success-with-mid-run-401/503 + 14
# case_failed-with-401/503 (classification script scanned every slot's
# kimi-code.log for balance/auth errors).
#
# reasoning_effort=high is injected via KIMI_MODEL_THINKING_EFFORT (official
# CLI env override -> request-body reasoning_effort); the runner now passes it
# through to the container.
#
# Isolation: gnho019 (fix6 job 89185 has ended), PG base 29300 + dedicated
# lease dir, wl relay on 19318 (opencode relay 19317 left untouched).
# Results land back in the ORIGINAL full-run dump root as new slots.

SCRIPT_DIR=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/run_sh/kimi-code-deepseek-v4-flash
PROJECT_ROOT=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym

# ── dedicated relay for the vLLM backend (idempotent) ────────────────────────
# RELAY_API_KEY + the vLLM backend host are supplied via the gitignored
# config.wl.env; keep real values out of version control.
PYTHON_BIN="${PYTHON_BIN:-/storage/lintaoLab/bowending/miniconda3/envs/dbw_dev/bin/python3}"
source "${SCRIPT_DIR}/config.wl.env"
if ! curl -sS -m 5 -o /dev/null -H "Authorization: Bearer ${MODEL_API_KEY}" \
     http://127.0.0.1:19318/v1/models 2>/dev/null; then
  nohup "$PYTHON_BIN" "$SCRIPT_DIR/api_relay.py" 19318 "${VLLM_BACKEND_HOST}" 8317 \
    >/dev/shm/api_relay_wl.log 2>&1 &
  sleep 1
fi
curl -sS -m 5 -o /dev/null -w "[wl] relay 19318 -> HTTP %{http_code}\n" \
  -H "Authorization: Bearer ${MODEL_API_KEY}" http://127.0.0.1:19318/v1/models

export KIMI_CONFIG_ENV="${SCRIPT_DIR}/config.wl.env"
export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
export RELAY_API_KEY="${RELAY_API_KEY:-$MODEL_API_KEY}"
export KIMI_SUBAGENTS=""
export KIMI_MODEL_THINKING_EFFORT="high"

export DUMP_ROOT="${DUMP_ROOT:-/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/dumps/kimi-code_deepseek-v4-flash-linslab-single-agent-full_20260817-002741}"
export RUN_ID="wl-high-$(date +%Y%m%d-%H%M%S)"
export PG_PORT_BASE="${PG_PORT_BASE:-29300}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/toolathlon_pg_port_leases_wl_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho019}"
export SLURM_MEM="${SLURM_MEM:-256G}"
export SLURM_CPUS="${SLURM_CPUS:-96}"
export MAX_CONCURRENT="${MAX_CONCURRENT:-8}"
export SLURM_TIME="${SLURM_TIME:-16:00:00}"

mapfile -t TASKS < "${SCRIPT_DIR}/wl_rerun_cases.txt"

echo "=== WL rerun reasoning_effort=high ($(date)) ==="
echo "Model:    $MODEL_NAME @ $MODEL_API_URL"
echo "Dump root: $DUMP_ROOT (run_id=$RUN_ID)"
echo "Tasks:    ${#TASKS[@]} tainted cases; node=$SLURM_NODELIST; pg_base=$PG_PORT_BASE"
echo ""

cd "$PROJECT_ROOT"
exec bash "${SCRIPT_DIR}/run_on_slurm.sh" "${TASKS[@]}"
