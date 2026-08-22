#!/bin/bash
# Launch deepseek-v4-flash eval on a slurm compute node.
#
# Why slurm: the login node's user slice has a 12GB cgroup memory limit, and a
# single kimi-code task maps ~22GB virtual memory + 4.7GB enroot rootfs shmem.
# Slurm compute nodes (gnho019 etc.) have 2TB RAM and no such cgroup cap, so
# high-concurrency evals run cleanly there.
#
# Compute nodes have NO external network. We bridge API calls by running a TCP
# relay on the login node (api_relay.py) that forwards login:19317 -> the API
# server. Compute nodes reach the login node over the internal 192.168.x net.
#
# Usage:
#   bash run_sh/kimi-code-deepseek-v4-flash/run_on_slurm.sh <task> [<task> ...]
#   MAX_CONCURRENT=6 bash .../run_on_slurm.sh task1 task2
#   bash .../run_on_slurm.sh --smoke   # single quick smoke test (canvas)
#
# Env knobs:
#   SLURM_MEM     memory per node to request      (default 128G)
#   SLURM_CPUS    cpus per node                   (default 64)
#   SLURM_TIME    wall time                       (default 04:00:00)
#   SLURM_PARTITION  partition                    (default linlab)
#   SLURM_NODELIST   optional fixed node, e.g. gnho019
#   RELAY_API_KEY  API key used only for relay health checks (required)

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"

# ── login-node API relay (idempotent) ────────────────────────────────────────
LOGIN_IP="192.168.180.240"
RELAY_PORT="${RELAY_PORT:-19317}"
RELAY_SKIP_AUTOSTART="${RELAY_SKIP_AUTOSTART:-0}"
PYTHON_BIN="${PYTHON_BIN:-/storage/lintaoLab/bowending/miniconda3/envs/dbw_dev/bin/python3}"
RELAY_API_KEY="${RELAY_API_KEY:?Set RELAY_API_KEY before launching the eval}"

RELAY_OWNED=0
RELAY_PID=""
cleanup_launcher() {
  local rc=$?
  trap - EXIT
  if [[ "$RELAY_OWNED" == "1" && -n "$RELAY_PID" ]]; then
    kill "$RELAY_PID" >/dev/null 2>&1 || true
    wait "$RELAY_PID" 2>/dev/null || true
  fi
  exit "$rc"
}
trap cleanup_launcher EXIT

if curl -sS -m 5 -o /dev/null \
     -H "Authorization: Bearer ${RELAY_API_KEY}" \
     "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
  echo "[slurm-launch] API relay healthy on 127.0.0.1:${RELAY_PORT}"
elif [[ "$RELAY_SKIP_AUTOSTART" == "1" ]]; then
  echo "[slurm-launch] FATAL: relay on 127.0.0.1:${RELAY_PORT} is down and RELAY_SKIP_AUTOSTART=1" >&2
  exit 1
else
  # Start relay on the login node (where this script runs). If the port was
  # already healthy above, it is shared and we leave it alone on exit. If this
  # launcher starts it, cleanup_launcher stops it when the Slurm run finishes.
  nohup "$PYTHON_BIN" "$SCRIPT_DIR/api_relay.py" >/dev/shm/api_relay.log 2>&1 &
  RELAY_PID=$!
  RELAY_OWNED=1
  sleep 1
  if curl -sS -m 5 -o /dev/null \
       -H "Authorization: Bearer ${RELAY_API_KEY}" \
       "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
    echo "[slurm-launch] API relay started on 127.0.0.1:${RELAY_PORT} pid=${RELAY_PID}"
  else
    echo "[slurm-launch] WARNING: relay health check failed, check /dev/shm/api_relay.log"
  fi
fi

# ── slurm resource request ───────────────────────────────────────────────────
SLURM_MEM="${SLURM_MEM:-128G}"
SLURM_CPUS="${SLURM_CPUS:-64}"
SLURM_TIME="${SLURM_TIME:-04:00:00}"
SLURM_PARTITION="${SLURM_PARTITION:-linlab}"
SLURM_NODELIST="${SLURM_NODELIST:-}"
SRUN_NODE_ARGS=()
if [[ -n "$SLURM_NODELIST" ]]; then
  SRUN_NODE_ARGS=(--nodelist="$SLURM_NODELIST")
fi

# ── enroot on compute nodes lives under the user's .local ───────────────────
ENROOT_ROOT="/storage/lintaoLab/bowending/.local/enroot"
ENROOT_BIN="/storage/lintaoLab/bowending/.local/bin/enroot"

# ── smoke mode: single fast task ─────────────────────────────────────────────
TASKS=("$@")
if [[ "${1:-}" == "--smoke" ]]; then
  TASKS=(canvas-announcement-summary)
  SLURM_MEM="32G"
  SLURM_CPUS="8"
  SLURM_TIME="01:00:00"
fi
if [[ ${#TASKS[@]} -eq 0 ]]; then
  echo "Usage: $0 <task> [<task> ...]   (or --smoke)"
  exit 1
fi

echo "[slurm-launch] partition=$SLURM_PARTITION node=${SLURM_NODELIST:-auto} mem=$SLURM_MEM cpus=$SLURM_CPUS time=$SLURM_TIME"
echo "[slurm-launch] tasks: ${TASKS[*]}"

# ── the command to run inside the slurm job ──────────────────────────────────
# Key points:
#   1. no_proxy must include the login node IP so kimi-code's HTTP client
#      doesn't try the (nonexistent) 127.0.0.1:7890 proxy for API calls.
#   2. ENROOT paths point to the compute node's /dev/shm (1TB tmpfs).
#   3. PATH must include enroot bin + sbin (unsquashfs).
read -r -d '' INNER <<'INNER_EOF' || true
set -euo pipefail
cd /lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym

# MODEL_NAME / KIMI_CONFIG_ENV are passed as $1 / $2 from the srun invocation
# (see the bash -c "$INNER" _ "$MODEL_NAME" "$KIMI_CONFIG_ENV" ... below). srun
# on this cluster does NOT reliably propagate KIMI_CONFIG_ENV through the
# environment (observed: it arrived as config.env instead of config.c2.env),
# which made run_eval_parallel.sh fall back to the default config.env
# (MODEL_NAME=deepseek-v4-flash, quota exhausted -> 503). Passing them as
# positional args is the robust fix.
export MODEL_NAME="$1"
export KIMI_CONFIG_ENV="$2"
export MODEL_API_URL="$3"
export MODEL_API_KEY="$4"
export KIMI_MAX_CONTEXT="$5"
export KIMI_TASK_TIMEOUT_S="$6"
export RELAY_PORT="$7"
export KIMI_MODEL_THINKING_EFFORT="$8"
# $9 = KIMI_SUBAGENTS with a sentinel: "__UNSET__" means leave it unset (harness
# default: coder/explore/plan); "" explicitly disables all sub-agents; "ten"
# selects the 10-agent roster; comma lists pass through verbatim.
if [ "$9" != "__UNSET__" ]; then
  export KIMI_SUBAGENTS="$9"
fi
shift 9
echo "[inner] ARGV: MODEL_NAME=$MODEL_NAME KIMI_CONFIG_ENV=$KIMI_CONFIG_ENV RELAY_PORT=$RELAY_PORT effort=${KIMI_MODEL_THINKING_EFFORT:-} subagents=${KIMI_SUBAGENTS:-<default 3>}"

# Compute nodes inherit http_proxy=127.0.0.1:7890 (nonexistent). Clear all
# proxy vars and whitelist the login node so API calls go direct to the relay.
unset http_proxy https_proxy HTTP_PROXY HTTPS_PROXY ALL_PROXY all_proxy
export no_proxy="127.0.0.1,localhost,::1,192.168.180.240,104.168.43.47,172.16.0.0/12,10.0.0.0/8"

# enroot runtime on this cluster.
# IMPORTANT: /storage/lintaoLab is an NFS mount. Compute nodes reading the
# enroot bash script + libraries directly from NFS intermittently hang on
# read() (NFS dithering), which deadlocks even `enroot version`. To avoid
# this, mirror the (tiny, ~40 MB) enroot install into node-local /dev/shm
# and put that copy first on PATH.
ENROOT_SRC="/storage/lintaoLab/bowending/.local/enroot"
ENROOT_LOCAL="/dev/shm/enroot_install"
if [[ ! -x "$ENROOT_LOCAL/bin/enroot" ]]; then
  mkdir -p "$ENROOT_LOCAL"
  rsync -a "$ENROOT_SRC/" "$ENROOT_LOCAL/" 2>/dev/null || cp -a "$ENROOT_SRC/." "$ENROOT_LOCAL/"
fi
export ENROOT_LIBRARY_PATH="${ENROOT_LOCAL}/lib"
export ENROOT_SYSCONF_PATH="${ENROOT_LOCAL}/etc"
export PATH="${ENROOT_LOCAL}/bin:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin:/storage/lintaoLab/bowending/miniconda3/envs/toolathlon_gym/bin"

# Use /dev/shm (1TB tmpfs on compute nodes) for all enroot working dirs.
export ENROOT_DATA_PATH="/dev/shm/enroot_data"
export ENROOT_TEMP_PATH="/dev/shm/enroot_tmp"
export ENROOT_RUNTIME_PATH="/dev/shm/enroot_runtime"
export ENROOT_CACHE_PATH="/dev/shm/enroot_cache"
mkdir -p "$ENROOT_DATA_PATH" "$ENROOT_TEMP_PATH" "$ENROOT_RUNTIME_PATH" "$ENROOT_CACHE_PATH"
export RUN_ID="${RUN_ID:-$(date +%Y%m%d-%H%M%S)}"

cleanup_inner() {
  local rc=$?
  trap - EXIT
  set +e
  echo "[inner-cleanup] run_id=${RUN_ID} rc=${rc}"
  if [[ -d "$ENROOT_DATA_PATH" ]]; then
    find "$ENROOT_DATA_PATH" -mindepth 1 -maxdepth 1 -type d \
      -name "agent-${RUN_ID}-*" -exec rm -rf -- {} + 2>/dev/null || true
    rmdir "$ENROOT_DATA_PATH" 2>/dev/null || true
  fi
  if [[ -d "/dev/shm/toolathlon_pg_${UID}" ]]; then
    find "/dev/shm/toolathlon_pg_${UID}" -mindepth 1 -maxdepth 1 -type d \
      -name "${RUN_ID}_*" -exec rm -rf -- {} + 2>/dev/null || true
    rmdir "/dev/shm/toolathlon_pg_${UID}" 2>/dev/null || true
  fi
  rmdir "/dev/shm/toolathlon_pg_port_leases_deepseek_${UID}" 2>/dev/null || true
  rmdir "$ENROOT_TEMP_PATH" "$ENROOT_RUNTIME_PATH" "$ENROOT_CACHE_PATH" 2>/dev/null || true
  if [[ "${SLURM_SELF_CANCEL_ON_EXIT:-1}" == "1" && -n "${SLURM_JOB_ID:-}" ]]; then
    echo "[inner-cleanup] self-scancel Slurm job ${SLURM_JOB_ID}"
    scancel "$SLURM_JOB_ID" >/dev/null 2>&1 || true
  fi
  exit "$rc"
}
trap cleanup_inner EXIT

# The base rootfs (toolathlon-pack) lives on NFS so every compute node can see
# it. Each worker enroot-creates its own copy into /dev/shm/enroot_data/agent-<task>.
# On compute nodes /dev/shm is 1TB so 6x4.7GB is trivial.
#
# c4 case-study (2026-08-15, dev_docs/2026-08-15-c4-rerun-batch-analysis.md §1):
# this line used to point at toolathlon-pack-rootfs — a WRITABLE copy that had
# been pip/uv-edited on the host, so its venv entry scripts carried host paths
# and 42/77 cases failed the MCP health check with infra_failed. The only
# artifact allowed here is the sqsh produced by enroot_build_agent.sh (immutable
# by construction). run_eval_parallel.sh's AGENT_TEMPLATE default prefers a
# directory copy, so unset it and export the sqsh path explicitly: workers then
# take the `enroot create -n <name> <sqsh>` path.
unset AGENT_TEMPLATE
export AGENT_SQSH="/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers/images/toolathlon-pack.sqsh"
if [[ ! -f "$AGENT_SQSH" ]]; then
  echo "[inner] FATAL: agent image $AGENT_SQSH missing — build it with scripts/enroot_build_agent.sh" >&2
  exit 1
fi
echo "[inner] agent image: $AGENT_SQSH (sqsh, build-produced)"

# Sanity: relay reachable from compute node.
# /v1/models can wait on the upstream gateway (llmapi.blsc.cn often >5s).
# Use /healthz when the HTTPS relay serves it; otherwise /v1/models with 30s.
echo "[inner] relay check:"
_relay_ok=0
_healthz_code="$(curl -sS -m 5 -o /dev/null -w "%{http_code}" \
  "http://192.168.180.240:${RELAY_PORT}/healthz" || true)"
if [[ "${_healthz_code}" == "200" ]]; then
  echo "  healthz HTTP 200"
  _relay_ok=1
else
  if curl -sS -m 30 -o /dev/null -w "  models HTTP %{http_code}\n" \
       -H "Authorization: Bearer ${RELAY_API_KEY}" \
       "http://192.168.180.240:${RELAY_PORT}/v1/models"; then
    _relay_ok=1
  fi
fi
if [[ "${_relay_ok}" != 1 ]]; then
  echo "[inner] FATAL: cannot reach API relay from compute node (port ${RELAY_PORT})"
  exit 1
fi

echo "[inner] enroot: $(enroot version 2>&1)"
echo "[inner] /dev/shm: $(df -h /dev/shm | tail -1 | awk '{print $2" total, "$4" free"}')"

# Hand off to the existing parallel launcher. Keep this shell alive so its
# EXIT trap can release the Slurm allocation and sweep this run's tmpfs roots.
#
# Important: model/evaluator case failures are valid eval outcomes and should
# not mark the Slurm job FAILED. Reconcile the summary after the runner exits:
# missing rows, rootfs/PG failures, or missing audit.html are infra failures;
# ordinary "failed" rows are case results.
set +e
bash run_sh/kimi-code-deepseek-v4-flash/run_eval_parallel.sh "$@"
RUNNER_RC=$?
set -e
echo "[inner] runner rc=${RUNNER_RC}; generating audit/index and reconciling"

if [[ "${AUTO_AUDIT_HTML:-1}" == "1" ]]; then
  python3 scripts/audit_html_gen.py "${DUMP_ROOT:-/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/dumps/kimi-code_deepseek-v4-flash}" || true
fi

python3 - "${DUMP_ROOT:-}" "${RUN_ID:-}" "$@" <<'PY'
import csv
import sys
from pathlib import Path

dump_root = Path(sys.argv[1] or ".")
run_id = sys.argv[2]
tasks = sys.argv[3:]
summary = dump_root / f"summary_parallel_{run_id}.csv"
infra_statuses = {
    "pg_fail", "rootfs_fail", "pg_test_fail", "pg_test_success",
    "missing_summary", "audit_missing", "interrupted",
    # Fail-fast / provider-invalid classifications (case-study 2026-08-12):
    # these mean the run was infra-invalid (bad seed / model backend down /
    # MCP server missing its venv) and must NOT count against the model's
    # ability score. infra_failed = MCP pre-flight health check failed
    # (P0-2, dev_docs/2026-08-13-c2-tz-fix-design.md §2).
    "preprocess_failed", "provider_invalid", "infra_failed",
}
case_failure_statuses = {"failed", "case_failed"}
allowed_statuses = {"success"} | case_failure_statuses

if not summary.exists():
    print(f"[reconcile] INFRA: missing summary {summary}")
    sys.exit(2)

rows = []
with summary.open(newline="") as f:
    for row in csv.DictReader(f):
        rows.append(row)

by_task = {r.get("task"): r for r in rows}
missing = [t for t in tasks if t not in by_task]
bad_status = [
    (r.get("task"), r.get("status"))
    for r in rows
    if r.get("status") not in allowed_statuses
]
missing_audit = []
for r in rows:
    out = Path(r.get("output_dir") or "")
    if r.get("status") in allowed_statuses and not (out / "audit.html").exists():
        missing_audit.append(r.get("task"))

case_failed = sum(1 for r in rows if r.get("status") in case_failure_statuses)
case_success = sum(1 for r in rows if r.get("status") == "success")
print(
    f"[reconcile] rows={len(rows)} tasks={len(tasks)} "
    f"success={case_success} failed={case_failed} "
    f"missing={len(missing)} bad_status={len(bad_status)} missing_audit={len(missing_audit)}"
)
if missing:
    print("[reconcile] missing summary tasks: " + ", ".join(missing))
if bad_status:
    print("[reconcile] infra statuses: " + ", ".join(f"{t}:{s}" for t, s in bad_status))
if missing_audit:
    print("[reconcile] missing audit tasks: " + ", ".join(missing_audit))

if missing or bad_status or missing_audit:
    sys.exit(2)
sys.exit(0)
PY
INNER_EOF

# Pass tasks as args to the inner script.
# Explicitly export MODEL_NAME/MODEL_API_URL/MODEL_API_KEY + KIMI_CONFIG_ENV so
# the srun'd INNER script (and the run_eval_parallel.sh it calls) receive the
# correct model alias. Without this, srun may not propagate KIMI_CONFIG_ENV
# and run_eval_parallel.sh falls back to the default config.env.
export SLURM_MEM SLURM_CPUS SLURM_TIME SLURM_PARTITION SLURM_NODELIST RELAY_API_KEY RELAY_PORT
export RUN_ID DUMP_ROOT MAX_CONCURRENT MAX_CONCURRENT_CAP AUTO_AUDIT_HTML MOCK_PORT_WAIT_LOOPS
export PG_PORT_BASE PG_RUNTIME_ROOT PG_PORT_LEASE_ROOT
export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_CONFIG_ENV KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
export KIMI_MODEL_THINKING_EFFORT="${KIMI_MODEL_THINKING_EFFORT:-}"
export KIMI_PLAN_FIRST="${KIMI_PLAN_FIRST:-}"

echo "[slurm-launch] dispatching to slurm..."
echo "[slurm-launch] DEBUG: MODEL_NAME=${MODEL_NAME:-<unset>} KIMI_CONFIG_ENV=${KIMI_CONFIG_ENV:-<unset>} RELAY_PORT=${RELAY_PORT} effort=${KIMI_MODEL_THINKING_EFFORT:-} subagents=${KIMI_SUBAGENTS-<default 3>}"
srun -p "$SLURM_PARTITION" \
     "${SRUN_NODE_ARGS[@]}" \
     -N1 -n1 -c"$SLURM_CPUS" \
     --mem="$SLURM_MEM" --time="$SLURM_TIME" \
     --job-name="${SLURM_JOB_NAME:-ds-v4-flash-eval}" \
     bash -c "$INNER" _ \
     "${MODEL_NAME:?MODEL_NAME must be set}" \
     "${KIMI_CONFIG_ENV:?KIMI_CONFIG_ENV must be set}" \
     "${MODEL_API_URL:?MODEL_API_URL must be set}" \
     "${MODEL_API_KEY:?MODEL_API_KEY must be set}" \
     "${KIMI_MAX_CONTEXT:-262144}" \
     "${KIMI_TASK_TIMEOUT_S:-7200}" \
     "${RELAY_PORT:-19317}" \
     "${KIMI_MODEL_THINKING_EFFORT:-}" \
     "${KIMI_SUBAGENTS-__UNSET__}" \
     "${TASKS[@]}"
