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
#   RELAY_API_KEY  API key used only for relay health checks (required)

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"

# ── login-node API relay (idempotent) ────────────────────────────────────────
LOGIN_IP="192.168.180.240"
RELAY_PORT=19317
PYTHON_BIN="${PYTHON_BIN:-/storage/lintaoLab/bowending/miniconda3/envs/dbw_dev/bin/python3}"
RELAY_API_KEY="${RELAY_API_KEY:?Set RELAY_API_KEY before launching the eval}"

# Start relay on the login node (where this script runs). api_relay.py is
# idempotent: it exits 0 immediately if the port is already taken, otherwise
# it blocks forever serving — so we background it.
nohup "$PYTHON_BIN" "$SCRIPT_DIR/api_relay.py" >/dev/shm/api_relay.log 2>&1 &
disown || true
# Give it a moment to bind, then confirm.
if curl -sS -m 5 -o /dev/null \
     -H "Authorization: Bearer ${RELAY_API_KEY}" \
     "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
  echo "[slurm-launch] API relay healthy on 127.0.0.1:${RELAY_PORT}"
else
  echo "[slurm-launch] WARNING: relay health check failed, check /dev/shm/api_relay.log"
fi

# ── slurm resource request ───────────────────────────────────────────────────
SLURM_MEM="${SLURM_MEM:-128G}"
SLURM_CPUS="${SLURM_CPUS:-64}"
SLURM_TIME="${SLURM_TIME:-04:00:00}"
SLURM_PARTITION="${SLURM_PARTITION:-linlab}"

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

echo "[slurm-launch] partition=$SLURM_PARTITION mem=$SLURM_MEM cpus=$SLURM_CPUS time=$SLURM_TIME"
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

# Compute nodes inherit http_proxy=127.0.0.1:7890 (nonexistent). Clear all
# proxy vars and whitelist the login node so API calls go direct to the relay.
unset http_proxy https_proxy HTTP_PROXY HTTPS_PROXY ALL_PROXY all_proxy
export no_proxy="127.0.0.1,localhost,::1,192.168.180.240,104.168.43.47,172.16.0.0/12,10.0.0.0/8"

# enroot runtime on this cluster.
# The user-installed enroot lives under ~/.local; its wrapper script
# (.local/bin/enroot) sets ENROOT_LIBRARY_PATH/SYSCONF_PATH itself, so we just
# need it first on PATH. Do NOT set ENROOT_LIBRARY_PATH manually here — the
# correct path is ~/.local/enroot/lib (not usr/lib/enroot).
#
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
export PATH="${ENROOT_LOCAL}/bin:/storage/lintaoLab/bowending/.local/bin:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin:/storage/lintaoLab/bowending/miniconda3/envs/toolathlon_gym/bin"

# Use /dev/shm (1TB tmpfs on compute nodes) for all enroot working dirs.
export ENROOT_DATA_PATH="/dev/shm/enroot_data"
export ENROOT_TEMP_PATH="/dev/shm/enroot_tmp"
export ENROOT_RUNTIME_PATH="/dev/shm/enroot_runtime"
export ENROOT_CACHE_PATH="/dev/shm/enroot_cache"
mkdir -p "$ENROOT_DATA_PATH" "$ENROOT_TEMP_PATH" "$ENROOT_RUNTIME_PATH" "$ENROOT_CACHE_PATH"

# The base rootfs (toolathlon-pack) lives on NFS so every compute node can see
# it. Each worker copies it into its own /dev/shm/enroot_data/agent-<task> (fast
# tmpfs-to-tmpfs copy). On compute nodes /dev/shm is 1TB so 6x4.7GB is trivial.
# The login-node shm copy is NOT visible here, so point at the NFS snapshot.
export AGENT_TEMPLATE="/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers/toolathlon-pack-rootfs"

# Sanity: relay reachable from compute node?
echo "[inner] relay check:"
curl -sS -m 5 -o /dev/null -w "  HTTP %{http_code}\n" \
  -H "Authorization: Bearer ${RELAY_API_KEY}" \
  http://192.168.180.240:19317/v1/models || {
    echo "[inner] FATAL: cannot reach API relay from compute node"
    exit 1
  }

echo "[inner] enroot: $(enroot version 2>&1)"
echo "[inner] /dev/shm: $(df -h /dev/shm | tail -1 | awk '{print $2" total, "$4" free"}')"

# Hand off to the existing parallel launcher.
exec bash run_sh/kimi-code-deepseek-v4-flash/run_eval_parallel.sh "$@"
INNER_EOF

# Pass tasks as args to the inner script.
export SLURM_MEM SLURM_CPUS SLURM_TIME SLURM_PARTITION RELAY_API_KEY

echo "[slurm-launch] dispatching to slurm..."
exec srun -p "$SLURM_PARTITION" \
     -N1 -n1 -c"$SLURM_CPUS" \
     --mem="$SLURM_MEM" --time="$SLURM_TIME" \
     --job-name="ds-v4-flash-eval" \
     bash -c "$INNER" _ "${TASKS[@]}"
