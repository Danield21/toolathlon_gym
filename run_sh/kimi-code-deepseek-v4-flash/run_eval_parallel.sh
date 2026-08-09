#!/bin/bash
# Parallel kimi-code eval for deepseek-v4-flash @ Toolathlon-GYM.
#
# Differences from run_sh/kimi-code/ (MiniMax-M3):
#   - MODEL_NAME=deepseek-v4-flash (same API URL/key)
#   - Dump root: dumps/kimi-code_deepseek-v4-flash/
#   - PG port base 26432 (avoids collision with MiniMax batch on 25432+)
#
# Usage (from toolathlon_gym/):
#   bash run_sh/kimi-code-deepseek-v4-flash/run_eval_parallel.sh
#   bash run_sh/kimi-code-deepseek-v4-flash/run_19cases.sh
#   MAX_CONCURRENT=6 bash run_sh/kimi-code-deepseek-v4-flash/run_eval_parallel.sh task1 task2 ...
#
# Ablation env vars (same as MiniMax launcher):
#   KIMI_SUBAGENTS=coder,explore
#   KIMI_EXAMPLES_FILE=examples_none.md
#   KIMI_COORDINATION_FILE=subagent_coordination_default.md

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"

export KIMI_CONFIG_ENV="${SCRIPT_DIR}/config.env"
export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_deepseek-v4-flash}"
# Separate port range + lease dir so this batch can run alongside MiniMax-M3.
export PG_PORT_BASE="${PG_PORT_BASE:-26432}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/toolathlon_pg_port_leases_deepseek_${UID}}"

# Use /dev/shm (tmpfs) for enroot data/temp to avoid NFS slowness and
# root-partition-full issues on this cluster. The root-partition (/) is 100%
# full from other users' training shards; tmpfs has 126G free. enroot create
# and cp -a of rootfs are near-instant on tmpfs vs 8+min on NFS.
export ENROOT_DATA_PATH="${ENROOT_DATA_PATH:-/dev/shm/enroot_data}"
export ENROOT_TEMP_PATH="${ENROOT_TEMP_PATH:-/dev/shm/enroot_tmp}"
export ENROOT_RUNTIME_PATH="${ENROOT_RUNTIME_PATH:-/dev/shm/enroot_runtime}"
export ENROOT_CACHE_PATH="${ENROOT_CACHE_PATH:-/dev/shm/enroot_cache}"
mkdir -p "$ENROOT_DATA_PATH" "$ENROOT_TEMP_PATH" "$ENROOT_RUNTIME_PATH" "$ENROOT_CACHE_PATH"

# Ensure /usr/sbin is on PATH (unsquashfs/mksquashfs live there); without it
# enroot create fails with "Command not found: unsquashfs".
export PATH="/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin:${PATH}"

exec bash "${SCRIPT_DIR}/../kimi-code/run_eval_parallel.sh" "$@"
