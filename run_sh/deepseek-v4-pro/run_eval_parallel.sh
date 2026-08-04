#!/bin/bash
# Safe high-concurrency eval for deepseek-v4-pro on login01.
#
# Isolation model (same idea as upstream run_parallel.sh):
#   - Each task gets its OWN PostgreSQL (unique port + pgdata)
#   - Each task gets its OWN ephemeral enroot rootfs (copied from template)
#   - Host dumps go to dumps/deepseek-v4-pro/<task>/<timestamp>/
#   => no shared DB / workspace cross-contamination
#
# Login01 resource guidance (2026-07-27 snapshot):
#   96 CPUs, load ~9–10, ~166 GiB mem available, NFS /lintaoLab2
#   Per slot ≈ 3.8G rootfs on NFS + ~0.5–2G RAM (PG + MCP/agent)
#   Recommended MAX_CONCURRENT on login01: 3 (default). Soft cap 5.
#   Higher concurrency mainly hurts NFS (parallel rootfs copies), not CPU.
#
# Usage:
#   bash run_sh/deepseek-v4-pro/run_eval_parallel.sh
#   MAX_CONCURRENT=3 bash run_sh/deepseek-v4-pro/run_eval_parallel.sh
#   bash run_sh/deepseek-v4-pro/run_eval_parallel.sh <task1> <task2> ...

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
RUNTIME_ROOT="${TOOLATHLON_EVAL_DOCKER_ROOT:-/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers}"
CONFIG_ENV="${SCRIPT_DIR}/config.env"

# shellcheck disable=SC1091
source "$RUNTIME_ROOT/env.sh"
# shellcheck disable=SC1091
source "$CONFIG_ENV"

export MODEL_NAME MODEL_PLATFORM MODEL_PROVIDER MODEL_API_KEY MODEL_API_URL
export MODEL_GREEDY MODEL_TEMPERATURE MODEL_TOP_P MODEL_N
export no_proxy="${no_proxy:+$no_proxy,}127.0.0.1,localhost,172.16.55.136"
export NO_PROXY="$no_proxy"

MAX_STEPS="${MAX_STEPS:-100}"
# Default 3 for login01; override via env
MAX_CONCURRENT="${MAX_CONCURRENT:-3}"
if (( MAX_CONCURRENT > 5 )); then
  echo "[warn] MAX_CONCURRENT=$MAX_CONCURRENT is high for login01; capping at 5 (NFS/rootfs pressure)." >&2
  MAX_CONCURRENT=5
fi

DUMP_ROOT="${DUMP_ROOT:-$PROJECT_ROOT/dumps/deepseek-v4-pro}"
PARALLEL_ROOT="${RUNTIME_ROOT}/parallel_runs"
PG_PORT_BASE="${PG_PORT_BASE:-25432}"
INIT_SQL="${PROJECT_ROOT}/db/init.sql.gz"
AGENT_TEMPLATE="${ENROOT_DATA_PATH}/toolathlon-pack"
AGENT_SQSH="${TOOLATHLON_AGENT_SQSH}"

TASKS=(
  arxiv-lit-review-gsheet
  canvas-assignment-effectiveness-ppt-notion-email
  fetch-howtocook-catering-excel-gcal-email
)
if [[ $# -gt 0 ]]; then
  TASKS=("$@")
fi

mkdir -p "$DUMP_ROOT" "$PARALLEL_ROOT" "${RUNTIME_ROOT}/logs"
cd "$PROJECT_ROOT"

RUN_ID="$(date +%Y%m%d-%H%M%S)"
SUMMARY="$DUMP_ROOT/summary_parallel_${RUN_ID}.csv"
echo "task,status,exit_code,output_dir,pg_port,duration_s" > "$SUMMARY"
SUMMARY_LOCK="$DUMP_ROOT/.summary_parallel.lock"

log() { echo "[$(date +%H:%M:%S)] $*"; }

need_bins() {
  for b in initdb pg_ctl pg_isready psql createdb enroot rsync; do
    command -v "$b" >/dev/null 2>&1 || { echo "[error] missing: $b (conda activate toolathlon_gym?)" >&2; exit 1; }
  done
  [[ -f "$INIT_SQL" ]] || { echo "[error] missing $INIT_SQL" >&2; exit 1; }
  [[ -d "$AGENT_TEMPLATE" || -f "$AGENT_SQSH" ]] || {
    echo "[error] need agent template $AGENT_TEMPLATE or sqsh $AGENT_SQSH" >&2
    exit 1
  }
}

append_summary() {
  while ! mkdir "$SUMMARY_LOCK" 2>/dev/null; do sleep 0.05; done
  echo "$1" >> "$SUMMARY"
  rmdir "$SUMMARY_LOCK"
}

find_free_port() {
  local port=$1
  while ss -ltn 2>/dev/null | awk '{print $4}' | rg -q ":${port}$"; do
    port=$((port + 1))
  done
  echo "$port"
}

start_isolated_pg() {
  local pgdata="$1" port="$2" logfile="$3"
  mkdir -p "$pgdata"
  if [[ ! -f "$pgdata/PG_VERSION" ]]; then
    initdb -D "$pgdata" -U "$PGUSER" --auth-local=trust --auth-host=md5 --encoding=UTF8 --locale=C >/dev/null
    cat >>"$pgdata/postgresql.conf" <<EOF
listen_addresses = '127.0.0.1'
port = ${port}
max_connections = 40
shared_buffers = 128MB
work_mem = 8MB
maintenance_work_mem = 64MB
EOF
    echo "host all all 127.0.0.1/32 md5" >>"$pgdata/pg_hba.conf"
    echo "host all all ::1/128 md5" >>"$pgdata/pg_hba.conf"
  fi
  pg_ctl -D "$pgdata" -l "$logfile" -o "-k /tmp -p ${port}" start >/dev/null

  # IMPORTANT: parent shell exports PGDATABASE=toolathlon_gym. pg_isready/psql
  # would otherwise probe that DB before it exists → FATAL spam / flaky ready.
  local i=0
  while (( i < 60 )); do
    if env -u PGDATABASE pg_isready -h /tmp -p "$port" -U "$PGUSER" -d postgres >/dev/null 2>&1; then
      break
    fi
    sleep 1
    i=$((i + 1))
  done
  env -u PGDATABASE pg_isready -h /tmp -p "$port" -U "$PGUSER" -d postgres >/dev/null 2>&1 || return 1

  # password + create DB (always against maintenance DB "postgres")
  env -u PGDATABASE psql -h /tmp -p "$port" -U "$PGUSER" -d postgres -v ON_ERROR_STOP=1 \
    -c "ALTER USER ${PGUSER} WITH PASSWORD '${PGPASSWORD}';" >/dev/null
  env -u PGDATABASE psql -h /tmp -p "$port" -U "$PGUSER" -d postgres -v ON_ERROR_STOP=1 \
    -c "SELECT 1 FROM pg_database WHERE datname='${PGDATABASE}'" | grep -q 1 \
    || env -u PGDATABASE createdb -h /tmp -p "$port" -U "$PGUSER" "$PGDATABASE"

  # Confirm target DB exists before loading dump
  env -u PGDATABASE psql -h /tmp -p "$port" -U "$PGUSER" -d postgres -tAc \
    "SELECT 1 FROM pg_database WHERE datname='${PGDATABASE}'" | grep -q 1 || return 1

  gunzip -c "$INIT_SQL" | env -u PGDATABASE psql -h /tmp -p "$port" -U "$PGUSER" -d "$PGDATABASE" \
    -v ON_ERROR_STOP=0 >/dev/null 2>&1 || true

  # FK fix used by emails MCP (TCP + password path, same as agents)
  PGPASSWORD="$PGPASSWORD" psql -h 127.0.0.1 -p "$port" -U "$PGUSER" -d "$PGDATABASE" -v ON_ERROR_STOP=0 <<'SQL' >/dev/null 2>&1 || true
ALTER TABLE email.sent_log DROP CONSTRAINT IF EXISTS sent_log_message_id_fkey;
ALTER TABLE email.sent_log ADD CONSTRAINT sent_log_message_id_fkey
  FOREIGN KEY (message_id) REFERENCES email.messages(id) ON DELETE CASCADE;
SQL

  # Final readiness on the real app DB
  PGPASSWORD="$PGPASSWORD" pg_isready -h 127.0.0.1 -p "$port" -U "$PGUSER" -d "$PGDATABASE" >/dev/null 2>&1 || return 1
  return 0
}

stop_isolated_pg() {
  local pgdata="$1"
  if [[ -d "$pgdata" ]] && pg_ctl -D "$pgdata" status >/dev/null 2>&1; then
    pg_ctl -D "$pgdata" -m fast stop >/dev/null 2>&1 || true
  fi
}

make_agent_rootfs() {
  local name="$1"
  # Prefer copying existing template (avoids re-unsquashfs under concurrency)
  if [[ -d "$AGENT_TEMPLATE" ]]; then
    rm -rf "${ENROOT_DATA_PATH}/${name}"
    # NFS: cp -a is heavy but more reliable than parallel unsquashfs storms
    cp -a "$AGENT_TEMPLATE" "${ENROOT_DATA_PATH}/${name}"
  else
    enroot create -n "$name" "$AGENT_SQSH"
  fi
}

run_one_task() {
  local TASK="$1"
  local SLOT="$2"
  local SAFE TASK_HASH TASK_ID PGPORT PGDATA INST OUTDIR TASK_LOG START_TS END_TS RC STATUS
  SAFE="$(echo "$TASK" | tr '/' '-')"
  TASK_HASH="$(echo -n "$TASK-$RUN_ID-$SLOT-$$" | md5sum | cut -c1-8)"
  TASK_ID="${SAFE}-${TASK_HASH}"
  PGPORT="$(find_free_port $((PG_PORT_BASE + SLOT)))"
  PGDATA="${PARALLEL_ROOT}/pgdata/${TASK_ID}"
  INST="agent-${TASK_ID}"
  OUTDIR="${DUMP_ROOT}/${TASK}/${RUN_ID}_slot${SLOT}"
  TASK_LOG="${OUTDIR}/run.log"
  mkdir -p "$OUTDIR" "$PGDATA"

  START_TS=$(date +%s)
  log "START  $TASK  (slot=$SLOT pg_port=$PGPORT)"

  {
    echo "=== isolated PG on :$PGPORT ==="
    if ! start_isolated_pg "$PGDATA" "$PGPORT" "${OUTDIR}/postgres.log"; then
      echo "[error] failed to start isolated postgres"
      append_summary "${TASK},pg_fail,1,${OUTDIR},${PGPORT},$(( $(date +%s) - START_TS ))"
      rm -rf "$PGDATA"
      return 1
    fi

    echo "=== create agent rootfs $INST ==="
    make_agent_rootfs "$INST"
    mkdir -p "${ENROOT_DATA_PATH}/${INST}/workspace/dumps"
    if [[ ! -f "${ENROOT_DATA_PATH}/${INST}/workspace/scripts/eval_config.json" ]]; then
      echo "[error] missing /workspace/scripts/eval_config.json in rootfs $INST"
      # best-effort copy from host project
      mkdir -p "${ENROOT_DATA_PATH}/${INST}/workspace/scripts"
      cp -f "${PROJECT_ROOT}/scripts/eval_config.json" \
        "${ENROOT_DATA_PATH}/${INST}/workspace/scripts/eval_config.json" || true
    fi
    if [[ ! -f "${ENROOT_DATA_PATH}/${INST}/workspace/main.py" ]]; then
      echo "[error] missing /workspace/main.py in rootfs"
      append_summary "${TASK},rootfs_fail,1,${OUTDIR},${PGPORT},$(( $(date +%s) - START_TS ))"
      stop_isolated_pg "$PGDATA"
      rm -rf "$PGDATA" "${ENROOT_DATA_PATH}/${INST}"
      return 1
    fi

    ENV_ARGS=(
      -e "PGHOST=127.0.0.1"
      -e "PG_HOST=127.0.0.1"
      -e "PGPORT=${PGPORT}"
      -e "PG_PORT=${PGPORT}"
      -e "PGUSER=${PGUSER}"
      -e "PG_USER=${PGUSER}"
      -e "PGPASSWORD=${PGPASSWORD}"
      -e "PG_PASSWORD=${PGPASSWORD}"
      -e "PGDATABASE=${PGDATABASE}"
      -e "PG_DATABASE=${PGDATABASE}"
      -e "LOCAL_SERVERS_PATH=/opt/local_servers"
      -e "PYTHON_BIN=/opt/venv/bin/python3"
      -e "PATH=/opt/venv/bin:/root/.local/bin:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin"
      -e "VIRTUAL_ENV=/opt/venv"
      -e "HOME=/root"
      -e "PWD=/workspace"
      -e "MODEL_NAME=${MODEL_NAME}"
      -e "MODEL_PLATFORM=${MODEL_PLATFORM}"
      -e "MODEL_PROVIDER=${MODEL_PROVIDER}"
      -e "MODEL_API_KEY=${MODEL_API_KEY}"
      -e "MODEL_API_URL=${MODEL_API_URL}"
      -e "MODEL_GREEDY=${MODEL_GREEDY:-1}"
      -e "MODEL_TEMPERATURE=${MODEL_TEMPERATURE:-0}"
      -e "MODEL_TOP_P=${MODEL_TOP_P:-1}"
      -e "no_proxy=${no_proxy}"
      -e "NO_PROXY=${NO_PROXY}"
    )
    [[ -n "${MODEL_N:-}" ]] && ENV_ARGS+=(-e "MODEL_N=${MODEL_N}")
    for var in http_proxy https_proxy HTTP_PROXY HTTPS_PROXY; do
      [[ -n "${!var:-}" ]] && ENV_ARGS+=(-e "${var}=${!var}")
    done

    echo "=== run main.py (cwd=/workspace) ==="
    set +e
    enroot start -r -w "${ENV_ARGS[@]}" "$INST" \
      /bin/bash -c "cd /workspace && exec /opt/venv/bin/python3 main.py --eval_config /workspace/scripts/eval_config.json --task_dir '${TASK}' --max_steps '${MAX_STEPS}' --debug"
    RC=$?
    set -e

    echo "=== sync dumps ==="
    rsync -a "${ENROOT_DATA_PATH}/${INST}/workspace/dumps/" "$OUTDIR/" || true

    echo "=== cleanup ==="
    enroot remove -f "$INST" >/dev/null 2>&1 || rm -rf "${ENROOT_DATA_PATH}/${INST}" || true
    stop_isolated_pg "$PGDATA"
    rm -rf "$PGDATA"

    END_TS=$(date +%s)
    if [[ $RC -eq 0 ]]; then STATUS=success; else STATUS=failed; fi
    append_summary "${TASK},${STATUS},${RC},${OUTDIR},${PGPORT},$((END_TS - START_TS))"
    log "DONE   $TASK -> $STATUS (exit=$RC, $((END_TS - START_TS))s, port=$PGPORT)"
    return "$RC"
  } >"$TASK_LOG" 2>&1
}

export -f log append_summary find_free_port start_isolated_pg stop_isolated_pg make_agent_rootfs run_one_task
export PROJECT_ROOT RUNTIME_ROOT DUMP_ROOT PARALLEL_ROOT PG_PORT_BASE INIT_SQL
export AGENT_TEMPLATE AGENT_SQSH ENROOT_DATA_PATH RUN_ID SUMMARY SUMMARY_LOCK
export MODEL_NAME MODEL_PLATFORM MODEL_PROVIDER MODEL_API_KEY MODEL_API_URL MAX_STEPS
export MODEL_GREEDY MODEL_TEMPERATURE MODEL_TOP_P MODEL_N
export PGUSER PGPASSWORD PGDATABASE no_proxy NO_PROXY http_proxy https_proxy HTTP_PROXY HTTPS_PROXY

need_bins

log "=============================================="
log "  Parallel safe eval — deepseek-v4-pro"
log "  Model:         $MODEL_NAME @ $MODEL_API_URL"
log "  Sampling:      greedy=${MODEL_GREEDY:-1} temp=${MODEL_TEMPERATURE:-0} top_p=${MODEL_TOP_P:-1}"
log "  Tasks:         ${#TASKS[@]}"
log "  Max concurrent:$MAX_CONCURRENT  (login01 recommend 3, soft-cap 5)"
log "  Dump root:     $DUMP_ROOT"
log "  PG port base:  $PG_PORT_BASE"
log "=============================================="

# API check
if command -v curl >/dev/null 2>&1; then
  code="$(curl -sS --connect-timeout 8 -o /dev/null -w '%{http_code}' \
    -H "Authorization: Bearer ${MODEL_API_KEY}" \
    "${MODEL_API_URL%/}/v1/models" || true)"
  log "API /v1/models -> HTTP $code"
fi

# FIFO semaphore on local tmpfs — NFS named pipes drop worker→parent tokens.
FIFO="/dev/shm/toolathlon_sem_${RUN_ID}_$$"
rm -f "$FIFO"
mkfifo "$FIFO"
exec 3<>"$FIFO"
rm -f "$FIFO"
cleanup_fifo() { exec 3>&- 2>/dev/null || true; rm -f "$FIFO" 2>/dev/null || true; }
trap cleanup_fifo EXIT
for ((i = 0; i < MAX_CONCURRENT; i++)); do
  echo >&3
done

PIDS=()
SLOT=0
FAILED=0
for TASK in "${TASKS[@]}"; do
  read -u 3
  # set +e: failed tasks must still return the FIFO token, or the pool deadlocks.
  (
    set +e
    run_one_task "$TASK" "$SLOT"
    rc=$?
    echo >&3
    exit "$rc"
  ) &
  PIDS+=($!)
  SLOT=$((SLOT + 1))
done

for pid in "${PIDS[@]}"; do
  wait "$pid" || FAILED=$((FAILED + 1))
done
cleanup_fifo
trap - EXIT

log "=============================================="
log "  Finished. Failed workers: $FAILED / ${#TASKS[@]}"
log "  Summary: $SUMMARY"
log "=============================================="
# Print summary table
if [[ -f "$SUMMARY" ]]; then
  column -t -s, "$SUMMARY" 2>/dev/null || cat "$SUMMARY"
fi

exit "$FAILED"
