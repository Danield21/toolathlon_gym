#!/bin/bash
# Safe high-concurrency eval for kimi-code (MiniMax-M3) on login01.
#
# Isolation model (same idea as upstream run_parallel.sh):
#   - Each task gets its OWN PostgreSQL (unique port + pgdata)
#   - Each task gets its OWN ephemeral enroot rootfs (copied from template)
#   - Host dumps go to dumps/kimi-code/<task>/<timestamp>/
#   => no shared DB / workspace cross-contamination
#
# Agent: kimi-code CLI (headless print mode) driven by kimi_harness/kimi_main.py
# inside the rootfs; only the task's MCP/local tools are exposed.
#
# Usage:
#   bash run_sh/kimi-code/run_eval_parallel.sh
#   MAX_CONCURRENT=3 bash run_sh/kimi-code/run_eval_parallel.sh
#   bash run_sh/kimi-code/run_eval_parallel.sh <task1> <task2> ...

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
RUNTIME_ROOT="${TOOLATHLON_EVAL_DOCKER_ROOT:-/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers}"
CONFIG_ENV="${KIMI_CONFIG_ENV:-${SCRIPT_DIR}/config.env}"
ENV_MAX_STEPS="${MAX_STEPS:-}"
ENV_MAX_CONCURRENT="${MAX_CONCURRENT:-}"
ENV_DUMP_ROOT="${DUMP_ROOT:-}"
ENV_PG_PORT_BASE="${PG_PORT_BASE:-}"
ENV_PG_RUNTIME_ROOT="${PG_RUNTIME_ROOT:-}"
ENV_RUN_ID="${RUN_ID:-}"
AUTO_AUDIT_HTML="${AUTO_AUDIT_HTML:-1}"

# shellcheck disable=SC1091
source "$RUNTIME_ROOT/env.sh"
# shellcheck disable=SC1091
source "$CONFIG_ENV"

export MODEL_NAME MODEL_API_KEY MODEL_API_URL
export KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S KIMI_PROVIDER_NAME KIMI_MODEL_ALIAS
export no_proxy="${no_proxy:+$no_proxy,}127.0.0.1,localhost,::1,192.168.180.240,104.168.43.47"
export NO_PROXY="$no_proxy"

MAX_STEPS="${ENV_MAX_STEPS:-${MAX_STEPS:-100}}"
# Default 6 for gnho019; override via env. Machine has 96 cores / 160 GB free
# RAM / 4.7 TB NFS, so the real ceiling is NFS bandwidth during rootfs copies
# and the remote API rate limit, not local CPU/RAM.
MAX_CONCURRENT="${ENV_MAX_CONCURRENT:-${MAX_CONCURRENT:-6}}"
if (( MAX_CONCURRENT > 8 )); then
  echo "[warn] MAX_CONCURRENT=$MAX_CONCURRENT is very high (NFS/API pressure); capping at 8." >&2
  MAX_CONCURRENT=8
fi

DUMP_ROOT="${ENV_DUMP_ROOT:-${DUMP_ROOT:-$PROJECT_ROOT/dumps/kimi-code}}"
KIMI_DIST="${KIMI_DIST:-$RUNTIME_ROOT/kimi_dist/kimi-code}"
PARALLEL_ROOT="${RUNTIME_ROOT}/parallel_runs"
PG_PORT_BASE="${ENV_PG_PORT_BASE:-${PG_PORT_BASE:-25432}}"
PG_RUNTIME_ROOT="${ENV_PG_RUNTIME_ROOT:-${PG_RUNTIME_ROOT:-/dev/shm/toolathlon_pg_${UID}}}"
PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/toolathlon_pg_port_leases_${UID}}"
PG_TEST_ONLY="${PG_TEST_ONLY:-0}"
INIT_SQL="${PROJECT_ROOT}/db/init.sql.gz"
AGENT_TEMPLATE="${AGENT_TEMPLATE:-${ENROOT_DATA_PATH}/toolathlon-pack}"
AGENT_SQSH="${TOOLATHLON_AGENT_SQSH}"

# Default: only real cases containing task_config.json. This intentionally
# excludes support directories such as tasks/finalpool/.utils.
mapfile -t TASKS < <(
  find "${PROJECT_ROOT}/tasks/finalpool" -mindepth 2 -maxdepth 2 \
    -type f -name task_config.json -printf '%h\n' \
    | sed 's#.*/##' \
    | sort
)
if [[ $# -gt 0 ]]; then
  TASKS=("$@")
fi
if [[ ${#TASKS[@]} -eq 0 ]]; then
  echo "[error] no tasks found under ${PROJECT_ROOT}/tasks/finalpool" >&2
  exit 1
fi
for task in "${TASKS[@]}"; do
  if [[ ! -f "${PROJECT_ROOT}/tasks/finalpool/${task}/task_config.json" ]]; then
    echo "[error] invalid task (missing task_config.json): $task" >&2
    exit 1
  fi
done

mkdir -p "$DUMP_ROOT" "$PARALLEL_ROOT" "${RUNTIME_ROOT}/logs"
cd "$PROJECT_ROOT"

RUN_ID="${ENV_RUN_ID:-$(date +%Y%m%d-%H%M%S)}"
SUMMARY="$DUMP_ROOT/summary_parallel_${RUN_ID}.csv"
echo "task,status,exit_code,output_dir,pg_port,duration_s" > "$SUMMARY"
SUMMARY_LOCK="/dev/shm/toolathlon_summary_${UID}_${RUN_ID}_$$.lock"
PORT_LEASES=()

log() { echo "[$(date +%H:%M:%S)] $*"; }

need_bins() {
  for b in initdb pg_ctl pg_isready psql createdb enroot rsync setsid ss; do
    command -v "$b" >/dev/null 2>&1 || { echo "[error] missing: $b (conda activate toolathlon_gym?)" >&2; exit 1; }
  done
  [[ -f "$INIT_SQL" ]] || { echo "[error] missing $INIT_SQL" >&2; exit 1; }
  [[ -d "$AGENT_TEMPLATE" || -f "$AGENT_SQSH" ]] || {
    echo "[error] need agent template $AGENT_TEMPLATE or sqsh $AGENT_SQSH" >&2
    exit 1
  }
  [[ "$PG_RUNTIME_ROOT" == /tmp/* || "$PG_RUNTIME_ROOT" == /dev/shm/* ]] || {
    echo "[error] PG_RUNTIME_ROOT must be below /tmp or /dev/shm: $PG_RUNTIME_ROOT" >&2
    exit 1
  }
  [[ "$PG_PORT_LEASE_ROOT" == /dev/shm/* ]] || {
    echo "[error] PG_PORT_LEASE_ROOT must be below /dev/shm: $PG_PORT_LEASE_ROOT" >&2
    exit 1
  }
  mkdir -p "$PG_RUNTIME_ROOT" "$PG_PORT_LEASE_ROOT"
  chmod 700 "$PG_RUNTIME_ROOT" "$PG_PORT_LEASE_ROOT" 2>/dev/null || true
  (( MAX_CONCURRENT >= 1 )) || { echo "[error] MAX_CONCURRENT must be >= 1" >&2; exit 1; }
  (( PG_PORT_BASE >= 1024 && PG_PORT_BASE <= 65535 )) || {
    echo "[error] PG_PORT_BASE must be between 1024 and 65535" >&2
    exit 1
  }
  [[ "$PG_TEST_ONLY" == 0 || "$PG_TEST_ONLY" == 1 ]] || {
    echo "[error] PG_TEST_ONLY must be 0 or 1" >&2
    exit 1
  }
  grep -q "Results are paginated and include summary totals" \
    "${PROJECT_ROOT}/local_servers/mcp-canvas-lms/build/index.js" || {
      echo "[error] Canvas MCP build is stale; rebuild local_servers/mcp-canvas-lms before running." >&2
      exit 1
    }
  local shm_path shm_free_kb template_kb reserve_kb need_kb
  shm_path="$ENROOT_DATA_PATH"
  [[ -d "$shm_path" ]] || shm_path="/dev/shm"
  shm_free_kb="$(df -Pk "$shm_path" | awk 'NR==2 {print $4}')"
  if [[ -d "$AGENT_TEMPLATE" ]]; then
    template_kb="$(du -sk "$AGENT_TEMPLATE" | awk '{print $1}')"
  else
    template_kb="$(du -sk "$AGENT_SQSH" | awk '{print $1}')"
  fi
  # /dev/shm is RAM-backed. Budget copied rootfs plus PG/MCP/browser scratch
  # before starting the fan-out, so a run fails cleanly instead of half-filling
  # tmpfs and leaving partial artifacts behind.
  reserve_kb=$(( MAX_CONCURRENT * ${SHM_PER_WORKER_RESERVE_GB:-8} * 1024 * 1024 ))
  need_kb=$(( template_kb * MAX_CONCURRENT + reserve_kb ))
  if (( shm_free_kb < need_kb )); then
    echo "[error] insufficient /dev/shm for MAX_CONCURRENT=$MAX_CONCURRENT" >&2
    echo "[error] free=${shm_free_kb}KB need~=${need_kb}KB template=${template_kb}KB reserve=${reserve_kb}KB" >&2
    exit 1
  fi
}

append_summary() {
  while ! mkdir "$SUMMARY_LOCK" 2>/dev/null; do sleep 0.05; done
  echo "$1" >> "$SUMMARY"
  rmdir "$SUMMARY_LOCK"
  if [[ -n "${WORKER_TASK:-}" ]]; then
    WORKER_SUMMARY_WRITTEN=1
  fi
}

port_is_listening() {
  local port="$1"
  local hexport
  hexport="$(printf '%04X' "$port")"
  # /proc/net/tcp is kernel-direct and unaffected by ss output format changes.
  # Column 4 (st) == 0A means LISTEN. We check both IPv4 and IPv6.
  awk -v p="$hexport" '
    $2 ~ ":" p && $4 == "0A" { found=1; exit }
    END { exit !found }
  ' /proc/net/tcp /proc/net/tcp6 2>/dev/null
}

# Called only by the parent shell. The atomic lease directory prevents another
# concurrently launched evaluator from choosing the same port between the
# listen check and PostgreSQL bind. Leases are retained for the whole run, so a
# port is never reused by two tasks in one summary.
TASK_LOCK_ROOT="/dev/shm/toolathlon_task_locks_${UID}_${RUN_ID}"
mkdir -p "$TASK_LOCK_ROOT"
chmod 700 "$TASK_LOCK_ROOT" 2>/dev/null || true

# claim_task prevents the same task from being launched twice (e.g. FIFO
# semaphore race where a closed FIFO causes read to return immediately).
claim_task() {
  local task="$1"
  local safe
  safe="$(printf '%s' "$task" | tr -cs 'A-Za-z0-9_.-' '-')"
  mkdir "$TASK_LOCK_ROOT/$safe" 2>/dev/null
}

allocate_free_port() {
  local lease owner
  while (( NEXT_PG_PORT <= 65535 )); do
    lease="${PG_PORT_LEASE_ROOT}/${NEXT_PG_PORT}"
    if mkdir "$lease" 2>/dev/null; then
      echo "$$" >"$lease/owner_pid"
      if port_is_listening "$NEXT_PG_PORT"; then
        rm -rf -- "$lease"
      else
        ALLOCATED_PG_PORT="$NEXT_PG_PORT"
        PORT_LEASES+=("$lease")
        NEXT_PG_PORT=$((NEXT_PG_PORT + 1))
        return 0
      fi
    else
      owner="$(sed -n '1p' "$lease/owner_pid" 2>/dev/null || true)"
      if [[ "$owner" =~ ^[0-9]+$ ]] && ! kill -0 "$owner" 2>/dev/null && ! port_is_listening "$NEXT_PG_PORT"; then
        rm -rf -- "$lease"
        continue
      fi
    fi
    NEXT_PG_PORT=$((NEXT_PG_PORT + 1))
  done
  echo "[error] exhausted PostgreSQL port range" >&2
  return 1
}

release_port_leases() {
  local lease
  for lease in "${PORT_LEASES[@]:-}"; do
    [[ "$lease" == "$PG_PORT_LEASE_ROOT"/* ]] && rm -rf -- "$lease"
  done
  PORT_LEASES=()
}

start_isolated_pg() {
  local pgdata="$1" socket_dir="$2" port="$3" logfile="$4" dbcheck_log="$5"
  mkdir -p "$pgdata" "$socket_dir"
  chmod 700 "$pgdata" "$socket_dir"
  if [[ ! -f "$pgdata/PG_VERSION" ]]; then
    # --auth-local=scram-sha-256 requires the superuser to have a password at
    # initdb time; pass it via a transient pwfile (deleted right after).
    local _pwfile; _pwfile="$(mktemp)"
    printf '%s\n' "$PGPASSWORD" > "$_pwfile"
    initdb -D "$pgdata" -U "$PGUSER" --pwfile="$_pwfile" \
      --auth-local=scram-sha-256 --auth-host=md5 --encoding=UTF8 --locale=C >/dev/null
    rm -f "$_pwfile"
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
  # Never continue to readiness checks when our own postmaster failed to bind.
  # Without this explicit return, pg_isready could discover another task's
  # PostgreSQL on the same port and silently cross-contaminate both tasks.
  if ! pg_ctl -D "$pgdata" -l "$logfile" -o "-k ${socket_dir} -p ${port}" start >/dev/null; then
    echo "[error] PostgreSQL failed to start for PGDATA=$pgdata port=$port" >&2
    return 1
  fi

  # IMPORTANT: parent shell exports PGDATABASE=toolathlon_gym. pg_isready/psql
  # would otherwise probe that DB before it exists → FATAL spam / flaky ready.
  local i=0
  while (( i < 60 )); do
    if env -u PGDATABASE pg_isready -h "$socket_dir" -p "$port" -U "$PGUSER" -d postgres >/dev/null 2>&1; then
      break
    fi
    sleep 1
    i=$((i + 1))
  done
  env -u PGDATABASE pg_isready -h "$socket_dir" -p "$port" -U "$PGUSER" -d postgres >/dev/null 2>&1 || return 1

  # Defense in depth: prove that the server reached through this port owns the
  # PGDATA created for this task. A mismatch is always fatal.
  local expected_pgdata actual_pgdata
  expected_pgdata="$(readlink -f "$pgdata")"
  actual_pgdata="$(env -u PGDATABASE psql -h "$socket_dir" -p "$port" -U "$PGUSER" -d postgres -Atc \
    'SHOW data_directory' 2>/dev/null || true)"
  actual_pgdata="$(readlink -f "$actual_pgdata" 2>/dev/null || echo "$actual_pgdata")"
  if [[ -z "$actual_pgdata" || "$actual_pgdata" != "$expected_pgdata" ]]; then
    echo "[error] PostgreSQL identity mismatch on port $port" >&2
    echo "[error] expected PGDATA=$expected_pgdata, got=$actual_pgdata" >&2
    return 1
  fi

  # password + create DB (always against maintenance DB "postgres")
  env -u PGDATABASE psql -h "$socket_dir" -p "$port" -U "$PGUSER" -d postgres -v ON_ERROR_STOP=1 \
    -c "ALTER USER ${PGUSER} WITH PASSWORD '${PGPASSWORD}';" >/dev/null
  if [[ "$(env -u PGDATABASE psql -h "$socket_dir" -p "$port" -U "$PGUSER" -d postgres -tAc \
      "SELECT 1 FROM pg_database WHERE datname='${PGDATABASE}'")" != 1 ]]; then
    env -u PGDATABASE createdb -h "$socket_dir" -p "$port" -U "$PGUSER" "$PGDATABASE"
  fi

  # Confirm target DB exists before loading dump
  env -u PGDATABASE psql -h "$socket_dir" -p "$port" -U "$PGUSER" -d postgres -tAc \
    "SELECT 1 FROM pg_database WHERE datname='${PGDATABASE}'" | grep -q 1 || return 1

  # A partially imported schema produces plausible-looking but invalid evals.
  # Treat any SQL error as a task setup failure instead of continuing.
  if ! gunzip -c "$INIT_SQL" | env -u PGDATABASE psql -h "$socket_dir" -p "$port" -U "$PGUSER" -d "$PGDATABASE" \
    -v ON_ERROR_STOP=1 >/dev/null; then
    echo "[error] failed to load clean Toolathlon schema on port $port" >&2
    return 1
  fi

  # FK fix used by emails MCP (TCP + password path, same as agents)
  if ! PGPASSWORD="$PGPASSWORD" psql -h 127.0.0.1 -p "$port" -U "$PGUSER" -d "$PGDATABASE" -v ON_ERROR_STOP=1 <<'SQL' >/dev/null
ALTER TABLE email.sent_log DROP CONSTRAINT IF EXISTS sent_log_message_id_fkey;
ALTER TABLE email.sent_log ADD CONSTRAINT sent_log_message_id_fkey
  FOREIGN KEY (message_id) REFERENCES email.messages(id) ON DELETE CASCADE;
SQL
  then
    echo "[error] failed to apply email.sent_log foreign-key fix on port $port" >&2
    return 1
  fi

  # Reject partial or polluted schemas before any MCP server can connect.
  if ! PGPASSWORD="$PGPASSWORD" psql -h 127.0.0.1 -p "$port" -U "$PGUSER" -d "$PGDATABASE" \
      -v ON_ERROR_STOP=1 >"$dbcheck_log" 2>&1 <<'SQL'
DO $$
BEGIN
  IF to_regclass('email.folders') IS NULL THEN
    RAISE EXCEPTION 'email.folders is missing';
  END IF;
  IF NOT EXISTS (
    SELECT 1 FROM pg_indexes
    WHERE schemaname = 'email'
      AND tablename = 'folders'
      AND indexdef ILIKE 'CREATE UNIQUE INDEX%'
      AND indexdef LIKE '%(name)%'
  ) THEN
    RAISE EXCEPTION 'email.folders(name) unique index is missing';
  END IF;
  IF EXISTS (SELECT 1 FROM pg_index WHERE NOT indisvalid OR NOT indisready) THEN
    RAISE EXCEPTION 'database contains invalid or unready indexes';
  END IF;
END
$$;
SELECT current_setting('data_directory') AS data_directory,
       current_setting('port') AS port,
       'database isolation check: OK' AS status;
SQL
  then
    echo "[error] database schema/isolation check failed on port $port" >&2
    return 1
  fi

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

WORKER_PGDATA=""
WORKER_PG_RUNTIME=""
WORKER_INST=""
WORKER_PGID=""
WORKER_TASK=""
WORKER_OUTDIR=""
WORKER_PGPORT=""
WORKER_START_TS=""
WORKER_SUMMARY_WRITTEN=0

stop_worker_process_group() {
  local pgid="${1:-}"
  local self_pgid=""
  local i

  [[ "$pgid" =~ ^[0-9]+$ ]] || return 0
  (( pgid > 1 )) || return 0
  self_pgid="$(ps -o pgid= -p "$$" 2>/dev/null | tr -d ' ')"
  if [[ -n "$self_pgid" && "$pgid" == "$self_pgid" ]]; then
    echo "[warn] refusing to terminate worker's own process group $pgid" >&2
    return 1
  fi
  kill -0 -- "-${pgid}" 2>/dev/null || return 0

  echo "[cleanup] stopping Enroot process group $pgid"
  kill -TERM -- "-${pgid}" 2>/dev/null || true
  for ((i = 0; i < 20; i++)); do
    kill -0 -- "-${pgid}" 2>/dev/null || return 0
    sleep 0.1
  done
  echo "[cleanup] force-killing Enroot process group $pgid"
  kill -KILL -- "-${pgid}" 2>/dev/null || true
  for ((i = 0; i < 20; i++)); do
    kill -0 -- "-${pgid}" 2>/dev/null || return 0
    sleep 0.1
  done
  echo "[warn] Enroot process group $pgid still has live members" >&2
  return 1
}

cleanup_worker() {
  set +e
  if [[ -n "${WORKER_PGID:-}" ]]; then
    stop_worker_process_group "$WORKER_PGID" || true
    WORKER_PGID=""
  fi
  if [[ -n "${WORKER_INST:-}" ]]; then
    enroot remove -f "$WORKER_INST" >/dev/null 2>&1 \
      || rm -rf -- "${ENROOT_DATA_PATH}/${WORKER_INST}" 2>/dev/null \
      || true
    if [[ -d "${ENROOT_DATA_PATH}/${WORKER_INST}" ]]; then
      echo "[warn] Enroot rootfs still exists after cleanup: ${ENROOT_DATA_PATH}/${WORKER_INST}" >&2
    fi
    WORKER_INST=""
  fi
  if [[ -n "${WORKER_PGDATA:-}" ]]; then
    stop_isolated_pg "$WORKER_PGDATA"
    WORKER_PGDATA=""
  fi
  if [[ -n "${WORKER_PG_RUNTIME:-}" && "$WORKER_PG_RUNTIME" == "$PG_RUNTIME_ROOT"/* ]]; then
    rm -rf -- "$WORKER_PG_RUNTIME"
    WORKER_PG_RUNTIME=""
  fi
}

cleanup_run_artifacts() {
  set +e
  local pid pgid path

  while read -r pid; do
    [[ "$pid" =~ ^[0-9]+$ ]] || continue
    pgid="$(ps -o pgid= -p "$pid" 2>/dev/null | tr -d ' ')"
    stop_worker_process_group "$pgid" || true
  done < <(pgrep -u "$USER" -f "agent-${RUN_ID}-" 2>/dev/null || true)

  if [[ -d "${ENROOT_DATA_PATH:-}" ]]; then
    while read -r path; do
      [[ "$path" == "$ENROOT_DATA_PATH"/agent-"$RUN_ID"-* ]] || continue
      enroot remove -f "$(basename "$path")" >/dev/null 2>&1 || true
      rm -rf -- "$path" 2>/dev/null || true
    done < <(find "$ENROOT_DATA_PATH" -mindepth 1 -maxdepth 1 -type d \
      -name "agent-${RUN_ID}-*" -print 2>/dev/null)
    rmdir "$ENROOT_DATA_PATH" 2>/dev/null || true
  fi

  if [[ -d "${PG_RUNTIME_ROOT:-}" ]]; then
    while read -r path; do
      [[ "$path" == "$PG_RUNTIME_ROOT"/"$RUN_ID"_* ]] || continue
      rm -rf -- "$path" 2>/dev/null || true
    done < <(find "$PG_RUNTIME_ROOT" -mindepth 1 -maxdepth 1 -type d \
      -name "${RUN_ID}_*" -print 2>/dev/null)
    rmdir "$PG_RUNTIME_ROOT" 2>/dev/null || true
  fi

  rmdir "$PG_PORT_LEASE_ROOT" 2>/dev/null || true
  rmdir "${ENROOT_TEMP_PATH:-}" "${ENROOT_RUNTIME_PATH:-}" "${ENROOT_CACHE_PATH:-}" 2>/dev/null || true
}

generate_audit_html() {
  local case_dir="$1"
  [[ "$AUTO_AUDIT_HTML" == "1" ]] || return 0
  [[ -f "${PROJECT_ROOT}/scripts/audit_html_gen.py" ]] || return 0
  echo "=== generate audit.html ==="
  if python3 "${PROJECT_ROOT}/scripts/audit_html_gen.py" "$case_dir"; then
    [[ -f "${case_dir}/audit.html" ]] && echo "[audit] wrote ${case_dir}/audit.html"
  else
    echo "[warn] audit.html generation failed for ${case_dir}" >&2
  fi
}

run_one_task() {
  local TASK="$1"
  local SLOT="$2"
  local PGPORT="$3"
  local SAFE TASK_HASH TASK_ID PG_RUNTIME PGDATA PGSOCKET INST OUTDIR TASK_LOG START_TS END_TS RC STATUS probe_count ENROOT_LAUNCH_PID
  SAFE="$(printf '%s' "$TASK" | tr -cs 'A-Za-z0-9_.-' '-')"
  TASK_HASH="$(echo -n "$TASK-$RUN_ID-$SLOT-$$" | md5sum | cut -c1-8)"
  TASK_ID="${SAFE}-${TASK_HASH}"
  # Keep PostgreSQL's mutable data and socket on node-local storage. The short
  # hash also keeps the Unix socket path below PostgreSQL's ~108-byte limit.
  PG_RUNTIME="${PG_RUNTIME_ROOT}/${RUN_ID}_${SLOT}_${TASK_HASH}"
  PGDATA="${PG_RUNTIME}/data"
  PGSOCKET="${PG_RUNTIME}/socket"
  INST="agent-${RUN_ID}-${SLOT}-${TASK_HASH}-${SAFE}"
  OUTDIR="${DUMP_ROOT}/${TASK}/${RUN_ID}_slot${SLOT}"
  TASK_LOG="${OUTDIR}/run.log"
  mkdir -p "$OUTDIR" "$PG_RUNTIME"

  WORKER_PG_RUNTIME="$PG_RUNTIME"
  WORKER_PGDATA="$PGDATA"
  WORKER_INST=""
  WORKER_TASK="$TASK"
  WORKER_OUTDIR="$OUTDIR"
  WORKER_PGPORT="$PGPORT"
  WORKER_SUMMARY_WRITTEN=0

  {
    echo "PGPORT=$PGPORT"
    echo "PGDATA=$PGDATA"
    echo "PGSOCKET=$PGSOCKET"
    echo "TASK_ID=$TASK_ID"
  } >"${OUTDIR}/isolation.env"

  START_TS=$(date +%s)
  WORKER_START_TS="$START_TS"
  log "START  $TASK  (slot=$SLOT pg_port=$PGPORT pgdata=$PGDATA)"

  {
    echo "=== isolated PG on 127.0.0.1:$PGPORT (socket=$PGSOCKET) ==="
    local pg_attempts=2
    local attempt=1
    local pg_started=0
    while (( attempt <= pg_attempts )); do
      if start_isolated_pg "$PGDATA" "$PGSOCKET" "$PGPORT" "${OUTDIR}/postgres.log" "${OUTDIR}/db_check.log"; then
        pg_started=1
        break
      fi
      if (( attempt < pg_attempts )); then
        echo "[warn] PostgreSQL start failed (attempt $attempt), retrying with new port..."
        # Clean up the failed PGDATA so initdb can rerun cleanly.
        rm -rf -- "$PGDATA"
        if allocate_free_port; then
          PGPORT="$ALLOCATED_PG_PORT"
          WORKER_PGPORT="$PGPORT"
          echo "PGPORT=$PGPORT" > "${OUTDIR}/isolation.env"
          log "[retry] $TASK moved to pg_port=$PGPORT"
        else
          echo "[error] unable to allocate retry port for $TASK"
          break
        fi
      fi
      attempt=$((attempt + 1))
    done
    if (( pg_started == 0 )); then
      echo "[error] failed to start isolated postgres after $pg_attempts attempts"
      append_summary "${TASK},pg_fail,1,${OUTDIR},${PGPORT},$(( $(date +%s) - START_TS ))"
      return 1
    fi
    # Legacy evaluators frequently hard-code port 5432.  Give this task's
    # private Unix socket a 5432 alias, then mount only this socket directory
    # into its Enroot rootfs.  The alias preserves isolation without Python
    # monkeypatching or a shared host TCP port.
    ln -sfn ".s.PGSQL.${PGPORT}" "${PGSOCKET}/.s.PGSQL.5432"

    if [[ "$PG_TEST_ONLY" == 1 ]]; then
      echo "=== PostgreSQL isolation probe ==="
      if ! PGPASSWORD="$PGPASSWORD" psql -h 127.0.0.1 -p "$PGPORT" -U "$PGUSER" -d "$PGDATABASE" \
          -v ON_ERROR_STOP=1 -v probe="$TASK_ID" <<'SQL'
CREATE TABLE IF NOT EXISTS public.toolathlon_isolation_probe (
  probe text PRIMARY KEY
);
INSERT INTO public.toolathlon_isolation_probe(probe) VALUES (:'probe');
SELECT probe FROM public.toolathlon_isolation_probe ORDER BY probe;
SQL
      then
        append_summary "${TASK},pg_test_fail,1,${OUTDIR},${PGPORT},$(( $(date +%s) - START_TS ))"
        return 1
      fi
      sleep 5
      probe_count="$(PGPASSWORD="$PGPASSWORD" psql -h 127.0.0.1 -p "$PGPORT" -U "$PGUSER" -d "$PGDATABASE" \
        -tAc 'SELECT count(*) FROM public.toolathlon_isolation_probe')"
      if [[ "$probe_count" != 1 ]]; then
        echo "[error] isolation probe saw $probe_count task markers instead of 1"
        append_summary "${TASK},pg_test_fail,1,${OUTDIR},${PGPORT},$(( $(date +%s) - START_TS ))"
        return 1
      fi
      END_TS=$(date +%s)
      append_summary "${TASK},pg_test_success,0,${OUTDIR},${PGPORT},$((END_TS - START_TS))"
      log "DONE   $TASK -> pg_test_success ($((END_TS - START_TS))s, port=$PGPORT)"
      return 0
    fi

    echo "=== create agent rootfs $INST ==="
    WORKER_INST="$INST"
    if ! make_agent_rootfs "$INST"; then
      echo "[error] failed to create Enroot rootfs $INST"
      append_summary "${TASK},rootfs_fail,1,${OUTDIR},${PGPORT},$(( $(date +%s) - START_TS ))"
      return 1
    fi
    # The sqsh/template is a build snapshot. Refresh the lightweight runtime
    # code needed by the evaluator without rebuilding 2.6 GB.
    if ! mkdir -p "${ENROOT_DATA_PATH}/${INST}/run/toolathlon_pg" \
      || ! cp -f "${PROJECT_ROOT}/utils/mcp/tool_servers.py" \
        "${ENROOT_DATA_PATH}/${INST}/workspace/utils/mcp/tool_servers.py" \
      || ! cp -f "${PROJECT_ROOT}/utils/roles/task_agent.py" \
        "${ENROOT_DATA_PATH}/${INST}/workspace/utils/roles/task_agent.py" \
      || ! cp -f "${PROJECT_ROOT}/utils/aux_tools/basic.py" \
        "${ENROOT_DATA_PATH}/${INST}/workspace/utils/aux_tools/basic.py" \
      || ! cp -f "${PROJECT_ROOT}/utils/api_model/model_provider.py" \
        "${ENROOT_DATA_PATH}/${INST}/workspace/utils/api_model/model_provider.py" \
      || ! mkdir -p "${ENROOT_DATA_PATH}/${INST}/workspace/tasks/finalpool" \
      || ! rsync -a --delete "${PROJECT_ROOT}/tasks/finalpool/" \
        "${ENROOT_DATA_PATH}/${INST}/workspace/tasks/finalpool/" \
      || ! cp -f "${PROJECT_ROOT}/local_servers/emails-mcp/src/emails_mcp/server.py" \
        "${ENROOT_DATA_PATH}/${INST}/opt/local_servers/emails-mcp/src/emails_mcp/server.py" \
      || ! cp -f "${PROJECT_ROOT}/local_servers/excel-mcp-server/src/excel_mcp/__main__.py" \
        "${ENROOT_DATA_PATH}/${INST}/opt/local_servers/excel-mcp-server/src/excel_mcp/__main__.py" \
      || ! cp -f "${PROJECT_ROOT}/local_servers/mcp-canvas-lms/build/index.js" \
        "${ENROOT_DATA_PATH}/${INST}/opt/local_servers/mcp-canvas-lms/build/index.js" \
      || ! cp -f "${PROJECT_ROOT}/local_servers/mcp-canvas-lms/build/client.js" \
        "${ENROOT_DATA_PATH}/${INST}/opt/local_servers/mcp-canvas-lms/build/client.js" \
      || ! cp -f "${PROJECT_ROOT}/local_servers/mcp-canvas-lms/build/pg-canvas-router.js" \
        "${ENROOT_DATA_PATH}/${INST}/opt/local_servers/mcp-canvas-lms/build/pg-canvas-router.js" \
      || ! cp -f "${PROJECT_ROOT}/local_servers/notion-mcp-server/scripts/notion-openapi.json" \
        "${ENROOT_DATA_PATH}/${INST}/opt/local_servers/notion-mcp-server/scripts/notion-openapi.json" \
      || ! cp -f "${PROJECT_ROOT}/local_servers/Office-Word-MCP-Server/word_document_server/main.py" \
        "${ENROOT_DATA_PATH}/${INST}/opt/local_servers/Office-Word-MCP-Server/word_document_server/main.py" \
      || ! cp -f "${PROJECT_ROOT}/configs/mcp_servers/word.yaml" \
        "${ENROOT_DATA_PATH}/${INST}/workspace/configs/mcp_servers/word.yaml" \
      || ! cp -f "${PROJECT_ROOT}/configs/mcp_servers/yahoo-finance.yaml" \
        "${ENROOT_DATA_PATH}/${INST}/workspace/configs/mcp_servers/yahoo-finance.yaml"; then
      echo "[error] failed to refresh Enroot runtime harness"
      append_summary "${TASK},rootfs_fail,1,${OUTDIR},${PGPORT},$(( $(date +%s) - START_TS ))"
      return 1
    fi
    # kimi harness + CLI distribution (excl. host pycache)
    if ! rm -rf "${ENROOT_DATA_PATH}/${INST}/workspace/kimi_harness" \
      || ! cp -a "${PROJECT_ROOT}/kimi_harness" "${ENROOT_DATA_PATH}/${INST}/workspace/kimi_harness" \
      || ! rm -rf "${ENROOT_DATA_PATH}/${INST}/workspace/kimi_harness/__pycache__" \
      || ! rm -rf "${ENROOT_DATA_PATH}/${INST}/opt/kimi-code" \
      || ! mkdir -p "${ENROOT_DATA_PATH}/${INST}/opt/kimi-code" \
      || ! cp -a "${KIMI_DIST}/." "${ENROOT_DATA_PATH}/${INST}/opt/kimi-code/"; then
      echo "[error] failed to stage kimi harness/CLI into rootfs"
      append_summary "${TASK},rootfs_fail,1,${OUTDIR},${PGPORT},$(( $(date +%s) - START_TS ))"
      return 1
    fi
    printf '#!/bin/bash\nexec /usr/bin/node /opt/kimi-code/dist/main.mjs "$@"\n' \
      > "${ENROOT_DATA_PATH}/${INST}/usr/local/bin/kimi"
    chmod +x "${ENROOT_DATA_PATH}/${INST}/usr/local/bin/kimi"

    # Bug B fix: v2 linkAttemptSignals missing `task.timeout <= 0` guard.
    # When task.timeout=0, setTimeout(fn,0) fires immediately, aborting spawn
    # and producing "not_started / Subagent timed out" for every AgentSwarm item.
    # v1 and v2 source both have this guard; only bundle 0.33.0's v2 compile missed it.
    # sed is idempotent: if bundle already fixed, grep won't match and patch is skipped.
    local _kimi_main="${ENROOT_DATA_PATH}/${INST}/opt/kimi-code/dist/main.mjs"
    if [[ -f "$_kimi_main" ]] && grep -q 'task\.timeout === void 0 ? void 0 : setTimeout' "$_kimi_main"; then
      sed -i 's/task\.timeout === void 0 ? void 0 : setTimeout/task.timeout === void 0 || task.timeout <= 0 ? void 0 : setTimeout/g' "$_kimi_main"
      echo "[bug-b-fix] patched v2 linkAttemptSignals (task.timeout <= 0 guard)"
    fi
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
      return 1
    fi

    ENV_ARGS=(
      -e "PGHOST=/run/toolathlon_pg"
      -e "PG_HOST=/run/toolathlon_pg"
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
      -e "MODEL_API_KEY=${MODEL_API_KEY}"
      -e "MODEL_API_URL=${MODEL_API_URL}"
      -e "KIMI_MAX_CONTEXT=${KIMI_MAX_CONTEXT:-262144}"
      -e "KIMI_TASK_TIMEOUT_S=${KIMI_TASK_TIMEOUT_S:-7200}"
      -e "no_proxy=${no_proxy}"
      -e "NO_PROXY=${NO_PROXY}"
      -e "MCP_STDIO_TIMEOUT_MIN=${MCP_STDIO_TIMEOUT_MIN:-90}"
    )
    if [[ -v KIMI_SUBAGENTS ]]; then
      ENV_ARGS+=(-e "KIMI_SUBAGENTS=${KIMI_SUBAGENTS}")
    fi
    [[ -n "${KIMI_EXAMPLES_FILE:-}" ]] && ENV_ARGS+=(-e "KIMI_EXAMPLES_FILE=${KIMI_EXAMPLES_FILE}")
    [[ -n "${KIMI_COORDINATION_FILE:-}" ]] && ENV_ARGS+=(-e "KIMI_COORDINATION_FILE=${KIMI_COORDINATION_FILE}")
    [[ -n "${MODEL_N:-}" ]] && ENV_ARGS+=(-e "MODEL_N=${MODEL_N}")
    # Note: no http_proxy passthrough — the eval endpoint is reached directly
    # (no_proxy covers it) and MCP servers stay fully local.

    # WooCommerce tasks need the PG REST backend on :8081 before MCP connects.
    # Enroot uses the host network namespace, so 127.0.0.1:8081 is SHARED across
    # all tasks/smokes. Never reuse a pre-existing listener — it may point at
    # another task's (or a dead smoke) Postgres. Always (re)bind to THIS task's PG.
    # NOTE: do NOT use `pkill -f pg-rest-server` here — the pattern appears in
    # this bash -c cmdline and would SIGTERM ourselves (exit 143).
    WC_START='
if [[ -f /opt/local_servers/woocommerce-mcp/dist/services/pg-rest-server.js ]]; then
  # Free :8081 via port only (safe; does not match this shell cmdline).
  fuser -k 8081/tcp >/dev/null 2>&1 || true
  sleep 0.3
  # Prefer TCP to the isolated PG: unix socket path differs across namespaces,
  # but host-network 127.0.0.1:$PG_PORT is unambiguous.
  nohup env PG_HOST=127.0.0.1 node /opt/local_servers/woocommerce-mcp/dist/services/pg-rest-server.js >/tmp/wc-rest.log 2>&1 &
  for i in $(seq 1 25); do
    if curl -fsS -m 2 http://127.0.0.1:8081/health >/dev/null 2>&1; then
      echo "[wc-rest] up on :8081 -> PG 127.0.0.1:${PG_PORT}"
      break
    fi
    sleep 0.4
  done
fi
'
    echo "=== run kimi_main.py (cwd=/workspace) ==="
    set +e
    # Enroot does not provide daemon lifecycle management.  Run each case in
    # its own session/process group so background services started by task
    # preprocessors (for example `nohup python -m http.server ... &`) remain
    # attributable to this worker even after the main agent process exits.
    # cleanup_worker terminates the whole group before deleting the rootfs,
    # preventing orphan listeners and NFS .nfs* remnants.
    setsid enroot start -r -w -m "${PGSOCKET}:/run/toolathlon_pg" "${ENV_ARGS[@]}" "$INST" \
      /bin/bash -c "${WC_START}cd /workspace && exec /opt/venv/bin/python3 kimi_harness/kimi_main.py --eval_config /workspace/scripts/eval_config.json --task_dir '${TASK}' --max_steps '${MAX_STEPS}' --debug" &
    ENROOT_LAUNCH_PID=$!
    WORKER_PGID="$(ps -o pgid= -p "$ENROOT_LAUNCH_PID" 2>/dev/null | tr -d ' ')"
    [[ -n "$WORKER_PGID" ]] || WORKER_PGID="$ENROOT_LAUNCH_PID"
    echo "[runner] Enroot launch pid=$ENROOT_LAUNCH_PID process_group=$WORKER_PGID"
    wait "$ENROOT_LAUNCH_PID"
    RC=$?
    set -e

    echo "=== sync dumps ==="
    rsync -a "${ENROOT_DATA_PATH}/${INST}/workspace/dumps/" "$OUTDIR/" || true
    generate_audit_html "$OUTDIR"

    echo "=== cleanup ==="
    cleanup_worker

    END_TS=$(date +%s)
    if [[ $RC -eq 0 ]]; then STATUS=success; else STATUS=failed; fi
    append_summary "${TASK},${STATUS},${RC},${OUTDIR},${PGPORT},$((END_TS - START_TS))"
    log "DONE   $TASK -> $STATUS (exit=$RC, $((END_TS - START_TS))s, port=$PGPORT)"
    return "$RC"
  } >"$TASK_LOG" 2>&1
}

export PROJECT_ROOT RUNTIME_ROOT DUMP_ROOT PARALLEL_ROOT PG_PORT_BASE PG_RUNTIME_ROOT PG_TEST_ONLY INIT_SQL
export AGENT_TEMPLATE AGENT_SQSH ENROOT_DATA_PATH RUN_ID SUMMARY SUMMARY_LOCK KIMI_DIST
export MODEL_NAME MODEL_API_KEY MODEL_API_URL MAX_STEPS
export KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
export PGUSER PGPASSWORD PGDATABASE no_proxy NO_PROXY

need_bins
[[ -d "${PROJECT_ROOT}/kimi_harness" ]] || { echo "[error] missing ${PROJECT_ROOT}/kimi_harness" >&2; exit 1; }
[[ -f "${KIMI_DIST}/dist/main.mjs" ]] || { echo "[error] missing kimi CLI dist: ${KIMI_DIST}" >&2; exit 1; }
[[ -n "${MODEL_API_KEY:-}" && -n "${MODEL_API_URL:-}" && -n "${MODEL_NAME:-}" ]] || {
  echo "[error] MODEL_NAME/MODEL_API_KEY/MODEL_API_URL must be set (config.env)" >&2; exit 1; }

log "=============================================="
log "  Parallel safe eval — kimi-code @ Toolathlon-GYM"
log "  Model:         $MODEL_NAME @ $MODEL_API_URL"
log "  Context:       max_context=${KIMI_MAX_CONTEXT:-262144} task_timeout=${KIMI_TASK_TIMEOUT_S:-7200}s"
log "  Tasks:         ${#TASKS[@]}"
log "  Max concurrent:$MAX_CONCURRENT  (gnho019 recommend 6, soft-cap 8)"
log "  Dump root:     $DUMP_ROOT"
log "  PG port base:  $PG_PORT_BASE"
log "  PG runtime:    $PG_RUNTIME_ROOT (unique data + socket per task)"
[[ "$PG_TEST_ONLY" == 1 ]] && log "  Mode:          PostgreSQL isolation test only"
log "=============================================="

# API check
if command -v curl >/dev/null 2>&1; then
  code="$(curl -sS --connect-timeout 8 -o /dev/null -w '%{http_code}' \
    -H "Authorization: Bearer ${MODEL_API_KEY}" \
    "${MODEL_API_URL%/}/v1/models" || true)"
  log "API /v1/models -> HTTP $code"
fi

# FIFO semaphore on local tmpfs — NFS named pipes (.nfs stubs under PARALLEL_ROOT)
# drop worker→parent token writes, which deadlocks the pool after the first failures.
FIFO="/dev/shm/toolathlon_sem_${RUN_ID}_$$"
rm -f "$FIFO"
mkfifo "$FIFO"
exec 3<>"$FIFO"
rm -f "$FIFO"
cleanup_fifo() {
  exec 3>&- 2>/dev/null || true
  rm -f "$FIFO" 2>/dev/null || true
  rm -rf -- "$TASK_LOCK_ROOT" 2>/dev/null || true
}
cleanup_parent() {
  local pid
  trap - HUP INT TERM
  for pid in "${PIDS[@]:-}"; do
    [[ -n "$pid" ]] && kill -TERM "$pid" 2>/dev/null || true
  done
  for pid in "${PIDS[@]:-}"; do
    [[ -n "$pid" ]] && wait "$pid" 2>/dev/null || true
  done
  cleanup_fifo
  release_port_leases
  cleanup_run_artifacts
  rmdir "$SUMMARY_LOCK" 2>/dev/null || true
}
trap cleanup_parent EXIT
trap 'exit 129' HUP
trap 'exit 130' INT TERM
for ((i = 0; i < MAX_CONCURRENT; i++)); do
  echo >&3
done

PIDS=()
SLOT=0
FAILED=0
NEXT_PG_PORT="$PG_PORT_BASE"
ALLOCATED_PG_PORT=""
for TASK in "${TASKS[@]}"; do
  read -u 3
  if ! claim_task "$TASK"; then
    log "[skip] $TASK already claimed by another slot (duplicate dispatch)"
    echo >&3
    continue
  fi
  if ! allocate_free_port; then
    log "[error] unable to allocate an isolated PostgreSQL port for $TASK"
    FAILED=$((FAILED + 1))
    echo >&3
    break
  fi
  PGPORT="$ALLOCATED_PG_PORT"
  (
    WORKER_PGDATA=""
    WORKER_PG_RUNTIME=""
    WORKER_INST=""
    WORKER_PGID=""
    WORKER_TASK=""
    WORKER_OUTDIR=""
    WORKER_PGPORT=""
    WORKER_START_TS=""
    WORKER_SUMMARY_WRITTEN=0
    worker_exit() {
      rc=$?
      trap - EXIT
      if [[ -n "${WORKER_TASK:-}" && "${WORKER_SUMMARY_WRITTEN:-0}" != 1 ]]; then
        now_ts="$(date +%s)"
        start_ts="${WORKER_START_TS:-$now_ts}"
        append_summary "${WORKER_TASK},interrupted,${rc:-1},${WORKER_OUTDIR:-},${WORKER_PGPORT:-},$((now_ts - start_ts))"
      fi
      cleanup_worker
      echo >&3
      exit "$rc"
    }
    trap worker_exit EXIT
    trap 'exit 129' HUP
    trap 'exit 130' INT TERM
    run_one_task "$TASK" "$SLOT" "$PGPORT"
  ) &
  PIDS+=($!)
  SLOT=$((SLOT + 1))
done

for pid in "${PIDS[@]}"; do
  wait "$pid" || FAILED=$((FAILED + 1))
done
cleanup_fifo
release_port_leases
cleanup_run_artifacts
trap - EXIT HUP INT TERM

log "=============================================="
log "  Finished. Failed workers: $FAILED / ${#TASKS[@]}"
log "  Summary: $SUMMARY"
log "=============================================="
# Print summary table
if [[ -f "$SUMMARY" ]]; then
  column -t -s, "$SUMMARY" 2>/dev/null || cat "$SUMMARY"
fi

exit "$FAILED"
