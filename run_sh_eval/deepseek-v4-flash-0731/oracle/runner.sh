#!/usr/bin/env bash
# Per-case PostgreSQL + Enroot isolation, adapted from the established evaluator.
# Called only by launch.py; all startup configuration is supplied explicitly.
set -euo pipefail
SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
: "${PROJECT_ROOT:?}" "${RUNTIME_ROOT:?}" "${DUMP_ROOT:?}" "${RUN_ID:?}"
: "${AGENT_SQSH:?}" "${KIMI_DIST:?}" "${MAX_CONCURRENT:?}" "${PG_PORT_BASE:?}"
INIT_SQL="$PROJECT_ROOT/db/init.sql.gz"
TASKS=("$@")
[[ ${#TASKS[@]} == 166 ]] || { echo "Expected all 166 MS-Eval cases" >&2; exit 2; }
cd "$PROJECT_ROOT"
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
  (( PG_PORT_BASE + 10000 + MAX_CONCURRENT - 1 <= 65535 )) || {
    echo "[error] mock-web ports (PG_PORT_BASE + 10000 + slot) would exceed 65535; lower PG_PORT_BASE (max $((65535 - 9999 - MAX_CONCURRENT + 1 ))) for MAX_CONCURRENT=${MAX_CONCURRENT}" >&2
    exit 1
  }
  grep -q "Results are paginated and include summary totals" \
    "${PROJECT_ROOT}/local_servers/mcp-canvas-lms/build/index.js" || {
      echo "[error] Canvas MCP build is stale; rebuild local_servers/mcp-canvas-lms before running." >&2
      exit 1
    }
  grep -q "Using PgCanvasRouter" \
    "${PROJECT_ROOT}/local_servers/mcp-canvas-lms/build/client.js" || {
      echo "[error] Canvas MCP client build is stale; rebuild local_servers/mcp-canvas-lms before running." >&2
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
  awk -v p="$hexport" '
    $2 ~ ":" p && $4 == "0A" { found=1; exit }
    END { exit !found }
  ' /proc/net/tcp /proc/net/tcp6 2>/dev/null
}

TASK_LOCK_ROOT="/dev/shm/toolathlon_task_locks_${UID}_${RUN_ID}"
mkdir -p "$TASK_LOCK_ROOT"
chmod 700 "$TASK_LOCK_ROOT" 2>/dev/null || true

claim_task() {
  local task="$1"
  local safe
  safe="$(printf '%s' "$task" | tr -cs 'A-Za-z0-9_.-' '-')"
  mkdir "$TASK_LOCK_ROOT/$safe" 2>/dev/null
}

MOCK_PORT_ROOT="/dev/shm/toolathlon_mock_ports_${UID}"
mkdir -p "$MOCK_PORT_ROOT" 2>/dev/null || true
mock_ports_of_task() {
  local task="${1:-}"
  local pp="$PROJECT_ROOT/tasks/finalpool/$task/preprocess/main.py"
  [[ -f "$pp" ]] || return 0
  grep -oiE '(^|[^a-z])_?port[^=]{0,3}=[[:space:]]*30[0-9]{3}' "$pp" 2>/dev/null \
    | grep -oE '30[0-9]{3}' | sort -u
}
acquire_mock_ports() {
  local task="${1:-}"
  local port lock
  ACQUIRED_MOCK_LOCKS=""
  local taken=()
  while read -r port; do
    [[ -n "$port" ]] || continue
    lock="$MOCK_PORT_ROOT/$port"
    if ! mkdir "$lock" 2>/dev/null; then
      local owner
      owner="$(sed -n '1p' "$lock/owner" 2>/dev/null || true)"
      if [[ "$owner" =~ ^[0-9]+$ ]] && ! kill -0 "$owner" 2>/dev/null; then
        rm -rf -- "$lock" 2>/dev/null || true
        mkdir "$lock" 2>/dev/null || { release_mock_ports "${taken[*]:-}"; return 1; }
      else
        local l
        release_mock_ports "${taken[*]:-}"
        return 1
      fi
    fi
    echo "$$" >"$lock/owner"
    taken+=("$lock")
  done < <(mock_ports_of_task "$task")
  ACQUIRED_MOCK_LOCKS="${taken[*]:-}"
  return 0
}
release_mock_ports() {
  local lock
  for lock in $1; do
    rm -f "$lock/owner" 2>/dev/null || true
    rmdir "$lock" 2>/dev/null || true
  done
}

DISPATCH_BUDGET_S="${DISPATCH_BUDGET_S:-$(( ${KIMI_TASK_TIMEOUT_S:-7200} + 1200 ))}"
_walltime_remaining_s() {
  if [[ -n "${SLURM_JOB_END_TIME:-}" && "${SLURM_JOB_END_TIME}" =~ ^[0-9]+$ ]]; then
    echo $(( SLURM_JOB_END_TIME - $(date +%s) ))
  else
    echo 999999999  # non-slurm (local) run: never block dispatch
  fi
}

allocate_free_port() {
  local lease owner
  while (( NEXT_PG_PORT <= 55535 )); do
    lease="${PG_PORT_LEASE_ROOT}/${NEXT_PG_PORT}"
    if mkdir "$lease" 2>/dev/null; then
      echo "$$" >"$lease/owner_pid"
      if port_is_listening "$NEXT_PG_PORT" || port_is_listening "$((NEXT_PG_PORT + 10000))"; then
        rm -rf -- "$lease"
      else
        ALLOCATED_PG_PORT="$NEXT_PG_PORT"
        PORT_LEASES+=("$lease")
        NEXT_PG_PORT=$((NEXT_PG_PORT + 1))
        return 0
      fi
    else
      owner="$(sed -n '1p' "$lease/owner_pid" 2>/dev/null || true)"
      if [[ "$owner" =~ ^[0-9]+$ ]] && ! kill -0 "$owner" 2>/dev/null && ! port_is_listening "$NEXT_PG_PORT" && ! port_is_listening "$((NEXT_PG_PORT + 10000))"; then
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
  if ! pg_ctl -D "$pgdata" -l "$logfile" -o "-k ${socket_dir} -p ${port}" start >/dev/null; then
    echo "[error] PostgreSQL failed to start for PGDATA=$pgdata port=$port" >&2
    return 1
  fi

  local i=0
  while (( i < 60 )); do
    if env -u PGDATABASE pg_isready -h "$socket_dir" -p "$port" -U "$PGUSER" -d postgres >/dev/null 2>&1; then
      break
    fi
    sleep 1
    i=$((i + 1))
  done
  env -u PGDATABASE pg_isready -h "$socket_dir" -p "$port" -U "$PGUSER" -d postgres >/dev/null 2>&1 || return 1

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

  env -u PGDATABASE psql -h "$socket_dir" -p "$port" -U "$PGUSER" -d postgres -v ON_ERROR_STOP=1 \
    -c "ALTER USER ${PGUSER} WITH PASSWORD '${PGPASSWORD}';" >/dev/null
  if [[ "$(env -u PGDATABASE psql -h "$socket_dir" -p "$port" -U "$PGUSER" -d postgres -tAc \
      "SELECT 1 FROM pg_database WHERE datname='${PGDATABASE}'")" != 1 ]]; then
    env -u PGDATABASE createdb -h "$socket_dir" -p "$port" -U "$PGUSER" "$PGDATABASE"
  fi

  env -u PGDATABASE psql -h "$socket_dir" -p "$port" -U "$PGUSER" -d postgres -tAc \
    "SELECT 1 FROM pg_database WHERE datname='${PGDATABASE}'" | grep -q 1 || return 1

  if ! gunzip -c "$INIT_SQL" | env -u PGDATABASE psql -h "$socket_dir" -p "$port" -U "$PGUSER" -d "$PGDATABASE" \
    -v ON_ERROR_STOP=1 >/dev/null; then
    echo "[error] failed to load clean Toolathlon schema on port $port" >&2
    return 1
  fi

  if ! PGPASSWORD="$PGPASSWORD" psql -h 127.0.0.1 -p "$port" -U "$PGUSER" -d "$PGDATABASE" -v ON_ERROR_STOP=1 <<'SQL' >/dev/null
ALTER TABLE email.sent_log DROP CONSTRAINT IF EXISTS sent_log_message_id_fkey;
ALTER TABLE email.sent_log ADD CONSTRAINT sent_log_message_id_fkey
  FOREIGN KEY (message_id) REFERENCES email.messages(id) ON DELETE CASCADE;
SQL
  then
    echo "[error] failed to apply email.sent_log foreign-key fix on port $port" >&2
    return 1
  fi

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
  if [[ -d "$AGENT_TEMPLATE" ]]; then
    rm -rf "${ENROOT_DATA_PATH}/${name}"
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

run_one_task() {
  local TASK="$1"
  local SLOT="$2"
  local PGPORT="$3"
  local SAFE TASK_HASH TASK_ID PG_RUNTIME PGDATA PGSOCKET INST OUTDIR TASK_LOG START_TS END_TS RC STATUS probe_count ENROOT_LAUNCH_PID WCPORT
  SAFE="$(printf '%s' "$TASK" | tr -cs 'A-Za-z0-9_.-' '-')"
  TASK_HASH="$(echo -n "$TASK-$RUN_ID-$SLOT-$$" | md5sum | cut -c1-8)"
  TASK_ID="${SAFE}-${TASK_HASH}"
  PG_RUNTIME="${PG_RUNTIME_ROOT}/${RUN_ID}_${SLOT}_${TASK_HASH}"
  PGDATA="${PG_RUNTIME}/data"
  PGSOCKET="${PG_RUNTIME}/socket"
  WCPORT=$((PGPORT + 10000))
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
    if ! start_isolated_pg "$PGDATA" "$PGSOCKET" "$PGPORT" "${OUTDIR}/postgres.log" "${OUTDIR}/db_check.log"; then
      append_summary "${TASK},pg_fail,1,${OUTDIR},${PGPORT},$(( $(date +%s) - START_TS ))"
      return 1
    fi
    ln -sfn ".s.PGSQL.${PGPORT}" "${PGSOCKET}/.s.PGSQL.5432"

    echo "=== create agent rootfs $INST ==="
    WORKER_INST="$INST"
    if ! make_agent_rootfs "$INST"; then
      echo "[error] failed to create Enroot rootfs $INST"
      append_summary "${TASK},rootfs_fail,1,${OUTDIR},${PGPORT},$(( $(date +%s) - START_TS ))"
      return 1
    fi
    local src dst rootfs="${ENROOT_DATA_PATH}/${INST}"
    mkdir -p "$rootfs/run/toolathlon_pg" "$rootfs/workspace/utils/evaluation" "$rootfs/workspace/tasks/finalpool"
    while read -r src dst; do
      mkdir -p "$(dirname "$rootfs/$dst")"
      cp -f "$PROJECT_ROOT/$src" "$rootfs/$dst" || return 1
    done <<'OVERLAYS'
utils/mcp/tool_servers.py workspace/utils/mcp/tool_servers.py
utils/roles/task_agent.py workspace/utils/roles/task_agent.py
utils/aux_tools/basic.py workspace/utils/aux_tools/basic.py
utils/api_model/model_provider.py workspace/utils/api_model/model_provider.py
utils/data_structures/task_config.py workspace/utils/data_structures/task_config.py
local_servers/emails-mcp/src/emails_mcp/server.py opt/local_servers/emails-mcp/src/emails_mcp/server.py
local_servers/excel-mcp-server/src/excel_mcp/__main__.py opt/local_servers/excel-mcp-server/src/excel_mcp/__main__.py
local_servers/mcp-google-sheets/src/mcp_google_sheets/server.py opt/local_servers/mcp-google-sheets/src/mcp_google_sheets/server.py
local_servers/mcp-google-sheets/src/mcp_google_sheets/pg_services.py opt/local_servers/mcp-google-sheets/src/mcp_google_sheets/pg_services.py
local_servers/mcp-canvas-lms/build/index.js opt/local_servers/mcp-canvas-lms/build/index.js
local_servers/mcp-canvas-lms/build/client.js opt/local_servers/mcp-canvas-lms/build/client.js
local_servers/mcp-canvas-lms/build/pg-canvas-router.js opt/local_servers/mcp-canvas-lms/build/pg-canvas-router.js
local_servers/notion-mcp-server/scripts/notion-openapi.json opt/local_servers/notion-mcp-server/scripts/notion-openapi.json
local_servers/notion-mcp-server/bin/cli.mjs opt/local_servers/notion-mcp-server/bin/cli.mjs
local_servers/notion-mcp-server/build/src/openapi-mcp-server/client/pg-client.js opt/local_servers/notion-mcp-server/build/src/openapi-mcp-server/client/pg-client.js
local_servers/Calendar-Autoauth-MCP-Server/build/pg-calendar.js opt/local_servers/Calendar-Autoauth-MCP-Server/build/pg-calendar.js
local_servers/Calendar-Autoauth-MCP-Server/build/index.js opt/local_servers/Calendar-Autoauth-MCP-Server/build/index.js
local_servers/google-forms-mcp/build/index.js opt/local_servers/google-forms-mcp/build/index.js
local_servers/google-forms-mcp/build/pg-forms.js opt/local_servers/google-forms-mcp/build/pg-forms.js
local_servers/Office-Word-MCP-Server/word_document_server/main.py opt/local_servers/Office-Word-MCP-Server/word_document_server/main.py
configs/mcp_servers/word.yaml workspace/configs/mcp_servers/word.yaml
configs/mcp_servers/yahoo-finance.yaml workspace/configs/mcp_servers/yahoo-finance.yaml
configs/mcp_servers/terminal.yaml workspace/configs/mcp_servers/terminal.yaml
OVERLAYS
    rsync -a --delete "$PROJECT_ROOT/utils/evaluation/" "$rootfs/workspace/utils/evaluation/" || return 1
    rsync -a --delete "$PROJECT_ROOT/tasks/finalpool/" "$rootfs/workspace/tasks/finalpool/" || return 1
    UV_PY_HOST_CACHE="${RUNTIME_ROOT}/uv_python_cache"
    if [[ -d "$UV_PY_HOST_CACHE" ]] && [[ ! -d "${ENROOT_DATA_PATH}/${INST}/opt/uv_python_cache" ]]; then
      mkdir -p "${ENROOT_DATA_PATH}/${INST}/opt/uv_python_cache"
      rsync -a "$UV_PY_HOST_CACHE/" "${ENROOT_DATA_PATH}/${INST}/opt/uv_python_cache/" \
        || echo "[warn] failed to stage uv_python_cache (uv MCP servers may not start)"
    fi
    if grep -q '"woocommerce"' "${PROJECT_ROOT}/tasks/finalpool/${TASK}/task_config.json"; then
      sed -i "s#http://localhost:8081#http://localhost:${WCPORT}#g" \
        "${ENROOT_DATA_PATH}/${INST}/workspace/configs/mcp_servers/woocommerce.yaml" || true
    fi
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
    _node_bin="/usr/local/bin/node"
    [[ -x "${ENROOT_DATA_PATH}/${INST}/usr/local/bin/node" ]] || _node_bin="/usr/bin/node"
    printf '#!/bin/bash\nexec %s /opt/kimi-code/dist/main.mjs "$@"\n' "$_node_bin" \
      > "${ENROOT_DATA_PATH}/${INST}/usr/local/bin/kimi"
    chmod +x "${ENROOT_DATA_PATH}/${INST}/usr/local/bin/kimi"

    local _kimi_main="${ENROOT_DATA_PATH}/${INST}/opt/kimi-code/dist/main.mjs"
    if [[ -f "$_kimi_main" ]] && grep -q 'task\.timeout === void 0 ? void 0 : setTimeout' "$_kimi_main"; then
      sed -i 's/task\.timeout === void 0 ? void 0 : setTimeout/task.timeout === void 0 || task.timeout <= 0 ? void 0 : setTimeout/g' "$_kimi_main"
      echo "[bug-b-fix] patched v2 linkAttemptSignals (task.timeout <= 0 guard)"
    fi
    if [[ -n "${KIMI_TASK_MD_DIR:-}" ]]; then
      if ! rm -rf "${ENROOT_DATA_PATH}/${INST}/workspace/oracle_briefs" \
        || ! cp -a "${KIMI_TASK_MD_DIR}" "${ENROOT_DATA_PATH}/${INST}/workspace/oracle_briefs"; then
        echo "[error] failed to stage oracle task_md_dir into rootfs"
        append_summary "${TASK},rootfs_fail,1,${OUTDIR},${PGPORT},$(( $(date +%s) - START_TS ))"
        return 1
      fi
    fi
    mkdir -p "${ENROOT_DATA_PATH}/${INST}/workspace/dumps"
    mkdir -p "${ENROOT_DATA_PATH}/${INST}/workspace/scripts"
    cp -f "$SCRIPT_DIR/eval_config.json" "${ENROOT_DATA_PATH}/${INST}/workspace/scripts/eval_config.json" || return 1
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
      -e "WC_REST_PORT=${WCPORT}"
    )
    if [[ -v KIMI_SUBAGENTS ]]; then
      ENV_ARGS+=(-e "KIMI_SUBAGENTS=${KIMI_SUBAGENTS}")
    fi
    if [[ -v KIMI_CODE_SUBAGENT ]]; then
      ENV_ARGS+=(-e "KIMI_CODE_SUBAGENT=${KIMI_CODE_SUBAGENT}")
    fi
    if [[ -n "${KIMI_TASK_MD_DIR:-}" ]]; then
      ENV_ARGS+=(-e "KIMI_TASK_MD_DIR=/workspace/oracle_briefs")
    fi
    ENV_ARGS+=(-e "KIMI_PROVIDER_TYPE=${KIMI_PROVIDER_TYPE:-openai}")
    [[ -n "${KIMI_EXAMPLES_FILE:-}" ]] && ENV_ARGS+=(-e "KIMI_EXAMPLES_FILE=${KIMI_EXAMPLES_FILE}")
    [[ -n "${KIMI_COORDINATION_FILE:-}" ]] && ENV_ARGS+=(-e "KIMI_COORDINATION_FILE=${KIMI_COORDINATION_FILE}")
    [[ -n "${MODEL_N:-}" ]] && ENV_ARGS+=(-e "MODEL_N=${MODEL_N}")
    [[ -n "${KIMI_MODEL_TEMPERATURE:-}" ]] && ENV_ARGS+=(-e "KIMI_MODEL_TEMPERATURE=${KIMI_MODEL_TEMPERATURE}")
    [[ -n "${KIMI_MODEL_TOP_P:-}" ]] && ENV_ARGS+=(-e "KIMI_MODEL_TOP_P=${KIMI_MODEL_TOP_P}")
    ENV_ARGS+=(-e "KIMI_MODEL_THINKING_EFFORT=${KIMI_MODEL_THINKING_EFFORT:-}")
    [[ -n "${KIMI_PLAN_FIRST:-}" ]] && ENV_ARGS+=(-e "KIMI_PLAN_FIRST=${KIMI_PLAN_FIRST}")

WC_START='
if [[ -n "${TASK_NAME:-}" ]] && grep -q "\"woocommerce\"" "/workspace/tasks/finalpool/${TASK_NAME}/task_config.json" && [[ -f /opt/local_servers/woocommerce-mcp/dist/services/pg-rest-server.js ]]; then
  WC_REST_PORT="${WC_REST_PORT:-8081}"
  sleep 0.3
  nohup env PG_HOST=127.0.0.1 node /opt/local_servers/woocommerce-mcp/dist/services/pg-rest-server.js >/tmp/wc-rest.log 2>&1 &
  for i in $(seq 1 25); do
    if curl -fsS -m 2 "http://127.0.0.1:${WC_REST_PORT}/health" >/dev/null 2>&1; then
      echo "[wc-rest] up on :${WC_REST_PORT} -> PG 127.0.0.1:${PG_PORT}"
      break
    fi
    sleep 0.4
  done
fi
'
    echo "=== run kimi_main.py (cwd=/workspace) ==="
    set +e
    setsid enroot start -r -w -m "${PGSOCKET}:/run/toolathlon_pg" "${ENV_ARGS[@]}" "$INST" \
      /bin/bash -c "TASK_NAME='${TASK}'; ${WC_START}cd /workspace && exec /opt/venv/bin/python3 kimi_harness/kimi_main.py --eval_config /workspace/scripts/eval_config.json --task_dir '${TASK}' --max_steps '${MAX_STEPS}' --debug" &
    ENROOT_LAUNCH_PID=$!
    WORKER_PGID="$(ps -o pgid= -p "$ENROOT_LAUNCH_PID" 2>/dev/null | tr -d ' ')"
    [[ -n "$WORKER_PGID" ]] || WORKER_PGID="$ENROOT_LAUNCH_PID"
    echo "[runner] Enroot launch pid=$ENROOT_LAUNCH_PID process_group=$WORKER_PGID"
    wait "$ENROOT_LAUNCH_PID"
    RC=$?
    set -e

        echo "=== sync dumps ==="
        if ! rsync -a "${ENROOT_DATA_PATH}/${INST}/workspace/dumps/" "$OUTDIR/"; then
      append_summary "${TASK},artifact_sync_failed,1,${OUTDIR},${PGPORT},$(( $(date +%s) - START_TS ))"
      return 1
    fi

        END_TS=$(date +%s)
    PREPROC_FAIL_MARKER="$(find "$OUTDIR" -maxdepth 4 -name '.preprocess_failed' -print -quit 2>/dev/null)"
    PROVIDER_INVALID_MARKER="$(find "$OUTDIR" -maxdepth 4 -name '.provider_invalid' -print -quit 2>/dev/null)"
    INFRA_FAIL_MARKER="$(find "$OUTDIR" -maxdepth 4 -name '.infra_failed' -print -quit 2>/dev/null)"
    if [[ -n "$PREPROC_FAIL_MARKER" && -f "$PREPROC_FAIL_MARKER" ]]; then
      STATUS=preprocess_failed
    elif [[ -n "$PROVIDER_INVALID_MARKER" && -f "$PROVIDER_INVALID_MARKER" ]]; then
      STATUS=provider_invalid
    elif [[ -n "$INFRA_FAIL_MARKER" && -f "$INFRA_FAIL_MARKER" ]]; then
      STATUS=infra_failed
    elif [[ $RC -eq 0 ]]; then
      STATUS=success
    else
      STATUS=case_failed
    fi
        append_summary "${TASK},${STATUS},${RC},${OUTDIR},${PGPORT},$((END_TS - START_TS))"
        WORKER_SUMMARY_WRITTEN=1
        log "DONE   $TASK -> $STATUS (exit=$RC, $((END_TS - START_TS))s, port=$PGPORT)"

        echo "=== cleanup ==="
        cleanup_worker

    case "$STATUS" in
      preprocess_failed|provider_invalid|infra_failed) return 1 ;;
    esac
    return 0
  } >"$TASK_LOG" 2>&1
}

export AGENT_TEMPLATE AGENT_SQSH ENROOT_DATA_PATH RUN_ID SUMMARY SUMMARY_LOCK KIMI_DIST
export MODEL_NAME MODEL_API_KEY MODEL_API_URL MAX_STEPS
export KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S KIMI_PROVIDER_TYPE
export PGUSER PGPASSWORD PGDATABASE no_proxy NO_PROXY

need_bins
[[ -d "${PROJECT_ROOT}/kimi_harness" ]] || { echo "[error] missing ${PROJECT_ROOT}/kimi_harness" >&2; exit 1; }
[[ -f "${KIMI_DIST}/dist/main.mjs" ]] || { echo "[error] missing kimi CLI dist: ${KIMI_DIST}" >&2; exit 1; }
[[ -n "${MODEL_API_KEY:-}" && -n "${MODEL_API_URL:-}" && -n "${MODEL_NAME:-}" ]] || {
  echo "[error] MODEL_NAME/MODEL_API_KEY/MODEL_API_URL must be set (config.env)" >&2; exit 1; }

log "=============================================="
log "  Parallel safe eval — kimi-code @ Toolathlon-GYM"
log "  Model:         $MODEL_NAME"
log "  Context:       max_context=${KIMI_MAX_CONTEXT:-262144} task_timeout=${KIMI_TASK_TIMEOUT_S:-7200}s"
log "  Tasks:         ${#TASKS[@]}"
log "  Max concurrent:$MAX_CONCURRENT"
log "  Dump root:     $DUMP_ROOT"
log "  PG port base:  $PG_PORT_BASE"
log "  PG runtime:    $PG_RUNTIME_ROOT (unique data + socket per task)"
log "=============================================="

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
INFRA_FAILED=0
NEXT_PG_PORT="$PG_PORT_BASE"
ALLOCATED_PG_PORT=""
UNPATCHED_SKIPPED=0
START_TS=$(date +%s)
for TASK in "${TASKS[@]}"; do
  read -u 3
  if (( $(_walltime_remaining_s) < DISPATCH_BUDGET_S )); then
    log "[walltime] remaining $(_walltime_remaining_s)s < dispatch budget ${DISPATCH_BUDGET_S}s — not dispatching $TASK (job ends cleanly)"
    UNPATCHED_SKIPPED=$((UNPATCHED_SKIPPED + 1))
    append_summary "${TASK},walltime_skipped,1,,,$(( $(date +%s) - START_TS ))"
    INFRA_FAILED=$((INFRA_FAILED + 1))
    echo >&3
    continue
  fi
  if ! claim_task "$TASK"; then
    log "[skip] $TASK already claimed by another slot (duplicate dispatch)"
    echo >&3
    continue
  fi
  _mock_wait=0
  _mock_wait_max="${MOCK_PORT_WAIT_LOOPS:-180}"
  until acquire_mock_ports "$TASK"; do
    sleep 20
    _mock_wait=$((_mock_wait + 1))
    if (( _mock_wait > _mock_wait_max )); then  # cap so a dead lock cannot wedge a slot
      log "[error] $TASK could not acquire its mock ports"
      release_mock_ports "$ACQUIRED_MOCK_LOCKS"
      append_summary "${TASK},mock_port_busy,1,,,0"
      INFRA_FAILED=$((INFRA_FAILED + 1))
      echo >&3
      continue 2
    fi
  done
  MOCK_LOCKS_HELD="$ACQUIRED_MOCK_LOCKS"
  if ! allocate_free_port; then
    log "[error] unable to allocate an isolated PostgreSQL port for $TASK"
    INFRA_FAILED=$((INFRA_FAILED + 1))
    [[ -n "$MOCK_LOCKS_HELD" ]] && release_mock_ports "$MOCK_LOCKS_HELD"
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
    WORKER_MOCK_LOCKS="$MOCK_LOCKS_HELD"
    worker_exit() {
      rc=$?
      trap - EXIT
      if [[ -n "${WORKER_TASK:-}" && "${WORKER_SUMMARY_WRITTEN:-0}" != 1 ]]; then
        now_ts="$(date +%s)"
        start_ts="${WORKER_START_TS:-$now_ts}"
        append_summary "${WORKER_TASK},interrupted,${rc:-1},${WORKER_OUTDIR:-},${WORKER_PGPORT:-},$((now_ts - start_ts))"
      fi
      [[ -n "${WORKER_MOCK_LOCKS:-}" ]] && release_mock_ports "$WORKER_MOCK_LOCKS"
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
  wait "$pid" || INFRA_FAILED=$((INFRA_FAILED + 1))
done
cleanup_fifo
release_port_leases
cleanup_run_artifacts
trap - EXIT HUP INT TERM

log "=============================================="
log "  Finished. Infra-failed workers: $INFRA_FAILED / ${#TASKS[@]}"
(( UNPATCHED_SKIPPED > 0 )) && log "  Walltime-skipped (not dispatched): $UNPATCHED_SKIPPED"
log "  Summary: $SUMMARY"
log "=============================================="
if [[ -f "$SUMMARY" ]]; then
  column -t -s, "$SUMMARY" 2>/dev/null || cat "$SUMMARY"
fi

DISPATCHED=${#TASKS[@]}
SUMMARY_ROWS=0
if [[ -f "$SUMMARY" ]]; then
  SUMMARY_ROWS=$(( $(wc -l < "$SUMMARY") - 1 ))
fi
if (( SUMMARY_ROWS < DISPATCHED )); then
  MISSING=$((DISPATCHED - SUMMARY_ROWS))
  log "[warn] $MISSING task(s) missing from summary (dispatched=$DISPATCHED rows=$SUMMARY_ROWS) — counting as infra-failed"
  INFRA_FAILED=$((INFRA_FAILED + MISSING))
fi

if (( INFRA_FAILED > 0 )); then
  exit 2
fi
exit 0
