#!/bin/bash
# Start/stop shared PostgreSQL for toolathlon_gym evaluation.
#
# Default backend: conda postgresql (reliable on this NFS login node).
# Enroot postgres image is kept for reference but not required.
#
# Usage:
#   bash scripts/enroot_postgres.sh start|stop|status|logs

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
RUNTIME_ROOT="${TOOLATHLON_EVAL_DOCKER_ROOT:-/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers}"

# shellcheck disable=SC1091
source "$RUNTIME_ROOT/env.sh"

PGDATA_HOST="${RUNTIME_ROOT}/pgdata/pg15"
INIT_SQL="${PROJECT_ROOT}/db/init.sql.gz"
PID_FILE="${RUNTIME_ROOT}/logs/postgres.pid"
LOG_FILE="${RUNTIME_ROOT}/logs/postgres.log"
READY_TIMEOUT="${READY_TIMEOUT:-300}"
MARKER="${PGDATA_HOST}/.toolathlon_initialized"

die()  { echo "[error] $*" >&2; exit 1; }
log()  { echo "[$(date +%H:%M:%S)] $*"; }

need_bins() {
  command -v initdb >/dev/null 2>&1 || die "initdb not found — conda activate toolathlon_gym"
  command -v pg_ctl >/dev/null 2>&1 || die "pg_ctl not found — conda activate toolathlon_gym"
  command -v pg_isready >/dev/null 2>&1 || die "pg_isready not found"
  command -v psql >/dev/null 2>&1 || die "psql not found"
}

is_ready() {
  # Probe maintenance DB so we don't require PGDATABASE to already exist.
  PGPASSWORD="${PGPASSWORD}" env -u PGDATABASE pg_isready -h "${PGHOST}" -p "${PGPORT}" -U "${PGUSER}" -d postgres >/dev/null 2>&1
}

cmd_status() {
  if [[ -f "$PID_FILE" ]] && kill -0 "$(cat "$PID_FILE")" 2>/dev/null; then
    echo "postgres process: running (pid=$(cat "$PID_FILE"))"
  elif pg_ctl -D "$PGDATA_HOST" status >/dev/null 2>&1; then
    echo "postgres process: running (pg_ctl)"
  else
    echo "postgres process: not running"
  fi
  if is_ready; then
    echo "pg_isready: accepting connections on ${PGHOST}:${PGPORT}"
  else
    echo "pg_isready: not ready"
  fi
}

cmd_stop() {
  if [[ -d "$PGDATA_HOST" ]] && pg_ctl -D "$PGDATA_HOST" status >/dev/null 2>&1; then
    log "Stopping postgres via pg_ctl ..."
    pg_ctl -D "$PGDATA_HOST" -m fast stop >>"$LOG_FILE" 2>&1 || true
  fi
  if [[ -f "$PID_FILE" ]]; then
    local pid
    pid="$(cat "$PID_FILE")"
    kill "$pid" 2>/dev/null || true
    rm -f "$PID_FILE"
  fi
  log "Stopped."
}

cmd_logs() {
  [[ -f "$LOG_FILE" ]] || die "No log file at $LOG_FILE"
  tail -n "${1:-80}" "$LOG_FILE"
}

bootstrap_if_needed() {
  mkdir -p "$(dirname "$PGDATA_HOST")" "$(dirname "$LOG_FILE")"
  if [[ ! -f "$PGDATA_HOST/PG_VERSION" ]]; then
    log "Initializing new PGDATA at $PGDATA_HOST ..."
    rm -rf "$PGDATA_HOST"
    mkdir -p "$PGDATA_HOST"
    # Local trust for simplicity on single-user login node; password still set for TCP clients.
    initdb -D "$PGDATA_HOST" -U "$PGUSER" --auth-local=trust --auth-host=md5 --encoding=UTF8 --locale=C
    cat >>"$PGDATA_HOST/postgresql.conf" <<EOF
listen_addresses = '127.0.0.1'
port = ${PGPORT}
max_connections = 100
shared_buffers = 256MB
EOF
    # password for TCP
    echo "host all all 127.0.0.1/32 md5" >>"$PGDATA_HOST/pg_hba.conf"
    echo "host all all ::1/128 md5" >>"$PGDATA_HOST/pg_hba.conf"
  fi
}

load_dump_if_needed() {
  if [[ -f "$MARKER" ]]; then
    log "Init dump already applied ($MARKER)."
    return 0
  fi
  [[ -f "$INIT_SQL" ]] || die "Missing $INIT_SQL"
  log "Creating database ${PGDATABASE} and loading init.sql.gz (may take a few minutes) ..."
  # Ensure role password for TCP auth (never let PGDATABASE leak into -d postgres calls)
  env -u PGDATABASE psql -h /tmp -p "$PGPORT" -U "$PGUSER" -d postgres -v ON_ERROR_STOP=1 <<SQL
ALTER USER ${PGUSER} WITH PASSWORD '${PGPASSWORD}';
SQL
  # Create DB if missing
  env -u PGDATABASE psql -h /tmp -p "$PGPORT" -U "$PGUSER" -d postgres -tc \
    "SELECT 1 FROM pg_database WHERE datname='${PGDATABASE}'" | grep -q 1 \
    || env -u PGDATABASE createdb -h /tmp -p "$PGPORT" -U "$PGUSER" "$PGDATABASE"

  # Prefer unix socket during load (faster, trust auth)
  if gunzip -c "$INIT_SQL" | env -u PGDATABASE psql -h /tmp -p "$PGPORT" -U "$PGUSER" -d "$PGDATABASE" -v ON_ERROR_STOP=0 >>"$LOG_FILE" 2>&1; then
    touch "$MARKER"
    log "Init dump loaded."
  else
    # Some dumps create the database themselves / have non-fatal notices
    touch "$MARKER"
    log "Init dump finished (see $LOG_FILE for notices)."
  fi
}

cmd_start() {
  need_bins
  bootstrap_if_needed

  if is_ready; then
    log "Postgres already accepting connections on ${PGHOST}:${PGPORT}"
    load_dump_if_needed
    return 0
  fi

  if pg_ctl -D "$PGDATA_HOST" status >/dev/null 2>&1; then
    log "pg_ctl reports running but TCP not ready yet ..."
  else
    log "Starting postgres (conda) ..."
    # Use /tmp sockets to avoid NFS socket issues
    mkdir -p /tmp
    pg_ctl -D "$PGDATA_HOST" -l "$LOG_FILE" -o "-k /tmp -p ${PGPORT}" start
    # record postmaster pid
    if [[ -f "$PGDATA_HOST/postmaster.pid" ]]; then
      head -1 "$PGDATA_HOST/postmaster.pid" >"$PID_FILE"
    fi
  fi

  local waited=0
  while (( waited < READY_TIMEOUT )); do
    if is_ready || pg_isready -h /tmp -p "$PGPORT" >/dev/null 2>&1; then
      log "Postgres is ready (${waited}s)."
      # Set password + load dump
      load_dump_if_needed
      # Verify TCP with password
      if PGPASSWORD="$PGPASSWORD" pg_isready -h "$PGHOST" -p "$PGPORT" -U "$PGUSER" >/dev/null 2>&1; then
        log "TCP auth path OK (${PGHOST}:${PGPORT})."
        return 0
      fi
      return 0
    fi
    sleep 2
    waited=$((waited + 2))
  done
  echo "---- last 40 log lines ----" >&2
  tail -n 40 "$LOG_FILE" >&2 || true
  die "Postgres not ready within ${READY_TIMEOUT}s"
}

case "${1:-}" in
  start)  cmd_start ;;
  stop)   cmd_stop ;;
  status) cmd_status ;;
  logs)   shift || true; cmd_logs "${1:-80}" ;;
  *)
    echo "Usage: $0 {start|stop|status|logs}" >&2
    exit 1
    ;;
esac
