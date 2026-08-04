#!/bin/bash
# Run a single task in an ephemeral enroot rootfs (replaces scripts/run_containerized.sh)
#
# Prerequisites:
#   1. source toolathlon_gym_eval_dockers/env.sh
#   2. bash scripts/enroot_build_agent.sh
#   3. bash scripts/enroot_postgres.sh start
#
# Usage:
#   MODEL_PLATFORM=openai_compatible MODEL_NAME=... MODEL_API_KEY=... MODEL_API_URL=... \
#     bash scripts/enroot_run_task.sh <task_name> [max_steps]

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
RUNTIME_ROOT="${TOOLATHLON_EVAL_DOCKER_ROOT:-/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers}"

# shellcheck disable=SC1091
source "$RUNTIME_ROOT/env.sh"

TASK="${1:?Usage: $0 <task_name> [max_steps]}"
MAX_STEPS="${2:-100}"
TASK_SOURCE="$PROJECT_ROOT/tasks/finalpool/$TASK"
# DUMP_ROOT lets callers group results, e.g. dumps/deepseek-v4-pro/<task>/<ts>/
DUMPS_DIR="${DUMP_ROOT:-$PROJECT_ROOT/dumps}"
LOCK_FILE="${LOCK_FILE:-$PROJECT_ROOT/dumps/.run.lock}"
TIMESTAMP="$(date +%Y%m%d-%H%M%S)"
SAFE_TASK="$(echo "$TASK" | tr '/' '-')"
INSTANCE_NAME="toolathlon-${SAFE_TASK}-${TIMESTAMP}"
OUTPUT_DIR="$DUMPS_DIR/$TASK/$TIMESTAMP"

die()  { echo "[$(date +%H:%M:%S)] [error] $*" >&2; exit 1; }
log()  { echo "[$(date +%H:%M:%S)] $*"; }
warn() { echo "[$(date +%H:%M:%S)] [warn] $*" >&2; }

[[ -d "$TASK_SOURCE" ]] || die "Task directory not found: $TASK_SOURCE"
[[ -f "$TOOLATHLON_AGENT_SQSH" ]] || die "Agent image missing: $TOOLATHLON_AGENT_SQSH — run scripts/enroot_build_agent.sh"
command -v pg_isready >/dev/null 2>&1 || die "pg_isready not found — activate conda env toolathlon_gym"
PGPASSWORD="${PGPASSWORD}" pg_isready -h "${PGHOST}" -p "${PGPORT}" -U "${PGUSER}" -d "${PGDATABASE}" >/dev/null 2>&1 \
  || die "Postgres not ready. Run: bash scripts/enroot_postgres.sh start"

mkdir -p "$OUTPUT_DIR" "$DUMPS_DIR"

cleanup() {
  log "Cleaning up enroot instance $INSTANCE_NAME ..."
  enroot remove -f "$INSTANCE_NAME" >/dev/null 2>&1 || rm -rf "${ENROOT_DATA_PATH}/${INSTANCE_NAME}" 2>/dev/null || true
}
trap cleanup EXIT

acquire_lock() {
  mkdir -p "$DUMPS_DIR"
  exec 9>"$LOCK_FILE"
  if ! flock --nonblock 9 2>/dev/null; then
    warn "Another task is running (lock: $LOCK_FILE). Waiting ..."
    flock 9
  fi
  log "Lock acquired."
}

log "=============================================="
log "  Task:      $TASK"
log "  Max steps: $MAX_STEPS"
log "  Instance:  $INSTANCE_NAME"
log "  Model:     ${MODEL_NAME:-<from eval_config>} (${MODEL_PROVIDER:-${MODEL_PLATFORM:-<from eval_config>}})"
log "  Output:    $OUTPUT_DIR"
log "  PG:        ${PGHOST}:${PGPORT}/${PGDATABASE}"
log "=============================================="

acquire_lock

log "Creating ephemeral rootfs from toolathlon-pack.sqsh ..."
enroot create -n "$INSTANCE_NAME" "$TOOLATHLON_AGENT_SQSH"

ENV_ARGS=(
  -e "PGHOST=${PGHOST}"
  -e "PG_HOST=${PGHOST}"
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
)
for var in MODEL_PROVIDER MODEL_PLATFORM MODEL_NAME MODEL_API_KEY MODEL_API_URL GEMINI_API_KEY \
           MODEL_GREEDY MODEL_TEMPERATURE MODEL_TOP_P MODEL_N; do
  [[ -n "${!var:-}" ]] && ENV_ARGS+=(-e "${var}=${!var}")
done
# Pass proxy so MCP/npm tools that hit network still work
for var in http_proxy https_proxy HTTP_PROXY HTTPS_PROXY no_proxy NO_PROXY; do
  [[ -n "${!var:-}" ]] && ENV_ARGS+=(-e "${var}=${!var}")
done

# Avoid bind-mounting dumps on NFS (enroot file/dir mounts are flaky here).
# Instead copy results out after the run via rsync from the ephemeral rootfs.
# For live output during the run, the task writes under /workspace/dumps inside
# the rootfs; we sync that directory to OUTPUT_DIR afterwards.

run_in() {
  enroot start -r -w \
    "${ENV_ARGS[@]}" \
    "$INSTANCE_NAME" \
    "$@"
}

# Ensure dumps dir exists inside the instance rootfs
mkdir -p "${ENROOT_DATA_PATH}/${INSTANCE_NAME}/workspace/dumps"

log "Fixing email.sent_log FK (best-effort) ..."
run_in /opt/venv/bin/python3 -c "
import psycopg2, os
conn = psycopg2.connect(host=os.environ['PGHOST'], port=os.environ.get('PGPORT','5432'),
                        database=os.environ['PGDATABASE'], user=os.environ['PGUSER'],
                        password=os.environ['PGPASSWORD'])
conn.autocommit = True
cur = conn.cursor()
try:
    cur.execute('ALTER TABLE email.sent_log DROP CONSTRAINT sent_log_message_id_fkey')
    cur.execute('ALTER TABLE email.sent_log ADD CONSTRAINT sent_log_message_id_fkey FOREIGN KEY (message_id) REFERENCES email.messages(id) ON DELETE CASCADE')
except Exception:
    pass
conn.close()
" 2>/dev/null || true

log "Running task ..."
set +e
run_in /bin/bash -c "cd /workspace && exec /opt/venv/bin/python3 main.py --eval_config /workspace/scripts/eval_config.json --task_dir '${TASK}' --max_steps '${MAX_STEPS}' --debug" \
  2>&1 | tee "$OUTPUT_DIR/run.log"
rc=${PIPESTATUS[0]}
set -e

log "Syncing dumps from container rootfs -> $OUTPUT_DIR ..."
rsync -a "${ENROOT_DATA_PATH}/${INSTANCE_NAME}/workspace/dumps/" "$OUTPUT_DIR/" || true

log "Done (exit=$rc). Results: $OUTPUT_DIR"
exit "$rc"
