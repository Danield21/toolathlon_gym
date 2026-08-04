#!/usr/bin/env bash
# Fully isolated Toolathlon evaluation with rootless Podman.
#
# Each task gets one private Podman pod containing:
#   - one ephemeral PostgreSQL container
#   - one ephemeral Toolathlon agent container
# Both containers share only that pod's network namespace, so every task can
# safely use PostgreSQL on 127.0.0.1:5432 without publishing a host port.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
CONFIG_ENV="${CONFIG_ENV:-$SCRIPT_DIR/config.env}"
ENV_MAX_CONCURRENT="${MAX_CONCURRENT:-}"
ENV_MAX_STEPS="${MAX_STEPS:-}"
ENV_DUMP_ROOT="${DUMP_ROOT:-}"

# shellcheck disable=SC1090
source "$CONFIG_ENV"

PODMAN="${PODMAN:-podman}"
AGENT_IMAGE="${AGENT_IMAGE:-localhost/toolathlon-pack:qwen36}"
POSTGRES_IMAGE="${POSTGRES_IMAGE:-docker.io/library/postgres:15}"
MAX_CONCURRENT="${ENV_MAX_CONCURRENT:-${MAX_CONCURRENT:-3}}"
MAX_STEPS="${ENV_MAX_STEPS:-${MAX_STEPS:-100}}"
DUMP_ROOT="${ENV_DUMP_ROOT:-${DUMP_ROOT:-$PROJECT_ROOT/dumps/qwen3.6-35B-A3B-podman}}"
INIT_SQL="$PROJECT_ROOT/db/init.sql.gz"
FAKE_UID_SRC="$SCRIPT_DIR/podman_fake_postgres_uid.c"
RUN_ID="$(date +%Y%m%d-%H%M%S)"
FAKE_UID_SO="/tmp/toolathlon_fake_postgres_uid_${RUN_ID}_$$.so"
SUMMARY="$DUMP_ROOT/summary_podman_${RUN_ID}.csv"
SUMMARY_LOCK="/tmp/toolathlon_podman_summary_${RUN_ID}_$$.lock"
POD_REGISTRY="/tmp/toolathlon_podman_pods_${RUN_ID}_$$"
POD_REGISTRY_LOCK="${POD_REGISTRY}.lock"

export MODEL_NAME MODEL_PLATFORM MODEL_PROVIDER MODEL_API_KEY MODEL_API_URL
export MODEL_GREEDY MODEL_TEMPERATURE MODEL_TOP_P MODEL_N
MODEL_N="${MODEL_N:-}"

log() { echo "[$(date +%H:%M:%S)] $*"; }
die() { echo "[error] $*" >&2; exit 1; }

append_summary() {
  while ! mkdir "$SUMMARY_LOCK" 2>/dev/null; do sleep 0.05; done
  echo "$1" >> "$SUMMARY"
  rmdir "$SUMMARY_LOCK"
}

register_pod() {
  while ! mkdir "$POD_REGISTRY_LOCK" 2>/dev/null; do sleep 0.05; done
  echo "$1" >> "$POD_REGISTRY"
  rmdir "$POD_REGISTRY_LOCK"
}

cleanup_all() {
  local pod
  if [[ -f "$POD_REGISTRY" ]]; then
    while IFS= read -r pod; do
      [[ -n "$pod" ]] && "$PODMAN" pod rm -f "$pod" >/dev/null 2>&1 || true
    done < "$POD_REGISTRY"
  fi
  rm -f "$POD_REGISTRY" "$POD_REGISTRY_LOCK"
  rm -f "$FAKE_UID_SO"
  rmdir "$SUMMARY_LOCK" 2>/dev/null || true
}
trap cleanup_all EXIT

need_prerequisites() {
  command -v "$PODMAN" >/dev/null 2>&1 || die "podman not found"
  command -v gcc >/dev/null 2>&1 || die "gcc not found (needed for PostgreSQL rootless UID shim)"
  [[ -f "$INIT_SQL" ]] || die "missing $INIT_SQL"
  [[ -f "$FAKE_UID_SRC" ]] || die "missing $FAKE_UID_SRC"
  "$PODMAN" image exists "$AGENT_IMAGE" || die "missing agent image: $AGENT_IMAGE"
  "$PODMAN" image exists "$POSTGRES_IMAGE" || die "missing postgres image: $POSTGRES_IMAGE"
  (( MAX_CONCURRENT >= 1 )) || die "MAX_CONCURRENT must be >= 1"
  gcc -shared -fPIC -O2 -Wall -Wextra -o "$FAKE_UID_SO" "$FAKE_UID_SRC" -ldl
}

wait_for_postgres() {
  local pg_container="$1" retries=300 state ready init_complete
  while (( retries > 0 )); do
    state="$("$PODMAN" inspect --format '{{.State.Status}}' "$pg_container" 2>/dev/null || echo missing)"
    [[ "$state" == exited || "$state" == missing ]] && return 1
    init_complete="$("$PODMAN" logs "$pg_container" 2>&1 | \
      grep -F 'PostgreSQL init process complete; ready for start up.' || true)"
    if [[ -n "$init_complete" ]]; then
      ready="$("$PODMAN" exec -e PGPASSWORD=camel "$pg_container" \
        psql -U eigent -d toolathlon_gym -tAc \
        "SELECT CASE WHEN to_regclass('email.folders') IS NOT NULL AND EXISTS (SELECT 1 FROM pg_indexes WHERE schemaname='email' AND tablename='folders' AND indexdef ILIKE 'CREATE UNIQUE INDEX%' AND indexdef LIKE '%(name)%') THEN 1 ELSE 0 END" \
        2>/dev/null | tr -d '[:space:]' || true)"
      [[ "$ready" == 1 ]] && return 0
    fi
    sleep 2
    retries=$((retries - 1))
  done
  return 1
}

check_database() {
  local pg_container="$1" output_file="$2"
  "$PODMAN" exec -i -e PGPASSWORD=camel "$pg_container" \
    psql -v ON_ERROR_STOP=1 -U eigent -d toolathlon_gym <<'SQL' >"$output_file" 2>&1
ALTER TABLE email.sent_log DROP CONSTRAINT IF EXISTS sent_log_message_id_fkey;
ALTER TABLE email.sent_log ADD CONSTRAINT sent_log_message_id_fkey
  FOREIGN KEY (message_id) REFERENCES email.messages(id) ON DELETE CASCADE;

DO $$
BEGIN
  IF to_regclass('email.folders') IS NULL THEN
    RAISE EXCEPTION 'email.folders is missing';
  END IF;
  IF NOT EXISTS (
    SELECT 1
    FROM pg_indexes
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

SELECT 'database isolation check: OK' AS status;
SQL
}

run_one_task() {
  local task="$1" slot="$2"
  local safe hash pod pg agent outdir task_log start_ts end_ts rc status api_code

  safe="$(echo "$task" | tr -cs 'A-Za-z0-9_.-' '-')"
  hash="$(printf '%s' "$task-$RUN_ID-$slot-$$" | md5sum | cut -c1-10)"
  pod="ta-q36-${hash}"
  pg="pg-${hash}"
  agent="agent-${hash}"
  outdir="$DUMP_ROOT/$task/${RUN_ID}_pod${slot}"
  task_log="$outdir/run.log"
  start_ts="$(date +%s)"
  mkdir -p "$outdir"

  ACTIVE_POD="$pod"
  register_pod "$pod"
  log "START  $task (pod=$pod slot=$slot)"

  if ! "$PODMAN" pod create --name "$pod" >"$outdir/pod_create.log" 2>&1; then
    append_summary "$task,pod_fail,1,$outdir,$(( $(date +%s) - start_ts ))"
    return 1
  fi

  if ! "$PODMAN" run -d \
      --pod "$pod" \
      --name "$pg" \
      -e POSTGRES_DB=toolathlon_gym \
      -e POSTGRES_USER=eigent \
      -e POSTGRES_PASSWORD=camel \
      -e LD_PRELOAD=/opt/toolathlon/fake_postgres_uid.so \
      -v "$INIT_SQL:/docker-entrypoint-initdb.d/00-init.sql.gz:ro" \
      -v "$FAKE_UID_SO:/opt/toolathlon/fake_postgres_uid.so:ro" \
      --health-cmd='pg_isready -U eigent -d toolathlon_gym' \
      --health-interval=3s \
      --health-timeout=5s \
      --health-retries=80 \
      "$POSTGRES_IMAGE" >"$outdir/postgres_start.log" 2>&1; then
    append_summary "$task,pg_start_fail,1,$outdir,$(( $(date +%s) - start_ts ))"
    return 1
  fi

  if ! wait_for_postgres "$pg"; then
    "$PODMAN" logs "$pg" >"$outdir/postgres.log" 2>&1 || true
    append_summary "$task,pg_unhealthy,1,$outdir,$(( $(date +%s) - start_ts ))"
    return 1
  fi

  if ! check_database "$pg" "$outdir/db_check.log"; then
    "$PODMAN" logs "$pg" >"$outdir/postgres.log" 2>&1 || true
    append_summary "$task,pg_schema_fail,1,$outdir,$(( $(date +%s) - start_ts ))"
    return 1
  fi

  if ! "$PODMAN" run -d \
      --pod "$pod" \
      --name "$agent" \
      -e PGHOST=127.0.0.1 \
      -e PG_HOST=127.0.0.1 \
      -e PGPORT=5432 \
      -e PG_PORT=5432 \
      -e PGUSER=eigent \
      -e PG_USER=eigent \
      -e PGPASSWORD=camel \
      -e PG_PASSWORD=camel \
      -e PGDATABASE=toolathlon_gym \
      -e PG_DATABASE=toolathlon_gym \
      -e LOCAL_SERVERS_PATH=/opt/local_servers \
      -e PYTHON_BIN=/opt/venv/bin/python3 \
      -e MODEL_NAME \
      -e MODEL_PLATFORM \
      -e MODEL_PROVIDER \
      -e MODEL_API_KEY \
      -e MODEL_API_URL \
      -e MODEL_GREEDY \
      -e MODEL_TEMPERATURE \
      -e MODEL_TOP_P \
      -e MODEL_N \
      -e NO_PROXY='*' \
      -e no_proxy='*' \
      -v "$outdir:/workspace/dumps:rw" \
      -w /workspace \
      "$AGENT_IMAGE" sleep infinity >"$outdir/agent_start.log" 2>&1; then
    append_summary "$task,agent_start_fail,1,$outdir,$(( $(date +%s) - start_ts ))"
    return 1
  fi

  api_code="$("$PODMAN" exec "$agent" curl -sS --connect-timeout 8 --max-time 20 \
    -o /dev/null -w '%{http_code}' "${MODEL_API_URL%/}/v1/models" 2>/dev/null || true)"
  echo "$api_code" > "$outdir/model_api_http_code.txt"
  if [[ "$api_code" != 200 ]]; then
    append_summary "$task,model_api_fail,1,$outdir,$(( $(date +%s) - start_ts ))"
    return 1
  fi

  set +e
  "$PODMAN" exec -w /workspace "$agent" \
    /opt/venv/bin/python3 main.py \
      --eval_config /workspace/scripts/eval_config.json \
      --task_dir "$task" \
      --max_steps "$MAX_STEPS" \
      --debug >"$task_log" 2>&1
  rc=$?
  set -e

  "$PODMAN" logs "$pg" >"$outdir/postgres.log" 2>&1 || true
  "$PODMAN" logs "$agent" >"$outdir/agent_container.log" 2>&1 || true

  end_ts="$(date +%s)"
  if [[ $rc -eq 0 ]]; then status=success; else status=failed; fi
  append_summary "$task,$status,$rc,$outdir,$((end_ts - start_ts))"
  log "DONE   $task -> $status (exit=$rc, $((end_ts - start_ts))s)"
  return "$rc"
}

export -f log append_summary register_pod wait_for_postgres check_database run_one_task
export PODMAN AGENT_IMAGE POSTGRES_IMAGE MAX_STEPS DUMP_ROOT INIT_SQL FAKE_UID_SRC FAKE_UID_SO RUN_ID SUMMARY SUMMARY_LOCK POD_REGISTRY POD_REGISTRY_LOCK
export MODEL_NAME MODEL_PLATFORM MODEL_PROVIDER MODEL_API_KEY MODEL_API_URL
export MODEL_GREEDY MODEL_TEMPERATURE MODEL_TOP_P MODEL_N

if [[ $# -gt 0 ]]; then
  TASKS=("$@")
else
  mapfile -t TASKS < <(find "$PROJECT_ROOT/tasks/finalpool" -mindepth 1 -maxdepth 1 -type d -printf '%f\n' | sort)
fi
[[ ${#TASKS[@]} -gt 0 ]] || die "no tasks selected"

mkdir -p "$DUMP_ROOT"
touch "$POD_REGISTRY"
echo "task,status,exit_code,output_dir,duration_s" > "$SUMMARY"
need_prerequisites

log "Podman isolated evaluation"
log "Tasks=${#TASKS[@]} concurrency=$MAX_CONCURRENT agent_image=$AGENT_IMAGE"
log "Model=$MODEL_NAME @ $MODEL_API_URL"
log "Summary=$SUMMARY"

FIFO="/tmp/toolathlon_podman_sem_${RUN_ID}_$$"
mkfifo "$FIFO"
exec 3<>"$FIFO"
rm -f "$FIFO"
for ((i = 0; i < MAX_CONCURRENT; i++)); do echo >&3; done

PIDS=()
slot=0
for task in "${TASKS[@]}"; do
  read -u 3
  (
    ACTIVE_POD=""
    trap '[[ -n "${ACTIVE_POD:-}" ]] && "$PODMAN" pod rm -f "$ACTIVE_POD" >/dev/null 2>&1 || true' EXIT
    set +e
    run_one_task "$task" "$slot"
    rc=$?
    echo >&3
    exit "$rc"
  ) &
  PIDS+=("$!")
  slot=$((slot + 1))
done

failed=0
for pid in "${PIDS[@]}"; do
  wait "$pid" || failed=$((failed + 1))
done
exec 3>&-

log "Finished: total=${#TASKS[@]} failed_workers=$failed"
log "Summary: $SUMMARY"
exit "$failed"
