#!/bin/bash
# Quick smoke: WooCommerce 8081 backend + agent boundary hardening.
# Usage: bash scripts/smoke_wc_boundary.sh [task_dir]
set -euo pipefail

TASK="${1:-ecommerce-market-benchmark}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
RUNTIME_ROOT="${TOOLATHLON_EVAL_DOCKER_ROOT:-/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers}"

# shellcheck disable=SC1091
source "$RUNTIME_ROOT/env.sh"

export PATH="/storage/lintaoLab/bowending/miniconda3/envs/toolathlon_gym/bin:${PATH}"
PGPORT="${SMOKE_PG_PORT:-25499}"
PG_RUNTIME="/dev/shm/smoke_pg_${UID}_${PGPORT}"
PGDATA="${PG_RUNTIME}/data"
PGSOCKET="${PG_RUNTIME}/socket"
INST="smoke-${TASK}"
INIT_SQL="${PROJECT_ROOT}/db/init.sql.gz"
AGENT_TEMPLATE="${ENROOT_DATA_PATH}/toolathlon-pack"
TASK_SOURCE="${PROJECT_ROOT}/tasks/finalpool/${TASK}"

die() { echo "[smoke][error] $*" >&2; exit 1; }
log() { echo "[smoke] $*"; }
pass() { echo "[smoke][PASS] $*"; }
fail() { echo "[smoke][FAIL] $*" >&2; exit 1; }

cleanup() {
  [[ -d "$PGDATA" ]] && pg_ctl -D "$PGDATA" -m fast stop >/dev/null 2>&1 || true
  rm -rf "$PG_RUNTIME" 2>/dev/null || true
  # Do NOT remove toolathlon-pack — it is the shared template.
}
trap cleanup EXIT

[[ -d "$TASK_SOURCE" ]] || die "task not found: $TASK"
[[ -d "$AGENT_TEMPLATE" ]] || die "missing template $AGENT_TEMPLATE — enroot create -n toolathlon-pack \$TOOLATHLON_AGENT_SQSH"
[[ -f "$INIT_SQL" ]] || die "missing $INIT_SQL"

log "=== Phase 1: isolated PG on port $PGPORT ==="
rm -rf "$PG_RUNTIME"
mkdir -p "$PGSOCKET"
initdb -D "$PGDATA" -U "$PGUSER" --auth-local=trust --auth-host=trust --encoding=UTF8 --locale=C >/dev/null
cat >>"$PGDATA/postgresql.conf" <<EOF
port = ${PGPORT}
listen_addresses = '127.0.0.1'
unix_socket_directories = '${PGSOCKET}'
EOF
pg_ctl -D "$PGDATA" -l "${PG_RUNTIME}/postgres.log" -w start
createdb -h "$PGSOCKET" -p "$PGPORT" -U "$PGUSER" "$PGDATABASE" 2>/dev/null || true
gunzip -c "$INIT_SQL" | psql -h "$PGSOCKET" -p "$PGPORT" -U "$PGUSER" -d "$PGDATABASE" -v ON_ERROR_STOP=1 >/dev/null
ln -sfn ".s.PGSQL.${PGPORT}" "${PGSOCKET}/.s.PGSQL.5432"
pass "PostgreSQL ready"

log "=== Phase 2: rootfs (reuse template, no 3.9GB copy) ==="
# Reuse the existing toolathlon-pack template in-place for infra checks only.
# Full agent eval copies per-task via run_eval_parallel.sh.
INST="toolathlon-pack"
[[ -d "${ENROOT_DATA_PATH}/${INST}" ]] || die "missing template — enroot create -n toolathlon-pack \$TOOLATHLON_AGENT_SQSH"
# Refresh harness only (lightweight)
rm -rf "${ENROOT_DATA_PATH}/${INST}/workspace/kimi_harness"
cp -a "${PROJECT_ROOT}/kimi_harness" "${ENROOT_DATA_PATH}/${INST}/workspace/kimi_harness"
pass "rootfs ready (template reuse)"

log "=== Phase 2b: skip task preprocess (read-only task, no injection) ==="
pass "preprocess skipped (not needed for infra smoke)"

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
)

run_in() {
  enroot start -r -w -m "${PGSOCKET}:/run/toolathlon_pg" "${ENV_ARGS[@]}" "$INST" "$@"
}

log "=== Phase 3: WooCommerce 8081 backend ==="
run_in /bin/bash -c '
WC=/opt/local_servers/woocommerce-mcp
test -f "$WC/dist/services/pg-rest-server.js" || { echo "missing pg-rest-server.js"; exit 1; }
nohup node "$WC/dist/services/pg-rest-server.js" >/tmp/wc-rest.log 2>&1 &
for i in $(seq 1 25); do
  curl -fsS -m 2 http://127.0.0.1:8081/health >/dev/null 2>&1 && exit 0
  sleep 0.4
done
echo "8081 failed"; cat /tmp/wc-rest.log; exit 1
' || fail "8081 backend failed to start"
pass "8081 /health OK"

PRODUCTS="$(run_in curl -fsS -m 5 "http://127.0.0.1:8081/wp-json/wc/v3/products?per_page=2" 2>/dev/null)" || fail "WC REST products failed"
echo "$PRODUCTS" | grep -q '"id"' || fail "WC REST returned no products"
pass "WC REST /products returns data: $(echo "$PRODUCTS" | head -c 120)..."

log "=== Phase 4: boundary hardening ==="
run_in /opt/venv/bin/python3 <<'PY' || fail "boundary checks failed"
import os, subprocess, sys
sys.path.insert(0, "/workspace")
from kimi_harness.kimi_main import _mask_blackbox, _unmask_blackbox

mounted = _mask_blackbox()
errors = []

# 1) bind-mount mask: blackbox dirs should look empty
for rel in ("tasks", "utils", "scripts"):
    p = f"/workspace/{rel}"
    try:
        entries = os.listdir(p)
        if entries:
            errors.append(f"mask leak: {p} has {len(entries)} entries")
    except OSError as e:
        errors.append(f"mask list {p}: {e}")

# 2) bind-mount mask: blackbox files should be empty
for rel in ("main.py", "README.md"):
    p = f"/workspace/{rel}"
    if os.path.isfile(p) and os.path.getsize(p) > 0:
        errors.append(f"mask leak: {p} size={os.path.getsize(p)}")

# 3) python_execute read guard on /opt/local_servers
from kimi_harness import local_tools_server as lts
lts.WORKSPACE = "/workspace/agent_workspace"
lts.MARKER = "/tmp/smoke.marker"
lts.ENABLED = {"python_execute"}
code = "open('/opt/local_servers/woocommerce-mcp/package.json').read()"
out = lts._python_execute(code)
if "PermissionError" not in out and "task boundary" not in out:
    errors.append(f"python_execute read guard failed: {out[:200]}")

# 4) setpriv: agent cannot umount masks (CAP_SYS_ADMIN stripped)
rc = subprocess.run(
    ["setpriv", "--bounding-set=-all", "--inh-caps=-all", "--ambient-caps=-all",
     "--no-new-privs", "umount", mounted[0] if mounted else "/workspace/tasks"],
    capture_output=True, text=True,
).returncode
if rc == 0:
    errors.append("setpriv agent was able to umount a mask (cap leak)")

_unmask_blackbox(mounted)

if errors:
    for e in errors:
        print(f"[smoke][FAIL] {e}", file=sys.stderr)
    sys.exit(1)
print("boundary checks OK")
PY
pass "bind-mount masks hide blackbox paths"
pass "python_execute blocks /opt/local_servers reads"
pass "setpriv agent cannot umount masks"

log "=== All smoke checks passed ==="
