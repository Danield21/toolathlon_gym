#!/bin/bash
# Smoke tests for enroot-based toolathlon_gym setup (replaces test_containerized.sh)

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"
RUNTIME_ROOT="${TOOLATHLON_EVAL_DOCKER_ROOT:-/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers}"

# shellcheck disable=SC1091
source "$RUNTIME_ROOT/env.sh"

PASS=0
FAIL=0
pass() { echo "  [PASS] $*"; PASS=$((PASS + 1)); }
fail() { echo "  [FAIL] $*" >&2; FAIL=$((FAIL + 1)); }
section() { echo; echo "=== $* ==="; }

section "Test 1: enroot CLI"
if command -v enroot >/dev/null 2>&1; then
  pass "enroot is available ($(enroot version 2>/dev/null | head -1))"
else
  fail "enroot not found"
fi

section "Test 2: runtime paths"
for p in "$TOOLATHLON_PG_SQSH" "$ENROOT_DATA_PATH/toolathlon_pg"; do
  if [[ -e "$p" ]]; then pass "exists: $p"; else fail "missing: $p"; fi
done
if [[ -f "$TOOLATHLON_AGENT_SQSH" ]]; then
  pass "agent sqsh: $TOOLATHLON_AGENT_SQSH"
elif [[ -d "$ENROOT_DATA_PATH/toolathlon-pack" ]] || [[ -d "$ENROOT_DATA_PATH/toolathlon-pack-build" ]]; then
  pass "agent rootfs present (sqsh may still be exporting)"
else
  fail "agent image not built yet — run scripts/enroot_build_agent.sh"
fi

section "Test 3: postgres health"
if command -v pg_isready >/dev/null 2>&1; then
  if PGPASSWORD="$PGPASSWORD" pg_isready -h "$PGHOST" -p "$PGPORT" -U "$PGUSER" -d "$PGDATABASE" >/dev/null 2>&1; then
    pass "postgres ready at ${PGHOST}:${PGPORT}"
  else
    fail "postgres not ready — run: bash scripts/enroot_postgres.sh start"
  fi
else
  fail "pg_isready not in PATH (activate conda env toolathlon_gym)"
fi

section "Test 4: DB query"
if command -v psql >/dev/null 2>&1 && PGPASSWORD="$PGPASSWORD" pg_isready -h "$PGHOST" -p "$PGPORT" -U "$PGUSER" -d "$PGDATABASE" >/dev/null 2>&1; then
  res="$(PGPASSWORD="$PGPASSWORD" psql -h "$PGHOST" -p "$PGPORT" -U "$PGUSER" -d "$PGDATABASE" -tAc 'SELECT 1' 2>/dev/null || echo error)"
  if [[ "$res" == "1" ]]; then pass "SELECT 1 ok"; else fail "query failed: $res"; fi
else
  fail "skip DB query (psql/pg not ready)"
fi

section "Test 5: agent paths (if image ready)"
AGENT_FS=""
if [[ -d "$ENROOT_DATA_PATH/toolathlon-pack" ]]; then
  AGENT_FS="$ENROOT_DATA_PATH/toolathlon-pack"
elif [[ -d "$ENROOT_DATA_PATH/toolathlon-pack-build" ]]; then
  AGENT_FS="$ENROOT_DATA_PATH/toolathlon-pack-build"
fi
if [[ -n "$AGENT_FS" ]]; then
  for rel in opt/venv/bin/python3 opt/local_servers workspace/main.py; do
    if [[ -e "$AGENT_FS/$rel" ]]; then pass "$rel"; else fail "missing in image: $rel"; fi
  done
else
  fail "no agent rootfs to inspect"
fi

section "Test 6: sequential lock"
DUMPS_DIR="$PROJECT_ROOT/dumps"
LOCK_FILE="$DUMPS_DIR/.run.lock"
mkdir -p "$DUMPS_DIR"
rm -f "$LOCK_FILE"
(
  exec 9>"$LOCK_FILE"
  flock 9
  sleep 2
) &
HOLDER=$!
sleep 0.3
if ( exec 9>"$LOCK_FILE"; flock --nonblock 9 ) 2>/dev/null; then
  fail "lock should have been held"
else
  pass "flock blocks while held"
fi
wait "$HOLDER" 2>/dev/null || true
if ( exec 9>"$LOCK_FILE"; flock --nonblock 9 ) 2>/dev/null; then
  pass "flock acquirable after release"
else
  fail "flock still blocked"
fi

echo
echo "=============================="
echo "  PASS: $PASS"
echo "  FAIL: $FAIL"
echo "=============================="
(( FAIL == 0 ))
