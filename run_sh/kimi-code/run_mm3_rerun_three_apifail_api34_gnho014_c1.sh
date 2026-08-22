#!/bin/bash
set -euo pipefail

# MiniMax-M3 auto-subagent rerun of 34 API-killed cases from
# dumps/kimi-code_MiniMax-M3-three-apifail:
#   26 last-hit 429 RPM/TPM
#   6  last-hit 402 insufficient balance (also had 429)
#   2  empty response
# Genuine model case_failed and successes are NOT in this list.
#
# Upstream: 172.16.55.136:8317 via login TCP relay (not official MiniMax).
# Same dump, new run_id. Concurrent=1 on gnho014. KIMI_SUBAGENTS=three.
# After the runner exits, drop replaced old slots for these 34 cases and
# rebuild summary_latest.csv + audit index.

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
DS_SCRIPT_DIR="${PROJECT_ROOT}/run_sh/kimi-code-deepseek-v4-flash"
PYTHON_BIN="${PYTHON_BIN:-/storage/lintaoLab/bowending/miniconda3/envs/dbw_dev/bin/python3}"

export KIMI_CONFIG_ENV="${KIMI_CONFIG_ENV:-${SCRIPT_DIR}/config.linslab8317.env}"
# shellcheck disable=SC1091
source "$KIMI_CONFIG_ENV"

export MODEL_NAME MODEL_API_URL MODEL_API_KEY KIMI_MAX_CONTEXT KIMI_TASK_TIMEOUT_S
unset KIMI_MODEL_THINKING_EFFORT || true
export KIMI_SUBAGENTS=three
unset KIMI_PLAN_FIRST || true

export RELAY_API_KEY="${RELAY_API_KEY:-$MODEL_API_KEY}"
export RELAY_PORT=19332
export RELAY_SKIP_AUTOSTART=1
export RELAY_KIND=tcp
export RELAY_UPSTREAM_HOST="${RELAY_UPSTREAM_HOST:-172.16.55.136}"
export RELAY_UPSTREAM_PORT="${RELAY_UPSTREAM_PORT:-8317}"
export MODEL_API_URL="http://192.168.180.240:${RELAY_PORT}"
export PG_PORT_BASE=45000
export MOCK_PORT_WAIT_LOOPS=360
export MAX_CONCURRENT=1
export MAX_CONCURRENT_CAP=1

export DUMP_ROOT="${DUMP_ROOT:-${PROJECT_ROOT}/dumps/kimi-code_MiniMax-M3-three-apifail}"
export RUN_ID="${RUN_ID:-mm3af3-$(date +%Y%m%d-%H%M%S)}"
export PG_PORT_LEASE_ROOT="${PG_PORT_LEASE_ROOT:-/dev/shm/tpl_mm3af3_${UID}}"
export PG_RUNTIME_ROOT="${PG_RUNTIME_ROOT:-/dev/shm/tpg_mm3af3_${UID}}"
export SLURM_NODELIST="${SLURM_NODELIST:-gnho014}"
export SLURM_MEM="${SLURM_MEM:-64G}"
export SLURM_CPUS="${SLURM_CPUS:-8}"
export SLURM_TIME="${SLURM_TIME:-5-00:00:00}"
export SLURM_JOB_NAME="${SLURM_JOB_NAME:-mm3-af34}"

TASK_LIST_FILE="${TASK_LIST:-${SCRIPT_DIR}/rerun_mm3af_api34.txt}"
mkdir -p "$DUMP_ROOT"

if curl -sS -m 8 -o /dev/null \
     -H "Authorization: Bearer ${MODEL_API_KEY}" \
     "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
  echo "[mm3-af34] relay healthy on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND}"
else
  nohup "$PYTHON_BIN" "${DS_SCRIPT_DIR}/api_relay.py" \
    "$RELAY_PORT" "$RELAY_UPSTREAM_HOST" "$RELAY_UPSTREAM_PORT" \
    >/dev/shm/api_relay_mm3af3_tcp.log 2>&1 &
  sleep 1
  if curl -sS -m 8 -o /dev/null \
       -H "Authorization: Bearer ${MODEL_API_KEY}" \
       "http://127.0.0.1:${RELAY_PORT}/v1/models"; then
    echo "[mm3-af34] relay started on 127.0.0.1:${RELAY_PORT} kind=${RELAY_KIND} pid=$!"
  else
    echo "[mm3-af34] FATAL: TCP relay failed on :${RELAY_PORT}" >&2
    tail -n 40 /dev/shm/api_relay_mm3af3_tcp.log 2>/dev/null || true
    exit 1
  fi
fi

echo "=== ping MiniMax-M3 through login TCP relay :${RELAY_PORT} (no reasoning_effort) ==="
ping_ok=0
for ping_try in 1 2 3; do
  if curl -sS -m 45 -w "\nHTTP %{http_code} time=%{time_total}\n" \
       -H "Authorization: Bearer ${MODEL_API_KEY}" \
       -H "Content-Type: application/json" \
       -d "{\"model\":\"${MODEL_NAME}\",\"messages\":[{\"role\":\"user\",\"content\":\"ping\"}],\"max_tokens\":8}" \
       "http://127.0.0.1:${RELAY_PORT}/v1/chat/completions" | tail -n 8; then
    ping_ok=1
    break
  fi
  echo "[mm3-af34] ping try ${ping_try} failed; retrying..."
  sleep 3
done
if [[ "$ping_ok" != 1 ]]; then
  echo "[mm3-af34] FATAL: MiniMax-M3 ping via :${RELAY_PORT} failed" >&2
  exit 1
fi

mapfile -t TASKS < "$TASK_LIST_FILE"
if (( ${#TASKS[@]} != 34 )); then
  echo "[mm3-af34] FATAL: expected 34 tasks, got ${#TASKS[@]}" >&2
  exit 1
fi

echo "=== MiniMax-M3 8317 API34 KIMI_SUBAGENTS=three c=1 ($(date)) ==="
echo "Model:    $MODEL_NAME @ $MODEL_API_URL  KIMI_SUBAGENTS=$KIMI_SUBAGENTS  effort=<omitted>"
echo "Relay:    login 0.0.0.0:${RELAY_PORT} -> ${RELAY_UPSTREAM_HOST}:${RELAY_UPSTREAM_PORT}"
echo "Dump:     $DUMP_ROOT  run_id=$RUN_ID"
echo "Tasks:    ${#TASKS[@]}  node=$SLURM_NODELIST  concurrent=$MAX_CONCURRENT  cap=$MAX_CONCURRENT_CAP  cpus=$SLURM_CPUS mem=$SLURM_MEM pg_base=$PG_PORT_BASE"
echo "Cleanup:  after runner, drop replaced old slots for these 34 + rebuild summary_latest"
echo "Subagents: plan/coder/explore (KIMI_SUBAGENTS=three via run_on_slurm_three.sh)"
echo "Harness:  live kimi_harness"
echo ""

cd "$PROJECT_ROOT"
set +e
bash "${DS_SCRIPT_DIR}/run_on_slurm_three.sh" "${TASKS[@]}"
RUN_RC=$?
set -e
echo "[mm3-af34] slurm runner rc=${RUN_RC}; cleaning replaced old slots/summaries"

"$PYTHON_BIN" - "$DUMP_ROOT" "$RUN_ID" "$TASK_LIST_FILE" <<'PY'
import csv, shutil, sys
from pathlib import Path

dump = Path(sys.argv[1])
run_id = sys.argv[2]
tasks = [ln.strip() for ln in Path(sys.argv[3]).read_text().splitlines() if ln.strip()]
new_prefix = run_id + "_"
removed_slots = 0
kept_old = []
for task in tasks:
    case = dump / task
    if not case.is_dir():
        continue
    slots = [p for p in case.iterdir() if p.is_dir() and "_slot" in p.name]
    new_slots = [p for p in slots if p.name.startswith(new_prefix)]
    if not new_slots:
        kept_old.append(task)
        continue
    for p in slots:
        if not p.name.startswith(new_prefix):
            shutil.rmtree(p, ignore_errors=True)
            removed_slots += 1
            print(f"[clean] rm slot {p.name}  ({task})")

by_outdir = {}
summaries = sorted(dump.glob("summary_parallel_*.csv"))
latest_existing = dump / "summary_latest.csv"
if latest_existing.exists():
    summaries = [latest_existing] + summaries
for sp in summaries:
    with sp.open(newline="") as f:
        for row in csv.DictReader(f):
            out = row.get("output_dir") or ""
            if out:
                by_outdir[out] = row

latest_rows = []
for case in sorted(p for p in dump.iterdir() if p.is_dir() and not p.name.startswith(".")):
    slots = [p for p in case.iterdir() if p.is_dir() and "_slot" in p.name]
    if not slots:
        continue
    slots.sort(key=lambda p: (p.stat().st_mtime, p.name))
    latest = slots[-1]
    row = by_outdir.get(str(latest))
    if row is None:
        row = {
            "task": case.name,
            "status": "unknown",
            "exit_code": "",
            "output_dir": str(latest),
            "pg_port": "",
            "duration_s": "",
        }
    latest_rows.append(row)

latest_path = dump / "summary_latest.csv"
if latest_rows:
    fields = ["task", "status", "exit_code", "output_dir", "pg_port", "duration_s"]
    with latest_path.open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for row in latest_rows:
            w.writerow({k: row.get(k, "") for k in fields})
    print(f"[clean] wrote {latest_path} rows={len(latest_rows)}")

new_summary = dump / f"summary_parallel_{run_id}.csv"
for sp in dump.glob("summary_parallel_*.csv"):
    if sp.resolve() == new_summary.resolve():
        continue
    sp.unlink(missing_ok=True)
    print(f"[clean] rm summary {sp.name}")

print(f"[clean] removed_old_slots={removed_slots} kept_old_no_new_slot={len(kept_old)}")
if kept_old:
    print("[clean] no new slot, kept previous: " + ", ".join(kept_old))
PY

if [[ "${AUTO_AUDIT_HTML:-1}" == "1" ]]; then
  echo "[mm3-af34] regenerating audit index"
  "$PYTHON_BIN" "${PROJECT_ROOT}/scripts/audit_html_gen.py" "$DUMP_ROOT" || true
fi

echo "[mm3-af34] done runner_rc=${RUN_RC}"
exit "$RUN_RC"
