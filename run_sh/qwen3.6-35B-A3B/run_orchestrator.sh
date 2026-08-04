#!/usr/bin/env bash
# Orchestrate: wait for SGLang healthy, then run all-task parallel eval + case study.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
ENV_DUMP_ROOT="${DUMP_ROOT:-}"
ENV_MAX_CONCURRENT="${MAX_CONCURRENT:-}"
ENV_MAX_STEPS="${MAX_STEPS:-}"
ENV_PG_PORT_BASE="${PG_PORT_BASE:-}"
ENV_PG_RUNTIME_ROOT="${PG_RUNTIME_ROOT:-}"
# shellcheck disable=SC1091
source "$SCRIPT_DIR/config.env"

[[ -n "$ENV_DUMP_ROOT" ]] && DUMP_ROOT="$ENV_DUMP_ROOT"
[[ -n "$ENV_MAX_CONCURRENT" ]] && export MAX_CONCURRENT="$ENV_MAX_CONCURRENT"
[[ -n "$ENV_MAX_STEPS" ]] && export MAX_STEPS="$ENV_MAX_STEPS"
[[ -n "$ENV_PG_PORT_BASE" ]] && export PG_PORT_BASE="$ENV_PG_PORT_BASE"
[[ -n "$ENV_PG_RUNTIME_ROOT" ]] && export PG_RUNTIME_ROOT="$ENV_PG_RUNTIME_ROOT"

DUMP_ROOT="${DUMP_ROOT:-$PROJECT_ROOT/dumps/qwen3.6-35B-A3B}"
export DUMP_ROOT
ORCH_LOG="$DUMP_ROOT/orchestrator_$(date +%Y%m%d-%H%M%S).log"
mkdir -p "$DUMP_ROOT"
exec > >(tee -a "$ORCH_LOG") 2>&1

echo "=============================================="
echo "  Orchestrator start: $(date)"
echo "  Model API: $MODEL_API_URL"
echo "  Dump root: $DUMP_ROOT"
echo "  Log: $ORCH_LOG"
echo "=============================================="

API_BASE="${MODEL_API_URL%/}"
# Accept either ...:port or ...:port/v1
if [[ "$API_BASE" == */v1 ]]; then
  MODELS_URL="${API_BASE}/models"
else
  MODELS_URL="${API_BASE}/v1/models"
fi

echo "[orch] waiting for SGLang at $MODELS_URL ..."
# login01 proxy must not intercept cluster-internal API (would return 502)
unset http_proxy https_proxy HTTP_PROXY HTTPS_PROXY ALL_PROXY all_proxy || true
export NO_PROXY="*"
export no_proxy="*"
READY=0
# gnho019 is offline; model load can take a while after sbatch. Keep waiting longer.
for i in $(seq 1 720); do
  code=$(curl -sS --noproxy '*' --connect-timeout 3 -o /tmp/qwen36_models.json -w '%{http_code}' \
    -H "Authorization: Bearer ${MODEL_API_KEY:-EMPTY}" \
    "$MODELS_URL" || true)
  if [[ "$code" == "200" ]]; then
    echo "[orch] SGLang healthy (HTTP 200) after try $i"
    head -c 500 /tmp/qwen36_models.json; echo
    READY=1
    break
  fi
  echo "[orch] not ready yet (HTTP ${code:-000}), sleep 10s ($i/720)"
  sleep 10
done
if [[ "$READY" != "1" ]]; then
  echo "[orch] FATAL: SGLang never became healthy" >&2
  exit 2
fi

# Prefer toolathlon_gym conda for PG/enroot eval
export PATH="/storage/lintaoLab/bowending/miniconda3/envs/toolathlon_gym/bin:$PATH"
source /lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers/env.sh

set +e
bash "$SCRIPT_DIR/run_eval_parallel.sh" "$@"
EVAL_RC=$?
set -e

echo "=============================================="
echo "  Eval finished rc=$EVAL_RC at $(date)"
echo "=============================================="

# Lightweight case study from latest summary
python3 - "$DUMP_ROOT" <<'PY'
import csv, json, re, sys
from datetime import datetime
from pathlib import Path

dump_root = Path(sys.argv[1])
summaries = sorted(dump_root.glob("summary_parallel_*.csv"), key=lambda p: p.stat().st_mtime)
if not summaries:
    print("[case_study] no summary csv"); sys.exit(0)
summary = summaries[-1]
rows = list(csv.DictReader(summary.open()))
md = [f"# Qwen3.6-35B-A3B Case Study\n",
      f"- Generated: `{datetime.now().isoformat(timespec='seconds')}`\n",
      f"- Summary: `{summary}`\n",
      f"- Sampling: greedy\n\n",
      f"**Tasks:** {len(rows)}\n\n",
      "| task | status | exit | duration_s |\n|---|---|---:|---:|\n"]
fails = 0
for r in rows:
    md.append(f"| `{r.get('task','')}` | {r.get('status','')} | {r.get('exit_code','')} | {r.get('duration_s','')} |\n")
    if r.get('status') != 'success' or str(r.get('exit_code','0')) != '0':
        fails += 1
md.append(f"\n**Failed workers:** {fails}/{len(rows)}\n")
out = dump_root / f"CASE_STUDY_{summary.stem}.md"
out.write_text("".join(md), encoding="utf-8")
(dump_root / "CASE_STUDY_latest.md").write_text("".join(md), encoding="utf-8")
print(f"[case_study] wrote {out}")
PY

exit "$EVAL_RC"
