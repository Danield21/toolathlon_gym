#!/bin/bash
# Orchestrate parallel deepseek-v4-pro eval + failure case study.
# Intended to run inside tmux session `toolathlon_dsv4pro`.

set -euo pipefail

PROJECT_ROOT=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym
RUNTIME_ROOT=/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym_eval_dockers
DUMP_ROOT="$PROJECT_ROOT/dumps/deepseek-v4-pro"
ORCH_LOG="$DUMP_ROOT/orchestrator_$(date +%Y%m%d-%H%M%S).log"

mkdir -p "$DUMP_ROOT"
exec > >(tee -a "$ORCH_LOG") 2>&1

source "$RUNTIME_ROOT/env.sh"
# Prefer toolathlon_gym conda binaries
export PATH="/storage/lintaoLab/bowending/miniconda3/envs/toolathlon_gym/bin:$PATH"

cd "$PROJECT_ROOT"

echo "=============================================="
echo "  Orchestrator start: $(date)"
echo "  Dump root: $DUMP_ROOT"
echo "  Log: $ORCH_LOG"
echo "=============================================="

set +e
bash "$PROJECT_ROOT/run_sh/deepseek-v4-pro/run_eval_parallel.sh"
EVAL_RC=$?
set -e

echo "=============================================="
echo "  Eval finished rc=$EVAL_RC at $(date)"
echo "  Writing failure case study ..."
echo "=============================================="

python3 - "$DUMP_ROOT" <<'PY'
import csv, json, os, sys, glob, re
from datetime import datetime
from pathlib import Path

dump_root = Path(sys.argv[1])
summaries = sorted(dump_root.glob("summary_parallel_*.csv"), key=lambda p: p.stat().st_mtime)
if not summaries:
    # also accept sequential summaries
    summaries = sorted(dump_root.glob("summary_*.csv"), key=lambda p: p.stat().st_mtime)
if not summaries:
    print("[case_study] no summary csv found", file=sys.stderr)
    sys.exit(0)

summary = summaries[-1]
rows = []
with summary.open() as f:
    reader = csv.DictReader(f)
    for r in reader:
        rows.append(r)

def read_tail(path, n=80):
    try:
        lines = Path(path).read_text(errors="replace").splitlines()
        return "\n".join(lines[-n:])
    except Exception as e:
        return f"<unreadable: {e}>"

def extract_signals(text: str):
    keys = []
    patterns = [
        (r"Pass:\s*(True|False)", "pass_flag"),
        (r"Status:\s*(\w+)", "agent_status"),
        (r"(?i)traceback", "traceback"),
        (r"(?i)error[:\s].{0,120}", "error_line"),
        (r"(?i)rate limit|429|timeout|connection", "network"),
        (r"(?i)AssertionError|assert ", "assert"),
        (r"(?i)tool.?call|MCP", "tooling"),
    ]
    found = {k: [] for _, k in patterns}
    for line in text.splitlines():
        for pat, name in patterns:
            if re.search(pat, line):
                found[name].append(line.strip()[:240])
    return found

def find_eval_json(outdir: Path):
    cands = list(outdir.rglob("eval_res.json")) + list(outdir.rglob("*eval*.json"))
    return cands[:5]

md = []
md.append(f"# DeepSeek-V4-Pro Failure Case Study\n")
md.append(f"- Generated: `{datetime.now().isoformat(timespec='seconds')}`\n")
md.append(f"- Summary CSV: `{summary}`\n")
md.append(f"- Dump root: `{dump_root}`\n")
md.append(f"- Sampling: greedy (`temperature=0`, `top_p=1`)\n")

total = len(rows)
fails = [r for r in rows if r.get("status") not in ("success",) or r.get("exit_code", "0") not in ("0",)]
# Also treat eval Pass:False as failure even if exit 0 — probe logs
soft_fails = []

md.append("\n## Overview\n")
md.append("| task | status | exit | duration_s | output_dir |\n")
md.append("|---|---|---:|---:|---|\n")
for r in rows:
    md.append(f"| `{r.get('task','')}` | {r.get('status','')} | {r.get('exit_code','')} | {r.get('duration_s','')} | `{r.get('output_dir','')}` |\n")

# Enrich with Pass flag from logs
for r in rows:
    out = Path(r.get("output_dir") or "")
    logp = out / "run.log"
    text = read_tail(logp, 400) if logp.exists() else ""
    m = re.search(r"Pass:\s*(True|False)", text)
    if m and m.group(1) == "False":
        soft_fails.append(r)
    if r.get("status") == "pg_fail":
        soft_fails.append(r)

# Unique failure set: hard fail or Pass False
fail_keys = set()
failure_rows = []
for r in rows:
    key = r.get("task")
    hard = r.get("status") != "success" or str(r.get("exit_code", "0")) != "0"
    out = Path(r.get("output_dir") or "")
    logp = out / "run.log"
    text = logp.read_text(errors="replace") if logp.exists() else ""
    pass_false = bool(re.search(r"Pass:\s*False", text))
    if hard or pass_false:
        if key not in fail_keys:
            fail_keys.add(key)
            failure_rows.append((r, text, hard, pass_false))

md.append(f"\n**Total tasks:** {total}  \n")
md.append(f"**Failure / Pass=False cases:** {len(failure_rows)}  \n")

if not failure_rows:
    md.append("\n## Result\n\nAll cases succeeded with `Pass: True`. No failure case study needed.\n")
else:
    md.append("\n## Failure Case Studies\n")
    for i, (r, text, hard, pass_false) in enumerate(failure_rows, 1):
        task = r.get("task", "")
        out = Path(r.get("output_dir") or "")
        md.append(f"\n### {i}. `{task}`\n")
        md.append(f"- **status (runner):** `{r.get('status')}` (exit={r.get('exit_code')})\n")
        md.append(f"- **duration_s:** {r.get('duration_s')}\n")
        md.append(f"- **pg_port:** {r.get('pg_port', 'n/a')}\n")
        md.append(f"- **output_dir:** `{out}`\n")
        md.append(f"- **eval Pass:False:** {pass_false}\n")
        md.append(f"- **hard runner failure:** {hard}\n")

        signals = extract_signals(text)
        md.append("\n#### Observed signals\n")
        for k, vals in signals.items():
            if vals:
                md.append(f"- **{k}:**\n")
                for v in vals[-5:]:
                    md.append(f"  - `{v}`\n")

        # eval_res.json snippets
        evals = find_eval_json(out) if out.exists() else []
        if evals:
            md.append("\n#### eval artifacts\n")
            for ep in evals:
                try:
                    data = json.loads(ep.read_text())
                    md.append(f"- `{ep.relative_to(out) if out in ep.parents else ep}`:\n")
                    md.append("```json\n")
                    md.append(json.dumps(data, indent=2, ensure_ascii=False)[:4000])
                    md.append("\n```\n")
                except Exception as e:
                    md.append(f"- `{ep}` unreadable: {e}\n")

        md.append("\n#### Hypothesis\n")
        # Simple heuristic hypotheses
        hyps = []
        low = text.lower()
        if "pg_fail" in (r.get("status") or "") or "postgres" in low and "fatal" in low:
            hyps.append("Isolated PostgreSQL failed to start or accept connections before the agent ran.")
        if "traceback" in low:
            hyps.append("Unhandled Python exception in agent / MCP / evaluation path (see traceback).")
        if re.search(r"429|rate limit|timeout|connection", low):
            hyps.append("LLM API connectivity / rate-limit / timeout interfered with tool-use trajectory.")
        if pass_false and not hard:
            hyps.append("Agent finished but automated evaluator checks failed (wrong/missing artifacts vs groundtruth).")
        if "pass: false" in low and "assertion" in low:
            hyps.append("Evaluation assertion mismatch on workspace outputs (xlsx/docx/email/db side effects).")
        if not hyps:
            hyps.append("Insufficient automatic signals; inspect full run.log and traj for tool-call mistakes.")
        for h in hyps:
            md.append(f"- {h}\n")

        md.append("\n#### Suggested next probes\n")
        md.append("- Read full `run.log` and CAMEL traj under the output dir.\n")
        md.append("- Diff agent workspace artifacts against `tasks/finalpool/<task>/groundtruth_workspace/`.\n")
        md.append("- Re-run this single task: `bash run_sh/deepseek-v4-pro/run_eval_parallel.sh %s`.\n" % task)

        md.append("\n#### run.log tail\n\n```text\n")
        md.append(read_tail(out / "run.log", 60) if (out / "run.log").exists() else "<missing run.log>")
        md.append("\n```\n")

md.append("\n## Notes\n")
md.append("- Parallel isolation: each task used a dedicated Postgres port + ephemeral enroot rootfs.\n")
md.append("- Decoding: greedy (`MODEL_GREEDY=1`, `temperature=0`, `top_p=1`).\n")

out_md = dump_root / f"CASE_STUDY_{summary.stem}.md"
out_md.write_text("".join(md), encoding="utf-8")
# also write/overwrite a stable name
stable = dump_root / "CASE_STUDY_latest.md"
stable.write_text("".join(md), encoding="utf-8")
print(f"[case_study] wrote {out_md}")
print(f"[case_study] wrote {stable}")
PY

echo "=============================================="
echo "  Orchestrator done: $(date)"
echo "  Case study: $DUMP_ROOT/CASE_STUDY_latest.md"
echo "=============================================="
exit "$EVAL_RC"
