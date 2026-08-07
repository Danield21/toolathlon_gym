#!/usr/bin/env python3
"""Analyze a kimi-code parallel eval batch: subagent usage, failure modes, timing."""

from __future__ import annotations

import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path


RUN_ID = "20260806-200223"
DUMP_ROOT = Path("/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/dumps/kimi-code")


def find_run_dirs() -> list[Path]:
    dirs = sorted(DUMP_ROOT.glob(f"*/{RUN_ID}_slot*/run.log"))
    return [d.parent for d in dirs]


def parse_run_log(runlog: Path) -> dict:
    text = runlog.read_text(errors="replace")
    out: dict = {"task": runlog.parent.parent.name, "slot": re.search(r"slot(\d+)", str(runlog)).group(1)}

    m = re.search(r"\[(\d\d:\d\d:\d\d)\] DONE\s+(\S+)\s+->\s+(\S+)", text)
    if m:
        out["done_time"] = m.group(1)
        out["worker_status"] = m.group(3)
    m = re.search(r"exit=(\d+),\s*(\d+)s", text)
    if m:
        out["duration_s"] = int(m.group(2))

    m = re.search(r"exited rc=(\d+) claim_done=(True|False)", text)
    if m:
        out["kimi_rc"] = int(m.group(1))
        out["claim_done"] = m.group(2) == "True"
    elif "failed to start isolated postgres" in text:
        out["failure_mode"] = "pg_fail"
        out["claim_done"] = False
    elif "preprocess] Done" in text and "kimi] launching" not in text:
        out["failure_mode"] = "running_or_stuck"
    elif "max_steps_exceeded" in text:
        out["failure_mode"] = "max_steps"
        out["claim_done"] = False

    m = re.search(r"^Pass:\s+(\S+)", text, re.M)
    if m and m.group(1) not in ("None", "null"):
        out["eval_pass"] = m.group(1) == "True"
    elif m:
        out["eval_pass"] = None

    m = re.search(r"=== RESULT: (\S+)", text)
    if m:
        out["eval_result"] = m.group(1)
        err = re.search(r"RESULT: \S+ \((\d+) (blocking )?errors?\)", text)
        if err:
            out["eval_errors"] = int(err.group(1))

    if "max_steps_exceeded" in text:
        out["failure_mode"] = "max_steps"
    elif "claim_done=False" in text and out.get("kimi_rc") == 0:
        out.setdefault("failure_mode", "no_claim_done")
    elif out.get("eval_pass") is False:
        out.setdefault("failure_mode", "eval_fail")

    return out


def analyze_traj(traj_path: Path) -> dict:
    try:
        raw = json.loads(traj_path.read_text())
    except Exception:
        return {}

    if isinstance(raw, dict):
        traj = raw.get("messages") or raw.get("trajectory") or []
    elif isinstance(raw, list):
        traj = raw
    else:
        return {}

    agent_calls = 0
    swarm_calls = 0
    subagent_types: Counter[str] = Counter()
    tool_calls: Counter[str] = Counter()
    swarm_timeouts = 0
    explore_hallucination = False

    for msg in traj:
        role = msg.get("role")
        if role == "assistant":
            for tc in msg.get("tool_calls") or []:
                fn = (tc.get("function") or {}).get("name") or ""
                tool_calls[fn] += 1
                if fn == "Agent":
                    agent_calls += 1
                    args = (tc.get("function") or {}).get("arguments") or ""
                    try:
                        a = json.loads(args) if isinstance(args, str) else args
                        st = a.get("subagent_type") or a.get("agent") or "unknown"
                        subagent_types[st] += 1
                    except Exception:
                        subagent_types["unknown"] += 1
                elif fn == "AgentSwarm":
                    swarm_calls += 1
        elif role == "tool":
            content = msg.get("content") or ""
            if "Subagent timed out" in content:
                swarm_timeouts += content.count("Subagent timed out")
            if "don't have access to Canvas" in content or "only have read-only tools (Read, Glob, Grep" in content:
                explore_hallucination = True

    steps = sum(1 for m in traj if m.get("role") == "assistant")
    return {
        "traj_steps": steps,
        "agent_calls": agent_calls,
        "swarm_calls": swarm_calls,
        "subagent_types": dict(subagent_types),
        "top_tools": tool_calls.most_common(8),
        "swarm_timeouts": swarm_timeouts,
        "explore_hallucination": explore_hallucination,
        "used_subagent": agent_calls + swarm_calls > 0,
    }


def classify_failure(row: dict) -> str:
    if row.get("failure_mode") == "pg_fail":
        return "infra:pg_conflict"
    if row.get("failure_mode") == "running_or_stuck":
        return "infra:still_running"
    if row.get("failure_mode") == "max_steps":
        return "agent:max_steps_exceeded"
    if row.get("claim_done") is False and row.get("kimi_rc") == 0:
        return "agent:no_claim_done"
    if row.get("claim_done") is False and row.get("kimi_rc") == 1:
        return "agent:kimi_error_exit"
    if row.get("eval_pass") is False:
        n = row.get("eval_errors", "?")
        return f"quality:eval_fail({n} errors)"
    if row.get("eval_pass") is None and row.get("claim_done"):
        return "quality:eval_not_run_or_null"
    return "unknown"


def main() -> None:
    rows = []
    for d in find_run_dirs():
        run = parse_run_log(d / "run.log")
        traj_glob = list(d.glob("kimi-code_MiniMax-M3/*/traj.json"))
        if traj_glob:
            run.update(analyze_traj(traj_glob[0]))
        run["failure_class"] = classify_failure(run)
        rows.append(run)

    rows.sort(key=lambda r: int(r.get("slot", 999)))

    # Dedupe yt-12306: keep slot19 as canonical, mark slot18 as duplicate
    seen_tasks: set[str] = set()
    deduped = []
    for r in rows:
        t = r["task"]
        if t in seen_tasks:
            r["failure_class"] = "infra:duplicate_dispatch"
            r["note"] = "duplicate of earlier slot"
        else:
            seen_tasks.add(t)
        deduped.append(r)

    canonical = [r for r in deduped if r.get("note") != "duplicate of earlier slot"]
    n = len(canonical)
    n_done = sum(1 for r in canonical if r.get("failure_class") != "infra:still_running")

    print("=" * 72)
    print(f"KIMI-CODE BATCH REPORT  run_id={RUN_ID}  model=MiniMax-M3")
    print("=" * 72)
    print()

    # Timing
    durations = [r["duration_s"] for r in canonical if r.get("duration_s")]
    if durations:
        print("## Timing")
        print(f"- Tasks (unique): {n}")
        print(f"- Completed: {n_done}/{n}")
        print(f"- Total wall time: ~{max(durations) // 60 + 15} min (batch started ~20:02, last done ~21:17+)")
        print(f"- Per-task duration: min={min(durations)}s max={max(durations)}s avg={sum(durations)//len(durations)}s")
        print(f"- Sum of task durations: {sum(durations)//60} min (serial equivalent)")
        print()

    # Pass rate
    pass_true = sum(1 for r in canonical if r.get("eval_pass") is True)
    pass_false = sum(1 for r in canonical if r.get("eval_pass") is False)
    claim_true = sum(1 for r in canonical if r.get("claim_done") is True)
    print("## Outcomes")
    print(f"- Eval Pass=True:  {pass_true}/{n_done} ({100*pass_true/max(n_done,1):.0f}%)")
    print(f"- Eval Pass=False: {pass_false}/{n_done}")
    print(f"- claim_done=True: {claim_true}/{n_done}")
    print()

    # Failure modes
    print("## Failure Mode Classification")
    fc = Counter(r["failure_class"] for r in canonical)
    for mode, cnt in fc.most_common():
        print(f"- {mode}: {cnt}")
    print()

    # Subagent usage
    used = [r for r in canonical if r.get("used_subagent")]
    print("## Sub-Agent Usage")
    print(f"- Tasks using sub-agents: {len(used)}/{n_done} ({100*len(used)/max(n_done,1):.0f}%)")
    total_agent = sum(r.get("agent_calls", 0) for r in canonical)
    total_swarm = sum(r.get("swarm_calls", 0) for r in canonical)
    print(f"- Total Agent (single) calls: {total_agent}")
    print(f"- Total AgentSwarm calls: {total_swarm}")
    type_totals: Counter[str] = Counter()
    for r in canonical:
        for k, v in (r.get("subagent_types") or {}).items():
            type_totals[k] += v
    if type_totals:
        print("- By subagent_type:")
        for t, c in type_totals.most_common():
            print(f"  - {t}: {c}")
    halluc = sum(1 for r in canonical if r.get("explore_hallucination"))
    timeouts = sum(r.get("swarm_timeouts", 0) for r in canonical)
    print(f"- Explore MCP hallucination detected: {halluc} tasks")
    print(f"- AgentSwarm subagent timeouts (total): {timeouts}")
    print()

    # Per-task table
    print("## Per-Task Detail")
    print(f"{'Task':<42} {'Dur':>5} {'Claim':>5} {'Pass':>5} {'Agent':>5} {'Swarm':>5} {'Failure'}")
    print("-" * 110)
    for r in canonical:
        task = r["task"][:41]
        dur = str(r.get("duration_s", "?"))
        claim = "Y" if r.get("claim_done") else ("?" if r.get("failure_class") == "infra:still_running" else "N")
        pas = "Y" if r.get("eval_pass") is True else ("N" if r.get("eval_pass") is False else "-")
        ag = str(r.get("agent_calls", 0))
        sw = str(r.get("swarm_calls", 0))
        fail = r["failure_class"]
        print(f"{task:<42} {dur:>5} {claim:>5} {pas:>5} {ag:>5} {sw:>5} {fail}")

    # Infra notes
    dupes = [r for r in deduped if r.get("note")]
    if dupes:
        print()
        print("## Infra Issues")
        print(f"- Duplicate dispatch: {dupes[0]['task']} (slot{dupes[0]['slot']}) -> pg_fail")
        print("  (slot19 succeeded; fix: claim_task lock already applied in run_eval_parallel.sh)")


if __name__ == "__main__":
    main()
