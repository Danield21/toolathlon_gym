#!/usr/bin/env python3
"""Scan all case dumps for infra problems the model had to fight instead of solving the task.

Categories:
  - mcp_down:        MCP server crashed / connection refused / not running
  - tool_crash:      a granted tool returns errors repeatedly (server-side)
  - pg_fail:         PostgreSQL failed to start (from run.log)
  - preprocess_fail: preprocess/injection errors
  - enroot_leftover: rootfs cleanup warning
  - duplicate_dispatch: same task run twice
  - timeout_loop:    repeated TIMEOUT tool results
"""

from __future__ import annotations

import json
import re
import sys
from collections import Counter
from pathlib import Path

INFRA_PAT = re.compile(
    r"(connection refused|connection reset|econnrefused|server.*(not running|down|crash|unavailable)"
    r"|mcp.*(fail|crash|down|disconnect|error)|failed to (start|connect|spawn|launch)"
    r"|no such file or directory.*(mcp|server)|spawn .* enoent|socket hang up"
    r"|postgresql.*(fail|error)|could not connect|timeout|timed out|broken pipe"
    r"|traceback \(most recent call last\))",
    re.I,
)


def load_json(p: Path):
    try:
        return json.loads(p.read_text(errors="replace"))
    except Exception:
        return None


def scan_wire(wire: Path, acc: Counter, samples: dict, cap=4):
    if not wire.exists():
        return
    for line in wire.read_text(errors="replace").splitlines():
        try:
            j = json.loads(line)
        except Exception:
            continue
        if j.get("type") != "context.append_loop_event":
            continue
        ev = j.get("event", {})
        if ev.get("type") != "tool.result":
            continue
        r = ev.get("result", {})
        out = str(r.get("output") if isinstance(r, dict) else r)
        is_err = bool(r.get("is_error")) if isinstance(r, dict) else False
        low = out.lower()
        if is_err and ("mcp" in low or "connect" in low or "server" in low or "refused" in low):
            acc["mcp_down"] += 1
            samples.setdefault("mcp_down", []).append(out[:160])
        elif "timeout" in low or "timed out" in low or "=== timeout ===" in low:
            acc["timeout_loop"] += 1
            samples.setdefault("timeout_loop", []).append(out[:120])
        elif is_err:
            acc["tool_crash"] += 1
            samples.setdefault("tool_crash", []).append(out[:160])
        elif INFRA_PAT.search(out):
            acc["infra_text"] += 1
            samples.setdefault("infra_text", []).append(out[:160])


def scan_case(case_dir: Path) -> dict:
    inner = None
    for model_dir in case_dir.iterdir():
        if model_dir.is_dir() and model_dir.name.startswith("kimi-code"):
            for c in model_dir.iterdir():
                if c.is_dir() and c.name.startswith("SingleUserTurn"):
                    inner = c
    task, run = case_dir.parent.name, case_dir.name
    row = {"task": task, "run": run}
    acc: Counter = Counter()
    samples: dict[str, list] = {}

    runlog = case_dir / "run.log"
    if runlog.exists():
        text = runlog.read_text(errors="replace")
        if "failed to start isolated postgres" in text or "pg_ctl: could not start" in text:
            acc["pg_fail"] += 1
            samples.setdefault("pg_fail", []).append("PostgreSQL failed to start")
        if "[warn] Enroot rootfs still exists" in text:
            acc["enroot_leftover"] += 1
        if "preprocess] " in text and re.search(r"preprocess.*(error|fail|traceback)", text, re.I):
            acc["preprocess_fail"] += 1
        m = re.search(r"exited rc=(\d+) claim_done=(\w+)", text)
        if m:
            row["claim_done"] = m.group(2) == "True"
            row["kimi_rc"] = int(m.group(1))

    if inner:
        for wire in inner.glob(".kimi_home/sessions/*/*/agents/*/wire.jsonl"):
            scan_wire(wire, acc, samples)
        # dedupe samples
        for k in samples:
            seen, uniq = set(), []
            for s in samples[k]:
                if s not in seen:
                    seen.add(s)
                    uniq.append(s)
                if len(uniq) >= 3:
                    break
            samples[k] = uniq

    row["issues"] = dict(acc)
    row["samples"] = samples
    row["total"] = sum(acc.values())
    return row


def main(argv):
    roots = [Path(a) for a in argv[1:]] or [Path("dumps/kimi-code")]
    all_rows = []
    for root in roots:
        if not root.is_dir():
            continue
        for task_dir in sorted(root.iterdir()):
            if not task_dir.is_dir():
                continue
            for run_dir in sorted(task_dir.iterdir()):
                if run_dir.is_dir() and re.match(r"\d{8}-\d{6}_slot\d+", run_dir.name):
                    all_rows.append((root.name, scan_case(run_dir)))

    flagged = [(rn, r) for rn, r in all_rows if r["total"] > 0]
    print("=" * 78)
    print("INFRA-ISSUE AUDIT — cases where the model fought infra instead of the task")
    print("=" * 78)
    print(f"scanned {len(all_rows)} runs, flagged {len(flagged)}\n")

    cat_totals: Counter = Counter()
    for rn, r in flagged:
        for k, v in r["issues"].items():
            cat_totals[k] += v

    print("## Category totals (result occurrences)")
    for k, v in cat_totals.most_common():
        print(f"  {k:22} {v}")
    print()

    print("## Per-case")
    print(f"{'root':<28}{'task':<40}{'issues':<38}{'claim'}")
    print("-" * 120)
    for rn, r in sorted(flagged, key=lambda x: -x[1]["total"]):
        iss = ",".join(f"{k}:{v}" for k, v in sorted(r["issues"].items(), key=lambda x: -x[1]))
        cd = r.get("claim_done")
        cd_s = "Y" if cd else ("N" if cd is False else "?")
        print(f"{rn:<28}{r['task'][:38]:<40}{iss[:36]:<38}{cd_s}")

    # detailed samples for worst offenders
    print("\n## Samples (top offenders)")
    for rn, r in sorted(flagged, key=lambda x: -x[1]["total"])[:8]:
        print(f"\n### [{rn}] {r['task']} / {r['run']}")
        for k, ss in r["samples"].items():
            for s in ss:
                print(f"  [{k}] {s}")

    out = {f"{rn}/{r['task']}/{r['run']}": r for rn, r in flagged}
    Path("dumps/infra_audit.json").write_text(json.dumps(out, indent=2))
    print("\n[written] dumps/infra_audit.json")


if __name__ == "__main__":
    main(sys.argv)
