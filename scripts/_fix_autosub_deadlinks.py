#!/usr/bin/env python3
"""One-off: fix dead-link output_dir in
dumps/deepseek-v4-flash-0731/auto_subagent/summary_parallel_latest.csv.

All 167 rows point into the non-existent mirror tree
dumps/kimi-code_deepseek-v4-flash-subagent/. Remap each to the real slot under
auto_subagent/{no_subagent,need_subagent/...}. Special-case
financial-audit-cross-system-reconciliation: its row is updated to the autorun
rerun (case_failed) and its old success slot was already deleted.
"""
import csv
import os
import re
import sys

A = "/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/dumps/deepseek-v4-flash-0731/auto_subagent"
SUM = f"{A}/summary_parallel_latest.csv"
SLOT_RE = re.compile(r"^(subagent|saf12|autorun)-\d{8}-\d{6}_slot\d+$")

# case -> {slotname: fullpath}
slotmap = {}
for r, ds, _ in os.walk(A):
    for d in ds:
        if SLOT_RE.match(d):
            case = os.path.basename(r)
            slotmap.setdefault(case, {})[d] = os.path.join(r, d)

FIN = "financial-audit-cross-system-reconciliation"
FIN_ROW = [
    FIN, "case_failed", "1",
    slotmap[FIN]["autorun-20260822-212016_slot0"],
    "51000", "2192",
]

rows = list(csv.reader(open(SUM)))
hdr, body = rows[0], rows[1:]
out = [hdr]
hit = fixed = 0
problems = []
for row in body:
    task = row[0]
    if task == FIN:
        out.append(FIN_ROW)
        fixed += 1
        continue
    m = re.search(r"/([^/]+)/((?:subagent|saf12|autorun)-\d{8}-\d{6}_slot\d+)$", row[3])
    if not m:
        problems.append(("noparse", task, row[3]))
        out.append(row)
        continue
    case, slot = m.group(1), m.group(2)
    real = slotmap.get(case, {}).get(slot)
    if real:
        row[3] = real
        hit += 1
    else:
        problems.append(("noslot", task, slot))
    out.append(row)

print(f"rows={len(body)} remapped={hit} financial_updated={fixed} problems={len(problems)}")
for p in problems:
    print("  PROBLEM", p)

# verify every written output_dir exists
dead = [r for r in out[1:] if not os.path.isdir(r[3])]
print(f"dead_after={len(dead)}")
for r in dead[:10]:
    print("  DEAD", r[0], r[3])

if "--write" in sys.argv and not problems and not dead:
    bak = SUM + ".bak_deadlinks"
    if not os.path.exists(bak):
        os.rename(SUM, bak)
        print("backup ->", bak)
    else:
        print("backup exists:", bak)
    with open(SUM, "w", newline="") as f:
        csv.writer(f).writerows(out)
    print("WROTE", SUM)
elif "--write" in sys.argv:
    print("REFUSED to write: unresolved problems/dead links", file=sys.stderr)
    sys.exit(1)
else:
    print("dry-run (pass --write to apply)")
