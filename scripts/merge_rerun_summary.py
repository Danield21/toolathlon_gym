#!/usr/bin/env python3
"""Merge rerun summary rows back into the original full-run summary.

The runner writes each launch's results to summary_parallel_<RUN_ID>.csv
inside the shared dump root. When a rerun (e.g. fix6) targets the same tasks
as an earlier run, the canonical full-run CSV should reflect the LATEST
attempt per task so audit_index/summary stay consistent with the newest
slot directories on disk.

Usage:
    python3 scripts/merge_rerun_summary.py <dump_root> <base_run_id> <rerun_run_id> [...]

The base summary file is rewritten in place (task,status,exit_code,output_dir,
pg_port,duration_s columns preserved); rows for tasks absent from the rerun
keep their original values. A backup summary_parallel_<base>.csv.bak is kept.
"""
from __future__ import annotations

import csv
import shutil
import sys
from pathlib import Path


def load_csv(path: Path) -> list[dict]:
    with path.open(newline="") as f:
        return list(csv.DictReader(f))


def main(argv: list[str]) -> int:
    if len(argv) < 4:
        print(__doc__)
        return 1
    dump_root = Path(argv[1])
    base_run = argv[2]
    rerun_runs = argv[3:]

    base_path = dump_root / f"summary_parallel_{base_run}.csv"
    if not base_path.exists():
        print(f"[error] missing {base_path}")
        return 1

    rows = load_csv(base_path)
    by_task = {r["task"]: r for r in rows}
    base_order = [r["task"] for r in rows]

    merged = 0
    new_tasks = []
    for run_id in rerun_runs:
        p = dump_root / f"summary_parallel_{run_id}.csv"
        if not p.exists():
            print(f"[warn] missing {p}; skipped")
            continue
        for r in load_csv(p):
            t = r["task"]
            if t in by_task:
                # newer attempt wins; keep the original row position
                old = by_task[t]
                old.update(r)
                merged += 1
            else:
                by_task[t] = r
                new_tasks.append(t)

    backup = base_path.with_suffix(".csv.bak")
    if not backup.exists():
        shutil.copy2(base_path, backup)
        print(f"[backup] {backup}")

    with base_path.open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["task", "status", "exit_code", "output_dir", "pg_port", "duration_s"])
        w.writeheader()
        for t in base_order:
            w.writerow(by_task[t])
        for t in new_tasks:
            w.writerow(by_task[t])

    print(f"[merged] {merged} rows updated, {len(new_tasks)} new tasks appended, total {len(base_order) + len(new_tasks)}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
