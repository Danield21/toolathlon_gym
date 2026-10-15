#!/usr/bin/env python3
"""Check all 166 case outcomes in a completed run, including missing/infra cases."""
import argparse
from collections import Counter
import csv
import json
from pathlib import Path


def summarize(directory):
    expected = set((Path(__file__).resolve().parent / "tasks.txt").read_text().splitlines())
    files = list(directory.glob("summary_parallel_*.csv"))
    if len(files) != 1:
        raise ValueError("Provide one run directory containing exactly one summary_parallel_*.csv")
    with files[0].open(newline="") as handle:
        rows = list(csv.DictReader(handle))
    counts = Counter(row["task"] for row in rows)
    invalid = sorted(task for task, count in counts.items() if count != 1 or task not in expected)
    missing = sorted(expected - counts.keys())
    statuses = Counter(row["status"] for row in rows)
    incomplete = sorted(row["task"] for row in rows if row["status"] not in ("success", "case_failed"))
    result = {
        "suite": "MS-Eval", "expected": 166, "summary_rows": len(rows),
        "statuses": dict(statuses), "missing": missing,
        "duplicate_or_unexpected": invalid, "infrastructure_or_skipped": incomplete,
        "complete": len(expected) == 166 and not (missing or invalid or incomplete),
        "note": "success/case_failed are harness execution statuses; inspect each case's evaluation artifacts for scores.",
    }
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("run_directory", type=Path)
    args = parser.parse_args()
    try:
        result = summarize(args.run_directory)
        print(json.dumps(result, indent=2))
        raise SystemExit(0 if result["complete"] else 2)
    except (ValueError, OSError, KeyError) as exc:
        parser.exit(2, f"Error: {exc}\n")
