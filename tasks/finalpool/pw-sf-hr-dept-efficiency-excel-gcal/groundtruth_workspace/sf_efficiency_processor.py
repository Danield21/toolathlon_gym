#!/usr/bin/env python3
"""SF HR efficiency processor.

Reads benchmark and internal HR data from JSON files in the workspace
and writes sf_efficiency_results.json comparing internal department
metrics with the external industry benchmark.
"""
import json
import os


def load_json(name):
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), name)
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def run():
    benchmark = load_json("benchmark.json")
    internal = load_json("internal_hr.json")
    rows = []
    for dept, metrics in internal.get("departments", {}).items():
        bench = benchmark.get("departments", {}).get(dept, {})
        rows.append({
            "department": dept,
            "our_avg_salary": metrics.get("avg_salary", 0),
            "industry_benchmark": bench.get("avg_salary", 0),
            "salary_gap": (metrics.get("avg_salary", 0)
                          - bench.get("avg_salary", 0)),
        })
    rows.sort(key=lambda r: (r.get("department") or "").lower())
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "sf_efficiency_results.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({"results": rows}, f, indent=2)
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    run()
