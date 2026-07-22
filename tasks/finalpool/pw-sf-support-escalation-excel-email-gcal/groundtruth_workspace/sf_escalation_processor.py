#!/usr/bin/env python3
"""SF support escalation processor.

Reads benchmark and internal support metrics from JSON files in the
workspace and writes sf_escalation_results.json comparing internal
priority response times with the external industry benchmark.
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
    internal = load_json("internal_tickets.json")
    rows = []
    for priority, metrics in internal.get("priorities", {}).items():
        bench = benchmark.get("priorities", {}).get(priority, {})
        rows.append({
            "priority": priority,
            "our_avg_response": metrics.get("avg_response", 0),
            "industry_avg": bench.get("avg_response", 0),
            "gap": (metrics.get("avg_response", 0)
                   - bench.get("avg_response", 0)),
        })
    rows.sort(key=lambda r: (r.get("priority") or "").lower())
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "sf_escalation_results.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({"results": rows}, f, indent=2)
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    run()
