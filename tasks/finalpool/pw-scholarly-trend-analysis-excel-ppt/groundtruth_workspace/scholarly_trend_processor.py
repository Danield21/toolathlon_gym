#!/usr/bin/env python3
"""Scholarly trend processor.

Reads benchmark and internal scholarly metrics from JSON files in the
workspace and writes scholarly_trend_results.json comparing internal
paper metrics against the external benchmark.
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
    papers = load_json("internal_papers.json")
    rows = []
    for p in papers.get("papers", []):
        bench_cit = benchmark.get("avg_citations_per_paper", 0)
        rows.append({
            "paper_id": p.get("id"),
            "citations": p.get("citations", 0),
            "relevance_score": p.get("relevance_score", 0),
            "benchmark_avg_citations": bench_cit,
            "gap": p.get("citations", 0) - bench_cit,
        })
    rows.sort(key=lambda r: str(r.get("paper_id", "")).lower())
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "scholarly_trend_results.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump({"results": rows}, f, indent=2)
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    run()
