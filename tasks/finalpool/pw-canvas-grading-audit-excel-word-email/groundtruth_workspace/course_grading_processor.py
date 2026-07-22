#!/usr/bin/env python3
"""Course grading audit processor.

Reads benchmark and internal data from JSON files in the workspace and
emits course_grading_results.json comparing internal metrics with
external benchmarks.
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
    benchmark = load_json("benchmark_data.json")
    internal = load_json("internal_data.json")
    courses = []
    for course_code, internal_metrics in internal.get("courses", {}).items():
        bench = benchmark.get("courses", {}).get(course_code, {})
        courses.append({
            "code": course_code,
            "name": internal_metrics.get("name", course_code),
            "enrollment": internal_metrics.get("enrollment", 0),
            "avg_score": internal_metrics.get("avg_score", 0),
            "pass_rate": internal_metrics.get("pass_rate", 0),
            "benchmark_avg_score": bench.get("avg_score", 0),
            "benchmark_pass_rate": bench.get("pass_rate", 0),
            "score_gap": (internal_metrics.get("avg_score", 0)
                         - bench.get("avg_score", 0)),
        })
    courses.sort(key=lambda c: c["code"].lower())
    out = {
        "courses": courses,
        "summary": {
            "course_count": len(courses),
            "average_pass_rate": (sum(c["pass_rate"] for c in courses)
                                  / max(len(courses), 1)),
        },
    }
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "course_grading_results.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(out, f, indent=2)
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    run()
