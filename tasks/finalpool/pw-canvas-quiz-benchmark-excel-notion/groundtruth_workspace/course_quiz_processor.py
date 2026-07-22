"""Course quiz processor.

Reads collected JSON data files (canvas course/quiz exports and external
benchmark data), computes per-course comparisons against benchmark, then
writes course_quiz_results.json.
"""
import json
import os


def load_json(path):
    if os.path.exists(path):
        with open(path, "r") as f:
            return json.load(f)
    return {}


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    canvas = load_json(os.path.join(here, "canvas_data.json"))
    external = load_json(os.path.join(here, "external_benchmark.json"))

    courses = canvas.get("courses", [])
    benchmarks = {b.get("course_code"): b for b in external.get("benchmarks", [])}

    results = []
    for c in courses:
        code = c.get("code")
        bench = benchmarks.get(code, {})
        results.append({
            "course": c.get("name"),
            "code": code,
            "enrollment": c.get("enrollment"),
            "avg_score": c.get("avg_score"),
            "pass_rate": c.get("pass_rate"),
            "benchmark_avg": bench.get("avg_score"),
        })

    out_path = os.path.join(here, "course_quiz_results.json")
    with open(out_path, "w") as f:
        json.dump({"results": results}, f, indent=2)


if __name__ == "__main__":
    main()
