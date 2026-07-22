"""Salary processor.

Reads benchmark_raw.json (industry benchmark) and internal_salaries.json
(company HR data), cross-references by department, and outputs
salary_comparison.json with merged analysis.
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
    bench = load_json(os.path.join(here, "benchmark_raw.json"))
    internal = load_json(os.path.join(here, "internal_salaries.json"))

    bench_map = {b.get("department"): b for b in bench.get("benchmarks", [])}

    rows = []
    for dept_row in internal.get("departments", []):
        dept = dept_row.get("department")
        b = bench_map.get(dept, {})
        our = dept_row.get("avg_salary")
        ind = b.get("benchmark")
        diff = (our or 0) - (ind or 0)
        rows.append({
            "department": dept,
            "employee_count": dept_row.get("employee_count"),
            "our_avg_salary": our,
            "industry_benchmark": ind,
            "difference": round(diff, 2),
            "difference_pct": round(diff / ind * 100, 1) if ind else None,
            "status": "Above" if diff >= 0 else "Below",
        })

    out_path = os.path.join(here, "salary_comparison.json")
    with open(out_path, "w") as f:
        json.dump({"rows": rows}, f, indent=2)


if __name__ == "__main__":
    main()
