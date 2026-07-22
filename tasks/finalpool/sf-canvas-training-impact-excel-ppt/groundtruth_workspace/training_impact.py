"""Training impact analysis script - reads employee/course data and outputs impact_analysis.json."""
import json
import os

def analyze():
    # Load employee data
    emp_path = "employee_data.json"
    course_path = "course_data.json"
    employees = []
    courses = []
    if os.path.exists(emp_path):
        with open(emp_path) as f:
            employees = json.load(f)
    if os.path.exists(course_path):
        with open(course_path) as f:
            courses = json.load(f)

    # Compute aggregations
    dept_perf = {}
    for e in employees:
        d = e.get("department")
        r = e.get("performance_rating")
        if d and r is not None:
            dept_perf.setdefault(d, []).append(r)
    dept_avg = {d: sum(v)/len(v) for d, v in dept_perf.items() if v}

    out = {
        "total_employees": len(employees),
        "total_courses": len(courses),
        "dept_avg_performance": dept_avg,
        "highest_performing_dept": max(dept_avg, key=dept_avg.get) if dept_avg else None,
        "lowest_performing_dept": min(dept_avg, key=dept_avg.get) if dept_avg else None,
    }
    with open("impact_analysis.json", "w") as f:
        json.dump(out, f, indent=2)
    print("Wrote impact_analysis.json")
    return out

if __name__ == "__main__":
    analyze()
