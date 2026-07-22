"""Identify at-risk students based on average scores."""
import json
import os


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    with open(os.path.join(here, "student_performance.json")) as f:
        perf = json.load(f)

    risk_assessment = {"per_course": {}, "students": []}
    course_stats = {}
    for s in perf.get("students", []):
        avg = s.get("avg_score", 0)
        if avg < 50:
            level = "Critical"
        elif avg < 65:
            level = "At Risk"
        else:
            level = "On Track"
        s_out = {**s, "risk_level": level}
        risk_assessment["students"].append(s_out)
        c = s.get("course_name", "")
        cs = course_stats.setdefault(c, {"Critical": 0, "At Risk": 0, "On Track": 0})
        cs[level] += 1

    risk_assessment["per_course"] = course_stats
    with open(os.path.join(here, "risk_assessment.json"), "w") as f:
        json.dump(risk_assessment, f, indent=2)
    print("Wrote risk_assessment.json")


if __name__ == "__main__":
    main()
