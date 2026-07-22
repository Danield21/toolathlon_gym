#!/usr/bin/env python3
"""Read grade_data.json, compute distributions and pass rates per course, write grade_report.json."""
import json
import os


def grade_letter(score):
    if score >= 90:
        return "A"
    if score >= 80:
        return "B"
    if score >= 70:
        return "C"
    if score >= 60:
        return "D"
    return "F"


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    in_path = os.path.join(here, "grade_data.json")
    out_path = os.path.join(here, "grade_report.json")
    with open(in_path) as f:
        data = json.load(f)

    courses_out = []
    total_students = 0
    weighted_avg_sum = 0.0
    weighted_pass_count = 0
    course_avgs = {}

    for course in data["courses"]:
        scores = course["scores"]
        n = len(scores)
        a = sum(1 for s in scores if s >= 90)
        b = sum(1 for s in scores if 80 <= s < 90)
        c = sum(1 for s in scores if 70 <= s < 80)
        d = sum(1 for s in scores if 60 <= s < 70)
        f = sum(1 for s in scores if s < 60)
        pass_count = a + b + c
        avg = sum(scores) / n if n else 0
        pass_rate = round(100.0 * pass_count / n, 1) if n else 0
        course_avgs[course["name"]] = avg
        courses_out.append({
            "name": course["name"],
            "A_Count": a, "B_Count": b, "C_Count": c, "D_Count": d, "F_Count": f,
            "Total_Students": n,
            "Pass_Rate_Pct": pass_rate,
            "Course_Avg": round(avg, 1),
        })
        total_students += n
        weighted_avg_sum += avg * n
        weighted_pass_count += pass_count

    overall_pass = round(100.0 * weighted_pass_count / total_students, 1) if total_students else 0
    overall_avg = round(weighted_avg_sum / total_students, 1) if total_students else 0
    highest = max(course_avgs, key=course_avgs.get)
    lowest = min(course_avgs, key=course_avgs.get)

    out = {
        "courses": courses_out,
        "summary": {
            "Total_Courses": len(courses_out),
            "Total_Students": total_students,
            "Overall_Pass_Rate": overall_pass,
            "Overall_Avg_Grade": overall_avg,
            "Highest_Avg_Course": highest,
            "Lowest_Avg_Course": lowest,
        },
    }
    with open(out_path, "w") as f:
        json.dump(out, f, indent=2)


if __name__ == "__main__":
    main()
