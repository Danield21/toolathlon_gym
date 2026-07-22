"""Evaluation for canvas-enrollment-notion."""
import argparse
import json
import os
import re
import sys

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="toolathlon_gym", user="eigent", password="camel")
PASS_COUNT = 0
FAIL_COUNT = 0


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1; print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1; print(f"  [FAIL] {name}: {str(detail)[:300]}")


def num_close(a, b, tol=1.0):
    try: return abs(float(a) - float(b)) <= tol
    except: return False


def _extract_text(json_obj):
    """Extract plain text from a Notion title/rich_text JSON."""
    if json_obj is None:
        return ""
    if isinstance(json_obj, str):
        return json_obj
    if isinstance(json_obj, list):
        out = []
        for item in json_obj:
            if isinstance(item, dict):
                t = item.get("plain_text") or (item.get("text") or {}).get("content") or ""
                out.append(t)
            elif isinstance(item, str):
                out.append(item)
        return "".join(out)
    if isinstance(json_obj, dict):
        # Could be {"title":[...]} or other
        for k in ("title", "rich_text", "plain_text"):
            if k in json_obj:
                return _extract_text(json_obj[k])
    return ""


def _normalize(s):
    return re.sub(r"\s+", " ", str(s or "").strip().lower())


def get_expected():
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""SELECT e.course_id, c.name, e.type, COUNT(*)
        FROM canvas.enrollments e JOIN canvas.courses c ON c.id=e.course_id
        GROUP BY e.course_id, c.name, e.type ORDER BY e.course_id""")
    courses = {}
    for cid, name, etype, cnt in cur.fetchall():
        if cid not in courses:
            courses[cid] = {"name": name, "students": 0, "teachers": 0, "tas": 0}
        if "Student" in etype: courses[cid]["students"] = cnt
        elif "Teacher" in etype: courses[cid]["teachers"] = cnt
        elif "Ta" in etype: courses[cid]["tas"] = cnt
    for c in courses.values():
        c["total"] = c["students"] + c["teachers"] + c["tas"]
    conn.close()
    return courses


def check_notion(expected):
    print("\n=== Checking Notion Database ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT id, title FROM notion.databases WHERE archived=false")
    dbs = cur.fetchall()
    record("Notion database created", len(dbs) >= 1, f"Found {len(dbs)}")

    # Find database with exact (or near-exact) title 'Course Enrollment Tracker'
    expected_title = "course enrollment tracker"
    target_db = None
    for did, title_json in dbs:
        title_str = _normalize(_extract_text(title_json))
        if title_str == expected_title:
            target_db = did
            break
    # Strict: must find by name; no fallback to dbs[0]
    record("Database titled 'Course Enrollment Tracker'", target_db is not None,
           f"Looking for exact title; dbs={[_extract_text(d[1]) for d in dbs]}")

    if target_db is None:
        cur.close()
        conn.close()
        return

    # Pages
    cur.execute("""SELECT id, properties FROM notion.pages
        WHERE parent->>'database_id' = %s AND archived=false""", (target_db,))
    pages = cur.fetchall()
    expected_count = len(expected)
    record(f"Database has all course entries (expected {expected_count})",
           len(pages) == expected_count,
           f"Found {len(pages)}, expected {expected_count}")

    # Build name -> expected map (lower)
    expected_by_name = {}
    for c in expected.values():
        expected_by_name[_normalize(c["name"])] = c

    # Check per-page: title matches a course; counts correct
    matched_titles = 0
    correct_props = 0
    pages_with_props_checked = 0
    for pid, props in pages:
        if not props:
            continue
        # Find the title property value
        title_text = ""
        # props is dict of {propname: {type:..., title:..., number:..., ...}}
        for pname, pval in props.items():
            if isinstance(pval, dict):
                ptype = pval.get("type")
                if ptype == "title":
                    title_text = _extract_text(pval.get("title", []))
                    break
        title_norm = _normalize(title_text)
        if title_norm in expected_by_name:
            matched_titles += 1
            exp = expected_by_name[title_norm]
            # Verify counts
            student_val = teacher_val = ta_val = total_val = None
            for pname, pval in props.items():
                if not isinstance(pval, dict): continue
                lname = _normalize(pname).replace(" ", "_")
                if pval.get("type") == "number":
                    val = pval.get("number")
                    # Match exact normalized property names per task.md
                    if lname in ("student_count", "studentcount", "students", "student"):
                        student_val = val
                    elif lname in ("teacher_count", "teachercount", "teachers", "teacher"):
                        teacher_val = val
                    elif lname in ("ta_count", "tacount", "tas", "ta"):
                        ta_val = val
                    elif lname in ("total_enrollment", "totalenrollment", "total"):
                        total_val = val
            pages_with_props_checked += 1
            ok = (
                student_val == exp["students"]
                and teacher_val == exp["teachers"]
                and ta_val == exp["tas"]
                and total_val == exp["total"]
            )
            if ok:
                correct_props += 1

    # Strict: every course must be present and counts must match
    record("All course names present as page titles",
           matched_titles == expected_count,
           f"Matched {matched_titles}/{expected_count}")
    record("All page property counts correct",
           correct_props == expected_count,
           f"Correct {correct_props}/{expected_count}")

    cur.close()
    conn.close()


def check_email(expected):
    print("\n=== Checking Email ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT subject, to_addr, body_text FROM email.messages")
    emails = cur.fetchall()
    record("At least 1 email sent", len(emails) >= 1, f"Found {len(emails)}")

    # Find the summary email by exact subject
    expected_subject = "Course Enrollment Summary Report"
    summary_email = None
    for subj, to, body in emails:
        if (subj or "").strip().lower() == expected_subject.lower():
            summary_email = (subj, to, body)
            break

    record("Email subject exactly 'Course Enrollment Summary Report'",
           summary_email is not None,
           f"Subjects seen: {[e[0] for e in emails]}")

    if summary_email is None:
        cur.close()
        conn.close()
        return

    subj, to, body = summary_email
    to_str = json.dumps(to).lower() if isinstance(to, list) else str(to).lower()
    record("Email to admin@university.example.com",
           "admin@university.example.com" in to_str, f"To: {to}")

    body_str = body or ""
    body_lower = body_str.lower()
    total_courses = len(expected)
    total_enrollment = sum(c["total"] for c in expected.values())

    record("Body mentions total course count",
           re.search(rf"\b{total_courses}\b", body_str) is not None,
           f"Expected {total_courses} in body")

    # Total enrollment number could be displayed with or without commas
    te_str = str(total_enrollment)
    te_str_with_comma = f"{total_enrollment:,}"
    record("Body mentions total enrollment",
           te_str in body_str or te_str_with_comma in body_str,
           f"Expected {te_str} (or {te_str_with_comma}) in body")

    # Top 5 by total
    sorted_courses = sorted(expected.values(), key=lambda c: c["total"], reverse=True)
    top5 = sorted_courses[:5]
    matched_top = 0
    for c in top5:
        cname = c["name"]
        # try short name (before paren) too
        short = cname.split("(")[0].strip()
        if short.lower() in body_lower or cname.lower() in body_lower:
            matched_top += 1
    record("Body lists top 5 courses by total enrollment",
           matched_top == 5,
           f"Matched {matched_top}/5; top names: {[c['name'] for c in top5]}")

    cur.close()
    conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", default=".")
    parser.add_argument("--groundtruth_workspace", default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()
    expected = get_expected()
    check_notion(expected)
    check_email(expected)
    print(f"\n=== SUMMARY: {PASS_COUNT} passed, {FAIL_COUNT} failed ===")
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump({"passed": PASS_COUNT, "failed": FAIL_COUNT, "success": FAIL_COUNT == 0}, f)
    sys.exit(0 if FAIL_COUNT == 0 else 1)

if __name__ == "__main__":
    main()
