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


def _norm_loose(s):
    """Lowercase, replace punctuation with spaces, collapse whitespace."""
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", str(s or "").lower())).strip()


def _levenshtein(a, b):
    """Plain Levenshtein edit distance (titles are short)."""
    if a == b:
        return 0
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[-1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def _titles_match(expected, actual):
    """Exact title match (task.md prescribes the exact database title).
    After normalization the titles must be identical — no containment /
    edit-distance / keyword fuzzy tolerance. Numeric tokens (years, term
    codes) are part of the normalized string, so e.g. 'Fall 2013' never
    matches 'Fall 2014'."""
    e = _norm_loose(expected)
    a = _norm_loose(actual)
    if not e or not a:
        return False
    return e == a


def _match_course_key(title_text, expected_by_name):
    """Map a page title to an expected course key. task.md requires each
    entry's title to be the course name exactly (extra identifiers such as
    term codes may be appended, but the full course name must remain intact).
    So: exact normalized match first, else the entry whose full normalized
    course name appears intact in the title. No edit-distance fuzzy."""
    tnorm = _normalize(title_text)
    if tnorm in expected_by_name:
        return tnorm
    tloose = _norm_loose(title_text)
    for name_norm in expected_by_name:
        if _norm_loose(name_norm) in tloose:
            return name_norm
    return None


def get_expected():
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    # Only student enrollment counts are required (per task.md). Match the
    # enrollment type exactly to avoid incidental substring matches on other
    # enrollment types (e.g. Observer, Designer) should they appear later.
    cur.execute("""SELECT e.course_id, c.name, COUNT(*)
        FROM canvas.enrollments e JOIN canvas.courses c ON c.id=e.course_id
        WHERE e.type = 'StudentEnrollment'
        GROUP BY e.course_id, c.name ORDER BY e.course_id""")
    courses = {}
    for cid, name, cnt in cur.fetchall():
        courses[cid] = {"name": name, "students": cnt}
    conn.close()
    return courses


def check_notion(expected):
    print("\n=== Checking Notion Database ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("SELECT id, title FROM notion.databases WHERE archived=false")
    dbs = cur.fetchall()
    record("Notion database created", len(dbs) >= 1, f"Found {len(dbs)}")

    # Find database titled 'Course Enrollment Tracker' (loose title match:
    # containment / edit distance <= 3 / all keywords present)
    expected_title = "course enrollment tracker"
    target_db = None
    for did, title_json in dbs:
        title_str = _extract_text(title_json)
        if _titles_match(expected_title, title_str):
            target_db = did
            break
    # Strict: must find by name; no fallback to dbs[0]
    record("Database titled 'Course Enrollment Tracker'", target_db is not None,
           f"Looking for title 'Course Enrollment Tracker' (loose match); dbs={[_extract_text(d[1]) for d in dbs]}")

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

    # Check per-page: title matches a course (loose match allowed); counts correct
    matched_keys = set()
    correct_props = 0
    pages_with_props_checked = 0
    seen_titles = []
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
        seen_titles.append(title_text)
        # Loose match: title must contain the course name core (or vice versa),
        # be within edit distance 3, or contain all course-name keywords
        key = _match_course_key(title_text, expected_by_name)
        if key is not None:
            matched_keys.add(key)
            exp = expected_by_name[key]
            # Verify the student count only (per task.md)
            student_val = None
            for pname, pval in props.items():
                if not isinstance(pval, dict): continue
                lname = _normalize(pname).replace(" ", "_")
                if pval.get("type") == "number":
                    val = pval.get("number")
                    if lname in ("student_count", "studentcount", "students", "student"):
                        student_val = val
            pages_with_props_checked += 1
            ok = (student_val == exp["students"])
            if ok:
                correct_props += 1

    # Strict: every course must be present and counts must match
    matched_titles = len(matched_keys)
    record("All course names present as page titles",
           matched_titles == expected_count,
           f"Matched {matched_titles}/{expected_count}; titles seen: {seen_titles}")
    record("All page student counts correct",
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
    total_students = sum(c["students"] for c in expected.values())

    record("Body mentions total course count",
           re.search(rf"\b{total_courses}\b", body_str) is not None,
           f"Expected {total_courses} in body")

    # Total student enrollment could be displayed with or without commas
    te_str = str(total_students)
    te_str_with_comma = f"{total_students:,}"
    record("Body mentions total student enrollment",
           te_str in body_str or te_str_with_comma in body_str,
           f"Expected {te_str} (or {te_str_with_comma}) in body")

    # Top 5 by student enrollment
    sorted_courses = sorted(expected.values(), key=lambda c: c["students"], reverse=True)
    top5 = sorted_courses[:5]
    matched_top = 0
    for c in top5:
        cname = c["name"]
        # try short name (before paren) too
        short = cname.split("(")[0].strip()
        if short.lower() in body_lower or cname.lower() in body_lower:
            matched_top += 1
    record("Body lists top 5 courses by student enrollment",
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
