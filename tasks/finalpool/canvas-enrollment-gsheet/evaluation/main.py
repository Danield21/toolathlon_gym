"""
Evaluation script for canvas-enrollment-gsheet task.

Checks:
1. Google Sheet "Fall 2014 Enrollment Tracker" exists with correct enrollment data
2. Email sent to planning@university.edu about under-enrolled courses
"""
import argparse
import json
import os
import sys

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0


def load_expected():
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("""
        SELECT c.name, c.course_code, COUNT(e.id) as enrollment_count
        FROM canvas.courses c
        LEFT JOIN canvas.enrollments e ON c.id = e.course_id
        WHERE c.name LIKE '%%Fall 2014%%'
        GROUP BY c.name, c.course_code
        ORDER BY COUNT(e.id) DESC
    """)
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return [{"name": r[0], "code": r[1], "count": int(r[2])} for r in rows]


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        detail_str = f": {str(detail)[:200]}" if detail else ""
        print(f"  [FAIL] {name}{detail_str}")


def check_gsheet():
    print("\n=== Checking Google Sheet ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("""
        SELECT id, title FROM gsheet.spreadsheets
        WHERE LOWER(title) LIKE '%%fall 2014%%' AND LOWER(title) LIKE '%%enrollment%%'
    """)
    sheets = cur.fetchall()
    check("Google Sheet with enrollment tracker exists",
          len(sheets) >= 1,
          f"Found {len(sheets)} matching spreadsheets")

    expected = load_expected()

    if sheets:
        ss_id = sheets[0][0]
        cur.execute("""
            SELECT c.row_index, c.col_index, c.value
            FROM gsheet.cells c
            WHERE c.spreadsheet_id = %s
            ORDER BY c.row_index, c.col_index
        """, (ss_id,))
        cells = cur.fetchall()
        grid = {}
        for r, col, v in cells:
            grid.setdefault(r, {})[col] = str(v) if v is not None else ''

        # Header detection
        header_row = None
        for r, row in sorted(grid.items()):
            joined = ' '.join(row.values()).lower()
            if 'course' in joined and ('code' in joined or 'name' in joined) and ('enroll' in joined or 'count' in joined):
                header_row = r
                break
        check('GSheet header row has Course Name/Code/Enrollment Count', header_row is not None, f'Headers not found')

        if header_row is not None:
            header_cells = grid[header_row]
            name_col = code_col = count_col = None
            for col, val in header_cells.items():
                vl = val.lower()
                if 'name' in vl:
                    name_col = col
                elif 'code' in vl:
                    code_col = col
                elif 'enroll' in vl or 'count' in vl:
                    count_col = col
            data_rows = [grid[r] for r in sorted(grid.keys()) if r > header_row]
            prev_count = None
            order_ok = True
            found_codes = []
            for row in data_rows:
                code = row.get(code_col, '').strip() if code_col is not None else ''
                try:
                    cnt = int(float(row.get(count_col, 0)))
                except (ValueError, TypeError):
                    continue
                if prev_count is not None and cnt > prev_count:
                    order_ok = False
                prev_count = cnt
                found_codes.append((code, cnt))
            check('GSheet rows are sorted by enrollment count descending', order_ok, 'Rows not in descending order')

            for course in expected:
                match = any(code.lower() == course['code'].lower() and abs(cnt - course['count']) <= 0 for code, cnt in found_codes)
                check(f"GSheet has correct row for '{course['code']}' with count {course['count']}",
                      match, f'Not found in structured rows')

    cur.close()
    conn.close()


def check_email():
    print("\n=== Checking Email ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("""
        SELECT subject, from_addr, to_addr, body_text
        FROM email.messages
    """)
    all_emails = cur.fetchall()

    def parse_recipients(to_addr):
        if to_addr is None:
            return []
        if isinstance(to_addr, list):
            return [str(r).strip().lower() for r in to_addr]
        to_str = str(to_addr).strip()
        try:
            parsed = json.loads(to_str)
            if isinstance(parsed, list):
                return [str(r).strip().lower() for r in parsed]
            return [to_str.lower()]
        except (json.JSONDecodeError, TypeError):
            return [to_str.lower()]

    target_email = "planning@university.edu"
    found = None
    for subj, from_addr, to_addr, body in all_emails:
        recipients = parse_recipients(to_addr)
        if target_email in recipients:
            found = (subj, from_addr, to_addr, body)
            break

    check(f"Email sent to {target_email}", found is not None,
          f"Found {len(all_emails)} total emails")

    if found:
        subj, from_addr, to_addr, body = found
        subj_lower = (subj or "").lower()
        body_lower = (body or "").lower()

        check("Email subject mentions 'under-enrolled'",
              "under" in subj_lower and "enroll" in subj_lower,
              f"Subject: {(subj or '')[:100]}")

        check("Email subject mentions 'Fall 2014'",
              "fall 2014" in subj_lower,
              f"Subject: {(subj or '')[:100]}")

        # Check under-enrolled courses mentioned in body
        expected = load_expected()
        under_enrolled = [c for c in expected if c["count"] < 1000]
        for course in under_enrolled:
            # Accept either course name or code
            name_parts = course["name"].lower().split("(")[0].strip()
            check(f"Email body mentions under-enrolled course '{course['code']}'",
                  course["code"].lower() in body_lower or name_parts in body_lower,
                  f"Not found in body")
            check(f"Email body mentions count {course['count']}",
                  str(course["count"]) in (body or ""),
                  f"Count not found in body")

    cur.close()
    conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_gsheet()
    check_email()

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    all_passed = FAIL_COUNT == 0
    print(f"  Overall: {'PASS' if all_passed else 'FAIL'}")
    sys.exit(0 if all_passed else 1)


if __name__ == "__main__":
    main()
