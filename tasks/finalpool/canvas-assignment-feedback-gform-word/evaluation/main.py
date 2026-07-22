"""Evaluation for canvas-assignment-feedback-gform-word."""
import argparse
import json
import os
import re
import sys

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432,
          dbname="toolathlon_gym", user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        detail_str = f": {detail[:200]}" if detail else ""
        print(f"  [FAIL] {name}{detail_str}")


def num_close(a, b, tol):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def _fetch_expected_from_db():
    """Fetch expected assignment statistics from canvas DB."""
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        # Course 16, exclude assignments with 0 submissions (e.g. Final Exam never submitted)
        cur.execute("""
            SELECT a.name,
                   COUNT(s.id) AS total,
                   ROUND(AVG(s.score)::numeric, 2)::float AS avg_score,
                   SUM(CASE WHEN s.late=true THEN 1 ELSE 0 END)::int AS late,
                   ROUND(SUM(CASE WHEN s.late=true THEN 1 ELSE 0 END) * 100.0 / NULLIF(COUNT(s.id),0), 2)::float AS late_rate
            FROM canvas.assignments a
            LEFT JOIN canvas.submissions s ON s.assignment_id=a.id
            WHERE a.course_id=16
            GROUP BY a.name
            HAVING COUNT(s.id) > 0
            ORDER BY a.name
        """)
        rows = cur.fetchall()
        cur.close()
        conn.close()
        return [(r[0], int(r[1]), float(r[2]) if r[2] is not None else None,
                 int(r[3]), float(r[4]) if r[4] is not None else None) for r in rows]
    except Exception as e:
        print(f"[warn] DB fetch fallback: {e}")
        return None


# Fallback in case DB unavailable
EXPECTED_ASSIGNMENTS_FALLBACK = [
    ("CMA 34878", 1470, 83.32, 36, 2.45),
    ("CMA 34879", 1352, 87.97, 43, 3.18),
    ("CMA 34880", 1252, 76.50, 61, 4.87),
    ("CMA 34881", 1224, 78.30, 59, 4.82),
    ("CMA 34882", 1193, 78.65, 75, 6.29),
    ("CMA 34883", 1196, 76.80, 119, 9.95),
    ("CMA 34884", 1160, 76.16, 88, 7.59),
    ("TMA 34873", 1859, 78.16, 157, 8.45),
    ("TMA 34874", 1661, 72.45, 193, 11.62),
    ("TMA 34875", 1402, 70.49, 376, 26.82),
    ("TMA 34876", 1313, 71.05, 268, 20.41),
    ("TMA 34877", 1158, 75.89, 148, 12.78),
]


def _norm_header(h):
    return re.sub(r'[^a-z0-9]', '', (h or "").lower())


def check_word(agent_workspace, expected):
    """Check Word document output."""
    print("\n=== Checking Word Document ===")
    try:
        from docx import Document
    except ImportError:
        check("python-docx available", False, "python-docx not installed")
        return

    agent_file = os.path.join(agent_workspace, "Assignment_Analysis.docx")
    if not os.path.isfile(agent_file):
        check("Word file exists", False, f"Expected {agent_file}")
        # Subordinate failures
        for label in ["Document title present", "Document has at least one table",
                      "Column 'Assignment_Name' present", "Column 'Total_Submissions' present",
                      "Column 'Avg_Score' present", "Column 'Late_Submissions' present",
                      "Column 'Late_Rate(%)' present",
                      "Table has expected number of data rows",
                      "All expected assignments present with correct numeric values",
                      "Rows are sorted alphabetically by Assignment_Name"]:
            check(label, False, "Word file missing")
        return
    check("Word file exists", True)

    try:
        doc = Document(agent_file)
    except Exception as e:
        check("Word file readable", False, str(e))
        return

    # Check title
    full_text = "\n".join(p.text for p in doc.paragraphs)
    title_ok = (
        "foundations of finance" in full_text.lower() and
        "fall 2013" in full_text.lower() and
        "assignment analysis" in full_text.lower()
    )
    check("Document title present", title_ok,
          f"Title not found in: {full_text[:200]}")

    # Find table
    if not doc.tables:
        check("Document has at least one table", False, "no table found")
        return
    check("Document has at least one table", True)

    tbl = doc.tables[0]

    # Check headers (exact normalized match)
    if not tbl.rows:
        check("Table has rows", False, "Empty table")
        return

    headers = [c.text.strip() for c in tbl.rows[0].cells]
    headers_norm = [_norm_header(h) for h in headers]

    expected_headers_norm = ["assignmentname", "totalsubmissions", "avgscore",
                             "latesubmissions", "laterate"]
    for col_norm, label in zip(expected_headers_norm,
                               ["Assignment_Name", "Total_Submissions", "Avg_Score",
                                "Late_Submissions", "Late_Rate(%)"]):
        present = any(h.startswith(col_norm) for h in headers_norm) or col_norm in headers_norm
        check(f"Column '{label}' present", present, f"Headers: {headers}")

    # Find column indices for header-based extraction
    def _col_index(target_norm_prefix):
        for i, h in enumerate(headers_norm):
            if h == target_norm_prefix or h.startswith(target_norm_prefix):
                return i
        return None

    col_name_idx = _col_index("assignmentname") or 0
    col_total_idx = _col_index("totalsubmissions")
    col_avg_idx = _col_index("avgscore")
    col_late_idx = _col_index("latesubmissions")
    col_rate_idx = _col_index("laterate")

    # Check rows
    data_rows = list(tbl.rows)[1:]
    expected_count = len(expected)
    # DB query already filters HAVING COUNT > 0, so row count must equal expected_count
    check("Table has expected number of data rows",
          len(data_rows) == expected_count,
          f"Found {len(data_rows)} rows, expected {expected_count}")

    # Build row map by name
    expected_by_name = {a[0].strip().lower(): a for a in expected}
    found_count = 0
    bad_values = []
    for row in data_rows:
        if not row.cells:
            continue
        name = row.cells[col_name_idx].text.strip()
        ekey = name.lower()
        if ekey not in expected_by_name:
            continue  # could be the optional zero-submission row
        e_name, e_total, e_avg, e_late, e_rate = expected_by_name[ekey]
        # Read agent values
        try:
            a_total = int(re.sub(r'[^\d-]', '', row.cells[col_total_idx].text)) if col_total_idx is not None else None
        except Exception:
            a_total = None
        try:
            a_avg = float(re.sub(r'[^\d.\-]', '', row.cells[col_avg_idx].text)) if col_avg_idx is not None else None
        except Exception:
            a_avg = None
        try:
            a_late = int(re.sub(r'[^\d-]', '', row.cells[col_late_idx].text)) if col_late_idx is not None else None
        except Exception:
            a_late = None
        try:
            a_rate = float(re.sub(r'[^\d.\-]', '', row.cells[col_rate_idx].text)) if col_rate_idx is not None else None
        except Exception:
            a_rate = None

        ok = True
        if a_total is None or a_total != e_total:
            ok = False
        if e_avg is None:
            pass
        elif a_avg is None or not num_close(a_avg, e_avg, 0.05):
            ok = False
        if a_late is None or a_late != e_late:
            ok = False
        if e_rate is None:
            pass
        elif a_rate is None or not num_close(a_rate, e_rate, 0.05):
            ok = False
        if ok:
            found_count += 1
        else:
            bad_values.append((name, (a_total, a_avg, a_late, a_rate),
                               (e_total, e_avg, e_late, e_rate)))

    check("All expected assignments present with correct numeric values",
          found_count == expected_count,
          f"matched={found_count}/{expected_count}, bad={bad_values[:3]}")

    # Sort check (alphabetical by name)
    names_in_order = [r.cells[col_name_idx].text.strip() for r in data_rows]
    sorted_names = sorted(names_in_order, key=lambda s: s.lower())
    check("Rows are sorted alphabetically by Assignment_Name",
          names_in_order == sorted_names,
          f"Order: {names_in_order[:5]}")


def check_gform():
    """Check Google Form creation."""
    print("\n=== Checking Google Form ===")
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        cur.execute("SELECT id, title FROM gform.forms")
        forms = cur.fetchall()
    except Exception as e:
        check("GForm DB query OK", False, str(e))
        return

    check("At least one form created", len(forms) >= 1, f"Found {len(forms)} forms")

    expected_title = "foundations of finance - assignment feedback survey"
    matching_form = None
    for fid, title in forms:
        if (title or "").strip().lower() == expected_title:
            matching_form = (fid, title)
            break

    check("Form title matches 'Foundations of Finance - Assignment Feedback Survey' exactly",
          matching_form is not None,
          f"Forms: {[(f[1] if len(f) > 1 else None) for f in forms]}")

    if matching_form:
        form_id = matching_form[0]
        try:
            cur.execute("SELECT COUNT(*) FROM gform.questions WHERE form_id = %s", (form_id,))
            q_count = cur.fetchone()[0]
        except Exception:
            q_count = 0
        check("Form has exactly 5 questions", q_count == 5, f"Found {q_count} questions")

    try:
        cur.close()
        conn.close()
    except Exception:
        pass


def check_emails(expected):
    """Check that summary email was sent."""
    print("\n=== Checking Emails ===")
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
        all_emails = cur.fetchall()
        conn.close()
    except Exception as e:
        check("Emails DB query OK", False, str(e))
        return

    def find_email_for_recipient(recipient):
        for subj, from_addr, to_addr, body in all_emails:
            if to_addr:
                recipients = []
                if isinstance(to_addr, list):
                    recipients = [str(r).strip().lower() for r in to_addr]
                elif isinstance(to_addr, str):
                    try:
                        parsed = json.loads(to_addr)
                        if isinstance(parsed, list):
                            recipients = [str(r).strip().lower() for r in parsed]
                        else:
                            recipients = [str(to_addr).strip().lower()]
                    except (json.JSONDecodeError, TypeError):
                        recipients = [str(to_addr).strip().lower()]
                if recipient.lower() in recipients:
                    return subj, from_addr, to_addr, body
        return None

    result = find_email_for_recipient("instructor@financeou.example.com")
    check("Summary email sent to instructor@financeou.example.com", result is not None,
          f"Total emails found: {len(all_emails)}")

    # Compute expected highlights
    if expected:
        # highest avg score
        with_avg = [a for a in expected if a[2] is not None]
        top_avg = max(with_avg, key=lambda a: a[2]) if with_avg else None
        # most late
        top_late = max(expected, key=lambda a: a[3]) if expected else None
    else:
        top_avg = None
        top_late = None

    if not result:
        for label in ["Email subject exact match", "Email from analytics",
                      "Email mentions count of assignments", "Email mentions highest-avg assignment",
                      "Email mentions most-late assignment", "Email contains form link"]:
            check(label, False, "no email")
        return

    subj, from_addr, to_addr, body = result
    expected_subject = "Assignment Analysis Report - Foundations of Finance (Fall 2013)"
    check("Email subject exact match",
          (subj or "").strip().lower() == expected_subject.lower(),
          f"Subject: {subj!r}, expected: {expected_subject!r}")

    check("Email from analytics@university.example.com",
          (from_addr or "").strip().lower() == "analytics@university.example.com",
          f"From: {from_addr}")

    body_lower = (body or "").lower()
    n = len(expected)
    body_clean = re.sub(r'\s+', ' ', body_lower)
    has_count = re.search(r'\b' + str(n) + r'\b', body_clean) is not None
    check(f"Email mentions count of {n} assignments", has_count, f"Body[:200]: {body_lower[:200]}")

    if top_avg:
        # require both name (full) and the avg score
        has_top_avg = top_avg[0].lower() in body_lower
        check(f"Email mentions highest-avg assignment ({top_avg[0]})",
              has_top_avg, f"Body[:200]: {body_lower[:200]}")
    if top_late:
        has_top_late = top_late[0].lower() in body_lower
        check(f"Email mentions most-late assignment ({top_late[0]})",
              has_top_late, f"Body[:200]: {body_lower[:200]}")

    # Form link check: should have the form URL or a UUID-style id
    has_form_link = (
        "form" in body_lower and
        (re.search(r'https?://[^\s]+', body or "") is not None or
         re.search(r'\b[a-f0-9-]{16,}\b', body or "") is not None)
    )
    check("Email contains form link", has_form_link,
          f"Body[:300]: {body[:300] if body else ''}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=" * 70)
    print("CANVAS ASSIGNMENT FEEDBACK GFORM WORD - EVALUATION")
    print("=" * 70)

    expected = _fetch_expected_from_db() or EXPECTED_ASSIGNMENTS_FALLBACK
    print(f"\nExpected {len(expected)} assignments from DB.")

    check_word(args.agent_workspace, expected)
    check_gform()
    check_emails(expected)

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    overall = FAIL_COUNT == 0
    print(f"  Overall: {'PASS' if overall else 'FAIL'}")

    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump({"passed": PASS_COUNT, "failed": FAIL_COUNT, "success": overall}, f)

    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
