"""
Evaluation script for canvas-quiz-performance-gcal-gform task.

Checks:
1. Excel file Quiz_Performance_Report.xlsx - 2 sheets with correct structure and data
2. Google Form "Quiz Improvement Feedback" exists with at least 3 questions
3. Google Calendar has 3 specific tutoring session events (date/time/title verified)
4. Email sent to ccc.instructor@university.edu with exact subject
"""

import argparse
import json
import os
import re
import sys

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def num_close(a, b, tol=1.0):
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return str(a).strip().lower() == str(b).strip().lower()


def str_match(a, b):
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return str(a).strip().lower() == str(b).strip().lower()


def load_sheet_by_name(wb, name):
    for sname in wb.sheetnames:
        if sname.strip().lower() == name.strip().lower():
            return [[cell.value for cell in row] for row in wb[sname].iter_rows()]
    return None


# ============================================================================
# Check 1: Excel file
# ============================================================================

def check_excel(agent_workspace, groundtruth_workspace):
    print("\n=== Checking Quiz_Performance_Report.xlsx ===")

    try:
        import openpyxl
    except ImportError:
        record("openpyxl available", False, "pip install openpyxl")
        return

    agent_file = os.path.join(agent_workspace, "Quiz_Performance_Report.xlsx")
    gt_file = os.path.join(groundtruth_workspace, "Quiz_Performance_Report.xlsx")

    if not os.path.isfile(agent_file):
        record("Excel file exists", False, f"Not found: {agent_file}")
        return
    record("Excel file exists", True)

    agent_wb = openpyxl.load_workbook(agent_file, data_only=True)
    gt_wb = openpyxl.load_workbook(gt_file, data_only=True)

    # Check Quiz Stats sheet
    a_quiz = load_sheet_by_name(agent_wb, "Quiz Stats")
    g_quiz = load_sheet_by_name(gt_wb, "Quiz Stats")
    record("Sheet 'Quiz Stats' exists", a_quiz is not None)

    if a_quiz is not None and g_quiz is not None:
        a_data = [r for r in a_quiz[1:] if any(v is not None for v in r)]
        g_data = [r for r in g_quiz[1:] if any(v is not None for v in r)]
        record("Quiz Stats row count matches", len(a_data) == len(g_data),
               f"Expected {len(g_data)}, got {len(a_data)}")

        # Build lookup by quiz title
        a_lookup = {}
        for row in a_data:
            if row and row[0] is not None:
                a_lookup[str(row[0]).strip().lower()] = row

        # Verify ascending sort by Avg_Score
        try:
            avg_scores = [float(r[2]) for r in a_data if r[2] is not None]
            ascending_ok = all(avg_scores[i] <= avg_scores[i+1] for i in range(len(avg_scores)-1))
        except Exception:
            ascending_ok = False
        record("Quiz Stats sorted ascending by Avg_Score", ascending_ok,
               f"Avg_Scores: {[r[2] for r in a_data]}")

        for g_row in g_data:
            if not g_row or g_row[0] is None:
                continue
            title = str(g_row[0]).strip()
            key = title.lower()
            a_row = a_lookup.get(key)
            if a_row is None:
                record(f"Quiz row exists: {title}", False, "Row not found")
                continue
            record(f"Quiz row exists: {title}", True)

            # Total_Submissions (col index 1)
            if len(g_row) > 1 and len(a_row) > 1:
                record(f"{title}: Total_Submissions",
                       num_close(a_row[1], g_row[1], 1),
                       f"got {a_row[1]}, expected {g_row[1]}")
            # Avg_Score (col index 2) - 2dp tol 0.5
            if len(g_row) > 2 and len(a_row) > 2:
                record(f"{title}: Avg_Score",
                       num_close(a_row[2], g_row[2], 0.5),
                       f"got {a_row[2]}, expected {g_row[2]}")
            # Min_Score (col index 3)
            if len(g_row) > 3 and len(a_row) > 3:
                record(f"{title}: Min_Score",
                       num_close(a_row[3], g_row[3], 1),
                       f"got {a_row[3]}, expected {g_row[3]}")
            # Max_Score (col index 4)
            if len(g_row) > 4 and len(a_row) > 4:
                record(f"{title}: Max_Score",
                       num_close(a_row[4], g_row[4], 1),
                       f"got {a_row[4]}, expected {g_row[4]}")
            # Pass_Rate_Pct (col index 5) - 2dp tol 0.5
            if len(g_row) > 5 and len(a_row) > 5:
                record(f"{title}: Pass_Rate_Pct",
                       num_close(a_row[5], g_row[5], 0.5),
                       f"got {a_row[5]}, expected {g_row[5]}")

    # Check Course Summary sheet
    a_summ = load_sheet_by_name(agent_wb, "Course Summary")
    g_summ = load_sheet_by_name(gt_wb, "Course Summary")
    record("Sheet 'Course Summary' exists", a_summ is not None)

    if a_summ is not None and g_summ is not None:
        a_data = [r for r in a_summ[1:] if any(v is not None for v in r)]
        g_data = [r for r in g_summ[1:] if any(v is not None for v in r)]

        a_lookup = {}
        for row in a_data:
            if row and row[0] is not None:
                a_lookup[str(row[0]).strip().lower()] = row

        for g_row in g_data:
            if not g_row or g_row[0] is None:
                continue
            key = str(g_row[0]).strip().lower()
            a_row = a_lookup.get(key)
            if a_row is None:
                record(f"Summary row: {g_row[0]}", False, "Row not found")
                continue
            record(f"Summary row: {g_row[0]}", True)

            if key == "total_quizzes":
                record("Total_Quizzes value",
                       num_close(a_row[1], g_row[1], 0),
                       f"got {a_row[1]}, expected {g_row[1]}")
            elif key == "total_quiz_submissions":
                record("Total_Quiz_Submissions value",
                       num_close(a_row[1], g_row[1], 1),
                       f"got {a_row[1]}, expected {g_row[1]}")
            elif key == "overall_avg_score":
                record("Overall_Avg_Score value",
                       num_close(a_row[1], g_row[1], 0.5),
                       f"got {a_row[1]}, expected {g_row[1]}")
            elif key == "lowest_avg_quiz":
                record("Lowest_Avg_Quiz value",
                       str_match(a_row[1], g_row[1]),
                       f"got {a_row[1]}, expected {g_row[1]}")
            elif key == "highest_avg_quiz":
                record("Highest_Avg_Quiz value",
                       str_match(a_row[1], g_row[1]),
                       f"got {a_row[1]}, expected {g_row[1]}")


# ============================================================================
# Check 2: Google Form
# ============================================================================

def check_gform():
    print("\n=== Checking Google Form ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("SELECT id, title FROM gform.forms")
    forms = cur.fetchall()
    print(f"[check_gform] Found {len(forms)} forms.")
    record("At least 1 form created", len(forms) >= 1, f"Found {len(forms)}")

    expected_title = "quiz improvement feedback"
    target_form_id = None
    for form_id, title in forms:
        if title and title.strip().lower() == expected_title:
            target_form_id = form_id
            break
    record("Form titled exactly 'Quiz Improvement Feedback' found",
           target_form_id is not None,
           f"Forms titles: {[f[1] for f in forms]}")

    if target_form_id is None:
        cur.close()
        conn.close()
        return

    cur.execute("SELECT title, question_type, required FROM gform.questions WHERE form_id=%s ORDER BY position",
                (target_form_id,))
    questions = cur.fetchall()
    record("Form has at least 3 questions", len(questions) >= 3,
           f"Found {len(questions)} questions")

    # Check for MC challenging-topic question with all four options
    has_mc_topic = False
    has_scale_confidence = False
    has_text_resources = False
    for q_title, q_type, q_required in questions:
        ql = (q_title or "").lower()
        qt = (q_type or "").lower()
        if (q_type and qt in ("multiple_choice", "radio", "multi_choice", "mc")) or qt == "multiple_choice":
            if "challeng" in ql or "topic" in ql:
                has_mc_topic = True
        if qt in ("scale", "rating", "linear_scale"):
            if "confiden" in ql or "prepar" in ql:
                has_scale_confidence = True
        if qt in ("text", "paragraph", "short_answer", "long_answer"):
            if "resource" in ql or "additional" in ql or "help" in ql:
                has_text_resources = True

    record("Multiple-choice 'most challenging topic' question present", has_mc_topic,
           f"Questions: {[(q[0], q[1]) for q in questions]}")
    record("Scale/linear question on confidence present", has_scale_confidence,
           f"Questions: {[(q[0], q[1]) for q in questions]}")
    record("Text question on additional resources present", has_text_resources,
           f"Questions: {[(q[0], q[1]) for q in questions]}")

    # Check the MC options match the four prescribed
    cur.execute("""SELECT q.title, o.value
                   FROM gform.questions q LEFT JOIN gform.question_options o ON o.question_id=q.id
                   WHERE q.form_id=%s""", (target_form_id,))
    rows = cur.fetchall()
    by_q = {}
    for qt, opt in rows:
        if opt is None:
            continue
        by_q.setdefault(qt, []).append((opt or "").strip().lower())

    expected_opts = {"programming fundamentals", "data structures", "algorithms", "user interface design"}
    found_all_opts = False
    for qt, opts in by_q.items():
        opts_set = set(opts)
        if expected_opts.issubset(opts_set):
            found_all_opts = True
            break
    record("MC options include all four prescribed topics", found_all_opts,
           f"Options by Q: {by_q}")

    cur.close()
    conn.close()


# ============================================================================
# Check 3: Google Calendar
# ============================================================================

def check_gcal():
    print("\n=== Checking Google Calendar ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("""
        SELECT summary, start_datetime, end_datetime
        FROM gcal.events
        ORDER BY start_datetime
    """)
    events = cur.fetchall()
    cur.close()
    conn.close()

    print(f"[check_gcal] Found {len(events)} calendar events.")
    record("At least 3 calendar events created", len(events) >= 3, f"Found {len(events)}")

    expected_sessions = [
        ("CCC Spring 2014 Tutoring Session 1", "2026-03-12"),
        ("CCC Spring 2014 Tutoring Session 2", "2026-03-19"),
        ("CCC Spring 2014 Tutoring Session 3", "2026-03-26"),
    ]
    for exp_title, exp_date in expected_sessions:
        match = None
        for summary, start_dt, end_dt in events:
            if summary and summary.strip().lower() == exp_title.lower():
                if start_dt and exp_date in str(start_dt):
                    match = (summary, start_dt, end_dt)
                    break
        record(f"Event '{exp_title}' on {exp_date} found", match is not None,
               f"Events: {[(e[0], str(e[1])[:10]) for e in events]}")
        if match:
            # Check time exactly 15:00 start and 17:00 end (parse hour/minute)
            from datetime import datetime as _dt
            sdt = match[1]; edt = match[2]
            try:
                if not hasattr(sdt, "hour"):
                    sdt = _dt.fromisoformat(str(sdt).replace("Z", "+00:00"))
                if not hasattr(edt, "hour"):
                    edt = _dt.fromisoformat(str(edt).replace("Z", "+00:00"))
                time_ok = (sdt.hour == 15 and sdt.minute == 0
                           and edt.hour == 17 and edt.minute == 0)
            except Exception:
                time_ok = False
            record(f"Event '{exp_title}' time 15:00-17:00",
                   time_ok, f"start={match[1]}, end={match[2]}")


# ============================================================================
# Check 4: Email
# ============================================================================

def check_emails():
    print("\n=== Checking Emails ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("""
        SELECT subject, from_addr, to_addr, body_text
        FROM email.messages
    """)
    all_emails = cur.fetchall()
    cur.close()
    conn.close()

    print(f"[check_emails] Found {len(all_emails)} total emails.")
    record("At least 1 email sent", len(all_emails) >= 1, f"Found {len(all_emails)}")

    # Strict: must be sent to ccc.instructor@university.edu
    target_email = None
    for subject, from_addr, to_addr, body_text in all_emails:
        to_str = json.dumps(to_addr).lower() if isinstance(to_addr, (list, dict)) else str(to_addr or "").lower()
        if "ccc.instructor@university.edu" in to_str:
            target_email = (subject, from_addr, to_addr, body_text)
            break
    record("Email to ccc.instructor@university.edu found", target_email is not None,
           f"Emails: {[(e[0], e[2]) for e in all_emails[:5]]}")

    if target_email is None:
        return

    subject, _, _, body_text = target_email
    expected_subject = "Creative Computing Quiz Performance Report"
    record("Email subject exactly 'Creative Computing Quiz Performance Report'",
           (subject or "").strip().lower() == expected_subject.lower(),
           f"Subject: {subject}")

    body_lower = (body_text or "").lower()
    record("Body summarizes quiz performance findings",
           ("quiz" in body_lower) and ("performance" in body_lower or "score" in body_lower or "average" in body_lower or "avg" in body_lower),
           "Missing performance/score/average mention")
    record("Body mentions tutoring sessions",
           "tutoring" in body_lower or "session" in body_lower,
           "Missing tutoring/session mention")


# ============================================================================
# Main
# ============================================================================

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    task_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    gt_dir = args.groundtruth_workspace or os.path.join(task_root, "groundtruth_workspace")

    check_excel(args.agent_workspace, gt_dir)
    check_gform()
    check_gcal()
    check_emails()

    all_passed = (FAIL_COUNT == 0)

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    print(f"  Overall: {'PASS' if all_passed else 'FAIL'}")

    if args.res_log_file:
        result = {
            "passed": PASS_COUNT,
            "failed": FAIL_COUNT,
            "success": all_passed,
        }
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    sys.exit(0 if all_passed else 1)


if __name__ == "__main__":
    main()
