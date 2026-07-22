"""Evaluation script for terminal-canvas-pdf-gsheet-email-word."""
import os
import argparse, json, os, sys

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent", "password": "camel"
}

PASS_COUNT = 0
FAIL_COUNT = 0


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        detail_str = str(detail)[:200] if detail else ""
        print(f"  [FAIL] {name}: {detail_str}")


def get_conn():
    import psycopg2
    return psycopg2.connect(**DB_CONFIG)


def check_gsheet():
    print("\n=== Checking Google Sheets ===")
    try:
        conn = get_conn()
        cur = conn.cursor()

        cur.execute("SELECT id, title FROM gsheet.spreadsheets WHERE LOWER(title) LIKE %s",
                    ("%faculty%workload%tracker%",))
        rows = cur.fetchall()
        if not rows:
            cur.execute("SELECT id, title FROM gsheet.spreadsheets")
            all_ss = cur.fetchall()
            check("Faculty_Workload_Tracker spreadsheet exists", False,
                  f"Found: {[r[1] for r in all_ss]}")
            cur.close()
            conn.close()
            return False

        check("Faculty_Workload_Tracker spreadsheet exists", True)
        ss_id = rows[0][0]

        cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s", (ss_id,))
        sheets = cur.fetchall()
        sheet_titles = [s[1].lower() for s in sheets]

        has_workload = any("workload" in t or "course" in t for t in sheet_titles)
        has_summary = any("summary" in t or "subject" in t for t in sheet_titles)
        check("Course_Workload sheet exists", has_workload, f"Sheets: {sheet_titles}")
        check("Subject_Summary sheet exists", has_summary, f"Sheets: {sheet_titles}")

        # Check Course_Workload content
        workload_sheet_id = None
        for sid, title in sheets:
            if "workload" in title.lower() or "course" in title.lower():
                workload_sheet_id = sid
                break

        if workload_sheet_id:
            cur.execute("SELECT row_index, col_index, value FROM gsheet.cells WHERE sheet_id = %s ORDER BY row_index, col_index",
                        (workload_sheet_id,))
            cells = cur.fetchall()
            data_rows = set(r for r, c, v in cells if r > 0 and v and str(v).strip())
            # Query dynamic course count from Canvas DB
            try:
                conn2 = get_conn()
                cur2 = conn2.cursor()
                cur2.execute("SELECT COUNT(*) FROM canvas.courses")
                expected_workload_rows = cur2.fetchone()[0]
                cur2.close(); conn2.close()
            except Exception:
                expected_workload_rows = 20
            check(f"Course_Workload has >= {expected_workload_rows} rows",
                  len(data_rows) >= expected_workload_rows,
                  f"Found {len(data_rows)} data rows")

            # Check headers
            headers = [str(v).lower() for r, c, v in cells if r == 0 and v]
            has_name = any("name" in h or "course" in h for h in headers)
            has_assign = any("assign" in h for h in headers)
            check("Course_Workload has course name column", has_name, f"Headers: {headers}")
            check("Course_Workload has assignment column", has_assign, f"Headers: {headers}")

        # Check Subject_Summary content
        summary_sheet_id = None
        for sid, title in sheets:
            if "summary" in title.lower() or "subject" in title.lower():
                summary_sheet_id = sid
                break

        if summary_sheet_id:
            cur.execute("SELECT row_index, col_index, value FROM gsheet.cells WHERE sheet_id = %s ORDER BY row_index, col_index",
                        (summary_sheet_id,))
            cells = cur.fetchall()
            data_rows = set(r for r, c, v in cells if r > 0 and v and str(v).strip())
            # Query dynamic unique subject count from Canvas DB
            try:
                conn3 = get_conn()
                cur3 = conn3.cursor()
                cur3.execute("SELECT COUNT(DISTINCT regexp_replace(name, ' \\(.*\\)$', '')) FROM canvas.courses")
                expected_subject_rows = cur3.fetchone()[0]
                cur3.close(); conn3.close()
            except Exception:
                expected_subject_rows = 7
            check(f"Subject_Summary has >= {expected_subject_rows} rows",
                  len(data_rows) >= expected_subject_rows,
                  f"Found {len(data_rows)} data rows")

            # Check that Heavy AND Light/Medium ratings are present (not just Heavy)
            all_values = [str(v).lower() for _, _, v in cells if v]
            has_heavy = any("heavy" in v for v in all_values)
            has_light_or_med = any(("light" in v) or ("medium" in v) or ("normal" in v) or ("moderate" in v) for v in all_values)
            check("Subject_Summary has Heavy rating", has_heavy, f"Values sample: {all_values[:15]}")
            check("Subject_Summary has at least one non-Heavy rating", has_light_or_med, f"Values sample: {all_values[:15]}")

        cur.close()
        conn.close()
        return True
    except Exception as e:
        check("Google Sheets accessible", False, str(e))
        return False


def check_word(agent_workspace):
    print("\n=== Checking Word Document ===")
    word_path = os.path.join(agent_workspace, "Faculty_Workload_Report.docx")
    check("Faculty_Workload_Report.docx exists", os.path.exists(word_path))
    if os.path.exists(word_path):
        from docx import Document
        doc = Document(word_path)
        text = " ".join(p.text for p in doc.paragraphs).lower()
        check("Word has substantial content", len(text) > 300, f"length: {len(text)}")
        check("Word mentions workload", "workload" in text)
        # Dynamically query subjects from Canvas DB instead of hardcoding
        try:
            conn = get_conn()
            cur = conn.cursor()
            cur.execute(r"SELECT regexp_replace(name, ' \(.*\)$', '') AS subj FROM canvas.courses GROUP BY subj")
            subjects = [r[0].lower() for r in cur.fetchall()]
            cur.close(); conn.close()
        except Exception:
            subjects = ["finance", "biochem"]
        # Require mention of >=3 distinct subjects (Subject Analysis section requirement)
        subject_hits = sum(1 for s in subjects if any(kw in text for kw in s.split() if len(kw) > 4))
        check(f"Word mentions >=3 distinct subjects", subject_hits >= 3,
              f"hits={subject_hits} of {len(subjects)} subjects")
        check("Word mentions recommendations", "recommend" in text)
        # Required sections from task.md
        section_keywords = {
            "overview": ["overview", "total courses"],
            "subject analysis": ["subject analysis"],
            "high workload areas": ["high workload"],
            "recommendations": ["recommendations"],
        }
        for sec_label, kws in section_keywords.items():
            has_section = any(kw in text for kw in kws)
            check(f"Word has '{sec_label}' section", has_section,
                  f"section keywords {kws} not found")
        # At least 3 actionable recommendations - look for indicator words in 'recommend' area
        rec_idx = text.find("recommend")
        rec_segment = text[rec_idx:rec_idx + 2000] if rec_idx >= 0 else ""
        rec_indicators = ["redistribut", "teaching assistant", "ta ", "ta\n", "balanc", "reduc",
                          "increase", "hire", "expand", "review", "rebalanc", "address",
                          "improve", "consider", "implement"]
        rec_hits = sum(1 for k in rec_indicators if k in rec_segment)
        check("Word recommendations section has >=3 actionable items", rec_hits >= 3,
              f"indicator hits={rec_hits}")


def check_emails():
    print("\n=== Checking Emails ===")
    try:
        conn = get_conn()
        cur = conn.cursor()
        # Email 1: subject "Faculty Workload Analysis Complete" to department-chairs@university.edu
        cur.execute("""SELECT subject, to_addr, body_text FROM email.messages
                       WHERE subject ILIKE '%faculty workload analysis complete%'""")
        emails = cur.fetchall()
        check("Email 'Faculty Workload Analysis Complete' sent", len(emails) >= 1, f"found {len(emails)}")
        if emails:
            to_str = str(emails[0][1] or "").lower()
            check("Email 1 to department-chairs@university.edu",
                  "department-chairs@university.edu" in to_str,
                  f"to: {emails[0][1]}")
            body = (emails[0][2] or "").lower()
            # Total course count, Heavy subjects, top recommendation
            check("Email 1 mentions courses analyzed",
                  "courses" in body and ("22" in body or "all" in body),
                  f"body sample: {body[:200]}")
            check("Email 1 mentions heavy workload subjects",
                  "heavy" in body, f"body sample: {body[:200]}")
            check("Email 1 mentions a recommendation",
                  "recommend" in body, f"body sample: {body[:200]}")

        # Email 2: subject "Workload Standards Compliance Report" to academic-affairs@university.edu
        cur.execute("""SELECT subject, to_addr, body_text FROM email.messages
                       WHERE subject ILIKE '%workload standards compliance report%'""")
        compliance = cur.fetchall()
        check("Email 'Workload Standards Compliance Report' sent",
              len(compliance) >= 1, f"found {len(compliance)}")
        if compliance:
            to_str2 = str(compliance[0][1] or "").lower()
            check("Email 2 to academic-affairs@university.edu",
                  "academic-affairs@university.edu" in to_str2,
                  f"to: {compliance[0][1]}")
            body2 = (compliance[0][2] or "").lower()
            check("Email 2 mentions compliance vs exceed",
                  ("comply" in body2 or "compliant" in body2 or "compliance" in body2)
                  and ("exceed" in body2 or "noncompliant" in body2 or "non-compliant" in body2),
                  f"body sample: {body2[:200]}")

        cur.close()
        conn.close()
    except Exception as e:
        check("Email checks", False, str(e))


def check_script(agent_workspace):
    print("\n=== Checking Terminal Script ===")
    check("workload_analyzer.py exists",
          os.path.exists(os.path.join(agent_workspace, "workload_analyzer.py")))
    # Verify intermediate JSON outputs explicitly referenced in task.md
    for fname in ("course_data.json", "standards.json", "workload_analysis.json"):
        fp = os.path.join(agent_workspace, fname)
        check(f"{fname} exists", os.path.exists(fp))
        if os.path.exists(fp):
            try:
                with open(fp, "r", encoding="utf-8") as f:
                    data = json.load(f)
                check(f"{fname} is valid non-empty JSON",
                      isinstance(data, (list, dict)) and len(data) > 0,
                      f"len={len(data) if hasattr(data, '__len__') else 0}")
            except Exception as e:
                check(f"{fname} is valid non-empty JSON", False, str(e)[:120])


def check_reverse_validation(agent_workspace):
    print("\n=== Reverse Validation ===")
    # Check no emails sent to wrong recipients
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("""
            SELECT to_addr FROM email.messages
            WHERE subject ILIKE '%%faculty%%workload%%'
               OR subject ILIKE '%%workload%%standards%%'
        """)
        emails = cur.fetchall()
        noise_recipients = ["all-staff@university.edu", "students@university.edu"]
        for email_row in emails:
            to_str = str(email_row[0]).lower()
            for noise in noise_recipients:
                if noise in to_str:
                    check("No workload emails sent to wrong recipients", False,
                          f"Sent to noise recipient: {noise}")
                    cur.close(); conn.close()
                    return
        check("No workload emails sent to wrong recipients", True)
        cur.close(); conn.close()
    except Exception as e:
        check("Reverse validation", False, str(e))


def run_evaluation(agent_workspace, groundtruth_workspace, launch_time, res_log_file):
    global PASS_COUNT, FAIL_COUNT
    PASS_COUNT = 0
    FAIL_COUNT = 0

    check_gsheet()
    check_word(agent_workspace)
    check_emails()
    check_script(agent_workspace)
    check_reverse_validation(agent_workspace)

    return FAIL_COUNT == 0, f"Passed {PASS_COUNT}/{PASS_COUNT + FAIL_COUNT} checks"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    success, message = run_evaluation(
        args.agent_workspace, args.groundtruth_workspace,
        args.launch_time, args.res_log_file
    )
    print(message)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
