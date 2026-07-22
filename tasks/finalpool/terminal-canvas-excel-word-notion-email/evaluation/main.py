"""Evaluation for terminal-canvas-excel-word-notion-email.
Checks:
1. Student_Risk_Analysis.xlsx with 4 sheets and correct data
2. Intervention_Plan.docx with required sections
3. Notion database "Student Risk Tracker" with 2 entries
4. Email sent to academic_advisors@university.edu
5. risk_scorer.py script exists
"""
import argparse
import json
import os
import sys

import openpyxl
import psycopg2
from docx import Document

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent", "password": "camel",
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
        print(f"  [FAIL] {name}: {str(detail)[:300]}")


def num_close(a, b, tol=2.0):
    if a is None or b is None:
        return False
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def _norm(s):
    return str(s or "").strip().lower().replace(" ", "_")


def _load_gt_excel(gt_dir):
    gt_path = os.path.join(gt_dir, "Student_Risk_Analysis.xlsx")
    if not os.path.exists(gt_path):
        return None
    return openpyxl.load_workbook(gt_path, data_only=True)


def check_excel(workspace, gt_dir):
    print("\n=== Check 1: Student_Risk_Analysis.xlsx ===")
    path = os.path.join(workspace, "Student_Risk_Analysis.xlsx")
    if not os.path.exists(path):
        check("Excel file exists", False, f"Not found at {path}")
        return
    check("Excel file exists", True)

    wb = openpyxl.load_workbook(path, data_only=True)
    gt_wb = _load_gt_excel(gt_dir)
    sheets = wb.sheetnames
    check("Has at least 4 sheets", len(sheets) >= 4, f"Found {len(sheets)}: {sheets}")

    sheets_lower = [s.lower().replace(" ", "_") for s in sheets]

    # ------- Course_Overview -------
    co_idx = next((i for i, s in enumerate(sheets_lower) if "course" in s and "overview" in s), None)
    if co_idx is None:
        check("Course_Overview sheet exists", False, f"Sheets: {sheets}")
    else:
        check("Course_Overview sheet exists", True)
        ws1 = wb[sheets[co_idx]]
        rows1 = list(ws1.iter_rows(values_only=True))
        data1 = [r for r in rows1[1:] if r and any(c is not None for c in r)]
        check("Course_Overview has exactly 2 course rows", len(data1) == 2, f"Found {len(data1)}")
        # Header validation
        if rows1:
            headers = [_norm(c) for c in rows1[0]]
            for h in ["course_id", "course_name", "enrollment_count", "avg_score", "pass_rate"]:
                check(f"Course_Overview has '{h}' header", h in headers, f"got {headers}")
            # Compare against GT row-by-row
            if gt_wb and "Course_Overview" in gt_wb.sheetnames:
                gt_rows = list(gt_wb["Course_Overview"].iter_rows(values_only=True))
                gt_data = [r for r in gt_rows[1:] if r and any(c is not None for c in r)]
                gt_headers = [_norm(c) for c in gt_rows[0]]
                # Build agent lookup keyed by course_id
                cid_idx = headers.index("course_id") if "course_id" in headers else 0
                a_lookup = {str(r[cid_idx]).strip(): r for r in data1 if cid_idx < len(r)}
                for g_row in gt_data:
                    gcid = str(g_row[gt_headers.index("course_id")]).strip()
                    a_row = a_lookup.get(gcid)
                    if not a_row:
                        check(f"Course_Overview has course {gcid}", False, f"Not in {list(a_lookup.keys())}")
                        continue
                    # course_name (substring 'foundations of finance' + year)
                    cn_idx = headers.index("course_name")
                    if cn_idx < len(a_row):
                        a_name = str(a_row[cn_idx] or "").lower()
                        g_name = str(g_row[gt_headers.index("course_name")] or "").lower()
                        check(f"course {gcid} name matches GT", a_name == g_name, f"a='{a_name}' g='{g_name}'")
                    # enrollment_count (tol 5)
                    ec_idx = headers.index("enrollment_count")
                    if ec_idx < len(a_row):
                        check(
                            f"course {gcid} enrollment_count",
                            num_close(a_row[ec_idx], g_row[gt_headers.index("enrollment_count")], 5),
                            f"a={a_row[ec_idx]} g={g_row[gt_headers.index('enrollment_count')]}",
                        )
                    # avg_score (tol 1.0)
                    as_idx = headers.index("avg_score")
                    if as_idx < len(a_row):
                        check(
                            f"course {gcid} avg_score",
                            num_close(a_row[as_idx], g_row[gt_headers.index("avg_score")], 1.0),
                            f"a={a_row[as_idx]} g={g_row[gt_headers.index('avg_score')]}",
                        )
                    # pass_rate (tol 1.0)
                    pr_idx = headers.index("pass_rate")
                    if pr_idx < len(a_row):
                        check(
                            f"course {gcid} pass_rate",
                            num_close(a_row[pr_idx], g_row[gt_headers.index("pass_rate")], 1.0),
                            f"a={a_row[pr_idx]} g={g_row[gt_headers.index('pass_rate')]}",
                        )

    # ------- Risk_Distribution -------
    rd_idx = next((i for i, s in enumerate(sheets_lower) if "risk" in s and "dist" in s), None)
    if rd_idx is None:
        check("Risk_Distribution sheet exists", False)
    else:
        check("Risk_Distribution sheet exists", True)
        ws2 = wb[sheets[rd_idx]]
        rows2 = list(ws2.iter_rows(values_only=True))
        data2 = [r for r in rows2[1:] if r and any(c is not None for c in r)]
        check("Risk_Distribution has 3 rows", len(data2) == 3, f"Found {len(data2)}")
        if rows2:
            headers2 = [_norm(c) for c in rows2[0]]
            for h in ["risk_level", "student_count", "pct"]:
                check(f"Risk_Distribution has '{h}' header", h in headers2, f"got {headers2}")
            if gt_wb and "Risk_Distribution" in gt_wb.sheetnames:
                gt_rows = list(gt_wb["Risk_Distribution"].iter_rows(values_only=True))
                gt_data = [r for r in gt_rows[1:] if r and any(c is not None for c in r)]
                gt_headers = [_norm(c) for c in gt_rows[0]]
                # Build agent lookup keyed by risk_level (lower)
                rl_idx = headers2.index("risk_level") if "risk_level" in headers2 else 0
                a_lookup = {str(r[rl_idx]).strip().lower(): r for r in data2 if rl_idx < len(r)}
                for g_row in gt_data:
                    grl = str(g_row[gt_headers.index("risk_level")]).strip().lower()
                    a_row = a_lookup.get(grl)
                    if not a_row:
                        check(f"Risk_Distribution has '{grl}' row", False)
                        continue
                    sc_idx = headers2.index("student_count")
                    pct_idx = headers2.index("pct")
                    if sc_idx < len(a_row):
                        check(
                            f"Risk_Distribution {grl} student_count",
                            num_close(a_row[sc_idx], g_row[gt_headers.index("student_count")], 5),
                            f"a={a_row[sc_idx]} g={g_row[gt_headers.index('student_count')]}",
                        )
                    if pct_idx < len(a_row):
                        check(
                            f"Risk_Distribution {grl} pct",
                            num_close(a_row[pct_idx], g_row[gt_headers.index("pct")], 1.0),
                            f"a={a_row[pct_idx]} g={g_row[gt_headers.index('pct')]}",
                        )

    # ------- At_Risk_Students -------
    ar_idx = next((i for i, s in enumerate(sheets_lower) if "at_risk" in s or "risk_student" in s), None)
    if ar_idx is None:
        check("At_Risk_Students sheet exists", False)
    else:
        check("At_Risk_Students sheet exists", True)
        ws3 = wb[sheets[ar_idx]]
        rows3 = list(ws3.iter_rows(values_only=True))
        data3 = [r for r in rows3[1:] if r and any(c is not None for c in r)]
        check("At_Risk_Students has 2 course rows", len(data3) == 2, f"Found {len(data3)}")
        if rows3:
            headers3 = [_norm(c) for c in rows3[0]]
            for h in ["course_name", "high_risk_count", "medium_risk_count", "low_risk_count"]:
                check(f"At_Risk_Students has '{h}' header", h in headers3, f"got {headers3}")

    # ------- Intervention_Plan -------
    ip_idx = next((i for i, s in enumerate(sheets_lower) if "intervention" in s), None)
    if ip_idx is None:
        check("Intervention_Plan sheet exists", False)
    else:
        check("Intervention_Plan sheet exists", True)
        ws4 = wb[sheets[ip_idx]]
        rows4 = list(ws4.iter_rows(values_only=True))
        data4 = [r for r in rows4[1:] if r and any(c is not None for c in r)]
        check("Intervention_Plan has 3 risk-level rows", len(data4) == 3, f"Found {len(data4)}")
        if rows4:
            headers4 = [_norm(c) for c in rows4[0]]
            for h in ["risk_level", "action", "timeline", "responsible"]:
                check(f"Intervention_Plan has '{h}' header", h in headers4, f"got {headers4}")
            # Validate per-row content
            rl_idx = headers4.index("risk_level") if "risk_level" in headers4 else 0
            a_lookup = {str(r[rl_idx]).strip().lower(): r for r in data4 if rl_idx < len(r)}
            EXPECTED = {
                "high": ("advis", "1 week", "academic advisor"),
                "medium": ("tutor", "2 week", "tutor"),
                "low": ("self-paced", "1 month", "success"),
            }
            for level, (kw_action, kw_time, kw_resp) in EXPECTED.items():
                a_row = a_lookup.get(level)
                if not a_row:
                    check(f"Intervention_Plan has '{level}' row", False)
                    continue
                row_text = " ".join(str(c or "").lower() for c in a_row)
                check(f"Intervention_Plan {level} action mentions '{kw_action}'", kw_action in row_text, f"row: {row_text}")
                check(f"Intervention_Plan {level} timeline mentions '{kw_time}'", kw_time in row_text, f"row: {row_text}")
                check(f"Intervention_Plan {level} responsible mentions '{kw_resp}'", kw_resp in row_text, f"row: {row_text}")


def check_word(workspace):
    print("\n=== Check 2: Intervention_Plan.docx ===")
    path = os.path.join(workspace, "Intervention_Plan.docx")
    if not os.path.exists(path):
        check("Word document exists", False, f"Not found at {path}")
        return
    check("Word document exists", True)

    doc = Document(path)
    full_text = " ".join(p.text for p in doc.paragraphs).lower()
    check("Document title 'Student Retention Intervention Plan' present",
          "student retention intervention plan" in full_text,
          f"Text begin: {full_text[:200]}")
    check("Document has executive summary section", "executive summary" in full_text or "summary" in full_text)
    check("Document mentions risk distribution", "risk distribution" in full_text)
    check("Document mentions interventions", "intervention" in full_text)
    # All 3 risk levels
    for lvl in ["high", "medium", "low"]:
        check(f"Document mentions '{lvl}' risk level", lvl in full_text)
    # Both course years
    check("Document mentions Fall 2013", "2013" in full_text)
    check("Document mentions Fall 2014", "2014" in full_text)
    check("Document has substantial content (>=300 chars)", len(full_text) > 300, f"Length: {len(full_text)}")


def check_notion():
    print("\n=== Check 3: Notion Student Risk Tracker ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        cur.execute("SELECT id, title FROM notion.databases")
        dbs = cur.fetchall()
        tracker_db = None
        for db_id, title in dbs:
            title_str = ""
            if isinstance(title, list):
                title_str = " ".join(item.get("text", {}).get("content", "") for item in title if isinstance(item, dict))
            elif isinstance(title, str):
                try:
                    parsed = json.loads(title)
                    if isinstance(parsed, list):
                        title_str = " ".join(item.get("text", {}).get("content", "") for item in parsed if isinstance(item, dict))
                    else:
                        title_str = str(title)
                except Exception:
                    title_str = str(title)
            else:
                title_str = str(title) if title else ""
            if "student" in title_str.lower() and "risk" in title_str.lower() and "tracker" in title_str.lower():
                tracker_db = (db_id, title_str)
                break
        check("Student Risk Tracker database exists", tracker_db is not None,
              f"Databases: {[d[1] for d in dbs]}")

        if tracker_db:
            cur.execute("""
                SELECT COUNT(*) FROM notion.pages
                WHERE parent->>'database_id' = %s AND archived = false
            """, (tracker_db[0],))
            count = cur.fetchone()[0]
            check("Tracker has exactly 2 course entries", count == 2, f"Found {count}")
            # Verify entries reference both course years
            cur.execute("""
                SELECT properties::text FROM notion.pages
                WHERE parent->>'database_id' = %s AND archived = false
            """, (tracker_db[0],))
            props_text = " ".join((r[0] or "").lower() for r in cur.fetchall())
            check("Notion entries mention 2013", "2013" in props_text, f"props: {props_text[:200]}")
            check("Notion entries mention 2014", "2014" in props_text, f"props: {props_text[:200]}")
    except Exception as e:
        check("Notion check", False, str(e))
    finally:
        cur.close()
        conn.close()


def check_email():
    print("\n=== Check 4: Email to Academic Advisors ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        cur.execute("""
            SELECT subject, to_addr, body_text
            FROM email.messages
            WHERE to_addr::text ILIKE '%academic_advisors@university.edu%'
        """)
        rows = cur.fetchall()
        check("Email sent to academic_advisors@university.edu", len(rows) >= 1, "No matching email")
        if rows:
            target_subj = "student retention risk analysis - action required"
            matched = None
            for subj, to_addr, body in rows:
                if (subj or "").strip().lower() == target_subj:
                    matched = (subj, to_addr, body)
                    break
            check(
                "Email subject 'Student Retention Risk Analysis - Action Required'",
                matched is not None,
                f"subjects: {[r[0] for r in rows]}",
            )
            if matched:
                _, _, body = matched
                body_l = (body or "").lower()
                # Require 'high' adjacent to a numeric count (within ~80 chars), not just any digit anywhere.
                import re as _re_email
                high_count_re = (
                    r"high[^\n]{0,80}\b\d+\b"
                    r"|\b\d+\b[^\n]{0,80}high"
                )
                check("Email body mentions high-risk count (with numeric adjacency)",
                      _re_email.search(high_count_re, body_l) is not None,
                      f"body: {body_l[:200]}")
                check("Email body mentions next steps", "step" in body_l or "intervention" in body_l, f"body: {body_l[:200]}")
    except Exception as e:
        check("Email check", False, str(e))
    finally:
        cur.close()
        conn.close()


def check_script(workspace):
    print("\n=== Check 5: risk_scorer.py ===")
    path = os.path.join(workspace, "risk_scorer.py")
    check("risk_scorer.py exists", os.path.exists(path))


def check_reverse_validation(workspace):
    print("\n=== Reverse Validation ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute("""
            SELECT to_addr FROM email.messages
            WHERE subject ILIKE '%risk%' OR subject ILIKE '%retention%'
        """)
        emails = cur.fetchall()
        noise_recipients = ["all-staff@university.edu", "it@university.edu",
                            "facilities@university.edu"]
        for email_row in emails:
            to_str = str(email_row[0]).lower()
            for noise in noise_recipients:
                if noise in to_str:
                    check("No risk emails sent to wrong recipients", False,
                          f"Sent to noise recipient: {noise}")
                    cur.close(); conn.close()
                    return
        check("No risk emails sent to wrong recipients", True)
        cur.close(); conn.close()
    except Exception as e:
        check("Reverse validation", False, str(e))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_excel(args.agent_workspace, args.groundtruth_workspace)
    check_word(args.agent_workspace)
    check_notion()
    check_email()
    check_script(args.agent_workspace)
    check_reverse_validation(args.agent_workspace)

    total = PASS_COUNT + FAIL_COUNT
    print(f"\nOverall: {PASS_COUNT}/{total} checks passed")

    result = {"total_passed": PASS_COUNT, "total_checks": total}
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    # Strict: all checks must pass (FAIL_COUNT == 0)
    if FAIL_COUNT == 0:
        print("PASS")
        sys.exit(0)
    else:
        print(f"FAIL ({FAIL_COUNT} failures)")
        sys.exit(1)


if __name__ == "__main__":
    main()
