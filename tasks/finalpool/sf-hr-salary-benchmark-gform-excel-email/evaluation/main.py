"""Evaluation for sf-hr-salary-benchmark-gform-excel-email.

Checks:
1. Salary_Analysis.xlsx has correct sheets and data
2. Google Forms survey "Compensation Satisfaction Survey" with 5 questions
3. Email to hr-leadership@company.example.com with subject matching pattern
"""
import argparse
import os
import sys

import openpyxl
import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="toolathlon_gym", user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0

# Note: GT comparison reads from groundtruth Excel directly; no per-dept constants needed.


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        print(f"  [FAIL] {name}: {str(detail)[:200]}")


def num_close(a, b, tol=1.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def str_match(a, b):
    if a is None or b is None:
        return a is None and b is None
    return str(a).strip().lower() == str(b).strip().lower()


def get_sheet(wb, name):
    for s in wb.sheetnames:
        if s.strip().lower() == name.strip().lower():
            return wb[s]
    return None


def check_excel(agent_workspace, groundtruth_workspace):
    print("\n=== Checking Salary_Analysis.xlsx ===")
    agent_file = os.path.join(agent_workspace, "Salary_Analysis.xlsx")
    gt_file = os.path.join(groundtruth_workspace, "Salary_Analysis.xlsx")

    check("Excel file exists", os.path.isfile(agent_file), agent_file)
    if not os.path.isfile(agent_file):
        return False

    try:
        agent_wb = openpyxl.load_workbook(agent_file, data_only=True)
        gt_wb = openpyxl.load_workbook(gt_file, data_only=True)
    except Exception as e:
        check("Excel readable", False, str(e))
        return False

    all_ok = True

    # Check Department_Stats sheet
    agent_dept = get_sheet(agent_wb, "Department_Stats")
    gt_dept = get_sheet(gt_wb, "Department_Stats")
    check("Sheet 'Department_Stats' exists", agent_dept is not None, f"Sheets: {agent_wb.sheetnames}")
    if agent_dept is None:
        all_ok = False
    else:
        a_rows = list(agent_dept.iter_rows(min_row=2, values_only=True))
        g_rows = list(gt_dept.iter_rows(min_row=2, values_only=True))
        check("Department_Stats has 7 rows", len(a_rows) == 7, f"Got {len(a_rows)}")
        if len(a_rows) != 7:
            all_ok = False

        # Build lookup by department name
        a_lookup = {}
        for r in a_rows:
            if r and r[0]:
                a_lookup[str(r[0]).strip().lower()] = r

        for g_row in g_rows:
            if not g_row or not g_row[0]:
                continue
            dept = str(g_row[0]).strip().lower()
            a_row = a_lookup.get(dept)
            if a_row is None:
                check(f"Dept '{g_row[0]}' present", False, "Missing")
                all_ok = False
                continue
            # Headcount (col 1)
            ok = num_close(a_row[1], g_row[1], 5)
            check(f"'{g_row[0]}' Headcount", ok, f"Expected {g_row[1]}, got {a_row[1]}")
            if not ok:
                all_ok = False
            # Min_Salary (col 2)
            ok = num_close(a_row[2], g_row[2], 1)
            check(f"'{g_row[0]}' Min_Salary", ok, f"Expected {g_row[2]}, got {a_row[2]}")
            if not ok:
                all_ok = False
            # Max_Salary (col 3)
            ok = num_close(a_row[3], g_row[3], 1)
            check(f"'{g_row[0]}' Max_Salary", ok, f"Expected {g_row[3]}, got {a_row[3]}")
            if not ok:
                all_ok = False
            # Avg_Salary (col 4)
            ok = num_close(a_row[4], g_row[4], 100)
            check(f"'{g_row[0]}' Avg_Salary", ok, f"Expected {g_row[4]}, got {a_row[4]}")
            if not ok:
                all_ok = False
            # Median_Salary (col 5)
            ok = num_close(a_row[5], g_row[5], 50)
            check(f"'{g_row[0]}' Median_Salary", ok, f"Expected {g_row[5]}, got {a_row[5]}")
            if not ok:
                all_ok = False

        # Sort by Avg_Salary descending verification
        try:
            avgs = [float(r[4]) for r in a_rows if r and r[4] is not None]
            sorted_ok = all(avgs[i] >= avgs[i+1] for i in range(len(avgs)-1))
        except Exception:
            sorted_ok = False
        check("Department_Stats sorted by Avg_Salary descending", sorted_ok,
              f"Avg_Salaries: {[r[4] for r in a_rows]}")
        if not sorted_ok:
            all_ok = False

    # Check Summary sheet
    agent_summary = get_sheet(agent_wb, "Summary")
    gt_summary = get_sheet(gt_wb, "Summary")
    check("Sheet 'Summary' exists", agent_summary is not None, f"Sheets: {agent_wb.sheetnames}")
    if agent_summary is None:
        all_ok = False
    else:
        a_summary = {}
        for row in agent_summary.iter_rows(min_row=2, values_only=True):
            if row and row[0]:
                a_summary[str(row[0]).strip().lower()] = row[1]

        # Total_Employees
        te = a_summary.get("total_employees")
        check("Total_Employees = 50000", num_close(te, 50000, 10), f"Got {te}")
        if not num_close(te, 50000, 10):
            all_ok = False

        # Company_Avg_Salary
        cas = a_summary.get("company_avg_salary")
        check("Company_Avg_Salary close to 58396", num_close(cas, 58396.14, 200), f"Got {cas}")

        # Highest_Paid_Dept (exact)
        hpd = str(a_summary.get("highest_paid_dept", "")).strip().lower()
        check("Highest_Paid_Dept exactly 'Engineering'", hpd == "engineering", f"Got '{hpd}'")

        # Lowest_Paid_Dept (exact)
        lpd = str(a_summary.get("lowest_paid_dept", "")).strip().lower()
        check("Lowest_Paid_Dept exactly 'Operations'", lpd == "operations", f"Got '{lpd}'")

    return all_ok


def check_gform():
    print("\n=== Checking Google Forms ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("SELECT id, title FROM gform.forms")
    forms = cur.fetchall()
    target_id = None
    for fid, title in forms:
        if (title or "").strip().lower() == "compensation satisfaction survey":
            target_id = fid
            break
    check("Form titled exactly 'Compensation Satisfaction Survey'", target_id is not None,
          f"Found titles: {[f[1] for f in forms]}")

    if target_id is not None:
        cur.execute("SELECT title, question_type FROM gform.questions WHERE form_id=%s ORDER BY position",
                    (target_id,))
        questions = cur.fetchall()
        check("Form has exactly 5 questions", len(questions) == 5, f"Got {len(questions)}")

        # Q1: satisfaction MC; Q2: pay competitive Yes/No/Not Sure; Q3: benefits short answer;
        # Q4: leaving Yes/No/Maybe; Q5: comments paragraph
        q_titles = [(q[0] or "").lower() for q in questions]
        q_types  = [(q[1] or "").lower() for q in questions]

        has_q1 = any(("satisfied" in t or "satisfaction" in t) and "compensation" in t for t in q_titles)
        check("Q1 satisfaction question on compensation present", has_q1, f"titles={q_titles}")
        has_q2 = any(("competitive" in t and ("industry" in t or "pay" in t)) for t in q_titles)
        check("Q2 competitive vs industry question present", has_q2, f"titles={q_titles}")
        has_q3 = any("benefit" in t for t in q_titles)
        check("Q3 benefits question present", has_q3, f"titles={q_titles}")
        has_q4 = any(("leav" in t or "leaving" in t) and ("pay" in t or "better" in t) for t in q_titles)
        check("Q4 leaving for better pay question present", has_q4, f"titles={q_titles}")
        has_q5 = any(("comment" in t or "additional" in t) for t in q_titles)
        check("Q5 additional comments question present", has_q5, f"titles={q_titles}")

        # Verify question types: 1 MC, 1 yes/no/notsure, 1 short text, 1 yes/no/maybe, 1 paragraph
        n_choice = sum(1 for t in q_types if t in ("multiple_choice", "radio", "choice", "mc", "choicequestion"))
        n_short = sum(1 for t in q_types if t in ("text", "short_answer", "textquestion"))
        n_para = sum(1 for t in q_types if t in ("paragraph", "long_answer", "long"))
        check("At least 3 multiple-choice/radio questions",
              n_choice >= 3, f"types={q_types}")
        check("At least 1 short-answer text question (benefits)",
              n_short >= 1, f"types={q_types}")
        check("At least 1 paragraph question (additional comments)",
              n_para >= 1, f"types={q_types}")
    cur.close()
    conn.close()


def check_email():
    print("\n=== Checking Email ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
    all_emails = cur.fetchall()
    cur.close()
    conn.close()

    expected_subject = "Compensation Analysis Report - Action Required"
    target = None
    for em in all_emails:
        if (em[0] or "").strip().lower() == expected_subject.lower():
            target = em
            break
    check("Email with exact subject 'Compensation Analysis Report - Action Required'",
          target is not None, f"Subjects seen: {[e[0] for e in all_emails]}")
    if target:
        subj, from_addr, to_addr, body = target
        to_str = str(to_addr or "").lower()
        check("Email to hr-leadership@company.example.com",
              "hr-leadership@company.example.com" in to_str, f"to: {to_addr}")
        check("Email from compensation@hr.example.com",
              "compensation@hr.example.com" in (from_addr or "").lower(), f"from: {from_addr}")
        body_l = (body or "").lower()
        # Tighten: anchor against trailing digit so '583961' / '58396×' don't false-positive
        import re
        has_avg = bool(re.search(r"\$?\s*58[, ]?396(?!\d)", body_l))
        check("Email body mentions company avg salary number 58396 (anchored)",
              has_avg, "expected 58396 (or 58,396) not followed by a digit")
        check("Email body identifies highest dept (Engineering)",
              "engineering" in body_l, "")
        check("Email body identifies lowest dept (Operations)",
              "operations" in body_l, "")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    task_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    gt_dir = args.groundtruth_workspace or os.path.join(task_root, "groundtruth_workspace")

    excel_ok = check_excel(args.agent_workspace, gt_dir)
    check_gform()
    check_email()

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    overall = excel_ok and FAIL_COUNT == 0
    print(f"  Overall: {'PASS' if overall else 'FAIL'}")

    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
