"""Evaluation for playwright-sf-hr-benefits-survey-gform-excel."""
import argparse
import os
import sys

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432,
          dbname=os.environ.get("PGDATABASE", "toolathlon_gym"),
          user="eigent", password="camel")


def num_close(a, b, tol=0.5):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return str(a).strip().lower() == str(b).strip().lower()


def load_sheet_rows(wb, sheet_name):
    for name in wb.sheetnames:
        if name.strip().lower() == sheet_name.strip().lower():
            return [[cell.value for cell in row] for row in wb[name].iter_rows()]
    return None


def check_excel(agent_workspace, gt_workspace):
    errors = []
    import openpyxl
    path = os.path.join(agent_workspace, "Benefits_Analysis.xlsx")
    if not os.path.exists(path):
        return ["Benefits_Analysis.xlsx not found"]
    gt_path = os.path.join(gt_workspace, "Benefits_Analysis.xlsx")
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        gt_wb = openpyxl.load_workbook(gt_path, data_only=True) if os.path.exists(gt_path) else None

        # ---------- Competitor Comparison (full validation, all 7 rows) ----------
        rows = load_sheet_rows(wb, "Competitor Comparison")
        if rows is None:
            errors.append("Sheet 'Competitor Comparison' not found")
        else:
            data_rows = [r for r in rows[1:] if r and r[0] is not None]
            if len(data_rows) != 7:
                errors.append(f"Competitor Comparison has {len(data_rows)} rows, expected 7")
            agent_lookup = {str(r[0]).strip().lower(): r for r in data_rows if r[0]}
            if gt_wb is not None:
                gt_rows = load_sheet_rows(gt_wb, "Competitor Comparison") or []
                gt_data = [r for r in gt_rows[1:] if r and r[0] is not None]
                for g_row in gt_data:
                    key = str(g_row[0]).strip().lower()
                    a_row = agent_lookup.get(key)
                    if a_row is None:
                        errors.append(f"Competitor row missing for {g_row[0]}")
                        continue
                    a_padded = list(a_row) + [None] * (4 - len(a_row))
                    # Health, PTO, Retirement
                    if not num_close(a_padded[1], g_row[1], 1):
                        errors.append(f"{g_row[0]} Health_Insurance_Pct: {a_padded[1]} vs {g_row[1]}")
                    if not num_close(a_padded[2], g_row[2], 1):
                        errors.append(f"{g_row[0]} PTO_Days: {a_padded[2]} vs {g_row[2]}")
                    if not num_close(a_padded[3], g_row[3], 0.2):
                        errors.append(f"{g_row[0]} Retirement_Match_Pct: {a_padded[3]} vs {g_row[3]}")
            # Sort order: rows must be alphabetical by Company
            company_names = [str(r[0]).strip() for r in data_rows if r and r[0]]
            sorted_names = sorted(company_names, key=lambda s: s.lower())
            if company_names != sorted_names:
                errors.append("Competitor Comparison not sorted alphabetically by Company")

        # ---------- Department Satisfaction (all 7 rows) ----------
        rows2 = load_sheet_rows(wb, "Department Satisfaction")
        if rows2 is None:
            errors.append("Sheet 'Department Satisfaction' not found")
        else:
            data_rows2 = [r for r in rows2[1:] if r and r[0] is not None]
            if len(data_rows2) != 7:
                errors.append(f"Department Satisfaction has {len(data_rows2)} rows, expected 7")
            agent_lookup2 = {str(r[0]).strip().lower(): r for r in data_rows2 if r[0]}
            if gt_wb is not None:
                gt_rows2 = load_sheet_rows(gt_wb, "Department Satisfaction") or []
                gt_data2 = [r for r in gt_rows2[1:] if r and r[0] is not None]
                for g_row in gt_data2:
                    key = str(g_row[0]).strip().lower()
                    a_row = agent_lookup2.get(key)
                    if a_row is None:
                        errors.append(f"Department row missing for {g_row[0]}")
                        continue
                    a_padded = list(a_row) + [None] * (5 - len(a_row))
                    # Avg_Job_Satisfaction (tol 0.15)
                    if not num_close(a_padded[1], g_row[1], 0.15):
                        errors.append(f"{g_row[0]} Avg_Job_Satisfaction: {a_padded[1]} vs {g_row[1]}")
                    # Avg_Work_Life_Balance (tol 0.15)
                    if not num_close(a_padded[2], g_row[2], 0.15):
                        errors.append(f"{g_row[0]} Avg_Work_Life_Balance: {a_padded[2]} vs {g_row[2]}")
                    # Headcount (deterministic from DB, exact)
                    if not num_close(a_padded[3], g_row[3], 0):
                        errors.append(f"{g_row[0]} Headcount: {a_padded[3]} vs {g_row[3]}")
                    # Satisfaction_Rating (string match, case-insensitive)
                    a_rating = str(a_padded[4] or "").strip().lower()
                    g_rating = str(g_row[4] or "").strip().lower()
                    if a_rating != g_rating:
                        errors.append(f"{g_row[0]} Satisfaction_Rating: {a_padded[4]} vs {g_row[4]}")
            # Sort order: alphabetical by Department
            dept_names = [str(r[0]).strip() for r in data_rows2 if r and r[0]]
            sorted_depts = sorted(dept_names, key=lambda s: s.lower())
            if dept_names != sorted_depts:
                errors.append("Department Satisfaction not sorted alphabetically by Department")

        # ---------- Gap Analysis (all 3 rows) ----------
        rows3 = load_sheet_rows(wb, "Gap Analysis")
        if rows3 is None:
            errors.append("Sheet 'Gap Analysis' not found")
        else:
            data_rows3 = [r for r in rows3[1:] if r and r[0] is not None]
            if len(data_rows3) != 3:
                errors.append(f"Gap Analysis has {len(data_rows3)} rows, expected 3")
            agent_lookup3 = {str(r[0]).strip().lower(): r for r in data_rows3 if r[0]}
            if gt_wb is not None:
                gt_rows3 = load_sheet_rows(gt_wb, "Gap Analysis") or []
                gt_data3 = [r for r in gt_rows3[1:] if r and r[0] is not None]
                for g_row in gt_data3:
                    key = str(g_row[0]).strip().lower()
                    a_row = agent_lookup3.get(key)
                    if a_row is None:
                        errors.append(f"Gap row missing for {g_row[0]}")
                        continue
                    a_padded = list(a_row) + [None] * (5 - len(a_row))
                    # Our_Value (tol 0.2)
                    if not num_close(a_padded[1], g_row[1], 0.2):
                        errors.append(f"{g_row[0]} Our_Value: {a_padded[1]} vs {g_row[1]}")
                    # Market_Average (tol 1.0)
                    if not num_close(a_padded[2], g_row[2], 1.0):
                        errors.append(f"{g_row[0]} Market_Average: {a_padded[2]} vs {g_row[2]}")
                    # Gap (tol 1.0)
                    if not num_close(a_padded[3], g_row[3], 1.0):
                        errors.append(f"{g_row[0]} Gap: {a_padded[3]} vs {g_row[3]}")
                    # Priority (string match)
                    a_pri = str(a_padded[4] or "").strip().lower()
                    g_pri = str(g_row[4] or "").strip().lower()
                    if a_pri != g_pri:
                        errors.append(f"{g_row[0]} Priority: {a_padded[4]} vs {g_row[4]}")

    except Exception as e:
        errors.append(f"Error reading Excel: {e}")
    return errors


def check_gform():
    errors = []
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        cur.execute("SELECT id, title FROM gform.forms")
        forms = cur.fetchall()
        if not forms:
            errors.append("No Google Forms found")
        else:
            # Exact title match
            target = "employee benefits improvement survey"
            form_id = None
            for fid, title in forms:
                if (title or "").strip().lower() == target:
                    form_id = fid
                    break
            if form_id is None:
                errors.append(
                    f"Form titled 'Employee Benefits Improvement Survey' not found (forms: {[f[1] for f in forms]})"
                )
            else:
                cur.execute("SELECT COUNT(*) FROM gform.questions WHERE form_id = %s", (form_id,))
                q_count = cur.fetchone()[0]
                if q_count != 5:
                    errors.append(f"Form has {q_count} questions, expected exactly 5")
                # Verify required + question types
                cur.execute(
                    "SELECT title, question_type, required FROM gform.questions WHERE form_id = %s ORDER BY id",
                    (form_id,),
                )
                qrows = cur.fetchall()
                if qrows:
                    # Check that at least one dropdown and one multiple_choice question exist
                    types_lower = [str(r[1] or "").lower() for r in qrows]
                    has_dropdown = any(("drop" in t or "select" in t) for t in types_lower)
                    has_mc = any(("multiple" in t or "radio" in t or "choice" in t) for t in types_lower)
                    has_scale = any(("scale" in t or "linear" in t or "rating" in t) for t in types_lower)
                    if not has_dropdown:
                        errors.append("Form missing dropdown question (Q1: department)")
                    if not has_mc:
                        errors.append("Form missing multiple-choice question (Q5: priority)")
                    if not has_scale:
                        errors.append("Form missing linear-scale questions (Q2-Q4)")
                    # All questions required
                    not_required = [r[0] for r in qrows if not r[2]]
                    if not_required:
                        errors.append(f"Some questions are not required: {not_required}")
        cur.close()
        conn.close()
    except Exception as e:
        errors.append(f"Error checking GForm: {e}")
    return errors


def check_email():
    errors = []
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        cur.execute("""
            SELECT subject, body_text FROM email.messages
            WHERE to_addr::text ILIKE '%hr_leadership@company.com%'
            ORDER BY id DESC LIMIT 5
        """)
        rows = cur.fetchall()
        cur.close()
        conn.close()
        if not rows:
            errors.append("No email found to hr_leadership@company.com")
        else:
            # Find email matching exact subject
            target_subj = "benefits competitiveness analysis & survey launch"
            matched = None
            for subj, body in rows:
                if (subj or "").strip().lower() == target_subj:
                    matched = (subj, body)
                    break
            if matched is None:
                errors.append(
                    f"Email subject 'Benefits Competitiveness Analysis & Survey Launch' not found (subjects: {[r[0] for r in rows]})"
                )
            else:
                _, body = matched
                body_l = (body or "").lower()
                # Body should mention key terms
                keys = ["health", "pto", "retirement", "survey"]
                missing = [k for k in keys if k not in body_l]
                if missing:
                    errors.append(f"Email body missing key terms: {missing}")
                # Body should mention specific gap data: largest gap is Health Insurance with priority 'High'
                if "high" not in body_l:
                    errors.append("Email body missing 'High' priority indicator from Gap Analysis")
                if not any(t in body_l for t in ["health insurance", "health"]):
                    errors.append("Email body missing reference to top-gap category 'Health Insurance'")
    except Exception as e:
        errors.append(f"Error checking email: {e}")
    return errors


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()
    task_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    agent_ws = args.agent_workspace or os.path.join(task_root, "groundtruth_workspace")
    gt_ws = args.groundtruth_workspace or os.path.join(task_root, "groundtruth_workspace")

    all_errors = []

    print("  Checking Excel file...")
    errs = check_excel(agent_ws, gt_ws)
    if errs:
        all_errors.extend(errs)
        for e in errs[:5]:
            print(f"    ERROR: {e}")
    else:
        print("    PASS")

    print("  Checking Google Form...")
    errs = check_gform()
    if errs:
        all_errors.extend(errs)
        for e in errs[:3]:
            print(f"    ERROR: {e}")
    else:
        print("    PASS")

    print("  Checking email...")
    errs = check_email()
    if errs:
        all_errors.extend(errs)
        for e in errs[:3]:
            print(f"    ERROR: {e}")
    else:
        print("    PASS")

    if all_errors:
        print(f"\n=== RESULT: FAIL ({len(all_errors)} errors) ===")
        for e in all_errors[:15]:
            print(f"  {e}")
        sys.exit(1)
    else:
        print("\n=== RESULT: PASS ===")
        sys.exit(0)


if __name__ == "__main__":
    main()
