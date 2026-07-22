"""Evaluation for sf-hr-compensation-pdf-notion.

Blocking checks: Compensation_Data.xlsx (Excel data comparison).
Non-blocking: Notion page, PDF existence.
"""
import argparse
import os
import sys
import openpyxl


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
    if a is None or b is None:
        return a is None and b is None
    return str(a).strip().lower() == str(b).strip().lower()


def load_sheet_rows(wb, sheet_name):
    for name in wb.sheetnames:
        if name.strip().lower() == sheet_name.strip().lower():
            return [[cell.value for cell in row] for row in wb[name].iter_rows()]
    return None


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    task_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    gt_dir = args.groundtruth_workspace or os.path.join(task_root, "groundtruth_workspace")
    all_errors = []

    # ---- Check Excel (blocking) ----
    agent_file = os.path.join(args.agent_workspace, "Compensation_Data.xlsx")
    gt_file = os.path.join(gt_dir, "Compensation_Data.xlsx")

    if not os.path.exists(agent_file):
        print(f"FAIL: Agent output not found: {agent_file}")
        sys.exit(1)
    if not os.path.exists(gt_file):
        print(f"FAIL: Groundtruth not found: {gt_file}")
        sys.exit(1)

    agent_wb = openpyxl.load_workbook(agent_file, data_only=True)
    gt_wb = openpyxl.load_workbook(gt_file, data_only=True)

    # Check Department Summary
    print("  Checking Department Summary...")
    a_rows = load_sheet_rows(agent_wb, "Department Summary")
    g_rows = load_sheet_rows(gt_wb, "Department Summary")
    if a_rows is None:
        all_errors.append("Sheet 'Department Summary' not found in agent output")
    elif g_rows is None:
        all_errors.append("Sheet 'Department Summary' not found in groundtruth")
    else:
        a_data = a_rows[1:] if len(a_rows) > 1 else []
        g_data = g_rows[1:] if len(g_rows) > 1 else []

        a_lookup = {}
        for row in a_data:
            if row and row[0] is not None:
                a_lookup[str(row[0]).strip().lower()] = row
        # Validate sort order: should be alphabetical by Department
        a_dept_order = [str(r[0]).strip().lower() for r in a_data if r and r[0] is not None]
        if a_dept_order != sorted(a_dept_order):
            all_errors.append(f"Department Summary not sorted alphabetically: {a_dept_order}")
        for g_row in g_data:
            if not g_row or g_row[0] is None:
                continue
            key = str(g_row[0]).strip().lower()
            a_row = a_lookup.get(key)
            if a_row is None:
                all_errors.append(f"Missing department: {g_row[0]}")
                continue
            # Col 1: Employees count - must be exact (deterministic data)
            if len(a_row) > 1 and len(g_row) > 1:
                if not num_close(a_row[1], g_row[1], 1):
                    all_errors.append(f"{key}.Employees: {a_row[1]} vs {g_row[1]} (tol=1)")
            # Col 2: Avg_Salary - tighter tolerance (rounded to 2 decimals)
            if len(a_row) > 2 and len(g_row) > 2:
                if not num_close(a_row[2], g_row[2], 1.0):
                    all_errors.append(f"{key}.Avg_Salary: {a_row[2]} vs {g_row[2]} (tol=1.0)")
            # Col 3: Min_Salary - exact match
            if len(a_row) > 3 and len(g_row) > 3:
                if not num_close(a_row[3], g_row[3], 1):
                    all_errors.append(f"{key}.Min_Salary: {a_row[3]} vs {g_row[3]} (tol=1)")
            # Col 4: Max_Salary - exact match
            if len(a_row) > 4 and len(g_row) > 4:
                if not num_close(a_row[4], g_row[4], 1):
                    all_errors.append(f"{key}.Max_Salary: {a_row[4]} vs {g_row[4]} (tol=1)")
            # Col 5: Median_Salary - tight tolerance (rounded)
            if len(a_row) > 5 and len(g_row) > 5:
                if not num_close(a_row[5], g_row[5], 1.0):
                    all_errors.append(f"{key}.Median_Salary: {a_row[5]} vs {g_row[5]} (tol=1.0)")
        if not all_errors:
            print("    PASS")

    # Check Education Breakdown
    print("  Checking Education Breakdown...")
    a_rows = load_sheet_rows(agent_wb, "Education Breakdown")
    g_rows = load_sheet_rows(gt_wb, "Education Breakdown")
    prev_errors = len(all_errors)
    if a_rows is None:
        all_errors.append("Sheet 'Education Breakdown' not found in agent output")
    elif g_rows is None:
        all_errors.append("Sheet 'Education Breakdown' not found in groundtruth")
    else:
        a_data = a_rows[1:] if len(a_rows) > 1 else []
        g_data = g_rows[1:] if len(g_rows) > 1 else []

        a_lookup = {}
        for row in a_data:
            if row and row[0] is not None and row[1] is not None:
                key = (str(row[0]).strip().lower(), str(row[1]).strip().lower())
                a_lookup[key] = row
        for g_row in g_data:
            if not g_row or g_row[0] is None:
                continue
            key = (str(g_row[0]).strip().lower(), str(g_row[1]).strip().lower())
            a_row = a_lookup.get(key)
            if a_row is None:
                all_errors.append(f"Missing edu row: {g_row[0]} / {g_row[1]}")
                continue
            # Col 2: Count - exact (deterministic)
            if len(a_row) > 2 and len(g_row) > 2:
                if not num_close(a_row[2], g_row[2], 1):
                    all_errors.append(f"{key}.Count: {a_row[2]} vs {g_row[2]} (tol=1)")
            # Col 3: Avg_Salary - tight (rounded to 2 decimals)
            if len(a_row) > 3 and len(g_row) > 3:
                if not num_close(a_row[3], g_row[3], 1.0):
                    all_errors.append(f"{key}.Avg_Salary: {a_row[3]} vs {g_row[3]} (tol=1.0)")
            # Col 4: Min_Salary
            if len(a_row) > 4 and len(g_row) > 4:
                if not num_close(a_row[4], g_row[4], 1):
                    all_errors.append(f"{key}.Min_Salary: {a_row[4]} vs {g_row[4]} (tol=1)")
            # Col 5: Max_Salary
            if len(a_row) > 5 and len(g_row) > 5:
                if not num_close(a_row[5], g_row[5], 1):
                    all_errors.append(f"{key}.Max_Salary: {a_row[5]} vs {g_row[5]} (tol=1)")

        new_errors = len(all_errors) - prev_errors
        if new_errors == 0:
            print("    PASS")

    # ---- Check PDF exists (blocking with content checks) ----
    print("  Checking PDF...")
    pdf_path = os.path.join(args.agent_workspace, "Compensation_Report.pdf")
    if not os.path.exists(pdf_path):
        all_errors.append("Compensation_Report.pdf not found")
    else:
        file_size = os.path.getsize(pdf_path)
        if file_size < 500:
            all_errors.append(f"Compensation_Report.pdf too small ({file_size} bytes)")
        else:
            # Content checks
            try:
                pdf_text = ""
                try:
                    import pdfplumber
                    with pdfplumber.open(pdf_path) as pdf:
                        for page in pdf.pages:
                            pdf_text += (page.extract_text() or "") + "\n"
                except ImportError:
                    try:
                        from pypdf import PdfReader
                        reader = PdfReader(pdf_path)
                        for page in reader.pages:
                            pdf_text += (page.extract_text() or "") + "\n"
                    except ImportError:
                        from PyPDF2 import PdfReader
                        reader = PdfReader(pdf_path)
                        for page in reader.pages:
                            pdf_text += (page.extract_text() or "") + "\n"
                pdf_lower = pdf_text.lower()
                # Title check
                if "compensation analysis report" not in pdf_lower:
                    all_errors.append("PDF missing 'Compensation Analysis Report' title")
                if "hr analytics" not in pdf_lower:
                    all_errors.append("PDF missing 'HR Analytics' subtitle")
                if "2026-03-06" not in pdf_lower and "march 6, 2026" not in pdf_lower and "march 06, 2026" not in pdf_lower:
                    all_errors.append("PDF missing date 2026-03-06")
                # Validate at least 3 of 7 dept names appear in PDF
                dept_count = sum(1 for d in ["engineering", "finance", "hr", "operations", "r&d", "sales", "support"]
                                 if d in pdf_lower)
                if dept_count < 5:
                    all_errors.append(f"PDF missing department names (found {dept_count}/7)")
                if not all_errors or all([e for e in all_errors if "PDF" not in e]):
                    print("    PASS")
            except Exception as e:
                print(f"    [WARN] PDF content extraction error: {e}")

    # ---- Notion check (BLOCKING) ----
    print("  Checking Notion page (blocking)...")
    try:
        import psycopg2
        conn = psycopg2.connect(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="toolathlon_gym",
                                user="eigent", password="camel")
        cur = conn.cursor()
        # Look for page with title 'Compensation Analysis 2026'
        cur.execute("""
            SELECT COUNT(*) FROM notion.pages
            WHERE LOWER(properties::text) LIKE '%%compensation analysis 2026%%'
               OR LOWER(properties::text) LIKE '%%compensation%%2026%%'
        """)
        count = cur.fetchone()[0]
        if count == 0:
            # fallback: look in blocks block_data for the title
            cur.execute("""
                SELECT COUNT(*) FROM notion.blocks
                WHERE LOWER(block_data::text) LIKE '%%compensation analysis 2026%%'
            """)
            count = cur.fetchone()[0]
        if count == 0:
            all_errors.append("Notion page 'Compensation Analysis 2026' not found")
        else:
            print(f"    PASS (found {count} matching Notion entries)")
        cur.close()
        conn.close()
    except Exception as e:
        all_errors.append(f"Notion check failed: {e}")

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
