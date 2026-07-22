"""Evaluation for sf-hr-manager-report."""
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
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    task_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    gt_dir = args.groundtruth_workspace or os.path.join(task_root, "groundtruth_workspace")

    agent_file = os.path.join(args.agent_workspace, "HR_Manager_Report.xlsx")
    gt_file = os.path.join(gt_dir, "HR_Manager_Report.xlsx")

    if not os.path.exists(agent_file):
        print(f"FAIL: Agent output not found: {agent_file}")
        sys.exit(1)
    if not os.path.exists(gt_file):
        print(f"FAIL: Groundtruth not found: {gt_file}")
        sys.exit(1)

    agent_wb = openpyxl.load_workbook(agent_file, data_only=True)
    gt_wb = openpyxl.load_workbook(gt_file, data_only=True)

    all_errors = []
    
    # Check sheet: Manager Report
    print(f"  Checking Manager Report...")
    a_rows = load_sheet_rows(agent_wb, "Manager Report")
    g_rows = load_sheet_rows(gt_wb, "Manager Report")
    if a_rows is None:
        all_errors.append("Sheet 'Manager Report' not found in agent output")
    elif g_rows is None:
        all_errors.append("Sheet 'Manager Report' not found in groundtruth")
    else:
        sheet_name = "Manager Report"
        errors = []
        a_data = a_rows[1:] if len(a_rows) > 1 else []
        g_data = g_rows[1:] if len(g_rows) > 1 else []
        
        a_lookup = {}
        for row in a_data:
            if row and row[0] is not None:
                a_lookup[str(row[0]).strip().lower()] = row
        # Sort order check: by Department alphabetically
        a_dept_order = [str(r[0]).strip().lower() for r in a_data if r and r[0] is not None]
        if a_dept_order != sorted(a_dept_order):
            errors.append(f"Manager Report not sorted alphabetically: {a_dept_order}")
        for g_row in g_data:
            if not g_row or g_row[0] is None:
                continue
            key = str(g_row[0]).strip().lower()
            a_row = a_lookup.get(key)
            if a_row is None:
                errors.append(f"Missing row: {g_row[0]}")
                continue

            # Total - exact (deterministic)
            if len(a_row) > 1 and len(g_row) > 1:
                if not num_close(a_row[1], g_row[1], 1):
                    errors.append(f"{key}.Total: {a_row[1]} vs {g_row[1]} (tol=1)")

            # High_Performers - tight tolerance for deterministic count
            if len(a_row) > 2 and len(g_row) > 2:
                if not num_close(a_row[2], g_row[2], 1):
                    errors.append(f"{key}.High_Performers: {a_row[2]} vs {g_row[2]} (tol=1)")

            # Low_Performers - tight tolerance
            if len(a_row) > 3 and len(g_row) > 3:
                if not num_close(a_row[3], g_row[3], 1):
                    errors.append(f"{key}.Low_Performers: {a_row[3]} vs {g_row[3]} (tol=1)")

            # Avg_Salary rounded to integer - tol=1
            if len(a_row) > 4 and len(g_row) > 4:
                if not num_close(a_row[4], g_row[4], 1.0):
                    errors.append(f"{key}.Avg_Salary: {a_row[4]} vs {g_row[4]} (tol=1.0)")

            if len(a_row) > 5 and len(g_row) > 5:
                if not num_close(a_row[5], g_row[5], 0.2):
                    errors.append(f"{key}.Avg_Experience: {a_row[5]} vs {g_row[5]} (tol=0.2)")

            if len(a_row) > 6 and len(g_row) > 6:
                if not num_close(a_row[6], g_row[6], 0.2):
                    errors.append(f"{key}.High_Perf_Pct: {a_row[6]} vs {g_row[6]} (tol=0.2)")
        if errors:
            all_errors.extend(errors)
            print(f"    ERRORS: {len(errors)}")
            for e in errors[:5]:
                print(f"      {e}")
        else:
            print(f"    PASS")


    # Check sheet: Summary
    print(f"  Checking Summary...")
    a_rows = load_sheet_rows(agent_wb, "Summary")
    g_rows = load_sheet_rows(gt_wb, "Summary")
    if a_rows is None:
        all_errors.append("Sheet 'Summary' not found in agent output")
    elif g_rows is None:
        all_errors.append("Sheet 'Summary' not found in groundtruth")
    else:
        sheet_name = "Summary"
        errors = []
        a_data = a_rows[1:] if len(a_rows) > 1 else []
        g_data = g_rows[1:] if len(g_rows) > 1 else []
        
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
                errors.append(f"Missing row: {g_row[0]}")
                continue

            if len(a_row) > 1 and len(g_row) > 1:
                gv = g_row[1]
                av = a_row[1]
                # Best_Dept is non-numeric -> exact str match; numeric metrics tighter tol
                try:
                    gv_f = float(gv) if gv is not None else None
                except (TypeError, ValueError):
                    gv_f = None
                if gv_f is not None:
                    # Choose tolerance per metric type
                    if "pct" in key or "percent" in key:
                        tol = 0.2
                    elif "total" in key or "count" in key or "performers" in key:
                        tol = 1.0
                    else:
                        tol = 1.0
                    if not num_close(av, gv, tol):
                        errors.append(f"{key}.Value: {av} vs {gv} (tol={tol})")
                else:
                    # String comparison - exact
                    if not str_match(av, gv):
                        errors.append(f"{key}.Value: '{av}' vs '{gv}' (str_match)")
        if errors:
            all_errors.extend(errors)
            print(f"    ERRORS: {len(errors)}")
            for e in errors[:5]:
                print(f"      {e}")
        else:
            print(f"    PASS")

    # Check Notion page (BLOCKING) - extract page title from properties JSON and
    # require the title to contain 'HR Department Performance Report' (case-insensitive).
    print("  Checking Notion page (blocking)...")
    try:
        import json as _json
        import psycopg2
        conn = psycopg2.connect(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="toolathlon_gym",
                                user="eigent", password="camel")
        cur = conn.cursor()
        cur.execute("SELECT id, properties FROM notion.pages WHERE archived = false")
        pages = cur.fetchall()
        target_phrase = "hr department performance report"
        found_count = 0
        for pid, props in pages:
            if not props:
                continue
            try:
                if isinstance(props, str):
                    props_obj = _json.loads(props)
                else:
                    props_obj = props
                titles = []
                for k, v in (props_obj.items() if isinstance(props_obj, dict) else []):
                    if isinstance(v, dict) and v.get("type") == "title":
                        for t in v.get("title", []):
                            if isinstance(t, dict):
                                pt = t.get("plain_text", "") or (
                                    t.get("text", {}).get("content", "")
                                    if isinstance(t.get("text"), dict) else ""
                                )
                                if pt:
                                    titles.append(pt)
                title_text = " ".join(titles).strip().lower()
                if target_phrase in title_text:
                    found_count += 1
            except Exception:
                continue
        if found_count == 0:
            all_errors.append(
                "Notion page titled 'HR Department Performance Report' not found "
                "(extracted via title-property field)"
            )
        else:
            print(f"    PASS (found {found_count} matching Notion page titles)")
        cur.close()
        conn.close()
    except Exception as e:
        all_errors.append(f"Notion check failed: {e}")

    if all_errors:
        print(f"\n=== RESULT: FAIL ({len(all_errors)} errors) ===")
        for e in all_errors[:10]:
            print(f"  {e}")
        sys.exit(1)
    else:
        print("\n=== RESULT: PASS ===")
        sys.exit(0)


if __name__ == "__main__":
    main()
