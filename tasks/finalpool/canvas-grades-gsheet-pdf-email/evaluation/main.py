"""Evaluation script for canvas-grades-gsheet-pdf-email."""
import os
import argparse, json, os, sys
import openpyxl

def num_close(a, b, rel_tol=0.15, abs_tol=0.5):
    return abs(float(a) - float(b)) <= max(abs_tol, abs(float(b)) * rel_tol)


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

def safe_float(val, default=None):
    try:
        if val is None: return default
        return float(str(val).replace(",", "").replace("%", "").replace("$", "").strip())
    except (ValueError, TypeError):
        return default

def get_conn():
    import psycopg2
    return psycopg2.connect(**DB_CONFIG)

def get_gt_data(groundtruth_workspace):
    """Load reference data from groundtruth Grade_Dashboard_Reference.xlsx."""
    gt_path = os.path.join(groundtruth_workspace, "Grade_Dashboard_Reference.xlsx")
    if not os.path.exists(gt_path):
        return None
    wb = openpyxl.load_workbook(gt_path, data_only=True)
    out = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if rows:
            headers = [str(c).strip().lower() if c else "" for c in rows[0]]
            out[sheet_name] = {"headers": headers, "rows": [tuple(r) for r in rows[1:]]}
    return out


def check_gsheet_against_gt(cur, gt_data):
    """Verify the 'Department Grade Dashboard' Google Sheet matches the reference data."""
    cur.execute("SELECT id, title FROM gsheet.spreadsheets")
    sheets = cur.fetchall()
    target_id = None
    target_title = None
    for sid, title in sheets:
        if title and "department grade dashboard" in title.strip().lower():
            target_id = sid
            target_title = title
            break
    check("Google Sheet 'Department Grade Dashboard' exists", target_id is not None,
          f"got: {[t for _, t in sheets]}")
    if target_id is None:
        return

    # List sheets / tabs
    cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s", (target_id,))
    tabs = cur.fetchall()
    tab_titles_lower = {t[1].strip().lower(): t[0] for t in tabs if t[1]}

    for required in ("grade_distribution", "department_summary"):
        present = any(required.replace("_", "") == k.replace("_", "").replace(" ", "")
                      or required in k for k in tab_titles_lower)
        check(f"Tab '{required}' present in Google Sheet", present,
              f"tabs: {list(tab_titles_lower)}")

    # Fetch all cells for the spreadsheet
    cur.execute(
        "SELECT sheet_id, row_index, col_index, value FROM gsheet.cells WHERE spreadsheet_id = %s",
        (target_id,)
    )
    cells = cur.fetchall()
    by_sheet = {}
    for sheet_id, ri, ci, val in cells:
        by_sheet.setdefault(sheet_id, {})[(ri, ci)] = val

    # For each GT sheet, locate the matching tab and compare
    for gt_sheet_name, info in (gt_data or {}).items():
        sheet_key = gt_sheet_name.strip().lower().replace("_", "")
        match_id = None
        for tab_title, tid in tab_titles_lower.items():
            if tab_title.replace("_", "").replace(" ", "") == sheet_key:
                match_id = tid; break
        if match_id is None:
            continue  # already flagged above
        cell_map = by_sheet.get(match_id, {})
        if not cell_map:
            check(f"Tab '{gt_sheet_name}' has cells", False, "no cells found")
            continue
        # Build header -> col_index from cell row 0
        header_cols = {}
        for (r, c), v in cell_map.items():
            if r == 0 and v:
                header_cols[str(v).strip().lower()] = c
        gt_headers = info["headers"]
        for h in gt_headers:
            if h and h not in header_cols:
                check(f"GSheet '{gt_sheet_name}' has column '{h}'", False,
                      f"got: {list(header_cols)}")
        # Compare every GT row by Course_Name (or Metric)
        # Find key column index in cell_map
        key_label = gt_headers[0]
        key_col = header_cols.get(key_label)
        if key_col is None:
            check(f"GSheet '{gt_sheet_name}' has key column '{key_label}'", False)
            continue
        # Build agent_lookup: {key_lower: {col_label: val}}
        agent_lookup = {}
        max_row = max((r for (r, c) in cell_map.keys()), default=0)
        for r in range(1, max_row + 1):
            kval = cell_map.get((r, key_col))
            if kval is None:
                continue
            key_lower = str(kval).strip().lower()
            row_data = {}
            for label, c in header_cols.items():
                row_data[label] = cell_map.get((r, c))
            agent_lookup[key_lower] = row_data

        for gt_row in info["rows"]:
            if not gt_row or gt_row[0] is None:
                continue
            key_lower = str(gt_row[0]).strip().lower()
            agent_row = agent_lookup.get(key_lower)
            if agent_row is None:
                check(f"GSheet '{gt_sheet_name}' row '{gt_row[0]}' present", False)
                continue
            for ci, h in enumerate(gt_headers):
                if ci == 0 or not h or ci >= len(gt_row):
                    continue
                gv = gt_row[ci]
                av = agent_row.get(h)
                gf = safe_float(gv)
                af = safe_float(av)
                if gf is not None and af is not None:
                    tol = max(0.5, abs(gf) * 0.05)
                    if abs(gf - af) > tol:
                        check(f"GSheet '{gt_sheet_name}' '{gt_row[0]}' {h}={gf:.1f}",
                              False, f"got {af}")
                elif gv is not None:
                    gs = str(gv).strip().lower()
                    avs = str(av or "").strip().lower()
                    if gs and gs != avs:
                        check(f"GSheet '{gt_sheet_name}' '{gt_row[0]}' {h}",
                              False, f"expected '{gs[:50]}', got '{avs[:50]}'")


def run_evaluation(agent_workspace, groundtruth_workspace, launch_time, res_log_file):
    global PASS_COUNT, FAIL_COUNT
    PASS_COUNT = 0
    FAIL_COUNT = 0

    gt_data = get_gt_data(groundtruth_workspace)

    # Check Python script exists (terminal usage required by task)
    py_files = [f for f in os.listdir(agent_workspace) if f.endswith(".py")]
    check("Python analysis script exists", len(py_files) >= 1, f"found: {py_files}")
    # Specifically grade_reporter.py mentioned in task
    check("grade_reporter.py exists",
          any("grade_reporter" in f for f in py_files),
          f"py files: {py_files}")
    # Output JSON files mentioned in task
    has_grade_data = os.path.exists(os.path.join(agent_workspace, "grade_data.json"))
    has_grade_report = os.path.exists(os.path.join(agent_workspace, "grade_report.json"))
    check("grade_data.json exists", has_grade_data)
    check("grade_report.json exists", has_grade_report)

    # Database checks
    try:
        conn = get_conn()
        cur = conn.cursor()

        # Email check - require exact subject and recipient
        cur.execute(
            "SELECT subject, to_addr, body_text FROM email.messages "
            "WHERE folder_id = (SELECT id FROM email.folders WHERE name = 'Sent' LIMIT 1)"
        )
        sent = cur.fetchall()
        target_subj = "q1 2026 grade distribution report"
        match = None
        for subj, to_addr, body in sent:
            subj_l = (subj or "").lower()
            to_l = str(to_addr or "").lower()
            if target_subj in subj_l and "dept-heads@university.edu" in to_l:
                match = (subj, to_addr, body)
                break
        check("Email to dept-heads@university.edu with correct subject", match is not None,
              f"{len(sent)} sent emails")
        if match is not None:
            body_l = (match[2] or "").lower()
            check("Email body mentions overall pass rate",
                  "pass" in body_l and ("rate" in body_l or "%" in body_l))
            # Tightened: require the actual highest/lowest course names from GT to appear in body
            # (not just generic words "highest"/"lowest")
            highest_name = None
            lowest_name = None
            if gt_data and "Department_Summary" in gt_data:
                for row in gt_data["Department_Summary"]["rows"]:
                    if row and len(row) >= 2:
                        label = str(row[0] or "").strip().lower()
                        val = str(row[1] or "").strip()
                        if label == "highest_avg_course":
                            highest_name = val
                        elif label == "lowest_avg_course":
                            lowest_name = val
            if highest_name:
                check(f"Email body mentions highest course '{highest_name}'",
                      highest_name.lower() in body_l)
            else:
                check("Email body mentions highest course (label-only fallback)",
                      "highest" in body_l)
            if lowest_name:
                check(f"Email body mentions lowest course '{lowest_name}'",
                      lowest_name.lower() in body_l)
            else:
                check("Email body mentions lowest course (label-only fallback)",
                      "lowest" in body_l)

        # Google Sheet check using gt_data
        if gt_data is not None:
            check_gsheet_against_gt(cur, gt_data)
        else:
            cur.execute("SELECT COUNT(*) FROM gsheet.spreadsheets")
            check("Google Sheet created", cur.fetchone()[0] >= 1)

        # Reverse verification: noise emails should not be forwarded
        cur.execute(
            "SELECT COUNT(*) FROM email.messages "
            "WHERE folder_id = (SELECT id FROM email.folders WHERE name = 'Sent' LIMIT 1) "
            "AND subject ILIKE '%newsletter%'"
        )
        noise_sent = cur.fetchone()[0]
        check("No noise emails forwarded", noise_sent == 0,
              f"found {noise_sent} noise emails in Sent")

        conn.close()
    except Exception as e:
        check("DB checks", False, str(e))

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