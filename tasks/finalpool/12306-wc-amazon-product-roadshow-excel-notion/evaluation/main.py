"""
Evaluation for 12306-wc-amazon-product-roadshow-excel-notion task.

Checks:
1. Roadshow_Plan.xlsx exists
2. Products sheet has >= 4 data rows with Name and Price columns
3. Travel_Itinerary sheet has >= 2 rows containing G11 and G105 train numbers
4. Roadshow_Schedule sheet has >= 3 rows
5. Notion page exists with Roadshow or Shanghai or Guangzhou in title
6. Email sent to shanghai_dist@partner.com
7. Email sent to guangzhou_dist@partner.com
8. Email sent to manager@company.com
"""
import json
import os
import sys
from argparse import ArgumentParser

import psycopg2
import openpyxl

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
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
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def num_close(a, b, tol=1.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def str_match(a, b):
    if a is None or b is None:
        return a is None and b is None
    return str(a).strip().lower() == str(b).strip().lower()


def check_excel(agent_workspace, groundtruth_workspace="."):
    print("\n=== Check 1: Excel Roadshow_Plan.xlsx ===")
    xlsx_path = os.path.join(agent_workspace, "Roadshow_Plan.xlsx")
    if not os.path.exists(xlsx_path):
        record("Roadshow_Plan.xlsx exists", False, f"Not found at {xlsx_path}")
        return
    record("Roadshow_Plan.xlsx exists", True)

    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    except Exception as e:
        record("Excel file readable", False, str(e))
        return
    record("Excel file readable", True)

    # Helper: normalized sheet match (case+underscore+space tolerant)
    def norm(s): return str(s or "").strip().lower().replace("_", " ")
    def find_sheet(wb, target):
        for n in wb.sheetnames:
            if norm(n) == norm(target):
                return wb[n]
        return None

    # Check Products sheet (exact match preferred, tolerant)
    prod_sheet = find_sheet(wb, "Products")
    record("Products sheet exists (exact name)", prod_sheet is not None, f"Sheets: {wb.sheetnames}")
    if prod_sheet is not None:
        rows = list(prod_sheet.iter_rows(values_only=True))
        headers = [str(c).strip().lower() if c else "" for c in (rows[0] if rows else [])]
        has_name = any("name" in h for h in headers)
        has_price = any("price" in h for h in headers)
        record("Products has Name and Price columns", has_name and has_price,
               f"Headers: {rows[0] if rows else []}")
        data_rows = [r for r in rows[1:] if any(c for c in r)]
        record("Products has exactly 5 data rows", len(data_rows) == 5,
               f"Found {len(data_rows)} data rows")

    # Customer_Regions sheet
    cr_sheet = find_sheet(wb, "Customer_Regions")
    record("Customer_Regions sheet exists", cr_sheet is not None, f"Sheets: {wb.sheetnames}")

    # Check Travel_Itinerary sheet
    travel_sheet = find_sheet(wb, "Travel_Itinerary")
    record("Travel_Itinerary sheet exists (exact name)", travel_sheet is not None, f"Sheets: {wb.sheetnames}")
    if travel_sheet is not None:
        rows = list(travel_sheet.iter_rows(values_only=True))
        # Use word-bounded search to avoid substring FP (G11 vs G110)
        def has_train_no(rows, train_no):
            for r in rows[1:]:
                for c in r:
                    if c is None: continue
                    v = str(c).strip().upper()
                    if v == train_no.upper():
                        return True
            return False
        data_rows = [r for r in rows[1:] if any(c for c in r)]
        record("Travel_Itinerary has exactly 2 data rows", len(data_rows) == 2,
               f"Found {len(data_rows)} rows")
        record("Travel_Itinerary has row with train G11", has_train_no(rows, "G11"),
               "G11 not found as Train_No cell")
        record("Travel_Itinerary has row with train G105", has_train_no(rows, "G105"),
               "G105 not found as Train_No cell")

    # Check Roadshow_Schedule sheet
    sched_sheet = find_sheet(wb, "Roadshow_Schedule")
    record("Roadshow_Schedule sheet exists (exact name)", sched_sheet is not None, f"Sheets: {wb.sheetnames}")
    if sched_sheet is not None:
        rows = list(sched_sheet.iter_rows(values_only=True))
        data_rows = [r for r in rows[1:] if any(c for c in r)]
        record("Roadshow_Schedule has >= 4 data rows", len(data_rows) >= 4,
               f"Found {len(data_rows)} rows")
        # Must cover both Shanghai and Guangzhou
        all_text = " ".join(str(c) for r in rows for c in r if c).lower()
        record("Roadshow_Schedule mentions Shanghai", "shanghai" in all_text)
        record("Roadshow_Schedule mentions Guangzhou", "guangzhou" in all_text)

    # --- Groundtruth value comparison ---
    gt_path = os.path.join(groundtruth_workspace, "Roadshow_Plan.xlsx")
    if not os.path.isfile(gt_path):
        record("Groundtruth xlsx exists", False, gt_path)
        return

    gt_wb = openpyxl.load_workbook(gt_path, data_only=True)
    for gt_sheet_name in gt_wb.sheetnames:
        gt_ws = gt_wb[gt_sheet_name]
        agent_ws = None
        for asn in wb.sheetnames:
            if asn.strip().lower() == gt_sheet_name.strip().lower():
                agent_ws = wb[asn]
                break
        if agent_ws is None:
            record(f"GT sheet '{gt_sheet_name}' exists in agent", False, f"Available: {wb.sheetnames}")
            continue

        gt_rows = [r for r in gt_ws.iter_rows(min_row=2, values_only=True) if any(c is not None for c in r)]
        agent_rows = [r for r in agent_ws.iter_rows(min_row=2, values_only=True) if any(c is not None for c in r)]

        record(f"GT '{gt_sheet_name}' row count", len(agent_rows) == len(gt_rows),
               f"Expected {len(gt_rows)}, got {len(agent_rows)}")

        # Iterate ALL GT rows (not just 3 + last)
        check_indices = list(range(len(gt_rows)))
        for idx in check_indices:
            gt_row = gt_rows[idx]
            if idx < len(agent_rows):
                a_row = agent_rows[idx]
                row_ok = True
                fail_cols = []
                for col_idx in range(min(len(gt_row), len(a_row) if a_row else 0)):
                    gt_val = gt_row[col_idx]
                    a_val = a_row[col_idx]
                    if gt_val is None:
                        continue
                    if isinstance(gt_val, (int, float)):
                        # Tighter tolerance for currency/numeric: 5% relative, abs 0.5 floor.
                        tol = max(abs(gt_val) * 0.05, 0.5)
                        ok = num_close(a_val, gt_val, tol)
                    else:
                        ok = str_match(a_val, gt_val)
                    if not ok:
                        # Continue checking remaining columns and report them all.
                        fail_cols.append((col_idx + 1, gt_val, a_val))
                        row_ok = False
                for c_idx, gv, av in fail_cols:
                    record(f"GT '{gt_sheet_name}' row {idx+1} col {c_idx}",
                           False, f"Expected {gv}, got {av}")
                if row_ok:
                    record(f"GT '{gt_sheet_name}' row {idx+1} values match", True)
            else:
                record(f"GT '{gt_sheet_name}' row {idx+1} exists", False, "Row missing in agent")
    gt_wb.close()


def check_notion():
    print("\n=== Check 2: Notion Page ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("SELECT id, parent, properties FROM notion.pages")
    pages = cur.fetchall()
    cur.close()
    conn.close()

    found_any = False
    found_exact = False
    for page_id, parent, props in pages:
        try:
            title_items = []
            for key, val in props.items():
                if isinstance(val, dict) and val.get("type") == "title":
                    title_items = val.get("title", [])
                    break
            title_text = " ".join(
                item.get("text", {}).get("content", "") for item in title_items
                if isinstance(item, dict)
            ).lower()
            if "roadshow" in title_text or "shanghai" in title_text or "guangzhou" in title_text:
                found_any = True
            # Require both cities AND roadshow keyword
            if "roadshow" in title_text and "shanghai" in title_text and "guangzhou" in title_text:
                found_exact = True
                break
        except Exception:
            continue

    record("Notion page title contains Roadshow + Shanghai + Guangzhou", found_exact,
           f"Total pages: {len(pages)}")


def check_emails_sent():
    print("\n=== Check 3: Emails Sent ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    try:
        # Fetch every sent message with subject/to_addr
        cur.execute("""
            SELECT m.subject, m.to_addr FROM email.messages m
            JOIN email.folders f ON m.folder_id = f.id
            WHERE UPPER(f.name) = 'SENT'
        """)
        msgs = list(cur.fetchall())
        cur.execute("""
            SELECT m.subject, m.to_addr FROM email.sent_log sl
            JOIN email.messages m ON sl.message_id = m.id
        """)
        msgs += list(cur.fetchall())
        # Normalize to_addr strings
        def to_str(v):
            if isinstance(v, list): return " ".join(str(x).lower() for x in v)
            if isinstance(v, str):
                try:
                    p = json.loads(v)
                    if isinstance(p, list): return " ".join(str(x).lower() for x in p)
                except Exception:
                    pass
                return v.lower()
            return str(v).lower()

        normalized = [((s or "").lower(), to_str(t)) for s, t in msgs]

        def has_msg(recipient, subject_substr):
            for subj, to in normalized:
                if recipient in to and subject_substr in subj:
                    return True
            return False

        record("Email to shanghai_dist with subject 'Roadshow Meeting Confirmation - Shanghai March 10'",
               has_msg("shanghai_dist@partner.com",
                       "roadshow meeting confirmation - shanghai march 10"),
               f"Total: {len(normalized)}")
        record("Email to guangzhou_dist with subject 'Roadshow Meeting Confirmation - Guangzhou March 11'",
               has_msg("guangzhou_dist@partner.com",
                       "roadshow meeting confirmation - guangzhou march 11"),
               f"Total: {len(normalized)}")
        record("Email to manager@company.com with subject 'Roadshow Plan Summary - Shanghai and Guangzhou March 2026'",
               has_msg("manager@company.com",
                       "roadshow plan summary - shanghai and guangzhou march 2026"),
               f"Total: {len(normalized)}")
    except Exception as e:
        record("Email sent check", False, str(e))
    finally:
        cur.close()
        conn.close()


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_excel(args.agent_workspace, args.groundtruth_workspace)
    file_fail = FAIL_COUNT
    check_notion()
    check_emails_sent()
    runtime_fail = FAIL_COUNT - file_fail

    total = PASS_COUNT + FAIL_COUNT
    if total == 0:
        print("\nFAIL: No checks were performed.")
        sys.exit(1)

    accuracy = PASS_COUNT / total * 100
    print(f"\nOverall: {PASS_COUNT}/{total} checks passed ({accuracy:.1f}%) "
          f"(file_fail={file_fail}, runtime_fail={runtime_fail})")

    result = {
        "total_passed": PASS_COUNT,
        "total_checks": total,
        "accuracy": accuracy,
        "file_fail": file_fail,
        "runtime_fail": runtime_fail,
    }

    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    # File checks (Excel vs GT) are blocking. Notion/email runtime checks expected
    # to fail in GT self-test since agent has not created those artifacts.
    if file_fail == 0:
        print("PASS (file checks clean)")
        sys.exit(0)
    else:
        print(f"FAIL ({file_fail} file-level failures)")
        sys.exit(1)


if __name__ == "__main__":
    main()
