"""
Evaluation for 12306-wc-supplier-visit-excel-email-gcal task.

Checks:
1. Supplier_Visit_Plan.xlsx exists with Products, Travel_Plan, Visit_Schedule sheets
2. Products sheet has >= 5 rows with Product_ID and Supplier_Name columns
3. Travel_Plan has Beijing-Shanghai and Beijing-Guangzhou train entries (word-boundary regex)
4. Visit_Schedule has >= 5 rows
5. GCal has >= 2 supplier visit events on 2026-03-10
6. Outgoing emails sent to required addresses with meeting/visit subject keywords
"""
import json
import os
import re
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
    print("\n=== Check 1: Excel Supplier_Visit_Plan.xlsx ===")

    xlsx_path = os.path.join(agent_workspace, "Supplier_Visit_Plan.xlsx")
    if not os.path.exists(xlsx_path):
        record("Supplier_Visit_Plan.xlsx exists", False, f"Not found at {xlsx_path}")
        return
    record("Supplier_Visit_Plan.xlsx exists", True)

    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    except Exception as e:
        record("Excel readable", False, str(e))
        return
    record("Excel readable", True)

    sheet_names_lower = [s.lower() for s in wb.sheetnames]
    has_products = any("product" in s for s in sheet_names_lower)
    has_travel = any("travel" in s for s in sheet_names_lower)
    has_schedule = any("schedule" in s or "visit" in s for s in sheet_names_lower)

    record("Excel has Products sheet", has_products, f"Sheets: {wb.sheetnames}")
    record("Excel has Travel_Plan sheet", has_travel, f"Sheets: {wb.sheetnames}")
    record("Excel has Visit_Schedule sheet", has_schedule, f"Sheets: {wb.sheetnames}")

    if has_products:
        ws_name = wb.sheetnames[next(i for i, s in enumerate(sheet_names_lower) if "product" in s)]
        ws = wb[ws_name]
        rows = list(ws.iter_rows(values_only=True))
        data_rows = [r for r in rows[1:] if any(c for c in r)]
        record("Products sheet has >= 5 rows", len(data_rows) >= 5, f"Found {len(data_rows)} rows")

        if rows:
            headers = [str(c).lower() if c else "" for c in rows[0]]
            has_id = any("id" in h or "product" in h for h in headers)
            has_supplier = any("supplier" in h for h in headers)
            record("Products has Product_ID column", has_id, f"Headers: {rows[0]}")
            record("Products has Supplier_Name column", has_supplier, f"Headers: {rows[0]}")

    if has_travel:
        ws_name = wb.sheetnames[next(i for i, s in enumerate(sheet_names_lower) if "travel" in s)]
        ws = wb[ws_name]
        rows = list(ws.iter_rows(values_only=True))
        data_rows = [r for r in rows[1:] if any(c for c in r)]
        record("Travel_Plan has >= 2 rows", len(data_rows) >= 2, f"Found {len(data_rows)} rows")

        all_text = " ".join(str(c) for row in rows for c in row if c)
        # Use word-boundary regex (case-insensitive) to avoid false positives like 'G110'/'G1050'.
        # Accept any high-speed rail (G-class) train number; require at least 2 distinct G-trains.
        g_trains = set(m.group(0).upper() for m in re.finditer(r"\b[Gg]\d{1,4}\b", all_text))
        record("Travel_Plan contains >=2 distinct G-class train codes",
               len(g_trains) >= 2, f"Trains found: {sorted(g_trains)}")
        # Validate route coverage: Shanghai and Guangzhou both referenced.
        text_lower = all_text.lower()
        record("Travel_Plan references Shanghai", "shanghai" in text_lower or "上海" in text_lower,
               f"Content: {all_text[:200]}")
        record("Travel_Plan references Guangzhou", "guangzhou" in text_lower or "广州" in text_lower,
               f"Content: {all_text[:200]}")

    if has_schedule:
        idx = next(i for i, s in enumerate(sheet_names_lower) if "schedule" in s or "visit" in s)
        ws_name = wb.sheetnames[idx]
        ws = wb[ws_name]
        rows = list(ws.iter_rows(values_only=True))
        data_rows = [r for r in rows[1:] if any(c for c in r)]
        record("Visit_Schedule has >= 5 rows", len(data_rows) >= 5, f"Found {len(data_rows)} rows")

    # --- Groundtruth value comparison ---
    gt_path = os.path.join(groundtruth_workspace, "Supplier_Visit_Plan.xlsx")
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

        # Allow agent to include >= GT rows (>=N consistency with the >=5 thresholds above).
        record(f"GT '{gt_sheet_name}' row count >= {len(gt_rows)}",
               len(agent_rows) >= len(gt_rows),
               f"Expected >={len(gt_rows)}, got {len(agent_rows)}")

        # Build agent supplier-name index for set-based comparison (order-independent).
        # GT row[5] (Products) / row[1] (Visit_Schedule) / row[0] (Travel_Plan) hold a supplier or city
        # but the order may differ between agent's choice and GT. Match GT rows to agent rows via
        # primary key (Supplier_Name for Products/Visit_Schedule, Supplier_City for Travel_Plan).
        sheet_lower = gt_sheet_name.lower()
        if "product" in sheet_lower:
            key_col = 5  # Supplier_Name
        elif "travel" in sheet_lower:
            key_col = 0  # Supplier_City
        elif "schedule" in sheet_lower or "visit" in sheet_lower:
            key_col = 1  # Supplier_Name
        else:
            key_col = 0

        def _key(row):
            try:
                return str(row[key_col] or "").strip().lower() if row else ""
            except IndexError:
                return ""

        agent_index = {}
        for r in agent_rows:
            k = _key(r)
            if k:
                agent_index.setdefault(k, []).append(r)

        check_count = min(3, len(gt_rows))
        # Check first N GT rows by primary key match.
        for idx in range(check_count):
            gt_row = gt_rows[idx]
            k = _key(gt_row)
            if not k:
                continue
            matched_rows = agent_index.get(k, [])
            if not matched_rows:
                record(f"GT '{gt_sheet_name}' row[{idx+1}] key='{gt_row[key_col]}' present",
                       False, f"Agent missing supplier/city key '{k}'")
                continue
            a_row = matched_rows[0]
            row_ok = True
            mismatched_field = None
            for col_idx in range(min(len(gt_row), len(a_row))):
                gt_val = gt_row[col_idx]
                a_val = a_row[col_idx]
                if gt_val is None:
                    continue
                if isinstance(gt_val, (int, float)):
                    ok = num_close(a_val, gt_val, max(abs(gt_val) * 0.15, 1.0))
                else:
                    ok = str_match(a_val, gt_val)
                if not ok:
                    mismatched_field = (col_idx + 1, gt_val, a_val)
                    row_ok = False
                    break
            if row_ok:
                record(f"GT '{gt_sheet_name}' row[{idx+1}] key='{gt_row[key_col]}' values match", True)
            else:
                col_n, ev, av = mismatched_field
                record(f"GT '{gt_sheet_name}' row[{idx+1}] key='{gt_row[key_col]}' col {col_n}",
                       False, f"Expected {ev}, got {av}")
    gt_wb.close()


def check_gcal():
    print("\n=== Check 2: GCal supplier visit events on 2026-03-10 ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("""
        SELECT summary, start_datetime FROM gcal.events
        WHERE start_datetime::date = '2026-03-10'
        ORDER BY start_datetime
    """)
    events = cur.fetchall()
    cur.close()
    conn.close()

    record("GCal has >= 2 events on 2026-03-10", len(events) >= 2,
           f"Found {len(events)} events")

    summaries = [str(e[0]).lower() for e in events]
    has_supplier = any("supplier" in s or "visit" in s or "meeting" in s or "shanghai" in s or "guangzhou" in s
                       for s in summaries)
    record("GCal has supplier/visit/meeting events", has_supplier,
           f"Summaries: {summaries}")


def check_emails():
    print("\n=== Check 3: Emails sent ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
    messages = cur.fetchall()
    cur.close()
    conn.close()

    def to_addresses(to_addr):
        if isinstance(to_addr, list):
            return " ".join(str(r).lower() for r in to_addr)
        elif to_addr:
            try:
                parsed = json.loads(str(to_addr))
                return " ".join(str(r).lower() for r in parsed) if isinstance(parsed, list) else str(to_addr).lower()
            except Exception:
                return str(to_addr).lower()
        return ""

    # Count outgoing messages (not the injected supplier emails which are FROM suppliers)
    outgoing_addrs = ["shanghai_supplier@techworld.com", "gz_supplier@supplier.com", "procurement@company.com"]
    outgoing = [m for m in messages if any(addr in to_addresses(m[2]) for addr in outgoing_addrs)]

    record("At least 1 email sent to supplier or procurement", len(outgoing) >= 1,
           f"Total messages: {len(messages)}, matching: {len(outgoing)}")
    # Count only outgoing emails: TO supplier/procurement (i.e., agent's actions),
    # not the noise injected emails (which are FROM suppliers TO buyer).
    record("At least 2 outgoing emails sent (visits + summary)", len(outgoing) >= 2,
           f"Found {len(outgoing)} outgoing (matching to_addr); total: {len(messages)}")

    to_shanghai = [m for m in messages if "shanghai_supplier@techworld.com" in to_addresses(m[2])]
    record("Email sent to shanghai_supplier@techworld.com", len(to_shanghai) >= 1,
           f"Total messages: {len(messages)}")
    to_gz = [m for m in messages if "gz_supplier@supplier.com" in to_addresses(m[2])]
    record("Email sent to gz_supplier@supplier.com", len(to_gz) >= 1,
           f"Total messages: {len(messages)}")
    to_proc = [m for m in messages if "procurement@company.com" in to_addresses(m[2])]
    record("Email sent to procurement@company.com (summary)", len(to_proc) >= 1,
           f"Total messages: {len(messages)}")

    # Subject must mention meeting / visit / supplier on outgoing supplier emails (rejects empty/random subjects).
    meeting_kw = ("meeting", "visit", "supplier", "appointment", "trip", "march")
    supplier_outgoing = to_shanghai + to_gz
    if supplier_outgoing:
        ok_subj = sum(
            1 for m in supplier_outgoing
            if any(kw in (str(m[0] or "").lower()) for kw in meeting_kw)
        )
        record("Outgoing supplier emails reference meeting/visit in subject",
               ok_subj >= len(supplier_outgoing),
               f"{ok_subj}/{len(supplier_outgoing)} outgoing supplier emails have keyword in subject")


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_excel(args.agent_workspace, args.groundtruth_workspace)
    check_gcal()
    check_emails()

    total = PASS_COUNT + FAIL_COUNT
    if total == 0:
        print("\nFAIL: No checks were performed.")
        sys.exit(1)

    print(f"\nOverall: {PASS_COUNT}/{total} checks passed")

    result = {
        "total_passed": PASS_COUNT,
        "total_checks": total,
        "failed": FAIL_COUNT,
    }

    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    if FAIL_COUNT == 0:
        print("PASS")
        sys.exit(0)
    else:
        print(f"FAIL ({FAIL_COUNT} failures)")
        sys.exit(1)


if __name__ == "__main__":
    main()
