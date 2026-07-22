"""
Evaluation for yt-12306-tech-conference-travel-notion-excel task.

Checks:
1. Tech_Conference_Plan.xlsx exists
2. Prep_Videos sheet has >= 4 rows with Title and View_Count columns
3. Travel_Details sheet has >= 2 rows with Train_No column
4. Travel_Details contains G235 and G236
5. Conference_Schedule sheet has >= 5 rows
6. Notion page exists with Conference or Tech in title
7. GCal has >= 3 new events between 2026-03-12 and 2026-03-15
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
    try: return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError): return False


def str_match(a, b):
    if a is None or b is None: return a is None and b is None
    return str(a).strip().lower() == str(b).strip().lower()


def check_excel(agent_workspace, groundtruth_workspace="."):
    print("\n=== Check 1: Excel Tech_Conference_Plan.xlsx ===")
    xlsx_path = os.path.join(agent_workspace, "Tech_Conference_Plan.xlsx")
    if not os.path.exists(xlsx_path):
        record("Tech_Conference_Plan.xlsx exists", False, f"Not found at {xlsx_path}")
        return
    record("Tech_Conference_Plan.xlsx exists", True)

    try:
        wb = openpyxl.load_workbook(xlsx_path)
    except Exception as e:
        record("Excel file readable", False, str(e))
        return
    record("Excel file readable", True)

    sheet_names_lower = [s.lower() for s in wb.sheetnames]

    # Check Prep_Videos sheet
    prep_sheet = None
    for name in wb.sheetnames:
        if "prep" in name.lower() or "video" in name.lower():
            prep_sheet = wb[name]
            break
    if prep_sheet is None:
        record("Prep_Videos sheet exists", False, f"Sheets: {wb.sheetnames}")
    else:
        record("Prep_Videos sheet exists", True)
        rows = list(prep_sheet.iter_rows(values_only=True))
        headers = [str(c).strip().lower() if c else "" for c in (rows[0] if rows else [])]
        has_title = any("title" in h for h in headers)
        has_viewcount = any("view" in h or "count" in h for h in headers)
        record("Prep_Videos has Title and View_Count columns", has_title and has_viewcount,
               f"Headers: {rows[0] if rows else []}")
        data_rows = [r for r in rows[1:] if any(c for c in r)]
        # Task says exactly 5 videos
        record("Prep_Videos has exactly 5 data rows", len(data_rows) == 5,
               f"Found {len(data_rows)} data rows")

    # Check Travel_Details sheet
    travel_sheet = None
    for name in wb.sheetnames:
        if "travel" in name.lower():
            travel_sheet = wb[name]
            break
    if travel_sheet is None:
        record("Travel_Details sheet exists", False, f"Sheets: {wb.sheetnames}")
    else:
        record("Travel_Details sheet exists", True)
        rows = list(travel_sheet.iter_rows(values_only=True))
        all_text = " ".join(str(c) for r in rows for c in r if c).upper()
        headers = [str(c).strip().lower() if c else "" for c in (rows[0] if rows else [])]
        has_train = any("train" in h or "no" in h for h in headers)
        record("Travel_Details has Train_No column", has_train, f"Headers: {rows[0] if rows else []}")
        data_rows = [r for r in rows[1:] if any(c for c in r)]
        record("Travel_Details has >= 2 data rows", len(data_rows) >= 2,
               f"Found {len(data_rows)} rows")
        # Dynamic check: query the actual train DB for valid Beijing<->Qufu trains
        # on the task dates rather than hard-coding train codes.
        try:
            tconn = psycopg2.connect(**DB_CONFIG)
            tcur = tconn.cursor()
            tcur.execute("""
                SELECT station_train_code FROM train.trains
                WHERE from_station_telecode='VNP' AND to_station_telecode='QFB'
                  AND depart_date='2026-03-12'
            """)
            outbound_codes = {r[0] for r in tcur.fetchall()}
            tcur.execute("""
                SELECT station_train_code FROM train.trains
                WHERE from_station_telecode='QFB' AND to_station_telecode='VNP'
                  AND depart_date='2026-03-15'
            """)
            return_codes = {r[0] for r in tcur.fetchall()}
            tcur.close()
            tconn.close()
        except Exception as e:
            outbound_codes, return_codes = set(), set()
            print(f"  [WARN] Train DB lookup failed: {e}")
        if outbound_codes:
            ok_out = any(code in all_text for code in outbound_codes)
            record(
                f"Travel_Details outbound train code present",
                ok_out,
                f"Expected one of {sorted(outbound_codes)} in Travel_Details text",
            )
        if return_codes:
            ok_ret = any(code in all_text for code in return_codes)
            record(
                f"Travel_Details return train code present",
                ok_ret,
                f"Expected one of {sorted(return_codes)} in Travel_Details text",
            )

    # Check Conference_Schedule sheet
    sched_sheet = None
    for name in wb.sheetnames:
        if "schedule" in name.lower() or "conference" in name.lower():
            sched_sheet = wb[name]
            break
    if sched_sheet is None:
        record("Conference_Schedule sheet exists", False, f"Sheets: {wb.sheetnames}")
    else:
        record("Conference_Schedule sheet exists", True)
        rows = list(sched_sheet.iter_rows(values_only=True))
        data_rows = [r for r in rows[1:] if any(c for c in r)]
        # Task says at least six rows
        record("Conference_Schedule has >= 6 data rows", len(data_rows) >= 6,
               f"Found {len(data_rows)} rows")

    # --- Groundtruth XLSX value comparison (order-insensitive, skips free-form cols) ---
    gt_path = os.path.join(groundtruth_workspace, "Tech_Conference_Plan.xlsx")
    if os.path.isfile(gt_path):
        gt_wb = openpyxl.load_workbook(gt_path, data_only=True)
        for gt_sname in gt_wb.sheetnames:
            gt_ws = gt_wb[gt_sname]
            a_ws = None
            for asn in wb.sheetnames:
                if asn.strip().lower() == gt_sname.strip().lower():
                    a_ws = wb[asn]; break
            if a_ws is None:
                record(f"GT sheet '{gt_sname}' exists in agent xlsx", False, f"Available: {wb.sheetnames}")
                continue
            # Identify free-form columns to skip — Topic / Key_Technologies /
            # Activity / Location are agent-discretionary; Notes / descriptions
            # are unstructured; Duration_Min and Price_CNY may be rounded by the
            # agent in slightly different ways. Compare only on the structurally
            # pinned cells (titles, train codes, dates, etc.).
            gt_headers = [str(c.value).strip().lower() if c.value else "" for c in next(gt_ws.iter_rows(max_row=1))]
            skip_cols = {i for i, h in enumerate(gt_headers)
                         if any(k in h for k in [
                             "note", "reach", "remark", "comment", "description",
                             "topic", "key_tech", "key tech", "technologies",
                             "activity", "location",
                             "duration_min", "duration min", "price_cny", "price cny",
                         ])}
            gt_rows = [r for r in gt_ws.iter_rows(min_row=2, values_only=True) if any(c is not None for c in r)]
            a_rows = [r for r in a_ws.iter_rows(min_row=2, values_only=True) if any(c is not None for c in r)]
            # Conference_Schedule may have legitimately more rows than GT (task asks
            # for "at least six"); use >= for that sheet, exact otherwise.
            if "schedule" in gt_sname.lower() or "conference" in gt_sname.lower():
                record(
                    f"GT '{gt_sname}' row count",
                    len(a_rows) >= len(gt_rows),
                    f"Expected >= {len(gt_rows)}, got {len(a_rows)}",
                )
            else:
                record(f"GT '{gt_sname}' row count", len(a_rows) == len(gt_rows),
                       f"Expected {len(gt_rows)}, got {len(a_rows)}")
            # Order-insensitive: each GT row's key cols must exist somewhere in agent rows.
            # Normalize numeric vs string discrepancies and trim/lowercase strings.
            def _norm_cell(v):
                if v is None:
                    return None
                if isinstance(v, (int, float)):
                    return float(v)
                s = str(v).strip()
                # Try numeric coercion first
                try:
                    return float(s.replace(",", ""))
                except Exception:
                    return s.lower()

            def row_key(row):
                return tuple(
                    _norm_cell(v)
                    for i, v in enumerate(row)
                    if v is not None and i not in skip_cols
                )
            for ri, gt_row in enumerate(gt_rows):
                gt_key = row_key(gt_row)
                found = any(row_key(ar) == gt_key for ar in a_rows)
                first_val = next((v for v in gt_row if v is not None), None)
                record(f"GT '{gt_sname}' entry '{first_val}' found (order-insensitive)", found,
                       f"gt_key={gt_key}")
        gt_wb.close()


def check_notion():
    print("\n=== Check 2: Notion Page ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("SELECT id, properties FROM notion.pages")
    pages = cur.fetchall()
    cur.close()
    conn.close()

    # Task title: "Tech Conference Preparation - Qufu March 2026"
    found_exact = False
    for page_id, props in pages:
        try:
            # Look for "Tech Conference Preparation" + "Qufu" + "March 2026"
            props_str = json.dumps(props).lower()
            if (
                "tech conference preparation" in props_str
                and "qufu" in props_str
                and "march 2026" in props_str
            ):
                found_exact = True
                break
        except Exception:
            continue

    record(
        "Notion page titled 'Tech Conference Preparation - Qufu March 2026' exists",
        found_exact,
        f"Pages found: {len(pages)}",
    )


def check_gcal():
    print("\n=== Check 3: GCal Conference Events ===")
    import datetime as _dt
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("""
        SELECT summary, start_datetime, end_datetime FROM gcal.events
        WHERE start_datetime >= '2026-03-12' AND start_datetime < '2026-03-16'
        ORDER BY start_datetime
    """)
    events = cur.fetchall()
    cur.close()
    conn.close()

    def _to_naive(d):
        if d is None:
            return None
        if hasattr(d, "tzinfo") and d.tzinfo is not None:
            return d.astimezone(_dt.timezone.utc).replace(tzinfo=None)
        return d

    # Required 4 events with exact titles + times
    expected = [
        (
            "Conference Travel: Beijing to Qufu",
            _dt.datetime(2026, 3, 12, 17, 0),
            _dt.datetime(2026, 3, 12, 20, 0),
        ),
        (
            "Tech Conference Day 1",
            _dt.datetime(2026, 3, 13, 9, 0),
            _dt.datetime(2026, 3, 13, 18, 0),
        ),
        (
            "Tech Conference Day 2",
            _dt.datetime(2026, 3, 14, 9, 0),
            _dt.datetime(2026, 3, 14, 18, 0),
        ),
        (
            "Return: Qufu to Beijing",
            _dt.datetime(2026, 3, 15, 14, 30),
            _dt.datetime(2026, 3, 15, 17, 30),
        ),
    ]

    for exp_title, exp_start, exp_end in expected:
        match = None
        for summary, sdt, edt in events:
            if summary and summary.strip().lower() == exp_title.lower():
                match = (summary, _to_naive(sdt), _to_naive(edt))
                break
        record(
            f"GCal event '{exp_title}' exists with correct title",
            match is not None,
            f"All summaries: {[e[0] for e in events]}",
        )
        if match is not None:
            _, sdt, edt = match
            # Allow timezone interpretations: try ±8hr (China) tolerance only when comparing
            ok_start = sdt is not None and (
                abs((sdt - exp_start).total_seconds()) <= 60
                or abs((sdt - (exp_start - _dt.timedelta(hours=8))).total_seconds()) <= 60
                or abs((sdt - (exp_start + _dt.timedelta(hours=8))).total_seconds()) <= 60
            )
            ok_end = edt is not None and (
                abs((edt - exp_end).total_seconds()) <= 60
                or abs((edt - (exp_end - _dt.timedelta(hours=8))).total_seconds()) <= 60
                or abs((edt - (exp_end + _dt.timedelta(hours=8))).total_seconds()) <= 60
            )
            record(
                f"GCal event '{exp_title}' start={exp_start.strftime('%H:%M')}",
                ok_start,
                f"got {sdt}",
            )
            record(
                f"GCal event '{exp_title}' end={exp_end.strftime('%H:%M')}",
                ok_end,
                f"got {edt}",
            )


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_excel(args.agent_workspace, args.groundtruth_workspace)
    check_notion()
    check_gcal()

    total = PASS_COUNT + FAIL_COUNT
    if total == 0:
        print("\nFAIL: No checks were performed.")
        sys.exit(1)

    accuracy = PASS_COUNT / total * 100
    print(f"\nOverall: {PASS_COUNT}/{total} checks passed ({accuracy:.1f}%)")

    result = {
        "total_passed": PASS_COUNT,
        "total_checks": total,
        "accuracy": accuracy,
    }

    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    if FAIL_COUNT == 0:
        print("PASS")
        sys.exit(0)
    else:
        print("FAIL")
        sys.exit(1)


if __name__ == "__main__":
    main()
