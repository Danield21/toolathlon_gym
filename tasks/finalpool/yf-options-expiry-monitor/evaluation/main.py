"""Evaluation for yf-options-expiry-monitor."""
import argparse
import json
import os
import sys
from datetime import date

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="toolathlon_gym", user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        detail_str = f": {str(detail)[:200]}" if detail else ""
        print(f"  [FAIL] {name}{detail_str}")


def num_close(a, b, tol=1.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return str(a).strip().lower() == str(b).strip().lower()


def load_sheet_rows(wb, sheet_name):
    for name in wb.sheetnames:
        if name.strip().lower() == sheet_name.strip().lower():
            return [[cell.value for cell in row] for row in wb[name].iter_rows()]
    return None


def date_str(v):
    if v is None:
        return ""
    if hasattr(v, "isoformat"):
        return v.isoformat()[:10]
    return str(v).strip()[:10]


def check_excel(agent_workspace, gt_workspace):
    print("\n=== Checking Excel ===")
    import openpyxl
    path = os.path.join(agent_workspace, "Options_Monitor.xlsx")
    gt_path = os.path.join(gt_workspace, "Options_Monitor.xlsx")
    check("Options_Monitor.xlsx exists", os.path.exists(path), f"Expected {path}")
    if not os.path.exists(path):
        return
    if not os.path.exists(gt_path):
        check("GT Options_Monitor.xlsx exists", False)
        return

    a_wb = openpyxl.load_workbook(path, data_only=True)
    g_wb = openpyxl.load_workbook(gt_path, data_only=True)

    # Position Analysis
    a_rows = load_sheet_rows(a_wb, "Position Analysis")
    g_rows = load_sheet_rows(g_wb, "Position Analysis")
    check("Sheet 'Position Analysis' present", a_rows is not None)
    if a_rows is not None and g_rows is not None:
        a_data = [r for r in a_rows[1:] if r and r[0] is not None]
        g_data = [r for r in g_rows[1:] if r and r[0] is not None]
        check(f"Position Analysis row count == {len(g_data)}",
              len(a_data) == len(g_data), f"got {len(a_data)}")

        # Sort: Symbol then Expiration
        keys = [(str(r[0]).upper(), date_str(r[1])) for r in a_data]
        check("Position Analysis sorted by Symbol then Expiration",
              keys == sorted(keys), f"got {keys[:3]}...")

        # Build (symbol, expiration, type) -> row mapping
        a_lookup = {(str(r[0]).strip().upper(), date_str(r[1]), str(r[2]).strip().lower()): r for r in a_data}
        for g_row in g_data:
            sym = str(g_row[0]).strip().upper()
            exp = date_str(g_row[1])
            typ = str(g_row[2]).strip().lower()
            key = (sym, exp, typ)
            a_row = a_lookup.get(key)
            check(f"Row ({sym}, {exp}, {typ}) present", a_row is not None)
            if a_row is None:
                continue
            # Num_Contracts (exact)
            check(f"  ({sym}, {exp}, {typ}).Num_Contracts == {g_row[3]}",
                  num_close(a_row[3], g_row[3], 0), f"got {a_row[3]}")
            # Avg_IV (±0.5)
            check(f"  ({sym}, {exp}, {typ}).Avg_IV ≈ {g_row[6]}",
                  num_close(a_row[6], g_row[6], 0.5), f"got {a_row[6]}")
            # Risk_Level (exact)
            check(f"  ({sym}, {exp}, {typ}).Risk_Level == '{g_row[7]}'",
                  str(a_row[7] or "").strip().lower() == str(g_row[7] or "").strip().lower(),
                  f"got '{a_row[7]}'")

    # Expiry Alerts
    a_rows = load_sheet_rows(a_wb, "Expiry Alerts")
    g_rows = load_sheet_rows(g_wb, "Expiry Alerts")
    check("Sheet 'Expiry Alerts' present", a_rows is not None)
    if a_rows is not None and g_rows is not None:
        a_data = [r for r in a_rows[1:] if r and r[0] is not None]
        g_data = [r for r in g_rows[1:] if r and r[0] is not None]
        check(f"Expiry Alerts row count == {len(g_data)}",
              len(a_data) == len(g_data), f"got {len(a_data)}")
        # Sort by Days_Until_Expiry ascending
        days = []
        for r in a_data:
            try:
                days.append(int(float(r[4])))
            except (TypeError, ValueError):
                pass
        check("Expiry Alerts sorted by Days_Until_Expiry ascending",
              days == sorted(days), f"got {days}")
        # Verify each expected row
        a_lookup = {(str(r[0]).strip().upper(), date_str(r[1]), str(r[2]).strip().lower()): r for r in a_data}
        for g_row in g_data:
            sym = str(g_row[0]).strip().upper()
            exp = date_str(g_row[1])
            typ = str(g_row[2]).strip().lower()
            a_row = a_lookup.get((sym, exp, typ))
            check(f"Expiry Alerts row ({sym}, {exp}, {typ}) present",
                  a_row is not None)
            if a_row is None:
                continue
            check(f"  Days_Until_Expiry == {g_row[4]}",
                  num_close(a_row[4], g_row[4], 0), f"got {a_row[4]}")

    # Summary
    a_rows = load_sheet_rows(a_wb, "Summary")
    g_rows = load_sheet_rows(g_wb, "Summary")
    check("Sheet 'Summary' present", a_rows is not None)
    if a_rows is not None and g_rows is not None:
        a_data = [r for r in a_rows[1:] if r and r[0] is not None]
        g_data = [r for r in g_rows[1:] if r and r[0] is not None]
        a_lookup = {str(r[0]).strip().lower(): r[1] for r in a_data}
        for g_row in g_data:
            metric = str(g_row[0]).strip().lower()
            val = a_lookup.get(metric)
            check(f"Summary metric '{metric}' present", val is not None)
            if val is None:
                continue
            if metric == "stocks_with_near_expiry":
                # Sort comma-separated symbols
                got_syms = sorted([s.strip().upper() for s in str(val).split(",") if s.strip()])
                exp_syms = sorted([s.strip().upper() for s in str(g_row[1]).split(",") if s.strip()])
                check(f"  Summary.{metric} == {exp_syms}",
                      got_syms == exp_syms, f"got {got_syms}")
            elif isinstance(g_row[1], (int, float)):
                check(f"  Summary.{metric} == {g_row[1]}",
                      num_close(val, g_row[1], 0), f"got {val}")
            else:
                check(f"  Summary.{metric} == '{g_row[1]}'",
                      str(val).strip().lower() == str(g_row[1]).strip().lower(),
                      f"got '{val}'")


def check_gcal(gt_workspace):
    print("\n=== Checking Calendar Events ===")
    import openpyxl
    gt_path = os.path.join(gt_workspace, "Options_Monitor.xlsx")
    g_wb = openpyxl.load_workbook(gt_path, data_only=True)
    # Use Expiry Alerts to determine which (symbol, expiration) need events
    g_alerts = load_sheet_rows(g_wb, "Expiry Alerts")
    expected_events = set()
    if g_alerts:
        for r in g_alerts[1:]:
            if r and r[0]:
                expected_events.add((str(r[0]).strip().upper(), date_str(r[1])))

    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
    except psycopg2.Error as e:
        check("DB connect", False, str(e))
        return

    cur.execute("SELECT summary, description, start_datetime FROM gcal.events WHERE summary ILIKE '%%options expiry%%'")
    events = cur.fetchall()
    check("At least 1 options expiry calendar event", len(events) >= 1, f"got {len(events)}")

    # Collect (symbol, date) tuples from events
    actual_events = set()
    for sm, desc, start in events:
        sm_upper = (sm or "").upper()
        date_part = start.date() if hasattr(start, "date") else start
        # Extract symbol from summary
        for sym in ("AMZN", "GOOGL", "JNJ", "JPM", "XOM"):
            if sym in sm_upper:
                actual_events.add((sym, date_part.isoformat() if hasattr(date_part, "isoformat") else str(date_part)))

    # Each unique (symbol, expiration) from expected_events should have an event
    expected_syms_dates = set()
    for s, e in expected_events:
        expected_syms_dates.add((s, e))

    for sym, exp in sorted(expected_syms_dates):
        check(f"Calendar event for {sym} on {exp} present",
              (sym, exp) in actual_events, f"actual events: {actual_events}")

    cur.close()
    conn.close()


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

    check_excel(agent_ws, gt_ws)
    check_gcal(gt_ws)

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    overall = FAIL_COUNT == 0
    print(f"  Overall: {'PASS' if overall else 'FAIL'}")
    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
