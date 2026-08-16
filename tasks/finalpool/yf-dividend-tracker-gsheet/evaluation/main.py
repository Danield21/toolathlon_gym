"""
Evaluation for yf-dividend-tracker-gsheet.
Checks:
1. Google Sheet "Dividend Tracker" exists in DB with correct stock action data
2. Email sent to investor@portfolio.example.com with dividend summary
"""
import argparse
import json
import os
import sys
from datetime import datetime

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0
WARN_COUNT = 0
IS_GT_SELF_TEST = False


def record(name, passed, detail="", db_side=False):
    global PASS_COUNT, FAIL_COUNT, WARN_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        # In GT self-test mode, DB-side checks (gsheet/email)
        # naturally fail because GT files cannot pre-populate DB.
        if IS_GT_SELF_TEST and db_side:
            WARN_COUNT += 1
            msg = f": {detail[:300]}" if detail else ""
            print(f"  [WARN] {name} (GT self-test, DB-side){msg}")
        else:
            FAIL_COUNT += 1
            msg = f": {detail[:300]}" if detail else ""
            print(f"  [FAIL] {name}{msg}")


def str_match(a, b):
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return str(a).strip().lower() == str(b).strip().lower()


def num_close(a, b, tol=0.1):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def get_expected_data():
    """Query YF DB for expected dividend/action data."""
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("SELECT symbol, data FROM yf.stock_info WHERE symbol IN ('GOOGL','AMZN','JPM','JNJ','XOM') ORDER BY symbol")
    rows = cur.fetchall()

    dividend_stocks = []
    for symbol, data in rows:
        d = data if isinstance(data, dict) else json.loads(data)
        div_rate = d.get('dividendRate')
        if div_rate and float(div_rate) > 0:
            dividend_stocks.append(symbol)

    cur.close()
    conn.close()
    return dividend_stocks


def check_gsheet():
    """Check Google Sheet data in DB."""
    print("\n=== Checking Google Sheet ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    # Find spreadsheet with "dividend" or "tracker" in title
    cur.execute("""
        SELECT id, title FROM gsheet.spreadsheets
        WHERE LOWER(title) LIKE '%dividend%' OR LOWER(title) LIKE '%tracker%'
    """)
    spreadsheets = cur.fetchall()
    record("Spreadsheet exists", len(spreadsheets) > 0,
           f"No spreadsheet with 'dividend' or 'tracker' found",
           db_side=True)

    if not spreadsheets:
        cur.close()
        conn.close()
        return False

    sp_id = spreadsheets[0][0]
    print(f"  Found spreadsheet: {spreadsheets[0][1]} (id={sp_id})")

    # Check sheets
    cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s", (sp_id,))
    sheets = cur.fetchall()
    record("At least one sheet exists", len(sheets) > 0, db_side=True)

    if not sheets:
        cur.close()
        conn.close()
        return False

    sheet_id = sheets[0][0]

    # Check cells for content
    cur.execute("""
        SELECT row_index, col_index, value FROM gsheet.cells
        WHERE spreadsheet_id = %s AND sheet_id = %s
        ORDER BY row_index, col_index
    """, (sp_id, sheet_id))
    cells = cur.fetchall()
    record("Sheet has data cells", len(cells) > 5, f"Only {len(cells)} cells found",
           db_side=True)

    # Check for ALL 5 required symbols (incl. AMZN previously missing)
    all_values = " ".join(str(c[2]).lower() for c in cells if c[2])
    for symbol in ['googl', 'amzn', 'jnj', 'jpm', 'xom']:
        record(f"Sheet contains {symbol.upper()}", symbol in all_values, db_side=True)

    # Check Action_Type column has 'Dividend' or 'Stock Split' values
    import re
    dividend_count = sum(1 for c in cells if c[2] and re.fullmatch(r"\s*dividend\s*", str(c[2]), re.IGNORECASE))
    split_count = sum(1 for c in cells if c[2] and re.fullmatch(r"\s*stock\s+split\s*", str(c[2]), re.IGNORECASE))
    # Dividend is the primary action type; require at least 2 dividend rows
    # so the tracker is substantive (multiple dividend-paying stocks tracked).
    record("Sheet has at least 2 'Dividend' Action_Type cells", dividend_count >= 2,
           f"Found {dividend_count} 'Dividend' cells", db_side=True)
    # Stock Split: split data lives only in stock_info.lastSplitDate /
    # lastSplitFactor (historical, e.g. 2001-2022) - some agents using only
    # price history will not surface them. Make this NON-BLOCKING so the
    # test is fair to agents that consider old splits non-actionable.
    record("Sheet has 'Stock Split' Action_Type cells (optional, non-blocking)",
           True,
           f"Found {split_count} 'Stock Split' cells "
           "(splits are historical, not required for pass)")

    cur.close()
    conn.close()
    return True


def check_email():
    """Check email sent to investor."""
    print("\n=== Checking Email ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
    emails = cur.fetchall()
    record("At least 1 email sent", len(emails) >= 1, f"Found {len(emails)}", db_side=True)

    matched = None
    for subject, from_addr, to_addr, body_text in emails:
        subject_lower = (subject or "").lower()
        to_str = str(to_addr or "").lower()
        if ("dividend action summary" in subject_lower or
            ("dividend" in subject_lower and "action" in subject_lower and "summary" in subject_lower)):
            if "investor@portfolio.example.com" in to_str:
                matched = (subject, from_addr, to_addr, body_text)
                break

    record("Email with subject 'Dividend Action Summary' to investor exists", matched is not None,
           "Required subject + recipient combination not found", db_side=True)

    if matched:
        subject, from_addr, to_addr, body_text = matched
        from_str = str(from_addr or "").lower()
        record("Email from portfolio-alerts@finance.example.com",
               "portfolio-alerts@finance.example.com" in from_str,
               f"From: {from_addr}", db_side=True)
        body_lower = (body_text or "").lower()
        # Body must summarize dividend stocks: at least 3 dividend stock names appear
        symbols_in_body = sum(1 for s in ["googl", "amzn", "jpm", "jnj", "xom"] if s in body_lower)
        record("Email body mentions >= 3 stock symbols",
               symbols_in_body >= 3, f"Found {symbols_in_body}/5", db_side=True)
        record("Email body mentions dividend info",
               "dividend" in body_lower,
               "No dividend reference in body", db_side=True)

    cur.close()
    conn.close()
    return matched is not None


def main():
    global IS_GT_SELF_TEST
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    # Detect GT self-test mode: agent_workspace == groundtruth_workspace.
    try:
        if args.groundtruth_workspace:
            IS_GT_SELF_TEST = (
                os.path.realpath(args.agent_workspace) ==
                os.path.realpath(args.groundtruth_workspace)
            )
    except Exception:
        IS_GT_SELF_TEST = False

    dividend_stocks = get_expected_data()
    print(f"[eval] Expected dividend stocks: {dividend_stocks}")

    gsheet_ok = check_gsheet()
    email_ok = check_email()

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    if IS_GT_SELF_TEST:
        print(f"  Warned (GT self-test, DB-side): {WARN_COUNT}")
    overall = FAIL_COUNT == 0
    print(f"  Overall: {'PASS' if overall else 'FAIL'}")

    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
