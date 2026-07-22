"""Evaluation for wc-customer-order-gsheet-email task."""

import argparse
import json
import os
import sys

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
        detail_str = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{detail_str}")


def num_close(a, b, tol=0.5):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def get_top10():
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT first_name, last_name, first_name || ' ' || last_name as name, email, orders_count,
               ROUND(total_spent::numeric, 2) as total_spent,
               CASE WHEN orders_count > 0 THEN ROUND((total_spent::numeric / orders_count), 2) ELSE 0 END as avg_order
        FROM wc.customers ORDER BY total_spent::numeric DESC LIMIT 10
    """)
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return rows


def get_summary():
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*), ROUND(SUM(total_spent::numeric), 2) FROM wc.customers")
    total_cust, total_rev = cur.fetchone()
    cur.execute("""
        SELECT ROUND(SUM(total_spent::numeric), 2)
        FROM (SELECT total_spent FROM wc.customers ORDER BY total_spent::numeric DESC LIMIT 10) t
    """)
    vip_rev = cur.fetchone()[0]
    vip_pct = round(float(vip_rev) / float(total_rev) * 100, 1) if float(total_rev) > 0 else 0
    avg_spend = round(float(total_rev) / total_cust, 2) if total_cust > 0 else 0
    cur.close()
    conn.close()
    return {
        "Total_Customers": total_cust,
        "Total_Revenue": float(total_rev),
        "VIP_Revenue_Share_Pct": vip_pct,
        "Avg_Customer_Spend": avg_spend,
    }


def get_grid(cur, ss_id, sheet_id):
    cur.execute("""
        SELECT row_index, col_index, value FROM gsheet.cells
        WHERE spreadsheet_id = %s AND sheet_id = %s
        ORDER BY row_index, col_index
    """, (ss_id, sheet_id))
    grid = {}
    for row_idx, col_idx, value in cur.fetchall():
        grid.setdefault(row_idx, {})[col_idx] = value
    return grid


def safe_float(s):
    if s is None:
        return None
    try:
        return float(str(s).strip())
    except (TypeError, ValueError):
        return None


def check_gsheet_top_customers():
    print("\n=== Checking Google Sheet: Top Customers ===")
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
    except psycopg2.Error as e:
        check("DB connect", False, str(e))
        return

    cur.execute("SELECT id, title FROM gsheet.spreadsheets WHERE LOWER(title) LIKE '%%vip%%' OR LOWER(title) LIKE '%%vip customer tracker%%'")
    spreadsheets = cur.fetchall()
    check("Spreadsheet 'VIP Customer Tracker' exists", len(spreadsheets) >= 1, f"Found {len(spreadsheets)}")
    if not spreadsheets:
        cur.close()
        conn.close()
        return
    ss_id = spreadsheets[0][0]

    cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s", (ss_id,))
    sheets = cur.fetchall()
    top_sheet_id = next((s_id for s_id, t in sheets if t and t.strip().lower() == "top customers"), None)
    check("Sheet 'Top Customers' present", top_sheet_id is not None, f"got {[t for _, t in sheets]}")
    if top_sheet_id is None:
        cur.close()
        conn.close()
        return

    grid = get_grid(cur, ss_id, top_sheet_id)
    # Find header row
    header_row = None
    for r in sorted(grid.keys()):
        joined = " ".join(str(v).lower() for v in grid[r].values())
        if "name" in joined and "email" in joined and "total" in joined:
            header_row = r
            break
    check("Top Customers header row found", header_row is not None)
    if header_row is None:
        cur.close()
        conn.close()
        return

    col_map = {}
    for c, v in grid[header_row].items():
        vl = str(v or "").strip().lower()
        if vl == "name":
            col_map["name"] = c
        elif vl == "email":
            col_map["email"] = c
        elif vl == "orders_count":
            col_map["orders"] = c
        elif vl == "total_spent":
            col_map["total"] = c
        elif vl == "avg_order_value":
            col_map["aov"] = c

    data_rows = [grid[r] for r in sorted(grid.keys()) if r > header_row]
    check("Top Customers has 10 data rows", len(data_rows) == 10, f"got {len(data_rows)}")

    # Sort order: total_spent descending
    totals = []
    for row in data_rows:
        v = row.get(col_map.get("total", -1))
        f = safe_float(v)
        totals.append(f)
    check("Top Customers sorted by Total_Spent descending",
          totals == sorted(totals, reverse=True, key=lambda x: x or 0),
          f"got {totals}")

    top10 = get_top10()
    # Email -> position lookup
    expected_by_email = {row[3].lower(): row for row in top10}
    for i, row in enumerate(data_rows):
        email_v = str(row.get(col_map.get("email", -1), "")).strip().lower()
        exp = expected_by_email.get(email_v)
        check(f"Top Customers row {i+1} email '{email_v}' is in expected top10",
              exp is not None, f"got {email_v}")
        if exp is None:
            continue
        # Check name
        name_v = str(row.get(col_map.get("name", -1), "")).strip().lower()
        check(f"  Row for {email_v}.Name == '{exp[2]}'",
              name_v == exp[2].strip().lower(), f"got '{name_v}'")
        # Check orders_count
        orders_v = safe_float(row.get(col_map.get("orders", -1)))
        check(f"  Row for {email_v}.Orders_Count == {exp[4]}",
              orders_v is not None and int(orders_v) == int(exp[4]), f"got {orders_v}")
        # Check total_spent
        total_v = safe_float(row.get(col_map.get("total", -1)))
        check(f"  Row for {email_v}.Total_Spent ≈ {exp[5]}",
              total_v is not None and abs(total_v - float(exp[5])) <= 0.5, f"got {total_v}")
        # Check avg_order_value
        aov_v = safe_float(row.get(col_map.get("aov", -1)))
        check(f"  Row for {email_v}.Avg_Order_Value ≈ {exp[6]}",
              aov_v is not None and abs(aov_v - float(exp[6])) <= 0.05, f"got {aov_v}")

    cur.close()
    conn.close()


def check_gsheet_summary():
    print("\n=== Checking Google Sheet: Store Summary ===")
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
    except psycopg2.Error as e:
        check("DB connect", False, str(e))
        return

    cur.execute("SELECT id FROM gsheet.spreadsheets WHERE LOWER(title) LIKE '%%vip%%'")
    spreadsheets = cur.fetchall()
    if not spreadsheets:
        check("Spreadsheet 'VIP' present", False)
        cur.close()
        conn.close()
        return
    ss_id = spreadsheets[0][0]
    cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id=%s", (ss_id,))
    sheets = cur.fetchall()
    sum_sheet_id = next((s_id for s_id, t in sheets if t and t.strip().lower() == "store summary"), None)
    check("Sheet 'Store Summary' present", sum_sheet_id is not None, f"got {[t for _, t in sheets]}")
    if sum_sheet_id is None:
        cur.close()
        conn.close()
        return

    grid = get_grid(cur, ss_id, sum_sheet_id)
    # Build (metric_name -> value) by reading rows where col 0 = metric, col 1 = value
    kv = {}
    for r in sorted(grid.keys()):
        row = grid[r]
        cols = sorted(row.keys())
        if len(cols) < 2:
            continue
        metric = str(row[cols[0]] or "").strip().lower()
        val = row[cols[1]]
        if metric and metric != "metric":
            kv[metric] = val

    expected = get_summary()
    check("Summary.Total_Customers == %d" % expected["Total_Customers"],
          int(safe_float(kv.get("total_customers")) or -1) == expected["Total_Customers"],
          f"got {kv.get('total_customers')}")
    check("Summary.Total_Revenue ≈ %s" % expected["Total_Revenue"],
          safe_float(kv.get("total_revenue")) is not None and
          abs(safe_float(kv.get("total_revenue")) - expected["Total_Revenue"]) <= 1.0,
          f"got {kv.get('total_revenue')}")
    check("Summary.VIP_Revenue_Share_Pct ≈ %s" % expected["VIP_Revenue_Share_Pct"],
          safe_float(kv.get("vip_revenue_share_pct")) is not None and
          abs(safe_float(kv.get("vip_revenue_share_pct")) - expected["VIP_Revenue_Share_Pct"]) <= 0.1,
          f"got {kv.get('vip_revenue_share_pct')}")
    check("Summary.Avg_Customer_Spend ≈ %s" % expected["Avg_Customer_Spend"],
          safe_float(kv.get("avg_customer_spend")) is not None and
          abs(safe_float(kv.get("avg_customer_spend")) - expected["Avg_Customer_Spend"]) <= 1.0,
          f"got {kv.get('avg_customer_spend')}")

    cur.close()
    conn.close()


def check_emails():
    print("\n=== Checking Emails ===")
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
    except psycopg2.Error as e:
        check("DB connect", False, str(e))
        return

    top10 = get_top10()
    top5 = top10[:5]

    cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
    all_emails = cur.fetchall()

    def parse_recipients(to_addr):
        if to_addr is None:
            return []
        if isinstance(to_addr, list):
            return [str(r).strip().lower() for r in to_addr]
        if isinstance(to_addr, str):
            try:
                parsed = json.loads(to_addr)
                if isinstance(parsed, list):
                    return [str(r).strip().lower() for r in parsed]
                return [str(to_addr).strip().lower()]
            except (json.JSONDecodeError, TypeError):
                return [str(to_addr).strip().lower()]
        return [str(to_addr).strip().lower()]

    for first_name, last_name, name, email_addr, orders, total, aov in top5:
        match = None
        for subj, frm, to, body in all_emails:
            if email_addr.lower() in parse_recipients(to):
                match = (subj, frm, to, body)
                break
        check(f"Email sent to top customer {email_addr}", match is not None)
        if match is None:
            continue
        subj, frm, to, body = match
        check(f"  Email to {email_addr}: subject == 'Thank You for Being a VIP Customer!'",
              (subj or "").strip().lower() == "thank you for being a vip customer!",
              f"got: {subj}")
        check(f"  Email to {email_addr}: from_addr == 'vip-program@store.example.com'",
              str(frm or "").strip().lower() == "vip-program@store.example.com",
              f"got: {frm}")
        check(f"  Email to {email_addr}: body addresses by first name '{first_name}'",
              first_name.lower() in (body or "").lower(),
              f"missing {first_name}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=" * 70)
    print("WC CUSTOMER ORDER GSHEET EMAIL - EVALUATION")
    print("=" * 70)

    check_gsheet_top_customers()
    check_gsheet_summary()
    check_emails()

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    all_ok = FAIL_COUNT == 0
    print(f"  Overall: {'PASS' if all_ok else 'FAIL'}")
    sys.exit(0 if all_ok else 1)


if __name__ == "__main__":
    main()
