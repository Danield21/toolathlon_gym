"""Evaluation for wc-vip-customer-gsheet-gcal-email.

Checks:
1. Google Sheet "VIP Customer Tracker" with >=10 rows including Gold/Silver/Bronze tiers
2. GCal event "VIP Appreciation Day" ~30 days from launch_time
3. 3 emails sent to Gold tier customers (Scarlett Wright, Ethan Martinez, Olivia Wilson)
   from vip@store.example.com with subject containing VIP and discount/20%
"""
import argparse
import json
import os
import sys
from datetime import datetime, timedelta, timezone

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=int(os.environ.get("PGPORT", "5432")), dbname="toolathlon_gym", user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0

# Top 10 customers by completed-order spend, derived live from WC DB at
# verifier startup (with hardcoded fallback). WC is read-only, so this
# query produces the same result run-to-run; the live query future-proofs
# the check against any data refresh.
def _fetch_top10_customers():
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
        cur.execute(
            """SELECT c.first_name, c.last_name, c.email,
                       ROUND(SUM(o.total::numeric), 2) AS total_spent
                FROM wc.customers c
                JOIN wc.orders o ON o.customer_id = c.id
                WHERE o.status = 'completed'
                GROUP BY c.id, c.first_name, c.last_name, c.email
                ORDER BY total_spent DESC
                LIMIT 10"""
        )
        rows = cur.fetchall()
        cur.close()
        conn.close()
        if rows and len(rows) >= 10:
            return [
                {"name": f"{r[0]} {r[1]}", "email": r[2].lower(),
                 "total_spent": float(r[3])}
                for r in rows
            ]
    except Exception:
        pass
    return None


_TOP10_FALLBACK = [
    {"name": "Scarlett Wright", "email": "scarlett.wright@x.dummyjson.com", "total_spent": 3328.27},
    {"name": "Ethan Martinez",  "email": "ethan.martinez@x.dummyjson.com",  "total_spent": 3053.80},
    {"name": "Olivia Wilson",   "email": "olivia.wilson@x.dummyjson.com",   "total_spent": 2942.96},
    {"name": "William Gonzalez", "email": "william.gonzalez@x.dummyjson.com", "total_spent": 2423.81},
    {"name": "Charlotte Lopez", "email": "charlotte.lopez@x.dummyjson.com", "total_spent": 2359.28},
    {"name": "Sophia Brown",    "email": "sophia.brown@x.dummyjson.com",    "total_spent": 1620.99},
    {"name": "Carter Baker",    "email": "carter.baker@x.dummyjson.com",    "total_spent": 1482.45},
    {"name": "Mason Parker",    "email": "mason.parker@x.dummyjson.com",    "total_spent": 1413.99},
    {"name": "Henry Hill",      "email": "henry.hill@x.dummyjson.com",      "total_spent": 1380.71},
    {"name": "Daniel Cook",     "email": "daniel.cook@x.dummyjson.com",     "total_spent": 1180.86},
]
TOP10 = _fetch_top10_customers() or _TOP10_FALLBACK
GOLD_CUSTOMERS = TOP10[:3]
SILVER_CUSTOMERS = TOP10[3:7]
BRONZE_CUSTOMERS = TOP10[7:10]
TOP_SPENT = TOP10[0]["total_spent"]


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        detail_str = f": {str(detail)[:200]}" if detail else ""
        print(f"  [FAIL] {name}{detail_str}")


def parse_recipients(to_addr):
    if to_addr is None:
        return []
    if isinstance(to_addr, list):
        return [str(r).strip().lower() for r in to_addr]
    to_str = str(to_addr).strip()
    try:
        parsed = json.loads(to_str)
        if isinstance(parsed, list):
            return [str(r).strip().lower() for r in parsed]
        return [to_str.lower()]
    except (json.JSONDecodeError, TypeError):
        return [to_str.lower()]


def check_gsheet():
    print("\n=== Checking Google Sheet ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("""
        SELECT id, title FROM gsheet.spreadsheets
        WHERE LOWER(title) LIKE '%vip%' AND LOWER(title) LIKE '%customer%'
    """)
    sheets = cur.fetchall()
    check("VIP Customer Tracker spreadsheet exists", len(sheets) >= 1,
          f"Found {len(sheets)} matching spreadsheets")

    if sheets:
        ss_id = sheets[0][0]
        # Read by row/col into matrix
        cur.execute("""
            SELECT row_index, col_index, value FROM gsheet.cells
            WHERE spreadsheet_id = %s ORDER BY row_index, col_index
        """, (ss_id,))
        cell_rows = cur.fetchall()
        grid = {}
        for r, c, v in cell_rows:
            grid.setdefault(r, {})[c] = v

        if grid:
            min_r = min(grid.keys())
            headers = [str(grid[min_r].get(i, "")).strip().lower() for i in sorted(grid[min_r].keys())]
            data_rows = sorted([r for r in grid if r > min_r])

            # Word-boundary tier counting (exact match within cell)
            import re
            def count_exact(token, rows_set):
                cnt = 0
                for r in rows_set:
                    for c in grid[r].values():
                        if c is None: continue
                        if re.fullmatch(rf"\s*{token}\s*", str(c), flags=re.IGNORECASE):
                            cnt += 1
                cnt_breakdown = cnt
                return cnt_breakdown

            gold_n = count_exact("gold", data_rows)
            silver_n = count_exact("silver", data_rows)
            bronze_n = count_exact("bronze", data_rows)

            check("GSheet has exactly 10 data rows",
                  len(data_rows) == 10, f"Got {len(data_rows)}")
            check("GSheet has 3 Gold tier rows", gold_n == 3, f"Got {gold_n}")
            check("GSheet has 4 Silver tier rows", silver_n == 4, f"Got {silver_n}")
            check("GSheet has 3 Bronze tier rows", bronze_n == 3, f"Got {bronze_n}")

            # Find tier and customer name columns
            tier_col = next((i for i, h in enumerate(headers) if "tier" in h or "vip" in h), None)
            name_col = next((i for i, h in enumerate(headers) if "customer" in h or "name" in h), None)
            spent_col = next((i for i, h in enumerate(headers) if "spent" in h or "total" in h and "spent" in h), None)

            if tier_col is not None and name_col is not None:
                # Verify EACH expected customer is tagged with the right tier
                # (per-rank mapping, not just gold count). Per QA: previous
                # version only checked Gold names + tier counts, not whether
                # silver/bronze rows mapped to the right ranks.
                row_lookup = {}
                for r in data_rows:
                    name = str(grid[r].get(name_col, "")).strip().lower()
                    tier = str(grid[r].get(tier_col, "")).strip().lower()
                    row_lookup[name] = tier

                def _find_tier(cust_name):
                    parts = cust_name.lower().split()
                    if not parts:
                        return None
                    first, last = parts[0], parts[-1]
                    for k, v in row_lookup.items():
                        if first in k and last in k:
                            return v
                    return None

                for cust in GOLD_CUSTOMERS:
                    found = _find_tier(cust["name"])
                    check(f"Customer '{cust['name']}' tagged Gold tier",
                          found == "gold", f"got '{found}'")
                for cust in SILVER_CUSTOMERS:
                    found = _find_tier(cust["name"])
                    check(f"Customer '{cust['name']}' tagged Silver tier",
                          found == "silver", f"got '{found}'")
                for cust in BRONZE_CUSTOMERS:
                    found = _find_tier(cust["name"])
                    check(f"Customer '{cust['name']}' tagged Bronze tier",
                          found == "bronze", f"got '{found}'")

            # Total_Spent verification: top customer's spent must match DB
            # (live-derived TOP_SPENT or fallback constant).
            if spent_col is not None and name_col is not None:
                top_cust = TOP10[0]
                top_first = top_cust["name"].split()[0].lower()
                top_last = top_cust["name"].split()[-1].lower()
                expected_top_spent = TOP_SPENT
                for r in data_rows:
                    name = str(grid[r].get(name_col, "")).strip().lower()
                    if top_first in name and top_last in name:
                        spent_val = grid[r].get(spent_col)
                        try:
                            spent_num = float(str(spent_val).replace("$", "").replace(",", "").strip())
                            check(f"Top customer '{top_cust['name']}' Total_Spent ~ {expected_top_spent:.2f}",
                                  abs(spent_num - expected_top_spent) <= 5.0,
                                  f"got {spent_val}")
                        except (TypeError, ValueError):
                            check(f"Top customer '{top_cust['name']}' Total_Spent ~ {expected_top_spent:.2f}",
                                  False, f"got {spent_val}")
                        break

    cur.close()
    conn.close()


def check_gcal(launch_time_str=None):
    print("\n=== Checking Google Calendar ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("""
        SELECT summary, start_datetime, end_datetime, description
        FROM gcal.events
        WHERE LOWER(summary) LIKE '%vip%' AND LOWER(summary) LIKE '%appreciation%'
    """)
    events = cur.fetchall()
    check("VIP Appreciation Day event created", len(events) >= 1,
          f"Found {len(events)} matching events")

    if events and launch_time_str:
        try:
            launch_time = datetime.fromisoformat(launch_time_str)
            if launch_time.tzinfo is None:
                launch_time = launch_time.replace(tzinfo=timezone.utc)
            target_date = launch_time + timedelta(days=30)
            for event in events:
                event_start = event[1]
                if event_start.tzinfo is None:
                    event_start = event_start.replace(tzinfo=timezone.utc)
                diff_days = abs((event_start.date() - target_date.date()).days)
                if diff_days <= 2:
                    check("VIP Appreciation Day is ~30 days from launch", True)
                    break
            else:
                check("VIP Appreciation Day is ~30 days from launch", False,
                      f"Closest event at {events[0][1]}, expected ~{target_date.date()}")
        except Exception as e:
            check("VIP Appreciation Day date check", False, str(e))

    cur.close()
    conn.close()


def check_emails():
    print("\n=== Checking VIP Emails ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("""
        SELECT subject, from_addr, to_addr, body_text
        FROM email.messages
    """)
    all_emails = cur.fetchall()
    conn.close()

    gold_emails_found = 0
    for customer in GOLD_CUSTOMERS:
        customer_email = customer["email"].lower()
        found = False
        for subj, from_addr, to_addr, body in all_emails:
            recipients = parse_recipients(to_addr)
            if customer_email in recipients:
                found = True
                subj_lower = (subj or "").lower()
                from_lower = (from_addr or "").lower()
                body_lower = (body or "").lower()

                check(f"Email sent to {customer['name']} ({customer['email']})", True)
                check(f"Email from vip@store.example.com for {customer['name']}",
                      "vip@store.example.com" in from_lower,
                      f"From: {from_addr}")
                check(f"Subject contains 'VIP' for {customer['name']}",
                      "vip" in subj_lower, f"Subject: {subj}")
                check(f"Subject contains 'discount' or '20%' for {customer['name']}",
                      "discount" in subj_lower or "20%" in subj_lower,
                      f"Subject: {subj}")
                gold_emails_found += 1
                break
        if not found:
            check(f"Email sent to {customer['name']} ({customer['email']})", False,
                  "No matching email found")

    check("All 3 Gold tier customers emailed", gold_emails_found == 3,
          f"Found {gold_emails_found}/3 emails")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=" * 70)
    print("WC VIP CUSTOMER GSHEET GCAL EMAIL - EVALUATION")
    print("=" * 70)

    check_gsheet()
    check_gcal(args.launch_time)
    check_emails()

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    overall = FAIL_COUNT == 0
    print(f"  Overall: {'PASS' if overall else 'FAIL'}")
    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
