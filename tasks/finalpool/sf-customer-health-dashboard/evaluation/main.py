"""
Evaluation script for sf-customer-health-dashboard task.

Checks:
1. Excel file Customer_Health.xlsx with 3 sheets
2. Notion page created
3. Calendar events for critical accounts with score < 15
"""
import argparse
import json
import os
import sys
from datetime import date

import openpyxl
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0

# Fixed analysis date. The task schedules follow-up calls starting on
# 2026-03-09; "today" for the recency calc is the day before. Pinning this
# removes the eval-flakiness from CURRENT_DATE (results drift day by day and
# across timezones otherwise).
AS_OF_DATE = "2026-03-08"


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        detail_str = f": {detail[:200]}" if detail else ""
        print(f"  [FAIL] {name}{detail_str}")


def num_close(a, b, abs_tol=1.0, rel_tol=0.05):
    try:
        a_f, b_f = float(a), float(b)
        return abs(a_f - b_f) <= max(abs_tol, abs(b_f) * rel_tol)
    except (TypeError, ValueError):
        return False


def str_match(a, b):
    if a is None or b is None:
        return a is None and b is None
    return str(a).strip().lower() == str(b).strip().lower()


def compute_expected_values():
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("""
        WITH order_stats AS (
            SELECT "CUSTOMER_ID",
                COUNT(*) as order_count,
                MAX("ORDER_DATE") as last_order_date
            FROM sf_data."SALES_DW__PUBLIC__ORDERS"
            GROUP BY "CUSTOMER_ID"
        ),
        health AS (
            SELECT c."CUSTOMER_ID", c."CUSTOMER_NAME", c."SEGMENT", c."REGION",
                ROUND(GREATEST(0, 100 - (DATE '2026-03-08' - COALESCE(os.last_order_date, c."SIGNUP_DATE")))::numeric, 2) as recency_score,
                ROUND(LEAST(100, COALESCE(os.order_count, 0) * 10)::numeric, 2) as frequency_score,
                ROUND(LEAST(100, c."LIFETIME_VALUE" / 50.0)::numeric, 2) as monetary_score,
                ROUND((0.4 * GREATEST(0, 100 - (DATE '2026-03-08' - COALESCE(os.last_order_date, c."SIGNUP_DATE"))) +
                0.3 * LEAST(100, COALESCE(os.order_count, 0) * 10) +
                0.3 * LEAST(100, c."LIFETIME_VALUE" / 50.0))::numeric, 2) as health_score
            FROM sf_data."SALES_DW__PUBLIC__CUSTOMERS" c
            LEFT JOIN order_stats os ON c."CUSTOMER_ID" = os."CUSTOMER_ID"
        )
        SELECT "CUSTOMER_NAME", "SEGMENT", "REGION",
            recency_score, frequency_score, monetary_score, health_score,
            CASE WHEN health_score >= 80 THEN 'Healthy'
                 WHEN health_score >= 50 THEN 'At Risk'
                 ELSE 'Critical' END as status
        FROM health
        ORDER BY health_score ASC
    """)
    all_rows = cur.fetchall()

    # Summary by segment
    segments = {}
    for r in all_rows:
        seg = r[1]
        if seg not in segments:
            segments[seg] = {"total": 0, "healthy": 0, "at_risk": 0, "critical": 0, "scores": []}
        segments[seg]["total"] += 1
        segments[seg]["scores"].append(float(r[6]))
        if r[7] == "Healthy":
            segments[seg]["healthy"] += 1
        elif r[7] == "At Risk":
            segments[seg]["at_risk"] += 1
        else:
            segments[seg]["critical"] += 1

    segment_summary = []
    for seg in sorted(segments.keys()):
        s = segments[seg]
        avg = round(sum(s["scores"]) / len(s["scores"]), 2)
        segment_summary.append({
            "Segment": seg,
            "Total_Customers": s["total"],
            "Healthy_Count": s["healthy"],
            "At_Risk_Count": s["at_risk"],
            "Critical_Count": s["critical"],
            "Avg_Health_Score": avg,
        })

    critical_rows = [r for r in all_rows if r[7] == "Critical"]
    below_15 = [r for r in all_rows if float(r[6]) < 15]

    cur.close()
    conn.close()

    return {
        "total_customers": len(all_rows),
        "critical_count": len(critical_rows),
        "at_risk_count": sum(1 for r in all_rows if r[7] == "At Risk"),
        "healthy_count": sum(1 for r in all_rows if r[7] == "Healthy"),
        "segment_summary": segment_summary,
        "below_15": below_15,
        "top5_lowest": all_rows[:5],
    }


def get_sheet(wb, name):
    for s in wb.sheetnames:
        if str_match(s, name):
            return wb[s]
    return None


def check_excel(agent_workspace, expected):
    print("\n=== Checking Excel Output ===")

    agent_file = os.path.join(agent_workspace, "Customer_Health.xlsx")
    check("Excel file exists", os.path.isfile(agent_file), f"Expected {agent_file}")
    if not os.path.isfile(agent_file):
        return

    try:
        wb = openpyxl.load_workbook(agent_file, data_only=True)
    except Exception as e:
        check("Excel file readable", False, str(e))
        return
    check("Excel file readable", True)

    for sn in ["Health Scores", "Critical Accounts", "Summary by Segment"]:
        found = any(str_match(s, sn) for s in wb.sheetnames)
        check(f"Sheet '{sn}' exists", found, f"Found: {wb.sheetnames}")

    # --- Health Scores ---
    print("\n--- Health Scores ---")
    ws = get_sheet(wb, "Health Scores")
    if ws:
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        # Strict: total customers exact (this is fixed by DB)
        check("Health Scores total rows",
              len(rows) == expected["total_customers"],
              f"Expected {expected['total_customers']}, got {len(rows)}")

        # Check that first 5 rows are the actual lowest scores (sorted ascending)
        if len(rows) >= 5:
            for i, exp_row in enumerate(expected["top5_lowest"][:5]):
                agent_row = rows[i]
                if agent_row and agent_row[6] is not None:
                    check(f"Row {i+1} health score == {exp_row[6]} (tol 0.5)",
                          num_close(agent_row[6], exp_row[6], abs_tol=0.5, rel_tol=0.0),
                          f"Expected ~{exp_row[6]}, got {agent_row[6]}")
                # Per-rank name match with tie-tolerance: agent's rank-i name should be
                # in the expected top-5 set (allows ties at the same score to swap order).
                if agent_row and agent_row[0] is not None:
                    agent_name = str(agent_row[0]).strip().lower()
                    expected_top5_names = {str(r[0] or "").strip().lower() for r in expected["top5_lowest"][:5]}
                    check(f"Row {i+1} customer name in expected top-5 (tie-tolerance)",
                          agent_name in expected_top5_names,
                          f"agent rank-{i+1}={agent_name}, expected_top5={expected_top5_names}")
            # Set-overlap requirement tightened to >=4/5 (was >=3/5).
            agent_top5_names = {str(r[0] or "").strip().lower() for r in rows[:5]}
            expected_top5_names = {str(r[0] or "").strip().lower() for r in expected["top5_lowest"][:5]}
            overlap = len(agent_top5_names & expected_top5_names)
            check(f"Top-5 lowest customer names overlap (>=4/5)",
                  overlap >= 4,
                  f"agent: {agent_top5_names}, expected: {expected_top5_names}")

            # Verify ascending sort (by health_score)
            scores = [float(r[6]) for r in rows if r[6] is not None]
            check("Health Scores sorted ascending by Health_Score",
                  scores == sorted(scores),
                  f"first 5: {scores[:5]}")

    # --- Critical Accounts ---
    print("\n--- Critical Accounts ---")
    ws = get_sheet(wb, "Critical Accounts")
    if ws:
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        # Critical count is deterministic from DB - require exact equality.
        check("Critical Accounts count",
              len(rows) == expected["critical_count"],
              f"Expected {expected['critical_count']}, got {len(rows)}")
        # Verify all entries have Status == Critical
        if rows:
            non_critical = [r for r in rows if r and r[7] and not str_match(r[7], "Critical")]
            check("All Critical Accounts rows have Status='Critical'",
                  len(non_critical) == 0,
                  f"Found {len(non_critical)} non-Critical rows")
            # Sort ascending
            scores = [float(r[6]) for r in rows if r[6] is not None]
            check("Critical Accounts sorted ascending",
                  scores == sorted(scores),
                  f"first 5: {scores[:5]}")

    # --- Summary by Segment ---
    print("\n--- Summary by Segment ---")
    ws = get_sheet(wb, "Summary by Segment")
    if ws:
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        check("Summary rows", len(rows) == 4,
              f"Expected 4, got {len(rows)}")

        # Alphabetical sort
        seg_names = [str(r[0]) for r in rows if r and r[0]]
        check("Summary sorted alphabetically by Segment",
              seg_names == sorted(seg_names),
              f"order: {seg_names}")

        for exp_seg in expected["segment_summary"]:
            seg = exp_seg["Segment"]
            matched = None
            for r in rows:
                if r and str_match(r[0], seg):
                    matched = r
                    break
            if matched:
                # Exact integer counts (was tol=5/10)
                check(f"{seg} Total_Customers == {exp_seg['Total_Customers']}",
                      num_close(matched[1], exp_seg["Total_Customers"], abs_tol=1, rel_tol=0.0),
                      f"Expected {exp_seg['Total_Customers']}, got {matched[1]}")
                check(f"{seg} Healthy_Count == {exp_seg['Healthy_Count']}",
                      num_close(matched[2], exp_seg["Healthy_Count"], abs_tol=2, rel_tol=0.0),
                      f"Expected {exp_seg['Healthy_Count']}, got {matched[2]}")
                check(f"{seg} At_Risk_Count == {exp_seg['At_Risk_Count']}",
                      num_close(matched[3], exp_seg["At_Risk_Count"], abs_tol=2, rel_tol=0.0),
                      f"Expected {exp_seg['At_Risk_Count']}, got {matched[3]}")
                check(f"{seg} Critical_Count == {exp_seg['Critical_Count']}",
                      num_close(matched[4], exp_seg["Critical_Count"], abs_tol=2, rel_tol=0.0),
                      f"Expected {exp_seg['Critical_Count']}, got {matched[4]}")
                # Tighten avg health score (was tol=2.0); use 0.5 (2-decimal precision)
                check(f"{seg} Avg_Health_Score ~= {exp_seg['Avg_Health_Score']}",
                      num_close(matched[5], exp_seg["Avg_Health_Score"], abs_tol=0.5, rel_tol=0.0),
                      f"Expected {exp_seg['Avg_Health_Score']}, got {matched[5]}")
            else:
                check(f"Segment '{seg}' found", False, "Not in output")


def check_notion(expected):
    print("\n=== Checking Notion Page ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    # Title-exact match for "Customer Health Dashboard"
    cur.execute("""
        SELECT id, properties FROM notion.pages
        WHERE properties::text ILIKE '%Customer Health Dashboard%'
        ORDER BY id
    """)
    pages = cur.fetchall()
    check("Notion page titled 'Customer Health Dashboard' exists", len(pages) > 0,
          f"Found {len(pages)} pages with that title")

    if pages:
        page_id = pages[0][0]
        cur.execute("""
            SELECT content FROM notion.blocks
            WHERE page_id = %s
        """, (page_id,))
        blocks = cur.fetchall()
        check("Notion page has content blocks", len(blocks) > 0,
              f"Found {len(blocks)} blocks")

        # Validate page content includes total counts and at least 5 lowest customers
        all_text = " ".join(str(b[0] or "") for b in blocks).lower()
        # Counts
        check(f"Notion mentions total customers ({expected['total_customers']})",
              str(expected["total_customers"]) in all_text,
              f"Sample: {all_text[:300]}")
        # Mention of categories
        check("Notion mentions 'Healthy' / 'At Risk' / 'Critical' status terms",
              "healthy" in all_text and ("at risk" in all_text or "at-risk" in all_text) and "critical" in all_text,
              f"Sample: {all_text[:300]}")
        # 5 lowest customers' names
        names_found = sum(1 for r in expected["top5_lowest"][:5]
                          if str(r[0] or "").lower() in all_text)
        check("Notion lists at least 4 of the 5 lowest-score customer names",
              names_found >= 4,
              f"Found {names_found}/5 names in content")

    cur.close()
    conn.close()


def check_calendar(expected):
    print("\n=== Checking Calendar Events ===")
    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    cur.execute("""
        SELECT summary, description, start_datetime, end_datetime
        FROM gcal.events
        WHERE LOWER(summary) LIKE '%follow%up%' OR LOWER(summary) LIKE '%follow-up%'
        ORDER BY start_datetime
    """)
    events = cur.fetchall()

    below_15_count = len(expected["below_15"])
    # Tighten to exact count
    check(f"Calendar events for critical accounts (score < 15) == {below_15_count}",
          len(events) == below_15_count,
          f"Expected {below_15_count}, got {len(events)}")

    # Check ALL events match lowest-score customers in order
    if events and expected["below_15"]:
        match_count = 0
        for i in range(min(len(events), len(expected["below_15"]))):
            evt_summary = events[i][0] or ""
            exp_name = expected["below_15"][i][0]
            if exp_name.lower() in evt_summary.lower():
                match_count += 1
        n_check = min(len(events), len(expected["below_15"]))
        check(f"All events match lowest-score order ({match_count}/{n_check})",
              match_count >= n_check - 1,
              f"Order mismatch")

    # Check events start from 2026-03-09 and one per day at 10:00
    from datetime import time
    if events:
        first_dt = events[0][2]
        if first_dt:
            first_date = first_dt.date() if hasattr(first_dt, 'date') else first_dt
            check("First event on 2026-03-09",
                  str(first_date) == "2026-03-09",
                  f"First event date: {first_date}")
            # 10:00 AM
            try:
                start_time = first_dt.time() if hasattr(first_dt, 'time') else None
                check("First event at 10:00 AM",
                      start_time is not None and start_time.hour == 10 and start_time.minute == 0,
                      f"Got time: {start_time}")
            except AttributeError:
                pass
        # 30-minute duration on each event
        if len(events) > 0:
            ok_30min = 0
            for ev in events:
                if ev[2] and ev[3]:
                    delta = (ev[3] - ev[2]).total_seconds() / 60
                    if 25 <= delta <= 35:
                        ok_30min += 1
            check(f"Events are 30 minutes long ({ok_30min}/{len(events)})",
                  ok_30min >= len(events) - 1,
                  f"{ok_30min}/{len(events)} have 30-min duration")
        # Description mentions segment, region, score
        desc_ok = 0
        for ev in events:
            d = str(ev[1] or "").lower()
            if any(s in d for s in ["segment", "region", "score"]):
                desc_ok += 1
        check(f"Event descriptions reference segment/region/score ({desc_ok}/{len(events)})",
              desc_ok >= len(events) - 1,
              f"{desc_ok}/{len(events)} have substantive description")

    cur.close()
    conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=== Computing Expected Values ===")
    try:
        expected = compute_expected_values()
        print(f"  Total customers: {expected['total_customers']}")
        print(f"  Critical: {expected['critical_count']}")
        print(f"  Customers with score < 15: {len(expected['below_15'])}")
    except Exception as e:
        print(f"  ERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    check_excel(args.agent_workspace, expected)
    check_notion(expected)
    check_calendar(expected)

    total = PASS_COUNT + FAIL_COUNT
    pass_rate = PASS_COUNT / total if total > 0 else 0
    # Require all checks pass (FAIL_COUNT == 0)
    success = FAIL_COUNT == 0

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    print(f"  Pass Rate: {pass_rate:.1%}")
    print(f"  Overall: {'PASS' if success else 'FAIL'}")

    if args.res_log_file:
        result = {
            "passed": PASS_COUNT,
            "failed": FAIL_COUNT,
            "pass_rate": round(pass_rate, 3),
            "success": success,
        }
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
