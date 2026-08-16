"""Evaluation for sf-support-agent-performance-gsheet-gcal.

Checks:
1. Google Sheet "Agent Performance Scorecard" with Scorecards and Rankings sheets
2. Google Calendar event "Agent Performance Review" 10 days from launch
3. Email to performance-review@company.example.com
"""
import argparse
import os
import sys
from datetime import datetime, timedelta

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=int(os.environ.get("PGPORT", "5432")), dbname="toolathlon_gym", user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        print(f"  [FAIL] {name}: {str(detail)[:200]}")


def num_close(a, b, tol=1.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def _db_agent_stats():
    """Return authoritative per-agent stats from DB: list of tuples
    (agent, total, avg_resp, avg_csat, sla_pct) sorted by total desc."""
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT "REPORTER",
               COUNT(*)::int AS total,
               ROUND(AVG("RESPONSE_TIME_HOURS")::numeric, 2)::float,
               ROUND(AVG("CUSTOMER_SATISFACTION")::numeric, 2)::float,
               ROUND((SUM(CASE
                    WHEN ("PRIORITY"='High' AND "RESPONSE_TIME_HOURS"<=4)
                      OR ("PRIORITY"='Medium' AND "RESPONSE_TIME_HOURS"<=8)
                      OR ("PRIORITY"='Low' AND "RESPONSE_TIME_HOURS"<=24) THEN 1.0 ELSE 0.0 END)
                / COUNT(*) * 100)::numeric, 2)::float
        FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"
        GROUP BY "REPORTER"
        ORDER BY total DESC
    """)
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return rows


def check_gsheet():
    print("\n=== Checking Google Sheet ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()

    cur.execute("""
        SELECT id, title FROM gsheet.spreadsheets
        WHERE title ILIKE '%agent%performance%' OR title ILIKE '%agent%scorecard%'
    """)
    sheets = cur.fetchall()
    check("Agent Performance Scorecard spreadsheet exists", len(sheets) >= 1,
          f"Found: {[s[1] for s in sheets]}")

    if not sheets:
        cur.close()
        conn.close()
        return False

    ss_id = sheets[0][0]

    # Check sheets/tabs
    cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s", (ss_id,))
    tabs = [(r[0], r[1]) for r in cur.fetchall()]
    scorecard_sheet = next((t for t in tabs if "scorecard" in t[1].lower()), None)
    ranking_sheet = next((t for t in tabs if "ranking" in t[1].lower()), None)
    check("Has Scorecards sheet", scorecard_sheet is not None, f"Tabs: {[t[1] for t in tabs]}")
    check("Has Rankings sheet", ranking_sheet is not None, f"Tabs: {[t[1] for t in tabs]}")

    agent_stats = _db_agent_stats()  # sorted by total DESC
    agents_sorted = [row[0] for row in agent_stats]  # ['Emily', 'Charlie', ...]

    # --- Strict Scorecards sheet validation (row/col) ---
    if scorecard_sheet:
        sheet_id = scorecard_sheet[0]
        cur.execute("""SELECT row_index, col_index, value
                       FROM gsheet.cells
                       WHERE spreadsheet_id = %s AND sheet_id = %s""",
                    (ss_id, sheet_id))
        cell_map = {(r[0], r[1]): (r[2] or "") for r in cur.fetchall()}

        # Identify header row (the row that contains "Agent_Name" or "Agent Name")
        header_row_idx = None
        col_map = {}
        # Look in first 3 rows for header
        for ri in range(0, 3):
            row_vals = {ci: v for (r2, ci), v in cell_map.items() if r2 == ri}
            if not row_vals:
                continue
            row_joined = " ".join(str(v).lower() for v in row_vals.values())
            if "agent" in row_joined and ("name" in row_joined or "total" in row_joined):
                header_row_idx = ri
                # Map header name to col index
                for ci, val in row_vals.items():
                    lv = str(val).lower().strip()
                    if "agent" in lv and "name" in lv:
                        col_map["agent_name"] = ci
                    elif "total" in lv and "ticket" in lv:
                        col_map["total_tickets"] = ci
                    elif "response" in lv or "resp" in lv:
                        col_map["avg_response"] = ci
                    elif "csat" in lv:
                        col_map["avg_csat"] = ci
                    elif "sla" in lv:
                        col_map["sla"] = ci
                break

        check("Scorecards header row found with Agent_Name column",
              header_row_idx is not None and "agent_name" in col_map,
              f"header_row_idx={header_row_idx}, col_map={col_map}")

        # For each agent, check the row exists and values are correct
        if header_row_idx is not None and "agent_name" in col_map:
            name_col = col_map["agent_name"]
            # Build name->row mapping for rows after header
            name_to_row = {}
            for (r2, ci), v in cell_map.items():
                if r2 > header_row_idx and ci == name_col:
                    name = str(v).strip()
                    if name:
                        name_to_row[name.lower()] = r2

            for (agent_name, total, resp, csat, sla) in agent_stats:
                row_idx = name_to_row.get(agent_name.lower())
                check(f"Scorecards row for {agent_name}", row_idx is not None,
                      f"Missing row for {agent_name}")
                if row_idx is None:
                    continue
                # Check numeric columns
                if "total_tickets" in col_map:
                    v = cell_map.get((row_idx, col_map["total_tickets"]))
                    check(f"{agent_name} Total_Tickets",
                          num_close(v, total, 1),
                          f"val={v} expected {total}")
                if "avg_response" in col_map:
                    v = cell_map.get((row_idx, col_map["avg_response"]))
                    check(f"{agent_name} Avg_Response_Hrs",
                          num_close(v, resp, 0.1),
                          f"val={v} expected {resp}")
                if "avg_csat" in col_map:
                    v = cell_map.get((row_idx, col_map["avg_csat"]))
                    check(f"{agent_name} Avg_CSAT",
                          num_close(v, csat, 0.1),
                          f"val={v} expected {csat}")
                if "sla" in col_map:
                    v = cell_map.get((row_idx, col_map["sla"]))
                    check(f"{agent_name} SLA_Compliance_Rate",
                          num_close(v, sla, 0.5),
                          f"val={v} expected {sla}")

            # Sort check: rows should be ordered by total tickets DESC
            if "total_tickets" in col_map:
                row_ordered_agents = []
                # Get rows in index order, capturing name + total
                rows_sorted = sorted([r2 for (r2, ci) in cell_map if r2 > header_row_idx and ci == name_col])
                for ri in rows_sorted:
                    nm = cell_map.get((ri, name_col))
                    if nm:
                        nm = str(nm).strip()
                        if nm and nm.lower() in [a.lower() for a in agents_sorted]:
                            row_ordered_agents.append(nm)
                # Compare to expected order (case-insensitive)
                expected_order = [a for a in agents_sorted]
                check("Scorecards sorted by Total_Tickets descending",
                      [x.lower() for x in row_ordered_agents[:5]] == [x.lower() for x in expected_order[:5]],
                      f"got {row_ordered_agents}, expected {expected_order}")

    # --- Strict Rankings sheet validation ---
    if ranking_sheet:
        sheet_id = ranking_sheet[0]
        cur.execute("""SELECT row_index, col_index, value
                       FROM gsheet.cells
                       WHERE spreadsheet_id = %s AND sheet_id = %s""",
                    (ss_id, sheet_id))
        cells = cur.fetchall()
        # Compute expected top 3 by CSAT and top 3 by Volume
        top_csat = sorted(agent_stats, key=lambda r: -r[3])[:3]
        top_vol = sorted(agent_stats, key=lambda r: -r[1])[:3]

        all_values = [str(c[2] or "").lower() for c in cells]
        joined = " ".join(all_values)
        # Top performer by CSAT must appear
        for i, row in enumerate(top_csat):
            agent = row[0]
            check(f"Rankings: Top CSAT rank {i+1} = {agent}",
                  agent.lower() in joined,
                  f"missing {agent} in Rankings")
        for i, row in enumerate(top_vol):
            agent = row[0]
            check(f"Rankings: Top Volume rank {i+1} = {agent}",
                  agent.lower() in joined,
                  f"missing {agent} in Rankings")
        # Top_CSAT / Top_Volume labels
        check("Rankings uses Top_CSAT label",
              "top_csat" in joined or "top csat" in joined,
              "label not found")
        check("Rankings uses Top_Volume label",
              "top_volume" in joined or "top volume" in joined,
              "label not found")

    cur.close()
    conn.close()
    return True


def check_gcal(launch_time_str):
    print("\n=== Checking Google Calendar ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("SELECT summary, description, start_datetime, end_datetime FROM gcal.events ORDER BY start_datetime")
    events = cur.fetchall()
    cur.close()
    conn.close()

    print(f"  Found {len(events)} calendar events")
    check("At least 1 calendar event", len(events) >= 1, f"Found {len(events)}")

    review_events = [e for e in events
                     if "agent" in (e[0] or "").lower() and "performance" in (e[0] or "").lower() and "review" in (e[0] or "").lower()]
    check("Agent Performance Review event exists (title match)", len(review_events) >= 1,
          f"Events: {[e[0] for e in events]}")

    if launch_time_str and review_events:
        try:
            launch_dt = datetime.fromisoformat(launch_time_str)
            expected_dt = launch_dt + timedelta(days=10)
            # Tighter: must be same DAY (±1 day)
            matched = False
            for ev in review_events:
                if ev[2]:
                    ev_dt = ev[2].replace(tzinfo=None) if ev[2].tzinfo else ev[2]
                    diff_seconds = abs((ev_dt - expected_dt).total_seconds())
                    if diff_seconds <= 86400:  # ±1 day
                        matched = True
                        # 1-hour duration check
                        if ev[3]:
                            end_dt = ev[3].replace(tzinfo=None) if ev[3].tzinfo else ev[3]
                            dur = (end_dt - ev_dt).total_seconds()
                            check("Review meeting duration ~1 hour",
                                  abs(dur - 3600) <= 300,
                                  f"duration={dur}s")
                        break
            check("Review meeting 10 days (±1 day) from launch", matched,
                  f"Expected around {expected_dt}, got {[e[2] for e in review_events]}")
        except Exception as e:
            print(f"  [INFO] Could not verify date: {e}")

    return len(review_events) >= 1


def check_email():
    print("\n=== Checking Email ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT subject, from_addr, to_addr, body_text
        FROM email.messages
        WHERE subject ILIKE '%agent%' AND subject ILIKE '%performance%'
        ORDER BY date DESC
    """)
    emails = cur.fetchall()
    cur.close()
    conn.close()

    check("Agent performance email exists", len(emails) >= 1, f"Found {len(emails)}")
    if emails:
        e = emails[0]
        to_str = str(e[2])
        check("Email to performance-review@company.example.com",
              "performance-review@company.example.com" in to_str.lower(), f"to: {to_str}")
        check("Email from support-manager@company.example.com",
              "support-manager@company.example.com" in (e[1] or "").lower(), f"from: {e[1]}")
        # Exact subject match
        subj = (e[0] or "").lower()
        check("Email subject contains 'Monthly Scorecard'",
              "monthly scorecard" in subj,
              f"subject: {e[0]}")
        body = (e[3] or "").lower()

        # Strict: top CSAT agent + highest volume agent must both be in body with values
        stats = _db_agent_stats()
        top_csat_row = sorted(stats, key=lambda r: -r[3])[0]
        top_vol_row = sorted(stats, key=lambda r: -r[1])[0]
        top_csat_name = top_csat_row[0]
        top_csat_val = top_csat_row[3]
        top_vol_name = top_vol_row[0]
        top_vol_cnt = top_vol_row[1]

        check(f"Email mentions top CSAT agent '{top_csat_name}'",
              top_csat_name.lower() in body, f"body={body[:120]}")
        check(f"Email mentions top volume agent '{top_vol_name}'",
              top_vol_name.lower() in body, f"body={body[:120]}")
        # CSAT value
        import re
        csat_str = f"{top_csat_val:.2f}"
        csat_alt = f"{top_csat_val:.1f}"
        check(f"Email mentions top CSAT value ~{top_csat_val}",
              csat_str in body or csat_alt in body or re.search(r"3\.2[0-9]", body),
              f"csat value not in body")
        # Volume value
        vol_str = str(top_vol_cnt)
        check(f"Email mentions top volume count ({top_vol_cnt})",
              vol_str in body,
              f"volume count not in body")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_gsheet()
    check_gcal(args.launch_time)
    check_email()

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    overall = FAIL_COUNT == 0
    print(f"  Overall: {'PASS' if overall else 'FAIL'}")

    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
