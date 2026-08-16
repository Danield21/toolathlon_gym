"""Evaluation for sf-support-resolution-gsheet."""
import argparse
import json
import os
import sys
import psycopg2

DB = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
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
        detail_str = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{detail_str}")


def num_close(a, b, tol=1.0):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def get_expected():
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute(
        """
    SELECT "ISSUE_TYPE",
           COUNT(*),
           COUNT(CASE WHEN "STATUS" = 'Closed' THEN 1 END),
           ROUND(AVG(CASE WHEN "RESOLVED_AT" IS NOT NULL THEN EXTRACT(EPOCH FROM ("RESOLVED_AT" - "CREATED_AT"))/3600.0 END)::numeric, 1),
           ROUND(AVG("CUSTOMER_SATISFACTION")::numeric, 2)
    FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"
    GROUP BY "ISSUE_TYPE"
    """
    )
    issue_stats = {}
    for r in cur.fetchall():
        issue_stats[r[0]] = {
            "total": int(r[1]),
            "resolved": int(r[2]),
            "avg_res_hrs": float(r[3]) if r[3] is not None else None,
            "avg_sat": float(r[4]),
        }
    cur.execute(
        """
    SELECT "PRIORITY",
           COUNT(*),
           ROUND(AVG("RESPONSE_TIME_HOURS")::numeric, 1),
           ROUND( (COUNT(CASE WHEN "RESPONSE_TIME_HOURS" <= "SLA_HOURS" THEN 1 END)::numeric / COUNT(*)) * 100, 1)
    FROM sf_data."SUPPORT_CENTER__PUBLIC__TICKETS"
    GROUP BY "PRIORITY"
    """
    )
    priority_stats = {}
    for r in cur.fetchall():
        priority_stats[r[0]] = {
            "count": int(r[1]),
            "avg_response": float(r[2]),
            "sla_pct": float(r[3]),
        }
    cur.close()
    conn.close()
    # Issue type with longest avg resolution
    longest_issue = max(issue_stats.keys(), key=lambda k: issue_stats[k]["avg_res_hrs"] or 0)
    return {"issue": issue_stats, "priority": priority_stats, "longest_issue": longest_issue}


def check_gsheet(expected):
    print("\n=== Checking Google Sheet ===")
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
    except psycopg2.Error as e:
        check("DB connect", False, str(e))
        return

    cur.execute("SELECT id, title FROM gsheet.spreadsheets WHERE LOWER(title) LIKE '%%support resolution analysis%%'")
    ss_rows = cur.fetchall()
    check("Spreadsheet 'Support Resolution Analysis' exists", len(ss_rows) >= 1, f"Found {len(ss_rows)}")
    if not ss_rows:
        cur.close()
        conn.close()
        return
    ss_id = ss_rows[0][0]

    # By Issue Type
    cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id=%s", (ss_id,))
    sheet_list = cur.fetchall()
    issue_sheet = next((s for s, t in sheet_list if t and t.strip().lower() == "by issue type"), None)
    priority_sheet = next((s for s, t in sheet_list if t and t.strip().lower() == "by priority"), None)
    check("Sheet 'By Issue Type' present", issue_sheet is not None, f"got {[t for _, t in sheet_list]}")
    check("Sheet 'By Priority' present", priority_sheet is not None, f"got {[t for _, t in sheet_list]}")

    if issue_sheet is not None:
        cur.execute("SELECT row_index, col_index, value FROM gsheet.cells WHERE sheet_id=%s ORDER BY row_index, col_index", (issue_sheet,))
        cells = cur.fetchall()
        grid = {}
        for r, c, v in cells:
            grid.setdefault(r, {})[c] = str(v).strip() if v is not None else ""
        # Header row
        header_row = None
        for r in sorted(grid.keys()):
            joined = " ".join(grid[r].values()).lower()
            if "issue_type" in joined.replace(" ", "_") or ("issue" in joined and "type" in joined):
                if "total" in joined and "tickets" in joined:
                    header_row = r
                    break
        check("By Issue Type header found", header_row is not None)
        if header_row is not None:
            col_map = {}
            for col, val in grid[header_row].items():
                v = val.strip().lower().replace(" ", "_")
                if v == "issue_type":
                    col_map["type"] = col
                elif v == "total_tickets":
                    col_map["total"] = col
                elif v == "resolved_count":
                    col_map["resolved"] = col
                elif v == "avg_resolution_hours":
                    col_map["avg_res"] = col
                elif v == "avg_satisfaction":
                    col_map["avg_sat"] = col

            data_rows = [grid[r] for r in sorted(grid.keys()) if r > header_row]
            # Sort order: by Avg_Resolution_Hours descending
            res_hours = []
            for row in data_rows:
                try:
                    res_hours.append(float(row.get(col_map.get("avg_res", -1), "0")))
                except (TypeError, ValueError):
                    pass
            check("By Issue Type sorted by Avg_Resolution_Hours descending",
                  res_hours == sorted(res_hours, reverse=True),
                  f"got {res_hours}")

            by_type = {}
            for row in data_rows:
                t = row.get(col_map.get("type", -1), "").strip()
                if t:
                    by_type[t] = row
            for itype, exp in expected["issue"].items():
                row = by_type.get(itype)
                check(f"GSheet Issue Type '{itype}' present", row is not None)
                if row is None:
                    continue
                try:
                    total = int(float(row.get(col_map.get("total", -1), "0")))
                except (TypeError, ValueError):
                    total = None
                try:
                    resolved = int(float(row.get(col_map.get("resolved", -1), "0")))
                except (TypeError, ValueError):
                    resolved = None
                try:
                    avg_res = float(row.get(col_map.get("avg_res", -1), "0"))
                except (TypeError, ValueError):
                    avg_res = None
                try:
                    avg_sat = float(row.get(col_map.get("avg_sat", -1), "0"))
                except (TypeError, ValueError):
                    avg_sat = None
                check(f"  {itype}.Total_Tickets == {exp['total']}", total == exp["total"], f"got {total}")
                check(f"  {itype}.Resolved_Count == {exp['resolved']}", resolved == exp["resolved"], f"got {resolved}")
                if exp["avg_res_hrs"] is not None:
                    check(f"  {itype}.Avg_Resolution_Hours ≈ {exp['avg_res_hrs']}",
                          avg_res is not None and abs(avg_res - exp["avg_res_hrs"]) <= 0.1,
                          f"got {avg_res}")
                check(f"  {itype}.Avg_Satisfaction ≈ {exp['avg_sat']}",
                      avg_sat is not None and abs(avg_sat - exp["avg_sat"]) <= 0.05,
                      f"got {avg_sat}")

    if priority_sheet is not None:
        cur.execute("SELECT row_index, col_index, value FROM gsheet.cells WHERE sheet_id=%s ORDER BY row_index, col_index", (priority_sheet,))
        cells = cur.fetchall()
        grid = {}
        for r, c, v in cells:
            grid.setdefault(r, {})[c] = str(v).strip() if v is not None else ""
        header_row = None
        for r in sorted(grid.keys()):
            joined = " ".join(grid[r].values()).lower()
            if "priority" in joined and ("ticket_count" in joined.replace(" ", "_") or ("ticket" in joined and "count" in joined)):
                header_row = r
                break
        check("By Priority header found", header_row is not None)
        if header_row is not None:
            col_map = {}
            for col, val in grid[header_row].items():
                v = val.strip().lower().replace(" ", "_")
                if v == "priority":
                    col_map["priority"] = col
                elif v == "ticket_count":
                    col_map["count"] = col
                elif v == "avg_response_hours":
                    col_map["resp"] = col
                elif v == "sla_compliance_pct":
                    col_map["sla"] = col

            data_rows = [grid[r] for r in sorted(grid.keys()) if r > header_row]
            priorities = [row.get(col_map.get("priority", -1), "").strip() for row in data_rows]
            check("By Priority sorted alphabetically", priorities == sorted(priorities), f"got {priorities}")

            by_p = {p.lower(): row for p, row in zip(priorities, data_rows) if p}
            for prio, exp in expected["priority"].items():
                row = by_p.get(prio.lower())
                check(f"GSheet Priority '{prio}' present", row is not None)
                if row is None:
                    continue
                try:
                    cnt = int(float(row.get(col_map.get("count", -1), "0")))
                except (TypeError, ValueError):
                    cnt = None
                try:
                    resp = float(row.get(col_map.get("resp", -1), "0"))
                except (TypeError, ValueError):
                    resp = None
                try:
                    sla = float(row.get(col_map.get("sla", -1), "0"))
                except (TypeError, ValueError):
                    sla = None
                check(f"  {prio}.Ticket_Count == {exp['count']}", cnt == exp["count"], f"got {cnt}")
                check(f"  {prio}.Avg_Response_Hours ≈ {exp['avg_response']}",
                      resp is not None and abs(resp - exp["avg_response"]) <= 0.1, f"got {resp}")
                check(f"  {prio}.SLA_Compliance_Pct ≈ {exp['sla_pct']}",
                      sla is not None and abs(sla - exp["sla_pct"]) <= 0.5, f"got {sla}")

    cur.close()
    conn.close()


def check_email(expected):
    print("\n=== Checking Email ===")
    try:
        conn = psycopg2.connect(**DB)
        cur = conn.cursor()
    except psycopg2.Error as e:
        check("DB connect", False, str(e))
        return
    cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
    all_emails = cur.fetchall()
    cur.close()
    conn.close()

    found = None
    for subject, from_addr, to_addr, body in all_emails:
        to_str = ""
        if isinstance(to_addr, list):
            to_str = " ".join(str(r).lower() for r in to_addr)
        else:
            try:
                parsed = json.loads(str(to_addr))
                to_str = " ".join(str(r).lower() for r in parsed) if isinstance(parsed, list) else str(to_addr).lower()
            except (json.JSONDecodeError, TypeError):
                to_str = str(to_addr).lower()
        if "support-lead@company.com" in to_str:
            found = (subject, from_addr, to_addr, body)
            break
    check("Email sent to support-lead@company.com", found is not None)
    if found is None:
        return
    subject, from_addr, to_addr, body = found
    check("Email subject == 'Support Resolution Performance Report'",
          (subject or "").strip().lower() == "support resolution performance report",
          f"got: {subject}")
    body_lower = (body or "").lower()
    longest = expected["longest_issue"].lower()
    check(f"Email body mentions longest-resolution issue type '{expected['longest_issue']}'",
          longest in body_lower, f"missing {expected['longest_issue']}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    expected = get_expected()
    check_gsheet(expected)
    check_email(expected)

    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}")
    print(f"  Failed: {FAIL_COUNT}")
    all_ok = FAIL_COUNT == 0
    print(f"  Overall: {'PASS' if all_ok else 'FAIL'}")
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump({"passed": PASS_COUNT, "failed": FAIL_COUNT, "success": all_ok}, f)
    sys.exit(0 if all_ok else 1)


if __name__ == "__main__":
    main()
