"""
Evaluation for 12306-howtocook-team-trip-catering-excel-gcal task.

Checks:
1. Team_Trip_Plan.xlsx exists with Travel, Menu, Timeline sheets
2. Travel sheet has G11 and 07:00 departure
3. Menu sheet has >= 5 rows with Course_Type and Dish_Name columns
4. Timeline sheet has >= 5 rows
5. Notion page exists with team/trip/march in title
6. GCal has >= 2 new events on 2026-03-10
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
    print("\n=== Check 1: Excel Team_Trip_Plan.xlsx ===")

    xlsx_path = os.path.join(agent_workspace, "Team_Trip_Plan.xlsx")
    if not os.path.exists(xlsx_path):
        record("Team_Trip_Plan.xlsx exists", False, f"Not found at {xlsx_path}")
        return
    record("Team_Trip_Plan.xlsx exists", True)

    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    except Exception as e:
        record("Excel readable", False, str(e))
        return
    record("Excel readable", True)

    sheet_names_lower = [s.lower() for s in wb.sheetnames]
    has_travel = any("travel" in s for s in sheet_names_lower)
    has_menu = any("menu" in s for s in sheet_names_lower)
    has_timeline = any("timeline" in s for s in sheet_names_lower)

    record("Excel has Travel sheet", has_travel, f"Sheets: {wb.sheetnames}")
    record("Excel has Menu sheet", has_menu, f"Sheets: {wb.sheetnames}")
    record("Excel has Timeline sheet", has_timeline, f"Sheets: {wb.sheetnames}")

    if has_travel:
        ws_name = wb.sheetnames[next(i for i, s in enumerate(sheet_names_lower) if "travel" in s)]
        ws = wb[ws_name]
        rows = list(ws.iter_rows(values_only=True))
        data_rows = [r for r in rows[1:] if any(c for c in r)]
        record("Travel sheet has exactly 1 data row", len(data_rows) == 1, f"Found {len(data_rows)} rows")

        all_text = " ".join(str(c) for row in rows for c in row if c).lower()
        # Use exact-token match to avoid 'g11' matching 'g111' etc.
        import re
        has_g11 = bool(re.search(r"\bg11\b", all_text))
        record("Travel sheet contains G11", has_g11, f"Content sample: {all_text[:200]}")
        has_time = bool(re.search(r"\b07:00(:00)?\b", all_text)) or bool(re.search(r"\b7:00\b", all_text))
        record("Travel sheet shows 07:00 departure", has_time, f"Content: {all_text[:200]}")

        # Validate Total_Budget_CNY = Per_Person_CNY * 15 if both columns are present
        if rows:
            headers = [str(c).strip().lower() if c else "" for c in rows[0]]
            try:
                pp_idx = headers.index("per_person_cny")
                tb_idx = headers.index("total_budget_cny")
            except ValueError:
                pp_idx = tb_idx = -1
            if pp_idx >= 0 and tb_idx >= 0 and data_rows:
                ok = True
                for r in data_rows:
                    try:
                        pp = float(r[pp_idx])
                        tb = float(r[tb_idx])
                        if abs(tb - pp * 15) > 1.0:
                            ok = False
                            break
                    except (TypeError, ValueError):
                        ok = False
                        break
                record("Total_Budget_CNY = Per_Person_CNY * 15", ok,
                       f"Travel rows: {data_rows}")

    if has_menu:
        ws_name = wb.sheetnames[next(i for i, s in enumerate(sheet_names_lower) if "menu" in s)]
        ws = wb[ws_name]
        rows = list(ws.iter_rows(values_only=True))
        data_rows = [r for r in rows[1:] if any(c for c in r)]
        # Task says exactly 6 dishes (2 appetizers + 3 mains + 1 dessert/soup)
        record("Menu sheet has 6 rows", len(data_rows) == 6, f"Found {len(data_rows)} rows")

        if rows:
            headers = [str(c).lower() if c else "" for c in rows[0]]
            has_course = any("course" in h or "type" in h or "category" in h for h in headers)
            has_dish = any("dish" in h or "name" in h for h in headers)
            record("Menu has course type column", has_course, f"Headers: {rows[0]}")
            record("Menu has dish name column", has_dish, f"Headers: {rows[0]}")

    if has_timeline:
        ws_name = wb.sheetnames[next(i for i, s in enumerate(sheet_names_lower) if "timeline" in s)]
        ws = wb[ws_name]
        rows = list(ws.iter_rows(values_only=True))
        data_rows = [r for r in rows[1:] if any(c for c in r)]
        record("Timeline sheet has >= 6 rows", len(data_rows) >= 6, f"Found {len(data_rows)} rows")

    # --- Groundtruth value comparison ---
    # Travel: exact-match (G11/07:00/12:31/etc are derivable from 12306 data
    # given the task's "earliest departure" constraint).
    # Menu: structural validation only - dish names not derivable from
    # task spec ("preferably popular Chinese home cooking") so accepting
    # any valid dish that meets the schema.
    # Timeline: structural validation only.
    gt_path = os.path.join(groundtruth_workspace, "Team_Trip_Plan.xlsx")
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
        sn_lower = gt_sheet_name.strip().lower()

        if "travel" in sn_lower:
            # Travel: exact row count and exact value comparison
            record(f"GT '{gt_sheet_name}' row count exact",
                   len(agent_rows) == len(gt_rows),
                   f"Expected {len(gt_rows)}, got {len(agent_rows)}")
            for idx in range(len(gt_rows)):
                gt_row = gt_rows[idx]
                if idx < len(agent_rows):
                    a_row = agent_rows[idx]
                    row_ok = True
                    fail_detail = None
                    for col_idx in range(min(len(gt_row), len(a_row) if a_row else 0)):
                        gt_val = gt_row[col_idx]
                        a_val = a_row[col_idx]
                        if gt_val is None:
                            continue
                        if isinstance(gt_val, int):
                            ok = num_close(a_val, gt_val, 0)
                        elif isinstance(gt_val, float):
                            ok = num_close(a_val, gt_val, max(abs(gt_val) * 0.02, 1.0))
                        else:
                            ok = str_match(a_val, gt_val)
                        if not ok:
                            fail_detail = (col_idx, gt_val, a_val)
                            row_ok = False
                            break
                    if row_ok:
                        record(f"GT '{gt_sheet_name}' row {idx+1} values match", True)
                    else:
                        record(f"GT '{gt_sheet_name}' row {idx+1} col {fail_detail[0]+1}",
                               False, f"Expected {fail_detail[1]}, got {fail_detail[2]}")
                else:
                    record(f"GT '{gt_sheet_name}' row {idx+1} exists", False, "Row missing in agent")

        elif "menu" in sn_lower:
            # Menu: structural validation. Task says 6 dishes (2 appetizer + 3 main + 1 dessert/soup).
            # Validate course type distribution and that dish names are non-empty.
            record(f"GT '{gt_sheet_name}' row count exact",
                   len(agent_rows) == 6,
                   f"Expected 6 (2 appetizer + 3 main + 1 dessert/soup), got {len(agent_rows)}")

            # Build column index map from agent's header
            agent_header_rows = list(agent_ws.iter_rows(min_row=1, max_row=1, values_only=True))
            if agent_header_rows:
                agent_headers = [str(c).lower() if c else "" for c in agent_header_rows[0]]
                course_idx = next((i for i, h in enumerate(agent_headers) if "course" in h or "category" in h or "type" in h), -1)
                dish_idx = next((i for i, h in enumerate(agent_headers) if "dish" in h or "name" in h), -1)
                cook_idx = next((i for i, h in enumerate(agent_headers) if "cook" in h or "time" in h), -1)
                serv_idx = next((i for i, h in enumerate(agent_headers) if "serv" in h or "yield" in h), -1)
            else:
                course_idx = dish_idx = cook_idx = serv_idx = -1

            # Count by course type (allow flexible matching)
            appetizer_count = 0
            main_count = 0
            dessert_count = 0
            dish_names_valid = True
            for r in agent_rows:
                course = (str(r[course_idx]) if course_idx >= 0 and course_idx < len(r) else "").strip().lower()
                dish = (str(r[dish_idx]) if dish_idx >= 0 and dish_idx < len(r) else "").strip()
                if not dish or dish.lower() == "none":
                    dish_names_valid = False
                if "appetizer" in course or "starter" in course or "凉" in course:
                    appetizer_count += 1
                elif "main" in course or "entr" in course or "主" in course:
                    main_count += 1
                elif "dessert" in course or "soup" in course or "汤" in course or "甜" in course:
                    dessert_count += 1

            record(f"Menu has 2 appetizers, 3 mains, 1 dessert/soup (counts derived from Course_Type)",
                   appetizer_count == 2 and main_count == 3 and dessert_count == 1,
                   f"appetizer={appetizer_count}, main={main_count}, dessert/soup={dessert_count}")
            record("Menu dish names are all non-empty",
                   dish_names_valid,
                   "All Dish_Name cells must be non-empty")

            # Cook time and servings yield should be numeric and reasonable
            servings_ok = True
            cook_time_ok = True
            for r in agent_rows:
                if cook_idx >= 0 and cook_idx < len(r):
                    try:
                        ct = float(r[cook_idx])
                        if ct <= 0 or ct > 600:  # 0 to 10 hours
                            cook_time_ok = False
                    except (TypeError, ValueError):
                        cook_time_ok = False
                if serv_idx >= 0 and serv_idx < len(r):
                    try:
                        sv = float(r[serv_idx])
                        # Should be 15 (team size) per task
                        if abs(sv - 15) > 1:
                            servings_ok = False
                    except (TypeError, ValueError):
                        servings_ok = False
            record("Menu cook times are numeric and reasonable (1-600 min)",
                   cook_time_ok, "")
            record("Menu servings_yield matches team size 15",
                   servings_ok, "")

        elif "timeline" in sn_lower:
            # Timeline: at least 6 rows with non-empty time and activity columns.
            record(f"GT '{gt_sheet_name}' row count >= 6",
                   len(agent_rows) >= 6,
                   f"Expected >= 6, got {len(agent_rows)}")
            agent_header_rows = list(agent_ws.iter_rows(min_row=1, max_row=1, values_only=True))
            if agent_header_rows:
                agent_headers = [str(c).lower() if c else "" for c in agent_header_rows[0]]
                time_idx = next((i for i, h in enumerate(agent_headers) if "time" in h), -1)
                act_idx = next((i for i, h in enumerate(agent_headers) if "activ" in h or "task" in h), -1)
            else:
                time_idx = act_idx = -1

            timeline_ok = True
            for r in agent_rows:
                t = (str(r[time_idx]) if time_idx >= 0 and time_idx < len(r) else "").strip()
                a = (str(r[act_idx]) if act_idx >= 0 and act_idx < len(r) else "").strip()
                if not t or not a or t.lower() == "none" or a.lower() == "none":
                    timeline_ok = False
                    break
            record("Timeline rows have non-empty Time and Activity",
                   timeline_ok, "")

            # Verify required milestone times are covered (06:30, 07:00, 12:31, 13:00, 18:00, 20:00)
            # Allow some tolerance: 30 min window for each.
            all_times_text = " ".join(str(r[time_idx] or "") for r in agent_rows if time_idx >= 0 and time_idx < len(r)).lower()
            has_morning = "06:" in all_times_text or "6:30" in all_times_text or "07:00" in all_times_text or "7:00" in all_times_text
            has_arrival = "12:" in all_times_text or "13:00" in all_times_text
            has_dinner = "18:" in all_times_text or "20:" in all_times_text
            record("Timeline covers morning departure, midday arrival, evening dinner times",
                   has_morning and has_arrival and has_dinner,
                   f"morning:{has_morning} arrival:{has_arrival} dinner:{has_dinner}")
    gt_wb.close()


def check_notion():
    print("\n=== Check 2: Notion page for team trip ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("SELECT id, properties FROM notion.pages")
    pages = cur.fetchall()
    cur.close()
    conn.close()

    trip_page = None
    for page_id, props in pages:
        title = ""
        if isinstance(props, dict):
            title_obj = props.get("title", {})
            if isinstance(title_obj, dict):
                title_list = title_obj.get("title", [])
                if isinstance(title_list, list):
                    for t in title_list:
                        if isinstance(t, dict):
                            title += t.get("text", {}).get("content", "")
        title_lower = title.lower()
        # Tighten: title must reference team-building/trip AND a date marker (march or 03/10 or shanghai)
        has_trip_keyword = any(kw in title_lower for kw in ["team", "trip", "建设", "出行"])
        has_date_or_dest = any(kw in title_lower for kw in ["march", "shanghai", "上海", "3月", "03-10", "03/10", "2026-03-10", "team building"])
        if has_trip_keyword and has_date_or_dest:
            trip_page = (page_id, title)
            break

    record("Notion page exists with trip+date/destination context (team/trip AND march/shanghai/team-building)",
           trip_page is not None,
           f"Pages found: {len(pages)}, titles: {[str(p[1])[:80] for p in pages[:3]]}")


def check_gcal():
    print("\n=== Check 3: GCal events on 2026-03-10 ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("""
        SELECT summary, start_datetime, end_datetime FROM gcal.events
        WHERE start_datetime::date = '2026-03-10'
        ORDER BY start_datetime
    """)
    events = cur.fetchall()
    cur.close()
    conn.close()

    # Required: 3 specific events with the correct time slots.
    # 1) Departure 06:30-07:30
    # 2) Arrival/Check-in 12:30-14:00
    # 3) Team Dinner 18:00-20:00
    def hhmm(dt):
        try:
            return dt.strftime("%H:%M")
        except Exception:
            return str(dt)[11:16]

    def has_event(summary_keywords, start_hhmm, end_hhmm):
        for s, st, et in events:
            sl = (s or "").lower()
            if any(kw in sl for kw in summary_keywords):
                if hhmm(st) == start_hhmm and hhmm(et) == end_hhmm:
                    return True
        return False

    has_depart = has_event(
        ["depart", "departure", "station", "train"], "06:30", "07:30"
    )
    has_arrival = has_event(
        ["arriv", "check", "hotel"], "12:30", "14:00"
    )
    has_dinner = has_event(
        ["dinner", "team dinner", "banquet"], "18:00", "20:00"
    )
    summaries = [(e[0], hhmm(e[1]), hhmm(e[2])) for e in events]
    record(
        "GCal departure event on 2026-03-10 06:30-07:30",
        has_depart,
        f"Mar10 events: {summaries}",
    )
    record(
        "GCal arrival/check-in event on 2026-03-10 12:30-14:00",
        has_arrival,
        f"Mar10 events: {summaries}",
    )
    record(
        "GCal team dinner event on 2026-03-10 18:00-20:00",
        has_dinner,
        f"Mar10 events: {summaries}",
    )


def check_email():
    print("\n=== Check 4: Email sent to events@company.com ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("SELECT subject, from_addr, to_addr, body_text FROM email.messages")
    messages = cur.fetchall()
    cur.close()
    conn.close()

    sent = [m for m in messages if m[1] and "company.com" not in str(m[1])]
    # Check for outgoing emails (not the preprocess-injected incoming email)
    outgoing = []
    for subject, from_addr, to_addr, body_text in messages:
        to_str = ""
        if isinstance(to_addr, list):
            to_str = " ".join(str(r).lower() for r in to_addr)
        elif to_addr:
            try:
                parsed = json.loads(str(to_addr))
                to_str = " ".join(str(r).lower() for r in parsed) if isinstance(parsed, list) else str(to_addr).lower()
            except Exception:
                to_str = str(to_addr).lower()
        if "events@company.com" in to_str:
            outgoing.append((subject, from_addr, to_addr, body_text))

    record("Email sent to events@company.com", len(outgoing) >= 1,
           f"Total messages: {len(messages)}, matching: {len(outgoing)}")

    if outgoing:
        subject, _, _, body = outgoing[0]
        body_lower = ((subject or "") + " " + (body or "")).lower()
        has_trip = any(kw in body_lower for kw in ["trip", "travel", "train", "shanghai", "g11", "excel", "plan"])
        record("Email mentions trip/travel/plan content", has_trip,
               f"Subject: {subject}")


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
    check_email()

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

    if FAIL_COUNT == 0 and PASS_COUNT > 0:
        print("PASS")
        sys.exit(0)
    else:
        print("FAIL")
        sys.exit(1)


if __name__ == "__main__":
    main()
