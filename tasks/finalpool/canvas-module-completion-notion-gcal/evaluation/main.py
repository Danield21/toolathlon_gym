"""Evaluation for canvas-module-completion-notion-gcal."""
import argparse
import json
import os
import sys
from datetime import datetime, timedelta

import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=5432, dbname="toolathlon_gym", user="eigent", password="camel")

PASS_COUNT = 0
FAIL_COUNT = 0

EXPECTED_MODULES = ["Introduction", "Week 1", "Week 3", "Week 5", "Week 8"]


def check(name, condition, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if condition:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        detail_str = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{detail_str}")


def _title_to_str(title):
    title_str = ""
    if isinstance(title, list):
        for part in title:
            if isinstance(part, dict):
                title_str += part.get("plain_text", "") or part.get("text", {}).get("content", "")
    elif isinstance(title, str):
        title_str = title
    elif title:
        title_str = str(title)
    return title_str


def _extract_property(props, prop_name):
    """Best-effort extract a value from a Notion properties dict by name."""
    if not isinstance(props, dict):
        return None
    # case-insensitive find
    target = prop_name.lower().replace(" ", "_").replace(" ", "")
    for k, v in props.items():
        kn = (k or "").lower().replace(" ", "_")
        if kn == target or kn == prop_name.lower():
            return v
    return None


def _val_from_property(p):
    if not isinstance(p, dict):
        return None
    t = p.get("type")
    if not t:
        return None
    if t == "title":
        arr = p.get("title", [])
        return "".join(x.get("plain_text", "") or x.get("text", {}).get("content", "") for x in arr)
    if t == "rich_text":
        arr = p.get("rich_text", [])
        return "".join(x.get("plain_text", "") or x.get("text", {}).get("content", "") for x in arr)
    if t == "number":
        return p.get("number")
    if t == "select":
        sel = p.get("select")
        return sel.get("name") if isinstance(sel, dict) else None
    return None


def find_target_db():
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("SELECT id, title FROM notion.databases")
    dbs = cur.fetchall()
    db_id = None
    found_title = None
    for did, title in dbs:
        title_str = _title_to_str(title).strip().lower()
        # Tightened: must equal one of the canonical task titles exactly
        if title_str in (
            "ccc fall 2014 - module tracker",
            "ccc fall 2014 module tracker",
            "creative computing fall 2014 - module tracker",
            "creative computing fall 2014 module tracker",
        ):
            db_id = did
            found_title = _title_to_str(title)
            break
    cur.close()
    conn.close()
    return db_id, found_title, dbs


def check_notion():
    print("\n=== Checking Notion Database ===")
    db_id, title, all_dbs = find_target_db()
    check("Notion database titled 'CCC Fall 2014 - Module Tracker'", db_id is not None,
          f"Found: {[(d, _title_to_str(t)) for d, t in all_dbs]}")
    if not db_id:
        return

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    # Scope pages to target DB
    cur.execute("""
        SELECT id, properties
        FROM notion.pages
        WHERE archived = false AND parent->>'database_id' = %s
    """, (db_id,))
    pages = cur.fetchall()
    cur.close()
    conn.close()

    check(f"Notion DB has exactly 5 entries", len(pages) == 5,
          f"Found {len(pages)} pages in target DB")

    # Inspect each page's properties
    found_modules = []
    found_status_not_started = 0
    item_count_present = 0
    type_summary_present = 0
    for pid, props in pages:
        props_dict = props if isinstance(props, dict) else {}
        # Find Module_Name
        for prop_name in ["Module_Name", "Module Name", "Name"]:
            p = _extract_property(props_dict, prop_name)
            if p is not None:
                v = _val_from_property(p)
                if v:
                    found_modules.append(v.strip())
                    break
        # Status
        for sk in ["Status", "status"]:
            p = _extract_property(props_dict, sk)
            v = _val_from_property(p)
            if v and v.strip().lower() == "not started":
                found_status_not_started += 1
                break
        # Item_Count
        for ik in ["Item_Count", "Item Count"]:
            p = _extract_property(props_dict, ik)
            v = _val_from_property(p)
            if isinstance(v, (int, float)):
                item_count_present += 1
                break
        # Type_Summary
        for tk in ["Type_Summary", "Type Summary"]:
            p = _extract_property(props_dict, tk)
            v = _val_from_property(p)
            if v and v.strip():
                type_summary_present += 1
                break

    found_modules_lower = [m.lower() for m in found_modules]
    for m in EXPECTED_MODULES:
        check(f"Module '{m}' exists in tracker",
              any(m.lower() == fm or m.lower() in fm for fm in found_modules_lower),
              f"Found modules: {found_modules}")
    check("All 5 modules have Status='Not Started'",
          found_status_not_started == 5,
          f"Got {found_status_not_started}/5 with Not Started")
    check("All 5 modules have Item_Count number",
          item_count_present == 5,
          f"Got {item_count_present}/5 with Item_Count")
    check("All 5 modules have Type_Summary text",
          type_summary_present == 5,
          f"Got {type_summary_present}/5 with Type_Summary")


def _first_saturday_after(launch_dt):
    """Return the first Saturday strictly after launch_dt (date-aware, UTC)."""
    d = launch_dt
    # weekday: Monday=0...Saturday=5
    days_ahead = (5 - d.weekday()) % 7
    if days_ahead == 0:
        days_ahead = 7
    return (d + timedelta(days=days_ahead)).replace(hour=10, minute=0, second=0, microsecond=0)


def check_gcal(launch_time_str=None):
    print("\n=== Checking Google Calendar Events ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    # Read both UTC-equivalent and the stored timezone for time-of-day check.
    cur.execute("""
        SELECT summary, description,
               start_datetime AT TIME ZONE 'UTC',
               end_datetime AT TIME ZONE 'UTC',
               start_timezone, start_datetime
        FROM gcal.events
        WHERE LOWER(summary) LIKE '%%ccc module review meeting%%'
        ORDER BY start_datetime
    """)
    events = cur.fetchall()
    cur.close()
    conn.close()

    check("Exactly 4 'CCC Module Review Meeting' events",
          len(events) == 4,
          f"Found {len(events)}: {[(e[0], str(e[2])) for e in events]}")
    if len(events) < 1:
        return

    # All events must have summary "CCC Module Review Meeting" exactly
    for i, ev in enumerate(events):
        summary = ev[0]
        check(f"Event {i+1} summary is 'CCC Module Review Meeting'",
              (summary or "").strip().lower() == "ccc module review meeting",
              f"Got: {summary}")

    # Each event 1 hour (use UTC-equivalent for duration)
    for i, ev in enumerate(events):
        st, et = ev[2], ev[3]
        if st and et:
            duration = (et - st).total_seconds() / 60
            check(f"Event {i+1} duration is 60 minutes",
                  abs(duration - 60) < 1,
                  f"Duration: {duration} min")

    # 4 events on 4 consecutive Saturdays
    if len(events) >= 4:
        # Use UTC-equivalent date for date/weekday/spacing
        starts_utc = [e[2] for e in events[:4]]
        # If event is at 10:00 in a positive-offset zone (e.g. +02:00), UTC is 08:00 same day.
        # If at 10:00 UTC, raw is 10:00. If at 10:00 in negative-offset (e.g. -05:00), UTC is 15:00 same day.
        # In all common cases the date is the same day. Check Saturday using UTC date.
        sats_ok = all(s.weekday() == 5 for s in starts_utc if s)
        check("All 4 events fall on Saturdays",
              sats_ok,
              f"Weekdays: {[s.weekday() if s else None for s in starts_utc]}")
        gaps_ok = True
        for i in range(1, 4):
            gap = (starts_utc[i] - starts_utc[i-1]).total_seconds() / 86400
            if abs(gap - 7) > 0.5:
                gaps_ok = False
                break
        check("All 4 events are 7 days apart (consecutive Saturdays)",
              gaps_ok,
              f"Dates: {[str(s) for s in starts_utc]}")

        # All events should start at hour 10:00. Since timezone interpretation varies,
        # accept 10:00 in either UTC (wall-clock UTC) or in the stored start_timezone.
        for i, ev in enumerate(events[:4]):
            st_utc = ev[2]
            st_raw = ev[5]
            st_tz_name = ev[4]
            ok = False
            reason = []
            if st_utc and st_utc.hour == 10 and st_utc.minute == 0:
                ok = True
                reason.append("10:00 UTC")
            if st_raw is not None:
                # st_raw has tzinfo from DB — wall-clock equals .hour/.minute
                if st_raw.hour == 10 and st_raw.minute == 0:
                    ok = True
                    reason.append(f"10:00 {st_tz_name or 'local-stored-tz'}")
                # Also try interpreting in named tz if present
                if st_tz_name:
                    try:
                        from zoneinfo import ZoneInfo
                        tz = ZoneInfo(st_tz_name)
                        wc = st_raw.astimezone(tz)
                        if wc.hour == 10 and wc.minute == 0:
                            ok = True
                            reason.append(f"10:00 {st_tz_name}")
                    except Exception:
                        pass
            check(f"Event {i+1} starts at 10:00",
                  ok,
                  f"start_utc={st_utc}, start_raw={st_raw}, tz={st_tz_name}")

        # First Saturday after launch_time (if provided)
        if launch_time_str:
            try:
                launch_dt = datetime.fromisoformat(launch_time_str.replace("Z", "+00:00"))
                if launch_dt.tzinfo:
                    launch_dt_naive = launch_dt.replace(tzinfo=None)
                else:
                    launch_dt_naive = launch_dt
                expected_first_date = _first_saturday_after(launch_dt_naive).date()
                actual_first_date = starts_utc[0].date() if starts_utc[0] else None
                check("First event is first Saturday after launch_time",
                      actual_first_date == expected_first_date,
                      f"Expected first Saturday {expected_first_date}, got {actual_first_date}")
            except Exception as e:
                print(f"  [INFO] launch_time validation skipped: {e}")


def check_emails():
    print("\n=== Checking Emails ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT subject, from_addr, to_addr, COALESCE(body_text, body_html, '')
        FROM email.messages
    """)
    all_emails = cur.fetchall()
    cur.close()
    conn.close()

    target_to = "teaching-team@ccc.example.com"
    target_from = "ta@university.example.com"
    target_subj = "module review setup - creative computing and culture (fall 2014)"

    match = None
    for subj, from_addr, to_addr, body in all_emails:
        recipients = []
        if isinstance(to_addr, list):
            recipients = [str(r).strip().lower() for r in to_addr]
        elif isinstance(to_addr, str):
            try:
                p = json.loads(to_addr)
                recipients = [str(r).strip().lower() for r in p] if isinstance(p, list) else [to_addr.lower()]
            except Exception:
                recipients = [to_addr.lower()]
        if target_to in recipients and target_subj in (subj or "").lower():
            match = (subj, from_addr, to_addr, body)
            break

    check(f"Email with subject '...{target_subj[:40]}...' to {target_to}",
          match is not None,
          f"Total emails: {len(all_emails)}")
    if not match:
        return

    subj, from_addr, to_addr, body = match
    check(f"Email from {target_from}",
          target_from in (from_addr or "").lower(),
          f"From: {from_addr}")
    body_lower = (body or "").lower()
    # 5 module names
    mods_in_body = sum(1 for m in EXPECTED_MODULES if m.lower() in body_lower)
    check("Email body lists all 5 module names",
          mods_in_body == 5,
          f"Found {mods_in_body}/5 module names")
    # Notion DB link or name
    check("Email body mentions Notion DB or 'Module Tracker'",
          "module tracker" in body_lower or "notion" in body_lower or "ccc fall 2014" in body_lower,
          "Expected Notion DB reference")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    print("=" * 70)
    print("CANVAS MODULE COMPLETION NOTION GCAL - EVALUATION")
    print("=" * 70)

    check_notion()
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
