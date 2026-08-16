"""
Evaluation script for fetch-arxiv-conference-schedule-gcal-notion task.

Checks:
1. Notion database "Conference Reading List" with 6+ entries
2. Google Calendar events for 3 reading group sessions
"""

import argparse
import json
import os
import sys

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
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


def str_contains(haystack, needle):
    if haystack is None or needle is None:
        return False
    return needle.strip().lower() in str(haystack).strip().lower()


def check_notion():
    """Check Notion database for Conference Reading List."""
    print("\n=== Checking Notion Database ===")

    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()

        # Find databases
        cur.execute("SELECT id, title, properties FROM notion.databases")
        dbs = cur.fetchall()

        found_db = None
        for db_id, title_raw, props in dbs:
            # title can be jsonb array or string
            if isinstance(title_raw, list):
                title = " ".join(t.get("plain_text", "") for t in title_raw if isinstance(t, dict))
            elif isinstance(title_raw, str):
                try:
                    parsed = json.loads(title_raw)
                    if isinstance(parsed, list):
                        title = " ".join(t.get("plain_text", "") for t in parsed if isinstance(t, dict))
                    else:
                        title = title_raw
                except (json.JSONDecodeError, TypeError):
                    title = title_raw
            else:
                title = str(title_raw) if title_raw else ""
            title_lower = title.lower()
            if "conference" in title_lower or "reading" in title_lower:
                found_db = db_id
                break

        if not found_db:
            record(
                "Conference Reading List database exists",
                False,
                f"Found {len(dbs)} databases: {[d[1] for d in dbs]}",
            )
            cur.close()
            conn.close()
            return False

        record("Conference Reading List database exists", True)

        # Check database properties - require ALL 6 properties from task.md
        if props:
            if isinstance(props, str):
                props = json.loads(props)
            prop_names = [k.lower() for k in props.keys()] if isinstance(props, dict) else []
            required_props = {
                "title": any("title" in p for p in prop_names),
                "authors": any("author" in p for p in prop_names),
                "source": any("source" in p for p in prop_names),
                "conference_session": any("session" in p for p in prop_names),
                "relevance_score": any("relevance" in p for p in prop_names),
                "read_by_date": any("read" in p and "date" in p for p in prop_names) or any("read_by" in p for p in prop_names),
            }
            for name, present in required_props.items():
                record(f"Database has '{name}' property", present, f"Properties: {prop_names}")
        else:
            record("Database has properties", False, "No properties found")

        # Check pages (entries) in the database
        cur.execute(
            "SELECT id, properties FROM notion.pages WHERE parent::text LIKE %s",
            (f'%{found_db}%',),
        )
        pages = cur.fetchall()

        record(
            "Database has >= 6 entries",
            len(pages) >= 6,
            f"Found {len(pages)} entries",
        )

        # Check that entries have diverse sources
        sources = set()
        for page_id, page_props in pages:
            if isinstance(page_props, str):
                page_props = json.loads(page_props)
            if isinstance(page_props, dict):
                for key, val in page_props.items():
                    if "source" in key.lower():
                        if isinstance(val, dict):
                            select_val = val.get("select", {})
                            if isinstance(select_val, dict):
                                name = select_val.get("name", "")
                                if name:
                                    sources.add(name.lower())

        # Require BOTH ArXiv and Scholar sources, not OR with page count
        has_arxiv = any("arxiv" in s for s in sources)
        has_scholar = any("scholar" in s for s in sources)
        record(
            "Entries include ArXiv source",
            has_arxiv,
            f"Sources found: {sources}",
        )
        record(
            "Entries include Scholar source",
            has_scholar,
            f"Sources found: {sources}",
        )

        # Check relevance scores - require all entries to have a numeric score 1-5
        # Also build per-page (relevance, read_by_date, title) records for downstream checks
        scores_valid = 0
        per_page_records = []  # list of (relevance:int|None, read_by:str|None, title:str)
        for page_id, page_props in pages:
            if isinstance(page_props, str):
                page_props = json.loads(page_props)

            page_relevance = None
            page_read_by = None
            page_title = ""
            if isinstance(page_props, dict):
                for key, val in page_props.items():
                    klow = key.lower()
                    if "relevance" in klow:
                        try:
                            num = None
                            if isinstance(val, dict):
                                num = val.get("number")
                                if num is None and isinstance(val.get("formula"), dict):
                                    num = val["formula"].get("number")
                            elif isinstance(val, (int, float)):
                                num = val
                            if num is not None and 1 <= float(num) <= 5:
                                scores_valid += 1
                                page_relevance = float(num)
                        except (TypeError, ValueError):
                            pass
                    if "read" in klow and "date" in klow:
                        try:
                            if isinstance(val, dict):
                                date_val = val.get("date")
                                if isinstance(date_val, dict):
                                    page_read_by = date_val.get("start")
                                elif isinstance(date_val, str):
                                    page_read_by = date_val
                                elif isinstance(val.get("formula"), dict):
                                    page_read_by = val["formula"].get("date")
                        except (TypeError, ValueError):
                            pass
                    if klow == "title" or "title" == klow.strip():
                        try:
                            if isinstance(val, dict):
                                title_arr = val.get("title", [])
                                if isinstance(title_arr, list):
                                    page_title = " ".join(
                                        t.get("plain_text", "") if isinstance(t, dict) else ""
                                        for t in title_arr
                                    )
                        except (TypeError, ValueError):
                            pass
            per_page_records.append((page_relevance, page_read_by, page_title))

        record(
            "All entries have relevance scores 1-5",
            scores_valid == len(pages) and len(pages) >= 6,
            f"valid={scores_valid}, total={len(pages)}",
        )

        # Read_By_Date conditional validation:
        #   - high relevance (4-5) -> 2026-03-20
        #   - lower relevance (1-3) -> 2026-03-23
        read_by_correct = 0
        read_by_total = 0
        for relevance, read_by, _title in per_page_records:
            if relevance is None or read_by is None:
                continue
            read_by_total += 1
            read_by_str = str(read_by)
            if relevance >= 4 and "2026-03-20" in read_by_str:
                read_by_correct += 1
            elif relevance <= 3 and "2026-03-23" in read_by_str:
                read_by_correct += 1
        record(
            "Read_By_Date follows relevance rule (>=4 -> 03-20, <=3 -> 03-23)",
            read_by_total >= 6 and read_by_correct == read_by_total,
            f"correct={read_by_correct}/{read_by_total}",
        )

        # Noise rejection: papers about protein folding, manifolds, climate change
        # should NOT appear in the reading list (they are noise injected in preprocess).
        noise_keywords = [
            "protein folding",
            "topological invariant",
            "four-dimensional manifold",
            "climate change",
            "coastal erosion",
        ]
        noise_titles_found = []
        for _r, _rb, title in per_page_records:
            tl = (title or "").lower()
            for nk in noise_keywords:
                if nk in tl:
                    noise_titles_found.append(title)
                    break
        record(
            "Noise (irrelevant) papers excluded from reading list",
            len(noise_titles_found) == 0,
            f"found noise titles: {noise_titles_found}",
        )

        cur.close()
        conn.close()
        return True

    except Exception as e:
        record("Notion DB accessible", False, str(e))
        return False


def check_calendar():
    """Check Google Calendar for reading group sessions."""
    print("\n=== Checking Google Calendar ===")

    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute(
            "SELECT summary, description, start_datetime, end_datetime FROM gcal.events"
        )
        events = cur.fetchall()
        cur.close()
        conn.close()
    except Exception as e:
        record("Calendar DB accessible", False, str(e))
        return False

    all_ok = True

    # Check for 3 reading group sessions
    reading_events = [
        e for e in events
        if "reading group" in (e[0] or "").lower()
    ]
    record(
        "3 reading group events exist",
        len(reading_events) >= 3,
        f"Found {len(reading_events)} reading group events",
    )
    if len(reading_events) < 3:
        all_ok = False

    # Check specific topics
    topics_found = {
        "transformer": False,
        "attention": False,
        "optim": False,
    }
    for summary, description, start_dt, end_dt in reading_events:
        summary_lower = (summary or "").lower()
        for topic in topics_found:
            if topic in summary_lower:
                topics_found[topic] = True

    for topic, found in topics_found.items():
        record(f"Reading group for '{topic}' topic exists", found)
        if not found:
            all_ok = False

    # Check exact dates and times
    # March 18 14:00-15:30: Transformer
    # March 19 14:00-15:30: Attention
    # March 20 14:00-15:30: Optimization
    expected_sessions = [
        ("transformer", "2026-03-18", "14:00", "15:30"),
        ("attention", "2026-03-19", "14:00", "15:30"),
        ("optim", "2026-03-20", "14:00", "15:30"),
    ]
    for keyword, exp_date, exp_start_hm, exp_end_hm in expected_sessions:
        match_date = False
        match_start_time = False
        match_end_time = False
        for summary, description, start_dt, end_dt in reading_events:
            if keyword in (summary or "").lower():
                sd = str(start_dt) if start_dt else ""
                ed = str(end_dt) if end_dt else ""
                if exp_date in sd:
                    match_date = True
                if exp_start_hm in sd:
                    match_start_time = True
                if exp_end_hm in ed:
                    match_end_time = True
        record(f"'{keyword}' event date is {exp_date}", match_date)
        record(f"'{keyword}' event starts at {exp_start_hm}", match_start_time)
        record(f"'{keyword}' event ends at {exp_end_hm}", match_end_time)
        if not (match_date and match_start_time and match_end_time):
            all_ok = False

    # Each reading group event description must have non-trivial content
    # listing paper titles (substantive content > 50 chars)
    for keyword, exp_date, _, _ in expected_sessions:
        for summary, description, start_dt, end_dt in reading_events:
            if keyword in (summary or "").lower():
                desc = description or ""
                record(f"'{keyword}' event description lists papers (>50 chars)",
                       len(desc) > 50, f"len={len(desc)}")
                break
        else:
            record(f"'{keyword}' event description present", False, "no event found")

    return all_ok


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    notion_ok = check_notion()
    cal_ok = check_calendar()

    print(f"\n=== SUMMARY ===")
    print(f"  Notion:   {'PASS' if notion_ok else 'FAIL'}")
    print(f"  Calendar: {'PASS' if cal_ok else 'FAIL'}")
    print(f"  Passed: {PASS_COUNT}, Failed: {FAIL_COUNT}")

    overall = notion_ok and cal_ok and FAIL_COUNT == 0
    print(f"  Overall:  {'PASS' if overall else 'FAIL'}")

    sys.exit(0 if overall else 1)


if __name__ == "__main__":
    main()
