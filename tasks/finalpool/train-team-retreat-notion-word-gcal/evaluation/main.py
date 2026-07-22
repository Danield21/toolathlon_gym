"""
Evaluation for train-team-retreat-notion-word-gcal task.

Checks:
1. Team_Retreat_Itinerary.docx exists with 4 sections
2. Notion page "Team Retreat March 2026" exists
3. 2 GCal events (Team Retreat Departs + Returns)
4. Email sent to hr@company.com
"""
import json
import os
import sys
from argparse import ArgumentParser

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}

PASS_COUNT = 0
FAIL_COUNT = 0


def get_valid_train_codes():
    """Query the rail mock for valid train codes on the retreat dates/routes.

    Returns a dict with keys 'outbound' and 'return', each mapping to a set of
    accepted train codes (lowercased, e.g. {'g235', 'g168'}). The two telecode
    pairs (VNP=Beijing, SHH=Shanghai) and (QFB=Qufu) are static for the mock.
    Falls back to the hardcoded set if DB is unavailable.
    """
    fallback = {
        "outbound": {"g235", "g168"},
        "return": {"g236", "g167"},
    }
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute(
            "SELECT lower(station_train_code) FROM train.trains "
            "WHERE depart_date = '2026-03-12' "
            "  AND from_station_telecode IN ('VNP','SHH') "
            "  AND to_station_telecode = 'QFB'"
        )
        outbound = {r[0] for r in cur.fetchall() if r[0]}
        cur.execute(
            "SELECT lower(station_train_code) FROM train.trains "
            "WHERE depart_date = '2026-03-15' "
            "  AND from_station_telecode = 'QFB' "
            "  AND to_station_telecode IN ('VNP','SHH')"
        )
        retn = {r[0] for r in cur.fetchall() if r[0]}
        cur.close()
        conn.close()
        if outbound and retn:
            return {"outbound": outbound, "return": retn}
    except Exception:
        pass
    return fallback


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def check_word(agent_workspace):
    print("\n=== Check 1: Team_Retreat_Itinerary.docx ===")

    docx_path = os.path.join(agent_workspace, "Team_Retreat_Itinerary.docx")
    if not os.path.exists(docx_path):
        record("Team_Retreat_Itinerary.docx exists", False, f"Not found at {docx_path}")
        return
    record("Team_Retreat_Itinerary.docx exists", True)

    try:
        from docx import Document
        doc = Document(docx_path)
    except Exception as e:
        record("Word file readable", False, str(e))
        return
    record("Word file readable", True)

    full_text = "\n".join(p.text for p in doc.paragraphs).lower()

    # Check sections (require both keywords from each section title)
    has_overview = "retreat overview" in full_text
    has_outbound = "outbound journey" in full_text
    has_return = "return journey" in full_text
    has_notes = "schedule notes" in full_text
    record("Document has 'Retreat Overview' section", has_overview, "Missing exact 'Retreat Overview'")
    record("Document has 'Outbound Journey' section", has_outbound, "Missing 'Outbound Journey'")
    record("Document has 'Return Journey' section", has_return, "Missing 'Return Journey'")
    record("Document has 'Schedule Notes' section", has_notes, "Missing 'Schedule Notes'")

    # All FIVE participant names must be present
    expected_names = [
        ("Alice Chen", "alice"), ("Bob Liu", "bob"), ("Carol Wang", "carol"),
        ("David Zhang", "david"), ("Emma Li", "emma"),
    ]
    for full_name, key in expected_names:
        record(f"Document mentions {full_name}", key in full_text,
               f"Looking for '{key}'")

    # Train codes: dynamically query rail mock for valid codes on the routes.
    # Agent must mention at least one outbound code per departure-city route
    # and at least one return code per return-city route.
    codes = get_valid_train_codes()
    out_present = sum(1 for c in codes["outbound"] if c in full_text)
    ret_present = sum(1 for c in codes["return"] if c in full_text)
    record(
        f"Document mentions outbound train codes (any of {sorted(codes['outbound'])})",
        out_present >= max(1, len(codes["outbound"])),
        f"matched {out_present}/{len(codes['outbound'])} outbound codes",
    )
    record(
        f"Document mentions return train codes (any of {sorted(codes['return'])})",
        ret_present >= max(1, len(codes["return"])),
        f"matched {ret_present}/{len(codes['return'])} return codes",
    )

    # Check destination
    has_qufu = "qufu" in full_text
    record("Document mentions Qufu as destination", has_qufu, "Qufu not found")


def check_notion():
    print("\n=== Check 2: Notion Page ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    # Strict: require exact title 'Team Retreat March 2026' (case-insensitive whole phrase)
    cur.execute("""
        SELECT id, properties FROM notion.pages
        WHERE properties::text ILIKE '%Team Retreat March 2026%'
    """)
    pages = cur.fetchall()
    cur.close()
    conn.close()

    record("Notion page 'Team Retreat March 2026' exists", len(pages) >= 1,
           f"Found {len(pages)} matching pages")

    if pages:
        page_id = pages[0][0]
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute("SELECT block_data FROM notion.blocks WHERE parent_id = %s", (page_id,))
        blocks = cur.fetchall()
        cur.close()
        conn.close()

        block_text = " ".join(str(b[0]) for b in blocks).lower()
        has_content = len(blocks) > 0
        record("Notion page has content blocks", has_content, f"Found {len(blocks)} blocks")

        if has_content:
            # Tighter: require Qufu AND at least 2 valid train codes (dynamic).
            valid_codes = get_valid_train_codes()
            all_codes = valid_codes["outbound"] | valid_codes["return"]
            train_codes = sum(1 for code in all_codes if code in block_text)
            record("Notion page mentions Qufu", "qufu" in block_text, block_text[:200])
            record("Notion page mentions at least 2 train codes",
                   train_codes >= 2, f"Found {train_codes} train codes (valid set: {sorted(all_codes)}); text: {block_text[:200]}")
            # Require all 5 participant names
            for key in ("alice", "bob", "carol", "david", "emma"):
                record(f"Notion page mentions {key}", key in block_text, block_text[:200])


def check_gcal():
    print("\n=== Check 3: Calendar Events ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    cur.execute("""
        SELECT summary, start_datetime, end_datetime, description
        FROM gcal.events
        WHERE summary ILIKE '%retreat%'
        ORDER BY start_datetime
    """)
    events = cur.fetchall()
    cur.close()
    conn.close()

    record("Exactly 2 retreat calendar events", len(events) == 2,
           f"Found {len(events)} events")

    depart_events = [e for e in events if "depart" in (e[0] or "").lower()]
    return_events = [e for e in events if "return" in (e[0] or "").lower()]
    record("'Team Retreat Departs' event exists", len(depart_events) >= 1,
           f"Found: {[e[0] for e in events]}")
    record("'Team Retreat Returns' event exists", len(return_events) >= 1,
           f"Found: {[e[0] for e in events]}")

    # Validate exact start/end datetimes per task.md
    def fmt_dt(dt):
        try:
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return str(dt) if dt else ""

    if depart_events:
        ev = depart_events[0]
        summary, start_dt, end_dt, desc = ev
        record("Depart event title is exactly 'Team Retreat Departs'",
               (summary or "").strip() == "Team Retreat Departs",
               f"Got: {summary}")
        record("Depart event start = 2026-03-12 17:00:00",
               fmt_dt(start_dt) == "2026-03-12 17:00:00", f"Got: {fmt_dt(start_dt)}")
        record("Depart event end = 2026-03-12 20:00:00",
               fmt_dt(end_dt) == "2026-03-12 20:00:00", f"Got: {fmt_dt(end_dt)}")
        desc_low = (desc or "").lower()
        # Description must mention all FIVE team members
        for key in ("alice", "bob", "carol", "david", "emma"):
            record(f"Depart description mentions {key}",
                   key in desc_low, f"Description: {desc_low[:200]}")

    if return_events:
        ev = return_events[0]
        summary, start_dt, end_dt, desc = ev
        record("Return event title is exactly 'Team Retreat Returns'",
               (summary or "").strip() == "Team Retreat Returns",
               f"Got: {summary}")
        record("Return event start = 2026-03-15 14:00:00",
               fmt_dt(start_dt) == "2026-03-15 14:00:00", f"Got: {fmt_dt(start_dt)}")
        record("Return event end = 2026-03-15 18:00:00",
               fmt_dt(end_dt) == "2026-03-15 18:00:00", f"Got: {fmt_dt(end_dt)}")
        desc_low = (desc or "").lower()
        for key in ("alice", "bob", "carol", "david", "emma"):
            record(f"Return description mentions {key}",
                   key in desc_low, f"Description: {desc_low[:200]}")


def check_email():
    print("\n=== Check 4: Email to HR ===")

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()
    # Strict: must be sent to hr@company.com AND have exact subject
    cur.execute("""
        SELECT to_addr, subject, body_text FROM email.messages
        WHERE to_addr::text ILIKE '%hr@company.com%'
    """)
    recipient_msgs = cur.fetchall()
    cur.close()
    conn.close()

    record("Email sent to hr@company.com",
           len(recipient_msgs) >= 1,
           f"Found {len(recipient_msgs)} emails to hr@company.com")

    target = None
    for to_addr, subject, body in recipient_msgs:
        if (subject or "").strip() == "Team Retreat Travel Confirmed":
            target = (to_addr, subject, body)
            break
    record("Email subject is exactly 'Team Retreat Travel Confirmed'",
           target is not None,
           f"Subjects: {[m[1] for m in recipient_msgs]}")

    if target:
        to_addr, subject, body = target
        body_low = (body or "").lower()
        # Body must mention all 5 participants and at least 2 train codes
        for key in ("alice", "bob", "carol", "david", "emma"):
            record(f"Email body mentions {key}", key in body_low, f"Body: {body_low[:200]}")
        valid_codes = get_valid_train_codes()
        all_codes = valid_codes["outbound"] | valid_codes["return"]
        train_codes = sum(1 for code in all_codes if code in body_low)
        record("Email body mentions at least 2 train codes",
               train_codes >= 2,
               f"Found {train_codes} (valid set: {sorted(all_codes)}); body: {body_low[:200]}")


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    check_word(args.agent_workspace)
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

    if FAIL_COUNT == 0:
        print("PASS")
        sys.exit(0)
    else:
        print(f"FAIL: {FAIL_COUNT} check(s) failed")
        sys.exit(1)


if __name__ == "__main__":
    main()
