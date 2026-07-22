"""Preprocess for terminal-howtocook-pdf-excel-word-gcal.
HowToCook is MCP-only (no DB). Clear gcal and inject noise events."""
import argparse
import glob
import json
import os
from datetime import datetime, timedelta

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent",
    "password": "camel",
}


def clear_gcal(cur):
    print("[preprocess] Clearing Google Calendar events...")
    cur.execute("DELETE FROM gcal.events")


def inject_noise(cur, launch_time):
    """Inject noise events on dates that do NOT overlap with target week (March 9-13, 2026)."""
    print("[preprocess] Injecting noise calendar events...")
    # Target week is March 9-13, 2026 (festival prep). Place noise outside this week.
    # Use fixed dates: week before (March 2-6) and week after (March 16-20).
    target_year = 2026

    noise_events = [
        # Week before target
        ("Weekly Team Standup", datetime(target_year, 3, 2, 9, 0),
         datetime(target_year, 3, 2, 9, 30), "Regular team meeting"),
        ("Budget Review Meeting", datetime(target_year, 3, 3, 14, 0),
         datetime(target_year, 3, 3, 15, 0), "Q1 budget review"),
        ("Facilities Maintenance", datetime(target_year, 3, 4, 8, 0),
         datetime(target_year, 3, 4, 12, 0), "Building maintenance check"),
        # Week after target
        ("Vendor Contract Review", datetime(target_year, 3, 17, 10, 0),
         datetime(target_year, 3, 17, 11, 0), "Catering vendor contract review"),
        ("Post-Event Debrief", datetime(target_year, 3, 18, 15, 0),
         datetime(target_year, 3, 18, 16, 0), "Post-festival retrospective"),
    ]

    for summary, start, end, desc in noise_events:
        cur.execute("""
            INSERT INTO gcal.events (summary, start_datetime, end_datetime, description, status)
            VALUES (%s, %s, %s, %s, 'confirmed')
        """, (summary, start, end, desc))
    print(f"[preprocess] Injected {len(noise_events)} noise events.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_gcal(cur)
        inject_noise(cur, args.launch_time)
        conn.commit()
        print("[preprocess] DB setup done.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    if args.agent_workspace:
        for pattern in ["Campus_Dining_Plan.xlsx", "Meal_Plan_Report.docx",
                        "budget_*.py", "budget_*.json", "selected_*.json"]:
            for f in glob.glob(os.path.join(args.agent_workspace, pattern)):
                os.remove(f)
                print(f"[preprocess] Removed {f}")

    print("[preprocess] Done.")


if __name__ == "__main__":
    main()
