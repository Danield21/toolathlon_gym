"""
Preprocess script for sf-hr-satisfaction-excel-gcal task.

Snowflake data is read-only. This script clears Google Calendar events.
"""
import os
import argparse
import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": "toolathlon_gym",
    "user": "eigent",
    "password": "camel",
}


def clear_gcal(cur):
    print("[preprocess] Clearing Google Calendar events...")
    cur.execute("DELETE FROM gcal.events")
    print("[preprocess] Google Calendar events cleared.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        clear_gcal(cur)
        conn.commit()
        print("[preprocess] Done.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    # Clean up stale Employee_Satisfaction.xlsx in agent workspace (not GT)
    if args.agent_workspace:
        ws_abs = os.path.abspath(args.agent_workspace)
        if "groundtruth" not in ws_abs.lower():
            stale = os.path.join(args.agent_workspace, "Employee_Satisfaction.xlsx")
            if os.path.exists(stale):
                try:
                    os.remove(stale)
                    print(f"[preprocess] Removed stale {stale}")
                except Exception as e:
                    print(f"[preprocess] Could not remove: {e}")


if __name__ == "__main__":
    main()
