"""Preprocess for sf-hr-attrition-gcal. Clears gcal and email data."""
import os
import argparse
import psycopg2

DB = dict(host=os.environ.get("PGHOST", "localhost"), port=int(os.environ.get("PGPORT", "5432")), dbname="toolathlon_gym", user="eigent", password="camel")

def clean_workspace(agent_workspace: str):
    if not agent_workspace:
        return
    stale = os.path.join(agent_workspace, "Satisfaction_Report.xlsx")
    if os.path.isfile(stale):
        try:
            os.remove(stale)
            print(f"[preprocess] Removed stale {stale}")
        except OSError as e:
            print(f"[preprocess] Warning: could not remove {stale}: {e}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    clean_workspace(args.agent_workspace)

    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    try:
        cur.execute("DELETE FROM gcal.events")
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        conn.commit()
        print("[preprocess] Cleared gcal events and email data.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

if __name__ == "__main__":
    main()
