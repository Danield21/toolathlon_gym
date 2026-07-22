"""Preprocess script for pw-notion-gform-survey-excel-word."""
import os
import argparse, json, os, sys, shutil, tarfile, subprocess, time
from datetime import datetime, timedelta

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"), "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent", "password": "camel"
}

TASK_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def get_conn():
    import psycopg2
    return psycopg2.connect(**DB_CONFIG)

def clear_writable_schemas():
    conn = get_conn()
    cur = conn.cursor()
    
    cur.execute("DELETE FROM notion.comments")
    cur.execute("DELETE FROM notion.blocks")
    cur.execute("DELETE FROM notion.pages")
    cur.execute("DELETE FROM notion.databases")

    cur.execute("DELETE FROM gform.responses")
    cur.execute("DELETE FROM gform.questions")
    cur.execute("DELETE FROM gform.forms")
    conn.commit()
    cur.close()
    conn.close()

def inject_data(launch_time):
    conn = get_conn()
    cur = conn.cursor()
    launch_dt = datetime.strptime(launch_time, "%Y-%m-%d %H:%M:%S")

    cur.execute("""INSERT INTO notion.pages (id, parent, properties, archived, in_trash, url) VALUES
        ('noise-notion_survey-001', '{"type": "workspace", "workspace": true}'::jsonb,
         '{"title": {"id": "title", "type": "title", "title": [{"type": "text", "text": {"content": "Old Project Notes"}}]}}'::jsonb,
         false, false, 'https://notion.so/old-notes')""")

    # Inject the page that holds internal survey metrics — these are the
    # Internal_Value figures the agent must read and compare against the mock
    # external benchmarks. Categories match those exposed by the mock dashboard.
    cur.execute(
        """INSERT INTO notion.pages (id, parent, properties, archived, in_trash, url) VALUES
        ('internal-survey-metrics-001', '{"type": "workspace", "workspace": true}'::jsonb,
         '{"title": {"id": "title", "type": "title", "title": [{"type": "text", "text": {"content": "Internal Survey Metrics"}}]}}'::jsonb,
         false, false, 'https://notion.so/internal-survey')"""
    )
    # Internal duration values per category (in minutes), to be compared
    # against external Typical_Duration_Min benchmarks.
    internal_rows = [
        ("Strategy Review", 105),
        ("Team Standup", 20),
        ("1-on-1", 25),
        ("All-Hands", 75),
        ("Training", 110),
    ]
    pos = 0
    for cat, val in internal_rows:
        block_id = f"internal-block-{pos}"
        block_text = f"{cat}: {val} minutes"
        cur.execute(
            """INSERT INTO notion.blocks
               (id, object, parent_type, parent_id, type, has_children, archived, in_trash, block_data, position)
               VALUES (%s, 'block', 'page_id', 'internal-survey-metrics-001', 'paragraph', false, false, false, %s::jsonb, %s)""",
            (
                block_id,
                json.dumps({"paragraph": {"rich_text": [{"type": "text", "text": {"content": block_text}, "plain_text": block_text}]}}),
                pos,
            ),
        )
        pos += 1
    conn.commit()
    cur.close()
    conn.close()


def setup_mock_server(port=30330):
    files_dir = os.path.join(TASK_ROOT, "files")
    tmp_dir = os.path.join(TASK_ROOT, "tmp")
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    # Kill existing process on port
    try:
        subprocess.run(f"kill -9 $(lsof -ti:30330) 2>/dev/null", shell=True, timeout=5)
    except Exception:
        pass
    time.sleep(0.5)

    # Extract mock pages
    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    if os.path.exists(tar_path):
        with tarfile.open(tar_path, "r:gz") as tar:
            try:
                tar.extractall(path=tmp_dir, filter="data")
            except TypeError:
                tar.extractall(path=tmp_dir)

    # Start HTTP server
    mock_dir = os.path.join(tmp_dir, "mock_pages")
    if os.path.exists(mock_dir):
        log_path = os.path.join(mock_dir, "server.log")
        subprocess.Popen(
            f"nohup python3 -m http.server 30330 --directory {mock_dir} > {log_path} 2>&1 &",
            shell=True
        )
        time.sleep(1)
        print(f"Mock server started on port 30330")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    args = parser.parse_args()

    clear_writable_schemas()
    inject_data(args.launch_time)
    setup_mock_server(30330)

if __name__ == "__main__":
    main()
