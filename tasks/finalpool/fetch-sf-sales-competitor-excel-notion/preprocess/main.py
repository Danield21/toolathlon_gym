"""Preprocess script for fetch-sf-sales-competitor-excel-notion."""
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
    conn.commit()
    cur.close()
    conn.close()

def inject_data(launch_time):
    conn = get_conn()
    cur = conn.cursor()
    launch_dt = datetime.strptime(launch_time, "%Y-%m-%d %H:%M:%S")
    
    cur.execute("""INSERT INTO notion.pages (id, parent, properties, archived, in_trash, url) VALUES
        ('noise-sf_competitor-001', '{"type": "workspace", "workspace": true}'::jsonb,
         '{"title": {"id": "title", "type": "title", "title": [{"type": "text", "text": {"content": "Old Project Notes"}}]}}'::jsonb,
         false, false, 'https://notion.so/old-notes')""")
    conn.commit()
    cur.close()
    conn.close()


def setup_mock_server(port=30335):
    files_dir = os.path.join(TASK_ROOT, "files")
    tmp_dir = os.path.join(TASK_ROOT, "tmp")
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    # Kill existing process on port
    try:
        subprocess.run(f"kill -9 $(lsof -ti:30335) 2>/dev/null", shell=True, timeout=5)
    except Exception:
        pass
    time.sleep(0.5)

    # Extract mock pages
    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    if os.path.exists(tar_path):
        with tarfile.open(tar_path, "r:gz") as tar:
            tar.extractall(path=tmp_dir)

    # CRITICAL FIX: Replace HR-flavored mock data with region-flavored benchmarks
    # so the agent has the external Market_Size_M data it needs to compute
    # Market_Penetration_Pct against internal sf_data SALES_DW.CUSTOMERS regions.
    mock_dir = os.path.join(tmp_dir, "mock_pages")
    api_dir = os.path.join(mock_dir, "api")
    os.makedirs(api_dir, exist_ok=True)
    region_data = {
        "benchmarks": [
            {"region": "Asia Pacific", "market_size_m": 439},
            {"region": "Europe", "market_size_m": 339},
            {"region": "Latin America", "market_size_m": 453},
            {"region": "Middle East", "market_size_m": 421},
            {"region": "North America", "market_size_m": 463},
        ],
        "source": "Sales Region Benchmarks API",
        "date": "2026-03-01",
        "notes": "Market_Size_M is regional addressable market in millions USD.",
    }
    with open(os.path.join(api_dir, "data.json"), "w") as f:
        json.dump(region_data, f, indent=2)

    # Start HTTP server
    if os.path.exists(mock_dir):
        log_path = os.path.join(mock_dir, "server.log")
        subprocess.Popen(
            f"nohup python3 -m http.server 30335 --directory {mock_dir} > {log_path} 2>&1 &",
            shell=True
        )
        time.sleep(1)
        print(f"Mock server started on port 30335")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    args = parser.parse_args()

    clear_writable_schemas()
    inject_data(args.launch_time)
    setup_mock_server(30335)

if __name__ == "__main__":
    main()
