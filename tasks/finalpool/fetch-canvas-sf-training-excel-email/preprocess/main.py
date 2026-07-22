"""Preprocess script for fetch-canvas-sf-training-excel-email."""
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
    
    cur.execute("DELETE FROM email.attachments")
    cur.execute("DELETE FROM email.sent_log")
    cur.execute("DELETE FROM email.messages")
    conn.commit()
    cur.close()
    conn.close()

def inject_data(launch_time):
    conn = get_conn()
    cur = conn.cursor()
    launch_dt = datetime.strptime(launch_time, "%Y-%m-%d %H:%M:%S")
    
    inbox_id = 1
    cur.execute("SELECT id FROM email.folders WHERE name = 'INBOX' LIMIT 1")
    row = cur.fetchone()
    if row: inbox_id = row[0]
    cur.execute("""INSERT INTO email.messages (folder_id, message_id, subject, from_addr, to_addr, date, body_text, is_read)
        VALUES (%s, %s, %s, %s, %s, %s, %s, true)""",
        (inbox_id, '<noise-canvas_sf_training-001@co.com>', 'Weekly Newsletter', 'newsletter@company.com',
         json.dumps(['all@company.com']), launch_dt - timedelta(hours=6),
         'This week in company news: New coffee machine installed in break room.'))
    conn.commit()
    cur.close()
    conn.close()


def setup_mock_server(port=30327):
    files_dir = os.path.join(TASK_ROOT, "files")
    tmp_dir = os.path.join(TASK_ROOT, "tmp")
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    # Kill existing process on port
    try:
        subprocess.run(f"kill -9 $(lsof -ti:30327) 2>/dev/null", shell=True, timeout=5)
    except Exception:
        pass
    time.sleep(0.5)

    # Extract mock pages
    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    if os.path.exists(tar_path):
        with tarfile.open(tar_path, "r:gz") as tar:
            tar.extractall(path=tmp_dir)

    mock_dir = os.path.join(tmp_dir, "mock_pages")
    # Overwrite mock JSON to expose course-level training benchmarks matching GT codes
    if os.path.exists(mock_dir):
        api_dir = os.path.join(mock_dir, "api")
        os.makedirs(api_dir, exist_ok=True)
        data_json_path = os.path.join(api_dir, "data.json")
        # Field names align with the verifier's expected output columns
        # (Pass_Rate, Avg_Score, Enrollment) so the agent can map directly
        # without renaming. Both `pass_rate` and a `completion_rate` alias
        # are emitted to keep the API self-documenting.
        mock_data = {
            "training_benchmarks": [
                {"course_code": "AAA-2013J", "course_name": "Applied Analytics & Algorithms (Fall 2013)",
                 "pass_rate": 78.5, "completion_rate": 78.5, "avg_score": 80.0, "enrollment": 1993},
                {"course_code": "AAA-2014J", "course_name": "Applied Analytics & Algorithms (Fall 2014)",
                 "pass_rate": 71.2, "completion_rate": 71.2, "avg_score": 75.0, "enrollment": 542},
                {"course_code": "BBB-2013J", "course_name": "Biochemistry & Bioinformatics (Fall 2013)",
                 "pass_rate": 75.0, "completion_rate": 75.0, "avg_score": 80.0, "enrollment": 725},
                {"course_code": "BBB-2014J", "course_name": "Biochemistry & Bioinformatics (Fall 2014)",
                 "pass_rate": 70.0, "completion_rate": 70.0, "avg_score": 72.0, "enrollment": 2024},
                {"course_code": "BBB-2013B", "course_name": "Biochemistry & Bioinformatics (Spring 2013)",
                 "pass_rate": 68.0, "completion_rate": 68.0, "avg_score": 70.0, "enrollment": 1920},
                {"course_code": "BBB-2014B", "course_name": "Biochemistry & Bioinformatics (Spring 2014)",
                 "pass_rate": 73.0, "completion_rate": 73.0, "avg_score": 75.0, "enrollment": 1447},
            ],
            "source": "Training Analytics API"
        }
        with open(data_json_path, "w") as f:
            json.dump(mock_data, f, indent=2)

        log_path = os.path.join(mock_dir, "server.log")
        subprocess.Popen(
            f"nohup python3 -m http.server 30327 --directory {mock_dir} > {log_path} 2>&1 &",
            shell=True
        )
        time.sleep(1)
        print(f"Mock server started on port 30327")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False, default="2026-03-07 10:00:00")
    args = parser.parse_args()

    clear_writable_schemas()
    inject_data(args.launch_time)
    setup_mock_server(30327)

if __name__ == "__main__":
    main()
