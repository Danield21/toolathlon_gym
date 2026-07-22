#!/usr/bin/env python3
"""Preprocess script for task setup.

Cleans stale email/gcal rows that match the verifier ILIKE filters,
extracts mock_pages.tar.gz, and starts a local HTTP server so the
playwright agent has stable competitor URLs to scrape rather than
relying on live web pages.
"""

from argparse import ArgumentParser
import os
import shutil
import subprocess
import tarfile
import time
from pathlib import Path

TASK_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MOCK_PORT = 30350


def cleanup_writable_schemas():
    """Clean stale data from writable schemas matching task ILIKE patterns.

    Verifier queries:
      - email.messages subject/body ILIKE %competitive%/%competitor%/%market%/%intelligence%
      - gcal.events summary ILIKE %briefing%/%strategy%/%competitive%
    Stale data from prior tasks could falsely pass these checks.
    """
    try:
        import psycopg2
        conn = psycopg2.connect(
            host=os.environ.get("PGHOST", "localhost"), port=5432,
            dbname=os.environ.get("PGDATABASE", "toolathlon_gym"),
            user="eigent", password="camel",
        )
        cur = conn.cursor()
        cur.execute(
            """DELETE FROM gcal.events
               WHERE summary ILIKE %s OR summary ILIKE %s OR summary ILIKE %s""",
            ('%briefing%', '%strategy%', '%competitive%'))
        cur.execute(
            """DELETE FROM email.messages
               WHERE subject ILIKE %s OR subject ILIKE %s OR subject ILIKE %s
                  OR body_text ILIKE %s""",
            ('%competitive%', '%competitor%', '%market%', '%intelligence%'))
        conn.commit()
        cur.close()
        conn.close()
        print("Cleaned stale competitive/market/intelligence data from gcal.events + email.messages")
    except Exception as e:
        print(f"Warning: cleanup_writable_schemas skipped: {e}")


def setup_mock_server(port=MOCK_PORT):
    """Extract mock_pages.tar.gz and start a local HTTP server.

    Provides reproducible competitor profile pages for the playwright
    browsing phase so runs are not dependent on live external sites.
    """
    files_dir = os.path.join(TASK_ROOT, "files")
    tmp_dir = os.path.join(TASK_ROOT, "tmp")
    if os.path.exists(tmp_dir):
        shutil.rmtree(tmp_dir)
    os.makedirs(tmp_dir, exist_ok=True)

    # Kill any existing process on the port
    try:
        subprocess.run(
            f"kill -9 $(lsof -ti:{port}) 2>/dev/null",
            shell=True, timeout=5)
    except Exception:
        pass
    time.sleep(0.5)

    # Extract mock pages
    tar_path = os.path.join(files_dir, "mock_pages.tar.gz")
    if not os.path.exists(tar_path):
        print(f"Warning: {tar_path} missing; mock server skipped")
        return
    with tarfile.open(tar_path, "r:gz") as tar:
        tar.extractall(path=tmp_dir)

    mock_dir = os.path.join(tmp_dir, "mock_pages")
    if not os.path.exists(mock_dir):
        print(f"Warning: extracted mock_pages dir not found")
        return

    log_path = os.path.join(mock_dir, "server.log")
    subprocess.Popen(
        f"nohup python3 -m http.server {port} --directory {mock_dir} > {log_path} 2>&1 &",
        shell=True,
    )
    time.sleep(1)
    print(f"Mock competitor pages server started on port {port}")


def main():
    parser = ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    # Clean writable schemas to avoid stale data false positives
    cleanup_writable_schemas()

    # Start mock competitor pages HTTP server
    setup_mock_server()

    # Ensure agent workspace exists
    if args.agent_workspace:
        agent_ws = Path(args.agent_workspace)
        agent_ws.mkdir(parents=True, exist_ok=True)

    print("Preprocess completed successfully")


if __name__ == "__main__":
    main()
