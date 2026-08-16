"""
Preprocess script for canvas-faculty-workload-review task.
Clears gsheet data, starts mock HTTP server on port 30220.
Canvas is read-only.
"""
import argparse
import asyncio
import json
import os
import subprocess
import sys
import time
from pathlib import Path
from urllib.request import urlopen

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent",
    "password": "camel",
}

PORT = 30220


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        cur.execute("DELETE FROM gsheet.cells")
        cur.execute("DELETE FROM gsheet.permissions")
        cur.execute("DELETE FROM gsheet.sheets")
        cur.execute("DELETE FROM gsheet.spreadsheets")
        cur.execute("DELETE FROM gsheet.folders")
        print("[preprocess] Cleared Google Sheet data.")

        conn.commit()
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Database error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    # Serve the mock Canvas API from tmp/mock_pages (Bug E in audit §A.8: a
    # prior commit pointed this at files/, which does not exist, so the HTTP
    # server never started and the agent saw 404s for workload_standards.json).
    task_root = Path(__file__).resolve().parent.parent
    serve_dir = task_root / "tmp" / "mock_pages"
    standards_path = serve_dir / "api" / "workload_standards.json"
    if not standards_path.is_file():
        raise FileNotFoundError(f"Missing workload standards: {standards_path}")

    server = subprocess.Popen(
        [
            sys.executable,
            "-m",
            "http.server",
            str(PORT),
            "--bind",
            "127.0.0.1",
            "--directory",
            str(serve_dir),
        ],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    deadline = time.monotonic() + 5
    url = f"http://127.0.0.1:{PORT}/api/workload_standards.json"
    while True:
        if server.poll() is not None:
            raise RuntimeError("Workload standards server exited before becoming ready")
        try:
            with urlopen(url, timeout=0.5) as response:
                observed = json.load(response)
            if observed["grading_hours_per_student"] != 0.5:
                raise RuntimeError("Workload standards server returned unexpected data")
            break
        except (OSError, KeyError, json.JSONDecodeError):
            if time.monotonic() >= deadline:
                server.terminate()
                raise RuntimeError("Workload standards server did not become ready")
            await asyncio.sleep(0.1)

    print(f"[preprocess] Mock server running at http://localhost:{PORT}")
    print("[preprocess] Done.")


if __name__ == "__main__":
    asyncio.run(main())
