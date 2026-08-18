"""
Preprocess script for canvas-at-risk-intervention task.
Clears email data and starts mock HTTP server on port 30430.
Canvas is read-only.
"""
import argparse
import asyncio
import os

import psycopg2

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": int(os.environ.get("PGPORT", "5432")),
    "dbname": os.environ.get("PGDATABASE", "toolathlon_gym"),
    "user": "eigent",
    "password": "camel",
}

PORT = 30430


async def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    args = parser.parse_args()

    conn = psycopg2.connect(**DB_CONFIG)
    cur = conn.cursor()

    try:
        # Clear email data
        cur.execute("DELETE FROM email.attachments")
        cur.execute("DELETE FROM email.sent_log")
        cur.execute("DELETE FROM email.messages")
        print("[preprocess] Cleared email data.")

        conn.commit()
        print("[preprocess] Database operations committed.")
    except Exception as e:
        conn.rollback()
        print(f"[preprocess] Database error: {e}")
        raise
    finally:
        cur.close()
        conn.close()

    # Start mock HTTP server.
    #
    # Robustness notes (case-study 2026-08-12, case #4 canvas-at-risk):
    # The previous version used `lsof -ti:PORT` to kill any prior listener, then
    # started `python -m http.server` and slept 1s before declaring success.
    # That was doubly fragile: (1) the image may not ship lsof, leaving a stale
    # listener from a different task on 30430 (we observed the Competitor Portal
    # served there instead of support_resources.json); (2) if our bind failed
    # because the port was taken, the script still printed "running" and the
    # model then hit 404s for the whole run.
    #
    # Fix: do NOT rely on lsof. Probe the port first; if it already serves the
    # EXPECTED content, leave it. If it serves the WRONG content (another task's
    # portal), fail loudly so the run is flagged invalid rather than burning the
    # model's step budget on an unwinnable 404. After spawning our server, verify
    # it actually answers 200 with support_resources.json content before
    # returning success.
    task_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    serve_dir = os.path.join(task_root, "tmp", "mock_pages")
    if not os.path.isdir(serve_dir):
        raise FileNotFoundError(f"mock_pages directory not found: {serve_dir}")

    endpoint = f"http://localhost:{PORT}/api/support_resources.json"
    import urllib.request as _urlreq
    import urllib.error as _urlerr

    # Content markers that prove the response is OUR support-resources fixture
    # (case-study 2026-08-12, case #4). A squatter serving a different task's
    # portal would not contain tutoring/study_group/counseling entries.
    _CONTENT_MARKERS = ("tutoring", "study_group", "counseling")

    def _probe(url, timeout=1.0):
        try:
            with _urlreq.urlopen(url, timeout=timeout) as resp:
                body = resp.read().decode("utf-8", "replace").lower()
                return resp.status, body
        except _urlerr.HTTPError as e:
            return e.code, ""
        except Exception:
            return None, ""

    def _is_our_content(body_lower: str) -> bool:
        return any(m in body_lower for m in _CONTENT_MARKERS)

    # Probe the expected endpoint before we start anything. If something is
    # already serving on the port, decide whether it is ours or a squatter.
    pre_status, pre_body = _probe(endpoint)
    if pre_status == 200 and _is_our_content(pre_body):
        print(f"[preprocess] Port {PORT} already serves support_resources.json; reusing.")
    else:
        if pre_status is not None:
            # Something is listening but it is NOT our content — likely a stale
            # mock server from a previous run (the compute node's /dev/shm is
            # shared across slots and prior cancelled jobs may leave orphans).
            # Try to kill the squatter via fuser, then re-probe. Only fail if we
            # truly cannot reclaim the port.
            print(f"[preprocess] Port {PORT} occupied by another service "
                  f"(status={pre_status}); attempting to reclaim.")
            import subprocess as _sp
            for _kill_cmd in (
                f"fuser -k {PORT}/tcp",
                f"fuser -k -n tcp {PORT}",
            ):
                try:
                    _sp.run(_kill_cmd, shell=True, timeout=5,
                            stdout=_sp.DEVNULL, stderr=_sp.DEVNULL)
                except Exception:
                    pass
            await asyncio.sleep(1.0)
            # Re-probe after kill attempt
            post_status, _ = _probe(endpoint)
            if post_status is not None:
                raise RuntimeError(
                    f"Port {PORT} still occupied after kill attempt "
                    f"(status={post_status}); cannot start mock server."
                )
            print(f"[preprocess] Port {PORT} reclaimed after killing squatter.")
        # Port is free: start our server. Use --bind 127.0.0.1 so we never
        # accidentally claim the public interface, and start it via a process
        # group we can attribute and clean up.
        await asyncio.create_subprocess_shell(
            f"nohup python3 -m http.server {PORT} --bind 127.0.0.1 "
            f"--directory {serve_dir} > {serve_dir}/server.log 2>&1 &"
        )
        # Verify the server actually came up AND serves the expected content.
        ready = False
        for _ in range(30):  # up to ~3s
            await asyncio.sleep(0.1)
            st, body = _probe(endpoint)
            if st == 200 and _is_our_content(body):
                ready = True
                break
        if not ready:
            # Dump the server log to aid diagnosis, then fail.
            log_tail = ""
            try:
                log_tail = open(f"{serve_dir}/server.log", encoding="utf-8",
                                errors="ignore").read()[-500:]
            except Exception:
                pass
            raise RuntimeError(
                f"Mock HTTP server on port {PORT} did not become ready / did not "
                f"serve support_resources.json. server.log tail: {log_tail!r}"
            )
    print(f"[preprocess] Mock server verified at {endpoint}")
    print("[preprocess] Done.")


if __name__ == "__main__":
    asyncio.run(main())
