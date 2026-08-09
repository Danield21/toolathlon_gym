#!/usr/bin/env python3
"""Monitor running eval dumps for infra problems and kill bad cases early.

Scans the kimi-code.log of each running/incomplete case for signatures that
indicate the agent is fighting infrastructure rather than solving the task:

  - MCP server unavailable / timed out
  - FastMCP / ImportError
  - connection refused / ECONNREFUSED to expected ports
  - repeated "server unavailable" (the agent retrying dead servers)
  - agent reading /opt/local_servers or /opt/kimi-code (boundary violation)

When a problem is detected, it prints a clear alert and (if --kill is given)
terminates the enroot process group for that case, then cleans up its
/dev/shm artifacts.

Usage:
    python3 monitor_traj.py [--kill] [--dump-root DIR] [--interval 30]
"""

import argparse
import glob
import os
import re
import signal
import subprocess
import sys
import time

# Patterns that signal infra trouble (case-insensitive).
BAD_PATTERNS = [
    re.compile(r"mcp server unavailable", re.I),
    re.compile(r"transport=stdio\s+status=failed", re.I),
    re.compile(r"Timed out after \d+ms", re.I),
    re.compile(r"ImportError.*FastMCP", re.I),
    re.compile(r"cannot import name 'FastMCP'", re.I),
    re.compile(r"ECONNREFUSED.*127\.0\.0\.1:(8081|8317|19317)", re.I),
    re.compile(r"connection refused.*8081", re.I),
    # Agent probing infra it shouldn't (boundary violation during solving).
    re.compile(r"PermissionError.*(/opt/local_servers|/opt/kimi-code)", re.I),
    # Agent stuck retrying a dead server repeatedly.
    re.compile(r"woocommerce.*503|woocommerce.*502", re.I),
]

# These are INFO lines that contain "server" but are benign — don't false-alarm.
BENIGN = [
    re.compile(r"toolCount=\d+", re.I),  # tool registration count
    re.compile(r"llm config", re.I),
    re.compile(r"llm request", re.I),
]


def find_kimi_logs(dump_root: str, since_ts: float):
    """Yield (log_path, task_name) for every kimi-code.log under dump_root."""
    for log in glob.glob(os.path.join(dump_root, "*/*/kimi-code*/**/kimi-code.log"), recursive=True):
        try:
            mtime = os.path.getmtime(log)
        except OSError:
            continue
        if mtime < since_ts:
            continue
        # Extract task name from path: dump_root/<task>/<slot>/...
        parts = log.replace(dump_root + "/", "").split("/")
        task = parts[0] if parts else "?"
        yield log, task


def scan_log(log_path: str):
    """Return list of (line_no, line, pattern_name) for bad lines."""
    try:
        with open(log_path, "r", errors="replace") as f:
            lines = f.readlines()
    except OSError:
        return []
    hits = []
    for i, line in enumerate(lines, 1):
        if any(b.search(line) for b in BENIGN):
            continue
        for pat in BAD_PATTERNS:
            if pat.search(line):
                hits.append((i, line.rstrip(), pat.pattern[:50]))
                break
    return hits


def find_enroot_pgid_for_task(task: str):
    """Try to find the enroot process group for a task, via run.log."""
    # Look for "[runner] Enroot launch pid=NNNN process_group=NNNN"
    dump_root = os.environ.get("DUMP_ROOT", "")
    for runlog in glob.glob(os.path.join(dump_root, task, "*/run.log")):
        try:
            with open(runlog) as f:
                content = f.read()
        except OSError:
            continue
        m = re.search(r"Enroot launch pid=(\d+)\s+process_group=(\d+)", content)
        if m:
            return m.group(2)
    return None


def kill_task(task: str, pgid: str):
    """Kill the enroot process group and clean up /dev/shm artifacts."""
    killed = []
    if pgid:
        try:
            os.killpg(int(pgid), signal.SIGTERM)
            killed.append(f"SIGTERM pgid={pgid}")
        except (OSError, ProcessLookupError):
            pass
    # Also kill any process whose cmdline mentions the task name + enroot
    try:
        out = subprocess.check_output(["pgrep", "-af", f"enroot.*{task}"], text=True)
        for line in out.strip().split("\n"):
            pid = line.split()[0]
            try:
                os.kill(int(pid), signal.SIGTERM)
                killed.append(f"SIGTERM pid={pid}")
            except (OSError, ProcessLookupError):
                pass
    except (subprocess.CalledProcessError, FileNotFoundError):
        pass
    # Clean /dev/shm artifacts for this task
    for d in glob.glob(f"/dev/shm/enroot_data/agent-{task}*"):
        try:
            subprocess.run(["rm", "-rf", d], timeout=30, check=False)
            killed.append(f"rm {d}")
        except Exception:
            pass
    return killed


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dump-root", default=os.environ.get(
        "DUMP_ROOT",
        "/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/dumps/kimi-code_deepseek-v4-flash"))
    ap.add_argument("--kill", action="store_true",
                    help="kill cases that show infra problems")
    ap.add_argument("--interval", type=int, default=30,
                    help="scan interval in seconds (default 30)")
    ap.add_argument("--since-minutes", type=int, default=360,
                    help="only look at logs modified in last N minutes (default 360)")
    args = ap.parse_args()

    since_ts = time.time() - args.since_minutes * 60
    seen_problems = set()  # (task, log) already reported

    print(f"[monitor] scanning {args.dump_root} every {args.interval}s, kill={args.kill}")
    print(f"[monitor] watching for: {[p.pattern[:40] for p in BAD_PATTERNS]}")
    sys.stdout.flush()

    while True:
        for log_path, task in find_kimi_logs(args.dump_root, since_ts):
            key = (task, log_path)
            hits = scan_log(log_path)
            if not hits:
                continue
            new_hits = [h for h in hits if (task, h[0], h[2]) not in seen_problems]
            if not new_hits:
                continue
            for ln, line, pat in new_hits:
                seen_problems.add((task, ln, pat))
                tag = "*** INFRA ALERT ***" if args.kill else "[warn]"
                print(f"\n{tag} task={task} line={ln} pattern={pat}")
                print(f"  {line[:200]}")
                sys.stdout.flush()
            if args.kill and key not in [k for k in seen_problems if isinstance(k, tuple) and len(k) == 2]:
                pgid = find_enroot_pgid_for_task(task)
                killed = kill_task(task, pgid or "")
                if killed:
                    print(f"[monitor] KILLED task={task}: {killed}")
                else:
                    print(f"[monitor] wanted to kill task={task} but found no process")
                sys.stdout.flush()
        time.sleep(args.interval)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[monitor] stopped")
