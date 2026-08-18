#!/usr/bin/env python3
"""Run task evaluators on no-claim runs without changing official eval_res.json.

Official Toolathlon scoring requires claim_done. For auditing, it is still useful
to know whether the artifacts already present on disk would satisfy the task
evaluator. This script bypasses only the task-status gate and writes the result
to artifact_eval_res.json next to eval_res.json.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path


RUN_DIR_RE = re.compile(r"^\d{8}-\d{6}(?:[-_][A-Za-z0-9_.-]+)?_slot\d+$")


def load_json(path: Path) -> dict:
    try:
        return json.loads(path.read_text(errors="replace"))
    except Exception:
        return {}


def find_run_dirs(root: Path) -> list[Path]:
    if not root.is_dir():
        return []
    if RUN_DIR_RE.match(root.name):
        return [root]
    out: list[Path] = []
    for p in root.rglob("*"):
        if p.is_dir() and RUN_DIR_RE.match(p.name):
            out.append(p)
    return sorted(out)


def find_inner_dir(run_dir: Path) -> Path | None:
    for model_dir in run_dir.iterdir():
        if not model_dir.is_dir() or not model_dir.name.startswith("kimi-code"):
            continue
        for inner in model_dir.iterdir():
            if inner.is_dir() and inner.name.startswith("SingleUserTurn"):
                return inner
    return None


def has_user_artifacts(workspace: Path) -> bool:
    if not workspace.is_dir():
        return False
    for p in workspace.rglob("*"):
        if not p.is_file():
            continue
        rel_parts = p.relative_to(workspace).parts
        if any(part.startswith(".") or part == "__pycache__" for part in rel_parts):
            continue
        return True
    return False


def parse_claim_done(run_dir: Path) -> bool | None:
    text = (run_dir / "run.log").read_text(errors="replace") if (run_dir / "run.log").exists() else ""
    m = re.search(r"exited rc=\d+ claim_done=(True|False)", text)
    if not m:
        return None
    return m.group(1) == "True"


def host_path(path_value: str, inner: Path) -> Path:
    if path_value.startswith("/workspace/dumps/"):
        rel = path_value.removeprefix("/workspace/dumps/")
        parts = Path(rel).parts
        if len(parts) >= 2:
            return inner.parent / parts[1] / Path(*parts[2:])
    if path_value.startswith("/workspace/"):
        return Path.cwd() / path_value.removeprefix("/workspace/")
    return Path(path_value)


def host_eval_command(command: str) -> str:
    command = command.replace("/opt/venv/bin/python3", sys.executable)
    command = command.replace("python3 ", f"{sys.executable} ", 1)
    return command


def run_artifact_eval(run_dir: Path, *, force: bool, timeout_s: int) -> dict | None:
    inner = find_inner_dir(run_dir)
    if inner is None:
        return None
    if parse_claim_done(run_dir) is True:
        return None

    out_path = inner / "artifact_eval_res.json"
    if out_path.exists() and not force:
        return load_json(out_path)

    traj = load_json(inner / "traj_log.json")
    cfg = traj.get("config") or {}
    eval_cfg = cfg.get("evaluation") or {}
    eval_command = eval_cfg.get("evaluation_command")
    if not eval_command:
        return None

    workspace = host_path(cfg.get("agent_workspace") or "", inner)
    if not has_user_artifacts(workspace):
        result = {
            "pass": None,
            "details": "No non-hidden workspace artifacts found; artifact-only evaluation skipped.",
        }
        out_path.write_text(json.dumps(result, indent=2))
        return result

    # Many task evaluators write their own report to --res_log_file. Never pass
    # the official traj_log.json here, or artifact-only auditing can destroy the
    # run metadata that explains why claim_done was missing.
    res_log_file = inner / "artifact_eval_log.json"
    groundtruth = Path.cwd() / (eval_cfg.get("groundtruth_workspace") or "")
    launch_time = cfg.get("launch_time") or ""
    command = (
        f"{host_eval_command(eval_command)} "
        f"--res_log_file {res_log_file} "
        f"--agent_workspace {workspace} "
        f"--groundtruth_workspace {groundtruth} "
        f"--launch_time \"{launch_time}\""
    )
    env = os.environ.copy()
    env["PYTHONPATH"] = str(Path.cwd()) + os.pathsep + env.get("PYTHONPATH", "")
    try:
        proc = subprocess.run(
            command,
            shell=True,
            cwd=Path.cwd(),
            env=env,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout_s,
        )
        result = {
            "pass": proc.returncode == 0,
            "returncode": proc.returncode,
            "command": command,
            "stdout": proc.stdout,
            "stderr": proc.stderr,
            "details": "Artifact-only evaluator run; official claim_done status is unchanged.",
        }
        if proc.returncode != 0:
            result["failure"] = proc.stdout
    except subprocess.TimeoutExpired as exc:
        result = {
            "pass": None,
            "returncode": None,
            "command": command,
            "stdout": exc.stdout or "",
            "stderr": exc.stderr or "",
            "details": f"Artifact-only evaluator timed out after {timeout_s}s.",
        }

    out_path.write_text(json.dumps(result, indent=2))
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("roots", nargs="+", type=Path)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--timeout-s", type=int, default=300)
    args = parser.parse_args()

    total = 0
    ran = 0
    for root in args.roots:
        for run_dir in find_run_dirs(root):
            total += 1
            result = run_artifact_eval(run_dir, force=args.force, timeout_s=args.timeout_s)
            if result is None:
                continue
            ran += 1
            task = run_dir.parent.name
            print(f"[artifact-eval] {task}/{run_dir.name} pass={result.get('pass')}")
    print(f"[artifact-eval] considered={total} wrote_or_loaded={ran}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
