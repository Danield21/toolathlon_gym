#!/usr/bin/env python3
"""Launch the fixed 166-case MS-Eval suite; no API calls during preflight."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import secrets
import shutil
import signal
import subprocess
import sys
import tempfile
import time
from urllib.parse import urlsplit

HERE = Path(__file__).resolve().parent
# Fixed MS-Eval membership is an internal invariant, not a model setting.
_TASKS_SHA256 = "b6c7e6be77e63d10a5f26ef3ef200f8de87767f09b883d0e90c3ec30ab0e5068"


def require(condition, message):
    if not condition:
        raise ValueError(message)


def integer(name, default, low=1, high=1_000_000):
    value = int(os.environ.get(name, str(default)))
    require(low <= value <= high, f"Invalid {name}; expected {low}..{high}")
    return value


def prepare():
    preset = json.loads((HERE / "model.json").read_text())
    project = Path(os.environ.get("PROJECT_ROOT", HERE.parents[2])).resolve()
    runtime = Path(os.environ.get("RUNTIME_ROOT", project.parent / "toolathlon_gym_eval_dockers")).resolve()
    manifest = (HERE / "tasks.txt").read_bytes()
    require(hashlib.sha256(manifest).hexdigest() == _TASKS_SHA256, "MS-Eval tasks.txt was changed")
    tasks = manifest.decode().splitlines()
    require(len(tasks) == len(set(tasks)) == 166, "MS-Eval requires exactly 166 unique cases")
    require(all(re.fullmatch(r"[A-Za-z0-9_-]+", t) for t in tasks), "Invalid task ID")
    oracle = Path(os.environ.get("ORACLE_TASK_DIR", project.parent / "metrics_eval_bw/oracle166")).resolve()
    for task in tasks:
        case = project / "tasks/finalpool" / task
        require((case / "task_config.json").is_file(), f"Missing task_config.json: {task}")
        brief = oracle / f"{task}.task.md" if preset["variant"] == "oracle" else case / "docs/task.md"
        require(brief.is_file(), f"Missing task brief: {task}")
    if preset["variant"] == "oracle":
        require({p.name.removesuffix('.task.md') for p in oracle.glob('*.task.md')} == set(tasks),
                "Oracle briefs must match the fixed 166-case manifest exactly")
    env = dict(os.environ)
    installation = os.environ.get("ENROOT_INSTALL_ROOT", "")
    executable = shutil.which("enroot")
    if not installation and executable:
        resolved = Path(executable).resolve()
        if resolved.name == "enroot-wrapper":
            installation = str(resolved.parent.parent)
    if installation:
        installation = Path(installation).resolve()
        env["PATH"] = str(installation / "bin") + os.pathsep + env.get("PATH", "")
        env["ENROOT_LIBRARY_PATH"] = str(installation / "lib")
        env["ENROOT_SYSCONF_PATH"] = str(installation / "etc")
    # A previous experiment's roster, prompts, or sampling must not leak in.
    for key in list(env):
        if key.startswith("KIMI_") or key.startswith("RELAY_"):
            del env[key]
    env.update({
        "PROJECT_ROOT": str(project), "RUNTIME_ROOT": str(runtime),
        "TOOLATHLON_EVAL_DOCKER_ROOT": str(runtime),
        "KIMI_DIST": os.environ.get("KIMI_DIST", str(runtime / "kimi_dist/kimi-code")),
        "AGENT_SQSH": os.environ.get("AGENT_SQSH", str(runtime / "images/toolathlon-pack.sqsh")),
        "AGENT_TEMPLATE": os.environ.get("AGENT_TEMPLATE", ""),
        "KIMI_PROVIDER_TYPE": os.environ.get("API_FORMAT", preset.get("api_format", "openai")),
        "KIMI_PROVIDER_NAME": "remote-eval", "KIMI_MODEL_ALIAS": "eval-model",
        "MODEL_NAME": os.environ.get("MODEL_NAME", preset["model"]),
        "KIMI_MAX_CONTEXT": str(integer("MAX_CONTEXT", 262144)),
        "KIMI_TASK_TIMEOUT_S": str(integer("TASK_TIMEOUT_S", 7200)),
        "MAX_STEPS": str(integer("MAX_STEPS", 100)),
        "MAX_CONCURRENT": str(integer("MAX_CONCURRENT", 1, high=166)),
        "PG_PORT_BASE": str(integer("PG_PORT_BASE", 25432, high=55000)),
        "PGUSER": "eigent", "PGDATABASE": "toolathlon_gym",
        # Public benchmark fixture: legacy preprocess/evaluation code uses this pair.
        "PGPASSWORD": "camel",
        "PG_TEST_ONLY": "0", "ENROOT_ROOTFS_WRITABLE": "yes",
        "MOCK_PORT_WAIT_LOOPS": "720",
    })
    require(env["KIMI_PROVIDER_TYPE"] in ("openai", "anthropic"), "API_FORMAT must be openai or anthropic")
    require(bool(re.fullmatch(r"[A-Za-z0-9_.:/()+-]+", env["MODEL_NAME"])), "Invalid MODEL_NAME")
    for source, target, default in [
        ("TEMPERATURE", "KIMI_MODEL_TEMPERATURE", preset["temperature"]),
        ("TOP_P", "KIMI_MODEL_TOP_P", preset["top_p"]),
        ("REASONING_EFFORT", "KIMI_MODEL_THINKING_EFFORT", preset["reasoning_effort"]),
    ]:
        env[target] = os.environ.get(source, default)
    for name in ("KIMI_MODEL_TEMPERATURE", "KIMI_MODEL_TOP_P"):
        if env[name]:
            value = float(env[name])
            require(0 <= value <= (1 if name.endswith("TOP_P") else 2), f"Invalid {name}")
    require(env["KIMI_MODEL_THINKING_EFFORT"] in ("", "low", "medium", "high", "xhigh", "max", "adaptive"),
            "Invalid REASONING_EFFORT")
    if preset["variant"] == "plan1st":
        env["KIMI_PLAN_FIRST"] = "1"
    else:
        env["KIMI_CODE_SUBAGENT"] = "7" if preset["variant"] == "spec" else "3"
    if preset["variant"] == "oracle":
        env["KIMI_TASK_MD_DIR"] = str(oracle)
    env["GATEWAY_MAX_OUTPUT_TOKENS"] = str(integer("MAX_OUTPUT_TOKENS", preset["max_output_tokens"], low=0))
    env.pop("MODEL_N", None)
    for option in ("ENABLE_THINKING", "BREAK_SCHEMA_CYCLES"):
        env["GATEWAY_" + option] = os.environ.get(option, str(preset[option.lower()]))
        require(env["GATEWAY_" + option] in ("0", "1"), f"{option} must be 0 or 1")
    plan = {
        "suite": "MS-Eval", "cases": len(tasks),
        "model": env["MODEL_NAME"], "variant": preset["variant"],
        "api_format": env["KIMI_PROVIDER_TYPE"],
        "subagent_roster": env.get("KIMI_CODE_SUBAGENT", "3 (default)"),
        "plan_first": env.get("KIMI_PLAN_FIRST") == "1",
        "briefs": str(oracle if preset["variant"] == "oracle" else project / "tasks/finalpool"),
        "max_steps": int(env["MAX_STEPS"]), "timeout_s": int(env["KIMI_TASK_TIMEOUT_S"]),
        "max_context": int(env["KIMI_MAX_CONTEXT"]), "concurrency": int(env["MAX_CONCURRENT"]),
        "temperature": env["KIMI_MODEL_TEMPERATURE"] or "model default",
        "top_p": env["KIMI_MODEL_TOP_P"] or "model default",
        "reasoning_effort": env["KIMI_MODEL_THINKING_EFFORT"] or "model default",
        "max_output_tokens": int(env["GATEWAY_MAX_OUTPUT_TOKENS"]),
        "enable_thinking": env["GATEWAY_ENABLE_THINKING"] == "1",
        "break_schema_cycles": env["GATEWAY_BREAK_SCHEMA_CYCLES"] == "1",
    }
    return env, tasks, plan


def preflight(env):
    require(os.geteuid() != 0, "Run as a regular user: PostgreSQL initdb cannot run as root")
    for name in ("bash", "enroot", "initdb", "pg_ctl", "pg_isready", "psql", "createdb",
                 "rsync", "setsid", "ss", "unsquashfs", "curl", "gzip", "pgrep", "md5sum"):
        require(shutil.which(name, path=env.get("PATH")) is not None, f"Missing executable: {name}; add the runtime tools to PATH")
    try:
        probe = subprocess.run(["enroot", "version"], env=env, capture_output=True, timeout=10)
        require(probe.returncode == 0, "Enroot cannot run; check ENROOT_INSTALL_ROOT and its libraries")
    except subprocess.TimeoutExpired:
        raise ValueError("Enroot version check timed out; check ENROOT_INSTALL_ROOT and avoid recursive wrapper symlinks") from None
    for path in [Path(env["KIMI_DIST"]) / "dist/main.mjs", HERE / "runner.sh", HERE / "api_gateway.py",
                 HERE / "schema.py", HERE / "eval_config.json"]:
        require(path.is_file(), f"Missing runtime file: {path}")
    require(Path(env["AGENT_SQSH"]).is_file() or bool(env["AGENT_TEMPLATE"]) and Path(env["AGENT_TEMPLATE"]).is_dir(),
            "Set AGENT_SQSH to a prepared Toolathlon image or AGENT_TEMPLATE to its unpacked rootfs")
    require((Path(env["PROJECT_ROOT"]) / "db/init.sql.gz").is_file(), "Missing db/init.sql.gz")
    require((Path(env["PROJECT_ROOT"]) / "kimi_harness/kimi_main.py").is_file(), "Missing benchmark harness")
    for line in (HERE / "runner.sh").read_text().split("<<'OVERLAYS'\n", 1)[1].split("\nOVERLAYS", 1)[0].splitlines():
        source = line.split()[0]
        require((Path(env["PROJECT_ROOT"]) / source).is_file(), f"Missing benchmark runtime overlay: {source}")
    for name, marker in [("index.js", "Results are paginated and include summary totals"),
                         ("client.js", "Using PgCanvasRouter")]:
        content = (Path(env["PROJECT_ROOT"]) / "local_servers/mcp-canvas-lms/build" / name).read_text()
        require(marker in content, "Rebuild the current Canvas MCP server before evaluation")
    require(bool(os.environ.get("API_KEY")), "Set API_KEY in config.env or the environment")
    url = urlsplit(os.environ.get("API_BASE_URL", ""))
    require(url.scheme in ("http", "https") and bool(url.hostname) and not url.username
            and not url.password and not url.query and not url.fragment,
            "Set API_BASE_URL to an HTTP(S) API base URL, without credentials, query, or fragment")


def launch(env, tasks, plan):
    run_id = time.strftime("%Y%m%d-%H%M%S") + "-" + secrets.token_hex(3)
    output = Path(os.environ.get("OUTPUT_ROOT", HERE / "results")).resolve() / run_id
    output.mkdir(parents=True, exist_ok=False)
    (output / "run.json").write_text(json.dumps(plan, indent=2) + "\n")
    (output / "tasks.txt").write_text("\n".join(tasks) + "\n")
    token = secrets.token_hex(24)
    gateway = runner = None
    def stop(signum, frame):
        if runner is not None and runner.poll() is None:
            runner.send_signal(signal.SIGTERM)
        raise KeyboardInterrupt
    previous = {s: signal.signal(s, stop) for s in (signal.SIGINT, signal.SIGTERM, signal.SIGHUP)}
    try:
        with tempfile.TemporaryDirectory(prefix=f"mse{os.getuid()}_", dir="/dev/shm") as temp:
            root = Path(temp)
            gateway_env = dict(env, GATEWAY_TOKEN=token, API_KEY=os.environ["API_KEY"],
                               API_BASE_URL=os.environ["API_BASE_URL"])
            gateway = subprocess.Popen([sys.executable, str(HERE / "api_gateway.py"), str(root / "ready")],
                                       env=gateway_env, start_new_session=True)
            for _ in range(100):
                if (root / "ready").is_file():
                    break
                require(gateway.poll() is None, "API gateway failed to start")
                time.sleep(0.05)
            require((root / "ready").is_file(), "API gateway startup timed out")
            # The real endpoint/key are held by the local gateway, never passed to the agent.
            env.update(MODEL_API_URL=(root / "ready").read_text().strip(), MODEL_API_KEY=token,
                       DUMP_ROOT=str(output), RUN_ID=run_id, PG_RUNTIME_ROOT=str(root / "pg"),
                       PG_PORT_LEASE_ROOT=f"/dev/shm/toolathlon_pg_port_leases_{os.getuid()}",
                       no_proxy="127.0.0.1,localhost,::1", NO_PROXY="127.0.0.1,localhost,::1")
            for key in ("API_KEY", "API_BASE_URL", "http_proxy", "https_proxy", "HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "all_proxy"):
                env.pop(key, None)
            for kind in ("DATA", "TEMP", "RUNTIME", "CACHE"):
                directory = root / kind.lower()
                directory.mkdir()
                env[f"ENROOT_{kind}_PATH"] = str(directory)
            print(f"MS-Eval: 166 cases; results: {output}", flush=True)
            runner = subprocess.Popen(["bash", str(HERE / "runner.sh"), *tasks], env=env, start_new_session=True)
            try:
                return runner.wait()
            finally:
                if runner.poll() is None:
                    runner.send_signal(signal.SIGTERM)
                    runner.wait()
    finally:
        if gateway is not None:
            gateway.terminate()
            try:
                gateway.wait(timeout=5)
            except subprocess.TimeoutExpired:
                gateway.kill()
                gateway.wait()
        for sig, handler in previous.items():
            signal.signal(sig, handler)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--dry-run", action="store_true", help="validate 166 cases and show settings; no credentials/runtime needed")
    mode.add_argument("--check", action="store_true", help="also check local runtime and credentials; no API calls")
    args = parser.parse_args()
    env, tasks, plan = prepare()
    print(json.dumps(plan, indent=2), flush=True)
    if args.dry_run:
        return 0
    preflight(env)
    if args.check:
        print("Preflight passed; no evaluation or API request was started.")
        return 0
    return launch(env, tasks, plan)


if __name__ == "__main__":
    os.umask(0o077)
    try:
        sys.exit(main())
    except (ValueError, FileNotFoundError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(2)
    except KeyboardInterrupt:
        sys.exit(130)
