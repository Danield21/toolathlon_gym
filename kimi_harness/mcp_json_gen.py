"""Generate a kimi-code mcp.json from a Toolathlon task config.

Replicates utils/mcp/tool_servers.py resolution semantics exactly:
- ${local_servers_paths} / ${agent_workspace} / ${task_dir} substitution
- env merge: yaml env, then os.environ wins for keys present in both
  (toolathlon: full_env = {**yaml_env, **os.environ}), then the PG_* bridge
  derived from PGHOST/PGPORT/PGDATABASE/PGUSER/PGPASSWORD.
- startupTimeoutMs = max(yaml client_session_timeout_seconds,
  MCP_STDIO_TIMEOUT_MIN) * 1000.

Only the resolved yaml env (+ PG bridge) is written into mcp.json; kimi-code
merges the parent process env underneath, so secrets stay out of the file.

Usage:
  python3 mcp_json_gen.py --task_dir <task> --agent_workspace <abs> \
      --task_src_dir <abs> --out <path> \
      [--local-tools claim_done,python_execute] \
      [--claim-marker <path>] \
      [--project-root /workspace]
"""
import argparse
import json
import os
from pathlib import Path

import yaml


def _resolve(value, local_servers_path: str, agent_workspace: str, task_dir: str):
    if not isinstance(value, str):
        return value
    return (value
            .replace("${local_servers_paths}", local_servers_path)
            .replace("${agent_workspace}", agent_workspace)
            .replace("${task_dir}", task_dir))


def _uv_project_dir(args, cwd):
    """The directory whose `.venv` uv will use.

    `uv --directory <dir> run ...` uses `<dir>/.venv`; otherwise uv falls back
    to the project root discovered from `cwd`. Returns an absolute path.
    """
    for i, a in enumerate(args):
        if a == "--directory" and i + 1 < len(args):
            return os.path.abspath(args[i + 1])
        if isinstance(a, str) and a.startswith("--directory="):
            return os.path.abspath(a.split("=", 1)[1])
    return os.path.abspath(cwd)


def _pin_uv_env(final_env, args, cwd):
    """Force `uv run` to reuse the image-build `.venv` offline.

    The eval container exports `VIRTUAL_ENV=/opt/venv`, which uv sees as
    mismatched with each MCP server's own `.venv` and therefore *ignores* the
    pre-synced env, triggering a full rebuild from source on every cold start.
    Under no-egress eval containers that rebuild hangs until the MCP
    startupTimeoutMs fires and kimi drops the server (root cause of the
    20260807-224427 batch failing). Pin VIRTUAL_ENV at the server's own .venv
    and lock uv to frozen+offline so it never tries to reach the network.
    """
    project_dir = _uv_project_dir(args, cwd)
    venv = os.path.join(project_dir, ".venv")
    final_env["VIRTUAL_ENV"] = venv
    # `--frozen`/`--offline`/`--no-sync` env equivalents (uv 0.5+). Frozen =
    # honor uv.lock as-is (no re-resolve); offline = never hit the network;
    # no-sync = skip the sync step entirely and run against the existing .venv.
    # Together they make `uv run` reuse the image-build .venv deterministically.
    #
    # NO_SYNC is critical: without it, `uv run <package>` still executes a sync
    # pass that *builds* the project from source (resolving build-system.requires
    # like hatchling). On a no-egress compute node hatchling is not in the cache,
    # the build fails, and the MCP server never starts — its tools silently
    # become "not found" (root cause of excel/google_sheet failing in C.3 rerun
    # 2026-08-13). NO_SYNC skips that build entirely.
    final_env["UV_FROZEN"] = "true"
    final_env["UV_OFFLINE"] = "true"
    final_env["UV_NO_SYNC"] = "true"
    # Pin the cache to the image-build location so any residual uv resolution
    # (e.g. `--no-build-isolation` paths) finds the pre-populated wheels. The
    # image bakes the cache at /root/.cache/uv; without this, a custom HOME or
    # XDG_CACHE_HOME inside the container could hide it.
    if "UV_CACHE_DIR" not in final_env and "UV_CACHE_DIR" not in os.environ:
        final_env["UV_CACHE_DIR"] = "/root/.cache/uv"
    return final_env


def build_mcp_servers(
    needed_servers,
    agent_workspace: str,
    task_src_dir: str = "",
    config_dir: str = "configs/mcp_servers",
    local_servers_path: str = None,
) -> dict:
    local_servers_path = local_servers_path or os.environ.get(
        "LOCAL_SERVERS_PATH", os.path.abspath("./local_servers"))
    agent_workspace = os.path.abspath(agent_workspace)
    task_src_dir = os.path.abspath(task_src_dir) if task_src_dir else ""

    timeout_min_s = float(os.environ.get("MCP_STDIO_TIMEOUT_MIN", "0"))
    servers = {}
    config_path = Path(config_dir)
    remaining = set(needed_servers)

    for config_file in sorted(config_path.glob("*.yaml")):
        with open(config_file, encoding="utf-8") as f:
            cfg = yaml.safe_load(f)
        if not cfg:
            continue
        name = cfg.get("name", config_file.stem)
        if name not in remaining:
            continue
        remaining.discard(name)

        params = cfg.get("params", {})

        def res(v):
            return _resolve(v, local_servers_path, agent_workspace, task_src_dir)

        command = res(params.get("command", ""))
        args = [res(a) for a in params.get("args", [])]
        yaml_env = {k: res(v) for k, v in params.get("env", {}).items()}
        cwd = res(params.get("cwd", agent_workspace))

        # filesystem MCP: allow the sessions/ subtree so the agent can read
        # full (untruncated) tool outputs from tool-oversize/. The exact
        # path <task_root>/.kimi_home/sessions/<wd>/<sid>/tool-oversize/
        # requires <wd> and <sid> which are not known at mcp.json generation
        # time, so we allow the sessions/ root as a pragmatic compromise.
        # wire.jsonl (under agents/) is protected by the terminal tool's
        # DENIED_PATHS, not by filesystem MCP boundary. See audit doc §4.2.
        if name == "filesystem":
            task_root = os.path.dirname(agent_workspace)
            sessions_dir = os.path.join(task_root, ".kimi_home", "sessions")
            if sessions_dir not in args:
                args.append(sessions_dir)

        # Boundary hardening: deny the terminal tool access to tool/harness
        # source trees (deny takes priority over ALLOWED_DIR). The agent runs
        # capability-masked, but this adds a second layer so `cat /opt/...`
        # and similar are refused at the tool boundary too.
        #
        # Also deny `.kimi_home` (a sibling of the workspace under task_root):
        # it holds session internals (wire.jsonl LLM transcripts, MCP config,
        # per-agent tool-result offloads) that are outside the declared
        # workspace boundary. See dev_docs/2026-08-13-c2-tz-fix-design.md §1 (P0-1).
        if name == "terminal" and os.environ.get("KIMI_DISABLE_BOUNDARY") != "1":
            denied = ["/opt/local_servers", "/opt/kimi-code"]
            kimi_home = os.environ.get("KIMI_CODE_HOME", "")
            if not kimi_home:
                kimi_home = os.path.join(os.path.dirname(os.path.abspath(agent_workspace)),
                                         ".kimi_home")
            if kimi_home:
                denied.append(os.path.realpath(os.path.normpath(kimi_home)))
            yaml_env["DENIED_PATHS"] = ",".join(denied)

        # toolathlon: full_env = {**yaml_env, **os.environ} — os.environ wins.
        # kimi merges process env underneath config env, so to reproduce the
        # same winner we inline only the yaml keys that os.environ overrides.
        final_env = dict(yaml_env)
        for k in yaml_env:
            if k in os.environ:
                final_env[k] = os.environ[k]

        # PG bridge: libpq-style PGHOST/... -> PG_HOST/... used by DB MCP servers
        pg_bridge = {
            "PG_HOST": os.environ.get("PGHOST"),
            "PG_PORT": os.environ.get("PGPORT"),
            "PG_DATABASE": os.environ.get("PGDATABASE"),
            "PG_USER": os.environ.get("PGUSER"),
            "PG_PASSWORD": os.environ.get("PGPASSWORD"),
        }
        for k, v in pg_bridge.items():
            if v is not None:
                final_env[k] = v

        # uv-launched MCP servers: pin to the pre-built `.venv` and forbid any
        # network re-sync, so `uv run` reuses the image-build env deterministically.
        # See _pin_uv_env for the VIRTUAL_ENV-mismatch root cause.
        if command == "uv":
            _pin_uv_env(final_env, args, cwd)

        timeout_s = max(float(cfg.get("client_session_timeout_seconds", 60)),
                        timeout_min_s)

        servers[name] = {
            "command": command,
            "args": args,
            "env": {k: str(v) for k, v in final_env.items()},
            "cwd": cwd,
            "startupTimeoutMs": int(timeout_s * 1000),
        }

    missing = sorted(remaining)
    if missing:
        print(f"[mcp_json_gen] Warning: no yaml config found for: {missing}")

    return servers


def local_tools_entry(workspace: str, marker: str, tools, python_bin: str,
                      harness_dir: str, timeout_s: float = 60.0) -> dict:
    return {
        "command": python_bin,
        "args": [
            os.path.join(harness_dir, "local_tools_server.py"),
            "--workspace", workspace,
            "--marker", marker,
            "--tools", ",".join(tools),
        ],
        "env": {},
        "cwd": workspace,
        "startupTimeoutMs": int(timeout_s * 1000),
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--task_dir", required=True)
    ap.add_argument("--agent_workspace", required=True)
    ap.add_argument("--task_src_dir", default="")
    ap.add_argument("--out", required=True)
    ap.add_argument("--project-root", default=".")
    ap.add_argument("--config-dir", default=None)
    ap.add_argument("--local-servers-path", default=None)
    ap.add_argument("--local-tools", default="")
    ap.add_argument("--claim-marker", default="")
    ap.add_argument("--python-bin", default=os.environ.get("PYTHON_BIN", "python3"))
    ap.add_argument("--harness-dir", default="/workspace/kimi_harness")
    args = ap.parse_args()

    project_root = os.path.abspath(args.project_root)
    config_dir = args.config_dir or os.path.join(project_root, "configs/mcp_servers")

    task_cfg_path = Path(project_root) / "tasks/finalpool" / args.task_dir / "task_config.json"
    with open(task_cfg_path, encoding="utf-8") as f:
        task_cfg = json.load(f)

    servers = build_mcp_servers(
        task_cfg.get("needed_mcp_servers", []),
        agent_workspace=args.agent_workspace,
        task_src_dir=args.task_src_dir or str(Path(project_root) / "tasks/finalpool" / args.task_dir),
        config_dir=config_dir,
        local_servers_path=args.local_servers_path,
    )

    local_tools = [t for t in args.local_tools.split(",") if t]
    if not local_tools:
        local_tools = task_cfg.get("needed_local_tools", []) or []
    if local_tools:
        servers["local"] = local_tools_entry(
            workspace=os.path.abspath(args.agent_workspace),
            marker=args.claim_marker,
            tools=local_tools,
            python_bin=args.python_bin,
            harness_dir=args.harness_dir,
        )

    out = {"mcpServers": servers}
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(out, f, indent=2, ensure_ascii=False)
    print(f"[mcp_json_gen] wrote {args.out} with servers: {sorted(servers)}")


if __name__ == "__main__":
    main()
