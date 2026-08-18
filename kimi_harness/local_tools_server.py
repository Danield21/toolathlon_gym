#!/usr/bin/env python3
"""Pure-stdlib MCP stdio server exposing Toolathlon local tools to kimi-code.

Implements just enough of the MCP stdio transport (newline-delimited
JSON-RPC 2.0) for initialize / ping / tools/list / tools/call.

Tools mirror utils/roles/task_agent.py::_build_local_tools semantics:
  claim_done              -> writes the completion marker file
  python_execute          -> run code under the agent workspace
  save_overlong_output    -> persist large text, return an id
  view_overlong_output    -> paginated read of a saved output
  sleep                   -> asyncio-less time.sleep
  manage_context/history  -> "OK" stubs (memory handled by the harness)

Nothing here prints to stdout except protocol frames.
"""
import argparse
import json
import os
import subprocess
import sys
import time
import uuid

OVERLONG_DIR = ".overlong_tool_outputs"
PAGE_SIZE = 10000

PROTOCOL_VERSION = "2024-11-05"

WORKSPACE = ""
MARKER = ""
ENABLED = set()

# Source trees that are runtime dependencies (the MCP servers execute from
# here) but whose *contents* the evaluated agent must not read. Reading them
# to debug infra or reverse-engineer the harness is out of the task boundary.
# This is a read-guard only: it never blocks execution, so MCP tools and the
# interpreter keep working — only agent-driven reads of these paths error out.
#
# `.kimi_home` (a sibling of the workspace under task_root) is also blocked:
# it holds session internals (wire.jsonl LLM transcripts, MCP config, other
# agents' tool-result offloads) that are outside the declared workspace
# boundary. See dev_docs/2026-08-13-c2-tz-fix-design.md §1 (P0-1).
_BLOCKED_READ_PREFIXES = [
    "/opt/local_servers",
    "/opt/kimi-code",
]
_BLOCKED_READ_MSG = (
    "Reading this path is outside your task boundary. The directory holds "
    "tool/harness implementation source, which you may use via the granted "
    "tools but may not inspect. Solve the task using the granted tools and "
    "your workspace only."
)

# Runtime guard prepended to user code: hooks open() so reads of the blocked
# source trees raise, while everything else (including normal execution and
# MCP subprocess spawns) is untouched. The guard normalizes both str and bytes
# paths and resolves symlinks/`..` so an agent cannot escape the block by
# writing `../.kimi_home/...` or `/proc/self/cwd/../.kimi_home`.
_READ_GUARD = (
    "import builtins as _bi, os as _os\n"
    "_BLK = {blocked!r}\n"
    "_MSG = {msg!r}\n"
    "_orig_open = _bi.open\n"
    "def _norm(file):\n"
    "    try:\n"
    "        if isinstance(file, bytes):\n"
    "            s = _os.fsdecode(file)\n"
    "        else:\n"
    "            s = _os.fspath(file)\n"
    "        return _os.path.realpath(_os.path.normpath(s))\n"
    "    except Exception:\n"
    "        return ''\n"
    "def _g(file, mode='r', *a, **k):\n"
    "    p = _norm(file)\n"
    "    if p and any(p == b or p.startswith(b + _os.sep) for b in _BLK) "
    "and any(m in mode for m in ('r', '+')):\n"
    "        raise PermissionError(_MSG + ' [' + p + ']')\n"
    "    return _orig_open(file, mode, *a, **k)\n"
    "_bi.open = _g\n"
)


def _claim_done() -> str:
    if MARKER:
        with open(MARKER, "w", encoding="utf-8") as f:
            f.write(str(time.time()))
    return "Task marked as done. The harness has been notified."


def _python_execute(code: str, filename: str = "", timeout: int = 30) -> str:
    timeout = min(int(timeout), 120)
    if not filename:
        filename = f"{uuid.uuid4()}.py"
    if not filename.endswith(".py"):
        filename += ".py"

    workspace = os.path.abspath(WORKSPACE)
    tmp_dir = os.path.join(workspace, ".python_tmp")
    os.makedirs(tmp_dir, exist_ok=True)

    file_path = os.path.join(tmp_dir, filename)
    guard = "" if os.environ.get("KIMI_DISABLE_BOUNDARY") == "1" else _READ_GUARD.format(
        blocked=list(_BLOCKED_READ_PREFIXES), msg=_BLOCKED_READ_MSG)
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(guard + code)

    python_bin = os.environ.get("PYTHON_EXECUTE_BIN") or sys.executable
    cmd = [python_bin, file_path]
    # Credential isolation (audit §A.3 / security-boundary): agent-authored
    # Python runs here as a child of the kimi-code CLI and would otherwise
    # inherit MODEL_API_KEY/MODEL_API_URL (the CLI needs them to call the
    # model, but the evaluated code must not). PG* are also stripped again
    # as defense-in-depth — launch_kimi already removes them from the agent
    # env, but if a future launcher re-adds them this tool stays safe.
    child_env = dict(os.environ)
    if os.environ.get("KIMI_DISABLE_BOUNDARY") != "1":
        for _secret in ("MODEL_API_KEY", "MODEL_API_URL",
                        "OPENAI_API_KEY", "ANTHROPIC_API_KEY"):
            child_env.pop(_secret, None)
        for _k in list(child_env.keys()):
            if _k.upper().startswith("PG") and _k.upper() in (
                "PGHOST", "PG_HOST", "PGPORT", "PG_PORT",
                "PGUSER", "PG_USER", "PGPASSWORD", "PG_PASSWORD",
                "PGDATABASE", "PG_DATABASE",
            ):
                child_env.pop(_k, None)
    start = time.time()
    try:
        result = subprocess.run(
            cmd, cwd=workspace, capture_output=True, text=True,
            encoding="utf-8", timeout=timeout, env=child_env,
        )
    except subprocess.TimeoutExpired:
        return f"=== TIMEOUT ===\nExceeded {timeout}s limit."

    elapsed = time.time() - start
    parts = []
    if result.stdout:
        parts += ["=== STDOUT ===", result.stdout.rstrip()]
    if result.stderr:
        parts += ["=== STDERR ===", result.stderr.rstrip()]
    parts += [
        "=== INFO ===",
        f"Return code: {result.returncode}",
        f"Time: {elapsed:.2f}s / {timeout}s limit",
    ]
    return "\n".join(parts) if parts else "No output."


def _overlong_dir() -> str:
    d = os.path.join(os.path.abspath(WORKSPACE), OVERLONG_DIR)
    os.makedirs(d, exist_ok=True)
    return d


def _save_overlong_output(content: str, label: str = "") -> str:
    fid = str(uuid.uuid4())[:8]
    path = os.path.join(_overlong_dir(), f"{fid}.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"id": fid, "label": label, "content": content,
                   "saved_at": time.time()}, f)
    preview = content[:200] + ("..." if len(content) > 200 else "")
    return (f"Saved {len(content)} chars as [{fid}] label='{label}'.\n"
            f"Preview: {preview}\n"
            f"Use view_overlong_output(id='{fid}') to read.")


def _view_overlong_output(id: str, page: int = 0) -> str:
    path = os.path.join(_overlong_dir(), f"{id}.json")
    if not os.path.exists(path):
        return f"No saved output with id '{id}'."
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    content = data["content"]
    total_pages = max(1, (len(content) + PAGE_SIZE - 1) // PAGE_SIZE)
    start = page * PAGE_SIZE
    chunk = content[start:start + PAGE_SIZE]
    return (f"[{id}] label='{data['label']}' | "
            f"page {page + 1}/{total_pages} | "
            f"chars {start}-{start + len(chunk)} of {len(content)}\n\n"
            f"{chunk}")


def _sleep(seconds: float = 1) -> str:
    time.sleep(float(seconds))
    return f"Slept {seconds} seconds."


def _stub(action: str = "") -> str:
    return "OK"


TOOL_DEFS = {
    "claim_done": {
        "description": "Call this tool when the task is fully completed.",
        "inputSchema": {"type": "object", "properties": {}},
        "fn": lambda a: _claim_done(),
    },
    "python_execute": {
        "description": "Execute Python code in the agent workspace and return stdout/stderr.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "code": {"type": "string", "description": "Python source code to execute."},
                "filename": {"type": "string",
                             "description": "Optional filename (with .py). A random UUID name is used if omitted."},
                "timeout": {"type": "integer",
                            "description": "Max execution time in seconds (capped at 120).", "default": 30},
            },
            "required": ["code"],
        },
        "fn": lambda a: _python_execute(a.get("code", ""), a.get("filename", ""), a.get("timeout", 30)),
    },
    "save_overlong_output": {
        "description": "Save a large text to disk and return a reference ID.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "content": {"type": "string", "description": "The large text content to store."},
                "label": {"type": "string", "description": "Optional human-readable label for this output."},
            },
            "required": ["content"],
        },
        "fn": lambda a: _save_overlong_output(a.get("content", ""), a.get("label", "")),
    },
    "view_overlong_output": {
        "description": "View a saved overlong output by ID, paginated.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "id": {"type": "string", "description": "The reference ID returned by save_overlong_output."},
                "page": {"type": "integer",
                         "description": "Page number (0-indexed, each page ~10000 chars).", "default": 0},
            },
            "required": ["id"],
        },
        "fn": lambda a: _view_overlong_output(a.get("id", ""), int(a.get("page", 0))),
    },
    "sleep": {
        "description": "Sleep for the given number of seconds.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "seconds": {"type": "number", "description": "Number of seconds to sleep (default 1).",
                            "default": 1},
            },
        },
        "fn": lambda a: _sleep(a.get("seconds", 1)),
    },
    "manage_context": {
        "description": "Manage conversation context (handled internally by the harness).",
        "inputSchema": {
            "type": "object",
            "properties": {"action": {"type": "string", "default": ""}},
        },
        "fn": _stub,
    },
    "history": {
        "description": "Inspect conversation history (handled internally by the harness).",
        "inputSchema": {
            "type": "object",
            "properties": {"action": {"type": "string", "default": ""}},
        },
        "fn": _stub,
    },
}


def _send(payload: dict):
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def _result(msg_id, result):
    _send({"jsonrpc": "2.0", "id": msg_id, "result": result})


def _error(msg_id, code, message):
    _send({"jsonrpc": "2.0", "id": msg_id, "error": {"code": code, "message": message}})


def _tools_list():
    tools = []
    for name in sorted(ENABLED):
        d = TOOL_DEFS[name]
        tools.append({"name": name, "description": d["description"],
                      "inputSchema": d["inputSchema"]})
    return tools


def _handle(msg: dict):
    method = msg.get("method")
    msg_id = msg.get("id")

    if method == "initialize":
        params = msg.get("params") or {}
        _result(msg_id, {
            "protocolVersion": params.get("protocolVersion", PROTOCOL_VERSION),
            "capabilities": {"tools": {}},
            "serverInfo": {"name": "toolathlon-local-tools", "version": "1.0.0"},
        })
    elif method == "ping":
        _result(msg_id, {})
    elif method == "tools/list":
        _result(msg_id, {"tools": _tools_list()})
    elif method == "tools/call":
        params = msg.get("params") or {}
        name = params.get("name", "")
        targs = params.get("arguments") or {}
        if name not in ENABLED or name not in TOOL_DEFS:
            _result(msg_id, {
                "content": [{"type": "text", "text": f"Unknown or disabled tool: {name}"}],
                "isError": True,
            })
            return
        try:
            text = TOOL_DEFS[name]["fn"](targs)
            _result(msg_id, {"content": [{"type": "text", "text": str(text)}]})
        except Exception as e:  # tool failures must not kill the server
            _result(msg_id, {
                "content": [{"type": "text", "text": f"Tool error: {e}"}],
                "isError": True,
            })
    elif method and method.startswith("notifications/"):
        pass  # no response for notifications
    elif msg_id is not None:
        _error(msg_id, -32601, f"Method not found: {method}")


def _compute_blocked_read_prefixes(workspace: str) -> None:
    """Extend the read block-list with workspace-external session internals.

    kimi_main.py lays out a task as ``<task_root>/workspace`` (the agent's
    declared boundary, and the cwd kimi-code runs under) alongside
    ``<task_root>/.kimi_home`` (session internals: wire.jsonl LLM transcripts,
    MCP config, per-agent tool-result offloads, etc.). The agent must never
    read outside its workspace. We block the kimi_home resolved from
    KIMI_CODE_HOME (set by the harness) plus the workspace's sibling/parent
    .kimi_home to cover both the standard and custom layouts.
    """
    global _BLOCKED_READ_PREFIXES
    seen = set(_BLOCKED_READ_PREFIXES)

    def _add(candidate: str) -> None:
        c = os.path.realpath(os.path.normpath(candidate))
        if c and c != workspace and not workspace.startswith(c + os.sep) and c not in seen:
            _BLOCKED_READ_PREFIXES.append(c)
            seen.add(c)

    env_home = os.environ.get("KIMI_CODE_HOME")
    if env_home:
        _add(env_home)
    if workspace:
        # Standard layout: .kimi_home is a sibling of workspace under task_root.
        _add(os.path.join(os.path.dirname(workspace), ".kimi_home"))
        # Defense-in-depth: a stray .kimi_home inside the workspace (should not
        # exist, but if it does it is not a legitimate deliverable).
        _add(os.path.join(workspace, ".kimi_home"))


def main():
    global WORKSPACE, MARKER, ENABLED
    ap = argparse.ArgumentParser()
    ap.add_argument("--workspace", required=True)
    ap.add_argument("--marker", default="")
    ap.add_argument("--tools", required=True, help="comma-separated tool names")
    args = ap.parse_args()

    WORKSPACE = os.path.abspath(args.workspace)
    MARKER = args.marker
    ENABLED = {t for t in args.tools.split(",") if t in TOOL_DEFS}
    unknown = {t for t in args.tools.split(",") if t and t not in TOOL_DEFS}
    if unknown:
        print(f"[local_tools] ignoring unknown tools: {sorted(unknown)}", file=sys.stderr)

    os.makedirs(WORKSPACE, exist_ok=True)
    _compute_blocked_read_prefixes(WORKSPACE)

    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            msg = json.loads(line)
        except json.JSONDecodeError:
            continue
        try:
            _handle(msg)
        except Exception as e:
            if msg.get("id") is not None:
                _error(msg["id"], -32603, f"Internal error: {e}")


if __name__ == "__main__":
    main()
