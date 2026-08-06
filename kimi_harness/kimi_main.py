"""kimi-code harness entry — runs INSIDE the enroot rootfs (cwd=/workspace).

Replaces the CAMEL TaskAgent portion of main.py while reusing Toolathlon's
TaskConfig / workspace setup / preprocess / TaskEvaluator untouched.

Flow:
  1. TaskConfig.build(...)            (paths identical to main.py)
  2. setup agent_workspace            (copy initial_workspace)
  3. run preprocess                   (resets PG schema state)
  4. generate per-task $KIMI_CODE_HOME (mcp.json + agents/ + config.toml)
  5. launch `kimi -p <task_str>` headless, stream-json to raw_stream.jsonl
  6. completion: claim_done marker | kimi exit | overall timeout
  7. write traj_log.json + traj.json  (evaluator-compatible)
  8. TaskEvaluator.evaluate_from_log_file
"""
import argparse
import asyncio
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

sys.path.insert(0, os.getcwd())

from utils.general.helper import read_json, print_color  # noqa: E402
from utils.data_structures.task_config import TaskConfig  # noqa: E402
# TaskEvaluator imported lazily in run_evaluation() — pulling it in at module
# level drags in camel (TaskStatus), which is unavailable outside the enroot
# rootfs and breaks standalone preview/rendering.

HARNESS_DIR = str(Path(__file__).resolve().parent)
CLAIM_TOOL = "mcp__local__claim_done"

CLAIM_DONE_PROTOCOL = """\
Completion Protocol (Mandatory):
- The task is **complete only after** you call `mcp__local__claim_done`.
- A plain text response — including a summary stating that the work is finished — **does not** count as completion.
- Do not stop after describing the result. You **must** call `mcp__local__claim_done` immediately after all required tool actions have succeeded.
- **Warning**: this call terminates the task instantly. You will have no further opportunity to act on this task afterward.\
"""

# --- Pluggable prompt sections (ablation-ready) -----------------------------
# Sub-agent type descriptions and examples are loaded from separate files so
# different subagent combinations can be swapped via env vars without touching
# code.  Set KIMI_SUBAGENTS="coder,explore" to drop plan, KIMI_EXAMPLES_FILE
# to an empty file to remove examples entirely, etc.

SECTIONS_DIR = os.path.join(HARNESS_DIR, "assets", "sections")
DEFAULT_SUBAGENTS = ("coder", "explore", "plan")

CORE_RESPONSIBILITIES_BASE = (
    "Core Responsibilities:\n"
    "- Efficiently and correctly solve the Toolathlon-GYM task using only the granted tools.\n"
    "- Ensure all required artifacts and external side effects are fully completed."
)
CORE_RESPONSIBILITIES_DELEGATION = (
    "- Assess whether sub-agent delegation is beneficial. Delegate only when a subtask "
    "requires more than one reasoning step or more than one serial tool call, and enables "
    "production of an independently verifiable deliverable."
)
CORE_RESPONSIBILITIES_PARALLEL = (
    "- Handle single-step work directly within the main agent; use parallel tool calls "
    "natively when sufficient."
)
CORE_RESPONSIBILITIES_TRUST = (
    "- Trust sub-agent outputs that satisfy their return contract. Reuse facts, artifact "
    "paths, resource IDs, digests, and verification results provided by sub-agents."
)


def _core_responsibilities(subagents: list) -> str:
    lines = [CORE_RESPONSIBILITIES_BASE]
    if subagents:
        lines += [CORE_RESPONSIBILITIES_DELEGATION, CORE_RESPONSIBILITIES_PARALLEL,
                  CORE_RESPONSIBILITIES_TRUST]
    else:
        lines.append(CORE_RESPONSIBILITIES_PARALLEL)
    return "\n".join(lines)

DELEGATION_RULES = """\
Subagent Delegation:
When assigning a sub-agent, provide a comprehensive and self-contained prompt that includes:
- Clear motivation and rationale.
- Sufficient prior context and background.
- Explicit constraints, scope, and limitations.
- Expected deliverables and acceptance criteria.

Keep the prompt concise and unambiguous — direct, not verbose.\
"""

ORCHESTRATION_RULES = """\
Sub-Agent Orchestration Rules:

- General Delegation:
  - Delegate focused subtasks to available sub-agents via the Agent tool. For parallel execution, use the AgentSwarm tool.
  - If a **Specified Sub-Agent Coordination** section appears later in this prompt, follow that prescribed workflow strictly when assigning and coordinating sub-agents; it overrides your default delegation judgment.

- Parallelism Guidelines:
  - For parallel sub-agents with different roles, issue multiple Agent tool calls within the same response.
  - For parallel sub-agents of one type, one prompt template, and distinct items, issue only the AgentSwarm tool call in that response (no other tool calls allowed).

- Environment & Context:
  - Sub-agents share the same task-approved tools and workspace directory as you — including terminal (shell), filesystem read/write, Python execution, database queries, spreadsheet/document creation, calendar/email, and any other tools granted by the current task.
  - Every sub-agent prompt must be **self-contained** — sub-agents operate in isolated contexts and cannot see the current user message or your previous reasoning steps.

- When & How to Delegate:
  - Delegate when a subtask is independent and would otherwise bloat your own context (e.g. exploring datasets, drafting documents, verifying intermediate results).
  - **Never delegate the final completion signal.** Only *you* may call `mcp__local__claim_done` after verifying that *every* requirement has been met.
  - Default to **foreground** sub-agents (`run_in_background=false`). Use `run_in_background=true` **only** for long-running work where you can proceed without the result immediately. Do not poll background agents, and do not restate a single background result unless integration requires it.
  - Prefer **resume** when an existing sub-agent already holds relevant context or the current task is a continuation of its prior work.\
"""


def _load_section(filename: str) -> str:
    path = os.path.join(SECTIONS_DIR, filename)
    if not os.path.exists(path):
        return ""
    with open(path, encoding="utf-8") as f:
        return f.read().strip()


def _active_subagents() -> list:
    """Return the list of sub-agent types to enable for this run.

    Unset → default (coder, explore, plan).
    Explicit empty string → none (disable all sub-agents).
    """
    if "KIMI_SUBAGENTS" not in os.environ:
        return list(DEFAULT_SUBAGENTS)
    return [s.strip() for s in os.environ["KIMI_SUBAGENTS"].split(",") if s.strip()]


def _subagent_types_section(subagents: list) -> str:
    """Load the subagent-types section file, filtered to active sub-agents."""
    raw = _load_section("subagent_types_default.md")
    if not raw:
        return ""
    lines = raw.splitlines()
    header = lines[0] if lines else "Available Sub-Agent Types:"
    kept = [header]
    for line in lines[1:]:
        # keep bullet lines that mention an active sub-agent name
        for name in subagents:
            if f"`{name}`" in line or line.strip().startswith("- " + name):
                kept.append(line)
                break
    if len(kept) <= 1:
        return ""
    return "\n".join(kept)


_SUBAGENT_NAMES = DEFAULT_SUBAGENTS  # coder, explore, plan
_EXAMPLE_BLOCK_RE = re.compile(r"<example>.*?</example>", re.DOTALL)


def _mentioned_subagents(text: str) -> set:
    """Return sub-agent names referenced via `name` backticks in *text*."""
    return {name for name in _SUBAGENT_NAMES if f"`{name}`" in text}


def _filter_examples_by_subagents(raw: str, subagents: list) -> str:
    """Keep only <example> blocks whose referenced sub-agents are all active."""
    if not raw.strip():
        return ""
    active = set(subagents)
    blocks = _EXAMPLE_BLOCK_RE.findall(raw)
    if not blocks:
        return raw
    kept = []
    for block in blocks:
        mentioned = _mentioned_subagents(block)
        if not mentioned or mentioned <= active:
            kept.append(block.strip())
    if not kept:
        return ""
    return "Examples:\n\n" + "\n\n".join(kept)


def _examples_section(subagents: list | None = None) -> str:
    """Load examples; filter <example> blocks to match *subagents* when set."""
    fname = os.environ.get("KIMI_EXAMPLES_FILE", "examples_default.md")
    raw = _load_section(fname)
    if not raw:
        return ""
    if all(line.lstrip().startswith("#") or not line.strip() for line in raw.splitlines()):
        return ""
    subagents = subagents if subagents is not None else _active_subagents()
    return _filter_examples_by_subagents(raw, subagents)


def _coordination_section() -> str:
    """Load optional prescribed sub-agent coordination workflow (ablation-ready).

    Default file is empty → section omitted. Set KIMI_COORDINATION_FILE to a
    non-empty markdown file to inject hard-coded coordination instructions
    immediately before Completion Protocol.
    """
    fname = os.environ.get("KIMI_COORDINATION_FILE", "subagent_coordination_default.md")
    raw = _load_section(fname)
    if not raw:
        return ""
    if all(line.lstrip().startswith("#") or not line.strip() for line in raw.splitlines()):
        return ""
    first = raw.splitlines()[0].strip()
    if first.lower().startswith("specified sub-agent coordination"):
        return raw
    return "Specified Sub-Agent Coordination (Mandatory):\n\n" + raw

# Read-only tool selection for `explore` / `plan` sub-agents.
#
# kimi matches MCP tool entries with picomatch, so `mcp__<server>__read_*`
# is a real prefix glob (verified in agent-core-v2/.../evaluate.ts). We
# allowlist by prefix so explore/plan can never see write/create/delete
# tools even if the model tries to call them — the tools are not in the
# schema sent to the model, and the executor re-checks before each call.
READONLY_TOOL_PREFIXES = (
    "read_", "list_", "search_", "get_", "view_", "query_", "fetch_",
    "describe_", "info_", "count_", "exists_", "find_", "select_",
    "inspect_", "show_", "lookup_",
)
# Read-only tools whose names don't share a common read prefix.
READONLY_TOOL_EXACT = {
    "directory_tree",        # filesystem MCP: read-only tree listing
    "list_allowed_directories",  # already covered by list_, kept for clarity
}


def build_readonly_tool_patterns(server_names):
    """Return mcp__<server>__<prefix>* + exact patterns for read-only tools."""
    patterns = []
    for s in server_names:
        for p in READONLY_TOOL_PREFIXES:
            patterns.append(f"mcp__{s}__{p}*")
        for name in READONLY_TOOL_EXACT:
            patterns.append(f"mcp__{s}__{name}")
    return patterns


def render_subagent_profile(name, tools_list, disallowed=None):
    """Read assets/subagents/<name>.md, keep its body, rewrite frontmatter
    with an explicit read-only tool allowlist."""
    src = os.path.join(HARNESS_DIR, "assets", "subagents", f"{name}.md")
    with open(src, encoding="utf-8") as f:
        raw = f.read()
    # split frontmatter (---\n...\n---\n) from body
    body = raw
    if raw.startswith("---"):
        parts = raw.split("---", 2)
        if len(parts) >= 3:
            body = parts[2].lstrip("\n")
    tools_yaml = "\n".join(f"  - {t}" for t in tools_list)
    disallowed_yaml = ""
    if disallowed:
        disallowed_yaml = "\ndisallowedTools:\n" + "\n".join(f"  - {t}" for t in disallowed)
    return (
        "---\n"
        f"name: {name}\n"
        f"override: true\n"
        "tools:\n"
        f"{tools_yaml}"
        f"{disallowed_yaml}\n"
        "subagents: []\n"
        "---\n\n"
        f"{body}\n"
    )


def copy_folder_contents(src: str, dst: str):
    for item in os.listdir(src):
        s, d = os.path.join(src, item), os.path.join(dst, item)
        if os.path.isdir(s):
            shutil.copytree(s, d, dirs_exist_ok=True)
        else:
            shutil.copy2(s, d)


def setup_workspace(task_config) -> str:
    workspace = os.path.abspath(task_config.agent_workspace)
    os.makedirs(workspace, exist_ok=True)
    init = task_config.initialization
    if init and init.workspace and os.path.exists(str(init.workspace)):
        copy_folder_contents(str(init.workspace), workspace)
    for srv, d in [("arxiv_local", "arxiv_local_storage"),
                   ("memory", "memory"),
                   ("playwright_with_chunk", ".playwright_output")]:
        if srv in (task_config.needed_mcp_servers or []):
            os.makedirs(os.path.join(workspace, d), exist_ok=True)
    return workspace


def run_preprocess(task_config, debug: bool):
    init = task_config.initialization
    if init and init.process_command:
        cmd = init.process_command
        cmd += f" --agent_workspace {task_config.agent_workspace}"
        lt = task_config.launch_time or ""
        lt_clean = " ".join(lt.split()[:2])
        cmd += f" --launch_time \"{lt_clean}\""
        print_color("[preprocess] running...", "yellow")
        r = subprocess.run(cmd, shell=True, capture_output=not debug, text=True)
        if r.returncode != 0:
            print_color(f"[preprocess] failed: {(r.stderr or '')[:300]}", "red")
        else:
            print_color("[preprocess] done.", "green")


# Redundant sentence in task-provided agent_system_prompt.md that duplicates
# CLAIM_DONE_PROTOCOL; strip it to keep the prompt tight.
_REDUNDANT_CLAIM_SENTENCE = (
    "If you believe the task is completed, you can either call the "
    "`{claim}` tool or respond without calling any tool to indicate "
    "completion. This will immediately terminate the task, and you will "
    "have no further opportunity to work on it."
)


def render_system_prompt(task_config) -> str:
    ws = os.path.abspath(task_config.agent_workspace)
    if task_config.system_prompts and task_config.system_prompts.agent:
        sp = task_config.system_prompts.agent
        sp = sp.replace("local-claim_done", CLAIM_TOOL)
        sp = sp.replace("claim_done", CLAIM_TOOL) if CLAIM_TOOL not in sp else sp
    else:
        sp = (f"You are a helpful AI assistant. Your workspace directory is: {ws}\n"
              "Complete the user's task using the provided tools.")

    # Strip the redundant claim_done sentence (both original and rewritten forms)
    for variant in ("local-claim_done", CLAIM_TOOL):
        redundant = _REDUNDANT_CLAIM_SENTENCE.format(claim=variant)
        sp = sp.replace(redundant, "")

    subagents = _active_subagents()
    sections = [sp.rstrip(), "", _core_responsibilities(subagents)]
    if subagents:
        sections += ["", DELEGATION_RULES, "", ORCHESTRATION_RULES]
        types_sec = _subagent_types_section(subagents)
        if types_sec:
            sections += ["", types_sec]
        examples_sec = _examples_section(subagents)
        if examples_sec:
            sections += ["", examples_sec]
        coord_sec = _coordination_section()
        if coord_sec:
            sections += ["", coord_sec]
    sections += ["", CLAIM_DONE_PROTOCOL]
    return "\n".join(sections)


def write_agentfile(path: str, system_prompt: str):
    body = system_prompt.replace("\\", "\\\\")
    subagents = _active_subagents()
    if subagents:
        tools_yaml = (
            "  - mcp__*\n"
            "  - Agent\n"
            "  - AgentSwarm\n"
            "  - TodoList"
        )
        subagents_yaml = "subagents:\n" + "\n".join(f"  - {s}" for s in subagents)
    else:
        tools_yaml = "  - mcp__*\n  - TodoList"
        subagents_yaml = "subagents: []"
    content = f"""---
name: toolathlon-main
description: Toolathlon-GYM task solver (main-agent)
tools:
{tools_yaml}
{subagents_yaml}
---

{body}
"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)


def write_kimi_home(home: str, task_config, workspace: str, marker: str,
                    max_steps: int):
    os.makedirs(home, exist_ok=True)
    os.makedirs(os.path.join(home, "agents"), exist_ok=True)
    os.makedirs(os.path.join(home, "plugins"), exist_ok=True)

    from kimi_harness.mcp_json_gen import build_mcp_servers, local_tools_entry

    servers = build_mcp_servers(
        task_config.needed_mcp_servers or [],
        agent_workspace=workspace,
        task_src_dir=os.path.abspath(os.path.join("tasks/finalpool", task_config.task_dir)),
        config_dir=os.path.abspath("configs/mcp_servers"),
    )
    local_tools = task_config.needed_local_tools or []
    if local_tools:
        servers["local"] = local_tools_entry(
            workspace=workspace,
            marker=marker,
            tools=local_tools,
            python_bin=os.environ.get("PYTHON_BIN", "/opt/venv/bin/python3"),
            harness_dir=HARNESS_DIR,
        )
    with open(os.path.join(home, "mcp.json"), "w", encoding="utf-8") as f:
        json.dump({"mcpServers": servers}, f, indent=2, ensure_ascii=False)

    # Sub-agent profiles. coder keeps the full MCP set (minus claim_done);
    # explore/plan get an explicit read-only allowlist computed from the
    # task's MCP servers so write/create/delete tools are not even visible.
    active = _active_subagents()
    server_names = list(servers.keys())
    readonly_patterns = build_readonly_tool_patterns(server_names)
    if "coder" in active:
        coder_path = os.path.join(HARNESS_DIR, "assets", "subagents", "coder.md")
        shutil.copy2(coder_path, os.path.join(home, "agents", "coder.md"))
    for name in ("explore", "plan"):
        if name not in active:
            continue
        content = render_subagent_profile(
            name,
            tools_list=readonly_patterns,
            disallowed=[CLAIM_TOOL],
        )
        with open(os.path.join(home, "agents", f"{name}.md"), "w", encoding="utf-8") as f:
            f.write(content)

    provider = os.environ.get("KIMI_PROVIDER_NAME", "remote-eval")
    model_alias = os.environ.get("KIMI_MODEL_ALIAS", "eval-model")
    base_url = os.environ["MODEL_API_URL"].rstrip("/")
    if not base_url.endswith("/v1"):
        base_url += "/v1"
    config = f"""default_model = "{model_alias}"

[providers."{provider}"]
type = "openai"
base_url = "{base_url}"
api_key = "{os.environ['MODEL_API_KEY']}"

[models."{model_alias}"]
provider = "{provider}"
model = "{os.environ['MODEL_NAME']}"
max_context_size = {int(os.environ.get('KIMI_MAX_CONTEXT', '262144'))}

[loop_control]
max_steps_per_turn = {max_steps}

[mcp]
startup_timeout_ms = {int(float(os.environ.get('MCP_STDIO_TIMEOUT_MIN', '90')) * 1000)}
"""
    with open(os.path.join(home, "config.toml"), "w", encoding="utf-8") as f:
        f.write(config)


def launch_kimi(task_str: str, agentfile: str, home: str, stream_path: str,
                workspace: str, marker: str, timeout_s: int, debug: bool):
    env = dict(os.environ)
    env["KIMI_CODE_HOME"] = home
    env["KIMI_CODE_BUILTIN_PRODUCT_SKILLS"] = "0"
    env.pop("http_proxy", None)
    env.pop("https_proxy", None)
    env.pop("HTTP_PROXY", None)
    env.pop("HTTPS_PROXY", None)

    cmd = ["kimi", "-p", task_str,
           "--agent-file", agentfile,
           "--output-format", "stream-json"]
    print_color(f"[kimi] launching: {' '.join(cmd[:3])} ... (home={home})", "cyan")

    with open(stream_path, "w", encoding="utf-8") as sf:
        proc = subprocess.Popen(
            cmd, cwd=workspace, env=env,
            stdout=sf, stderr=subprocess.STDOUT,
            start_new_session=True,
        )
        start = time.time()
        marker_seen_at = None
        while True:
            rc = proc.poll()
            if rc is not None:
                break
            if marker_seen_at is None and os.path.exists(marker):
                marker_seen_at = time.time()
                print_color("[kimi] claim_done marker detected; grace period ...", "green")
            if marker_seen_at is not None and time.time() - marker_seen_at > 15:
                print_color("[kimi] grace over, terminating.", "yellow")
                break
            if time.time() - start > timeout_s:
                print_color(f"[kimi] overall timeout {timeout_s}s, terminating.", "red")
                break
            time.sleep(1.0)
        if proc.poll() is None:
            try:
                proc.terminate()
                proc.wait(timeout=10)
            except Exception:
                proc.kill()
                proc.wait(timeout=10)
        return proc.returncode, marker_seen_at is not None or os.path.exists(marker)


def write_traj_logs(task_config, status: str, start_time: datetime,
                    stream_path: str):
    log_path = task_config.log_file
    Path(log_path).parent.mkdir(parents=True, exist_ok=True)
    record = {
        "config": task_config.to_dict(),
        "status": status,
        "start_time": start_time.isoformat(),
        "end_time": datetime.now().isoformat(),
    }
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(record, f, ensure_ascii=False, indent=2)

    from kimi_harness.traj_convert import convert_stream
    traj = convert_stream(stream_path)
    traj.update({
        "status": status,
        "start_time": record["start_time"],
        "end_time": record["end_time"],
    })
    traj_path = str(Path(log_path).parent / "traj.json")
    with open(traj_path, "w", encoding="utf-8") as f:
        json.dump(traj, f, ensure_ascii=False, indent=2)


async def amain():
    parser = argparse.ArgumentParser()
    parser.add_argument("--eval_config", default="scripts/eval_config.json")
    parser.add_argument("--task_dir", required=True)
    parser.add_argument("--max_steps", type=int, default=None)
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--cn_mode", action="store_true")
    args = parser.parse_args()

    cfg = read_json(args.eval_config)
    dump_path = cfg.get("dump_path", "./dumps/")
    max_steps = args.max_steps or cfg.get("global_task_config", {}).get(
        "max_steps_under_single_turn_mode", 100)

    model_name = os.environ.get("MODEL_NAME", "unknown-model")

    task_config = TaskConfig.build(
        args.task_dir,
        agent_short_name=f"kimi-code/{model_name}",
        global_task_config={"dump_path": dump_path,
                            "max_steps_under_single_turn_mode": max_steps},
        single_turn_mode=True,
        cn_mode=args.cn_mode,
    )

    print_color(f"====== {args.task_dir} | kimi-code/{model_name} | steps={max_steps} ======",
                "yellow")
    print_color(f"workspace : {task_config.agent_workspace}", "cyan")

    start_time = datetime.now()
    task_root = task_config.task_root
    os.makedirs(task_root, exist_ok=True)
    marker = os.path.join(task_root, ".claim_done")
    if os.path.exists(marker):
        os.remove(marker)

    workspace = setup_workspace(task_config)
    run_preprocess(task_config, args.debug)

    kimi_home = os.path.join(task_root, ".kimi_home")
    write_kimi_home(kimi_home, task_config, workspace, marker, max_steps)

    agentfile = os.path.join(task_root, "agent_main.md")
    write_agentfile(agentfile, render_system_prompt(task_config))

    stream_path = os.path.join(task_root, "raw_stream.jsonl")
    timeout_s = int(os.environ.get("KIMI_TASK_TIMEOUT_S", "7200"))
    rc, done = launch_kimi(
        task_str=task_config.task_str,
        agentfile=agentfile,
        home=kimi_home,
        stream_path=stream_path,
        workspace=workspace,
        marker=marker,
        timeout_s=timeout_s,
        debug=args.debug,
    )

    status = "success" if done else "failed"
    print_color(f"[kimi] exited rc={rc} claim_done={done} -> status={status}",
                "green" if done else "red")

    write_traj_logs(task_config, status, start_time, stream_path)

    print_color("\n====== Evaluating ======", "yellow")
    from utils.evaluation.evaluator import TaskEvaluator
    eval_res = await TaskEvaluator.evaluate_from_log_file(task_config.log_file)
    print(f"Pass:    {eval_res.get('pass', False)}")
    print(f"Details: {eval_res.get('details', 'N/A')}")
    return 0 if eval_res.get("pass", False) else 1


def main():
    raise SystemExit(asyncio.run(amain()))


if __name__ == "__main__":
    main()
