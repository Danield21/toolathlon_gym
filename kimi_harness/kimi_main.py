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
import ctypes
import fnmatch
import glob
import html
import http.server
import json
import mimetypes
import os
import posixpath
import re
import shlex
import shutil
import signal
import subprocess
import sys
import threading
import time
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

import yaml

sys.path.insert(0, os.getcwd())

from utils.general.helper import read_json, print_color  # noqa: E402
from utils.data_structures.task_config import TaskConfig  # noqa: E402
# TaskEvaluator imported lazily in run_evaluation() — pulling it in at module
# level drags in camel (TaskStatus), which is unavailable outside the enroot
# rootfs and breaks standalone preview/rendering.

HARNESS_DIR = str(Path(__file__).resolve().parent)
CLAIM_TOOL = "mcp__local__claim_done"
BUILTIN_DISALLOWED_TOOLS = (
    "Bash",
    "Shell",
    "Terminal",
    "Read",
    "ReadMediaFile",
    "Write",
    "Edit",
    "Glob",
    "Grep",
    "WebFetch",
    "WebSearch",
    "FetchURL",
    "Fetch",
    "Browser",
)

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
# code.  Set KIMI_SUBAGENTS="plan,academic-literature-researcher" to pin a
# subset, KIMI_EXAMPLES_FILE to an empty file to remove examples entirely, etc.
#
# 2026-08-21 roster redesign: coder and explore are RETIRED from the final
# design.  The 10-agent roster is plan + 7 domain specialists + 2 cross-cutting
# agents (evidence-integrator, deliverable-auditor); select it explicitly with
# KIMI_SUBAGENTS=ten (aliases: 10 / all).  The DEFAULT (unset) stays on the
# legacy trio (coder/explore/plan) so launch scripts that predate the redesign
# keep their original behavior; KIMI_SUBAGENTS=three (3) pins the same trio.

SECTIONS_DIR = os.path.join(HARNESS_DIR, "assets", "sections")
LEGACY_SUBAGENTS = ("coder", "explore", "plan")  # retired; kept for the "three" preset
SPECIALIZED_SUBAGENTS = (
    "academic-literature-researcher",
    "web-domain-researcher",
    "enterprise-data-analyst",
    "financial-market-analyst",
    "workspace-data-engineer",
    "office-report-builder",
    "external-workflow-operator",
)
# Cross-cutting function agents (not bound to one tool domain): merge parallel
# evidence, and independently audit finished deliverables.
CROSSCUT_SUBAGENTS = (
    "evidence-integrator",
    "deliverable-auditor",
)
PROFILE_SUBAGENTS = SPECIALIZED_SUBAGENTS + CROSSCUT_SUBAGENTS
TEN_SUBAGENTS = ("plan",) + PROFILE_SUBAGENTS
# Default (unset) stays the legacy trio: existing launch scripts keep producing
# pre-redesign prompts; opt in to the new roster with KIMI_SUBAGENTS=ten.
DEFAULT_SUBAGENTS = LEGACY_SUBAGENTS
# Named presets accepted by _active_subagents() (case-insensitive).
SUBAGENT_PRESETS = {
    "ten": list(TEN_SUBAGENTS),
    "10": list(TEN_SUBAGENTS),
    "all": list(TEN_SUBAGENTS),
    "three": list(LEGACY_SUBAGENTS),
    "3": list(LEGACY_SUBAGENTS),
}

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
    "- Reuse sub-agent outputs only when they satisfy their named versioned return "
    "contract. Preserve facts, artifact paths, resource IDs, digests, provenance, and "
    "verification results without silently reshaping them."
)
CORE_RESPONSIBILITIES_SPECIALIST = (
    "- When a domain specialist matches the sub-task (papers, web evidence, enterprise "
    "data, market data, workspace engineering, office artifacts, external workflows), "
    "delegate to the narrowest matching profile. When no specialist matches, do the "
    "work yourself rather than inventing a generalist delegate."
)
CORE_RESPONSIBILITIES_CROSSCUT = (
    "- Route cross-cutting phases to their dedicated agents: merge parallel evidence "
    "packets through `evidence-integrator` instead of reconciling them in your own "
    "context, and, when the audit threshold applies, obtain an independent per-criterion "
    "verdict from `deliverable-auditor` before signaling completion."
)


def _boundary_section(workspace: str) -> str:
    """Explicit visible-boundary statement injected into every agent prompt."""
    return (
        "Visible Boundary (Strictly Enforced):\n"
        f"- You may ONLY read and write files inside your accessible workspace: {workspace}\n"
        "- You may use ONLY the task-granted tools to interact with external systems "
        "(databases, calendars, email, spreadsheets, etc.).\n"
        "- Do NOT access, read, list, or probe anything outside your workspace, including "
        "tool/MCP server source code, evaluation logic, ground-truth data, harness or "
        "benchmark internals, or any system paths (e.g. under /opt, /workspace/tasks, "
        "/workspace/utils). Those are outside your task boundary.\n"
        "- Assume the environment and all granted tools are functioning correctly. Never "
        "attempt to diagnose, fix, or work around infrastructure; if a granted tool fails, "
        "report the failure and continue with the tools that work."
    )


def _core_responsibilities(subagents: list, workspace: str = "") -> str:
    lines = [CORE_RESPONSIBILITIES_BASE]
    if subagents:
        lines += [CORE_RESPONSIBILITIES_DELEGATION, CORE_RESPONSIBILITIES_PARALLEL,
                  CORE_RESPONSIBILITIES_TRUST]
        if any(s in PROFILE_SUBAGENTS for s in subagents):
            lines.append(CORE_RESPONSIBILITIES_SPECIALIST)
        if any(s in CROSSCUT_SUBAGENTS for s in subagents):
            lines.append(CORE_RESPONSIBILITIES_CROSSCUT)
    else:
        lines.append(CORE_RESPONSIBILITIES_PARALLEL)
    out = "\n".join(lines)
    if workspace:
        out += "\n\n" + _boundary_section(workspace)
    return out

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
  - By default, sub-agents inherit the same task-scoped tools and workspace permissions as the parent agent, including terminal access, filesystem read/write operations, Python execution, database interactions, spreadsheet/document generation, calendar/email capabilities, and other tools enabled for the current task. However, if a sub-agent YAML configuration defines customized tool permissions or workspace directory access, the customized configuration overrides the inherited defaults.
  - Current roster YAML customizations (the Agent tool description lists each type's Tools; do not assume any sub-agent has every tool you have):
    - Write-capable specialists (`workspace-data-engineer`, `office-report-builder`, `external-workflow-operator`) receive a task-scoped write tool ceiling for their domain.
    - Research specialists, `plan`, and `deliverable-auditor` operate **read-only** over their domains. They cannot create, update, delete, send, or otherwise mutate persistent state (Notion pages, email, calendar events, Canvas/WooCommerce writes, spreadsheet writes, etc.).
    - `evidence-integrator` cannot query domain systems or mutate external state, but it may write one named canonical intermediate dataset inside the task workspace.
    - Do not ask a read-only agent to call a write tool.
  - Every sub-agent prompt must be **self-contained** — sub-agents operate in isolated contexts and cannot see the current user message or your previous reasoning steps.

- Role Boundaries & Staged Handoffs:
  - Select the narrowest profile whose semantic responsibility and task-scoped tool ceiling both fit the subtask; tool availability alone is not a reason to select a role.
  - When one apparent subtask crosses profile boundaries, split it into **producer and consumer phases**: a domain researcher returns EvidencePacket v1, then `workspace-data-engineer`, `office-report-builder`, or `external-workflow-operator` consumes the frozen packet and returns its versioned artifact packet or **DeliverableReceipt v1**.
  - **Never broaden a profile's tool ceiling** to keep a cross-domain subtask monolithic. If a step is genuinely atomic and no profile fits, execute that residual step in the main agent with the task-granted tools.

- Evidence Integration:
  - Research specialists return **EvidencePacket v1**. When two or more parallel packets must feed shared deliverables, delegate their merge to `evidence-integrator`; do not reconcile multi-packet joins ad hoc in your own context.
  - Inline small packets verbatim. Packets larger than 32 KiB, packets already materialized by a sub-agent, or packets whose duplication would materially bloat context must be passed as **named workspace files** with path, size, and digest. If a required packet cannot be materialized or its size and digest cannot be computed, you must stop before integration and report the blocker. The integrator holds no domain tools and cannot re-fetch missing facts.
  - Treat **CanonicalEvidence v1** from the integrator as the single source for downstream construction; resolve its `conflicts` and `unresolved` lists yourself before dispatching builders.
  - Skip the integrator when a task has only one evidence branch or the packets need no joining — project directly.

- Pre-Completion Audit:
  - Before calling `mcp__local__claim_done`, delegate `deliverable-auditor` **whenever deliverables span two or more systems or the acceptance list has several independently checkable criteria**; hand it the per-criterion checklist plus the deliverable inventory (paths/IDs) and require **AuditReport v1**.
  - Treat the audit as fail-closed: repair every FAIL and re-audit the affected criteria; investigate every UNKNOWN yourself with authoritative read-backs. You **must not call `mcp__local__claim_done`** while any applicable criterion remains FAIL or UNKNOWN. The gate opens only when every criterion is PASS.
  - For a single trivial deliverable, verify it yourself with read-backs instead — the audit must earn its cost.

- When & How to Delegate:
  - Delegate when a subtask is independent and would otherwise bloat your own context (e.g. researching sources, drafting documents, verifying intermediate results).
  - **Never delegate the final completion signal.** Only *you* may call `mcp__local__claim_done` after verifying that *every* requirement has been met (directly, or via the auditor's verdict).
  - Default to **foreground** sub-agents (`run_in_background=false`). Use `run_in_background=true` **only** for long-running work where you can proceed without the result immediately. Do not poll background agents, and do not restate a single background result unless integration requires it.
  - Prefer **resume** when an existing sub-agent already holds relevant context or the current task is a continuation of its prior work.\
"""

# Three-agent orchestration text. Environment & Context uses the same
# inherit-then-YAML-override wording as the ten-agent rules; the rest of this
# block stays on the pre-redesign structure so KIMI_SUBAGENTS=three still
# selects the legacy catalog/examples rather than the specialist roster.
ORCHESTRATION_RULES_LEGACY = """\
Sub-Agent Orchestration Rules:

- General Delegation:
  - Delegate focused subtasks to available sub-agents via the Agent tool. For parallel execution, use the AgentSwarm tool.
  - If a **Specified Sub-Agent Coordination** section appears later in this prompt, follow that prescribed workflow strictly when assigning and coordinating sub-agents; it overrides your default delegation judgment.

- Parallelism Guidelines:
  - For parallel sub-agents with different roles, issue multiple Agent tool calls within the same response.
  - For parallel sub-agents of one type, one prompt template, and distinct items, issue only the AgentSwarm tool call in that response (no other tool calls allowed).

- Environment & Context:
  - By default, sub-agents inherit the same task-scoped tools and workspace permissions as the parent agent, including terminal access, filesystem read/write operations, Python execution, database interactions, spreadsheet/document generation, calendar/email capabilities, and other tools enabled for the current task. However, if a sub-agent YAML configuration defines customized tool permissions or workspace directory access, the customized configuration overrides the inherited defaults.
  - Every sub-agent prompt must be **self-contained** — sub-agents operate in isolated contexts and cannot see the current user message or your previous reasoning steps.

- When & How to Delegate:
  - Delegate when a subtask is independent and would otherwise bloat your own context (e.g. exploring datasets, drafting documents, verifying intermediate results).
  - **Never delegate the final completion signal.** Only *you* may call `mcp__local__claim_done` after verifying that *every* requirement has been met.
  - Default to **foreground** sub-agents (`run_in_background=false`). Use `run_in_background=true` **only** for long-running work where you can proceed without the result immediately. Do not poll background agents, and do not restate a single background result unless integration requires it.
  - Prefer **resume** when an existing sub-agent already holds relevant context or the current task is a continuation of its prior work.\
"""

# Retired pre-redesign agents: their presence in the active roster switches the
# whole prompt (orchestration rules + section files) back to the legacy variant.
RETIRED_SUBAGENTS = ("coder", "explore")


def _is_legacy_roster(subagents: list) -> bool:
    return any(s in RETIRED_SUBAGENTS for s in subagents)


def _orchestration_rules(subagents: list) -> str:
    return ORCHESTRATION_RULES_LEGACY if _is_legacy_roster(subagents) else ORCHESTRATION_RULES


def _load_section(filename: str) -> str:
    path = os.path.join(SECTIONS_DIR, filename)
    if not os.path.exists(path):
        return ""
    with open(path, encoding="utf-8") as f:
        return f.read().strip()


def _active_subagents() -> list:
    """Return the list of sub-agent types to enable for this run.

    Unset → legacy trio (coder/explore/plan) — the default, matching all
    pre-redesign runs.
    "ten" / "10" / "all" → new 10-agent roster (plan + 7 specialists +
    integrator/auditor), pinned explicitly.
    "three" / "3" → same legacy trio, pinned explicitly.
    Explicit empty string → none (disable all sub-agents).
    Comma list → exactly those agents (e.g. "plan,evidence-integrator").
    """
    if "KIMI_SUBAGENTS" not in os.environ:
        return list(DEFAULT_SUBAGENTS)
    raw = os.environ["KIMI_SUBAGENTS"]
    preset = SUBAGENT_PRESETS.get(raw.strip().lower())
    if preset is not None:
        return list(preset)
    return [s.strip() for s in raw.split(",") if s.strip()]


def _subagent_types_section(subagents: list) -> str:
    """Load the subagent-types section file, filtered to active sub-agents."""
    fname = "subagent_types_default.md"
    if _is_legacy_roster(subagents):
        legacy = _load_section("subagent_types_legacy.md")
        if legacy:
            fname = "subagent_types_legacy.md"
    raw = _load_section(fname)
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


# Full roster for example-block filtering: current 10 plus the retired trio's
# coder/explore so legacy example files filter correctly under KIMI_SUBAGENTS=three.
_SUBAGENT_NAMES = TEN_SUBAGENTS + RETIRED_SUBAGENTS
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
    subagents = subagents if subagents is not None else _active_subagents()
    if "KIMI_EXAMPLES_FILE" in os.environ:
        fname = os.environ["KIMI_EXAMPLES_FILE"]
    else:
        fname = ("examples_legacy.md" if _is_legacy_roster(subagents)
                 else "examples_default.md")
        if _is_legacy_roster(subagents) and not _load_section(fname):
            fname = "examples_default.md"
    raw = _load_section(fname)
    if not raw:
        return ""
    if all(line.lstrip().startswith("#") or not line.strip() for line in raw.splitlines()):
        return ""
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


def _plan_first_enabled() -> bool:
    """Plan-first arm switch (ablation-ready).

    Unset / empty / 0 → off (baseline subagent arm).
    Any other value → inject the plan-first section into the main prompt.
    Requires sub-agents (plan must be active); silently off otherwise.
    """
    val = os.environ.get("KIMI_PLAN_FIRST", "").strip().lower()
    return val not in ("", "0", "false", "no", "off")


def _plan_first_section(subagents: list) -> str:
    """Load the plan-first mandate section; '' when disabled or inapplicable."""
    if not _plan_first_enabled():
        return ""
    if not subagents or "plan" not in subagents:
        return ""
    if "KIMI_PLAN_FIRST_FILE" in os.environ:
        fname = os.environ["KIMI_PLAN_FIRST_FILE"]
    else:
        fname = ("plan_first_legacy.md" if _is_legacy_roster(subagents)
                 else "plan_first_default.md")
        if _is_legacy_roster(subagents) and not _load_section(fname):
            fname = "plan_first_default.md"
    return _load_section(fname)

# Read-only tool selection for `explore` / `plan` sub-agents.
#
# kimi matches MCP names with picomatch against the FULL `mcp__<server>__<tool>`
# string (agent-core-v2/.../evaluate.ts). A prefix-only glob such as
# `mcp__canvas__get_*` misses real names like `canvas_get_course_grades`,
# `woo_products_list`, `API-get-user`, `videos_getVideo`, `get-tickets`,
# `search-arxiv`, `mcp_howtocook_getAllRecipes`.
#
# Allow by read-verb *substring* (`*{verb}*`) so those layouts match, then
# deny `write_*` so snowflake `write_query` is not pulled in by `*query*`.
# `select` stays prefix-only: `*select*` would match playwright
# `browser_select_option`.
READONLY_TOOL_VERBS = (
    "read", "list", "search", "get", "view", "query", "fetch",
    "describe", "info", "count", "exists", "find", "inspect",
    "show", "lookup", "retrieve", "snapshot",
)
READONLY_TOOL_PREFIX_ONLY_VERBS = ("select",)
# Names that are read-only but contain none of the verbs above.
READONLY_TOOL_EXTRA_GLOBS = (
    "directory_tree",
    "list_allowed_directories",
    "*_reports",
    "*_reports_*",
    "reports_*",
    "*_health_check",
    "*_system_status",
    "browser_navigate",
    "browser_navigate_back",
    "browser_navigate_forward",
    "browser_console_messages",
    "browser_network_requests",
    "browser_take_screenshot",
    "browser_wait_for",
    "*recommend*",
    "*whatToEat",
)
# Defense in depth: keep write_query / write_file out even if an allow glob
# would otherwise match (e.g. `*query*` vs snowflake write_query).
READONLY_TOOL_DENY_PATTERNS = (
    "mcp__*__write_*",
    "mcp__*__*_write_*",
    "mcp__*__*_write",
)


def build_readonly_tool_patterns(server_names):
    """Return picomatch globs that grant read-oriented MCP tools for explore/plan."""
    patterns = []
    seen = set()

    def add(pattern):
        if pattern not in seen:
            seen.add(pattern)
            patterns.append(pattern)

    for server in server_names:
        for verb in READONLY_TOOL_VERBS:
            add(f"mcp__{server}__*{verb}*")
        for verb in READONLY_TOOL_PREFIX_ONLY_VERBS:
            add(f"mcp__{server}__{verb}_*")
        for extra in READONLY_TOOL_EXTRA_GLOBS:
            add(f"mcp__{server}__{extra}")
    return patterns


def mcp_tool_granted_to_readonly_agent(name, allow_patterns, deny_patterns=None):
    """Whether an MCP tool name would be active under explore/plan policy.

    Uses fnmatch, which agrees with kimi's picomatch on these `*` / `mcp__`
    patterns (no `**`, no `/` in names).
    """
    deny_patterns = (
        deny_patterns if deny_patterns is not None else READONLY_TOOL_DENY_PATTERNS
    )
    if not any(fnmatch.fnmatchcase(name, pattern) for pattern in allow_patterns):
        return False
    return not any(fnmatch.fnmatchcase(name, pattern) for pattern in deny_patterns)


_SUBAGENT_BOUNDARY = (
    "Visible Boundary (Strictly Enforced):\n"
    "- Read and write files ONLY inside the task workspace directory.\n"
    "- Interact with external systems ONLY through the task-granted tools.\n"
    "- Do NOT access, read, list, or probe anything outside the workspace, including "
    "tool/MCP server source code, evaluation logic, ground-truth data, harness or "
    "benchmark internals, or system paths (e.g. under /opt, /workspace/tasks, "
    "/workspace/utils).\n"
    "- Assume the environment and all granted tools work correctly. Never diagnose, fix, "
    "or work around infrastructure; if a granted tool fails, report it and continue."
)


def render_subagent_profile(name, tools_list, disallowed=None):
    """Read assets/subagents/<name>.md, keep its body, rewrite frontmatter
    with an explicit read-only tool allowlist."""
    src = os.path.join(HARNESS_DIR, "assets", "subagents", f"{name}.md")
    with open(src, encoding="utf-8") as f:
        raw = f.read()
    # split frontmatter (---\n...\n---\n) from body
    body = raw
    metadata = {}
    if raw.startswith("---"):
        parts = raw.split("---", 2)
        if len(parts) >= 3:
            metadata = yaml.safe_load(parts[1]) or {}
            body = parts[2].lstrip("\n")
    description = metadata.get("description")
    if not isinstance(description, str) or not description.strip():
        raise ValueError(f"subagent profile {src} must define description")
    when_to_use = metadata.get("whenToUse")
    body = body.rstrip() + "\n\n" + _SUBAGENT_BOUNDARY + "\n"
    # Fail closed: an empty allowlist must serialize as `tools: []`, never as
    # a bare `tools:` (YAML null) which some runtimes read as "inherit all".
    tools_block = (
        "tools:\n" + "\n".join(f"  - {t}" for t in tools_list)
        if tools_list
        else "tools: []"
    )
    disallowed_yaml = ""
    if disallowed:
        disallowed_yaml = "\ndisallowedTools:\n" + "\n".join(f"  - {t}" for t in disallowed)
    when_to_use_yaml = (
        f"whenToUse: {when_to_use}\n"
        if isinstance(when_to_use, str) and when_to_use.strip()
        else ""
    )
    return (
        "---\n"
        f"name: {name}\n"
        f"description: {description}\n"
        f"{when_to_use_yaml}"
        f"override: true\n"
        f"{tools_block}"
        f"{disallowed_yaml}\n"
        "subagents: []\n"
        "---\n\n"
        f"{body}\n"
    )


def profile_tools_for_task(name: str, server_names: list) -> list:
    """Intersect a specialized profile's declared tool ceiling with the tools
    actually reachable in the current task (identified by MCP server name).

    Specialized profiles declare exact `mcp__<server>__<tool>` entries (no
    wildcards). Entries whose server is not part of this task's MCP config are
    dropped; if nothing remains the caller must render `tools: []` (fail
    closed) — the agent then simply has no MCP tools instead of inheriting
    the main agent's set.
    """
    src = os.path.join(HARNESS_DIR, "assets", "subagents", f"{name}.md")
    with open(src, encoding="utf-8") as f:
        raw = f.read()
    parts = raw.split("---", 2)
    if len(parts) != 3:
        raise ValueError(f"invalid subagent profile frontmatter: {src}")
    metadata = yaml.safe_load(parts[1]) or {}
    declared = metadata.get("tools")
    if not isinstance(declared, list) or not all(isinstance(t, str) for t in declared):
        raise ValueError(f"invalid tools list for subagent profile: {src}")
    if any("*" in t for t in declared):
        raise ValueError(f"wildcard tool is not allowed for specialized profile: {src}")
    prefixes = tuple(f"mcp__{server}__" for server in server_names)
    return sorted({t for t in declared if t.startswith(prefixes)})


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


def run_preprocess(task_config, debug: bool) -> bool:
    """Run the task preprocess. Returns True on success, False on failure.

    On failure this writes a `.preprocess_failed` marker into the task root so the
    outer runner can classify the run as INFRA (PREPROCESS_FAILED) rather than a
    model CASE_FAILED, and the caller aborts before launching the model. Without
    this fail-fast, a broken seed (e.g. a missing fixture file) silently leaves
    the DB in a partial state and the model then burns its whole step budget on
    an unwinnable task (case-study 2026-08-12, case #19 research-lab).
    """
    init = task_config.initialization
    if not (init and init.process_command):
        return True
    cmd = init.process_command
    cmd += f" --agent_workspace {shlex.quote(str(task_config.agent_workspace))}"
    lt = task_config.launch_time or ""
    lt_clean = " ".join(lt.split()[:2])
    cmd += f" --launch_time \"{lt_clean}\""
    print_color("[preprocess] running...", "yellow")
    r = subprocess.run(cmd, shell=True, capture_output=not debug, text=True)
    if r.returncode != 0:
        msg = (r.stderr or "")[:300]
        print_color(f"[preprocess] FAILED: {msg}", "red")
        # Marker consumed by run_eval_parallel.sh to classify as PREPROCESS_FAILED.
        try:
            marker = os.path.join(task_config.task_root, ".preprocess_failed")
            with open(marker, "w", encoding="utf-8") as fh:
                fh.write(msg + "\n")
        except Exception:
            pass
        return False
    print_color("[preprocess] done.", "green")
    return True


# Provider-error signals that indicate the model backend was unavailable, not
# that the model produced a wrong answer. A run consisting ENTIRELY of these
# (with no real model progress) is classified as provider_invalid / infra.
_PROVIDER_ERR_PATTERNS = (
    "provider.connection_error",
    "provider.rate_limit",
    "provider.auth_error",
    "provider.overloaded",
    "APITimeoutError",
    "APIStatusError",
    "auth_unavailable",
    "insufficient_quota",
    "insufficient balance",
    "Request timed out",
)
# Indicators that the model actually made progress (so provider errors, if any,
# were recoverable retries rather than the whole run being invalid).
_MODEL_PROGRESS_PATTERNS = (
    '"type":"tool"',
    "tool_use",
    '"role":"assistant"',
    "function_call",
)


def _detect_provider_invalid(stream_path: str) -> str:
    """Return a non-empty reason string if the run looks provider-invalid, else ''.

    A run is provider-invalid when the raw stream contains provider-error signals
    AND no genuine model progress (no tool calls, no assistant content). This
    catches the "10 retries then die" failure mode without misclassifying a run
    that recovered after a transient 503.
    """
    try:
        text = Path(stream_path).read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return ""
    if not text:
        return ""
    err_hits = [p for p in _PROVIDER_ERR_PATTERNS if p in text]
    if not err_hits:
        return ""
    progress = any(p in text for p in _MODEL_PROGRESS_PATTERNS)
    if progress:
        return ""
    # No model progress + provider errors present -> invalid.
    return f"provider errors ({', '.join(err_hits[:3])}) with no model progress"


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
    sections = [sp.rstrip(), "", _core_responsibilities(subagents, workspace=ws)]
    if subagents:
        sections += ["", DELEGATION_RULES, "", _orchestration_rules(subagents)]
        types_sec = _subagent_types_section(subagents)
        if types_sec:
            sections += ["", types_sec]
        examples_sec = _examples_section(subagents)
        if examples_sec:
            sections += ["", examples_sec]
        coord_sec = _coordination_section()
        if coord_sec:
            sections += ["", coord_sec]
    plan_first_sec = _plan_first_section(subagents)
    if plan_first_sec:
        sections += ["", plan_first_sec]
        print_color("[plan-first] section injected (KIMI_PLAN_FIRST=1)", "cyan")
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
    disallowed_yaml = "\n".join(f"  - {t}" for t in BUILTIN_DISALLOWED_TOOLS)
    content = f"""---
name: toolathlon-main
description: Toolathlon-GYM task solver (main-agent)
override: true
tools:
{tools_yaml}
disallowedTools:
{disallowed_yaml}
{subagents_yaml}
---

{body}
"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)


_SUPPORTED_MODEL_PROVIDER_TYPES = {"openai", "anthropic"}


def resolve_model_provider_config():
    """Return the Kimi provider wire type and a correctly shaped base URL.

    OpenAI-compatible endpoints expect a ``/v1`` base URL. The Anthropic SDK
    appends ``/v1/messages`` itself, so retaining a user-supplied trailing
    ``/v1`` would incorrectly produce ``/v1/v1/messages``. The default stays
    exactly compatible with the historical OpenAI-only harness.
    """
    provider_type = os.environ.get("KIMI_PROVIDER_TYPE", "openai").strip().lower()
    if provider_type not in _SUPPORTED_MODEL_PROVIDER_TYPES:
        supported = ", ".join(sorted(_SUPPORTED_MODEL_PROVIDER_TYPES))
        raise ValueError(
            f"unsupported KIMI_PROVIDER_TYPE={provider_type!r}; expected one of: {supported}"
        )

    base_url = os.environ["MODEL_API_URL"].rstrip("/")
    if provider_type == "openai":
        if not base_url.endswith("/v1"):
            base_url += "/v1"
    elif base_url.endswith("/v1"):
        base_url = base_url[:-3]
    return provider_type, base_url


def write_kimi_home(home: str, task_config, workspace: str, marker: str,
                    max_steps: int, rewrite: dict = None):
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
    servers = _apply_path_rewrite(servers, rewrite or {})
    with open(os.path.join(home, "mcp.json"), "w", encoding="utf-8") as f:
        json.dump({"mcpServers": servers}, f, indent=2, ensure_ascii=False)

    # Sub-agent profiles. Legacy coder (retired, "three" preset only) keeps the
    # full MCP set (minus claim_done); legacy explore and plan get an explicit
    # read-only allowlist computed from the task's MCP servers so
    # write/create/delete tools are not even visible.  Specialist and
    # cross-cutting profiles (10-agent roster) keep their own declared tool
    # ceiling intersected with this task's MCP servers — no match renders
    # `tools: []` (fail closed), never the full MCP set.
    active = _active_subagents()
    server_names = list(servers.keys())
    readonly_patterns = build_readonly_tool_patterns(server_names)
    if "coder" in active:
        coder_body = render_subagent_profile(
            "coder",
            tools_list=["mcp__*"],
            disallowed=[CLAIM_TOOL, *BUILTIN_DISALLOWED_TOOLS],
        )
        with open(os.path.join(home, "agents", "coder.md"), "w", encoding="utf-8") as f:
            f.write(coder_body)
    for name in ("explore", "plan"):
        if name not in active:
            continue
        content = render_subagent_profile(
            name,
            tools_list=readonly_patterns,
            disallowed=[
                CLAIM_TOOL,
                *BUILTIN_DISALLOWED_TOOLS,
                *READONLY_TOOL_DENY_PATTERNS,
            ],
        )
        with open(os.path.join(home, "agents", f"{name}.md"), "w", encoding="utf-8") as f:
            f.write(content)
    for name in PROFILE_SUBAGENTS:
        if name not in active:
            continue
        specialist_tools = profile_tools_for_task(name, server_names)
        content = render_subagent_profile(
            name,
            tools_list=specialist_tools,
            disallowed=["Agent", "AgentSwarm", CLAIM_TOOL, *BUILTIN_DISALLOWED_TOOLS],
        )
        with open(os.path.join(home, "agents", f"{name}.md"), "w", encoding="utf-8") as f:
            f.write(content)

    provider = os.environ.get("KIMI_PROVIDER_NAME", "remote-eval")
    model_alias = os.environ.get("KIMI_MODEL_ALIAS", "eval-model")
    provider_type, base_url = resolve_model_provider_config()
    # For Anthropic-native calls, advertise thinking explicitly. This keeps
    # newer Claude aliases usable even when the installed Kimi capability
    # catalog has not caught up with their model-name prefix yet.
    model_capabilities = ""
    thinking_config = ""
    requested_effort = ""
    if provider_type == "anthropic":
        # Kimi 0.34's static capability catalog predates Sonnet 5. Declare
        # the native adaptive-thinking contract here so a requested `max`
        # survives Kimi's model-level effort normalization.
        model_capabilities = (
            'capabilities = ["thinking"]\n'
            'adaptive_thinking = true\n'
            'support_efforts = ["low", "medium", "high", "xhigh", "max"]\n'
        )
        requested_effort = os.environ.get("KIMI_MODEL_THINKING_EFFORT", "").strip()
        thinking_config = "\n[thinking]\nenabled = true\n"
        if requested_effort:
            thinking_config += f'effort = "{requested_effort}"\n'
    config = f"""default_model = "{model_alias}"

[providers."{provider}"]
type = "{provider_type}"
base_url = "{base_url}"
api_key = "{os.environ['MODEL_API_KEY']}"

[models."{model_alias}"]
provider = "{provider}"
model = "{os.environ['MODEL_NAME']}"
max_context_size = {int(os.environ.get('KIMI_MAX_CONTEXT', '262144'))}
{model_capabilities}{thinking_config}

[loop_control]
max_steps_per_turn = {max_steps}

[mcp]
startup_timeout_ms = {int(float(os.environ.get('MCP_STDIO_TIMEOUT_MIN', '90')) * 1000)}
"""
    with open(os.path.join(home, "config.toml"), "w", encoding="utf-8") as f:
        f.write(config)
    if provider_type == "anthropic" and requested_effort == "max":
        # Kimi 0.34 treats persisted max as a legacy UI value and migrates it
        # to high before its first request. Each Toolathlon task owns a fresh
        # KIMI_CODE_HOME, so mark that one-shot migration complete to retain
        # this explicitly requested, session-scoped max effort.
        with open(os.path.join(home, "migrations-effort.json"), "w", encoding="utf-8") as f:
            json.dump({"thinking-effort-max-to-high": "harness-requested"}, f)


def _check_mcp_servers_health(mcp_json_path: str) -> list:
    """Pre-flight health check for every declared MCP server.

    Returns a list of human-readable failure reasons. An empty list means every
    server passed.

    This catches the most common image-build/runtime defects WITHOUT spawning
    the servers (which kimi-code owns):
      - the launch ``command`` binary is missing/not executable (e.g. ``uv``,
        ``node`` not on PATH inside the rootfs);
      - for ``uv``-launched servers, the project's ``.venv/bin/python`` (built
        at image time) is missing — the exact failure mode that silently dropped
        the ``google_sheet`` MCP in the §C.2 rerun (canvas-course-comparison /
        arxiv-latex-review-notion-word). A missing .venv cannot be recovered at
        runtime on the no-egress compute node, so we fail-fast as infra.
      - for ``node``/``npx``-launched servers, the entry script (first arg) is
        missing.

    See dev_docs/2026-08-13-c2-tz-fix-design.md §2 (P0-2).
    """
    failures = []
    try:
        with open(mcp_json_path, encoding="utf-8") as f:
            spec = json.load(f)
    except Exception as e:
        return [f"cannot read mcp.json: {e}"]

    servers = spec.get("mcpServers", {}) or {}
    if not servers:
        return failures

    for name, cfg in servers.items():
        command = cfg.get("command", "")
        args = cfg.get("args", []) or []
        cwd = cfg.get("cwd", "") or ""

        # 1) The launch binary must resolve.
        bin_path = None
        if os.path.isabs(command):
            bin_path = command
        else:
            # Resolve against PATH (shutil.which mirrors what subprocess does).
            import shutil
            bin_path = shutil.which(command)
        if not bin_path or not os.path.isfile(bin_path):
            failures.append(
                f"MCP server '{name}': command '{command}' not found "
                f"(resolved='{bin_path}')")
            continue
        if not os.access(bin_path, os.X_OK):
            failures.append(
                f"MCP server '{name}': command '{bin_path}' is not executable")
            continue

        # 2) uv-launched Python MCP servers need their image-built .venv.
        #    `uv run` will try to (re)create it on a no-egress node and either
        #    hang past startupTimeoutMs or exit non-zero — either way kimi drops
        #    the server and every tool it provides becomes "NOT FOUND".
        #
        #    The venv lives next to the uv *project* (the package that `uv run`
        #    executes).  Toolathlon uv servers pass the project via the
        #    `--directory <dir>` flag (e.g. excel-mcp-server, mcp-google-sheets),
        #    while `cwd` is the agent workspace — so the venv must be resolved
        #    from `--directory`, NOT from `cwd`.  Checking cwd/.venv caused every
        #    excel/google_sheet task to be misclassified as infra_failed.
        if command == "uv":
            uv_project = ""
            # Resolve ${local_servers_paths} style placeholders the same way the
            # generator does: local_servers live under /opt/local_servers in the
            # image, which is also AGENT_TEMPLATE/opt/local_servers on the host.
            if "--directory" in args:
                try:
                    di = args.index("--directory")
                    if di + 1 < len(args):
                        uv_project = args[di + 1]
                except ValueError:
                    pass
            if not uv_project:
                uv_project = cwd  # fall back; some servers may set cwd to the project
            if uv_project:
                venv_python = os.path.join(uv_project, ".venv", "bin", "python")
                if not os.path.isfile(venv_python):
                    failures.append(
                        f"MCP server '{name}': uv project '{uv_project}' has no "
                        f".venv/bin/python — the image build did not sync this "
                        f"server's venv. Rebuild the image (it is now in "
                        f"UV_REQUIRED_DIRS) or fix the .venv before rerunning.")
                    continue
                # Host-path contamination guard (case-study 2026-08-14, P0-2):
                # if the runtime rootfs was hand-edited on the host, venv entry
                # scripts (dotenv/httpx/mcp/uvicorn/...) keep HOST-absolute
                # interpreter paths that cannot resolve inside the container.
                # Every executable under .venv/bin must stay container-relative.
                venv_bin = os.path.join(uv_project, ".venv", "bin")
                try:
                    bin_entries = os.listdir(venv_bin)
                except OSError:
                    bin_entries = []
                for entry in bin_entries:
                    entry_path = os.path.join(venv_bin, entry)
                    if os.path.isdir(entry_path) or os.path.islink(entry_path):
                        continue
                    try:
                        with open(entry_path, "rb") as f:
                            head = f.read(4096).decode("utf-8", "replace")
                    except OSError:
                        continue
                    # A venv entry script's shebang/exec references its own
                    # interpreter. Under enroot that must be the container path
                    # (/opt/... or /usr/...); a /lintaoLab2 or /storage path is
                    # definitive evidence of host-side contamination.
                    for bad_prefix in ("/lintaoLab2/", "/storage/lintaoLab/"):
                        if bad_prefix in head:
                            failures.append(
                                f"MCP server '{name}': venv entry '{entry}' "
                                f"references host path '{bad_prefix}...' — the "
                                f"runtime rootfs was contaminated by a host-side "
                                f"edit. Rebuild the image; do not patch the "
                                f"rootfs on the host.")
                            break

        # 3) node/npx servers: the entry script (first non-flag arg) should
        #    exist when it is a path-like value.
        if command in ("node", "npx") and args:
            entry = args[0]
            if entry and not entry.startswith("-") and entry.endswith((".js", ".mjs", ".cjs")):
                entry_full = entry if os.path.isabs(entry) else os.path.join(cwd, entry)
                if not os.path.isfile(entry_full):
                    failures.append(
                        f"MCP server '{name}': entry script '{entry_full}' "
                        f"not found")
    return failures


_LIBC = ctypes.CDLL("libc.so.6", use_errno=True)
_LIBC.mount.argtypes = (ctypes.c_char_p, ctypes.c_char_p, ctypes.c_char_p,
                        ctypes.c_ulong, ctypes.c_void_p)
_LIBC.umount2.argtypes = (ctypes.c_char_p, ctypes.c_int)
_MS_BIND = 4096
_MNT_DETACH = 2

# Directories that are Toolathlon-GYM internals, not part of the agent's
# granted boundary. They are masked (hidden) while the agent runs, then
# unmasked before the evaluator runs. Only source the agent never legitimately
# needs is masked. The MCP servers / venv / harness-tools that the agent's
# tools depend on are left in place (they are runtime deps) and guarded by the
# fallback blocker instead.
_BLACKBOX_DIRS = ("tasks", "utils", "configs", "scripts", "db",
                  "local_servers", "explorer")
# Toolathlon-GYM's own README/build scripts/source mirrors under /workspace
# that are not needed by the agent at runtime.
_BLACKBOX_FILES = ("main.py", "run_parallel.sh", "test_mcp_servers.py",
                   "docker-compose.yml", "Dockerfile", "README.md",
                   "toolathlon-gym.mdx", "pyproject.toml", "uv.lock",
                   "LICENSE", "Weekly_Meal_Plan.xlsx")
# Absolute single files to mask.
_BLACKBOX_ABS_FILES = ("/opt/provision_agent.sh",)


_MOCK_HTTP_MAX_BYTES = int(os.environ.get("KIMI_MOCK_HTTP_MAX_BYTES",
                                          str(50 * 1024 * 1024)))
_MOCK_HTTP_SKIP_FILES = {"server.log", "http.log"}


def _is_relative_to(path: str, base: str) -> bool:
    try:
        return os.path.commonpath([path, base]) == base
    except ValueError:
        return False


def _parse_http_server_cmdline(argv: list[str], cwd: str) -> tuple[int, str] | None:
    """Return (port, directory) for `python -m http.server ...` cmdlines."""
    if "http.server" not in argv:
        return None
    idx = argv.index("http.server") + 1
    port = None
    directory = None
    i = idx
    while i < len(argv):
        arg = argv[i]
        if arg in ("--directory", "-d") and i + 1 < len(argv):
            directory = argv[i + 1]
            i += 2
            continue
        if arg in ("--bind", "-b", "--protocol") and i + 1 < len(argv):
            i += 2
            continue
        if arg.startswith("-"):
            i += 1
            continue
        if port is None:
            try:
                port = int(arg)
            except ValueError:
                pass
        i += 1
    if port is None:
        port = 8000
    if not (1024 < port < 65536):
        return None
    directory = directory or cwd
    if not os.path.isabs(directory):
        directory = os.path.abspath(os.path.join(cwd, directory))
    return port, directory


def _iter_http_server_processes() -> list[dict]:
    """Scan /proc for live `python -m http.server` processes."""
    out = []
    for proc_dir in glob.glob("/proc/[0-9]*"):
        pid_s = os.path.basename(proc_dir)
        try:
            pid = int(pid_s)
            with open(os.path.join(proc_dir, "cmdline"), "rb") as f:
                raw = f.read()
            if not raw:
                continue
            argv = [a.decode(errors="replace") for a in raw.split(b"\0") if a]
            cwd = os.readlink(os.path.join(proc_dir, "cwd"))
        except (OSError, ValueError):
            continue
        parsed = _parse_http_server_cmdline(argv, cwd)
        if parsed is None:
            continue
        port, directory = parsed
        out.append({"pid": pid, "port": port, "directory": directory, "argv": argv})
    return out


def _snapshot_static_tree(root: str) -> tuple[dict[str, bytes], set[str], int]:
    """Read a static tree into memory, excluding symlinks and server logs."""
    root_real = os.path.realpath(root)
    files: dict[str, bytes] = {}
    directories: set[str] = {""}
    total = 0
    for dirpath, dirnames, filenames in os.walk(root_real, followlinks=False):
        dirnames[:] = [
            d for d in dirnames
            if not os.path.islink(os.path.join(dirpath, d))
        ]
        rel_dir = os.path.relpath(dirpath, root_real)
        rel_dir = "" if rel_dir == "." else rel_dir.replace(os.sep, "/")
        directories.add(rel_dir)
        for d in dirnames:
            directories.add(posixpath.join(rel_dir, d) if rel_dir else d)
        for name in filenames:
            src = os.path.join(dirpath, name)
            if name in _MOCK_HTTP_SKIP_FILES or os.path.islink(src):
                continue
            with open(src, "rb") as f:
                data = f.read()
            total += len(data)
            if total > _MOCK_HTTP_MAX_BYTES:
                raise ValueError(
                    f"mock HTTP tree exceeds {_MOCK_HTTP_MAX_BYTES} bytes"
                )
            rel = os.path.relpath(src, root_real).replace(os.sep, "/")
            files[rel] = data
    return files, directories, total


def _make_snapshot_handler(files: dict[str, bytes], directories: set[str]):
    class SnapshotHTTPRequestHandler(http.server.BaseHTTPRequestHandler):
        server_version = "ToolathlonMockHTTP/1.0"

        def log_message(self, fmt, *args):
            return

        def do_HEAD(self):
            self._handle(send_body=False)

        def do_GET(self):
            self._handle(send_body=True)

        def _handle(self, send_body: bool):
            parsed = urllib.parse.urlsplit(self.path)
            raw_path = urllib.parse.unquote(parsed.path)
            had_trailing = raw_path.endswith("/")
            rel = posixpath.normpath(raw_path.lstrip("/"))
            rel = "" if rel in ("", ".") else rel
            if rel.startswith("../") or rel == "..":
                self.send_error(404, "File not found")
                return

            if rel in directories:
                if not had_trailing:
                    self.send_response(301)
                    self.send_header("Location", raw_path + "/")
                    self.end_headers()
                    return
                for index_name in ("index.html", "index.htm"):
                    index_rel = posixpath.join(rel, index_name) if rel else index_name
                    if index_rel in files:
                        self._send_file(index_rel, files[index_rel], send_body)
                        return
                self._send_listing(rel, send_body)
                return

            if rel in files:
                self._send_file(rel, files[rel], send_body)
                return
            self.send_error(404, "File not found")

        def _send_file(self, rel: str, data: bytes, send_body: bool):
            ctype = mimetypes.guess_type(rel)[0] or "application/octet-stream"
            self.send_response(200)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            if send_body:
                self.wfile.write(data)

        def _send_listing(self, rel: str, send_body: bool):
            prefix = rel + "/" if rel else ""
            entries = set()
            for d in directories:
                if d and d.startswith(prefix):
                    rest = d[len(prefix):]
                    if rest and "/" not in rest:
                        entries.add(rest + "/")
            for f in files:
                if f.startswith(prefix):
                    rest = f[len(prefix):]
                    if rest and "/" not in rest:
                        entries.add(rest)
            title = f"Directory listing for /{html.escape(rel)}"
            lines = [
                "<!DOCTYPE HTML>",
                "<html><head>",
                f"<title>{title}</title>",
                "</head><body>",
                f"<h1>{title}</h1>",
                "<hr><ul>",
            ]
            for name in sorted(entries):
                q = urllib.parse.quote(name)
                lines.append(
                    f'<li><a href="{q}">{html.escape(name)}</a></li>'
                )
            lines += ["</ul><hr>", "</body></html>"]
            data = "\n".join(lines).encode()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            if send_body:
                self.wfile.write(data)

    return SnapshotHTTPRequestHandler


def _terminate_pid(pid: int):
    for sig in (signal.SIGTERM, signal.SIGKILL):
        try:
            os.kill(pid, sig)
        except ProcessLookupError:
            return
        except PermissionError:
            return
        deadline = time.time() + (1.5 if sig == signal.SIGTERM else 0.5)
        while time.time() < deadline:
            try:
                os.kill(pid, 0)
            except ProcessLookupError:
                return
            time.sleep(0.05)


def _start_snapshot_http_server(port: int, files: dict[str, bytes],
                                directories: set[str]):
    class ReusableThreadingHTTPServer(http.server.ThreadingHTTPServer):
        allow_reuse_address = True

    handler = _make_snapshot_handler(files, directories)
    server = ReusableThreadingHTTPServer(("127.0.0.1", port), handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    return server, thread


def _wait_http_ready(port: int, timeout_s: float = 3.0) -> bool:
    deadline = time.time() + timeout_s
    url = f"http://127.0.0.1:{port}/"
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=0.5) as resp:
                return 200 <= resp.status < 400
        except Exception:
            time.sleep(0.1)
    return False


def _relocate_mock_servers(task_config) -> list:
    """Keep task mock HTTP data reachable after `tasks/` is masked.

    Some preprocess scripts start `python -m http.server --directory` inside the
    task tree. A later bind-mount hides `tasks/`, turning those servers into
    404 machines. We snapshot only the active server directory for this task and
    serve it back from memory on the same localhost port. No task source,
    evaluator, or ground truth path is exposed, and no new /tmp copy is created.
    """
    if os.environ.get("KIMI_DISABLE_BOUNDARY") == "1":
        return []
    task_root = os.path.realpath(os.path.join(
        os.getcwd(), "tasks", "finalpool", task_config.task_dir
    ))
    task_name = task_config.task_dir
    relocated = []
    seen_ports = set()
    for proc in _iter_http_server_processes():
        port = proc["port"]
        serve_dir = os.path.realpath(proc["directory"])
        if port in seen_ports or not os.path.isdir(serve_dir):
            continue
        if not _is_relative_to(serve_dir, task_root):
            continue
        rel = os.path.relpath(serve_dir, task_root).replace(os.sep, "/")
        allowed = (
            rel == "tmp" or rel.startswith("tmp/") or
            rel == "files/mock_pages" or rel.startswith("files/mock_pages/")
        )
        if not allowed:
            print_color(
                f"[mock-http] skip unsafe serve dir for {task_name}: {rel}",
                "yellow",
            )
            continue
        try:
            files, directories, total = _snapshot_static_tree(serve_dir)
        except Exception as exc:
            print_color(
                f"[mock-http] skip localhost:{port}: snapshot failed: {exc}",
                "yellow",
            )
            continue
        _terminate_pid(proc["pid"])
        try:
            server, thread = _start_snapshot_http_server(port, files, directories)
        except OSError as exc:
            print_color(
                f"[mock-http] failed to restart localhost:{port}: {exc}",
                "yellow",
            )
            continue
        if not _wait_http_ready(port):
            print_color(f"[mock-http] localhost:{port} did not become ready",
                        "yellow")
            server.shutdown()
            server.server_close()
            continue
        relocated.append((server, thread, port))
        seen_ports.add(port)
        print_color(
            f"[mock-http] snapshotted {len(files)} files ({total} bytes) "
            f"for {task_name}; serving localhost:{port} from memory",
            "cyan",
        )
    return relocated


def _shutdown_mock_servers(servers: list):
    for server, thread, port in servers:
        try:
            server.shutdown()
            server.server_close()
            thread.join(timeout=1)
        except Exception as exc:
            print_color(f"[mock-http] warn: shutdown localhost:{port}: {exc}",
                        "yellow")


def _mount(src: str, tgt: str, fstype=None, flags=0) -> bool:
    r = _LIBC.mount(src.encode(), tgt.encode(),
                    fstype.encode() if fstype else None, flags, None)
    return r == 0


def _mask_blackbox() -> list:
    """Hide Toolathlon-GYM internals from the agent by bind-mounting an empty
    directory over each blackbox path. Returns the masked mount points so the
    caller can unmask them before the evaluator runs.

    The agent runs with a cleared capability bounding set, so it cannot umount
    these masks — the blackbox paths simply do not exist for it.
    """
    if os.environ.get("KIMI_DISABLE_BOUNDARY") == "1":
        return []
    root = os.getcwd()
    empty = "/tmp/.blackbox_empty"
    empty_file = "/tmp/.blackbox_empty_file"
    os.makedirs(empty, exist_ok=True)
    if not os.path.exists(empty_file):
        open(empty_file, "w").close()
    mounted = []
    for rel in _BLACKBOX_DIRS:
        tgt = os.path.join(root, rel)
        if os.path.isdir(tgt) and _mount(empty, tgt, None, _MS_BIND):
            mounted.append(tgt)
        elif os.path.isdir(tgt):
            print_color(f"[boundary] warn: failed to mask {tgt}", "yellow")
    for rel in _BLACKBOX_FILES:
        tgt = os.path.join(root, rel)
        if os.path.isfile(tgt) and _mount(empty_file, tgt, None, _MS_BIND):
            mounted.append(tgt)
    for tgt in _BLACKBOX_ABS_FILES:
        if os.path.isfile(tgt) and _mount(empty_file, tgt, None, _MS_BIND):
            mounted.append(tgt)
    print_color(f"[boundary] masked {len(mounted)} blackbox paths", "cyan")
    return mounted


def _unmask_blackbox(mounted: list):
    for tgt in mounted:
        _LIBC.umount2(tgt.encode(), _MNT_DETACH)


def launch_kimi(task_str: str, agentfile: str, home: str, stream_path: str,
                workspace: str, marker: str, timeout_s: int, debug: bool,
                rewrite: dict = None):
    env = dict(os.environ)
    env["KIMI_CODE_HOME"] = home
    env["KIMI_CODE_BUILTIN_PRODUCT_SKILLS"] = "0"
    env.pop("http_proxy", None)
    env.pop("https_proxy", None)
    env.pop("HTTP_PROXY", None)
    env.pop("HTTPS_PROXY", None)

    # Credential isolation (audit §A.3 / security-boundary): the evaluated
    # agent must NOT be able to read backing-DB credentials from its own
    # process environment. The MCP servers receive PG credentials through the
    # per-server `env` block written into mcp.json (see mcp_json_gen.py), so
    # the agent's global environment does not need them. Strip every libpq /
    # Toolathlon spelling of PGHOST/PGPORT/PGUSER/PGPASSWORD/PGDATABASE so
    # `os.environ['PGPASSWORD']` and `env | grep PG` both come back empty.
    # MODEL_API_KEY is kept here because the kimi-code CLI needs it to call
    # the model; python_execute strips it again at the tool boundary (see
    # local_tools_server.py) so agent-authored code still cannot read it.
    if os.environ.get("KIMI_DISABLE_BOUNDARY") != "1":
        for _k in list(env.keys()):
            if _k.upper().startswith("PG") and _k.upper() in (
                "PGHOST", "PG_HOST", "PGPORT", "PG_PORT",
                "PGUSER", "PG_USER", "PGPASSWORD", "PG_PASSWORD",
                "PGDATABASE", "PG_DATABASE", "PGSERVICE", "PG_SERVICE",
                "PGSERVICEFILE", "PG_SERVICE_FILE",
                "PGSSLMODE", "PG_SSLMODE",
            ):
                env.pop(_k, None)
        print_color("[boundary] stripped PG credentials from agent environment",
                    "cyan")

    # Rewrite env paths that point into masked dirs (e.g. /opt/venv on PATH)
    # so the agent's tools resolve to the staged copies.
    rewrite = rewrite or {}
    if rewrite:
        def rw(v):
            for src, dst in rewrite.items():
                if v == src:
                    return dst
                if v.startswith(src + os.sep):
                    return dst + v[len(src):]
            return v
        for key in ("PATH", "VIRTUAL_ENV", "PYTHON_BIN", "LOCAL_SERVERS_PATH"):
            if key in env:
                sep = ":" if key == "PATH" else None
                parts = env[key].split(":") if key == "PATH" else [env[key]]
                parts = [rw(p) for p in parts]
                env[key] = ":".join(parts) if key == "PATH" else parts[0]

    kimi_cmd = ["kimi", "-p", task_str,
                "--agent-file", agentfile,
                "--output-format", "stream-json"]
    # Boundary hardening: the agent subtree must not be able to undo the
    # blackbox bind-mount masks. enroot runs everything as uid 0 (single-UID
    # user namespace), so file permissions cannot isolate the agent — but
    # dropping the capability bounding set can. Strip every capability so the
    # agent cannot mount/umount (CAP_SYS_ADMIN) and re-expose masked source.
    cmd = kimi_cmd
    if os.environ.get("KIMI_DISABLE_BOUNDARY") != "1":
        cmd = (["setpriv", "--bounding-set=-all", "--inh-caps=-all",
                "--ambient-caps=-all", "--no-new-privs"] + kimi_cmd)
        print_color("[boundary] agent launched with capability bounding set cleared",
                    "cyan")
    print_color(f"[kimi] launching: {' '.join(kimi_cmd[:3])} ... (home={home})", "cyan")

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


def _stage_agent_runtime(task_root: str) -> dict:
    """No staging copies are needed: the MCP servers, venv, and harness tools
    the agent depends on are left in place (not masked). Only Toolathlon task
    source / evaluation / gym internals — which the agent never runs — are
    masked. Returns an empty path-rewrite map (kept for interface stability).
    """
    return {}


def _apply_path_rewrite(servers: dict, rewrite: dict):
    if not rewrite:
        return servers
    def rw(v):
        if not isinstance(v, str):
            return v
        for src, dst in rewrite.items():
            if v == src:
                return dst
            if v.startswith(src + os.sep):
                return dst + v[len(src):]
        return v
    for srv in servers.values():
        srv["command"] = rw(srv.get("command", ""))
        srv["args"] = [rw(a) for a in srv.get("args", [])]
        srv["cwd"] = rw(srv.get("cwd", ""))
        env = srv.get("env", {})
        srv["env"] = {k: rw(v) for k, v in env.items()}
    return servers


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
    # Dump/workspace paths are interpolated into shell preprocess commands.
    # MODEL_NAME may contain a thinking suffix such as gpt-5.6-sol(xhigh);
    # keep that string for config.toml, but strip shell-special chars here.
    dump_model = re.sub(r"[^\w.+-]+", "-", model_name).strip("-") or "unknown-model"

    task_config = TaskConfig.build(
        args.task_dir,
        agent_short_name=f"kimi-code/{dump_model}",
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
    if not run_preprocess(task_config, args.debug):
        # Fail-fast: do NOT launch the model on a task whose DB seed failed.
        # The .preprocess_failed marker lets the outer runner classify this as
        # PREPROCESS_FAILED (infra), not a model CASE_FAILED.
        print_color("[kimi] aborting: preprocess failed; not launching model.", "red")
        return 3
    mock_servers = _relocate_mock_servers(task_config)

    # Stage agent-runtime deps (MCP servers, harness tools, venv) into an
    # agent-visible dir so the originals can be masked without breaking tools.
    rewrite = _stage_agent_runtime(task_root)

    kimi_home = os.path.join(task_root, ".kimi_home")
    write_kimi_home(kimi_home, task_config, workspace, marker, max_steps,
                    rewrite=rewrite)

    # Pre-flight MCP health check (P0-2): verify every declared MCP server's
    # launch binary / venv / entry script exists BEFORE launching the model. A
    # missing .venv (the google_sheet failure mode) cannot be recovered at
    # runtime on a no-egress compute node; fail-fast here as infra so the run
    # is classified as INFRA_FAILED and auto-rerun, not scored as a model fail.
    mcp_json_path = os.path.join(kimi_home, "mcp.json")
    mcp_failures = _check_mcp_servers_health(mcp_json_path)
    if mcp_failures and os.environ.get("KIMI_DISABLE_BOUNDARY") != "1":
        msg = "MCP server health check failed:\n  - " + "\n  - ".join(mcp_failures)
        print_color(f"[infra] {msg}", "red")
        try:
            with open(os.path.join(task_root, ".infra_failed"), "w",
                      encoding="utf-8") as fh:
                fh.write(msg + "\n")
        except Exception:
            pass
        return 3

    agentfile = os.path.join(task_root, "agent_main.md")
    write_agentfile(agentfile, render_system_prompt(task_config))

    stream_path = os.path.join(task_root, "raw_stream.jsonl")
    timeout_s = int(os.environ.get("KIMI_TASK_TIMEOUT_S", "7200"))

    # Mask Toolathlon internals from the agent, run kimi, then unmask so the
    # evaluator (which reads tasks/, utils/) can run. finally guarantees unmask.
    masked = _mask_blackbox()
    try:
        rc, done = launch_kimi(
            task_str=task_config.task_str,
            agentfile=agentfile,
            home=kimi_home,
            stream_path=stream_path,
            workspace=workspace,
            marker=marker,
            timeout_s=timeout_s,
            debug=args.debug,
            rewrite=rewrite,
        )
    finally:
        _unmask_blackbox(masked)

    status = "success" if done else "failed"
    print_color(f"[kimi] exited rc={rc} claim_done={done} -> status={status}",
                "green" if done else "red")

    write_traj_logs(task_config, status, start_time, stream_path)

    # Provider-invalid detection (case-study 2026-08-12, case #1 sf-support): a run
    # that never produced a usable assistant turn — only provider connection /
    # auth / balance / overload errors — is an infra failure, not a model failure.
    # We scan the raw stream for provider-error signals and for any real model
    # progress (tool calls or assistant text content). If there is no real model
    # progress AND provider errors are present, write a marker so the outer runner
    # classifies the run as provider_invalid instead of case_failed, and the task
    # gets an automatic rerun rather than being scored as a model ability deficit.
    try:
        pv = _detect_provider_invalid(stream_path)
        if pv:
            print_color(f"[provider] invalid run detected: {pv}", "magenta")
            try:
                with open(os.path.join(task_root, ".provider_invalid"), "w",
                          encoding="utf-8") as fh:
                    fh.write(pv + "\n")
            except Exception:
                pass
            # Skip evaluation: there is nothing to evaluate, and the run is invalid.
            return 4

        print_color("\n====== Evaluating ======", "yellow")
        from utils.evaluation.evaluator import TaskEvaluator
        eval_res = await TaskEvaluator.evaluate_from_log_file(task_config.log_file)
        print(f"Pass:    {eval_res.get('pass', False)}")
        print(f"Details: {eval_res.get('details', 'N/A')}")
        # Only a completed lifecycle may return 0 to the outer runner. A
        # no-claim artifact-only PASS is valuable audit signal, but the worker
        # status must remain case_failed so downstream tables can distinguish
        # "answer quality passed" from "agent followed completion protocol".
        return 0 if (done and eval_res.get("pass", False)) else 1
    finally:
        # Keep mock services alive through evaluation so no-claim runs with
        # deliverables can still be judged against local service side effects.
        _shutdown_mock_servers(mock_servers)


def main():
    raise SystemExit(asyncio.run(amain()))


if __name__ == "__main__":
    main()
