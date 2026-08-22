#!/usr/bin/env python3
"""Generate an audit HTML page per evaluated case, placed inside the case dump dir.

Anthropic-inspired light UI. Rich timeline: reasoning, structured tool calls,
tool results, sub-agent delegation. Critical Steps per formula 2, where a step
that fires multiple Agent calls concurrently counts as ONE parallel phase.
"""

from __future__ import annotations

import html
import json
import math
import re
import sys
from collections import defaultdict
from pathlib import Path

DELEGATION_TOOLS = {"Agent", "AgentSwarm"}
# Slot dir names are "<RUN_ID>_slot<N>" where RUN_ID may be a bare timestamp
# (20260817-002742) or carry a prefix (rerun-fix6-20260817-102737,
# subagent-20260817-120650). Accept an optional alphanumeric prefix before
# the timestamp so prefixed runs still match.
RUN_DIR_RE = re.compile(r"^(?:[A-Za-z0-9][A-Za-z0-9_.-]*[-_])?(?:\d{8}-\d{6}|\d{8})(?:[-_][A-Za-z0-9_.-]+)?_slot\d+$")
USAGE_FIELDS = ("inputOther", "inputCacheRead", "inputCacheCreation", "output")

# Model/provider/relay failures only.  Do not include generic "API" or
# localhost connection failures here: many Toolathlon tasks intentionally use
# business/data APIs, and failed local probing is task behavior rather than a
# model-provider outage.
API_ERROR_RE = re.compile(
    r"("
    r"provider_invalid|No providers available|MODEL_API_URL|RELAY_API_KEY|"
    r"API(?:Connection|Timeout)?Error|APITimeoutError|APIConnectionError|"
    r"AuthenticationError|PermissionDeniedError|RateLimitError|"
    r"429 Too Many Requests|401 Unauthorized|403 Forbidden|"
    r"502 Bad Gateway|503 Service Unavailable|504 Gateway Timeout|"
    r"invalid api key|api key invalid|quota exceeded|"
    r"provider error|upstream error|relay error|model api"
    r")",
    re.I,
)


def load_tokenizer():
    try:
        import tiktoken  # type: ignore

        return tiktoken.get_encoding("cl100k_base")
    except Exception:
        return None


TOKENIZER = load_tokenizer()


def estimate_tokens(text: str) -> int:
    if not text:
        return 0
    if TOKENIZER is not None:
        return len(TOKENIZER.encode(text))
    return max(1, math.ceil(len(text) / 4))


def int_or_none(value) -> int | None:
    if value is None or value == "":
        return None
    try:
        return int(float(value))
    except Exception:
        return None


def usage_totals(usage: dict | None) -> dict[str, int]:
    usage = usage or {}
    return {k: int_or_none(usage.get(k)) or 0 for k in USAGE_FIELDS}


def add_usage(dst: dict[str, int], src: dict[str, int]) -> None:
    for k in USAGE_FIELDS:
        dst[k] = dst.get(k, 0) + src.get(k, 0)


def step_token_total(step: dict) -> int:
    u = step.get("usage_totals") or {}
    return sum(int(u.get(k, 0)) for k in USAGE_FIELDS)


def wire_total_tokens(wire: dict) -> int:
    u = wire.get("usage_totals") or {}
    return sum(int(u.get(k, 0)) for k in USAGE_FIELDS)


def prompt_tokens_from_usage(usage: dict) -> int:
    return int(usage.get("inputOther", 0)) + int(usage.get("inputCacheRead", 0)) + int(usage.get("inputCacheCreation", 0))


def is_api_error_step(step: dict) -> bool:
    chunks: list[str] = [str(step.get("content") or "")]
    for call in step.get("calls", []):
        chunks.append(str(call.get("name") or ""))
        chunks.append(str(call.get("args_text") or ""))
        try:
            chunks.append(json.dumps(call.get("args") or {}, ensure_ascii=False))
        except Exception:
            chunks.append(str(call.get("args") or ""))
        chunks.append(str(call.get("result") or ""))
    return bool(API_ERROR_RE.search("\n".join(chunks)))


def load_json(p: Path):
    try:
        return json.loads(p.read_text(errors="replace"))
    except Exception:
        return None


def find_case_dirs(dump_root: Path) -> list[Path]:
    dirs = []
    for task_dir in sorted(dump_root.iterdir() if dump_root.is_dir() else []):
        if not task_dir.is_dir():
            continue
        for run_dir in sorted(task_dir.iterdir()):
            if run_dir.is_dir() and RUN_DIR_RE.match(run_dir.name):
                dirs.append(run_dir)
    return dirs


def find_inner_dir(case_dir: Path) -> Path | None:
    for model_dir in case_dir.iterdir():
        if model_dir.is_dir() and model_dir.name.startswith("kimi-code"):
            for inner in model_dir.iterdir():
                if inner.is_dir() and inner.name.startswith("SingleUserTurn"):
                    return inner
    return None


def parse_run_log(case_dir: Path) -> dict:
    p = case_dir / "run.log"
    out: dict = {}
    if not p.exists():
        return out
    text = p.read_text(errors="replace")
    m = re.search(r"\[(\d\d:\d\d:\d\d)\] DONE\s+\S+\s+->\s+(\S+)\s+\(exit=(\d+),\s*(\d+)s", text)
    if m:
        out["done_time"], out["worker_status"], out["duration_s"] = m.group(1), m.group(2), int(m.group(4))
    m = re.search(r"exited rc=(\d+) claim_done=(\w+)", text)
    if m:
        out["kimi_rc"], out["claim_done"] = int(m.group(1)), m.group(2) == "True"
    m = re.search(r"^Pass:\s+(\S+)", text, re.M)
    if m:
        out["eval_pass_text"] = m.group(1)
    m = re.search(r"Overall:\s*(\d+)/(\d+)\s*\(([\d.]+)%\)", text)
    if m:
        out["ckpt_pass"], out["ckpt_total"], out["ckpt_pct"] = int(m.group(1)), int(m.group(2)), float(m.group(3))
    return out


def parse_eval_res(inner: Path | None) -> dict:
    out: dict = {}
    p = (inner / "eval_res.json") if inner else None
    if p and p.exists():
        d = load_json(p) or {}
        out["pass"] = d.get("pass")
        failure = d.get("failure") or ""
        out["failure_text"] = failure
        m = re.search(r"ERRORS:\s*(\d+)", failure)
        if m:
            out["eval_errors"] = int(m.group(1))
        m = re.search(r"RESULT:\s*(\S+)", failure)
        if m:
            out["eval_result"] = m.group(1)
    return out


def parse_artifact_eval_res(inner: Path | None) -> dict:
    """Load supplemental artifact-only evaluation for no-claim runs."""
    out: dict = {}
    p = (inner / "artifact_eval_res.json") if inner else None
    if p and p.exists():
        d = load_json(p) or {}
        out["pass"] = d.get("pass")
        out["returncode"] = d.get("returncode")
        out["details"] = d.get("details") or ""
        out["stdout"] = d.get("stdout") or ""
        out["stderr"] = d.get("stderr") or ""
        out["command"] = d.get("command") or ""
        failure = d.get("failure") or d.get("stdout") or d.get("details") or ""
        out["failure_text"] = failure
        m = re.search(r"ERRORS:\s*(\d+)", failure)
        if m:
            out["eval_errors"] = int(m.group(1))
        m = re.search(r"RESULT:\s*(\S+)", failure)
        if m:
            out["eval_result"] = m.group(1)
    return out


def count_checkpoints(failure_text: str):
    """Return (pass_count, fail_count) parsed from [PASS]/[FAIL] markers."""
    if not failure_text:
        return None
    p = len(re.findall(r"^\s*\[PASS\]", failure_text, re.M))
    f = len(re.findall(r"^\s*\[FAIL\]", failure_text, re.M))
    return (p, f) if p + f > 0 else None


def tier_from_mcp_count(n_mcp: int) -> str:
    """Map MCP-server count to the dataset tier label used in the README
    (4=T1, 5=T2, 6=T3, 7-8=T4)."""
    return {4: "T1", 5: "T2", 6: "T3", 7: "T4", 8: "T4"}.get(n_mcp, "T?")


# ---------------------------------------------------------------------------
# Wire parsing
# ---------------------------------------------------------------------------


def parse_wire(wire_path: Path) -> dict:
    steps: dict[int, dict] = {}
    cur_step: dict | None = None
    profile_name = None
    sub_prompt = None
    result_by_id: dict[str, dict] = {}
    fallback_usage_by_step: dict[int, list[dict[str, int]]] = defaultdict(list)

    if not wire_path.exists():
        return {"steps": [], "all_steps": [], "n_steps": 0, "raw_n_steps": 0,
                "api_error_steps": 0, "profile": None, "prompt": None,
                "usage_totals": {k: 0 for k in USAGE_FIELDS},
                "excluded_usage_totals": {k: 0 for k in USAGE_FIELDS},
                "observation_tokens_est": 0, "excluded_observation_tokens_est": 0}

    for line in wire_path.read_text(errors="replace").splitlines():
        try:
            j = json.loads(line)
        except Exception:
            continue
        t = j.get("type")

        if t == "profile.bind":
            profile_name = j.get("profileName") or j.get("profile")
        elif t == "turn.prompt":
            inp = j.get("input") or []
            if isinstance(inp, list):
                sub_prompt = " ".join(x.get("text", "") for x in inp if isinstance(x, dict))[:4000]
            else:
                sub_prompt = str(j.get("prompt") or "")[:4000]
        elif t == "usage.record":
            if cur_step is not None:
                sn = cur_step.get("step")
                if isinstance(sn, int):
                    fallback_usage_by_step[sn].append(usage_totals(j.get("usage")))
        elif t == "context.append_loop_event":
            ev = j.get("event", {})
            et = ev.get("type")
            sn = ev.get("step")
            if et == "step.begin":
                cur_step = {"step": sn, "calls": [], "content": "",
                            "usage_totals": {k: 0 for k in USAGE_FIELDS},
                            "observation_tokens_est": 0, "has_step_end_usage": False,
                            "api_error": False}
                steps[sn] = cur_step
            elif et == "tool.call":
                raw_args = ev.get("args")
                if isinstance(raw_args, str):
                    try:
                        parsed = json.loads(raw_args)
                        call = {"name": ev.get("name"), "args": parsed if isinstance(parsed, dict) else {},
                                "args_text": raw_args,
                                "id": ev.get("toolCallId"), "result": None, "is_error": False}
                    except Exception:
                        call = {"name": ev.get("name"), "args": {},
                                "args_text": raw_args,
                                "id": ev.get("toolCallId"), "result": None, "is_error": False}
                else:
                    call = {"name": ev.get("name"), "args": raw_args or {},
                            "id": ev.get("toolCallId"), "result": None, "is_error": False}
                if cur_step is not None:
                    cur_step["calls"].append(call)
            elif et == "tool.result":
                res = ev.get("result") or {}
                output = res.get("output") if isinstance(res, dict) else str(res)
                is_error = bool(res.get("is_error")) if isinstance(res, dict) else False
                result_by_id[ev.get("toolCallId")] = {"output": str(output), "is_error": is_error}
                if cur_step is not None:
                    cur_step["observation_tokens_est"] += estimate_tokens(str(output or ""))
            elif et == "content.part":
                part = ev.get("part") or {}
                txt = part.get("text") or ""
                if cur_step is not None and txt:
                    cur_step["content"] += txt + "\n"
            elif et == "step.end":
                if isinstance(sn, int):
                    st = steps.setdefault(sn, {"step": sn, "calls": [], "content": "",
                                               "usage_totals": {k: 0 for k in USAGE_FIELDS},
                                               "observation_tokens_est": 0, "has_step_end_usage": False,
                                               "api_error": False})
                    st["usage_totals"] = usage_totals(ev.get("usage"))
                    st["has_step_end_usage"] = True

    step_list = [steps[k] for k in sorted(steps)]
    for st in step_list:
        for call in st["calls"]:
            r = result_by_id.get(call["id"])
            if r:
                call["result"], call["is_error"] = r["output"], r["is_error"]
        if not st.get("has_step_end_usage"):
            total = {k: 0 for k in USAGE_FIELDS}
            sn = st.get("step")
            for u in fallback_usage_by_step.get(sn, []):
                add_usage(total, u)
            st["usage_totals"] = total
        st["api_error"] = is_api_error_step(st)

    included = [st for st in step_list if not st.get("api_error")]
    totals = {k: 0 for k in USAGE_FIELDS}
    excluded_totals = {k: 0 for k in USAGE_FIELDS}
    observation = 0
    excluded_observation = 0
    for st in step_list:
        if st.get("api_error"):
            add_usage(excluded_totals, st.get("usage_totals") or {})
            excluded_observation += int(st.get("observation_tokens_est") or 0)
        else:
            add_usage(totals, st.get("usage_totals") or {})
            observation += int(st.get("observation_tokens_est") or 0)

    return {"steps": included, "all_steps": step_list, "n_steps": len(included),
            "raw_n_steps": len(step_list), "api_error_steps": len(step_list) - len(included),
            "profile": profile_name, "prompt": sub_prompt,
            "usage_totals": totals, "excluded_usage_totals": excluded_totals,
            "observation_tokens_est": observation,
            "excluded_observation_tokens_est": excluded_observation}


def compute_critical_steps(main_wire: dict, sub_wires: list[dict]) -> dict:
    """Compute critical-path steps after excluding model/API-error steps.

    API-error steps remain visible in the audit timeline, but are removed from
    step counts and token totals for both main and sub-agents.
    """
    phases: list[dict] = []
    sub_cursor = 0
    phase_idx = 0
    critical_tokens_api = 0
    critical_observation_tokens = 0

    for st in main_wire["steps"]:
        deleg = [c for c in st["calls"] if c["name"] in DELEGATION_TOOLS]
        main_tokens = step_token_total(st)
        main_obs = int(st.get("observation_tokens_est") or 0)
        if not deleg:
            critical_tokens_api += main_tokens
            critical_observation_tokens += main_obs
            continue
        n_sub = 0
        tools_used = set()
        for call in deleg:
            tools_used.add(call["name"])
            if call["name"] == "AgentSwarm":
                items = call["args"].get("items") or call["args"].get("tasks") or []
                n_sub += len(items) if isinstance(items, list) and items else 1
            else:
                n_sub += 1
        mode = "parallel" if n_sub > 1 else "sequential"
        taken = sub_wires[sub_cursor:sub_cursor + n_sub]
        sub_cursor += len(taken)
        sub_step_counts = [w["n_steps"] for w in taken]
        sub_token_counts = [wire_total_tokens(w) for w in taken]
        sub_obs_counts = [int(w.get("observation_tokens_est") or 0) for w in taken]
        slowest = max(sub_step_counts) if sub_step_counts else 0
        slowest_tokens = max(sub_token_counts) if sub_token_counts else 0
        slowest_obs = max(sub_obs_counts) if sub_obs_counts else 0
        phase_idx += 1
        phases.append({"phase": phase_idx, "mode": mode,
                       "deleg_tool": "/".join(sorted(tools_used)),
                       "n_sub": len(taken), "s_main": 1,
                       "sub_steps": sub_step_counts, "slowest": slowest,
                       "cost": 1 + slowest,
                       "main_tokens_api": main_tokens,
                       "sub_tokens_api": sub_token_counts,
                       "slowest_tokens_api": slowest_tokens,
                       "token_cost_api": main_tokens + slowest_tokens})
        critical_tokens_api += main_tokens + slowest_tokens
        critical_observation_tokens += main_obs + slowest_obs

    total_main = main_wire["n_steps"]
    n_deleg_phases = len(phases)
    solo = max(0, total_main - n_deleg_phases)
    if solo:
        solo_steps = [st for st in main_wire["steps"] if not any(c["name"] in DELEGATION_TOOLS for c in st["calls"])]
        solo_tokens = sum(step_token_total(st) for st in solo_steps)
        phases.insert(0, {"phase": 0, "mode": "solo", "deleg_tool": None,
                          "n_sub": 0, "s_main": solo, "sub_steps": [],
                          "slowest": 0, "cost": solo,
                          "main_tokens_api": solo_tokens, "sub_tokens_api": [],
                          "slowest_tokens_api": 0, "token_cost_api": solo_tokens})
    critical = sum(p["cost"] for p in phases)
    serial = total_main + sum(w["n_steps"] for w in sub_wires)
    n_parallel = sum(p["n_sub"] for p in phases if p["mode"] == "parallel")
    n_sequential = sum(p["n_sub"] for p in phases if p["mode"] == "sequential")
    total_usage = {k: 0 for k in USAGE_FIELDS}
    excluded_usage = {k: 0 for k in USAGE_FIELDS}
    observation = int(main_wire.get("observation_tokens_est") or 0)
    excluded_observation = int(main_wire.get("excluded_observation_tokens_est") or 0)
    add_usage(total_usage, main_wire.get("usage_totals") or {})
    add_usage(excluded_usage, main_wire.get("excluded_usage_totals") or {})
    for w in sub_wires:
        add_usage(total_usage, w.get("usage_totals") or {})
        add_usage(excluded_usage, w.get("excluded_usage_totals") or {})
        observation += int(w.get("observation_tokens_est") or 0)
        excluded_observation += int(w.get("excluded_observation_tokens_est") or 0)
    prompt_tokens = prompt_tokens_from_usage(total_usage)
    assistant_tokens = int(total_usage.get("output", 0))
    excluded_prompt_tokens = prompt_tokens_from_usage(excluded_usage)
    excluded_assistant_tokens = int(excluded_usage.get("output", 0))
    serial_tokens_api = prompt_tokens + assistant_tokens
    return {"phases": phases, "critical_steps": critical, "serial_steps": serial,
            "n_parallel_subs": n_parallel, "n_sequential_subs": n_sequential,
            "main_steps": total_main, "main_steps_raw": main_wire.get("raw_n_steps", total_main),
            "main_api_error_steps": main_wire.get("api_error_steps", 0),
            "sub_api_error_steps": sum(w.get("api_error_steps", 0) for w in sub_wires),
            "api_error_steps": main_wire.get("api_error_steps", 0) + sum(w.get("api_error_steps", 0) for w in sub_wires),
            "prompt_tokens_api": prompt_tokens,
            "prompt_uncached_tokens_api": int(total_usage.get("inputOther", 0)),
            "prompt_cache_read_tokens_api": int(total_usage.get("inputCacheRead", 0)),
            "prompt_cache_creation_tokens_api": int(total_usage.get("inputCacheCreation", 0)),
            "assistant_tokens_api": assistant_tokens,
            "total_tokens_api": serial_tokens_api,
            "observation_tokens_est": observation,
            "critical_tokens_api": critical_tokens_api,
            "serial_tokens_api": serial_tokens_api,
            "critical_token_saving_api": max(0, serial_tokens_api - critical_tokens_api),
            "critical_observation_tokens_est": critical_observation_tokens,
            "serial_observation_tokens_est": observation,
            "critical_all_tokens_est": critical_tokens_api + critical_observation_tokens,
            "serial_all_tokens_est": serial_tokens_api + observation,
            "excluded_prompt_tokens_api": excluded_prompt_tokens,
            "excluded_assistant_tokens_api": excluded_assistant_tokens,
            "excluded_total_tokens_api": excluded_prompt_tokens + excluded_assistant_tokens,
            "excluded_observation_tokens_est": excluded_observation}


def compute_plan_first_metrics(main_wire: dict) -> dict:
    """Compliance metrics for the plan-first arm (monitoring only, no gating).

    plan_first_ok   : the first Agent/AgentSwarm dispatch on the main wire was
                      an Agent call with subagent_type == 'plan'.
    pre_plan_actions: number of main-agent tool calls before that plan dispatch
                      (0 = planned immediately; large = worked solo first).
    n_plan_dispatches: how many plan-sub-agent dispatches happened in total.
    """
    out = {"plan_first_ok": None, "pre_plan_actions": None,
           "n_plan_dispatches": 0, "first_dispatch": None}
    actions_before = 0
    for st in main_wire.get("steps", []):
        for call in st.get("calls", []):
            name = call.get("name")
            if name == "Agent":
                sub_type = (call.get("args") or {}).get("subagent_type")
                if sub_type == "plan":
                    out["n_plan_dispatches"] += 1
                    if out["plan_first_ok"] is None:
                        out["plan_first_ok"] = True
                        out["pre_plan_actions"] = actions_before
                        out["first_dispatch"] = sub_type
                    actions_before += 1
                    continue
                if out["plan_first_ok"] is None:
                    out["plan_first_ok"] = False
                    out["pre_plan_actions"] = actions_before
                    out["first_dispatch"] = sub_type
            elif name == "AgentSwarm":
                items = (call.get("args") or {}).get("items") or []
                if isinstance(items, list):
                    for it in items:
                        if isinstance(it, dict) and it.get("subagent_type") == "plan":
                            out["n_plan_dispatches"] += 1
                if out["plan_first_ok"] is None:
                    out["plan_first_ok"] = False
                    out["pre_plan_actions"] = actions_before
                    out["first_dispatch"] = "AgentSwarm"
            else:
                actions_before += 1
    if out["plan_first_ok"] is None and main_wire.get("n_steps"):
        # delegated but never dispatched a plan sub-agent
        out["plan_first_ok"] = False
    return out


def parse_tools_inventory(inner: Path) -> dict:
    out: dict = {"mcp_servers": [], "subagents": []}
    mcp = load_json(inner / ".kimi_home" / "mcp.json")
    if mcp and "mcpServers" in mcp:
        out["mcp_servers"] = list(mcp["mcpServers"].keys())
    agents_dir = inner / ".kimi_home" / "agents"
    if agents_dir.is_dir():
        for f in sorted(agents_dir.glob("*.md")):
            text = f.read_text(errors="replace")
            tools = []
            m = re.search(r"^tools:\n((?:\s+-\s+.+\n)+)", text, re.M)
            if m:
                tools = [ln.strip().lstrip("- ") for ln in m.group(1).splitlines() if ln.strip()]
            body = re.sub(r"^---\n.*?\n---\n", "", text, flags=re.S)
            out["subagents"].append({"name": f.stem, "tools": tools, "template": body})
    return out


# ---------------------------------------------------------------------------
# Render helpers
# ---------------------------------------------------------------------------


def fmt_args(name: str, args: dict) -> str:
    """Human-readable one-liner for a tool call."""
    if name == "Bash":
        return args.get("command", "")[:400]
    if name == "Read":
        return args.get("path") or args.get("file_path") or json.dumps(args)[:200]
    if name == "Write":
        return args.get("path") or args.get("file_path") or json.dumps(args)[:200]
    if name == "Grep":
        return f"pattern={args.get('pattern','')} path={args.get('path','')}"
    if name == "Glob":
        return args.get("glob_pattern") or args.get("pattern") or json.dumps(args)[:200]
    if name in ("Agent", "AgentSwarm"):
        return args.get("description") or args.get("subagent_type") or ""
    if name == "TodoList":
        todos = args.get("todos") or []
        return "; ".join(f"[{t.get('status','?')}] {t.get('title','')}" for t in todos if isinstance(t, dict))[:300]
    if name == "python_execute" or name.endswith("__python_execute"):
        code = args.get("code") or ""
        n_lines = code.count("\n") + 1 if code else 0
        return f"python_execute · {n_lines} line{'s' if n_lines != 1 else ''}"
    return json.dumps(args, ensure_ascii=False)[:300]


_PY_KEYWORDS = (
    "import", "from", "as", "def", "class", "return", "if", "elif", "else",
    "for", "while", "in", "not", "and", "or", "is", "None", "True", "False",
    "try", "except", "finally", "with", "raise", "pass", "break", "continue",
    "lambda", "yield", "del", "global", "nonlocal", "assert", "print",
)


_PY_STRING_RE = re.compile(
    r'('
    r'"(?:[^"\\]|\\.)*"'      # double-quoted
    r"|'(?:[^'\\]|\\.)*'"     # single-quoted
    r'|"""(?:[^"]|"(?!""))*"""'  # triple double
    r"|'''(?:[^']|'(?!''))*'''"  # triple single
    r')', re.S)
_PY_COMMENT_RE = re.compile(r'#(?!\\).*$', re.M)


# Cap rendered code/JSON so a single pathological call cannot balloon the
# audit page into tens of MB.
_MAX_CODE_CHARS = 8000
_MAX_JSON_CHARS = 6000
_MAX_PLAIN_RESULT_CHARS = 3000


def _render_python_block(code: str) -> str:
    """Render Python source with light syntax highlighting (no deps).

    Tokenize first (strings/comments), then escape each segment and wrap in
    spans so we never run keyword regex over already-escaped HTML.
    """
    if len(code) > _MAX_CODE_CHARS:
        code = code[:_MAX_CODE_CHARS] + "\n… (truncated)"
    # find all strings first
    spans: list[tuple[int, int, str]] = []  # (start, end, css_class)
    for m in _PY_STRING_RE.finditer(code):
        spans.append((m.start(), m.end(), "py-str"))
    # find comments that are not inside a string
    for m in _PY_COMMENT_RE.finditer(code):
        s, e = m.start(), m.end()
        inside_str = any(a <= s < b for a, b, c in spans if c == "py-str")
        if not inside_str:
            spans.append((s, e, "py-com"))
    spans.sort()

    kw_pat = re.compile(r'\b(' + '|'.join(_PY_KEYWORDS) + r')\b')
    num_pat = re.compile(r'\b(\d+(?:\.\d+)?)\b')

    def _plain(text: str) -> str:
        esc = html.escape(text)
        esc = kw_pat.sub(r'<span class="py-kw">\1</span>', esc)
        esc = num_pat.sub(r'<span class="py-num">\1</span>', esc)
        return esc

    out: list[str] = []
    pos = 0
    for s, e, cls in spans:
        if s > pos:
            out.append(_plain(code[pos:s]))
        out.append(f'<span class="{cls}">{html.escape(code[s:e])}</span>')
        pos = e
    if pos < len(code):
        out.append(_plain(code[pos:]))
    return f'<pre class="code-block py">{"".join(out)}</pre>'


def _render_shell_block(cmd: str) -> str:
    esc = html.escape(cmd)
    esc = re.sub(r'(?m)^(\s*[$#]\s+)', r'<span class="sh-prompt">\1</span>', esc)
    return f'<pre class="code-block sh">{esc}</pre>'


_EMBEDDED_JSON_RE = re.compile(
    r'(\{(?:[^{}"]|"(?:[^"\\]|\\.)*"|"(?:[^"\\]|\\.)*"\s*:\s*\{(?:[^{}]|\{(?:[^{}])*\})*\})*\})')


def _split_embedded_json(text: str):
    """Split mixed text like 'prose <mcp-structured-result>{json}</mcp-structured-result>'
    into segments; returns list of (kind, value) where kind is 'text' or 'json'."""
    segments = []
    pos = 0
    for m in re.finditer(r'\{[\s\S]*\}', text):
        start, end = m.start(), m.end()
        if pos < start:
            segments.append(("text", text[pos:start]))
        segments.append(("json", text[start:end]))
        pos = end
    if pos < len(text):
        segments.append(("text", text[pos:]))
    # keep only well-formed JSON segments as json; others stay text
    out = []
    for kind, val in segments:
        if kind == "json":
            try:
                json.loads(val)
                out.append(("json", val))
                continue
            except Exception:
                pass
        out.append(("text", val))
    return out


def _render_json_block(text: str, cls: str = "code-block js") -> str:
    """Pretty-print JSON-ish text with key/string/number highlighting.

    Tries a real json.loads round-trip first (handles objects, arrays and
    top-level scalars); on failure falls back to best-effort bracket/quote
    highlighting so the block still reads better than one raw line.
    """
    raw = text.strip()
    if not raw:
        return ""
    truncated = len(raw) > _MAX_JSON_CHARS
    raw = raw[:_MAX_JSON_CHARS]
    pretty = None
    try:
        obj = json.loads(raw)
        pretty = json.dumps(obj, ensure_ascii=False, indent=2)
    except Exception:
        # String args are often single-quoted Python dicts; try ast.literal_eval
        try:
            import ast
            obj = ast.literal_eval(raw)
            pretty = json.dumps(obj, ensure_ascii=False, indent=2, default=str)
        except Exception:
            pretty = None
    if truncated:
        pretty = None  # don't re-format a truncated document
    if pretty is not None:
        esc = html.escape(pretty)
        # html.escape turns " into &quot;. String literals may contain escaped
        # quotes (\"), which escape() renders as \&quot;. Match a full quoted
        # literal including inner escapes so nested quotes stay one span.
        esc = re.sub(r'&quot;(?:[^&\\]|\\.|&(?!quot;))*?&quot;(?=\s*:)', lambda m: f'<span class="js-key">{m.group(0)}</span>', esc)
        esc = re.sub(r'(?<=[:\[,])(\s*)(&quot;(?:[^&\\]|\\.|&(?!quot;))*?&quot;)', lambda m: f'{m.group(1)}<span class="js-str">{m.group(2)}</span>', esc)
        esc = re.sub(r':\s*(-?\b\d+(?:\.\d+)?\b)', r': <span class="js-num">\1</span>', esc)
        esc = re.sub(r':\s*\b(true|false|null)\b', r': <span class="js-kw">\1</span>', esc)
        return f'<pre class="{cls}">{esc}</pre>'
    # Mixed prose + embedded JSON (e.g. MCP results wrapped in
    # <mcp-structured-result>...</mcp-structured-result>): render each part
    # in its own style so the JSON payload is still readable.
    parts = _split_embedded_json(raw)
    if len(parts) > 1 or (parts and parts[0][0] == "json"):
        rendered = []
        for kind, val in parts:
            if kind == "json":
                rendered.append(_render_json_block(val, cls))
            else:
                t = html.escape(val)
                if truncated and val is parts[-1][1]:
                    t += "\n… (truncated)"
                rendered.append(f'<pre class="code-block mixed">{t}</pre>')
        return "".join(rendered)
    esc = html.escape(raw)
    if truncated:
        esc += "\n… (truncated)"
    return f'<pre class="{cls}">{esc}</pre>'


def _json_block_from_value(value, cls: str = "code-block js") -> str:
    """Pretty-print an already-parsed Python value as highlighted JSON."""
    try:
        pretty = json.dumps(value, ensure_ascii=False, indent=2)
    except Exception:
        return _render_json_block(str(value), cls)
    return _render_json_block(pretty, cls)


def _is_python_tool(name: str) -> bool:
    return name == "python_execute" or name.endswith("__python_execute")


def _is_shell_tool(name: str) -> bool:
    return name == "Bash" or name.endswith("__terminal") or name.endswith("__run_terminal_cmd")


def render_call(call: dict, api_error: bool = False) -> str:
    name = call["name"]
    is_deleg = name in DELEGATION_TOOLS
    cls = "call deleg" if is_deleg else "call"
    if api_error:
        cls += " api-error"
    oneline = fmt_args(name, call["args"])
    err = '<span class="err-badge">error</span>' if call["is_error"] else ""
    api_badge = '<span class="api-badge">API error step · excluded from metrics</span>' if api_error else ""

    parts = [f'<div class="{cls}"><div class="call-h">'
             f'<span class="tool-tag">{html.escape(name)}</span> {err} {api_badge}'
             f'<span class="call-line">{html.escape(oneline)}</span></div>']

    # Code-bearing tools get a dedicated, formatted block instead of the
    # generic JSON dump. This keeps the one-liner compact and lets the
    # auditor read the script in isolation.
    rendered_code = None
    code_label = None
    if _is_python_tool(name):
        code = call["args"].get("code") or ""
        if code:
            rendered_code = _render_python_block(code)
            code_label = "code"
    elif _is_shell_tool(name):
        cmd = call["args"].get("command") or call["args"].get("cmd") or ""
        if cmd:
            rendered_code = _render_shell_block(cmd)
            code_label = "command"

    if rendered_code is not None:
        parts.append(f'<details class="args code-args"><summary>{code_label}</summary>'
                     f'{rendered_code}</details>')
        # Other args besides the code field, if any, still go to JSON.
        other_args = {k: v for k, v in call["args"].items()
                      if k not in ("code", "command", "cmd")}
        if other_args:
            parts.append(f'<details class="args"><summary>other args</summary>'
                         f'{_json_block_from_value(other_args)}</details>')
    else:
        # Prefer a real JSON parse so nested structures render as a readable
        # tree instead of one long line; _render_json_block handles the
        # fallback path (plain strings, numbers, malformed text).
        arg_text = call.get("args_text")
        has_args = bool(call["args"]) or bool(arg_text and arg_text.strip() not in ("{}", ""))
        if has_args and isinstance(call["args"], dict) and not arg_text:
            parts.append(f'<details class="args"><summary>arguments</summary>'
                         f'{_json_block_from_value(call["args"])}</details>')
        elif has_args and arg_text:
            parts.append(f'<details class="args"><summary>arguments</summary>'
                         f'{_render_json_block(arg_text)}</details>')

    # structured delegation
    if name == "Agent":
        a = call["args"]
        prompt = (a.get("prompt") or "")
        parts.append(f'<div class="deleg-box"><div class="deleg-meta">'
                     f'<b>subagent_type:</b> {html.escape(str(a.get("subagent_type","?")))} · '
                     f'<b>description:</b> {html.escape(str(a.get("description","")))}</div>'
                     f'<div class="md-pre prompt">{html.escape(prompt)}</div></div>')
    elif name == "AgentSwarm":
        a = call["args"]
        items = a.get("items") or a.get("tasks") or []
        parts.append(f'<div class="deleg-box"><div class="deleg-meta"><b>AgentSwarm</b> — {len(items)} parallel item(s)</div>')
        for it in items[:30]:
            if isinstance(it, dict):
                parts.append(f'<div class="md-pre prompt">{html.escape(str(it.get("prompt") or it.get("task") or it))}</div>')
        parts.append("</div>")

    if call["result"] is not None:
        res = call["result"]
        if len(res) <= _MAX_PLAIN_RESULT_CHARS:
            res_cls = "call-r err" if call["is_error"] else "call-r"
            parts.append(f'<details class="res"><summary>result</summary>'
                         f'<div class="{res_cls}">{_render_json_block(res, "code-block js res-js")}</div></details>')
        else:
            short = res[:_MAX_PLAIN_RESULT_CHARS] + "\n… (truncated)"
            res_cls = "call-r err" if call["is_error"] else "call-r"
            parts.append(f'<details class="res"><summary>result ({len(res)} chars)</summary>'
                         f'<div class="{res_cls} md-pre">{html.escape(short)}</div></details>')
    parts.append("</div>")
    return "".join(parts)


def render_timeline(main_wire: dict) -> str:
    parts = []
    for st in main_wire.get("all_steps") or main_wire["steps"]:
        sn = st["step"]
        content = (st.get("content") or "").strip()
        think = ""
        m = re.search(r"<think>(.*?)</think>", content, re.S)
        if m:
            think = m.group(1).strip()
        narrative = re.sub(r"<think>.*?</think>", "", content, flags=re.S).strip()

        n_calls = len(st["calls"])
        deleg_n = sum(1 for c in st["calls"] if c["name"] in DELEGATION_TOOLS)
        badge = f'{n_calls} tool{"s" if n_calls!=1 else ""}'
        if deleg_n:
            badge += f' · <span class="par">{deleg_n} delegation{"s" if deleg_n!=1 else ""}</span>'

        api_error = bool(st.get("api_error"))
        step_cls = "step api-error" if api_error else "step"
        if api_error:
            badge += ' · <span class="api-text">API error · excluded</span>'
        parts.append(f'<div class="{step_cls}"><div class="step-h"><span class="step-no">Step {sn}</span>'
                     f'<span class="badge">{badge}</span></div>')
        if think:
            parts.append(f'<details class="think"><summary>reasoning</summary>'
                         f'<div class="md-pre">{html.escape(think)}</div></details>')
        if narrative:
            parts.append(f'<div class="narrative md-pre">{html.escape(narrative)}</div>')
        for call in st["calls"]:
            parts.append(render_call(call, api_error=api_error))
        if n_calls == 0 and not narrative and not think:
            parts.append('<div class="muted small">(no tool call / no output this step)</div>')
        parts.append("</div>")
    return "\n".join(parts)


def render_subagents(sub_wires: list[dict]) -> str:
    if not sub_wires:
        return "<p class='muted'>No sub-agents were delegated in this run.</p>"
    parts = []
    for i, w in enumerate(sub_wires):
        prof = w.get("profile") or "?"
        head = (f'Sub-agent #{i} <span class="tag">{html.escape(prof)}</span> '
                f'<span class="muted">{w["n_steps"]} step(s)</span>')
        parts.append(f'<details class="sub"><summary>{head}</summary>')
        if w.get("prompt"):
            parts.append(f'<h4>Prompt</h4><div class="md-pre prompt">{html.escape(w["prompt"])}</div>')
        for st in w.get("all_steps") or w["steps"]:
            api_error = bool(st.get("api_error"))
            substep_cls = "substep api-error" if api_error else "substep"
            api_note = ' <span class="api-badge">API error step · excluded from metrics</span>' if api_error else ""
            parts.append(f'<div class="{substep_cls}"><div class="substep-h">step {st["step"]}{api_note}</div>')
            content = (st.get("content") or "").strip()
            if content:
                content = re.sub(r"<think>.*?</think>", "", content, flags=re.S).strip()
            if content:
                parts.append(f'<div class="md-pre narrative">{html.escape(content[:600])}</div>')
            for call in st["calls"]:
                parts.append(render_call(call, api_error=api_error))
            if not st["calls"]:
                parts.append('<div class="muted small">(no tool this step)</div>')
            parts.append("</div>")
        parts.append("</details>")
    return "\n".join(parts)


CSS = """
:root{--bg:#faf9f7;--card:#fff;--bd:#e8e6e1;--fg:#1f1e1b;--mut:#6b6862;--acc:#b4552d;
--acc-l:#f6ede7;--ok:#2a7d4f;--bad:#c13434;--code:#f4f2ee;--blue:#2f5f8f;--par:#7d54b2;--api:#d97706;--api-l:#fff4d6}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.6 'Söhne','Inter',-apple-system,'Segoe UI',sans-serif}
header{padding:22px 32px;border-bottom:1px solid var(--bd);background:var(--card);position:sticky;top:0;z-index:5}
h1{font-size:20px;margin:0;font-weight:600;letter-spacing:-.01em}
.muted{color:var(--mut)} .small{font-size:12px}
nav{margin-top:8px;font-size:13px} nav a{color:var(--acc);margin-right:16px;text-decoration:none}
nav a:hover{text-decoration:underline}
main{padding:26px 32px;max-width:1120px;margin:auto}
section{background:var(--card);border:1px solid var(--bd);border-radius:10px;padding:20px 22px;margin-bottom:20px}
h2{font-size:16px;margin:0 0 14px;font-weight:600}
h3{font-size:12px;margin:18px 0 6px;color:var(--mut);text-transform:uppercase;letter-spacing:.5px;font-weight:600}
h4{font-size:13px;margin:12px 0 4px;font-weight:600}
table{border-collapse:collapse;width:100%;font-size:13.5px}
td,th{padding:8px 12px;border-bottom:1px solid var(--bd);text-align:left;vertical-align:top}
th{color:var(--mut);font-weight:600;font-size:12px;text-transform:uppercase;letter-spacing:.4px}
.pill{display:inline-block;padding:3px 12px;border-radius:14px;font-weight:600;font-size:12.5px}
.pill.ok{background:#e6f3ec;color:var(--ok)}
.pill.bad{background:#fbeaea;color:var(--bad)}
.kv{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:10px 24px;font-size:14px}
.kv b{color:var(--mut);font-weight:600;display:block;font-size:12px;text-transform:uppercase;letter-spacing:.4px}
.bar{height:8px;border-radius:4px;background:#efece7;overflow:hidden;margin-top:6px}
.bar>div{height:100%;background:linear-gradient(90deg,var(--ok),#5aa97e)}
.md-pre{background:var(--code);border:1px solid var(--bd);border-radius:8px;padding:11px 13px;
font:12.5px/1.55 'SF Mono','Menlo',monospace;white-space:pre-wrap;word-break:break-word;
max-height:430px;overflow:auto;margin:6px 0}
.md-pre.prompt{background:#fdf6ee;border-color:#f0e0cd}
.narrative{background:#f0f4f9;border-color:#dbe4ee}
details{margin:5px 0}
details summary{cursor:pointer;font-size:13px;color:var(--acc);padding:3px 0;list-style:none}
details summary::before{content:'▸ '}
details[open] summary::before{content:'▾ '}
details.args summary,details.res summary,details.think summary{color:var(--mut);font-size:12px}
.step{border:1px solid var(--bd);border-radius:10px;padding:14px 16px;margin:12px 0;background:#fdfcfb}
.step-h{display:flex;align-items:center;gap:10px;margin-bottom:8px}
.step-no{font-weight:600;font-size:14px}
.badge{background:#f0ede8;color:var(--mut);border-radius:12px;padding:2px 10px;font-size:11.5px}
.badge .par{color:var(--par);font-weight:600}
.call{border-left:3px solid var(--blue);background:#fbfcfd;border:1px solid var(--bd);border-radius:8px;padding:8px 12px;margin:8px 0}
.call.deleg{border-left-color:var(--par);background:#faf7fc}
.step.api-error,.substep.api-error{border-color:#f0b65d;background:var(--api-l)}
.call.api-error{border-left-color:var(--api);background:#fffaf0}
.api-badge{background:#fff0bf;color:#9a5700;border:1px solid #e8b85a;border-radius:6px;padding:1px 7px;font-size:11px;font-weight:700}
.api-text{color:var(--api);font-weight:700}
.call-h{display:flex;align-items:baseline;gap:8px;flex-wrap:wrap}
.tool-tag{background:var(--blue);color:#fff;border-radius:6px;padding:1px 8px;font-size:11px;font-weight:600;font-family:monospace}
.call.deleg .tool-tag{background:var(--par)}
.call-line{font-family:monospace;font-size:12px;color:#3a3a38;word-break:break-all}
.err-badge{background:#fbeaea;color:var(--bad);border-radius:6px;padding:1px 7px;font-size:11px;font-weight:600}
.deleg-box{margin:8px 0;padding:8px;border:1px dashed #d8c8e6;border-radius:8px}
.deleg-meta{font-size:12px;color:var(--mut);margin-bottom:4px}
.call-r{background:#f6f6f4}.call-r.err{background:#fbeaea;color:var(--bad)}
.call-r{border:1px solid var(--bd);border-radius:8px;padding:4px}
.call-r.err{color:var(--bad)}
details.sub{border:1px solid var(--bd);border-radius:10px;background:#fdfcfb;margin:10px 0;padding:0}
details.sub>summary{padding:12px 16px;font-weight:600;font-size:14px;color:var(--fg)}
details.sub[open]>summary{border-bottom:1px solid var(--bd)}
details.sub>*:not(summary){margin-left:16px;margin-right:16px}
.tag{background:var(--acc-l);color:var(--acc);border-radius:8px;padding:2px 9px;font-size:12px;font-weight:600}
.substep{border-left:2px solid var(--bd);padding:6px 0 6px 14px;margin:10px 0}
.substep-h{font-weight:600;font-size:12px;color:var(--mut)}
.phase.parallel td:first-child{border-left:3px solid var(--par)}
.phase.sequential td:first-child{border-left:3px solid var(--blue)}
.phase.solo td:first-child{border-left:3px solid #cfc9c0}
.footer{text-align:center;color:var(--mut);font-size:12px;padding:26px}
a{color:var(--acc)}

/* code blocks */
.code-block{background:#2b2b2b;color:#e8e6e3;border:1px solid #1d1d1d;border-radius:8px;
padding:10px 12px;font:12px/1.55 'SF Mono','Menlo','DejaVu Sans Mono',monospace;
white-space:pre;overflow-x:auto;overflow-y:auto;max-height:480px;margin:6px 0;
tab-size:4;-moz-tab-size:4}
.code-block.py{background:#1e1e2e;border-color:#2f2f42}
.code-block.py .py-kw{color:#cba6f7;font-weight:600}
.code-block.py .py-str{color:#a6e3a1}
.code-block.py .py-num{color:#fab387}
.code-block.py .py-com{color:#7f849c;font-style:italic}
.code-block.sh{background:#1b2327;border-color:#233038}
.code-block.sh .sh-prompt{color:#7aa2f7;font-weight:600}
.code-block.js{background:#161b26;border-color:#232a3a}
.code-block.mixed{background:#1d222c;border-color:#2a3040;color:#c8ccd4}
.code-block.js .js-key{color:#7aa2f7}
.code-block.js .js-str{color:#a6e3a1}
.code-block.js .js-num{color:#fab387}
.code-block.js .js-kw{color:#cba6f7;font-weight:600}
.call-r .code-block.js{background:#20261d;border-color:#2c3526}
.call-r.err .code-block.js{background:#2a1d1d;border-color:#3a2626}
div.call-r:has(> .code-block), div.call-r.err:has(> .code-block){background:transparent;border:none;padding:0;margin:0;max-height:none;overflow:visible}
details.code-args summary{color:var(--mut);font-size:12px}
details.code-args[open] summary{margin-bottom:4px}
.btn-code{cursor:pointer;border:1px solid var(--bd);background:var(--card);color:var(--acc);
border-radius:8px;padding:4px 12px;font-size:12.5px;font-weight:600;margin:4px 0}
.btn-code:hover{background:var(--acc-l)}
"""

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Audit — {task} / {run}</title>
<style>{css}</style></head><body>
<header>
<h1>Case Audit — {task} <span class="muted">/ {run}</span></h1>
<nav><a href="#summary">Summary</a><a href="#critical">Critical Steps</a><a href="#prompts">Prompts &amp; Tools</a><a href="#subdef">Sub-agent Defs</a><a href="#timeline">Timeline</a><a href="#subrun">Sub-agent Runs</a><a href="#eval">Eval</a></nav>
</header><main>
{body}
<div class="footer">generated by scripts/audit_html_gen.py</div>
<script>
function toggleCodeBlocks(btn, sectionId) {{
  var sec = document.getElementById(sectionId);
  if (!sec) return;
  var open = btn.dataset.open === '1';
  var target = !open;
  sec.querySelectorAll('details.code-args').forEach(function(d) {{ d.open = target; }});
  btn.dataset.open = target ? '1' : '0';
  btn.textContent = target ? 'collapse all code' : 'expand all code';
}}
</script>
</main></body></html>"""


def build_case_html(case_dir: Path) -> dict | None:
    inner = find_inner_dir(case_dir)
    if inner is None:
        return None
    task, run = case_dir.parent.name, case_dir.name
    runlog = parse_run_log(case_dir)
    evres = parse_eval_res(inner)
    artifact_evres = parse_artifact_eval_res(inner)
    traj_log = load_json(inner / "traj_log.json") or {}
    cfg = traj_log.get("config") or {}

    agent_main = inner / "agent_main.md"
    agent_main_text = agent_main.read_text(errors="replace") if agent_main.exists() else "(agent_main.md not found)"
    inv = parse_tools_inventory(inner)

    sessions = list((inner / ".kimi_home" / "sessions").glob("*/*/agents"))
    main_wire: dict = {"steps": [], "n_steps": 0, "profile": None, "prompt": None}
    sub_wires: list[dict] = []
    if sessions:
        agents_root = sessions[0]
        mw = agents_root / "main" / "wire.jsonl"
        if mw.exists():
            main_wire = parse_wire(mw)
        for sub_dir in sorted(agents_root.iterdir()):
            if sub_dir.name.startswith("agent-") and (sub_dir / "wire.jsonl").exists():
                sub_wires.append(parse_wire(sub_dir / "wire.jsonl"))

    crit = compute_critical_steps(main_wire, sub_wires)
    pf = compute_plan_first_metrics(main_wire)
    plan_first_arm = "Plan-First Protocol" in agent_main_text

    ck_pass, ck_total = runlog.get("ckpt_pass"), runlog.get("ckpt_total")
    ck_fail = None
    if ck_pass is None:
        # count_checkpoints returns (pass_count, fail_count), NOT (pass, total).
        cnt = count_checkpoints(evres.get("failure_text", ""))
        if cnt:
            ck_pass, ck_fail = cnt
            ck_total = ck_pass + ck_fail
    else:
        ck_fail = (ck_total - ck_pass) if (ck_pass is not None and ck_total is not None) else None

    passed = evres.get("pass")
    pass_pill = ('<span class="pill ok">PASS</span>' if passed is True
                 else '<span class="pill bad">FAIL</span>' if passed is False
                 else '<span class="pill bad">NO EVAL</span>')
    artifact_passed = artifact_evres.get("pass")
    artifact_pill = ('<span class="pill ok">PASS</span>' if artifact_passed is True
                     else '<span class="pill bad">FAIL</span>' if artifact_passed is False
                     else '<span class="pill">not run</span>')
    duration = runlog.get("duration_s")
    dur_txt = f"{duration}s ({duration//60}m{duration%60}s)" if duration else "—"
    ck_txt = f"{ck_pass} / {ck_total} pass · {ck_fail} fail" if ck_pass is not None else "—"
    ck_pct = (100 * ck_pass / ck_total) if (ck_pass and ck_total) else 0

    n_mcp = len(cfg.get("needed_mcp_servers") or [])
    if n_mcp == 0:
        # traj_log.json may have been overwritten by the evaluator, wiping the
        # config. Fall back to the MCP servers actually registered for the run.
        n_mcp = len(inv.get("mcp_servers") or [])
    tier = tier_from_mcp_count(n_mcp)

    if plan_first_arm:
        if pf["plan_first_ok"] is True:
            pf_pill = (f'<span class="pill ok">OK</span> <span class="muted">'
                       f'{pf["pre_plan_actions"]} pre-action(s) · '
                       f'{pf["n_plan_dispatches"]} plan dispatch(es)</span>')
        elif pf["plan_first_ok"] is False:
            first_disp = pf.get("first_dispatch") or "none"
            pf_pill = (f'<span class="pill bad">VIOLATED</span> <span class="muted">'
                       f'first dispatch: {html.escape(str(first_disp))} · '
                       f'{pf["pre_plan_actions"] if pf["pre_plan_actions"] is not None else "—"} '
                       f'pre-action(s) before any plan dispatch</span>')
        else:
            pf_pill = '<span class="pill">n/a</span> <span class="muted">no dispatch observed</span>'
    else:
        pf_pill = ""

    body = []
    body.append(f"""<section id="summary"><h2>Run Summary</h2>
<div class="kv">
<div><b>Eval</b>{pass_pill}</div>
<div><b>Artifact Eval</b>{artifact_pill}</div>
<div><b>Split / Tier</b><span class="tag">{tier}</span> <span class="muted">{n_mcp} MCP server(s)</span></div>
<div><b>Checkpoints</b>{html.escape(ck_txt)}<div class="bar"><div style="width:{ck_pct:.0f}%"></div></div></div>
<div><b>Duration</b>{html.escape(dur_txt)}</div>
<div><b>claim_done</b>{runlog.get('claim_done')}</div>
<div><b>Main steps</b>{crit['main_steps']}</div>
<div><b>Critical Steps</b><span style="color:var(--acc);font-weight:700;font-size:17px">{crit['critical_steps']}</span> <span class="muted">(serial {crit['serial_steps']})</span></div>
<div><b>Sub-agents</b>{len(sub_wires)} — {crit['n_parallel_subs']} parallel / {crit['n_sequential_subs']} sequential</div>
{f'<div><b>Plan-first</b>{pf_pill}</div>' if plan_first_arm else ''}
</div></section>""")

    rows = []
    for p in crit["phases"]:
        if p["mode"] == "solo":
            rows.append(f'<tr class="phase solo"><td>{p["phase"]}</td><td>solo</td><td>—</td>'
                        f'<td>{p["s_main"]}</td><td>—</td><td>{p["cost"]}</td><td>{p.get("token_cost_api", 0)}</td></tr>')
        else:
            subs = ", ".join(map(str, p["sub_steps"])) or "—"
            rows.append(f'<tr class="phase {p["mode"]}"><td>{p["phase"]}</td>'
                        f'<td>{p["mode"]} <span class="muted">({html.escape(str(p["deleg_tool"]))})</span></td>'
                        f'<td>{p["n_sub"]}</td><td>{p["s_main"]}</td><td>[{subs}] → {p["slowest"]}</td><td>{p["cost"]}</td><td>{p.get("token_cost_api", 0)}</td></tr>')
    body.append(f"""<section id="critical"><h2>Critical Steps</h2>
<p class="muted">CriticalSteps = Σ<sub>t</sub>(S<sub>main</sub><sup>(t)</sup> + max<sub>i</sub> S<sub>sub,i</sub><sup>(t)</sup>) = <b style="color:var(--acc)">{crit['critical_steps']}</b>
· serial equivalent {crit['serial_steps']} · parallel saving {max(0, crit['serial_steps']-crit['critical_steps'])}.
A single step firing N Agent calls concurrently counts as one parallel phase.</p>
<table><tr><th>Phase</th><th>Mode</th><th>#Sub</th><th>S_main</th><th>S_sub,i (max)</th><th>Cost</th></tr>{''.join(rows)}</table></section>""")

    needed_mcp = cfg.get("needed_mcp_servers") or []
    needed_local = cfg.get("needed_local_tools") or []
    sys_agent = (cfg.get("system_prompts") or {}).get("agent") or ""
    task_str = cfg.get("task_str") or ""

    # Fallback: some evaluators overwrite traj_log.json with just {pass,fail},
    # wiping the config the harness wrote. Recover what we can from sibling
    # files so the audit page always shows the prompt + instruction.
    if not sys_agent:
        _ap = inner / "agent_main.md"
        if _ap.exists():
            _t = _ap.read_text(errors="replace")
            # agent_main.md frontmatter ends at the second '---'; the body
            # after that is the system prompt text.
            _parts = _t.split("---", 2)
            if len(_parts) >= 3:
                sys_agent = _parts[2].strip()
    if not task_str:
        _rl = (case_dir / "run.log").read_text(errors="replace") if (case_dir / "run.log").exists() else ""
        import re as _re
        _m = _re.search(r"\[kimi\] launching: kimi -p (.+?)(?:\s*\.\.\.\s*\(home=|$)", _rl, _re.S)
        if _m:
            task_str = _m.group(1).strip()
    body.append(f"""<section id="prompts"><h2>Prompts &amp; Tool Grants</h2>
<h3>Allowed MCP servers ({len(needed_mcp)})</h3><p>{html.escape(', '.join(needed_mcp) or '—')}</p>
<h3>Allowed local tools ({len(needed_local)})</h3><p>{html.escape(', '.join(needed_local) or '—')}</p>
<h3>MCP servers registered (.kimi_home/mcp.json)</h3><p>{html.escape(', '.join(inv['mcp_servers']) or '—')}</p>
<h3>Sub-agents enabled</h3><p>{html.escape(', '.join(s['name'] for s in inv['subagents']) or 'none')}</p>
<h3>Task system prompt</h3><div class="md-pre">{html.escape(sys_agent or '—')}</div>
<h3>Task instruction</h3><div class="md-pre">{html.escape(task_str or '—')}</div>
<h3>Main agent prompt (agent_main.md)</h3><div class="md-pre">{html.escape(agent_main_text)}</div>
</section>""")

    sd = []
    for s in inv["subagents"]:
        sd.append(f"<details><summary><b>{html.escape(s['name'])}</b> — {len(s['tools'])} tool pattern(s)</summary>"
                  f"<h3>Tool patterns</h3><div class='md-pre'>{html.escape(chr(10).join(s['tools']))}</div>"
                  f"<h3>Template</h3><div class='md-pre'>{html.escape(s['template'])}</div></details>")
    body.append(f"""<section id="subdef"><h2>Sub-agent Definitions ({len(inv['subagents'])})</h2>{''.join(sd) or '<p class="muted">none</p>'}</section>""")

    body.append(f"""<section id="timeline"><h2>Main Agent Execution Timeline ({main_wire['n_steps']} steps)</h2>
<p class="muted">Purple = sub-agent delegation · blue = tool call. Expand for full arguments/results.
Python code is rendered with syntax highlighting inside a dedicated block. Amber steps are model/provider API-error steps and are excluded from step/token metrics.</p>
<button class="btn-code" onclick="toggleCodeBlocks(this, 'timeline')">expand all code</button>
{render_timeline(main_wire)}</section>""")

    body.append(f"""<section id="subrun"><h2>Sub-agent Runs ({len(sub_wires)})</h2>
<p class="muted">{crit['n_parallel_subs']} in parallel · {crit['n_sequential_subs']} sequential · {crit['sub_api_error_steps']} API-error sub-agent step(s) excluded.</p>
<button class="btn-code" onclick="toggleCodeBlocks(this, 'subrun')">expand all code</button>
{render_subagents(sub_wires)}</section>""")

    artifact_detail = ""
    if artifact_evres:
        artifact_detail = f"""
<h3>Artifact-only Evaluation</h3>
<p class="muted">Supplemental audit only: this ignores the missing claim_done lifecycle signal and evaluates the artifacts currently present on disk. It does not change the official Eval above.</p>
<div class="md-pre">{html.escape(artifact_evres.get('failure_text') or artifact_evres.get('details') or '(empty artifact eval)')}</div>"""
        if artifact_evres.get("stderr"):
            artifact_detail += f"""
<h3>Artifact Eval STDERR</h3>
<div class="md-pre">{html.escape(artifact_evres.get('stderr') or '')}</div>"""

    body.append(f"""<section id="eval"><h2>Evaluation Detail</h2>
<h3>Official Evaluation</h3>
<div class="md-pre">{html.escape(evres.get('failure_text') or '(no eval_res.json)')}</div>
{artifact_detail}</section>""")

    (case_dir / "audit.html").write_text(
        HTML_TEMPLATE.format(task=html.escape(task), run=html.escape(run),
                             css=CSS, body="\n".join(body)))
    return {"task": task, "run": run, "path": str(case_dir / "audit.html"),
            "pass": passed, "ckpt_pass": ck_pass, "ckpt_total": ck_total,
            "duration_s": duration, "critical_steps": crit["critical_steps"],
            "serial_steps": crit["serial_steps"], "main_steps": crit["main_steps"],
            "main_steps_raw": crit["main_steps_raw"], "api_error_steps": crit["api_error_steps"],
            "main_api_error_steps": crit["main_api_error_steps"], "sub_api_error_steps": crit["sub_api_error_steps"],
            "prompt_tokens_api": crit["prompt_tokens_api"], "assistant_tokens_api": crit["assistant_tokens_api"],
            "total_tokens_api": crit["total_tokens_api"], "critical_tokens_api": crit["critical_tokens_api"],
            "serial_tokens_api": crit["serial_tokens_api"], "observation_tokens_est": crit["observation_tokens_est"],
            "critical_all_tokens_est": crit["critical_all_tokens_est"], "serial_all_tokens_est": crit["serial_all_tokens_est"],
            "excluded_total_tokens_api": crit["excluded_total_tokens_api"],
            "excluded_observation_tokens_est": crit["excluded_observation_tokens_est"],
            "n_subs": len(sub_wires), "n_parallel": crit["n_parallel_subs"],
            "n_sequential": crit["n_sequential_subs"], "claim_done": runlog.get("claim_done"),
            "plan_first_arm": plan_first_arm,
            "plan_first_ok": pf["plan_first_ok"] if plan_first_arm else None,
            "pre_plan_actions": pf["pre_plan_actions"] if plan_first_arm else None,
            "n_plan_dispatches": pf["n_plan_dispatches"] if plan_first_arm else None}


def write_dump_index(dump_root: Path, records: list[dict]) -> None:
    rows = []
    has_pf = any(r.get("plan_first_arm") for r in records)
    pf_stat = ""
    if has_pf:
        arm = [r for r in records if r.get("plan_first_arm")]
        ok = sum(1 for r in arm if r.get("plan_first_ok") is True)
        pf_stat = (f"<p>Plan-first arm: {len(arm)} run(s) · compliance {ok}/{len(arm)}"
                   f" ({100 * ok / max(1, len(arm)):.0f}%)</p>")
    for r in sorted(records, key=lambda x: (x["task"], x["run"])):
        rel = f"{r['task']}/{r['run']}/audit.html"
        pill = ("ok'>PASS" if r["pass"] is True else "bad'>FAIL") if r["pass"] is not None else "bad'>N/A"
        ck = f"{r['ckpt_pass']}/{r['ckpt_total']}" if r.get("ckpt_pass") is not None else "—"
        dur = f"{r['duration_s']}s" if r.get("duration_s") else "—"
        pf_cell = ""
        if has_pf:
            if r.get("plan_first_arm"):
                pf_cell = ("<td><span class='pill ok'>OK</span></td>" if r.get("plan_first_ok") is True
                           else f"<td><span class='pill bad'>VIOLATED</span>"
                                f"<div style='color:#6b6862;font-size:11px'>{r.get('pre_plan_actions') if r.get('pre_plan_actions') is not None else '—'} pre</div></td>")
            else:
                pf_cell = "<td>—</td>"
        rows.append(f"<tr><td><a href='{rel}'>{r['task']}</a></td><td>{r['run']}</td>"
                    f"<td><span class='pill {pill}</span></td><td>{ck}</td><td>{dur}</td>"
                    f"<td>{r['critical_steps']}</td><td>{r['serial_steps']}</td>"
                    f"<td><span class='api-text'>{r.get('api_error_steps', 0)}</span></td>"
                    f"<td>{r.get('total_tokens_api', 0)}</td><td>{r.get('critical_tokens_api', 0)}</td>"
                    f"<td>{r['n_subs']} ({r['n_parallel']}p/{r['n_sequential']}s)</td>{pf_cell}</tr>")
    pf_head = "<th>Plan-first</th>" if has_pf else ""
    doc = f"""<!DOCTYPE html><html><head><meta charset='utf-8'>
<meta name='viewport' content='width=device-width,initial-scale=1'>
<title>Audit Index — {dump_root.name}</title>
<style>body{{background:#faf9f7;color:#1f1e1b;font:15px 'Söhne','Inter',sans-serif;padding:28px;max-width:1200px;margin:auto}}
table{{border-collapse:collapse;width:100%;background:#fff;border:1px solid #e8e6e1;border-radius:10px;overflow:hidden}}
td,th{{padding:9px 12px;border-bottom:1px solid #e8e6e1;text-align:left;font-size:13.5px}}
th{{background:#f4f2ee;color:#6b6862;font-size:12px;text-transform:uppercase}}
a{{color:#b4552d;text-decoration:none}}a:hover{{text-decoration:underline}}
.pill{{padding:3px 12px;border-radius:14px;font-size:12.5px;font-weight:600}}
.pill.ok{{background:#e6f3ec;color:#2a7d4f}}.pill.bad{{background:#fbeaea;color:#c13434}}.api-text{{color:#d97706;font-weight:700}}
h1{{font-size:20px}}p{{color:#6b6862}}</style></head><body>
<h1>Audit Index — {dump_root.name}</h1><p>{len(records)} case(s)</p>{pf_stat}
<table><tr><th>Task</th><th>Run</th><th>Pass</th><th>Checkpoints</th><th>Duration</th><th>Critical</th><th>Serial</th><th>API-error steps</th><th>Total tokens API</th><th>Critical tokens API</th><th>Subs</th>{pf_head}</tr>
{''.join(rows)}</table></body></html>"""
    (dump_root / "audit_index.html").write_text(doc)


def main(argv: list[str]) -> int:
    roots = [Path(a) for a in argv[1:]] or [Path("dumps/kimi-code")]
    for root in roots:
        if not root.is_dir():
            print(f"[skip] {root} not a directory")
            continue
        if RUN_DIR_RE.match(root.name):
            rec = build_case_html(root)
            if rec:
                print(f"[ok] {rec['task']}/{rec['run']}  pass={rec['pass']}  crit={rec['critical_steps']}  subs={rec['n_subs']}({rec['n_parallel']}p/{rec['n_sequential']}s)")
            continue
        records = []
        for case_dir in find_case_dirs(root):
            rec = build_case_html(case_dir)
            if rec:
                records.append(rec)
                print(f"[ok] {rec['task']}/{rec['run']}  pass={rec['pass']}  crit={rec['critical_steps']}  subs={rec['n_subs']}({rec['n_parallel']}p/{rec['n_sequential']}s)")
        if records:
            write_dump_index(root, records)
            (root / "audit_index.json").write_text(json.dumps(records, indent=2))
            print(f"[index] {root}/audit_index.html  ({len(records)} cases)")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
