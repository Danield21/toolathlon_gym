#!/usr/bin/env python3
"""Generate an audit HTML page per evaluated case, placed inside the case dump dir.

Anthropic-inspired light UI. Rich timeline: reasoning, structured tool calls,
tool results, sub-agent delegation. Critical Steps per formula 2, where a step
that fires multiple Agent calls concurrently counts as ONE parallel phase.
"""

from __future__ import annotations

import html
import json
import re
import sys
from pathlib import Path

DELEGATION_TOOLS = {"Agent", "AgentSwarm"}


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
            if run_dir.is_dir() and re.match(r"\d{8}-\d{6}_slot\d+", run_dir.name):
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


def count_checkpoints(failure_text: str):
    if not failure_text:
        return None
    p = len(re.findall(r"^\s*\[PASS\]", failure_text, re.M))
    f = len(re.findall(r"^\s*\[FAIL\]", failure_text, re.M))
    return (p, f) if p + f > 0 else None


# ---------------------------------------------------------------------------
# Wire parsing
# ---------------------------------------------------------------------------


def parse_wire(wire_path: Path) -> dict:
    steps: dict[int, dict] = {}
    cur_step: dict | None = None
    profile_name = None
    sub_prompt = None
    result_by_id: dict[str, dict] = {}

    if not wire_path.exists():
        return {"steps": [], "n_steps": 0, "profile": None, "prompt": None}

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
        elif t == "context.append_loop_event":
            ev = j.get("event", {})
            et = ev.get("type")
            sn = ev.get("step")
            if et == "step.begin":
                cur_step = {"step": sn, "calls": [], "content": ""}
                steps[sn] = cur_step
            elif et == "tool.call":
                call = {"name": ev.get("name"), "args": ev.get("args") or {},
                        "id": ev.get("toolCallId"), "result": None, "is_error": False}
                if cur_step is not None:
                    cur_step["calls"].append(call)
            elif et == "tool.result":
                res = ev.get("result") or {}
                output = res.get("output") if isinstance(res, dict) else str(res)
                is_error = bool(res.get("is_error")) if isinstance(res, dict) else False
                result_by_id[ev.get("toolCallId")] = {"output": str(output), "is_error": is_error}
            elif et == "content.part":
                part = ev.get("part") or {}
                txt = part.get("text") or ""
                if cur_step is not None and txt:
                    cur_step["content"] += txt + "\n"

    step_list = [steps[k] for k in sorted(steps)]
    for st in step_list:
        for call in st["calls"]:
            r = result_by_id.get(call["id"])
            if r:
                call["result"], call["is_error"] = r["output"], r["is_error"]
    return {"steps": step_list, "n_steps": len(step_list),
            "profile": profile_name, "prompt": sub_prompt}


def compute_critical_steps(main_wire: dict, sub_wires: list[dict]) -> dict:
    """A step that fires N Agent calls concurrently = ONE parallel phase.
    A lone Agent call = sequential phase. AgentSwarm with N items = parallel.
    """
    phases: list[dict] = []
    sub_cursor = 0
    phase_idx = 0
    for st in main_wire["steps"]:
        deleg = [c for c in st["calls"] if c["name"] in DELEGATION_TOOLS]
        if not deleg:
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
        slowest = max(sub_step_counts) if sub_step_counts else 0
        phase_idx += 1
        phases.append({"phase": phase_idx, "mode": mode,
                       "deleg_tool": "/".join(sorted(tools_used)),
                       "n_sub": len(taken), "s_main": 1,
                       "sub_steps": sub_step_counts, "slowest": slowest,
                       "cost": 1 + slowest})

    total_main = main_wire["n_steps"]
    n_deleg_phases = len(phases)
    solo = max(0, total_main - n_deleg_phases)
    if solo:
        phases.insert(0, {"phase": 0, "mode": "solo", "deleg_tool": None,
                          "n_sub": 0, "s_main": solo, "sub_steps": [],
                          "slowest": 0, "cost": solo})
    critical = sum(p["cost"] for p in phases)
    serial = total_main + sum(w["n_steps"] for w in sub_wires)
    n_parallel = sum(p["n_sub"] for p in phases if p["mode"] == "parallel")
    n_sequential = sum(p["n_sub"] for p in phases if p["mode"] == "sequential")
    return {"phases": phases, "critical_steps": critical, "serial_steps": serial,
            "n_parallel_subs": n_parallel, "n_sequential_subs": n_sequential,
            "main_steps": total_main}


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
    if name == "python_execute":
        return (args.get("code") or "")[:300]
    return json.dumps(args, ensure_ascii=False)[:300]


def render_call(call: dict) -> str:
    name = call["name"]
    is_deleg = name in DELEGATION_TOOLS
    cls = "call deleg" if is_deleg else "call"
    oneline = fmt_args(name, call["args"])
    err = '<span class="err-badge">error</span>' if call["is_error"] else ""

    parts = [f'<div class="{cls}"><div class="call-h">'
             f'<span class="tool-tag">{html.escape(name)}</span> {err}'
             f'<span class="call-line">{html.escape(oneline)}</span></div>']

    full_args = json.dumps(call["args"], ensure_ascii=False, indent=2)
    if len(full_args) > len(oneline) + 40:
        parts.append(f'<details class="args"><summary>arguments</summary>'
                     f'<div class="md-pre">{html.escape(full_args)}</div></details>')

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
        short = res if len(res) <= 800 else res[:800] + "\n… (truncated)"
        res_cls = "call-r err" if call["is_error"] else "call-r"
        parts.append(f'<details class="res"><summary>result</summary>'
                     f'<div class="{res_cls} md-pre">{html.escape(short)}</div></details>')
    parts.append("</div>")
    return "".join(parts)


def render_timeline(main_wire: dict) -> str:
    parts = []
    for st in main_wire["steps"]:
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

        parts.append(f'<div class="step"><div class="step-h"><span class="step-no">Step {sn}</span>'
                     f'<span class="badge">{badge}</span></div>')
        if think:
            parts.append(f'<details class="think"><summary>reasoning</summary>'
                         f'<div class="md-pre">{html.escape(think)}</div></details>')
        if narrative:
            parts.append(f'<div class="narrative md-pre">{html.escape(narrative)}</div>')
        for call in st["calls"]:
            parts.append(render_call(call))
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
        for st in w["steps"]:
            parts.append(f'<div class="substep"><div class="substep-h">step {st["step"]}</div>')
            content = (st.get("content") or "").strip()
            if content:
                content = re.sub(r"<think>.*?</think>", "", content, flags=re.S).strip()
            if content:
                parts.append(f'<div class="md-pre narrative">{html.escape(content[:600])}</div>')
            for call in st["calls"]:
                parts.append(render_call(call))
            if not st["calls"]:
                parts.append('<div class="muted small">(no tool this step)</div>')
            parts.append("</div>")
        parts.append("</details>")
    return "\n".join(parts)


CSS = """
:root{--bg:#faf9f7;--card:#fff;--bd:#e8e6e1;--fg:#1f1e1b;--mut:#6b6862;--acc:#b4552d;
--acc-l:#f6ede7;--ok:#2a7d4f;--bad:#c13434;--code:#f4f2ee;--blue:#2f5f8f;--par:#7d54b2}
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
.call-h{display:flex;align-items:baseline;gap:8px;flex-wrap:wrap}
.tool-tag{background:var(--blue);color:#fff;border-radius:6px;padding:1px 8px;font-size:11px;font-weight:600;font-family:monospace}
.call.deleg .tool-tag{background:var(--par)}
.call-line{font-family:monospace;font-size:12px;color:#3a3a38;word-break:break-all}
.err-badge{background:#fbeaea;color:var(--bad);border-radius:6px;padding:1px 7px;font-size:11px;font-weight:600}
.deleg-box{margin:8px 0;padding:8px;border:1px dashed #d8c8e6;border-radius:8px}
.deleg-meta{font-size:12px;color:var(--mut);margin-bottom:4px}
.call-r{background:#f6f6f4}.call-r.err{background:#fbeaea;color:var(--bad)}
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
</main></body></html>"""


def build_case_html(case_dir: Path) -> dict | None:
    inner = find_inner_dir(case_dir)
    if inner is None:
        return None
    task, run = case_dir.parent.name, case_dir.name
    runlog = parse_run_log(case_dir)
    evres = parse_eval_res(inner)
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

    ck_pass, ck_total = runlog.get("ckpt_pass"), runlog.get("ckpt_total")
    if ck_pass is None:
        cnt = count_checkpoints(evres.get("failure_text", ""))
        if cnt:
            ck_pass, ck_total = cnt
    ck_fail = (ck_total - ck_pass) if (ck_pass is not None and ck_total is not None) else None

    passed = evres.get("pass")
    pass_pill = ('<span class="pill ok">PASS</span>' if passed is True
                 else '<span class="pill bad">FAIL</span>' if passed is False
                 else '<span class="pill bad">NO EVAL</span>')
    duration = runlog.get("duration_s")
    dur_txt = f"{duration}s ({duration//60}m{duration%60}s)" if duration else "—"
    ck_txt = f"{ck_pass} / {ck_total} pass · {ck_fail} fail" if ck_pass is not None else "—"
    ck_pct = (100 * ck_pass / ck_total) if (ck_pass and ck_total) else 0

    body = []
    body.append(f"""<section id="summary"><h2>Run Summary</h2>
<div class="kv">
<div><b>Eval</b>{pass_pill}</div>
<div><b>Checkpoints</b>{html.escape(ck_txt)}<div class="bar"><div style="width:{ck_pct:.0f}%"></div></div></div>
<div><b>Duration</b>{html.escape(dur_txt)}</div>
<div><b>claim_done</b>{runlog.get('claim_done')}</div>
<div><b>Main steps</b>{crit['main_steps']}</div>
<div><b>Critical Steps</b><span style="color:var(--acc);font-weight:700;font-size:17px">{crit['critical_steps']}</span> <span class="muted">(serial {crit['serial_steps']})</span></div>
<div><b>Sub-agents</b>{len(sub_wires)} — {crit['n_parallel_subs']} parallel / {crit['n_sequential_subs']} sequential</div>
</div></section>""")

    rows = []
    for p in crit["phases"]:
        if p["mode"] == "solo":
            rows.append(f'<tr class="phase solo"><td>{p["phase"]}</td><td>solo</td><td>—</td>'
                        f'<td>{p["s_main"]}</td><td>—</td><td>{p["cost"]}</td></tr>')
        else:
            subs = ", ".join(map(str, p["sub_steps"])) or "—"
            rows.append(f'<tr class="phase {p["mode"]}"><td>{p["phase"]}</td>'
                        f'<td>{p["mode"]} <span class="muted">({html.escape(str(p["deleg_tool"]))})</span></td>'
                        f'<td>{p["n_sub"]}</td><td>{p["s_main"]}</td><td>[{subs}] → {p["slowest"]}</td><td>{p["cost"]}</td></tr>')
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
<p class="muted">Purple = sub-agent delegation · blue = tool call. Expand for full arguments/results.</p>
{render_timeline(main_wire)}</section>""")

    body.append(f"""<section id="subrun"><h2>Sub-agent Runs ({len(sub_wires)})</h2>
<p class="muted">{crit['n_parallel_subs']} in parallel · {crit['n_sequential_subs']} sequential.</p>
{render_subagents(sub_wires)}</section>""")

    body.append(f"""<section id="eval"><h2>Evaluation Detail</h2>
<div class="md-pre">{html.escape(evres.get('failure_text') or '(no eval_res.json)')}</div></section>""")

    (case_dir / "audit.html").write_text(
        HTML_TEMPLATE.format(task=html.escape(task), run=html.escape(run),
                             css=CSS, body="\n".join(body)))
    return {"task": task, "run": run, "path": str(case_dir / "audit.html"),
            "pass": passed, "ckpt_pass": ck_pass, "ckpt_total": ck_total,
            "duration_s": duration, "critical_steps": crit["critical_steps"],
            "serial_steps": crit["serial_steps"], "main_steps": crit["main_steps"],
            "n_subs": len(sub_wires), "n_parallel": crit["n_parallel_subs"],
            "n_sequential": crit["n_sequential_subs"], "claim_done": runlog.get("claim_done")}


def write_dump_index(dump_root: Path, records: list[dict]) -> None:
    rows = []
    for r in sorted(records, key=lambda x: (x["task"], x["run"])):
        rel = f"{r['task']}/{r['run']}/audit.html"
        pill = ("ok'>PASS" if r["pass"] is True else "bad'>FAIL") if r["pass"] is not None else "bad'>N/A"
        ck = f"{r['ckpt_pass']}/{r['ckpt_total']}" if r.get("ckpt_pass") is not None else "—"
        dur = f"{r['duration_s']}s" if r.get("duration_s") else "—"
        rows.append(f"<tr><td><a href='{rel}'>{r['task']}</a></td><td>{r['run']}</td>"
                    f"<td><span class='pill {pill}</span></td><td>{ck}</td><td>{dur}</td>"
                    f"<td>{r['critical_steps']}</td><td>{r['serial_steps']}</td>"
                    f"<td>{r['n_subs']} ({r['n_parallel']}p/{r['n_sequential']}s)</td></tr>")
    doc = f"""<!DOCTYPE html><html><head><meta charset='utf-8'>
<meta name='viewport' content='width=device-width,initial-scale=1'>
<title>Audit Index — {dump_root.name}</title>
<style>body{{background:#faf9f7;color:#1f1e1b;font:15px 'Söhne','Inter',sans-serif;padding:28px;max-width:1200px;margin:auto}}
table{{border-collapse:collapse;width:100%;background:#fff;border:1px solid #e8e6e1;border-radius:10px;overflow:hidden}}
td,th{{padding:9px 12px;border-bottom:1px solid #e8e6e1;text-align:left;font-size:13.5px}}
th{{background:#f4f2ee;color:#6b6862;font-size:12px;text-transform:uppercase}}
a{{color:#b4552d;text-decoration:none}}a:hover{{text-decoration:underline}}
.pill{{padding:3px 12px;border-radius:14px;font-size:12.5px;font-weight:600}}
.pill.ok{{background:#e6f3ec;color:#2a7d4f}}.pill.bad{{background:#fbeaea;color:#c13434}}
h1{{font-size:20px}}p{{color:#6b6862}}</style></head><body>
<h1>Audit Index — {dump_root.name}</h1><p>{len(records)} case(s)</p>
<table><tr><th>Task</th><th>Run</th><th>Pass</th><th>Checkpoints</th><th>Duration</th><th>Critical</th><th>Serial</th><th>Subs</th></tr>
{''.join(rows)}</table></body></html>"""
    (dump_root / "audit_index.html").write_text(doc)


def main(argv: list[str]) -> int:
    roots = [Path(a) for a in argv[1:]] or [Path("dumps/kimi-code")]
    for root in roots:
        if not root.is_dir():
            print(f"[skip] {root} not a directory")
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
