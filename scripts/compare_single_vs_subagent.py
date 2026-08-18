#!/usr/bin/env python3
"""Compare single-agent and auto-subagent kimi-code evaluation dumps.

The comparison is anchored on the single-agent result table, so the split
labels and selected reruns are exactly the ones used in the 64-case run.  The
auto-subagent batch is then matched by case name.
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import statistics
from collections import defaultdict
from pathlib import Path
from typing import Any


DELEGATION_TOOLS = {"Agent", "AgentSwarm"}
USAGE_FIELDS = ("inputOther", "inputCacheRead", "inputCacheCreation", "output")


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
    # Fallback for environments without tiktoken. It is intentionally simple
    # because the primary prompt/assistant counts come from recorded API usage.
    return max(1, math.ceil(len(text) / 4))


def as_bool(value: Any) -> bool | None:
    if value is True or value is False or value is None:
        return value
    s = str(value).strip().lower()
    if s == "true":
        return True
    if s == "false":
        return False
    return None


def int_or_none(value: Any) -> int | None:
    if value is None or value == "":
        return None
    try:
        return int(float(value))
    except Exception:
        return None


def resolve_audit_path(path_value: str, toolathlon_root: Path) -> Path:
    p = Path(path_value)
    if p.is_absolute():
        return p
    return toolathlon_root / p


def find_inner_dir(run_dir: Path) -> Path | None:
    if not run_dir.is_dir():
        return None
    for model_dir in run_dir.iterdir():
        if not model_dir.is_dir() or not model_dir.name.startswith("kimi-code"):
            continue
        for inner in model_dir.iterdir():
            if inner.is_dir() and inner.name.startswith("SingleUserTurn"):
                return inner
    return None


def usage_totals(usage: dict[str, Any] | None) -> dict[str, int]:
    usage = usage or {}
    return {k: int_or_none(usage.get(k)) or 0 for k in USAGE_FIELDS}


def add_usage(dst: dict[str, int], src: dict[str, int]) -> None:
    for k in USAGE_FIELDS:
        dst[k] = dst.get(k, 0) + src.get(k, 0)


def step_token_total(step: dict[str, Any]) -> int:
    u = step.get("usage_totals") or {}
    return int(u.get("inputOther", 0)) + int(u.get("inputCacheRead", 0)) + int(u.get("inputCacheCreation", 0)) + int(u.get("output", 0))


def parse_wire(wire_path: Path) -> dict[str, Any]:
    steps: dict[int, dict[str, Any]] = {}
    cur_step: dict[str, Any] | None = None
    fallback_usage_by_step: dict[int, list[dict[str, int]]] = defaultdict(list)
    profile_name = None

    if not wire_path.exists():
        return {"steps": [], "n_steps": 0, "profile": None}

    for line in wire_path.read_text(errors="replace").splitlines():
        try:
            rec = json.loads(line)
        except Exception:
            continue

        typ = rec.get("type")
        if typ == "profile.bind":
            profile_name = rec.get("profileName") or rec.get("profile")
            continue

        if typ == "usage.record":
            if cur_step is not None:
                sn = cur_step.get("step")
                if isinstance(sn, int):
                    fallback_usage_by_step[sn].append(usage_totals(rec.get("usage")))
            continue

        if typ != "context.append_loop_event":
            continue

        ev = rec.get("event") or {}
        ev_type = ev.get("type")
        sn = ev.get("step")

        if ev_type == "step.begin":
            if isinstance(sn, int):
                cur_step = {
                    "step": sn,
                    "calls": [],
                    "usage_totals": {k: 0 for k in USAGE_FIELDS},
                    "observation_tokens_est": 0,
                    "has_step_end_usage": False,
                }
                steps[sn] = cur_step
        elif ev_type == "tool.call":
            if cur_step is not None:
                cur_step["calls"].append(
                    {
                        "name": ev.get("name"),
                        "args": ev.get("args") or {},
                        "id": ev.get("toolCallId"),
                    }
                )
        elif ev_type == "tool.result":
            if cur_step is not None:
                res = ev.get("result")
                if isinstance(res, dict):
                    output = res.get("output")
                else:
                    output = res
                cur_step["observation_tokens_est"] += estimate_tokens(str(output or ""))
        elif ev_type == "step.end":
            if isinstance(sn, int):
                st = steps.setdefault(
                    sn,
                    {
                        "step": sn,
                        "calls": [],
                        "usage_totals": {k: 0 for k in USAGE_FIELDS},
                        "observation_tokens_est": 0,
                        "has_step_end_usage": False,
                    },
                )
                st["usage_totals"] = usage_totals(ev.get("usage"))
                st["has_step_end_usage"] = True

    for sn, st in steps.items():
        if st.get("has_step_end_usage"):
            continue
        total = {k: 0 for k in USAGE_FIELDS}
        for u in fallback_usage_by_step.get(sn, []):
            add_usage(total, u)
        st["usage_totals"] = total

    ordered = [steps[k] for k in sorted(steps)]
    totals = {k: 0 for k in USAGE_FIELDS}
    observation = 0
    for st in ordered:
        add_usage(totals, st.get("usage_totals") or {})
        observation += int(st.get("observation_tokens_est") or 0)

    return {
        "steps": ordered,
        "n_steps": len(ordered),
        "profile": profile_name,
        "usage_totals": totals,
        "observation_tokens_est": observation,
    }


def collect_wires(inner_dir: Path | None) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    empty = {"steps": [], "n_steps": 0, "usage_totals": {k: 0 for k in USAGE_FIELDS}, "observation_tokens_est": 0}
    if inner_dir is None:
        return empty, []

    session_agents = sorted((inner_dir / ".kimi_home" / "sessions").glob("*/*/agents"))
    if not session_agents:
        return empty, []

    agents_root = session_agents[0]
    main_wire = parse_wire(agents_root / "main" / "wire.jsonl")
    sub_wires: list[dict[str, Any]] = []
    for sub_dir in sorted(agents_root.iterdir()):
        if sub_dir.name.startswith("agent-") and (sub_dir / "wire.jsonl").exists():
            sub_wires.append(parse_wire(sub_dir / "wire.jsonl"))
    return main_wire, sub_wires


def wire_total_tokens(wire: dict[str, Any]) -> int:
    u = wire.get("usage_totals") or {}
    return int(u.get("inputOther", 0)) + int(u.get("inputCacheRead", 0)) + int(u.get("inputCacheCreation", 0)) + int(u.get("output", 0))


def compute_critical_tokens(main_wire: dict[str, Any], sub_wires: list[dict[str, Any]]) -> dict[str, int]:
    sub_cursor = 0
    critical_api = 0
    critical_observation = 0

    for st in main_wire.get("steps") or []:
        deleg = [c for c in st.get("calls", []) if c.get("name") in DELEGATION_TOOLS]
        main_cost = step_token_total(st)
        main_observation_cost = int(st.get("observation_tokens_est") or 0)
        if not deleg:
            critical_api += main_cost
            critical_observation += main_observation_cost
            continue

        n_sub = 0
        for call in deleg:
            if call.get("name") == "AgentSwarm":
                items = (call.get("args") or {}).get("items") or (call.get("args") or {}).get("tasks") or []
                n_sub += len(items) if isinstance(items, list) and items else 1
            else:
                n_sub += 1
        taken = sub_wires[sub_cursor : sub_cursor + n_sub]
        sub_cursor += len(taken)
        critical_api += main_cost + (max([wire_total_tokens(w) for w in taken]) if taken else 0)
        critical_observation += main_observation_cost + (
            max([int(w.get("observation_tokens_est") or 0) for w in taken]) if taken else 0
        )

    serial = wire_total_tokens(main_wire) + sum(wire_total_tokens(w) for w in sub_wires)
    serial_observation = int(main_wire.get("observation_tokens_est") or 0) + sum(
        int(w.get("observation_tokens_est") or 0) for w in sub_wires
    )
    return {
        "critical_tokens_api": critical_api,
        "serial_tokens_api": serial,
        "critical_token_saving_api": max(0, serial - critical_api),
        "critical_observation_tokens_est": critical_observation,
        "serial_observation_tokens_est": serial_observation,
        "critical_all_tokens_est": critical_api + critical_observation,
        "serial_all_tokens_est": serial + serial_observation,
    }


def parse_token_metrics(run_dir: Path) -> dict[str, int]:
    inner = find_inner_dir(run_dir)
    main_wire, sub_wires = collect_wires(inner)

    total_usage = {k: 0 for k in USAGE_FIELDS}
    observation = int(main_wire.get("observation_tokens_est") or 0)
    add_usage(total_usage, main_wire.get("usage_totals") or {})
    for w in sub_wires:
        add_usage(total_usage, w.get("usage_totals") or {})
        observation += int(w.get("observation_tokens_est") or 0)

    prompt_tokens = total_usage["inputOther"] + total_usage["inputCacheRead"] + total_usage["inputCacheCreation"]
    assistant_tokens = total_usage["output"]
    out = {
        "prompt_tokens_api": prompt_tokens,
        "prompt_uncached_tokens_api": total_usage["inputOther"],
        "prompt_cache_read_tokens_api": total_usage["inputCacheRead"],
        "prompt_cache_creation_tokens_api": total_usage["inputCacheCreation"],
        "assistant_tokens_api": assistant_tokens,
        "total_tokens_api": prompt_tokens + assistant_tokens,
        "observation_tokens_est": observation,
    }
    out.update(compute_critical_tokens(main_wire, sub_wires))
    return out


def load_single_rows(single_root: Path) -> list[dict[str, Any]]:
    table = single_root / "single_agent_results_table.csv"
    rows: list[dict[str, Any]] = []
    with table.open(newline="") as f:
        for row in csv.DictReader(f):
            rows.append(row)
    return rows


def load_subagent_index(subagent_root: Path) -> dict[str, dict[str, Any]]:
    data = json.loads((subagent_root / "audit_index.json").read_text(errors="replace"))
    return {r["task"]: r for r in data}


def mode_record_from_single(row: dict[str, Any], toolathlon_root: Path) -> dict[str, Any]:
    audit = resolve_audit_path(row["audit_html"], toolathlon_root)
    run_dir = audit.parent
    duration_s = int_or_none(row.get("duration_s"))
    if duration_s is None:
        duration_s = duration_from_summary_csv(Path(row.get("summary") or ""), row.get("case") or "", run_dir)
    rec = {
        "pass": as_bool(row.get("pass")),
        "duration_s": duration_s or 0,
        "critical_steps": int_or_none(row.get("critical_steps")) or 0,
        "serial_steps": None,
        "main_steps": None,
        "subagents": int_or_none(row.get("subagents")) or 0,
        "parallel_subagents": int_or_none(row.get("parallel_subagents")) or 0,
        "sequential_subagents": int_or_none(row.get("sequential_subagents")) or 0,
        "audit_html": str(audit),
        "run_dir": str(run_dir),
        "run": row.get("run") or run_dir.name,
        "status": row.get("status") or "",
    }
    rec.update(parse_token_metrics(run_dir))
    return rec


def duration_from_summary_csv(summary_path: Path, case: str, run_dir: Path) -> int | None:
    if not summary_path.exists():
        return None
    try:
        with summary_path.open(newline="") as f:
            for row in csv.DictReader(f):
                if row.get("task") == case and Path(row.get("output_dir") or "") == run_dir:
                    return int_or_none(row.get("duration_s"))
    except Exception:
        return None
    return None


def mode_record_from_subagent(row: dict[str, Any], toolathlon_root: Path) -> dict[str, Any]:
    audit = resolve_audit_path(row["path"], toolathlon_root)
    run_dir = audit.parent
    rec = {
        "pass": as_bool(row.get("pass")),
        "duration_s": int_or_none(row.get("duration_s")) or 0,
        "critical_steps": int_or_none(row.get("critical_steps")) or 0,
        "serial_steps": int_or_none(row.get("serial_steps")),
        "main_steps": int_or_none(row.get("main_steps")),
        "subagents": int_or_none(row.get("n_subs")) or 0,
        "parallel_subagents": int_or_none(row.get("n_parallel")) or 0,
        "sequential_subagents": int_or_none(row.get("n_sequential")) or 0,
        "audit_html": str(audit),
        "run_dir": str(run_dir),
        "run": row.get("run") or run_dir.name,
        "status": "",
    }
    rec.update(parse_token_metrics(run_dir))
    return rec


def ratio(numer: float | int | None, denom: float | int | None) -> float | None:
    if numer is None or denom in (None, 0):
        return None
    return float(numer) / float(denom)


def fmt_num(v: Any, digits: int = 1) -> str:
    if v is None:
        return ""
    if isinstance(v, float):
        return f"{v:.{digits}f}"
    return str(v)


def median(values: list[int]) -> float | None:
    return statistics.median(values) if values else None


def mean(values: list[int]) -> float | None:
    return statistics.mean(values) if values else None


def percentile(values: list[int], pct: float) -> float | None:
    if not values:
        return None
    ordered = sorted(values)
    if len(ordered) == 1:
        return float(ordered[0])
    pos = (len(ordered) - 1) * pct
    lo = math.floor(pos)
    hi = math.ceil(pos)
    if lo == hi:
        return float(ordered[lo])
    return ordered[lo] + (ordered[hi] - ordered[lo]) * (pos - lo)


def aggregate(split: str, rows: list[dict[str, Any]], mode: str) -> dict[str, Any]:
    prefix = f"{mode}_"
    vals = [r for r in rows if r["split"] == split]

    def col(name: str) -> list[int]:
        return [int(r.get(prefix + name) or 0) for r in vals]

    passes = [as_bool(r.get(prefix + "pass")) for r in vals]
    n = len(vals)
    pass_true = sum(p is True for p in passes)
    pass_false = sum(p is False for p in passes)
    pass_none = sum(p is None for p in passes)
    durations = col("duration_s")
    total_tokens = col("total_tokens_api")
    crit_tokens = col("critical_tokens_api")
    crit_all_tokens = col("critical_all_tokens_est")
    crit_observation_tokens = col("critical_observation_tokens_est")
    crit_steps = col("critical_steps")
    return {
        f"{mode}_n": n,
        f"{mode}_pass": pass_true,
        f"{mode}_fail": pass_false,
        f"{mode}_pass_null": pass_none,
        f"{mode}_pass_rate": (pass_true / n) if n else None,
        f"{mode}_duration_sum_s": sum(durations),
        f"{mode}_duration_mean_s": mean(durations),
        f"{mode}_duration_median_s": median(durations),
        f"{mode}_duration_p95_s": percentile(durations, 0.95),
        f"{mode}_duration_max_s": max(durations) if durations else None,
        f"{mode}_critical_steps_mean": mean(crit_steps),
        f"{mode}_critical_steps_sum": sum(crit_steps),
        f"{mode}_subagents_sum": sum(col("subagents")),
        f"{mode}_subagent_used_cases": sum(1 for r in vals if int(r.get(prefix + "subagents") or 0) > 0),
        f"{mode}_parallel_subagents_sum": sum(col("parallel_subagents")),
        f"{mode}_sequential_subagents_sum": sum(col("sequential_subagents")),
        f"{mode}_prompt_tokens_sum": sum(col("prompt_tokens_api")),
        f"{mode}_assistant_tokens_sum": sum(col("assistant_tokens_api")),
        f"{mode}_observation_tokens_est_sum": sum(col("observation_tokens_est")),
        f"{mode}_total_tokens_api_sum": sum(total_tokens),
        f"{mode}_total_tokens_api_mean": mean(total_tokens),
        f"{mode}_critical_tokens_api_sum": sum(crit_tokens),
        f"{mode}_critical_tokens_api_mean": mean(crit_tokens),
        f"{mode}_critical_observation_tokens_est_sum": sum(crit_observation_tokens),
        f"{mode}_critical_all_tokens_est_sum": sum(crit_all_tokens),
        f"{mode}_critical_all_tokens_est_mean": mean(crit_all_tokens),
    }


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    if not rows:
        return
    fieldnames = list(rows[0].keys())
    with path.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def write_markdown_table(path: Path, rows: list[dict[str, Any]], columns: list[str]) -> None:
    lines = []
    lines.append("| " + " | ".join(columns) + " |")
    lines.append("|" + "|".join("---" for _ in columns) + "|")
    for row in rows:
        lines.append("| " + " | ".join(format_cell(c, row.get(c)) for c in columns) + " |")
    path.write_text("\n".join(lines) + "\n")


def format_cell(column: str, value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if "rate" in column or "ratio" in column:
            return f"{value:.3f}"
        return f"{value:.1f}"
    return str(value)


def build_case_rows(single_root: Path, subagent_root: Path, toolathlon_root: Path) -> list[dict[str, Any]]:
    single_rows = load_single_rows(single_root)
    sub_index = load_subagent_index(subagent_root)
    case_rows: list[dict[str, Any]] = []

    for row in single_rows:
        case = row["case"]
        if case not in sub_index:
            raise RuntimeError(f"case {case!r} is missing from subagent audit_index.json")

        single = mode_record_from_single(row, toolathlon_root)
        subagent = mode_record_from_subagent(sub_index[case], toolathlon_root)
        out: dict[str, Any] = {"split": row["split"], "case": case}
        for mode, rec in (("single", single), ("subagent", subagent)):
            for k, v in rec.items():
                out[f"{mode}_{k}"] = v

        out["pass_delta_subagent_minus_single"] = (1 if subagent["pass"] is True else 0) - (1 if single["pass"] is True else 0)
        out["duration_delta_subagent_minus_single_s"] = subagent["duration_s"] - single["duration_s"]
        out["duration_ratio_single_over_subagent"] = ratio(single["duration_s"], subagent["duration_s"])
        out["critical_steps_delta_subagent_minus_single"] = subagent["critical_steps"] - single["critical_steps"]
        out["critical_tokens_delta_subagent_minus_single"] = subagent["critical_tokens_api"] - single["critical_tokens_api"]
        out["critical_tokens_ratio_single_over_subagent"] = ratio(single["critical_tokens_api"], subagent["critical_tokens_api"])
        out["total_tokens_delta_subagent_minus_single"] = subagent["total_tokens_api"] - single["total_tokens_api"]
        out["total_tokens_ratio_single_over_subagent"] = ratio(single["total_tokens_api"], subagent["total_tokens_api"])
        case_rows.append(out)

    return case_rows


def build_split_rows(case_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    split_rows: list[dict[str, Any]] = []
    for split in sorted({r["split"] for r in case_rows}):
        base = {"split": split, "cases": sum(r["split"] == split for r in case_rows)}
        single = aggregate(split, case_rows, "single")
        subagent = aggregate(split, case_rows, "subagent")
        row: dict[str, Any] = {**base, **single, **subagent}
        row["pass_rate_delta_subagent_minus_single"] = (
            row["subagent_pass_rate"] - row["single_pass_rate"]
            if row["subagent_pass_rate"] is not None and row["single_pass_rate"] is not None
            else None
        )
        row["duration_mean_ratio_single_over_subagent"] = ratio(row["single_duration_mean_s"], row["subagent_duration_mean_s"])
        row["duration_sum_ratio_single_over_subagent"] = ratio(row["single_duration_sum_s"], row["subagent_duration_sum_s"])
        row["total_tokens_sum_ratio_single_over_subagent"] = ratio(row["single_total_tokens_api_sum"], row["subagent_total_tokens_api_sum"])
        row["critical_tokens_sum_ratio_single_over_subagent"] = ratio(row["single_critical_tokens_api_sum"], row["subagent_critical_tokens_api_sum"])
        row["critical_steps_sum_ratio_single_over_subagent"] = ratio(row["single_critical_steps_sum"], row["subagent_critical_steps_sum"])
        split_rows.append(row)

    total = {"split": "TOTAL", "cases": len(case_rows)}
    total_single = aggregate("TOTAL", [{**r, "split": "TOTAL"} for r in case_rows], "single")
    total_subagent = aggregate("TOTAL", [{**r, "split": "TOTAL"} for r in case_rows], "subagent")
    total.update(total_single)
    total.update(total_subagent)
    total["pass_rate_delta_subagent_minus_single"] = total["subagent_pass_rate"] - total["single_pass_rate"]
    total["duration_mean_ratio_single_over_subagent"] = ratio(total["single_duration_mean_s"], total["subagent_duration_mean_s"])
    total["duration_sum_ratio_single_over_subagent"] = ratio(total["single_duration_sum_s"], total["subagent_duration_sum_s"])
    total["total_tokens_sum_ratio_single_over_subagent"] = ratio(total["single_total_tokens_api_sum"], total["subagent_total_tokens_api_sum"])
    total["critical_tokens_sum_ratio_single_over_subagent"] = ratio(total["single_critical_tokens_api_sum"], total["subagent_critical_tokens_api_sum"])
    total["critical_steps_sum_ratio_single_over_subagent"] = ratio(total["single_critical_steps_sum"], total["subagent_critical_steps_sum"])
    split_rows.append(total)
    return split_rows


def write_report(out_dir: Path, split_rows: list[dict[str, Any]], case_rows: list[dict[str, Any]], single_root: Path, subagent_root: Path) -> None:
    lines = [
        "# DeepSeek v4 Flash: single-agent vs auto-subagent",
        "",
        f"- Single-agent root: `{single_root}`",
        f"- Auto-subagent root: `{subagent_root}`",
        f"- Case set: {len(case_rows)} exact-name matched cases, using the single-agent split labels.",
        "- Prompt/assistant tokens are API-recorded `usage.record` values: prompt = `inputOther + inputCacheRead + inputCacheCreation`, assistant = `output`.",
        "- Observation tokens are estimated from `tool.result` output text with `cl100k_base`; they are reported separately because the API usage does not expose observation-only counts.",
        "- `critical_tokens_api` mirrors Critical Steps over API-recorded prompt+assistant tokens: main non-delegation step tokens are summed; for each delegation phase, cost = main delegation step tokens + the slowest delegated sub-agent token total.",
        "- `critical_all_tokens_est` additionally includes estimated observation tokens in the same critical-path fashion.",
        "",
        "## Split Summary",
        "",
    ]
    summary_cols = [
        "split",
        "cases",
        "single_pass",
        "subagent_pass",
        "single_pass_rate",
        "subagent_pass_rate",
        "pass_rate_delta_subagent_minus_single",
        "single_duration_mean_s",
        "subagent_duration_mean_s",
        "duration_mean_ratio_single_over_subagent",
        "single_critical_steps_sum",
        "subagent_critical_steps_sum",
        "critical_steps_sum_ratio_single_over_subagent",
        "single_subagents_sum",
        "subagent_subagent_used_cases",
        "subagent_subagents_sum",
        "subagent_parallel_subagents_sum",
        "subagent_sequential_subagents_sum",
    ]
    lines.extend(markdown_lines(split_rows, summary_cols))
    lines.extend(["", "## Token Summary", ""])
    token_cols = [
        "split",
        "single_prompt_tokens_sum",
        "subagent_prompt_tokens_sum",
        "single_assistant_tokens_sum",
        "subagent_assistant_tokens_sum",
        "single_observation_tokens_est_sum",
        "subagent_observation_tokens_est_sum",
        "single_total_tokens_api_sum",
        "subagent_total_tokens_api_sum",
        "total_tokens_sum_ratio_single_over_subagent",
        "single_critical_tokens_api_sum",
        "subagent_critical_tokens_api_sum",
        "critical_tokens_sum_ratio_single_over_subagent",
        "single_critical_all_tokens_est_sum",
        "subagent_critical_all_tokens_est_sum",
    ]
    lines.extend(markdown_lines(split_rows, token_cols))

    flip_rows = [
        {
            "split": r["split"],
            "case": r["case"],
            "single_pass": r["single_pass"],
            "subagent_pass": r["subagent_pass"],
            "single_duration_s": r["single_duration_s"],
            "subagent_duration_s": r["subagent_duration_s"],
            "single_critical_steps": r["single_critical_steps"],
            "subagent_critical_steps": r["subagent_critical_steps"],
            "subagent_subagents": r["subagent_subagents"],
        }
        for r in case_rows
        if as_bool(r["single_pass"]) != as_bool(r["subagent_pass"])
    ]
    lines.extend(["", "## Pass Flips", ""])
    if flip_rows:
        lines.extend(
            markdown_lines(
                flip_rows,
                [
                    "split",
                    "case",
                    "single_pass",
                    "subagent_pass",
                    "single_duration_s",
                    "subagent_duration_s",
                    "single_critical_steps",
                    "subagent_critical_steps",
                    "subagent_subagents",
                ],
            )
        )
    else:
        lines.append("No pass flips.")

    lines.extend(
        [
            "",
            "## Artifacts",
            "",
            "- `split_comparison_summary.csv` / `.md`: split-level aggregate table.",
            "- `case_comparison_detail.csv` / `.md`: per-case comparison table.",
            "- `README.md`: this report.",
            "",
        ]
    )
    (out_dir / "README.md").write_text("\n".join(lines))


def markdown_lines(rows: list[dict[str, Any]], columns: list[str]) -> list[str]:
    out = ["| " + " | ".join(columns) + " |", "|" + "|".join("---" for _ in columns) + "|"]
    for row in rows:
        cells = []
        for c in columns:
            v = row.get(c)
            if isinstance(v, float):
                if "rate" in c or "ratio" in c:
                    cells.append(f"{v:.3f}")
                else:
                    cells.append(f"{v:.1f}")
            else:
                cells.append("" if v is None else str(v))
        out.append("| " + " | ".join(cells) + " |")
    return out


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--single-root", required=True, type=Path)
    parser.add_argument("--subagent-root", required=True, type=Path)
    parser.add_argument("--toolathlon-root", default=Path(__file__).resolve().parents[1], type=Path)
    parser.add_argument("--out-dir", required=True, type=Path)
    args = parser.parse_args()

    args.out_dir.mkdir(parents=True, exist_ok=True)
    case_rows = build_case_rows(args.single_root, args.subagent_root, args.toolathlon_root)
    split_rows = build_split_rows(case_rows)

    write_csv(args.out_dir / "case_comparison_detail.csv", case_rows)
    write_csv(args.out_dir / "split_comparison_summary.csv", split_rows)

    split_md_cols = [
        "split",
        "cases",
        "single_pass",
        "subagent_pass",
        "single_pass_rate",
        "subagent_pass_rate",
        "single_duration_mean_s",
        "subagent_duration_mean_s",
        "duration_mean_ratio_single_over_subagent",
        "single_total_tokens_api_sum",
        "subagent_total_tokens_api_sum",
        "total_tokens_sum_ratio_single_over_subagent",
        "single_critical_tokens_api_sum",
        "subagent_critical_tokens_api_sum",
        "critical_tokens_sum_ratio_single_over_subagent",
        "single_critical_all_tokens_est_sum",
        "subagent_critical_all_tokens_est_sum",
    ]
    write_markdown_table(args.out_dir / "split_comparison_summary.md", split_rows, split_md_cols)

    case_md_cols = [
        "split",
        "case",
        "single_pass",
        "subagent_pass",
        "single_duration_s",
        "subagent_duration_s",
        "duration_ratio_single_over_subagent",
        "single_critical_steps",
        "subagent_critical_steps",
        "single_subagents",
        "subagent_subagents",
        "subagent_parallel_subagents",
        "subagent_sequential_subagents",
        "single_total_tokens_api",
        "subagent_total_tokens_api",
        "total_tokens_ratio_single_over_subagent",
        "single_critical_tokens_api",
        "subagent_critical_tokens_api",
        "critical_tokens_ratio_single_over_subagent",
        "single_critical_all_tokens_est",
        "subagent_critical_all_tokens_est",
    ]
    write_markdown_table(args.out_dir / "case_comparison_detail.md", case_rows, case_md_cols)
    write_report(args.out_dir, split_rows, case_rows, args.single_root, args.subagent_root)

    print(f"[ok] wrote {args.out_dir}")
    print(f"[ok] cases={len(case_rows)} splits={len(split_rows) - 1}+TOTAL")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
