#!/usr/bin/env python3
"""Scan dump trees for mock-port overflow / MCP Invalid URL contamination.

Two independent detectors per case slot:
  A) mcp.json static check: any env value that is a URL with port > 65535
     (Node `new URL()` rejects these -> "Error: Invalid URL" on every call
     of that MCP server's tools).
  B) wire.jsonl dynamic check: count of 'Invalid URL' tool results in the
     main agent wire (proxy for a broken MCP server during the run).

Usage: scan_mcp_invalid_url.py <dump_root> [--out <tsv>]
Emits TSV rows: experiment, bucket, case, slot, run_id,
mcp_bad_ports(server:port,...) , invalid_url_hits, case_status.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import defaultdict

URL_PORT_RE = re.compile(r"^https?://[^\s:/]+:(\d+)(?:/|$)")
SLOT_RE = re.compile(
    r"^(?:[A-Za-z0-9][A-Za-z0-9_.-]*[-_])?(?:\d{8}-\d{6}|\d{8})(?:[-_][A-Za-z0-9_.-]+)?_slot\d+$"
)
DONE_RE = re.compile(r"DONE\s+\S+ -> ([a-z_]+)")


def iter_case_slots(root: str):
    """Yield (bucket, case, slot_path) walking the hierarchical dump."""
    for dirpath, dirnames, _ in os.walk(root):
        base = os.path.basename(dirpath)
        if SLOT_RE.match(base):
            case = os.path.basename(os.path.dirname(dirpath))
            # bucket = path fragment between dump root's mode dir and case
            rel = os.path.relpath(os.path.dirname(dirpath), root)
            bucket = os.path.dirname(rel)
            yield bucket, case, dirpath


def find_mcp_json(slot: str):
    for dirpath, dirnames, filenames in os.walk(slot):
        if "mcp.json" in filenames and ".kimi_home" in dirpath:
            return os.path.join(dirpath, "mcp.json")
    return None


def check_mcp_json(path: str):
    """Return list of 'server:port' entries whose URL port exceeds 65535."""
    bad = []
    try:
        data = json.load(open(path))
    except Exception:
        return bad
    servers = data.get("mcpServers", data)
    if not isinstance(servers, dict):
        return bad
    for name, cfg in servers.items():
        if not isinstance(cfg, dict):
            continue
        vals = []
        env = cfg.get("env")
        if isinstance(env, dict):
            vals.extend(str(v) for v in env.values())
        args = cfg.get("args")
        if isinstance(args, list):
            vals.extend(str(a) for a in args)
        for key in ("url", "serverUrl", "baseUrl", "base_url"):
            if cfg.get(key):
                vals.append(str(cfg[key]))
        for v in vals:
            m = URL_PORT_RE.match(v.strip())
            if m and int(m.group(1)) > 65535:
                bad.append(f"{name}:{m.group(1)}")
                break
    return bad


def count_invalid_url(slot: str):
    """Count 'Invalid URL' occurrences in the main agent wire.jsonl."""
    total = 0
    for dirpath, _, filenames in os.walk(slot):
        if "wire.jsonl" not in filenames:
            continue
        if os.path.basename(dirpath) != "main":
            continue  # main agent only; subagent echoes would double-count
        fp = os.path.join(dirpath, "wire.jsonl")
        try:
            with open(fp, errors="ignore") as f:
                total += f.read().count("Invalid URL")
        except Exception:
            pass
    return total


def case_status(slot: str):
    try:
        text = open(os.path.join(slot, "run.log"), errors="ignore").read()
    except Exception:
        return "no_runlog"
    m = DONE_RE.findall(text)
    return m[-1] if m else "incomplete"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("root")
    ap.add_argument("--out")
    args = ap.parse_args()
    root = args.root.rstrip("/")
    experiment = os.path.basename(root)

    rows = []
    for bucket, case, slot in iter_case_slots(root):
        mcp = find_mcp_json(slot)
        bad_ports = check_mcp_json(mcp) if mcp else []
        hits = count_invalid_url(slot)
        if not bad_ports and hits == 0:
            continue  # clean slot: skip to keep output focused
        run_id = os.path.basename(slot)
        status = case_status(slot)
        rows.append(
            (experiment, bucket, case, run_id, ",".join(bad_ports), hits, status)
        )

    out = sys.stdout
    fh = open(args.out, "w") if args.out else None
    if fh:
        out = fh
    print("experiment\tbucket\tcase\tslot\tbad_mcp_ports\tinvalid_url_hits\tstatus", file=out)
    for r in rows:
        print("\t".join(str(x) for x in r), file=out)
    if fh:
        fh.close()
        # summary to stdout
        by_case = defaultdict(lambda: {"bad_ports": set(), "hits": 0, "status": ""})
        for _, bucket, case, _, bp, hits, status in rows:
            e = by_case[(bucket, case)]
            if bp:
                e["bad_ports"].update(bp.split(","))
            e["hits"] += int(hits)
            e["status"] = status
        affected = {k: v for k, v in by_case.items() if v["bad_ports"] or v["hits"] > 0}
        print(f"[summary] experiment={experiment} affected_cases={len(affected)}")
        for (bucket, case), v in sorted(affected.items()):
            print(f"  {bucket or '.'}\t{case}\tports={sorted(v['bad_ports'])}\thits={v['hits']}\tstatus={v['status']}")


if __name__ == "__main__":
    main()
