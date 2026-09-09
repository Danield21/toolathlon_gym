#!/usr/bin/env python3
"""Keep latest slot per task in kimi-code_MiniMax-M3-three-apifail and rebuild one summary."""
from __future__ import annotations

import csv
import re
import shutil
from pathlib import Path

DUMP = Path(
    "/lintaoLab2/bowending/project_agent_swarm_benchmark/toolathlon_gym/dumps/kimi-code_MiniMax-M3-three-apifail"
)
STAMP_RE = re.compile(r"(\d{8}-\d{6})")
SKIP_DIRS = {
    "no_subagent",
    "need_subagent",
    "_unlisted",
    "_incoming",
}
FIELDS = ["task", "status", "exit_code", "output_dir", "pg_port", "duration_s"]


def slot_stamp(name: str) -> str:
    m = STAMP_RE.search(name)
    return m.group(1) if m else ""


def is_slot(p: Path) -> bool:
    return p.is_dir() and "_slot" in p.name


def slot_has_payload(p: Path) -> bool:
    if (p / "audit.html").is_file() or (p / "eval_res.json").is_file():
        return True
    if (p / "run.log").is_file() and (p / "run.log").stat().st_size > 0:
        return True
    for hit in p.rglob("traj.json"):
        if hit.is_file() and hit.stat().st_size > 0:
            return True
        break
    return False


def pick_latest(slots: list[Path]) -> Path:
    real = [p for p in slots if slot_has_payload(p)]
    cand = real or slots
    cand.sort(key=lambda p: (slot_stamp(p.name), p.stat().st_mtime, p.name))
    return cand[-1]


def load_summary_rows() -> dict[str, dict[str, str]]:
    by_outdir: dict[str, dict[str, str]] = {}
    paths = []
    latest = DUMP / "summary_latest.csv"
    if latest.exists():
        paths.append(latest)
    paths.extend(sorted(DUMP.glob("summary_parallel_*.csv")))
    for sp in paths:
        with sp.open(newline="") as f:
            for row in csv.DictReader(f):
                out = (row.get("output_dir") or "").rstrip("/")
                if out:
                    by_outdir[out] = row
    return by_outdir


def infer_row(case: Path, slot: Path, by_outdir: dict[str, dict[str, str]]) -> dict[str, str]:
    row = by_outdir.get(str(slot))
    if row:
        return {k: row.get(k, "") for k in FIELDS}
    status = "unknown"
    if (slot / "eval_res.json").is_file():
        status = "success"
    elif list(slot.rglob(".provider_invalid")):
        status = "provider_invalid"
    elif (slot / "run.log").is_file():
        status = "case_failed"
    return {
        "task": case.name,
        "status": status,
        "exit_code": "",
        "output_dir": str(slot),
        "pg_port": "",
        "duration_s": "",
    }


def main() -> None:
    if not DUMP.is_dir():
        raise SystemExit(f"missing dump {DUMP}")
    by_outdir = load_summary_rows()
    removed = 0
    kept_rows = []
    multi = []
    for case in sorted(p for p in DUMP.iterdir() if p.is_dir() and p.name not in SKIP_DIRS and not p.name.startswith(".")):
        slots = [p for p in case.iterdir() if is_slot(p)]
        if not slots:
            continue
        latest = pick_latest(slots)
        old = [p for p in slots if p != latest]
        if old:
            multi.append((case.name, latest.name, [p.name for p in old]))
        for p in old:
            shutil.rmtree(p, ignore_errors=True)
            removed += 1
            print(f"[clean] rm {case.name}/{p.name}")
        kept_rows.append(infer_row(case, latest, by_outdir))

    latest_path = DUMP / "summary_latest.csv"
    with latest_path.open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS, extrasaction="ignore")
        w.writeheader()
        w.writerows(kept_rows)
    print(f"[clean] wrote {latest_path} rows={len(kept_rows)}")

    for sp in DUMP.glob("summary_parallel_*.csv"):
        print(f"[clean] rm summary {sp.name}")
        sp.unlink(missing_ok=True)

    # also drop leftover per-wave notes that are not the merged latest
    counts: dict[str, int] = {}
    for row in kept_rows:
        counts[row["status"]] = counts.get(row["status"], 0) + 1
    print(f"[clean] removed_old_slots={removed} tasks={len(kept_rows)} status={counts}")
    print(f"[clean] multi_slot_tasks={len(multi)}")
    for name, keep, old in multi:
        print(f"  keep {name}/{keep}  drop {old}")


if __name__ == "__main__":
    main()
