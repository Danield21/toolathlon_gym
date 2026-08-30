#!/usr/bin/env python3
"""Generate caidian v2 evidence, metrics, CSV, and Markdown reports."""

from __future__ import annotations

import argparse
import csv
import json
import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from caidian_v2 import (
    VERSION,
    TrajectoryIntegrityError,
    aggregate_scores,
    apply_verified_overrides,
    extract_actual_ir,
    extract_runtime_prompt,
    load_source_map,
    parse_planned_ir,
    resolve_run_artifacts,
    score_task,
    sha256_file,
    validate_override_catalog,
)


def write_checksums(root: Path) -> None:
    checksum_path = root / "SHA256SUMS"
    files = sorted(
        path for path in root.rglob("*")
        if path.is_file() and path != checksum_path and not path.name.startswith("._")
    )
    checksum_path.write_text(
        "\n".join(f"{sha256_file(path)}  {path.relative_to(root)}" for path in files) + "\n",
        encoding="utf-8",
    )


def percent(value: float | None) -> str:
    return "N/A" if value is None else f"{100.0 * value:.1f}%"


def read_json(path: Path | None, default: Any) -> Any:
    if path is None:
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def legacy_summary(path: Path | None) -> dict[str, Any] | None:
    if path is None:
        return None
    payload = read_json(path, {})
    oracle = ((payload.get("modes") or {}).get("oracle") or {}) if isinstance(payload, dict) else {}
    summary = oracle.get("summary_agg")
    if not isinstance(summary, dict):
        raise ValueError(f"Legacy data has no modes.oracle.summary_agg: {path}")
    return {
        "source": str(path),
        "source_sha256": sha256_file(path),
        "wide_recall": summary.get("wr"),
        "wide_precision": summary.get("wp"),
        "strict_recall": summary.get("sr"),
        "strict_precision": summary.get("sp"),
        "coverage": summary.get("cover"),
    }


def write_task_csv(path: Path, records: list[dict[str, Any]]) -> None:
    columns = [
        "task",
        "trajectory_valid",
        "trajectory_error",
        "source_label",
        "evaluator_pass",
        "prompt_provenance",
        "prompt_matches_reference",
        "planned_waves_eligible",
        "planned_waves_ineligible",
        "planned_wave_hits",
        "planned_wave_recall",
        "fresh_waves",
        "fresh_wave_hits",
        "fresh_wave_precision",
        "fine_planned_nodes",
        "fine_planned_hits",
        "fine_node_recall",
        "fine_actual_nodes",
        "fine_actual_hits",
        "fine_node_precision",
        "accepted_resumes",
        "orphan_resumes",
        "rejected_calls",
        "repair_fresh_waves",
        "overhead_rate",
        "raw_stream",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns)
        writer.writeheader()
        for record in records:
            if record.get("score") is None:
                writer.writerow(
                    {
                        "task": record["task"],
                        "trajectory_valid": False,
                        "trajectory_error": record.get("trajectory_error", ""),
                        "source_label": record["source"].get("source_label", ""),
                        "evaluator_pass": record["source"].get("evaluator_pass", ""),
                        "prompt_provenance": record["prompt"]["provenance"],
                        "prompt_matches_reference": record["prompt"]["matches_reference"],
                        "raw_stream": record["source"]["raw_stream"],
                    }
                )
                continue
            score = record["score"]
            fine = score["fine"]
            overhead = score["overhead"]
            writer.writerow(
                {
                    "task": record["task"],
                    "trajectory_valid": True,
                    "trajectory_error": "",
                    "source_label": record["source"].get("source_label", ""),
                    "evaluator_pass": record["source"].get("evaluator_pass", ""),
                    "prompt_provenance": record["prompt"]["provenance"],
                    "prompt_matches_reference": record["prompt"]["matches_reference"],
                    "planned_waves_eligible": score["planned_wave_total"],
                    "planned_waves_ineligible": score["ineligible_wave_total"],
                    "planned_wave_hits": score["planned_wave_hit"],
                    "planned_wave_recall": percent(score["planned_wave_recall"]),
                    "fresh_waves": score["fresh_wave_total"],
                    "fresh_wave_hits": score["fresh_wave_hit"],
                    "fresh_wave_precision": percent(score["fresh_wave_precision"]),
                    "fine_planned_nodes": fine["planned_total"],
                    "fine_planned_hits": fine["planned_hit"],
                    "fine_node_recall": percent(fine["recall"]),
                    "fine_actual_nodes": fine["actual_total"],
                    "fine_actual_hits": fine["actual_hit"],
                    "fine_node_precision": percent(fine["precision"]),
                    "accepted_resumes": overhead["accepted_resumes"],
                    "orphan_resumes": overhead["orphan_resumes"],
                    "rejected_calls": overhead["rejected_calls"],
                    "repair_fresh_waves": overhead["repair_fresh_waves"],
                    "overhead_rate": percent(overhead["overhead_rate"]),
                    "raw_stream": record["source"]["raw_stream"],
                }
            )


def write_source_manifest(path: Path, records: list[dict[str, Any]]) -> None:
    columns = [
        "task",
        "source_label",
        "target_run",
        "raw_stream",
        "raw_stream_bytes",
        "raw_stream_sha256",
        "prompt_provenance",
        "prompt_source_path",
        "runtime_prompt_sha256",
        "reference_prompt_sha256",
        "prompt_matches_reference",
    ]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns, delimiter="\t")
        writer.writeheader()
        for record in records:
            source, prompt = record["source"], record["prompt"]
            writer.writerow(
                {
                    "task": record["task"],
                    "source_label": source.get("source_label", ""),
                    "target_run": source.get("target_run", ""),
                    "raw_stream": source["raw_stream"],
                    "raw_stream_bytes": source["raw_stream_bytes"],
                    "raw_stream_sha256": source["raw_stream_sha256"],
                    "prompt_provenance": prompt["provenance"],
                    "prompt_source_path": prompt["source_path"],
                    "runtime_prompt_sha256": prompt["sha256"],
                    "reference_prompt_sha256": prompt["reference_sha256"],
                    "prompt_matches_reference": prompt["matches_reference"],
                }
            )


def render_report(
    output: Path,
    records: list[dict[str, Any]],
    summary: dict[str, Any],
    legacy: dict[str, Any] | None,
    args: argparse.Namespace,
) -> None:
    lines = [
        "# Oracle 踩点率 v2 报告",
        "",
        f"> 算法版本: `{VERSION}`",
        f"> GT 参考目录: `{args.gt}`",
        f"> 精确轨迹源清单: `{args.source_map}`",
        f"> 轨迹根: `{args.dump}`",
        "> 新版不覆盖 legacy 产物；实际 prompt 优先取 `traj_log.json:config.task_str`，否则取 `run.log` 启动参数。",
        "",
        "## 1. 四组分离指标（P4）",
        "",
        "| 指标 | 命中 / 分母 | 值 |",
        "|---|---:|---:|",
        f"| Planned Wave Recall | {summary['planned_wave_hit']} / {summary['planned_wave_total']} | {percent(summary['planned_wave_recall'])} |",
        f"| Fresh Wave Precision | {summary['fresh_wave_hit']} / {summary['fresh_wave_total']} | {percent(summary['fresh_wave_precision'])} |",
        f"| Fine-grained Node Recall | {summary['fine_planned_hit']} / {summary['fine_planned_total']} | {percent(summary['fine_node_recall'])} |",
        f"| Fine-grained Node Precision | {summary['fine_actual_hit']} / {summary['fine_actual_total']} | {percent(summary['fine_node_precision'])} |",
        "",
        f"- 清单任务: **{summary['tasks']}**；可评分轨迹: **{summary['tasks_scored']}**；损坏/截断且未进入任何分子分母: **{summary['tasks_invalid_trajectory']}**。",
        "",
        "Resume/Repair 不伪装成新 wave，单独报告：",
        "",
        "| 项 | 数量 |",
        "|---|---:|",
        f"| Accepted resume | {summary['accepted_resumes']} |",
        f"| Linked resume | {summary['linked_resumes']} |",
        f"| Orphan resume | {summary['orphan_resumes']} |",
        f"| Rejected Agent/AgentSwarm call（不计入 wave） | {summary['rejected_calls']} |",
        f"| Repair fresh wave | {summary['repair_fresh_waves']} |",
        f"| Repair fresh node | {summary['repair_fresh_nodes']} |",
        f"| Resume/Repair Overhead | {percent(summary['resume_repair_overhead_rate'])} |",
        "",
        "## 2. P0 证据固定结果",
        "",
        "| prompt 来源 | 任务数 |",
        "|---|---:|",
        f"| traj_log.config.task_str | {summary['prompt_from_traj_log']} |",
        f"| run.log launch command | {summary['prompt_from_run_log']} |",
        f"| reference fallback | {summary['prompt_reference_fallback']} |",
        f"| 实际 prompt 与当前参考 prompt 哈希不同 | {summary['prompt_reference_mismatch']} |",
        "",
        f"- 条件分支 N/A wave: **{summary['ineligible_wave_total']}**（不进 Recall 分母，保留原因和证据）。",
        f"- 由委派前 runtime manifest 证据固定的动态 fan-out wave: **{summary['verified_runtime_fanout_waves']}**。",
        f"- 已知且显式保留的 schema 诊断行: **{summary['known_diagnostic_lines']}**（其他非 JSON 行 fail closed）。",
        "- 每题的 raw stream 路径、大小、SHA-256、prompt 来源/哈希在 `caidian_v2_source_manifest.tsv`。",
        "- 每题的完整实际 prompt 副本在 `runtime_prompts/`。",
        "",
        "## 3. 新旧口径边界",
        "",
        "v2 粗判是有序 wave 对齐，同时校验 action set 和 fan-out 区间；v2 细判仅在已对齐 wave 内做 1:1 节点匹配，同域的 22 个 Canvas owner 仍是 22 个点。",
        "",
    ]
    if legacy:
        lines += [
            "Legacy 数字只用于回放对照，不与 v2 混合成一个指标：",
            "",
            "| Legacy 指标 | 值 |",
            "|---|---:|",
            f"| 宽判 Recall | {legacy['wide_recall']:.1f}% |",
            f"| 宽判 Precision | {legacy['wide_precision']:.1f}% |",
            f"| 严判 Recall | {legacy['strict_recall']:.1f}% |",
            f"| 严判 Precision | {legacy['strict_precision']:.1f}% |",
            "",
        ]
    lines += [
        "## 4. 需人工优先复核的执行偏差与额外事件",
        "",
        "| task | Wave R | Wave P | Node R | Node P | resume | reject | repair wave | prompt |",
        "|---|---:|---:|---:|---:|---:|---:|---:|---|",
    ]
    priority = sorted(
        [record for record in records if record.get("score") is not None],
        key=lambda record: (
            record["score"]["planned_wave_recall"] if record["score"]["planned_wave_recall"] is not None else 1.0,
            record["score"]["fine"]["recall"] if record["score"]["fine"]["recall"] is not None else 1.0,
            record["task"],
        ),
    )
    for record in priority:
        score, fine, overhead = record["score"], record["score"]["fine"], record["score"]["overhead"]
        if (
            (score["planned_wave_recall"] in {None, 1.0})
            and (score["fresh_wave_precision"] in {None, 1.0})
            and (fine["recall"] in {None, 1.0})
            and (fine["precision"] in {None, 1.0})
            and not overhead["accepted_resumes"]
            and not overhead["rejected_calls"]
            and not overhead["repair_fresh_waves"]
        ):
            continue
        lines.append(
            f"| {record['task']} | {percent(score['planned_wave_recall'])} | {percent(score['fresh_wave_precision'])} | "
            f"{percent(fine['recall'])} | {percent(fine['precision'])} | {overhead['accepted_resumes']} | "
            f"{overhead['rejected_calls']} | {overhead['repair_fresh_waves']} | {record['prompt']['provenance']} |"
        )
    lines += [
        "",
        "### 不可评分轨迹（fail closed）",
        "",
    ]
    invalid_records = [record for record in records if record.get("score") is None]
    if invalid_records:
        lines += ["| task | 原因 |", "|---|---|"]
        for record in invalid_records:
            reason = str(record.get("trajectory_error") or "").replace("|", "\\|")
            lines.append(f"| {record['task']} | {reason} |")
    else:
        lines.append("无。")
    lines += [
        "",
        "## 5. 可复现性",
        "",
        "- `caidian_v2_detail.json`：统一 IR、请求/结果接受证据、有序 wave 对齐、1:1 节点匹配、条件分母。",
        "- `caidian_v2_tasks.csv`：每题四组指标与 resume/repair/reject 数量。",
        "- `caidian_v2_source_manifest.tsv`：精确来源与哈希，不使用 mtime 选 slot。",
        "- `caidian_v2_summary.json`：机器可读汇总。",
        "",
    ]
    output.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--gt", type=Path, required=True)
    parser.add_argument("--dump", type=Path, required=True)
    parser.add_argument("--source-map", type=Path, required=True)
    parser.add_argument("--outdir", type=Path, required=True)
    parser.add_argument("--eligibility-overrides", type=Path)
    parser.add_argument("--legacy-data", type=Path)
    parser.add_argument("--expected-tasks", type=int, default=167)
    parser.add_argument(
        "--replace-output",
        action="store_true",
        help="Atomically replace an existing output directory and retain its previous version.",
    )
    args = parser.parse_args()

    gt_files = sorted(path for path in args.gt.glob("*.task.md") if not path.name.startswith("._"))
    if len(gt_files) != args.expected_tasks:
        raise SystemExit(f"Expected {args.expected_tasks} GT files, found {len(gt_files)}")
    source_map = load_source_map(args.source_map)
    tasks = [path.name.removesuffix(".task.md") for path in gt_files]
    if set(tasks) != set(source_map):
        missing = sorted(set(tasks) - set(source_map))
        extra = sorted(set(source_map) - set(tasks))
        raise SystemExit(f"GT/SOURCE_MAP mismatch: missing={missing}, extra={extra}")
    overrides = read_json(args.eligibility_overrides, {})
    validate_override_catalog(overrides, tasks)
    legacy = legacy_summary(args.legacy_data)

    publish_outdir = args.outdir
    publish_outdir.parent.mkdir(parents=True, exist_ok=True)
    if publish_outdir.exists() and not args.replace_output:
        raise SystemExit(
            f"Output directory already exists: {publish_outdir}. "
            "Use --replace-output for a transactional replacement."
        )
    work_outdir = Path(
        tempfile.mkdtemp(
            prefix=f".{publish_outdir.name}.staging-",
            dir=publish_outdir.parent,
        )
    )
    prompt_dir = work_outdir / "runtime_prompts"
    prompt_dir.mkdir(parents=True, exist_ok=True)
    records: list[dict[str, Any]] = []
    for gt_path in gt_files:
        task = gt_path.name.removesuffix(".task.md")
        reference_prompt = gt_path.read_text(encoding="utf-8", errors="replace")
        source_row = source_map[task]
        artifacts = resolve_run_artifacts(args.dump, task, source_row)
        prompt = extract_runtime_prompt(artifacts, reference_prompt)
        (prompt_dir / f"{task}.task.md").write_text(prompt["text"], encoding="utf-8")
        planned = parse_planned_ir(task, prompt["text"])
        apply_verified_overrides(planned, overrides, artifacts.raw_stream)
        source = {
            **source_row,
            "run_dir": str(artifacts.run_dir),
            "raw_stream": str(artifacts.raw_stream),
            "raw_stream_bytes": artifacts.raw_stream.stat().st_size,
            "raw_stream_sha256": sha256_file(artifacts.raw_stream),
            "traj_log": str(artifacts.traj_log) if artifacts.traj_log else None,
            "run_log": str(artifacts.run_log) if artifacts.run_log else None,
        }
        try:
            actual = extract_actual_ir(task, artifacts.raw_stream)
            score = score_task(planned, actual)
            trajectory_error = None
        except TrajectoryIntegrityError as exc:
            actual = None
            score = None
            trajectory_error = str(exc)
        records.append(
            {
                "task": task,
                "trajectory_valid": score is not None,
                "trajectory_error": trajectory_error,
                "source": source,
                "prompt": {key: value for key, value in prompt.items() if key != "text"},
                "planned_ir": planned,
                "actual_ir": actual,
                "score": score,
            }
        )

    summary = aggregate_scores(records)
    meta = {
        "version": VERSION,
        "generated_utc": datetime.now(timezone.utc).isoformat(),
        "gt": str(args.gt),
        "dump": str(args.dump),
        "source_map": str(args.source_map),
        "source_map_sha256": sha256_file(args.source_map),
        "eligibility_overrides": str(args.eligibility_overrides) if args.eligibility_overrides else None,
        "eligibility_overrides_sha256": sha256_file(args.eligibility_overrides) if args.eligibility_overrides else None,
        "legacy": legacy,
        "scorer_sha256": sha256_file(Path(__file__).with_name("caidian_v2.py")),
        "generator_sha256": sha256_file(Path(__file__)),
    }
    (work_outdir / "caidian_v2_summary.json").write_text(
        json.dumps({"meta": meta, "summary": summary}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (work_outdir / "caidian_v2_detail.json").write_text(
        json.dumps({"meta": meta, "summary": summary, "tasks": records}, ensure_ascii=False, indent=1),
        encoding="utf-8",
    )
    write_task_csv(work_outdir / "caidian_v2_tasks.csv", records)
    write_source_manifest(work_outdir / "caidian_v2_source_manifest.tsv", records)
    render_report(work_outdir / "caidian_v2_report.md", records, summary, legacy, args)
    write_checksums(work_outdir)

    previous_outdir: Path | None = None
    if publish_outdir.exists():
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S.%fZ")
        previous_outdir = publish_outdir.parent / f".{publish_outdir.name}.previous-{stamp}"
        os.replace(publish_outdir, previous_outdir)
    try:
        os.replace(work_outdir, publish_outdir)
    except BaseException:
        if previous_outdir is not None and previous_outdir.exists() and not publish_outdir.exists():
            os.replace(previous_outdir, publish_outdir)
        raise
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print(f"wrote {publish_outdir}")
    if previous_outdir is not None:
        print(f"retained previous output at {previous_outdir}")


if __name__ == "__main__":
    main()
