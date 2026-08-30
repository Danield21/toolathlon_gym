from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from caidian_v2 import (  # noqa: E402
    RunArtifacts,
    TrajectoryIntegrityError,
    apply_verified_overrides,
    derive_predelegation_fanout,
    extract_actual_ir,
    extract_runtime_prompt,
    make_node_ir,
    ordered_wave_alignment,
    parse_planned_ir,
    score_task,
    sha256_file,
    validate_override_catalog,
)


def wave(wave_id: str, nodes: list[dict], exact: int) -> dict:
    ordinal = int(wave_id.rsplit("_", 1)[-1]) if wave_id.rsplit("_", 1)[-1].isdigit() else 1
    prefix = wave_id.rsplit("_", 1)[0]
    return {
        "wave_id": wave_id,
        "ordinal": ordinal,
        "dependency": "main-agent-prerequisites" if ordinal == 1 else f"{prefix}_{ordinal - 1}",
        "fanout": {"min": exact, "max": exact, "exact": exact, "dynamic": False, "source": "test"},
        "condition": {"eligible": True},
        "nodes": nodes,
        "actions": list(dict.fromkeys(value for node in nodes for value in node["actions"])),
        "domains": list(dict.fromkeys(value for node in nodes for value in node["domains"])),
        "write_roles": list(dict.fromkeys(value for node in nodes for value in node["write_roles"])),
        "targets": list(dict.fromkeys(value for node in nodes for value in node["targets"])),
        "object_ids": list(dict.fromkeys(value for node in nodes for value in node["object_ids"])),
        "agent_types": list(dict.fromkeys(node["agent_type"] for node in nodes)),
        "node_instances": sum(node.get("multiplicity", 1) for node in nodes),
    }


class PlannedIrTests(unittest.TestCase):
    def test_declared_22_agents_are_22_fine_points(self) -> None:
        prompt = """Task body.

To solve this task efficiently, freeze the course manifest first.

1. Wave 1 — dispatch exactly 22 explore sub-agents in parallel, one per verified course:

   1. Each course Explore agent reads Canvas assignments and returns its literal course_id metrics.

After Wave 1, aggregate directly.
"""
        parsed = parse_planned_ir("canvas-test", prompt)
        self.assertEqual(len(parsed["waves"]), 1)
        self.assertEqual(parsed["waves"][0]["fanout"]["exact"], 22)
        self.assertEqual(parsed["waves"][0]["node_instances"], 22)

    def test_verified_conditional_gate_is_na(self) -> None:
        prompt = """Task.
To solve this task efficiently, if required.csv is missing, stop before delegation. The waves apply only if it is present.
1. Wave 1 — dispatch exactly 1 coder sub-agent:
   1. Read required.csv and create payload.json.
"""
        parsed = parse_planned_ir("conditional", prompt)
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            raw.write_text('{"role":"assistant","content":"verified missing"}\n', encoding="utf-8")
            overrides = {
                "tasks": {
                    "conditional": {
                        "raw_stream_sha256": sha256_file(raw),
                        "waves": {
                            "wave_1": {
                                "eligible": False,
                                "reason": "verified missing",
                                "evidence": [{"source": "raw_stream", "contains": "verified missing"}],
                            }
                        },
                    }
                }
            }
            apply_verified_overrides(parsed, overrides, raw)
        self.assertFalse(parsed["waves"][0]["condition"]["eligible"])
        actual = {"fresh_waves": [], "counts": {key: 0 for key in (
            "fresh_waves", "fresh_nodes", "accepted_resumes", "linked_resumes", "orphan_resumes",
            "rejected_calls", "repair_fresh_waves", "repair_fresh_nodes")}}
        score = score_task(parsed, actual)
        self.assertEqual(score["planned_wave_total"], 0)
        self.assertEqual(score["ineligible_wave_total"], 1)
        self.assertIsNone(score["planned_wave_recall"])

    def test_only_explicit_transition_creates_wave_dependency(self) -> None:
        prompt = """Task.
To solve this task efficiently, use two workstreams.
Wave 1 — dispatch exactly 1 Explore sub-agent:
1. Read Canvas.
Wave 2 — dispatch exactly 1 Coder sub-agent:
1. Compute an independent summary.
"""
        parsed = parse_planned_ir("dependencies", prompt)
        self.assertEqual(
            [item["dependency"] for item in parsed["waves"]],
            ["main-agent-prerequisites", "main-agent-prerequisites"],
        )

    def test_explicit_dependency_keeps_named_wave(self) -> None:
        prompt = """Task.
To solve this task efficiently, use three workstreams.
Wave 1 — dispatch exactly 1 Explore sub-agent:
1. Read Canvas source A.
Wave 2 — dispatch exactly 1 Explore sub-agent:
1. Independently read source B.
After Wave 1, continue with the final computation.
Wave 3 — dispatch exactly 1 Coder sub-agent:
1. Compute the final report from source A.
"""
        parsed = parse_planned_ir("named-dependency", prompt)
        self.assertEqual(parsed["waves"][2]["dependency"], "wave_1")

    def test_dependency_in_current_header_is_parsed_without_prior_body_leakage(self) -> None:
        prompt = """Task.
To solve this task efficiently, use three workstreams.
Wave 1 — dispatch exactly 1 Explore sub-agent:
1. Read source A using outputs from Wave 9 as a field label.
Wave 2 — dispatch exactly 1 Explore sub-agent:
1. Independently read source B.
Wave 3 — after Wave 1, dispatch exactly 1 Coder sub-agent:
1. Compute the final report.
"""
        parsed = parse_planned_ir("header-dependency", prompt)
        self.assertEqual(parsed["waves"][1]["dependency"], "main-agent-prerequisites")
        self.assertEqual(parsed["waves"][2]["dependency"], "wave_1")

    def test_agent_internal_after_step_does_not_create_wave_dependency(self) -> None:
        prompt = """Task.
To solve this task efficiently, use two independent workstreams.
Wave 1 — dispatch exactly 1 Explore sub-agent:
Read all Canvas pages.
After reading all pages, return a compact packet.
Wave 2 — dispatch exactly 1 Coder sub-agent independently:
Compute an unrelated local summary.
"""
        parsed = parse_planned_ir("internal-after", prompt)
        self.assertEqual(parsed["waves"][1]["dependency"], "main-agent-prerequisites")

    def test_while_wave_runs_is_explicit_parallel_not_dependency(self) -> None:
        prompt = """Task.
To solve this task efficiently, use two parallel workstreams.
Wave 1 — dispatch exactly 1 Explore sub-agent:
Read Canvas source A.
While Wave 1 runs, launch the independent second source branch below.
Wave 2 — dispatch exactly 1 Explore sub-agent:
Read source B independently.
"""
        parsed = parse_planned_ir("while-parallel", prompt)
        self.assertEqual(parsed["waves"][1]["dependency"], "main-agent-prerequisites")

    def test_explicit_wave_dependency_outranks_intra_wave_concurrently(self) -> None:
        prompt = """Task.
To solve this task efficiently, use two workstreams.
Wave 1 — dispatch exactly 1 Explore sub-agent:
Read source A.
Wave 2 — after Wave 1, concurrently dispatch exactly 2 Coder sub-agents:
Write the two independent outputs using Wave 1 evidence.
"""
        parsed = parse_planned_ir("explicit-over-parallel", prompt)
        self.assertEqual(parsed["waves"][1]["dependency"], "wave_1")

    def test_legacy_two_shards_is_prompt_derived_exact_fanout(self) -> None:
        prompt = """Task.
To solve this task efficiently, use the following sub-task:
1. Please dispatch multiple explore sub-agents to work in parallel on the following sub-tasks:
Collect the first half of the course manifest, keeping the two sections represented across the two shards.
"""
        parsed = parse_planned_ir("two-shards", prompt)
        self.assertEqual(parsed["waves"][0]["fanout"]["exact"], 2)
        self.assertEqual(parsed["waves"][0]["node_instances"], 2)

    def test_negative_write_constraint_does_not_create_write_action(self) -> None:
        node = make_node_ir(
            "Read Canvas pages and return a compact report. Do not write files or create Excel workbooks.",
            node_id="n",
            declared_type="explore",
        )
        self.assertEqual(node["actions"], ["read"])
        self.assertEqual(node["write_roles"], [])

    def test_numbered_recipe_is_one_repeated_owner_template(self) -> None:
        prompt = """Task.
To solve this task efficiently, use one collector wave.
Wave 1 — dispatch exactly 7 coder sub-agents in parallel, one per manifest course. Every collector follows the same literal recipe with only its course substituted:
1. Read page 1.
2. Follow pagination.
3. Compute the mean.
4. Return counts.
5. Stop.
After Wave 1, aggregate directly.
"""
        parsed = parse_planned_ir("shared-recipe", prompt)
        planned_wave = parsed["waves"][0]
        self.assertEqual(len(planned_wave["nodes"]), 1)
        self.assertEqual(planned_wave["node_instances"], 7)

        actual_nodes = [
            make_node_ir(
                f"Collect and compute stats for one Canvas course_id={course_id} with the same literal recipe.",
                node_id=f"a{course_id}",
                declared_type="coder",
            )
            for course_id in range(1, 8)
        ]
        actual_ir = {"fresh_waves": [wave("fresh_wave_1", actual_nodes, 7)], "counts": {
            "fresh_waves": 1, "fresh_nodes": 7, "accepted_resumes": 0, "linked_resumes": 0,
            "orphan_resumes": 0, "rejected_calls": 0, "repair_fresh_waves": 0, "repair_fresh_nodes": 0}}
        score = score_task(parsed, actual_ir)
        self.assertEqual(score["fine"]["planned_hit"], 7)

    def test_main_agent_tail_is_not_attached_to_wave_nodes(self) -> None:
        prompt = """Task.
To solve this task efficiently, use one writer wave.
Wave 1 — dispatch exactly 2 coder sub-agents in parallel:
1. Create Report.xlsx.
2. Create Summary.docx.
While Wave 1 runs, the main agent sends email and creates a calendar event.
"""
        parsed = parse_planned_ir("tail", prompt)
        nodes = parsed["waves"][0]["nodes"]
        self.assertEqual([node["write_roles"] for node in nodes], [["excel"], ["word"]])

    def test_legacy_multiple_without_items_is_dynamic(self) -> None:
        prompt = """Task.
To solve this task efficiently, use the following sub-task:
1. Please dispatch multiple explore sub-agents to work in parallel on the following sub-tasks:
Collect the first and second halves of the Canvas course manifest.
"""
        parsed = parse_planned_ir("legacy-multiple", prompt)
        fanout = parsed["waves"][0]["fanout"]
        self.assertTrue(fanout["dynamic"])
        self.assertEqual(fanout["min"], 2)

    def test_payload_entities_are_not_automatically_shards(self) -> None:
        pagination = make_node_ir(
            "Course collector reads Canvas page=1 and page=2 for course_id=8.",
            node_id="course",
            declared_type="coder",
        )
        writer = make_node_ir(
            "The Word writer creates one section for every frozen paper.",
            node_id="writer",
            declared_type="coder",
        )
        self.assertEqual(pagination["shard_key"], "course")
        self.assertIsNone(writer["shard_key"])

    def test_google_spreadsheet_is_a_write_target(self) -> None:
        node = make_node_ir(
            "The sole owner creates a Google spreadsheet and reads it back.",
            node_id="gsheet",
            declared_type="coder",
        )
        self.assertEqual(node["actions"], ["read", "write"])
        self.assertEqual(node["write_roles"], ["gsheet"])

    def test_range_numbered_items_preserve_mixed_roles_and_fanout(self) -> None:
        prompt = """Task.
To solve this task efficiently, use one mixed source wave.
Wave 1 — start exactly 6 Explore source owners in parallel:
1-2. One Coder per Canvas course computes course metrics.
3-6. One Explore agent per order shard reads WooCommerce.
After Wave 1, merge directly.
"""
        parsed = parse_planned_ir("mixed-ranges", prompt)
        planned_wave = parsed["waves"][0]
        self.assertEqual(planned_wave["node_instances"], 6)
        self.assertEqual(
            [(node["agent_type"], node["multiplicity"]) for node in planned_wave["nodes"]],
            [("coder", 2), ("explore", 4)],
        )

    def test_numbered_item_role_overrides_wave_default(self) -> None:
        prompt = """Task.
To solve this task efficiently, use one wave.
Wave 1 — dispatch exactly 2 Explore sub-agents:
1. Use one Explore sub-agent to read a browser page.
2. Use one strictly read-only Coder to query a recipe category.
"""
        parsed = parse_planned_ir("mixed-default", prompt)
        self.assertEqual(
            [node["agent_type"] for node in parsed["waves"][0]["nodes"]],
            ["explore", "coder"],
        )

    def test_action_plan_noun_does_not_override_declared_coder(self) -> None:
        node = make_node_ir(
            "Create action_plans.xlsx with an evidence-backed action plan.",
            node_id="coder",
            declared_type="coder",
        )
        self.assertEqual(node["agent_type"], "coder")

    def test_actual_description_scopes_write_role(self) -> None:
        node = make_node_ir(
            "Create Report.xlsx, a Notion database, and calendar events.",
            node_id="notion",
            declared_type="coder",
            signature_text="Notion DB + action pages",
        )
        self.assertEqual(node["write_roles"], ["notion"])

    def test_read_only_published_metadata_is_not_a_write(self) -> None:
        node = make_node_ir(
            "Read papers for a knowledge base and return each published year. Do not create Notion or write files.",
            node_id="reader",
            declared_type="explore",
        )
        self.assertNotIn("write", node["actions"])

    def test_read_json_input_does_not_become_excel_writers_json_role(self) -> None:
        node = make_node_ir(
            "The sole owner of Report.xlsx reads input.json and computes metrics. Create the workbook and read it back.",
            node_id="writer",
            declared_type="coder",
        )
        self.assertEqual(node["write_roles"], ["excel"])


class ActualIrTests(unittest.TestCase):
    def test_acceptance_resume_and_rejection_semantics(self) -> None:
        rows = [
            {"role": "assistant", "content": "fresh", "tool_calls": [{"id": "c1", "function": {"name": "Agent", "arguments": json.dumps({"subagent_type": "explore", "description": "Read Canvas course 1", "prompt": "Query course_id=1 read-only"})}}]},
            {"role": "tool", "tool_call_id": "c1", "content": "agent_id: agent-7\nactual_subagent_type: explore\nstatus: failed\nsubagent stopped"},
            {"role": "assistant", "content": "continue", "tool_calls": [{"id": "c2", "function": {"name": "Agent", "arguments": json.dumps({"resume": "agent-7", "description": "Resume owner", "prompt": "Continue Canvas course_id=1 collection"})}}]},
            {"role": "tool", "tool_call_id": "c2", "content": "agent_id: agent-7\nactual_subagent_type: explore\nstatus: completed"},
            {"role": "assistant", "content": "bad", "tool_calls": [{"id": "c3", "function": {"name": "AgentSwarm", "arguments": json.dumps({"subagent_type": "explore", "items": [1, 2], "prompt_template": "Read Canvas course {{item}}", "extra": "bad"})}}]},
            {"role": "tool", "tool_call_id": "c3", "content": "ToolError: invalid argument combination; rejected by harness"},
        ]
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            raw.write_text("\n".join(json.dumps(row) for row in rows) + "\n", encoding="utf-8")
            actual = extract_actual_ir("t", raw)
        self.assertEqual(actual["counts"]["fresh_waves"], 1)
        self.assertEqual(actual["counts"]["fresh_nodes"], 1)
        self.assertEqual(actual["counts"]["accepted_resumes"], 1)
        self.assertEqual(actual["counts"]["linked_resumes"], 1)
        self.assertEqual(actual["counts"]["rejected_calls"], 1)
        self.assertIn("RESUME", actual["fresh_waves"][0]["nodes"][0]["text"])

    def test_agentswarm_items_expand_one_to_one(self) -> None:
        rows = [
            {"role": "assistant", "content": "swarm", "tool_calls": [{"id": "s1", "function": {"name": "AgentSwarm", "arguments": json.dumps({"subagent_type": "coder", "items": ["2401.00001", "2401.00002", "2401.00003"], "prompt_template": "Retrieve paper {{item}} and write {{item}}.tex"})}}]},
            {"role": "tool", "tool_call_id": "s1", "content": '<agent_swarm_result><summary>completed: 3</summary><subagent agent_id="agent-0" item="2401.00001" outcome="completed"/><subagent agent_id="agent-1" item="2401.00002" outcome="completed"/><subagent agent_id="agent-2" item="2401.00003" outcome="completed"/></agent_swarm_result>'},
        ]
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            raw.write_text("\n".join(json.dumps(row) for row in rows), encoding="utf-8")
            actual = extract_actual_ir("t", raw)
        self.assertEqual(actual["counts"]["fresh_waves"], 1)
        self.assertEqual(actual["counts"]["fresh_nodes"], 3)
        self.assertEqual([node["agent_id"] for node in actual["fresh_waves"][0]["nodes"]], ["agent-0", "agent-1", "agent-2"])

    def test_agentswarm_counts_only_accepted_result_envelopes(self) -> None:
        rows = [
            {"role": "assistant", "tool_calls": [{"id": "s1", "function": {"name": "AgentSwarm", "arguments": json.dumps({"subagent_type": "explore", "items": ["A", "B", "C"], "prompt_template": "Read item {{item}}"})}}]},
            {"role": "tool", "tool_call_id": "s1", "content": '<agent_swarm_result><summary>completed: 2</summary><subagent agent_id="agent-7" item="B" outcome="completed"/><subagent agent_id="agent-8" item="C" outcome="completed"/></agent_swarm_result>'},
        ]
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            raw.write_text("\n".join(json.dumps(row) for row in rows), encoding="utf-8")
            actual = extract_actual_ir("partial", raw)
        self.assertEqual(actual["counts"]["fresh_nodes"], 2)
        self.assertEqual(actual["call_records"][0]["requested_nodes"], 3)
        self.assertEqual(actual["call_records"][0]["accepted_nodes"], 2)
        self.assertEqual(
            [node["agent_id"] for node in actual["fresh_waves"][0]["nodes"]],
            ["agent-7", "agent-8"],
        )
        self.assertEqual(
            [node["text"].split("SWARM_ITEM=")[-1] for node in actual["fresh_waves"][0]["nodes"]],
            ["B", "C"],
        )

    def test_agentswarm_resume_map_links_each_existing_agent(self) -> None:
        rows = [
            {"role": "assistant", "tool_calls": [{"id": "s1", "function": {"name": "AgentSwarm", "arguments": json.dumps({"subagent_type": "explore", "items": ["course=1", "course=2"], "prompt_template": "Read {{item}}"})}}]},
            {"role": "tool", "tool_call_id": "s1", "content": '<agent_swarm_result><subagent agent_id="agent-0" outcome="failed"/><subagent agent_id="agent-1" outcome="failed"/></agent_swarm_result>'},
            {"role": "assistant", "tool_calls": [{"id": "s2", "function": {"name": "AgentSwarm", "arguments": json.dumps({"description": "resume", "resume_agent_ids": {"agent-0": "Continue course=1", "agent-1": "Continue course=2"}})}}]},
            {"role": "tool", "tool_call_id": "s2", "content": '<agent_swarm_result><subagent mode="resume" agent_id="agent-0" outcome="completed"/><subagent mode="resume" agent_id="agent-1" outcome="completed"/></agent_swarm_result>'},
        ]
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            raw.write_text("\n".join(json.dumps(row) for row in rows), encoding="utf-8")
            actual = extract_actual_ir("resume-map", raw)
        self.assertEqual(actual["counts"]["fresh_nodes"], 2)
        self.assertEqual(actual["counts"]["accepted_resumes"], 2)
        self.assertEqual(actual["counts"]["linked_resumes"], 2)

    def test_truncated_swarm_preview_uses_accepted_summary_count(self) -> None:
        rows = [
            {"role": "assistant", "tool_calls": [{"id": "s1", "function": {"name": "AgentSwarm", "arguments": json.dumps({"subagent_type": "explore", "items": [1, 2, 3, 4, 5], "prompt_template": "Read {{item}}"})}}]},
            {"role": "tool", "tool_call_id": "s1", "content": 'Tool output exceeded 50000 characters; showing a preview only.\n<agent_swarm_result><summary>completed: 4, failed: 1</summary><subagent agent_id="agent-0"/><subagent agent_id="agent-1"/>'},
        ]
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            raw.write_text("\n".join(json.dumps(row) for row in rows), encoding="utf-8")
            actual = extract_actual_ir("preview", raw)
        self.assertEqual(actual["counts"]["fresh_nodes"], 5)
        self.assertEqual(
            [node["agent_id"] for node in actual["fresh_waves"][0]["nodes"]],
            ["agent-0", "agent-1", "agent-2", "agent-3", "agent-4"],
        )

    def test_incomplete_full_accept_requires_preview_marker(self) -> None:
        rows = [
            {"role": "assistant", "tool_calls": [{"id": "s1", "function": {"name": "AgentSwarm", "arguments": json.dumps({"items": ["A", "B", "C"], "prompt_template": "Read {{item}}"})}}]},
            {"role": "tool", "tool_call_id": "s1", "content": '<agent_swarm_result><summary>completed: 3</summary><subagent agent_id="agent-2" item="B"/></agent_swarm_result>'},
        ]
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            raw.write_text("\n".join(json.dumps(row) for row in rows), encoding="utf-8")
            with self.assertRaisesRegex(TrajectoryIntegrityError, "without an explicit full-accept preview"):
                extract_actual_ir("not-preview", raw)

    def test_truncated_preview_uses_visible_nonfirst_item_mapping(self) -> None:
        rows = [
            {"role": "assistant", "tool_calls": [{"id": "s1", "function": {"name": "AgentSwarm", "arguments": json.dumps({"items": ["A", "B", "C"], "prompt_template": "Read {{item}}"})}}]},
            {"role": "tool", "tool_call_id": "s1", "content": 'Tool output exceeded 50000 characters; showing a preview only.\n<agent_swarm_result><summary>completed: 3</summary><subagent agent_id="agent-2" item="B"/></agent_swarm_result>'},
        ]
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            raw.write_text("\n".join(json.dumps(row) for row in rows), encoding="utf-8")
            actual = extract_actual_ir("preview-nonfirst", raw)
        self.assertEqual(
            [node["agent_id"] for node in actual["fresh_waves"][0]["nodes"]],
            ["agent-1", "agent-2", "agent-3"],
        )
        self.assertEqual(
            [node["text"].split("SWARM_ITEM=")[-1] for node in actual["fresh_waves"][0]["nodes"]],
            ["A", "B", "C"],
        )

    def test_corrupt_jsonl_fails_closed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            raw.write_text('{"role":"assistant"}\n{broken\n', encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "corrupt raw_stream JSONL"):
                extract_actual_ir("corrupt", raw)

    def test_corrupt_tool_arguments_fail_closed(self) -> None:
        rows = [
            {"role": "assistant", "tool_calls": [{"id": "a1", "function": {"name": "Agent", "arguments": "{broken"}}]},
            {"role": "tool", "tool_call_id": "a1", "content": "agent_id: agent-1\nstatus: completed"},
        ]
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            raw.write_text("\n".join(json.dumps(row) for row in rows), encoding="utf-8")
            with self.assertRaisesRegex(TrajectoryIntegrityError, "corrupt arguments"):
                extract_actual_ir("corrupt-arguments", raw)

    def test_non_object_tool_arguments_fail_closed(self) -> None:
        rows = [
            {"role": "assistant", "tool_calls": [{"id": "a1", "function": {"name": "Agent", "arguments": "[]"}}]},
            {"role": "tool", "tool_call_id": "a1", "content": "agent_id: agent-1\nstatus: completed"},
        ]
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            raw.write_text("\n".join(json.dumps(row) for row in rows), encoding="utf-8")
            with self.assertRaisesRegex(TrajectoryIntegrityError, "non-object JSON value"):
                extract_actual_ir("non-object-arguments", raw)

    def test_duplicate_agent_request_id_fails_closed(self) -> None:
        rows = [
            {"role": "assistant", "tool_calls": [{"id": "dup", "function": {"name": "Agent", "arguments": json.dumps({"prompt": "Read A"})}}]},
            {"role": "assistant", "tool_calls": [{"id": "dup", "function": {"name": "Agent", "arguments": json.dumps({"prompt": "Read B"})}}]},
            {"role": "tool", "tool_call_id": "dup", "content": "agent_id: agent-2\nstatus: completed"},
        ]
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            raw.write_text("\n".join(json.dumps(row) for row in rows), encoding="utf-8")
            with self.assertRaisesRegex(TrajectoryIntegrityError, "duplicate Agent/AgentSwarm request ID"):
                extract_actual_ir("duplicate-request", raw)

    def test_duplicate_tool_result_id_fails_closed(self) -> None:
        rows = [
            {"role": "assistant", "tool_calls": [{"id": "dup", "function": {"name": "Agent", "arguments": json.dumps({"prompt": "Read A"})}}]},
            {"role": "tool", "tool_call_id": "dup", "content": "agent_id: agent-1\nstatus: completed"},
            {"role": "tool", "tool_call_id": "dup", "content": "agent_id: agent-2\nstatus: completed"},
        ]
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            raw.write_text("\n".join(json.dumps(row) for row in rows), encoding="utf-8")
            with self.assertRaisesRegex(TrajectoryIntegrityError, "duplicate tool result ID"):
                extract_actual_ir("duplicate-result", raw)

    def test_known_schema_diagnostic_is_retained_not_silent(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            raw.write_text(
                '{"role":"assistant","content":"ok"}\n'
                'unknown format "json" ignored in schema at path "#/properties/icon"\n',
                encoding="utf-8",
            )
            actual = extract_actual_ir("diagnostic", raw)
        self.assertEqual(actual["counts"]["known_diagnostic_lines"], 1)
        self.assertEqual(actual["diagnostic_lines"][0]["line"], 2)

    def test_agentswarm_item_defines_shard_not_pagination_text(self) -> None:
        rows = [
            {"role": "assistant", "content": "swarm", "tool_calls": [{"id": "s1", "function": {"name": "AgentSwarm", "arguments": json.dumps({"subagent_type": "coder", "items": ["course_id=8"], "prompt_template": "Collect one course via page=1,2,3 for {{item}}"})}}]},
            {"role": "tool", "tool_call_id": "s1", "content": '<agent_swarm_result><summary>completed: 1</summary><subagent agent_id="agent-0" item="course_id=8" outcome="completed"/></agent_swarm_result>'},
        ]
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            raw.write_text("\n".join(json.dumps(row) for row in rows), encoding="utf-8")
            actual = extract_actual_ir("t", raw)
        self.assertEqual(actual["fresh_waves"][0]["nodes"][0]["shard_key"], "course")

    def test_agentswarm_page_interval_item_precedes_course_id(self) -> None:
        node = make_node_ir(
            "Read one Canvas page interval for the assigned item.",
            node_id="page",
            declared_type="explore",
            item="course_id=8;page_start=1;page_end=4",
        )
        self.assertEqual(node["shard_key"], "page")

    def test_explicit_teacher_assignment_precedes_generic_course_item(self) -> None:
        node = make_node_ir(
            "Wave 1 teacher lookup owner reads TeacherEnrollment for one Spring course.",
            node_id="teacher",
            declared_type="explore",
            item="course_id=8",
        )
        self.assertEqual(node["shard_key"], "teacher")


class MatchingTests(unittest.TestCase):
    def test_order_is_not_ignored(self) -> None:
        planned = [
            wave("wave_1", [make_node_ir("Explore reads Canvas course_id=1", node_id="p1", declared_type="explore")], 1),
            wave("wave_2", [make_node_ir("Coder creates Report.xlsx", node_id="p2", declared_type="coder")], 1),
        ]
        actual = [
            wave("fresh_wave_1", [make_node_ir("Coder creates Report.xlsx", node_id="a1", declared_type="coder")], 1),
            wave("fresh_wave_2", [make_node_ir("Explore reads Canvas course_id=1", node_id="a2", declared_type="explore")], 1),
        ]
        aligned = ordered_wave_alignment(planned, actual)
        self.assertEqual(len(aligned["pairs"]), 1)

    def test_fanout_mismatch_is_not_contract_hit(self) -> None:
        planned_node = make_node_ir("Explore reads Canvas one per course", node_id="p", declared_type="explore", multiplicity=22)
        actual_nodes = [make_node_ir(f"Explore reads Canvas course_id={index}", node_id=f"a{index}", declared_type="explore") for index in range(1, 22)]
        aligned = ordered_wave_alignment([wave("wave_1", [planned_node], 22)], [wave("fresh_wave_1", actual_nodes, 21)])
        self.assertEqual(len(aligned["pairs"]), 1)
        self.assertFalse(aligned["pairs"][0]["contract_hit"])

    def test_extra_fresh_wave_lowers_precision(self) -> None:
        planned_wave = wave("wave_1", [make_node_ir("Coder creates Report.xlsx", node_id="p", declared_type="coder")], 1)
        actual_good = wave("fresh_wave_1", [make_node_ir("Coder creates Report.xlsx", node_id="a", declared_type="coder")], 1)
        actual_extra = wave("fresh_wave_2", [make_node_ir("Explore reads Canvas", node_id="x", declared_type="explore")], 1)
        planned_ir = {"waves": [planned_wave]}
        actual_ir = {"fresh_waves": [actual_good, actual_extra], "counts": {
            "fresh_waves": 2, "fresh_nodes": 2, "accepted_resumes": 0, "linked_resumes": 0,
            "orphan_resumes": 0, "rejected_calls": 0, "repair_fresh_waves": 0, "repair_fresh_nodes": 0}}
        score = score_task(planned_ir, actual_ir)
        self.assertEqual(score["planned_wave_recall"], 1.0)
        self.assertEqual(score["fresh_wave_precision"], 0.5)

    def test_contract_hit_wins_alignment_tie_before_extra_wave(self) -> None:
        planned = wave(
            "wave_1",
            [make_node_ir("Coder creates Report.xlsx", node_id="p", declared_type="coder")],
            1,
        )
        actual_good = wave(
            "fresh_wave_1",
            [make_node_ir("Coder creates Report.xlsx", node_id="good", declared_type="coder")],
            1,
        )
        actual_extra = wave(
            "fresh_wave_2",
            [make_node_ir("Coder creates Report.xlsx", node_id="extra", declared_type="coder")],
            1,
        )
        aligned = ordered_wave_alignment([planned], [actual_good, actual_extra])
        self.assertEqual(aligned["pairs"][0]["actual_index"], 0)
        self.assertTrue(aligned["pairs"][0]["contract_hit"])

    def test_dependency_compares_aligned_target_not_only_non_main(self) -> None:
        planned = [
            wave("wave_1", [make_node_ir("Explore reads source A", node_id="p1", declared_type="explore")], 1),
            wave("wave_2", [make_node_ir("Explore reads source B", node_id="p2", declared_type="explore")], 1),
            wave("wave_3", [make_node_ir("Coder computes report C", node_id="p3", declared_type="coder")], 1),
        ]
        planned[1]["dependency"] = "main-agent-prerequisites"
        planned[2]["dependency"] = "wave_1"
        actual = [
            wave("fresh_wave_1", [make_node_ir("Explore reads source A", node_id="a1", declared_type="explore")], 1),
            wave("fresh_wave_2", [make_node_ir("Explore reads source B", node_id="a2", declared_type="explore")], 1),
            wave("fresh_wave_3", [make_node_ir("Coder computes report C", node_id="a3", declared_type="coder")], 1),
        ]
        actual[1]["dependency"] = "main-agent-prerequisites"
        wrong = ordered_wave_alignment(planned, actual)
        self.assertFalse(wrong["pairs"][2]["dependency_adherent"])
        self.assertFalse(wrong["pairs"][2]["contract_hit"])
        actual[2]["dependency"] = "fresh_wave_1"
        correct = ordered_wave_alignment(planned, actual)
        self.assertTrue(correct["pairs"][2]["dependency_adherent"])
        self.assertTrue(correct["pairs"][2]["contract_hit"])

    def test_dynamic_template_requires_verified_manifest(self) -> None:
        template = make_node_ir("Explore reads one Canvas course", node_id="p", declared_type="explore")
        planned_wave = wave("wave_1", [template], 1)
        planned_wave["fanout"] = {"min": 1, "max": 128, "exact": None, "dynamic": True, "source": "test"}
        actual_nodes = [make_node_ir(f"Explore reads Canvas course_id={index}", node_id=f"a{index}", declared_type="explore") for index in range(1, 22)]
        actual_ir = {"fresh_waves": [wave("fresh_wave_1", actual_nodes, 21)], "counts": {
            "fresh_waves": 1, "fresh_nodes": 21, "accepted_resumes": 0, "linked_resumes": 0,
            "orphan_resumes": 0, "rejected_calls": 0, "repair_fresh_waves": 0, "repair_fresh_nodes": 0}}
        planned_ir = {"task": "dynamic", "waves": [planned_wave]}
        with self.assertRaisesRegex(ValueError, "lack a verified runtime manifest"):
            score_task(planned_ir, actual_ir)
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            course_rows = [{"id": index, "name": f"Course {index}"} for index in range(1, 23)]
            rows = [
                {"role": "assistant", "tool_calls": [{"id": "courses", "function": {"name": "mcp__canvas__canvas_list_courses", "arguments": "{}"}}]},
                {"role": "tool", "tool_call_id": "courses", "content": json.dumps(course_rows)},
                {"role": "assistant", "content": "I claim only 7 owners", "tool_calls": [{"id": "agents", "function": {"name": "AgentSwarm", "arguments": json.dumps({"items": list(range(7)), "prompt_template": "Read {{item}}"})}}]},
            ]
            raw.write_text("\n".join(json.dumps(row) for row in rows), encoding="utf-8")
            overrides = {"tasks": {"dynamic": {
                "raw_stream_sha256": sha256_file(raw),
                "waves": {"wave_1": {
                    "fanout": {"exact": 22, "derivation": {"kind": "canvas_course_count"}},
                    "reason": "pre-dispatch manifest",
                    "evidence": [{"source": "pre_delegation_tool_results", "derivation": "canvas_course_count"}],
                }},
            }}}
            apply_verified_overrides(planned_ir, overrides, raw)
        score = score_task(planned_ir, actual_ir)
        self.assertEqual(score["fine"]["planned_total"], 22)
        self.assertEqual(score["fine"]["planned_hit"], 21)
        proof = planned_ir["waves"][0]["fanout"]["derivation_proof"]
        self.assertEqual(proof["tool_result_lines"], [2])
        self.assertEqual(proof["first_delegation_line"], 3)

    def test_grade_chunk_fanout_is_derived_before_delegation(self) -> None:
        rows = [
            {"role": "assistant", "tool_calls": [{"id": "courses", "function": {"name": "mcp__canvas__canvas_list_courses", "arguments": "{}"}}]},
            {"role": "tool", "tool_call_id": "courses", "content": json.dumps([{"id": 1, "name": "A"}, {"id": 2, "name": "B"}])},
            {"role": "assistant", "tool_calls": [
                {"id": "g1", "function": {"name": "mcp__canvas__canvas_get_course_grades", "arguments": json.dumps({"course_id": 1, "page": 1, "per_page": 1, "type": ["StudentEnrollment"]})}},
                {"id": "g2", "function": {"name": "mcp__canvas__canvas_get_course_grades", "arguments": json.dumps({"course_id": 2, "page": 1, "per_page": 1, "type": ["StudentEnrollment"]})}},
            ]},
            {"role": "tool", "tool_call_id": "g1", "content": json.dumps({"enrollments": [], "pagination": {"total_count": 401}})},
            {"role": "tool", "tool_call_id": "g2", "content": json.dumps({"enrollments": [], "pagination": {"total_count": 800}})},
            {"role": "assistant", "content": "I claim one owner", "tool_calls": [{"id": "agents", "function": {"name": "AgentSwarm", "arguments": json.dumps({"items": [1], "prompt_template": "Read {{item}}"})}}]},
        ]
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            raw.write_text("\n".join(json.dumps(row) for row in rows), encoding="utf-8")
            exact, proof = derive_predelegation_fanout(
                "grade-chunks",
                raw,
                {"kind": "canvas_grade_page_chunks", "page_size": 100, "max_pages_per_item": 4},
            )
        self.assertEqual(exact, 4)
        self.assertEqual(proof["tool_result_lines"], [2, 4, 5])
        self.assertEqual(proof["first_delegation_line"], 6)

    def test_multi_template_multiplicities_must_match_derived_components(self) -> None:
        page = make_node_ir("Explore one Canvas page owner", node_id="wave_1.node_1", declared_type="explore")
        teacher = make_node_ir("Explore one teacher owner", node_id="wave_1.node_2", declared_type="explore")
        planned_wave = wave("wave_1", [page, teacher], 2)
        planned_wave["fanout"] = {"min": 2, "max": 128, "exact": None, "dynamic": True, "source": "test"}
        planned_ir = {"task": "component-proof", "waves": [planned_wave]}
        courses = [
            {"id": 1, "name": "A (Spring 2014)"},
            {"id": 2, "name": "B (Spring 2014)"},
        ]
        rows = [
            {"role": "assistant", "tool_calls": [{"id": "courses", "function": {"name": "mcp__canvas__canvas_list_courses", "arguments": "{}"}}]},
            {"role": "tool", "tool_call_id": "courses", "content": json.dumps(courses)},
            {"role": "assistant", "tool_calls": [
                {"id": "g1", "function": {"name": "mcp__canvas__canvas_get_course_grades", "arguments": json.dumps({"course_id": 1, "page": 1, "per_page": 1, "type": ["StudentEnrollment"]})}},
                {"id": "g2", "function": {"name": "mcp__canvas__canvas_get_course_grades", "arguments": json.dumps({"course_id": 2, "page": 1, "per_page": 1, "type": ["StudentEnrollment"]})}},
            ]},
            {"role": "tool", "tool_call_id": "g1", "content": json.dumps({"enrollments": [], "pagination": {"total_count": 401}})},
            {"role": "tool", "tool_call_id": "g2", "content": json.dumps({"enrollments": [], "pagination": {"total_count": 800}})},
            {"role": "assistant", "tool_calls": [{"id": "agents", "function": {"name": "AgentSwarm", "arguments": json.dumps({"items": [1], "prompt_template": "Read {{item}}"})}}]},
        ]
        with tempfile.TemporaryDirectory() as directory:
            raw = Path(directory) / "raw_stream.jsonl"
            raw.write_text("\n".join(json.dumps(row) for row in rows), encoding="utf-8")
            overrides = {"tasks": {"component-proof": {
                "raw_stream_sha256": sha256_file(raw),
                "waves": {"wave_1": {
                    "fanout": {
                        "exact": 15,
                        "node_multiplicities": {"wave_1.node_1": 12, "wave_1.node_2": 3},
                        "derivation": {
                            "kind": "canvas_grade_pages_plus_course_count",
                            "course_name_suffix": "(Spring 2014)",
                            "page_size": 100,
                        },
                    },
                    "reason": "tampered components",
                    "evidence": [{
                        "source": "pre_delegation_tool_results",
                        "derivation": "canvas_grade_pages_plus_course_count",
                    }],
                }},
            }}}
            with self.assertRaisesRegex(ValueError, "disagree with derived components"):
                apply_verified_overrides(planned_ir, overrides, raw)

    def test_split_dynamic_templates_keep_full_planned_denominator(self) -> None:
        page = make_node_ir("Explore one Canvas page owner", node_id="page", declared_type="explore")
        teacher = make_node_ir("Explore one teacher owner", node_id="teacher", declared_type="explore")
        planned_wave = wave("wave_1", [page, teacher], 2)
        planned_wave["fanout"] = {"min": 3, "max": 3, "exact": 3, "dynamic": False, "source": "verified-runtime-manifest"}
        planned_wave["nodes"][0]["multiplicity"] = 2
        planned_wave["nodes"][1]["multiplicity"] = 1
        planned_wave["node_instances"] = 3
        page_nodes = [make_node_ir("Explore one Canvas page owner", node_id=f"p{i}", declared_type="explore") for i in range(2)]
        teacher_nodes = [make_node_ir("Explore one teacher owner", node_id="t", declared_type="explore")]
        actual_waves = [wave("fresh_wave_1", page_nodes, 2), wave("fresh_wave_2", teacher_nodes, 1)]
        actual_ir = {"fresh_waves": actual_waves, "counts": {
            "fresh_waves": 2, "fresh_nodes": 3, "accepted_resumes": 0, "linked_resumes": 0,
            "orphan_resumes": 0, "rejected_calls": 0, "repair_fresh_waves": 0, "repair_fresh_nodes": 0}}
        score = score_task({"task": "split", "waves": [planned_wave]}, actual_ir)
        self.assertEqual(score["fine"]["planned_total"], 3)
        self.assertLess(score["fine"]["planned_hit"], 3)

    def test_unplanned_write_and_missing_compute_do_not_hit(self) -> None:
        planned_read = wave("wave_1", [make_node_ir("Explore reads Canvas", node_id="p1", declared_type="explore")], 1)
        actual_write = wave("fresh_wave_1", [make_node_ir("Explore reads Canvas and creates report.json", node_id="a1", declared_type="explore")], 1)
        self.assertFalse(ordered_wave_alignment([planned_read], [actual_write])["pairs"])

        planned_compute = wave("wave_1", [make_node_ir("Coder reads Canvas and computes totals", node_id="p2", declared_type="coder")], 1)
        actual_read = wave("fresh_wave_1", [make_node_ir("Coder reads Canvas", node_id="a2", declared_type="coder")], 1)
        pair = ordered_wave_alignment([planned_compute], [actual_read])["pairs"][0]
        self.assertFalse(pair["contract_hit"])

    def test_missing_one_explicit_write_role_does_not_hit(self) -> None:
        planned = wave("wave_1", [make_node_ir("Coder creates Report.pptx and Report.pdf", node_id="p", declared_type="coder")], 1)
        actual = wave("fresh_wave_1", [make_node_ir("Coder creates Report.pptx", node_id="a", declared_type="coder")], 1)
        pair = ordered_wave_alignment([planned], [actual])["pairs"][0]
        self.assertFalse(pair["write_role_adherent"])
        self.assertFalse(pair["contract_hit"])

    def test_invalid_override_types_and_unknown_tasks_fail_closed(self) -> None:
        invalid = {"tasks": {"known": {
            "raw_stream_sha256": "0" * 64,
            "waves": {"wave_1": {
                "eligible": "false",
                "reason": "bad type",
                "evidence": [{"source": "raw_stream", "contains": "proof"}],
            }},
        }}}
        with self.assertRaisesRegex(ValueError, "eligible must be a JSON boolean"):
            validate_override_catalog(invalid, ["known"])
        unknown = {"tasks": {"unknown": {
            "raw_stream_sha256": "0" * 64,
            "waves": {"wave_1": {
                "eligible": False,
                "reason": "proof",
                "evidence": [{"source": "raw_stream", "contains": "proof"}],
            }},
        }}}
        with self.assertRaisesRegex(ValueError, "unknown tasks"):
            validate_override_catalog(unknown, ["known"])


class PromptProvenanceTests(unittest.TestCase):
    def test_traj_log_precedes_run_log_and_reference(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            raw = root / "raw_stream.jsonl"
            raw.write_text("x" * 1001, encoding="utf-8")
            traj = root / "traj_log.json"
            traj.write_text(json.dumps({"config": {"task_str": "runtime prompt from traj"}}), encoding="utf-8")
            run_log = root / "run.log"
            run_log.write_text("[kimi] launching: kimi -p runtime prompt from log\n ... (home=/x)", encoding="utf-8")
            artifacts = RunArtifacts("t", root, raw, traj, run_log)
            prompt = extract_runtime_prompt(artifacts, "reference")
        self.assertEqual(prompt["provenance"], "traj_log.config.task_str")
        self.assertFalse(prompt["matches_reference"])

    def test_run_log_fallback(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            raw = root / "raw_stream.jsonl"
            raw.write_text("x" * 1001, encoding="utf-8")
            run_log = root / "run.log"
            runtime = "This is a sufficiently long runtime prompt. To solve this task efficiently, do not dispatch sub-agents."
            run_log.write_text(f"[kimi] launching: kimi -p {runtime}\n ... (home=/x)", encoding="utf-8")
            artifacts = RunArtifacts("t", root, raw, None, run_log)
            prompt = extract_runtime_prompt(artifacts, "reference")
        self.assertEqual(prompt["provenance"], "run.log launch command")
        self.assertEqual(prompt["text"], runtime)


if __name__ == "__main__":
    unittest.main()
