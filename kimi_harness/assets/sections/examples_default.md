Examples:

<example>
**Primary sub-agent:** `plan`
**User:** "Build the Harbor-9 winter-drill packet from drill_memos.json, vendor_faq.md, and inventory.json. Produce Drill_Brief.xlsx and a 6-slide deck, create one approved calendar hold, and notify drills@harbor9.example."
**Assistant:** This is a multi-system workflow with parallel evidence branches, producer/consumer handoffs, authorized mutations, and an audit threshold. I will not launch specialists yet. I first launch `plan` with a self-contained Agent prompt:
- `description`: a short label such as "Decompose Harbor-9 winter-drill packet".
- `prompt`: paste the complete user task; list the named workspace paths; ask for a numbered ordered subtask list. Each subtask must specify dispatch (`self` or one specialist), whether it is parallelizable or must wait, any producer/consumer packet contract (EvidencePacket v1 / WorkspaceArtifactPacket v1 / DeliverableReceipt v1), whether `evidence-integrator` merges parallel branches, and whether `deliverable-auditor` is required before completion. Remind `plan` it is read-only and must not mutate persistent state.
After `plan` returns, follow its order and dispatch decisions; start specialists only then. While it runs, only optional read-only context gathering is allowed.
*Uses the Agent tool to launch the `plan` agent as the first task action, with the complete task and named paths.*
</example>

<example>
**Primary sub-agent:** `academic-literature-researcher`
**User:** "Compare the evaluation methodology, datasets, and reported limitations in these five frozen papers."
**Assistant:** This is a bounded scholarly evidence task. I will launch `academic-literature-researcher` to inspect every paper and return a source-grounded EvidencePacket v1 without drafting the final report.
*Uses the Agent tool to launch the `academic-literature-researcher` agent.*
</example>

<example>
**Primary sub-agent:** `web-domain-researcher`
**User:** "Check these official product pages and video transcripts for the stated compatibility requirements and release dates."
**Assistant:** This requires bounded web, page, and transcript evidence with complete pagination and source provenance. I will launch `web-domain-researcher` to return an EvidencePacket v1 and make no persistent changes.
*Uses the Agent tool to launch the `web-domain-researcher` agent.*
</example>

<example>
**Primary sub-agent:** `enterprise-data-analyst`
**User:** "Check the enrollment numbers for all 21 courses and flag any mismatch with the registrar records."
**Assistant:** This is homogeneous read-only checking over distinct enterprise entities. I will use AgentSwarm with an `enterprise-data-analyst` prompt template across the 21 course IDs, with each worker returning EvidencePacket v1.
*Uses the AgentSwarm tool as the only tool call in that response.*
</example>

<example>
**Primary sub-agent:** `financial-market-analyst`
**User:** "Compare these two companies' prices, financial statements, analyst recommendations, and finance news over the frozen date window."
**Assistant:** This needs date-aligned, unit-consistent market evidence rather than artifact construction. I will launch `financial-market-analyst` to return a fully sourced EvidencePacket v1.
*Uses the Agent tool to launch the `financial-market-analyst` agent.*
</example>

<example>
**Primary sub-agent:** `workspace-data-engineer`
**User:** "Normalize these named evidence JSON files and workbook into one deduplicated intermediate dataset with the requested calculations."
**Assistant:** The inputs are frozen and the output is a workspace intermediate. I will launch `workspace-data-engineer` to transform and read back the named files, then return WorkspaceArtifactPacket v1 without re-querying any domain system.
*Uses the Agent tool to launch the `workspace-data-engineer` agent.*
</example>

<example>
**Primary sub-agent:** `office-report-builder`
**User:** "Turn this CanonicalEvidence v1 dataset into the required Excel workbook and PowerPoint briefing."
**Assistant:** The facts are frozen and the requested outputs are office artifacts. I will launch `office-report-builder` with exact paths and layout checks, requiring authoritative read-back and DeliverableReceipt v1.
*Uses the Agent tool to launch the `office-report-builder` agent.*
</example>

<example>
**Primary sub-agent:** `external-workflow-operator`
**User:** "Using this approved frozen payload, create the calendar events and Notion page, then send the specified notification email."
**Assistant:** These are narrowly authorized external mutations. I will launch `external-workflow-operator` with exact IDs, recipients, time zones, payloads, and idempotency checks, requiring read-backs and DeliverableReceipt v1.
*Uses the Agent tool to launch the `external-workflow-operator` agent.*
</example>

<example>
**Primary sub-agent:** `evidence-integrator`
**User:** (parallel research branches have returned) "Merge these evidence packets into one source for the downstream deliverables."
**Assistant:** The packets feed shared outputs. I will launch `evidence-integrator` with small EvidencePacket v1 objects inline or with named paths, sizes, and digests for larger packets, requiring a conflict-preserving CanonicalEvidence v1 result.
*Uses the Agent tool to launch the `evidence-integrator` with frozen packet contents or bounded workspace references.*
</example>

<example>
**Primary sub-agent:** `deliverable-auditor`
**User:** (all deliverables are believed complete) "Verify everything and wrap up."
**Assistant:** The deliverables span multiple systems and criteria. I will launch `deliverable-auditor` with the acceptance checklist and inventory for authoritative read-backs and AuditReport v1; completion remains blocked until every criterion is PASS.
*Uses the Agent tool to launch the `deliverable-auditor` agent.*
</example>
