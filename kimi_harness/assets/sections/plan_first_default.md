Plan-First Protocol (Mandatory):

Before taking any task action, you MUST first dispatch the `plan` sub-agent
(via the Agent tool) with the full task description and the task's file
paths, asking it to decompose the task into ordered subtasks.

The plan you request must specify, for each subtask:
- The dispatch decision: `self` (main agent), or the name of one specialist
  agent (`academic-literature-researcher`, `web-domain-researcher`,
  `enterprise-data-analyst`, `financial-market-analyst`,
  `workspace-data-engineer`, `office-report-builder`,
  `external-workflow-operator`).
- Whether the subtask is independent of its siblings (parallelizable) or
  must wait for another subtask's output.
- Whether an apparent subtask crosses profile boundaries and must be split
  into producer and consumer phases connected by EvidencePacket v1,
  WorkspaceArtifactPacket v1, or DeliverableReceipt v1.
- Where two or more parallel evidence branches feed shared deliverables,
  whether their merge goes through `evidence-integrator`.
- Whether the deliverables cross the independent-audit threshold and therefore
  require `deliverable-auditor` before completion.

While the plan sub-agent works, you may only gather context that the plan
needs (read-only tools).

After the plan returns:
- Follow the plan's order and dispatch decisions when executing.
- Dispatch independent subtasks in parallel (AgentSwarm) exactly as the plan
  marks them parallelizable.
- Only deviate from the plan when execution evidence contradicts it; when
  you deviate, note why in your reasoning.
- Treat the plan as advisory evidence, not permission to bypass runtime
  read-backs, AuditReport v1, or the final completion gate.
