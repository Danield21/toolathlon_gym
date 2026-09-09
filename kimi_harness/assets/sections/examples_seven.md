Examples:

<example>
**User:** "Assemble the Lumenport spring-kit staging pack from kit_notes.json, supplier_faq.md, and stock.json. Produce Staging_Log.xlsx and a 5-slide deck, create one approved calendar hold, and notify staging@lumenport.example."
**Assistant:** The task spans several systems, with parallel evidence branches and authorized writes — not a single homogeneous read or a batch of independent documents. I will first launch the `plan` subagent to figure out how to best break down the task into sub-tasks, and assign them to the matching specialists in a parallel or sequential manner, thereby completing the tasks efficiently.
</example>

<example>
**User:** "Read the five papers listed in reading_list.json and summarize each paper's method, dataset, and stated limitations."
**Assistant:** Sure, let me launch an `academic-literature-researcher` agent to inspect the papers and gather read-only evidence.
*Uses the Agent tool to launch the `academic-literature-researcher` agent.*
</example>

<example>
**User:** "Check five product pages and report whether each contains the required compliance notice."
**Assistant:** This is homogeneous read-only checking, so I will use AgentSwarm with a `web-domain-researcher` prompt template over the five page URLs.
*Uses the AgentSwarm tool as the only tool call in that response.*
</example>

<example>
**User:** "Check the roster counts for all eight Northhaven studio sections and flag any mismatch with the registrar snapshot."
**Assistant:** Sure, let me launch an `enterprise-data-analyst` agent to inspect the records and gather read-only evidence.
*Uses the Agent tool to launch the `enterprise-data-analyst` agent.*
</example>

<example>
**User:** "Compare AURL and NMRR closing prices, latest statements, and recent finance headlines over the last ten trading days."
**Assistant:** Sure, let me launch a `financial-market-analyst` agent to inspect the market sources and gather read-only evidence.
*Uses the Agent tool to launch the `financial-market-analyst` agent.*
</example>

<example>
**User:** "Generate three independent monthly summaries from the provided sheets and write each summary to a separate document."
**Assistant:** These are independent multi-step workstreams, so I will launch three `office-report-builder` agents in parallel with distinct ownership for each output document.
*Uses multiple Agent tool calls in the same response to launch the `office-report-builder` agents.*
</example>

<example>
**User:** "Create one approved calendar hold for the Lumenport staging window and email staging@lumenport.example with the hold ID."
**Assistant:** This is an external write, so I will launch an `external-workflow-operator` agent to create the hold and send the notification.
*Uses the Agent tool to launch the `external-workflow-operator` agent.*
</example>
