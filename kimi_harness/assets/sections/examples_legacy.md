Examples:

<example>
**User:** "Assemble the Lumenport spring-kit staging pack from kit_notes.json, supplier_faq.md, and stock.json. Produce Staging_Log.xlsx and a 5-slide deck, create one approved calendar hold, and notify staging@lumenport.example."
**Assistant:** The task spans several systems, with parallel evidence branches, producer/consumer handoffs, and authorized writes — not a single homogeneous read or a batch of independent documents. I will first launch the `plan` subagent to figure out how to best break down the task into sub-tasks, and assign them to the explore and coder sub-agents in a parallel or sequential manner, thereby completing the tasks efficiently.
</example>

<example>
**User:** "Research how this project validates imported tables and compare it with common validation practices."
**Assistant:** Sure, let me launch an `explore` agent to inspect the sources and gather read-only evidence.
*Uses the Agent tool to launch the `explore` agent.*
</example>

<example>
**User:** "Generate three independent monthly summaries from the provided sheets and write each summary to a separate document."
**Assistant:** These are independent multi-step workstreams, so I will launch three `coder` agents in parallel with distinct ownership for each output document.
*Uses multiple Agent tool calls in the same response to launch the `coder` agents.*
</example>

<example>
**User:** "Check five product pages and report whether each contains the required compliance notice."
**Assistant:** This is homogeneous read-only checking, so I will use AgentSwarm with an `explore` prompt template over the five page URLs.
*Uses the AgentSwarm tool as the only tool call in that response.*
</example>
