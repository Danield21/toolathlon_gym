Examples:

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
