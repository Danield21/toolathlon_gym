I am writing a comparative analysis of chain-of-thought reasoning methods for my research group. Please look up the chain-of-thought reasoning papers in the arxiv-latex database and help me organize the findings. The database holds several papers, but only five are about chain-of-thought reasoning methods; the others are about word representations and are out of scope. The five reasoning papers, by arXiv ID, are: 2201.11903, 2203.11171, 2205.11916, 2210.03493, and 2305.10601. Use the arxiv-latex tools (get_paper_prompt or get_paper_abstract) with these IDs to retrieve each paper's title and abstract.

Create a Google Sheet called "Reasoning Methods Comparison" with a sheet named "Papers" that lists the reasoning-related papers. The columns should be Title, Year, Method, and Key_Contribution. For the Method column, try to identify the main reasoning approach described in each paper (for example "chain-of-thought prompting", "self-consistency", "zero-shot CoT", "automatic CoT", or "tree of thoughts"). For Key_Contribution, write a brief one-sentence description of the paper's main contribution. Only include papers that are about reasoning methods, not papers about word embeddings or word representations.

Also create a Word document called "Reasoning_Methods_Review.docx" in my workspace. The document should have the title "Chain-of-Thought Reasoning: Methods Comparison" at the top, followed by today's date (2026-03-06). Then include a section for each reasoning paper, mentioning its title and briefly describing its contribution. End with a short conclusion paragraph.

Note: this task has two deliverables — the Google Sheet and the Word document. If you are running as a team of agents, divide the work between agents (for example, assign the spreadsheet to one agent and the Word document to another) or otherwise serialize your writes, so that both deliverables are completed without conflicting writes to the same file.

To solve this task efficiently, follow this exact orchestration plan. A wave must finish and return all stated handoffs before the next wave begins. Within a wave, dispatch exactly one sub-agent for each numbered sub-task and run those sub-agents concurrently.

Before Wave 1, the main agent must freeze the five explicit arXiv IDs, today's required displayed date 2026-03-06, the five allowed reasoning-method labels, and the two independent deliverables. It must preserve the task's instruction to exclude all word-representation distractors.

1. Wave 1 — dispatch exactly 5 explore sub-agents in parallel, one per explicit paper ID:

   1. Use one explore sub-agent to retrieve and analyze only arXiv ID `2201.11903`.
   2. Use one explore sub-agent to retrieve and analyze only arXiv ID `2203.11171`.
   3. Use one explore sub-agent to retrieve and analyze only arXiv ID `2205.11916`.
   4. Use one explore sub-agent to retrieve and analyze only arXiv ID `2210.03493`.
   5. Use one explore sub-agent to retrieve and analyze only arXiv ID `2305.10601`.

Each Wave 1 sub-agent must call get_paper_prompt or get_paper_abstract for its one explicit ID, verify the returned identity, classify Method only from title/abstract evidence, and return one compact record containing ID, title, year derived from the ID/source, Method, one-sentence Key_Contribution, relevance decision, and source-call controls. Explore is read-only, so do not write files, Google Sheets, or Word documents.

After Wave 1, the main agent must verify exactly five distinct requested IDs, reject every embedding paper, normalize one canonical five-row packet, and freeze it for both output owners.

2. Wave 2 — dispatch exactly 2 sub-agents in parallel, one for each numbered sub-task:

   1. Use one coder sub-agent to own the Google Sheet Reasoning Methods Comparison exclusively. Create it once with exactly one Papers sheet, exact Title, Year, Method, Key_Contribution columns, and exactly the five frozen rows; write literal years and read back all rows. Return spreadsheet/sheet identifiers, headers, row count, ID-to-title coverage, and validation. Do not write the Word document.

   2. Use one coder sub-agent to own Reasoning_Methods_Review.docx exclusively. Create it exactly once with the required title, displayed date 2026-03-06, one source-backed section for each of the same five papers, and a conclusion. Read it back and return its path, heading/date/paper coverage, and validation. Do not create or modify the Google Sheet.

After Wave 2, the main agent must reconcile five titles, methods, and contributions across both deliverables, verify single-writer provenance and file/external-state integrity, and complete without adding unrelated papers.
