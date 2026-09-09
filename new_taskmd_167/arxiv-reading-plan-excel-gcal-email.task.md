You are a research coordinator for an LLM agent research reading group. Your workspace contains paper_ids.txt with a list of eight arXiv paper IDs and Reading_Guidelines.md describing the reading group format and discussion expectations.

Your task is to create a structured reading plan for the group. Begin by retrieving paper metadata from the preprint repository for every paper ID listed in paper_ids.txt.

How the repository (the arXiv tools) works: it does NOT support a direct lookup by arXiv ID. It offers two relevant tools. One lists all papers currently stored in the repository and returns each paper's title and abstract. The other is a keyword search that matches on words in a paper's title and abstract and returns, for each match, the arXiv ID, title, authors, abstract, categories, and published date. Because the search matches on title/abstract text rather than the ID string, searching for an ID such as "2302.01560" will return nothing. Recommended flow: first list all stored papers to obtain their titles and abstracts, then for each paper run a keyword search using distinctive words from its title to retrieve the full metadata (published date and categories), and match the returned arXiv ID back to the corresponding entry in paper_ids.txt.

Then create an Excel file called Reading_Plan.xlsx in your workspace with two sheets. The first sheet must be named Papers and contain exactly eight rows (one per paper) with these columns: ArXiv_ID, Title, Authors (the full author list, comma-separated), Published_Date (formatted as YYYY-MM-DD), Category (the paper's primary arXiv category, e.g. cs.CL), Abstract_Summary (the first 100 characters of the abstract, exactly as returned by the repository), Assigned_Session (a number from 1 to 8 assigned in the order the papers appear in paper_ids.txt). The second sheet must be named Schedule and contain exactly eight rows (one per reading session) with these columns: Session_Number, Session_Date (formatted as YYYY-MM-DD), Paper_Count, Topics_Covered. Session 1 starts on Monday March 9, 2026 and each subsequent session is one week later. Each session covers one paper. Topics_Covered should be the abbreviated title or topic area of that session's paper.

Next, create eight events on the shared calendar, one for each reading session. Each event title must follow the pattern "Reading Session N: [abbreviated paper title]" where N is the session number. Schedule each event on the corresponding Monday starting from the Monday of the launch week, at 10:00 to 11:30 UTC. Include the full paper title and arXiv ID in the event description.

Finally, send an email from coordinator@lab.example.com to reading-group@lab.example.com. The subject must contain "LLM Agent Research Reading Plan". The body must mention the total number of papers, the date range covered by the reading sessions, and the names or topics of at least three of the eight papers.

When all outputs are produced, call claim_done.

To solve this task efficiently, follow this exact orchestration plan. A wave must finish and return all stated handoffs before the next wave begins. Within a wave, dispatch exactly one sub-agent for each numbered sub-task and run those sub-agents concurrently.

Before Wave 1, the main agent must read paper_ids.txt and Reading_Guidelines.md in read-only mode, preserve the exact eight-ID source order, session numbering, Monday schedule, UTC time, workbook schema, recipient, and subject contract, and compute the eight Mondays from 2026-03-09 without changing order. It must then list the stored repository collection exactly once, match each requested ID to its listed title, and freeze an explicit eight-row manifest containing `row_no`, requested arXiv ID, and the corresponding full title. If any requested ID is missing or ambiguous in that listing, stop rather than dispatching a guessed search.

1. Wave 1 — dispatch exactly 8 explore sub-agents in parallel, one for each source-position row 1 through 8 in the frozen manifest:

   1. Use one explore sub-agent (RP1) to own manifest row 1 only.
   2. Use one explore sub-agent (RP2) to own manifest row 2 only.
   3. Use one explore sub-agent (RP3) to own manifest row 3 only.
   4. Use one explore sub-agent (RP4) to own manifest row 4 only.
   5. Use one explore sub-agent (RP5) to own manifest row 5 only.
   6. Use one explore sub-agent (RP6) to own manifest row 6 only.
   7. Use one explore sub-agent (RP7) to own manifest row 7 only.
   8. Use one explore sub-agent (RP8) to own manifest row 8 only.

For each numbered sub-task, provide that agent with only its frozen `row_no`, requested arXiv ID, and full title. The explore sub-agent must search using distinctive words from that exact title rather than the ID string, select the result whose returned ID exactly equals its requested ID, and return one normalized record containing the original row number, exact ID/title, full authors, published date, primary category, and the first 100 abstract characters exactly as returned, plus result-count and ambiguity controls. It must not list the collection again, process another manifest row, download bodies, or write files.

After Wave 1, the main agent must verify eight unique requested IDs in source order, exact 100-character slicing, dates/categories, and session dates, then freeze one eight-row papers/schedule packet.

2. Wave 2 — dispatch exactly 2 sub-agents in parallel, one for each numbered sub-task:

   1. Use one coder sub-agent to own Reading_Plan.xlsx exclusively. Create it exactly once with Papers and Schedule, exact headers, exactly eight rows per sheet, literal session numbers/counts, original paper order, weekly dates from 2026-03-09, and abbreviated topics. Read back both sheets and return path, sheet order, row counts, ID/date controls, and validation. Do not create calendar or email state.

   2. Use one coder sub-agent to own all eight calendar events exclusively. Search the relevant eight-week window first, create at most one missing Reading Session N event per frozen record from 10:00 to 11:30 UTC with full title and ID in its description, then read back the entire event set. Return eight event IDs/titles/dates, deduplication results, and validation. Do not write the workbook or send email.

After Wave 2, the main agent must reconcile workbook rows and event dates/titles. It must then search sent mail and send at most one missing email from coordinator@lab.example.com to reading-group@lab.example.com with a subject containing LLM Agent Research Reading Plan, total eight, the exact date range, and at least three verified paper topics. After final file, event, and message checks, it must call claim_done exactly once.
