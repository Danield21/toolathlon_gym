I am preparing to present at the International Conference on Machine Learning Methods. The conference schedule is available as a JSON endpoint at http://localhost:30412/api/conference.json. Please fetch the conference schedule data first.

Next, I need to review the research papers in our LaTeX paper archive. The archive holds several papers, but only three are relevant to this conference. They are identified by arXiv ID: 2301.07041 (scaling laws for neural language models), 2203.11171 (training language models to follow instructions with human feedback — RLHF / instruction tuning), and 2205.01068 (OPT, open-source pre-trained language models). All other papers in the archive are out of scope, even if they mention machine learning. For each of these three relevant papers, retrieve the full content including all section titles, abstracts, and key findings from its LaTeX source.

Note: the LaTeX source for each paper (its section headings and content) is available through the arxiv-latex tools. Use list_paper_sections, get_paper_section, and get_paper_prompt with the paper's arXiv ID (the three IDs above) to retrieve each paper's full content. The three listed IDs are the only relevant papers; do not include any other papers from the archive.

Create and run a Python script called conference_prep_builder.py in the workspace. The script should read the conference schedule from conference_schedule.json and the paper analysis data from paper_analysis.json (both files you create from the fetched data), cross-reference which papers are relevant to which conference sessions, extract word counts per section, and output conference_prep_results.json.

Create an Excel file called Conference_Prep_Tracker.xlsx with three sheets. The first sheet Paper_Sections should have columns Paper_ID, Paper_Title, Section_Title, and Section_Word_Count (approximate word count of each section; write these counts as literal numbers, not Excel formulas). Include all sections from all relevant papers found. Sort by Paper_ID then section order.

The second sheet Conference_Schedule should have columns Session_ID, Session_Title, Date, Time, Room, and Related_Papers (comma-separated paper titles that match the session topic). The data comes from the conference JSON.

The third sheet Presentation_Notes should have columns Slide_Number (1 through at least 8), Topic, Key_Points (2-3 sentence summary), and Source_Paper. Create a logical presentation flow starting with introduction, covering each paper's main findings, and ending with a conclusion.

Finally, create a page in the knowledge base titled "Conference Prep Notes" with structured content. The page should include properties for Conference_Name (text), Presentation_Date (text showing the date), Status (select with value "In Progress"), and Paper_Count (number, equal to the number of relevant papers). Also add descriptive content blocks — at least one block per paper — summarizing each paper's key contribution and how it relates to the conference themes.

To solve this task efficiently, use exactly one paper-retrieval wave. The main agent calls `fetch_json` once for the task-supplied conference endpoint, validates the complete conference/session records, and freezes that compact schedule.

1. Wave 1 — dispatch exactly 3 coder sub-agents in parallel through one same-template AgentSwarm call:

   1. Coder 1 owns only arXiv `2301.07041`.
   2. Coder 2 owns only arXiv `2203.11171`.
   3. Coder 3 owns only arXiv `2205.01068`.

Each coder keeps all arXiv-LaTeX calls for its one ID in one context, obtains the prompt, ordered section inventory, and every retrievable section, and writes one unique `<paper_id>.json` containing exact title, abstract, ordered section titles, full section text, word counts, and compact findings. Return only path, hash, ID, section count/order, and read-back controls; do not create required final files or modify Notion.

After Wave 1, the main agent verifies the three IDs/paths once, writes `conference_schedule.json` and `paper_analysis.json`, writes and runs `conference_prep_builder.py`, and verifies `conference_prep_results.json`. It then directly creates `Conference_Prep_Tracker.xlsx` and the single `Conference Prep Notes` page as independent output branches. Use exactly the three required workbook sheets, preserve every paper section and conference session, include at least eight presentation-note rows, and write literal counts. For Notion, exact-title search first, create at most one page, set all four required properties including `Status = In Progress` and `Paper_Count = 3`, and add at least one evidence block per paper. Read each final output once; do not dispatch another wave.
