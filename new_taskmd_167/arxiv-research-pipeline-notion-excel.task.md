I need to build a comprehensive research knowledge base for our team studying large language models. Start by searching for recent scholarly papers on topics like "large language models", "prompt engineering", and "in-context learning". Find at least 5 relevant papers.

Then use the paper repository to download and read the full content of the most relevant papers. Analyze their LaTeX source when available to extract detailed methodology sections.

Write and run a Python script called research_synthesizer.py in the workspace that reads papers_metadata.json and papers_contents.json (create both first), extracts key methods, creates a citation network map, calculates relevance scores, and outputs research_synthesis.json.

Create an Excel file called Research_Knowledge_Base.xlsx with three sheets. The first sheet Paper_Catalog should have columns Paper_ID, Title, Authors, Year, Category, and Citation_Count, sorted by Citation_Count descending. The second sheet Method_Comparison should have columns Method_Name, Paper_Source, Key_Innovation, Benchmark_Result, and Applicability ("High", "Medium", "Low"). The third sheet Research_Gaps should have Gap_Area, Current_State, Opportunity, and Priority ("Critical", "Important", "Nice-to-have") columns with at least 4 identified gaps. In the Excel file, write all values as literal text/numbers directly into the cells; do not use Excel formulas.

Create a Notion page titled "LLM Research Hub" with a heading "Large Language Model Research Dashboard". Add paragraphs covering research landscape overview, key papers summary, methodology comparison highlights, and identified research gaps with recommendations for future work.

To solve this task efficiently, use exactly one paper-processing wave. The main agent first reads all task-visible guide and team-interest files, then issues bounded native-parallel searches for the three literal topics `large language models`, `prompt engineering`, and `in-context learning` in both Scholarly and local arXiv. Union by stable ID, reject off-topic results by title/abstract, and freeze exactly five strongest papers covering all three topics in a literal dispatch table `{agent_no, paper_id, title, citation_count, category}`.

1. Wave 1 — dispatch exactly 5 coder sub-agents through one same-template AgentSwarm call:

   1. Coder 1 owns only dispatch-table row 1 and its literal paper ID.
   2. Coder 2 owns only dispatch-table row 2 and its literal paper ID.
   3. Coder 3 owns only dispatch-table row 3 and its literal paper ID.
   4. Coder 4 owns only dispatch-table row 4 and its literal paper ID.
   5. Coder 5 owns only dispatch-table row 5 and its literal paper ID.

For its one row, each coder retrieves/downloads and reads the full local paper, inspects available arXiv-LaTeX sections, and writes one unique `paper_<safe_id>.json` containing exact metadata, authors, year, citation count, category, abstract, full-content/method evidence, benchmark evidence, and source-supported references. Return only its path, ID, schema/count controls, and read status. Do not write any required final filename or modify Notion.

After Wave 1, the main agent checks the five paths and IDs once, consolidates them into `papers_metadata.json` and `papers_contents.json`, writes and runs `research_synthesizer.py` exactly once, and verifies `research_synthesis.json`. It then creates `Research_Knowledge_Base.xlsx` and the single `LLM Research Hub` page as independent native-parallel output branches from that verified synthesis. The workbook must use exactly `Paper_Catalog`, `Method_Comparison`, and `Research_Gaps`, literal values, citation-descending catalog rows, controlled applicability labels, and at least four evidence-backed gaps. The page must contain the exact dashboard heading and the four requested substantive sections. Read each final output once and do not dispatch another wave.
