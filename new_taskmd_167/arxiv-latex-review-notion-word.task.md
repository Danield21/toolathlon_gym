I need to build a knowledge base from recent papers about large language model fine-tuning and alignment. Please access our academic paper repository and read the full content of papers that are relevant to the topics of LLM fine-tuning, alignment, and model architecture improvements. For each relevant paper, read through its sections carefully to understand the methodology and key results. Skip any papers that are about unrelated topics like robotics or physical systems.

Create a Notion page titled LLM Fine-Tuning Knowledge Base. The page should contain blocks for each relevant paper, where each block covers the paper title, the key method introduced, and the main findings from the paper.

Also create a Word document called LLM_Paper_Synthesis.docx in my workspace. The document should start with a heading called LLM Fine-Tuning and Alignment Survey, followed by an introduction paragraph. Then include a dedicated section for each relevant paper that presents the paper title, the author names, a description of the method, and the key results extracted from reading the paper sections.

Finally, create a Google Sheet called LLM Paper Registry with columns for ArXiv_ID, Title, Authors, Published_Date, Key_Contribution, and Method_Category, with one row per relevant paper.

To solve this task efficiently, follow this exact orchestration plan. A wave must finish and return all stated handoffs before the next wave begins. Within a wave, dispatch exactly one sub-agent for each numbered sub-task and run those sub-agents concurrently.

1. Wave 1 — dispatch exactly 1 sub-agent for the following numbered sub-task:

   1. Use one explore sub-agent to cover the task-relevant local arXiv candidate space in read-only mode through complementary fine-tuning, alignment, and architecture queries at the supported per-call bound; union and deduplicate all returned IDs, since the search exposes no cursor. Identify every genuinely relevant record in that observed union, reject robotics or physical-system distractors, and use arxiv-latex only after an ID is discovered to read its available full sections. Return a compact normalized result with ID, title, complete authors, published date, key method, source-backed findings, method category, and section coverage, plus per-query/union/retained/rejected controls. Explore is read-only, so do not create Notion, Word, Google Sheet, or intermediate files.

After Wave 1, the main agent must deduplicate IDs/titles, verify every retained paper against the task's relevance rule and every rejected paper against the exclusion rule, determine both counts only from runtime results, and freeze one normalized packet without importing outside facts.

2. Wave 2 — dispatch exactly 3 sub-agents in parallel, one for each numbered sub-task:

   1. Use one coder sub-agent to own the single Notion page exclusively. Search for LLM Fine-Tuning Knowledge Base, create at most one page, add one substantive block per frozen paper with title, method, and findings, and read it back. Return page ID/title, paper coverage, block count, and validation. Do not write the Word document or Google Sheet.

   2. Use one coder sub-agent to own LLM_Paper_Synthesis.docx exclusively. Create it exactly once with the required heading, introduction, and one author/method/results section for every frozen paper. Read it back and return path, headings, retained-paper coverage, and validation. Do not modify Notion or Google Sheets.

   3. Use one coder sub-agent to own the Google Sheet LLM Paper Registry exclusively. Create it once with the exact six columns and one row per frozen paper, preserving full authors and source dates and using controlled Method_Category values. Read it back and return spreadsheet/sheet IDs, headers, row count, and validation. Do not write Notion or Word state.

After Wave 2, the main agent must reconcile the same stable-ID set, authors, methods, and findings across all deliverables, verify every rejected paper appears nowhere, and complete only after file and external-state checks.
