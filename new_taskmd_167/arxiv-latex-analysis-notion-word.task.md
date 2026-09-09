You are a research group coordinator who needs to document analysis of three foundational papers in language model research. Your task is to analyze three specific papers available in the paper database, create structured notes in your team's Notion workspace, and compile a formal Word report.

The three papers you need to analyze are: "Scaling Laws for Neural Language Models" (arXiv ID 2301.07041), "Training language models to follow instructions with human feedback" (arXiv ID 2203.11171), and "OPT: Open Pre-trained Transformer Language Models" (arXiv ID 2205.01068). Use the arXiv and LaTeX analysis tools to retrieve full text and analyze each paper's structure, sections, and key contributions.

For each paper, create a Notion page in your workspace with the paper title as the page title, prefixed with "Paper: ". Include structured notes in the page body covering the paper's main contributions, methodology, and key findings based on the full text you retrieve.

Create a Word document called Paper_Analysis_Report.docx in your workspace. The document should start with a title "Comparative Paper Analysis Report" and include a section for each of the three papers. Each section heading should identify the paper clearly. Within each section, describe the paper's title, key contributions, and methodology. After the three individual sections, include a final section titled "Comparative Analysis" that draws connections between the three papers, discussing how they relate to each other in terms of their contributions to large language model research.

Send an email to research_lead@university.edu with subject "Paper Analysis Report Ready" summarizing what you have done and noting that the Word document and Notion pages are ready for review.

To solve this task efficiently,

The three fixed-paper source stage is small enough for the main agent; use sub-agents only for the two independent output owners. All research is read-only.

Resolve the task-visible workspace directory from the runtime allowed-directory listing (or the literal accessible-workspace directory already supplied by the harness), freeze it as `<workspace_root>`, and reject bare `/workspace` and dumps/run-specific paths. Freeze the three task-specified ID/title slots: `2301.07041 / Scaling Laws for Neural Language Models`, `2203.11171 / Training language models to follow instructions with human feedback`, and `2205.01068 / OPT: Open Pre-trained Transformer Language Models`.

Issue the arXiv record and arxiv-latex section calls for all three IDs in one native-parallel source response. Verify each returned stable ID/title even if the task's pairing is unusual; do not silently replace a task slot from outside knowledge. For each slot extract title, complete authors when returned, contributions, methodology, findings, section list/count, provenance, and missing/conflict status. Reconcile once and freeze one literal canonical three-paper packet; the Notion pages, Word report, and email consume this exact packet without re-researching or recomputing claims.

1. Wave 1 — dispatch exactly 2 coder sub-agents in parallel as individual Agent calls:

   1. Notion owner: search exact titles `Paper: <canonical title>`; reuse one exact match, create at most one page on zero matches, and block on duplicates. Own all three pages in one agent so parent IDs and idempotency stay consistent. Each page must contain substantive Contribution, Methodology, and Findings blocks copied from the canonical packet. Read all three pages back and return IDs, exact titles, stable-ID coverage, and block validation. Do not write the Word file or email.
   2. Word owner: create `<workspace_root>/Paper_Analysis_Report.docx` exactly once, preferably atomically with `python-docx`, with the required report title, one clearly headed section per canonical paper, and a final Comparative Analysis grounded only in those records. Read it back and return path, heading order, exact paper/title coverage, and validation. Do not modify Notion or email.

In the same response that launches Wave 1, the main agent paginates Sent search for exact To `research_lead@university.edu` and Subject `Paper Analysis Report Ready`, reading every candidate. After both outputs return, send exactly once only if no exact match exists, mentioning the verified Word filename and the three verified Notion page titles/IDs. Never retry an ambiguous send. Accept only after all pages, the document, and exactly one message read back against the canonical packet.
