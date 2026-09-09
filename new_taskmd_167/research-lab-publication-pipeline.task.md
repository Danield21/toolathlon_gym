Your research laboratory must establish a systematic process for tracking relevant publications, managing research literature, and identifying gaps between current research and future opportunities. The first phase involves comprehensively searching for recent publications in your research domain. Identify publications from the past three years across multiple sources including preprint servers and peer-reviewed journals. Focus on papers addressing your laboratory's research questions and methodological approaches. Compile a comprehensive list of all relevant publications with bibliographic information.

In the second phase, you will retrieve full-text papers and supplementary materials for all identified publications. Download PDF versions of papers where available and extract supplementary materials including datasets, code, and additional analyses. Organize all materials in a centralized repository accessible to all laboratory members. Document the source and access date for each publication for citation tracking purposes.

The third phase requires carefully analyzing each publication to extract key methodology and experimental details. For each paper, document the research questions addressed, experimental design, data collection methods, analytical approaches, and key findings. Create detailed summaries highlighting novel methods or tools that might be applicable to your laboratory's work. Extract citations from each paper to identify additional relevant publications not captured in the initial search.

In the fourth phase, you will consolidate all literature into a structured database for easy retrieval and reference. Create an organized repository categorized by topic, methodology, and relevance to specific research projects. Add tags and keywords enabling efficient searching and filtering. Include researcher notes on the significance of each publication to your laboratory's work.

The fifth phase involves comprehensively analyzing your laboratory's current work and publications against the landscape of external research. Identify significant gaps where your laboratory's work is ahead of the field, where you are aligned with current trends, and where external research suggests new directions. Cross-reference your laboratory's ongoing projects with the literature to ensure you are building on existing work appropriately. Document major research gaps and opportunities.

The sixth phase requires synthesizing all findings into a comprehensive research roadmap document. Outline major research directions supported by the literature review. Identify emerging trends and technologies that may impact your research area. Provide recommendations for future research projects, potential collaborators, and interdisciplinary opportunities. Share this roadmap through email with all laboratory members and schedule meetings to discuss priorities and resource allocation.

## Deliverables / Expected Output

Produce your outputs inside the workspace directory. The names and formats below are the expected ones (reasonably similar names/formats are also acceptable, but following them guarantees your work is captured by the evaluation):

1. **Literature database** — consolidate the literature search into a single structured spreadsheet or CSV named `literature_database.csv` (or `literature_database.xlsx`). Include one row per publication with at least the columns `title`, `author`, `year`, and optionally `journal`, `keywords`, `relevance_to_lab`, `source`. It should cover all the publications retrieved during the search. Write values as literal text; do not use spreadsheet formulas.

2. **Bibliography** — produce a `bibliography.bib` file (or a file named like `references...`) containing at least 5 bibliography entries (`@article`, `@misc`, `@inproceedings`, etc.) for the retrieved publications.

3. **Per-paper analysis summaries** — save a summary document for each analyzed paper (e.g. `manuscript_paper1.docx`, `manuscript_paper2.docx`, ...) describing the research questions, methodology, and key findings of that paper.

4. **Research roadmap** — produce a document named `research_roadmap.docx` (or `.md`) that is at least 400 characters long and explicitly contains a **Future Directions** section and a **Recommendations** section. The document should discuss research directions, gaps, and recommendations identified in the gap analysis.

5. **Email to laboratory members** — send an email via the email tool to the laboratory members (a group/lab mailing-list address or an individual member's address) whose subject or body mentions the research roadmap (for example a subject such as "Research Roadmap Update").

6. **Calendar meeting** — schedule a calendar event to discuss the roadmap and priorities (for example a summary such as "Research Roadmap Discussion" or "Lab Meeting").

To solve this task efficiently, the main agent completes this bounded seeded-literature task directly; do not dispatch sub-agents. Finished artifacts plus real email/calendar side effects are required, while an exhaustive open-web survey is not.

Dynamically resolve the task workspace, read config.json, data.csv and every sheet of research_data.xlsx, and call local `list_papers` once to enumerate the complete available inventory. Use the task runtime to define the inclusive past-three-years window. For each inventory title/ID, use one bounded exact-ID/title `search_papers(max_results=10)` call to recover complete local metadata; use at most one additional local topic query and one Scholarly topic query for each laboratory topic still uncovered by the inventory. Stop a topic only after its complete local candidates have been classified and a refinement returns no new relevant normalized ID/title. For every retained ID, use the available download/read operations to retrieve full text or sections before analysis. Deduplicate by normalized paper ID and title. Do not use public curl, fabricate papers/DOIs/venues, or substitute unrelated famous papers.

Traverse the complete task-visible local inventory and include every distinct publication that falls inside the runtime three-year window and is relevant to the laboratory topics/methods visible in the seeded files. The five-entry bibliography requirement is a format floor, not a stopping rule for the literature database. For every selected publication with verified metadata, retrieve/read the available local full text or LaTeX sections and record unavailable PDFs/supplements honestly. The main agent is the sole writer of all local deliverables:

- literature_database.csv with one literal row per selected publication and at least title, author, year, journal, keywords, relevance_to_lab, source, access_date;
- bibliography.bib with one syntactically valid entry for every selected row and at least five entries;
- one manuscript_<stable-id>.docx per selected publication, each with explicit Research Question, Experimental Design, Data Collection, Analytical Methods, Key Findings, Limitations, Reusable Methods or Tools, Relevance to Lab, and Citation Leads sections grounded in the retrieved source;
- research_roadmap.docx of at least 400 characters with explicit Future Directions and Recommendations headings, research gaps/trends, possible collaborators, interdisciplinary opportunities, and citations to selected IDs.

Read the CSV, BibTeX, every manuscript and the roadmap back. Reconcile unique titles/IDs/years, entry/document counts, required headings, access dates, and source/unavailable-material notes before any external write.

Then perform email and calendar writes in one native-parallel response. Paginate Sent mail and read candidates for a message whose subject/body mentions Research Roadmap or Literature Review. If no valid laboratory-member address is exposed in task files or mailbox, use the operational group alias research-lab-members@lab.example.com required to fulfill the task and state that assumption in the body; do not silently skip the email. Send at most once with subject Research Roadmap Update and the verified paper count, gaps and next steps, then read the retained message back. Search for an existing roadmap/priorities meeting; if absent, create one future Research Roadmap Discussion event at a conflict-free visible calendar slot, with priorities and resource allocation in the description, then get_event. Never retry an ambiguous send/create.

Completion requires external read-back of a valid recipient email and retained event plus every local artifact. Do not report a missing optional PDF as completion failure when its unavailability is documented, and do not add synthesis, per-paper, writer or verification waves.
