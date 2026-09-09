I am a PhD student conducting a systematic literature review on "prompt engineering for large language models." I need your help to search for relevant papers, analyze them, and compile the results into a structured Google Sheet.

Here is what I need you to do:

First, search the arxiv database for papers related to prompt engineering. I expect there to be around 5 relevant papers about prompting techniques for reasoning and problem-solving with large language models. Ignore any papers that are about word embeddings or word representations (e.g., word2vec, GloVe, distributed representations) since those are from a different era and not related to prompt engineering. Do not include any word-embedding papers in any of the deliverables described below.

Then, for each relevant paper you find, use the scholarly search to look up citation counts and venue information so we have a complete picture of each paper's impact.

Next, use the arxiv-latex manuscript source to read the methodology sections of each paper. I need to understand what prompting technique each paper proposes and how the approaches differ from one another.

As you work through the papers, store notes in the memory system about your progress and key findings. This will help you keep track of which papers you have analyzed and what the main themes are.

Then, create a Google Sheet spreadsheet titled "Prompt Engineering Literature Review" with two sheets. Use the column names exactly as written below.

   The first sheet should be named "Paper Comparison" with the following columns: Paper_ID, Title, Authors, Year, Venue, Citation_Count, Primary_Category, Methodology_Summary. The Methodology_Summary column should contain a brief one to two sentence summary of each paper's prompting approach based on what you read from the manuscript source. There should be one row per relevant paper (approximately 5 rows). Citation_Count should be written as a number (e.g., 6500).

   The second sheet should be named "Technique Analysis" with the following columns: Technique_Name, Paper_ID, Key_Innovation, Reasoning_Type, Requires_Examples. The Reasoning_Type should describe what kind of reasoning the technique targets. You may use one or more of the following categories, separated by commas if you list several: arithmetic, commonsense, symbolic, creative, general. The Requires_Examples column should say Yes or No depending on whether the technique needs few-shot demonstrations. There should be one row per technique.

Finally, save a file called review_summary.txt in your workspace that lists the titles of all relevant papers you found and a one-paragraph synthesis of the overall findings.

Please be thorough in your analysis and make sure the Google Sheet contains accurate information for all the relevant papers you find.

To solve this task efficiently,

This five-paper task is small enough for the main agent to complete directly; do not dispatch any sub-agent. Keep all research operations read-only until the three requested outputs are ready.

First resolve the task-visible workspace directory from the runtime allowed-directory listing (or the literal accessible-workspace directory already supplied by the harness), freeze it as `<workspace_root>`, and reject bare `/workspace` and dumps/run-specific paths. Freeze the two exact schemas, the relevance rule (prompt-engineering methods for LLM reasoning/problem solving), and the exclusion rule (no word embeddings or word representations).

Run the local-arXiv and Scholarly discovery branches in native parallel. For local arXiv, issue complementary prompt-engineering/reasoning queries at the supported per-call bound, union and deduplicate by stable paper ID, and retain title, complete authors, abstract, date/year, and primary category. For Scholarly, issue complementary task-derived queries (at most ten records per query), union/deduplicate by normalized title or stable ID, and retain venue plus numeric citation count. Do not claim pagination that the handlers do not expose.

Cross-reference the two manifests, exclude every embedding-era distractor, and freeze exactly five genuinely relevant distinct paper IDs/titles. Then issue the five independent manuscript reads in native parallel, one exact stable ID per call. For each paper extract a source-backed one-to-two-sentence methodology summary and the technique fields `Technique_Name, Paper_ID, Key_Innovation, Reasoning_Type, Requires_Examples`; allowed reasoning categories are arithmetic, commonsense, symbolic, creative, and general, comma-separated when needed. A missing or empty methodology summary is a blocker, never a blank cell.

Reconcile all sources once and freeze one literal canonical packet with exactly five Paper Comparison rows and exactly five Technique Analysis rows. Each Paper Comparison row contains `Paper_ID, Title, Authors, Year, Venue, Citation_Count, Primary_Category, Methodology_Summary`; citation counts are numbers. No output branch may re-search, reclassify, or recompute these fields.

Using that packet, interleave the three independent output branches whenever their tool dependencies permit:

1. Memory: append one concise source-backed progress note per paper plus one synthesis note without deleting unrelated memory; read back the stored notes and require all five stable IDs.
2. Google Sheets: search exact title `Prompt Engineering Literature Review`; reuse one exact match, create once on zero matches, and block on duplicates. Keep exactly the tabs `Paper Comparison` and `Technique Analysis`. Write only `Paper Comparison!A1:H6` and `Technique Analysis!A1:E6`; never write diagnostic, test, checksum, or status cells outside those ranges. Read those same ranges back and require exact headers, five rows each, numeric citations, nonempty methodology in every row, and no embedding papers.
3. Text file: create `<workspace_root>/review_summary.txt` exactly once with all five exact titles and one coherent synthesis paragraph based on the canonical methodologies; read it back and require complete title coverage.

Accept only after memory, both exact Sheet ranges, and the text file all read back against the same five-paper packet. Never retry an ambiguous create or write.
