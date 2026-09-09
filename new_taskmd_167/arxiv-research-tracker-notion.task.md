I am an NLP research group lead and I need to set up a research tracker for recent papers on efficient transformer architectures, including topics like flash attention, token merging and pruning, low-rank adaptation, quantization-aware training, and sparse mixture of experts. The database contains exactly five papers relevant to these topics, along with some unrelated papers on other subjects. Include exactly the five relevant papers and do not add any others.

For each paper, please search the arxiv database for papers related to efficient transformers and model compression. Look for papers whose titles or abstracts mention topics such as attention efficiency, token merging, LoRA, quantization, or mixture of experts in the context of transformers. The five relevant papers cover exactly these topics: (1) attention efficiency / flash attention, (2) token merging and pruning, (3) low-rank adaptation (LoRA), (4) quantization-aware training, and (5) sparse mixture of experts. Exclude papers that do not fit this efficient-transformer theme, such as robotics or drone navigation and molecular property prediction papers, even if they appear in the search results. Then for each paper you find, look up its citation count and publication venue information through a scholarly search. Cross-reference by title to get accurate citation metrics. Note: the arXiv search tools do not return citation counts or publication venues, so you must obtain those from the scholarly search tools, matching each paper by its title.

Once you have gathered all the information, create an Excel file called Transformer_Research_Tracker.xlsx in the workspace with two sheets.

The first sheet should be named "Paper Comparison" and should have the following columns: Paper_ID, Title, Authors, Published_Date, Venue, Citation_Count, Primary_Category, Key_Contribution. Each row should represent one of the relevant papers. The Key_Contribution column should contain a brief one-sentence summary of each paper's main contribution. Sort the papers by Citation_Count in descending order.

The second sheet should be named "Statistics" and should have two columns: Metric and Value. Include the following metrics: Total_Papers (the number of papers tracked), Avg_Citations (the average citation count across all tracked papers rounded to one decimal place), Most_Cited_Paper (the title of the paper with the highest citation count), Top_Venue (the venue that appears most frequently among the tracked papers), and Date_Range (the range of publication dates in the format "YYYY-MM-DD to YYYY-MM-DD").

Finally, create a page in the team knowledge base titled "Efficient Transformers Research Tracker" that provides a structured overview of the research landscape. The page should include a heading or introduction summarizing the research area, a section listing each paper with its title, authors, citation count, and a brief description, a section with key observations or trends across the papers, and information about the most impactful papers and common themes.

Please complete all steps and save the Excel file to the workspace. Make sure the knowledge base page has substantive content that would be useful for tracking research progress.

To solve this task efficiently, follow this exact orchestration plan. A wave must finish and return all stated handoffs before the next wave begins. Within a wave, dispatch exactly one sub-agent for each numbered sub-task and run those sub-agents concurrently.

Before Wave 1, the main agent must freeze exactly the five task-defined topic slots—flash/efficient attention, token merging or pruning, LoRA, quantization-aware training, and sparse mixture of experts—and the exclusion rule for robotics, navigation, molecules, and other unrelated work. It must assign one topic per agent without assuming paper identities or source values.

1. Wave 1 — dispatch exactly 5 sub-agents in parallel, one for each numbered sub-task:

   1. Use one explore sub-agent to research attention efficiency or flash attention.
   2. Use one explore sub-agent to research token merging and pruning.
   3. Use one explore sub-agent to research low-rank adaptation.
   4. Use one explore sub-agent to research quantization-aware training.
   5. Use one explore sub-agent to research sparse mixture of experts.

Each Wave 1 sub-agent must search arXiv and Scholarly read-only for only its assigned topic, validate relevance from title/abstract, cross-reference venue and numeric citation count by title, and return the strongest source-backed paper matching that slot as one compact record containing stable ID, title, full authors, published date, venue, numeric citation count, primary category, one-sentence contribution, search controls, provenance, and missing/conflict status. Explore is read-only, so do not write files or create Excel or Notion state, and do not retain an unrelated paper merely to fill a slot.

After Wave 1, the main agent must verify exactly one distinct relevant paper per supported topic, resolve duplicates and source conflicts, sort the frozen five-paper set by citation count, and compute literal statistics only from that set.

2. Wave 2 — dispatch exactly 2 sub-agents in parallel, one for each numbered sub-task:

   1. Use one coder sub-agent to own Transformer_Research_Tracker.xlsx exclusively. Create exactly Paper Comparison and Statistics with the exact headers/metrics, five citation-descending rows, literal numeric counts and one-decimal average, correctly resolved most-cited paper, runtime top venue, and date range. Read back both sheets and return path, row/metric controls, sort order, and validation. Do not modify Notion.

   2. Use one coder sub-agent to own the single Notion page exclusively. Search for Efficient Transformers Research Tracker, create at most one page, and add a substantive overview, one source-backed section per frozen paper with authors/citations/contribution, cross-paper trends, impactful papers, and common themes. Read it back and return page ID/title, five-paper/section coverage, and validation. Do not write the workbook.

After Wave 2, the main agent must reconcile all five IDs, citations, venues, contributions, statistics, and exclusions across workbook and page and verify both external state and file integrity before completion.
