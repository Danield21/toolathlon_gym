I am organizing my research on federated learning and need to set up a knowledge base. Please search for papers about federated learning from the academic literature repository and organize them in two ways.

First, create a page in the team knowledge base titled "Federated Learning Research Hub". Inside this page, create a child database called "Paper Index" that contains an entry for each federated learning paper. Each entry should include the paper title, authors, and a short summary.

Second, create an Excel file called "Federated_Learning_Papers.xlsx" in my workspace. It should have two sheets. The first sheet "Paper Details" should have columns Paper_ID, Title, Authors, Year, and Abstract_Length (the character count of the abstract). The second sheet "Summary" should have columns Metric and Value, listing Total_Papers (the number of papers found), Avg_Abstract_Length (the average character count of all abstracts rounded to the nearest integer), Earliest_Year, and Latest_Year.

Please include all papers that are related to federated learning, including foundational papers about federated approaches and surveys of the field.

Note on the spreadsheet: write all numeric values (Abstract_Length and every Summary metric) as literal numbers in the cells. Do not rely on Excel formulas for these values, since formulas are not computed by the reading tool. Compute the values yourself (e.g. with your code tool) and fill in the resulting numbers.

To solve this task efficiently, follow this exact orchestration plan. A wave must finish and return all stated handoffs before the next wave begins. Within a wave, dispatch exactly one sub-agent for each numbered sub-task and run those sub-agents concurrently.

1. Wave 1 — dispatch exactly 1 sub-agent for the following numbered sub-task:

   1. Use one explore sub-agent to own complete read-only paper discovery. Issue complementary task-derived federated-learning queries to the local arXiv search at its supported per-call result bound, union and deduplicate the returned records by stable ID, and retain every genuinely relevant result including foundational and survey work. Because the search exposes no cursor or list-all operation, do not claim coverage beyond this observed query union. Preserve full title, complete authors, abstract, published date/year, and abstract character count, and return the normalized union directly with per-query/union/retained counts, IDs, year range, and abstract-length controls. Explore is read-only, so do not write files or create Notion or Excel state.

After Wave 1, the main agent must verify the complementary-query controls and deduplicated observed union, determine the related-paper count only from returned records, and freeze literal summary metrics and one title/author/summary record per retained paper using only task-visible runtime evidence.

2. Wave 2 — dispatch exactly 2 sub-agents in parallel, one for each numbered sub-task:

   1. Use one coder sub-agent to own the Notion hierarchy exclusively. Search for an existing Federated Learning Research Hub, create at most one parent page, then create exactly one Paper Index child database under that returned parent and add one entry per frozen paper with title, authors, and a source-faithful short summary. Keep the whole create-parent → create-database → populate → read-back sequence in this agent. Return page/database IDs, entry count, covered IDs/titles, and validation. Do not write Excel.

   2. Use one coder sub-agent to own Federated_Learning_Papers.xlsx exclusively. Create exactly Paper Details and Summary with the exact required headers, one row per frozen paper, literal numeric abstract lengths and metrics, and the rounded average, earliest year, and latest year. Read back both sheets and return path, row counts, summary values, and header/value validation. Do not modify Notion.

After Wave 2, the main agent must reconcile paper IDs, counts, authors, and summary metrics across the Notion database and workbook, verify both external state and file integrity, and then complete without adding unrelated papers or extra deliverables.
