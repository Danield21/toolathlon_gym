I need help with a wc yf finance analysis. There is external benchmark data available that I need you to fetch the data from http://localhost:30333/api/data.json and extract the relevant metrics.

Then check our online store for current product and order data.

Use the terminal to create and run a Python script called wc_yf_finance_processor.py in the workspace that reads the collected data from JSON files you create, performs the analysis, and outputs wc_yf_finance_results.json.

Create an Excel file called Yf_Financial_Report.xlsx with three sheets. The first sheet Data_Analysis should contain the main comparison data with relevant columns. The second sheet Metrics should summarize key metrics. The third sheet Recommendations should list actionable items.

The Data_Analysis sheet should have columns: Category, Our_Avg_Price, Market_Avg_Price, Price_Gap_Pct. Sort the data alphabetically by Category. The Metrics sheet should have two columns Metric and Value summarizing total counts, averages, and key statistics. The Recommendations sheet should list priority actions based on the gap analysis. Also create a Word document called Yf_Financial_Analysis.docx with an executive summary, key findings, and recommendations sections. Send an email to team-lead@company.com with subject "Analysis Report Complete" summarizing the key findings.

To solve this task efficiently, split the two overlong WooCommerce sources, then preserve the true source-to-processor-to-report dependency. Resolve the accessible workspace dynamically.

In the same orchestration response that launches the two Wave 1 collectors, the main agent fetches and validates exactly `http://localhost:30333/api/data.json` and reads the task-visible `analysis_guide.md` and `Analysis_Guidelines.pdf`. Freeze the literal benchmark rows and require the two visible guides to agree on price, gap, sorting, metric, and recommendation rules; a guide conflict is a blocker. This small independent context branch must not add a separate pre-Wave barrier.

1. Wave 1 — dispatch exactly 2 coder sub-agents in parallel:

   1. The product Coder is strictly read-only toward WooCommerce. Exhaust `woo_products_list` with `perPage<=100` and increasing pages through a short/empty terminal page. Consume any overlong response from its returned task-workspace `output_path` rather than relying on a preview. Persist a runtime-unique JSON containing the exact product ID, name, numeric current `price`, and complete category identities required for the analysis, plus page/row/null/duplicate controls. Return only path, SHA-256 and controls.
   2. The order Coder is strictly read-only toward WooCommerce. Exhaust the required order population with perPage<=100, increasing pages and terminal-page evidence; preserve exact order ID, status and relevant line-item product_id/category/quantity/unit-price fields in a runtime-unique JSON. Return only path, SHA-256, status distribution, page/order/line-item/null/duplicate controls. Do not mutate orders.

After Wave 1, the main agent verifies both hashes and coverage and freezes the benchmark payload plus the two artifact references.

2. Wave 2 — dispatch exactly 1 coder sub-agent:

   1. The processor Coder verifies both source hashes, writes the required canonical JSON input files, creates and runs `wc_yf_finance_processor.py`, and produces `wc_yf_finance_results.json`. For each product-category membership, use that product's literal current `price`; compute each category's `Our_Avg_Price` as the arithmetic mean of its member product prices. Eligible categories are the exact intersection of observed product categories and benchmark categories. Compute `Price_Gap_Pct = (Our_Avg_Price - Market_Avg_Price) / Market_Avg_Price * 100`, where a positive value means our current price is higher, and round only the final outputs according to the visible guides. Use the complete orders only for required coverage and order-derived metrics; never substitute order line-item prices for current catalog prices. Preserve unmatched/null evidence, sort deterministically, read all inputs/script/result back, and return their paths/hashes, execution status, eligible-category coverage, metrics, and schema controls. Do not write Excel/Word or send email.

After Wave 2, the main agent verifies the processor arithmetic and freezes the result path/hash as the sole report source.

3. Wave 3 — dispatch exactly 2 coder sub-agents in parallel:

   1. One Coder solely writes Yf_Financial_Report.xlsx from the verified result: exact Data_Analysis(Category,Our_Avg_Price,Market_Avg_Price,Price_Gap_Pct), Metrics(Metric,Value), and Recommendations(Priority,Action) sheets, literal values, alphabetical category order, and exact-range read-back.
   2. One Coder solely writes Yf_Financial_Analysis.docx from the same result hash with executive summary, key findings and recommendations, then reads its full text/headings back.

In the same response that launches Wave 3, the main agent searches Sent mail for complete To team-lead@company.com / subject Analysis Report Complete, sends only if absent using literal frozen findings, and reads the retained message back. Finally reconcile both reports and email against the identical result hash; do not add a verification wave.
