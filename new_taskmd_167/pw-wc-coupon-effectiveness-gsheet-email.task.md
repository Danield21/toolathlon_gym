I need help with a wc coupons analysis. There is external benchmark data available that I need you to visit http://localhost:30309 and extract the relevant metrics.

Then check our online store for current product and order data.

Use the terminal to create and run a Python script called wc_coupons_processor.py in the workspace that reads the collected data from JSON files you create, performs the analysis, and outputs wc_coupons_results.json.

Create an Excel file called Coupon_Effectiveness_Report.xlsx with three sheets. The first sheet Data_Analysis should contain the main comparison data with relevant columns. The second sheet Metrics should summarize key metrics. The third sheet Recommendations should list actionable items.

The Data_Analysis sheet should have columns: Category, Product_Count, Our_Avg_Price, Total_Sales, Market_Avg_Price, Price_Gap_Pct. Sort the data alphabetically by Category. The Metrics sheet should have two columns Metric and Value summarizing total counts, averages, and key statistics. The Recommendations sheet should list priority actions based on the gap analysis. Send an email to team-lead@company.com with subject "Analysis Report Complete" summarizing the key findings. Create a Google Sheet called "Wc Coupons Tracker" with the key data points.

To solve this task efficiently, preserve the real source-to-result dependency, but do not make one sub-agent perform the web extraction, two complete store scans, Python processing, and all verification serially.

At the start, the main agent navigates exactly once to `http://localhost:30309` and takes one browser snapshot. It extracts only the benchmark category metrics required by the task and freezes `{category,market_avg_price}` with row/missing-field controls. It must not install a browser, try alternate URLs, or use WooCommerce while extracting the page.

1. Wave 1 — dispatch exactly 1 coder sub-agent:

   1. The Coder is strictly read-only toward WooCommerce. Exhaust the complete product catalog and required order population with `perPage` at most 100, increasing pages, and terminal-page controls; omit an optional status filter for the unfiltered order collection and filter observed statuses locally. Persist exact unique product/category records with all three literal price fields `price`, `regular_price`, and `sale_price`, plus exact order/line-item product IDs, quantities, totals, and statuses, in runtime-unique JSON files inside the task workspace. Return only `{product_path,product_sha256,order_path,order_sha256,product_pages,product_rows,order_pages,order_rows,line_item_rows,duplicate_product_ids,duplicate_order_ids,null_controls,status_distribution}`. Do not use Playwright, run the processor, create a report, send email, or mutate the store.

After Wave 1, the main agent verifies both hashes and coverage. Using only the frozen benchmark rows and the two verified Woo files, it creates the required JSON inputs and `wc_coupons_processor.py`. The script must define each product's effective current price as numeric nonempty `sale_price`, otherwise numeric nonempty `regular_price`, otherwise numeric `price`; `Our_Avg_Price` is computed from that field exactly once. Run the script once, read `wc_coupons_results.json` back, and verify exact alphabetical rows with `Category, Product_Count, Our_Avg_Price, Total_Sales, Market_Avg_Price, Price_Gap_Pct`, plus metrics and source-supported recommendations. Freeze the result path/hash as the only report source. In the same response as Wave 2, start the nonduplicate `Analysis Report Complete` email branch using only those verified findings.

2. Wave 2 — dispatch exactly 2 coder sub-agents in parallel:

   1. One coder sub-agent is the sole writer of `Coupon_Effectiveness_Report.xlsx`. Read only the verified result. Create exactly `Data_Analysis`, `Metrics`, and `Recommendations`: `Data_Analysis` has literal columns `Category, Product_Count, Our_Avg_Price, Total_Sales, Market_Avg_Price, Price_Gap_Pct` sorted alphabetically by Category; `Metrics` has `Metric, Value`; `Recommendations` contains source-supported priority actions. Write literal values and read back every populated range.
   2. One coder sub-agent is the sole owner of the Google Sheet `Wc Coupons Tracker`. Read only the verified result, exhaust `search_spreadsheets` for the exact title, create it once only when absent, and block on multiple exact matches. Write the exact alphabetical category rows with `Category, Product_Count, Our_Avg_Price, Total_Sales, Market_Avg_Price, Price_Gap_Pct` plus a compact verified metrics section, then read both ranges back.

After Wave 2, the main agent reconciles the Excel and Google Sheet to the same result hash and verifies exactly one email to `team-lead@company.com` with subject `Analysis Report Complete`. Do not dispatch more sub-agents.
