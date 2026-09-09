The product team wants to analyze customer review quality across product categories to prioritize product improvement efforts. Query WooCommerce product review data and link it to product categories to understand how different categories perform with customers.

Create an Excel file called Product_Review_Analysis.xlsx with two sheets. The first sheet should be called "Category Analysis" with columns: Category, Product_Count, Review_Count, Avg_Rating (rounded to 2 decimal places), Five_Star_Count, and Five_Star_Rate (the percentage of reviews that are 5 stars, rounded to 1 decimal). Sort by Category alphabetically. The second sheet should be called "Top Products" and list the top 10 highest-rated products that have at least 3 reviews, with columns: Product_Name (truncated to 50 characters if needed), Category, Review_Count, and Avg_Rating (rounded to 2 decimal places). Sort by Avg_Rating descending, then Review_Count descending.

Create a Notion database called "Product Review Insights" with properties for Category (title), Product_Count (number), Review_Count (number), Avg_Rating (number), and Five_Star_Rate (number). Populate it with one entry per category from the Category Analysis sheet.

Email the analysis to product_team@store.com with subject "Product Review Analysis Report" and include a summary of which categories have the highest and lowest ratings in the email body.

To solve this task efficiently, use two genuine dependency waves with exactly three sub-agents total. The first agent computes the complete store join once; the two artifact writers then run in parallel from one canonical category table.

1. Wave 1 — dispatch exactly 1 Explore sub-agent:

   1. Use one Explore sub-agent to retrieve the complete product and review endpoints read-only using `perPage=100` and increasing pages until each has a short or empty terminal page. Omit the optional review `status` argument, preserve every observed status, and filter only as the original task requires. Join strictly by Product_ID; a product contributes to each observed category exactly according to its source categories. Compute the compact per-category Product_Count, Review_Count, rating sum/average, Five_Star_Count/Rate, and the eligible products with at least three reviews. Rank Top Products by average rating descending, review count descending, then stable Product_ID. Return the full compact category and top-ten tables directly, plus endpoint page/row/status controls, join/category coverage, and duplicate/tie checks. Do not modify the store or create deliverables.

After Wave 1, the main agent independently checks the counts, rates, rankings, and tie order, then freezes one immutable category/top-product handoff for both writers.

2. Wave 2 — dispatch exactly 2 Coder sub-agents in parallel:

   1. Use one Coder sub-agent as the sole writer of `Product_Review_Analysis.xlsx`. Create the exact two sheets once from the frozen tables, with exact headers, truncation, sorting, row limits, rounding, and literal values. Read both sheets back and return the path plus complete row/value controls.

   2. Use one Coder sub-agent as the sole owner of the exact-title `Product Review Insights` database. Search exact title first; create or complete one nonduplicate database with the exact property types, then upsert exactly one row per frozen category. Retrieve the database/schema/rows and return its ID plus source-to-row and duplicate checks. Do not write the workbook.

While Wave 2 runs, the main agent may prepare—but must not send—the email body from the frozen highest/lowest category values. After both writers return, reconcile every category across source, workbook, and database; then check Sent mail and send only one missing exact-recipient/subject message with those verified categories. Read back all three outputs. Repairs belong only to the original writer, and there is no separate validation wave.
