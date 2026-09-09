You are a business analyst preparing materials for a quarterly product performance review meeting with the sales leadership team.

Your goal is to analyze product sales data from the company database and create a comprehensive set of deliverables for the board presentation.

First, query the product and order data to determine which products generated the most revenue. Look at all orders and join them with product information to calculate total revenue and units sold per product.

Create an Excel file called "Product_Rankings.xlsx" in your workspace with two sheets:

The first sheet should be named "Top Products" and contain columns: Rank, Product_Name, Category, Total_Revenue, Units_Sold, Avg_Order_Value. Include the top 20 products ranked by total revenue in descending order. The Rank column should be numbered 1 through 20. For Product_Name, use the product base name as it appears in the data source; promotional or offer suffixes (such as "with No Cost EMI ...") are not required. Aggregate revenue and units by the product base name (PRODUCT_NAME): several distinct product IDs may share the same base product name, and those must be combined into one row — do not aggregate by PRODUCT_ID.

The second sheet should be named "Category Summary" and contain columns: Category, Product_Count, Total_Revenue, Avg_Revenue_Per_Product. This sheet should group products by category and show aggregate statistics.

Next, create a PowerPoint presentation called "Product_Performance_Review.pptx" in your workspace. The presentation should have at least 4 slides:
The first slide should be a title slide with the text "Q1 2026 Product Performance Review".
The second slide should show the top 10 products by revenue.
The third slide should show a category breakdown of sales.
The fourth slide should be a summary with key metrics such as total revenue across all products, number of products sold, and top performing product.

Finally, create a Google Sheet titled "Product Rankings Dashboard" with a sheet named "Top 20" that contains the same top 20 product data as the Excel file.

To solve this task efficiently, the main agent should first run set-based SQL over all orders, aggregate by PRODUCT_NAME across duplicate product IDs, compute the complete category summary, and freeze the top 20 rows with deterministic ranks and reconciled revenue/units. Then create the three independent deliverables in one wave.

1. Wave 1 — dispatch exactly 3 coder sub-agents in parallel:

   1. One coder sub-agent owns `Product_Rankings.xlsx` and writes the exact Top Products and Category Summary sheets with literal values. It must aggregate by `PRODUCT_NAME` rather than product ID, produce exactly ranks 1–20 in revenue-descending order, and write the complete category totals. Use one bulk write per sheet and one read-back.

   2. One coder sub-agent owns `Product_Performance_Review.pptx` under one presentation ID and creates exactly the requested title, top-10, category-breakdown, and numeric-summary content. Use at most one bounded populate operation per slide, save once, and extract the complete deck text once for validation.

   3. One coder sub-agent owns exactly one `Product Rankings Dashboard`. Search the exact title, create only if absent, ensure one `Top 20` sheet, bulk-write the same frozen rank block, and read it back once.

After Wave 1, the main agent performs one compact cross-output check of the 20 product names/ranks, category totals, and headline metrics and then finishes without another wave or repair agent.
