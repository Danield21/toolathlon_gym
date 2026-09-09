You are a financial auditor responsible for reconciling sales data between the company's data warehouse and the online store. The company uses two separate systems to track orders. The data warehouse captures all order transactions across multiple regions, while the online store handles e-commerce operations with its own product catalog. Your job is to perform a thorough cross-system audit.

Start by reading the Audit_Procedures.pdf document in your workspace. It outlines the reconciliation methodology and the specific metrics you need to compute.

First, query the data warehouse to get a summary of all orders. You need to group orders by their status (Delivered, Shipped, Processing, Cancelled) and calculate the total count and total revenue for each status. Also compute the overall average order value across all orders.

Next, query the online store system to retrieve all products and their details. Get the list of all orders from the store, including order totals and statuses. Also retrieve customer information from the store. Note: the online store API returns results in pages of 10 by default; increase the page size (for example to 100) or paginate through every page so that you retrieve the complete set of products, orders, and customers rather than only the first 10 of each.

Now write and run a Python script called audit_analysis.py in your workspace. This script should read the data from both systems (you can hardcode the values you gathered) and compute the following: the total order count from the data warehouse, the total order count from the online store, the difference between them, the average order value from each system, and a category-level revenue breakdown from the data warehouse by ship mode (Economy, Express, Next Day, Standard).

Create an Excel workbook called Order_Audit_Report.xlsx in your workspace with four sheets.

The first sheet should be named DW_Summary and contain columns Status, Order_Count, and Total_Revenue. Include one row for each order status from the data warehouse: Cancelled, Delivered, Processing, and Shipped. Add a totals row at the bottom.

The second sheet should be named Store_Summary and contain columns Metric and Value. Include rows for Total Products (82 products in the store), Total Orders, Total Customers, and Average Order Value from the online store.

The third sheet should be named ShipMode_Breakdown and contain columns Ship_Mode, Order_Count, and Total_Revenue. Include rows for Economy, Express, Next Day, and Standard shipping modes from the data warehouse.

The fourth sheet should be named Reconciliation and contain columns Metric, DW_Value, Store_Value, and Difference. Include rows comparing Total Order Count, Average Order Value, and Total Revenue between the two systems. The Difference column should be the data warehouse value minus the store value. For the store side, Total Revenue is the sum of the order `total` field across all store orders and Average Order Value is computed across all store orders (every status), consistent with the data warehouse side, which sums revenue across all statuses including Cancelled.

Finally, create a Word document called Audit_Findings.docx in your workspace. The document should have a title "Order Reconciliation Audit Report". Include a section discussing the data warehouse summary with order counts and revenue by status. Include another section on the online store summary. Add a section on the ship mode breakdown showing how revenue distributes across shipping methods. End with a reconciliation findings section that highlights the key differences between the two systems and provides recommendations for alignment.

To solve this task efficiently, preserve the two real dependency stages but keep both stages compact. The main agent first reads `Audit_Procedures.pdf` and freezes the all-status reconciliation rules.

1. Wave 1 — dispatch exactly 2 explore sub-agents in parallel:

   1. One explore sub-agent runs set-based Snowflake queries that return only the four literal status rows `Cancelled`, `Delivered`, `Processing`, and `Shipped`, the overall count/revenue/average, and the four literal ship-mode rows `Economy`, `Express`, `Next Day`, and `Standard`, with total controls.
   2. One explore sub-agent fully paginates WooCommerce products, orders, and customers at `perPage=100`. Omit the optional order status argument to request all observed statuses and compute locally only total products, total orders, total customers, all-status revenue, and all-status average order value plus page/count controls. Do not persist raw rows or modify the store.

After Wave 1, the main agent checks both compact tables, writes and runs `audit_analysis.py`, and freezes the four workbook tables.

2. Wave 2 — dispatch exactly 2 coder sub-agents in parallel:

   1. The Excel coder solely creates `Order_Audit_Report.xlsx`, using one bulk write for each of the exact four sheets and literal values, then one full workbook read-back.
   2. The Word coder solely creates `Audit_Findings.docx` with one bulk structured body and one read-back. It must contain the exact title `Order Reconciliation Audit Report`, the four required sections plus recommendations, and must literally spell all four statuses `Delivered`, `Cancelled`, `Processing`, `Shipped` and all four ship modes `Economy`, `Express`, `Next Day`, `Standard` in the document text.

After Wave 2, the main agent performs one compact totals/status/mode consistency check and finishes; do not dispatch repair or verification agents.
