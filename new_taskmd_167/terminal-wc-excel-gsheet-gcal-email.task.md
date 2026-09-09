You are an inventory manager responsible for monitoring product stock levels and coordinating restocking activities for an online store. Your goal is to analyze the current product inventory, forecast demand, generate reports, and ensure timely restock scheduling.

Start by reading the Inventory_Policy.pdf in your workspace, which describes the reorder thresholds and lead time expectations. Also review the warehouse_config.json file that contains warehouse capacity limits and supplier contact information.

Query the e-commerce platform to retrieve the complete product catalog. For each product, gather the product name, category, regular price, stock quantity, stock status, and total sales count. You will use this data to assess inventory health and compute reorder priorities.

Write a Python script called demand_forecast.py in your workspace and execute it using command-line tools. The script should compute a simple demand velocity metric for each product by dividing total sales by an assumed 180-day selling period to get daily sales rate. Then calculate days of supply by dividing current stock quantity by the daily sales rate (use 999 for products with zero sales rate). Determine a reorder point for each product as 14 times the daily sales rate (representing a 14-day lead time buffer). Flag products as Critical urgency if their current stock is at or below the reorder point and they are not out of stock, flag out-of-stock products as Out_of_Stock urgency, and flag all others as Normal.

Create an Excel workbook called Inventory_Lifecycle_Report.xlsx in your workspace with four sheets.

The first sheet should be named Product_Inventory and contain columns for product_name, category, price, stock_qty, stock_status, total_sales, and days_of_supply. Include one row for every product in the catalog, sorted by days_of_supply ascending so the most urgent items appear first.

The second sheet should be named Reorder_Alerts and contain columns for product_name, current_stock, reorder_point, and urgency. Include only the products that have Critical or Out_of_Stock urgency, sorted by urgency with Out_of_Stock items first.

The third sheet should be named Category_Summary and contain columns for category, product_count, avg_stock, and total_value. Total value should be the sum of price times stock quantity for each product in the category. Include one row per product category.

The fourth sheet should be named Restock_Schedule and contain columns for product_name, reorder_date, quantity, and supplier. For each Critical or Out_of_Stock product, set the reorder date to March 10, 2026. The quantity to order should be the reorder point minus current stock (minimum 1 unit). The supplier should be "Primary Supplier" for all items.

Publish the Category_Summary data to the shared spreadsheet system so the team can view real-time inventory status. Create a spreadsheet titled "Inventory Dashboard" with a sheet containing the category summary data.

Schedule a restock review meeting on the shared calendar for March 11, 2026 at 10:00 AM lasting one hour. The event summary should be "Inventory Restock Review Meeting" and the description should mention the number of critical items identified.

Send an email to purchasing@company.com with the subject "Critical Inventory Alert - Restock Required" summarizing the number of out-of-stock and critical products, along with the top three most urgent items that need immediate restocking.
Note: Use the America/New_York timezone for all calendar events in this task.


To solve this task efficiently, the main agent completes this bounded 82-product inventory task directly; do not dispatch sub-agents. Dynamically resolve the workspace and read Inventory_Policy.pdf and warehouse_config.json.

Retrieve the complete WooCommerce catalog with woo_products_list(page=1,perPage=100), continuing only when pagination requires it, and retrieve complete categories. Require 82 unique live product IDs, preserve every category membership plus literal regular_price, stock_quantity, stock_status and total_sales, and use full_output_path if the response is oversized. Write and run demand_forecast.py: daily_sales=total_sales/180; days_supply=stock/daily_sales or 999; reorder_point=14*daily_sales; urgency Out_of_Stock first, then Critical when stock<=reorder_point, else Normal. Compute category summaries, regular_price*stock inventory value, top-three urgent rows and March 10 restock rows. Read back script/output and freeze all 82 product rows with source/formula controls.

The main agent then performs four independent output branches directly, batching ready calls:

- sole-write Inventory_Lifecycle_Report.xlsx with exact four sheets, all 82 products, literal values, required sorts/date/quantity rule and full range read-back;
- fully paginate exact-title search for cloud spreadsheet Inventory Dashboard, create/reuse at most once, block duplicate titles/tabs, own exactly Category_Summary, batch write and read back;
- search exact Inventory Restock Review Meeting on 2026-03-11 10:00-11:00 America/New_York, create at most once with critical-item count, and get_event;
- paginate Sent search/read candidates for purchasing@company.com and exact subject Critical Inventory Alert - Restock Required, send at most once with out-of-stock/critical counts and top three complete item names, then read back.

Final acceptance compares external workbook, cloud sheet, event and email read-backs to the 82-row canonical output. No source, writer, calendar/email, auditor or repair agents are allowed.
