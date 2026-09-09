I need to forecast our inventory needs and plan restocking. There is a supplier catalog API at http://localhost:30306/api/supplier_catalog.json with lead times and minimum order quantities for our suppliers. Please fetch that data.

Then check our online store for current product stock levels, sales figures, and product details.

Use the terminal to create and run a Python script called inventory_forecaster.py in the workspace that reads supplier_data.json and product_stock.json (create both), calculates daily sales rates and days of remaining stock for each product, and outputs restock_plan.json. Daily rate is Total_Sales divided by 90 (assuming 90-day sales window). Days remaining is Current_Stock divided by daily rate. A product needs restock if days remaining is less than 30.

Create an Excel file called Inventory_Forecast_Report.xlsx with three sheets. The first sheet Stock_Status should have columns Product, Current_Stock, Total_Sales, Daily_Rate (round to 2 decimals), Days_Remaining (round to 1 decimal), and Needs_Restock ("Yes" if under 30 days, "No" otherwise). Sort by Product name alphabetically.

The second sheet Supplier_Info should list each supplier with columns Supplier, Lead_Time_Days, Min_Order_Qty, and Reliability_Score.

The third sheet Restock_Summary should have Metric and Value columns with Total_Products_Analyzed, Products_Need_Restock, Products_Healthy, Avg_Days_Remaining (round to 1 decimal).

Schedule a calendar event "Inventory Review Meeting" on March 12, 2026 from 9:00 AM to 10:00 AM UTC with a description listing products that need restocking. Also send an email to procurement@company.com with subject "Urgent Restock Alert" listing the products that need immediate restocking.

To solve this task efficiently, use exactly one source wave. The WooCommerce owner is a read-only Coder because the complete product records must cross the Wave boundary through a lossless file rather than a long natural-language handoff.

1. Wave 1 — dispatch exactly 2 sub-agents in parallel:

   1. One explore sub-agent fetches `http://localhost:30306/api/supplier_catalog.json` exactly once and returns every supplier row with field and count controls.
   2. One coder sub-agent, strictly read-only with respect to WooCommerce, calls `woo_products_list(page=1,perPage=20)` and increments `page` until the returned terminal condition/source exhaustion. Project every observed product to exact `{id,name,stock_quantity,total_sales}` without summarizing or rewriting names; reject duplicate IDs and preserve nulls explicitly. Write the complete ordered array once to a main-assigned runtime-unique workspace JSON path, read it back, and return only `{path,sha256,page_count,row_count,unique_id_count,null_counts,terminal_page}`. Do not return the product rows in chat and do not modify the store.

After Wave 1, the main agent verifies the supplier controls and the Coder JSON path/hash/count controls, then materializes `supplier_data.json` and `product_stock.json` from those exact sources without re-keying product names. Write and run `inventory_forecaster.py` once and verify `restock_plan.json`. Use `daily_rate = total_sales / 90`; for zero sales use a non-restock infinite/explicit sentinel consistently rather than dividing by zero; otherwise use `days_remaining = stock / daily_rate` and `Needs_Restock = Yes` iff days remaining is below 30. Create `Inventory_Forecast_Report.xlsx` directly with the exact three sheets, headers, product sorting, supplier rows, four Summary metrics, literal values, and requested rounding. In parallel with the final workbook branch, search/create/read back the fixed `Inventory Review Meeting` at 2026-03-12 09:00–10:00 UTC; after the verified restock list exists, send exactly one `Urgent Restock Alert` email to `procurement@company.com`. Read each final artifact once and do not dispatch another wave.
