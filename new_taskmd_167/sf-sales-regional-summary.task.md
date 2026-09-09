I need a regional sales summary focusing on delivered orders from our data warehouse. Please pull the data and break it down by region.

Create an Excel file called Regional_Sales.xlsx in the workspace with two sheets. The first sheet "Regional Breakdown" should have columns Region, Orders (count of distinct delivered orders), Customers (count of distinct customers with delivered orders), Revenue (total amount rounded to 2 decimals), Avg_Order_Value rounded to 2 decimals, and Revenue_Share_Pct showing each region's share of total revenue rounded to 1 decimal. Sort by Region alphabetically.

The second sheet "Summary" with Metric and Value should include Total_Regions, Total_Revenue rounded to 2 decimals, Total_Orders, Total_Customers, and Top_Region being the region name with highest revenue.

Also please create a Google Sheet spreadsheet titled "Regional Sales Dashboard" with a sheet called "Overview" containing the same Regional Breakdown data so the team can access it online.

To solve this task efficiently, complete it directly in the main agent and dispatch no sub-agent. One delivered-order regional aggregate is compact, and the local and cloud workbooks consume one frozen table.

Discover and verify `SALES_DW.PUBLIC.ORDERS` and `SALES_DW.PUBLIC.CUSTOMERS`. Join on `CUSTOMER_ID`, filter exact `o.STATUS='Delivered'`, and issue one grouped query by nonblank `c.REGION` returning `COUNT(DISTINCT o.ORDER_ID)`, `COUNT(DISTINCT c.CUSTOMER_ID)`, raw `SUM(o.TOTAL_AMOUNT)`, and raw `AVG(o.TOTAL_AMOUNT)`. Run one independent delivered-population control query for distinct orders/customers and revenue. Compute `Revenue_Share_Pct=100*region_revenue/total_revenue` from unrounded values, sort regions alphabetically, and freeze the complete table plus `Total_Regions,Total_Revenue,Total_Orders,Total_Customers,Top_Region`; choose Top_Region by raw revenue with alphabetical ties and round only displayed money to two decimals/share to one decimal.

Resolve the runtime workspace dynamically. The main agent is the sole writer of `Regional_Sales.xlsx`: create exact `Regional Breakdown(Region,Orders,Customers,Revenue,Avg_Order_Value,Revenue_Share_Pct)` and `Summary(Metric,Value)` with all five literal metrics, then read both full ranges back.

In native parallel with that local write, call `search_spreadsheets(query="Regional Sales Dashboard", max_results=100)` and locally exact-match names. Reuse exactly one result or call `create_spreadsheet(title="Regional Sales Dashboard")` exactly once if absent; block on duplicate exact titles or an ambiguous create. Ensure exactly one `Overview` sheet with the same six headers and every frozen region row, write one complete rectangle with `update_cells`, then read `get_spreadsheet_info` and `get_sheet_data` back. Reconcile cloud/local rows and all regional sums to the source controls without a Wave.
