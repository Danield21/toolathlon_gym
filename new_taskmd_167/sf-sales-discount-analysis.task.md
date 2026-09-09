I need to understand the impact of discounts on our sales. Group orders into the following discount bands and analyze the revenue contribution of each band. Use these inclusive boundaries:
- "No Discount" — DISCOUNT exactly 0
- "1-10%" — DISCOUNT greater than 0 and up to 0.10 inclusive
- "11-20%" — DISCOUNT greater than 0.10 and up to 0.20 inclusive (so an order with DISCOUNT exactly 0.20 belongs here)
- "20%+" — DISCOUNT strictly greater than 0.20

Note: the warehouse currently has no orders with a discount above 0.20, so the "20%+" band will be empty. You may omit the "20%+" row from the report (or include it with zeros); either is acceptable.

Create an Excel file called Sales_Discount_Report.xlsx with two sheets. The "Discount Analysis" should have Discount_Band, Orders count, Revenue rounded to 2 decimals, and Avg_Order_Value rounded to 2 decimals. Sort alphabetically by band name.

The "Summary" should have Total_Orders, Total_Revenue, No_Discount_Revenue, and Discounted_Revenue. Include a header row with columns `Metric` and `Value`, with one data row per metric below it.

Send an email to finance@company.com with subject "Discount Impact Analysis" summarizing the findings.

To solve this task efficiently, complete this compact order aggregation directly in the main agent; do not dispatch any sub-agent. Run one set-based query over all orders with no status filter and assign each row to the literal mutually exclusive bands: `No Discount` for discount equal to 0, `1-10%` for `0 < discount <= 0.10`, `11-20%` for `0.10 < discount <= 0.20`, and `20%+` only for discount strictly above 0.20. Return per-band order count, revenue sum, and average order value plus company-wide total orders/revenue and no-discount/discounted revenue controls. Preserve the empty `20%+` case as either an omitted row or one zero row, sort present band labels alphabetically, and freeze raw totals before rounding.

After the aggregates are frozen, resolve the runtime workspace root and start creation of `Sales_Discount_Report.xlsx` there plus the exact sent-mail search in the same native-parallel response. Create only `Discount Analysis(Discount_Band,Orders,Revenue,Avg_Order_Value)` and `Summary(Metric,Value)` with the four requested metrics; write literal numbers, round money to two decimals, and read both sheets back. Search all pages of Sent mail for recipient `finance@company.com` and subject `Discount Impact Analysis`; send once only if no exact match exists, summarize the verified band contribution, and read the retained message back. Reconcile all band totals to the company-wide controls before completion.
