I need to understand the impact of discounts on our sales. Group orders into the following discount bands and analyze the revenue contribution of each band. Use these inclusive boundaries:
- "No Discount" — DISCOUNT exactly 0
- "1-10%" — DISCOUNT greater than 0 and up to 0.10 inclusive
- "11-20%" — DISCOUNT greater than 0.10 and up to 0.20 inclusive (so an order with DISCOUNT exactly 0.20 belongs here)
- "20%+" — DISCOUNT strictly greater than 0.20

Note: the warehouse currently has no orders with a discount above 0.20, so the "20%+" band will be empty. You may omit the "20%+" row from the report (or include it with zeros); either is acceptable.

Create an Excel file called Sales_Discount_Report.xlsx with two sheets. The "Discount Analysis" should have Discount_Band, Orders count, Revenue rounded to 2 decimals, and Avg_Order_Value rounded to 2 decimals. Sort alphabetically by band name.

The "Summary" should have Total_Orders, Total_Revenue, No_Discount_Revenue, and Discounted_Revenue. Include a header row with columns `Metric` and `Value`, with one data row per metric below it.

Send an email to finance@company.com with subject "Discount Impact Analysis" summarizing the findings.
