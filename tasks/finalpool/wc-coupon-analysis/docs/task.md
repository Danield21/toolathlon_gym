I need to analyze our coupon usage in the shop. Pull all coupons and their usage statistics.

Create an Excel file called WC_Coupon_Report.xlsx with two sheets. The "Coupon Analysis" should have Code, Discount_Type, Amount, Usage_Count (use each coupon's stored usage_count field from the coupon list — not order-level coupon line counts), Usage_Limit (leave blank when the coupon has no usage limit), and Utilization_Pct which is usage count divided by usage limit times 100 rounded to 1 decimal (use 0 when the coupon has no usage limit). Sort by Usage_Count descending.

The "Summary" should have Total_Coupons, Total_Usage across all coupons, Most_Used_Code, and Avg_Utilization which is the average of Utilization_Pct across only those coupons that have a Usage_Limit (ignore coupons with no limit), rounded to 1 decimal.

Create a Word document called Coupon_Strategy.docx with recommendations based on the analysis.
