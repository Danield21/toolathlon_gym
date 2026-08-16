I need a customer loyalty analysis broken down by both segment and region. Pull data from the sales data warehouse showing customer lifetime value patterns across different segment-region combinations.

Create an Excel file called Customer_Loyalty_Report.xlsx with two sheets. The "Customer Loyalty" sheet should have Segment_Region (formatted as "Segment - Region"), Customers count, Avg_LTV (average lifetime value) rounded to 2 decimals, and Total_Orders. Sort by Segment alphabetically then Avg_LTV descending within each segment.

The "Summary" should have Total_Combinations, Highest_LTV_Group (the segment-region combo with highest avg LTV), and Overall_Avg_LTV as a weighted average by customer count. Include a header row with columns `Metric` and `Value`, with one data row per metric below it. Base all metrics on ALL orders in the data warehouse — do not filter orders by status unless explicitly asked.

Send an email to marketing@company.com with subject "Customer Loyalty Analysis" highlighting the highest-value segment-region combinations.
