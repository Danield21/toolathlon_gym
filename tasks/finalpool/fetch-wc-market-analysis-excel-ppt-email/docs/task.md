Hey, the business development team needs a competitive market analysis done. We want to see how our online store stacks up against the broader market in each product category.

There is a competitive strategy framework document called Competitive_Strategy.pdf in the workspace that explains how to classify price positioning, compute market share, and identify growth opportunities. Please read through that first so you understand the methodology.

We have a market intelligence portal running at http://localhost:30195/api/market_data.json that has the latest industry benchmarks including average prices, market revenue, and growth rates for each product category. Go grab that data.

Then pull all our product catalog data and order history from the company's online store. We need to know how many products we have in each category, what our average prices look like, and how much revenue we are generating per category from actual orders.

Once you have both datasets, use the terminal to cross-reference everything and compute the competitive metrics described in the PDF. Figure out our price positioning for each category by comparing our average price to the market average. Calculate our market share percentage in each category. Identify which categories represent growth opportunities based on high market growth rate combined with our low current share.

Create an Excel file called Competitive_Analysis.xlsx with three sheets.

The first sheet Category_Comparison should have columns Category, Own_Products, Own_Avg_Price, Market_Avg_Price, Price_Position, Own_Revenue, Market_Revenue, Market_Share_Pct, and Market_Growth_Rate. Sort alphabetically by category. Round prices and revenue to 2 decimal places. Market share to 2 decimal places. Price_Position uses exactly three labels per the PDF methodology: Price Leader when our average price is below the market average, Competitive when our average price is at or within 10 percent above the market average, and Premium when our average price exceeds the market average by more than 10 percent.

The second sheet Strategic_Matrix should have columns Category, Price_Position, Market_Share_Pct, Market_Growth_Rate, Growth_Opportunity, Strategic_Priority, and Recommended_Action. Growth_Opportunity is Yes when market growth rate is above 10 percent and our market share is below 5 percent, otherwise No. The two Strategic_Priority factors are: factor A is Growth_Opportunity being Yes, factor B is Price_Position being either Competitive or Price Leader. Assign High when both factors A and B are true, Medium when exactly one of them is true, and Low when neither is true. Add a recommended action for each category. Sort alphabetically by category.

The third sheet Executive_Summary should have two columns Label and Value with these rows: Total_Own_Products, Total_Own_Revenue, Total_Market_Size, Overall_Market_Share_Pct (total own revenue divided by total market size times 100 rounded to 2 decimals), Categories_Price_Leader (count), Categories_Premium (count), High_Priority_Categories (count), Growth_Opportunities_Count (count).

Build a PowerPoint presentation called Strategy_Presentation.pptx with at least 6 slides covering a title slide, market overview, category comparison highlights, strategic priority matrix, growth opportunities, and strategic recommendations.

Send two emails. First one to ceo@company.com with subject "Competitive Market Analysis - Executive Summary" containing the key findings from the executive summary sheet including our overall market share, number of high priority categories, and growth opportunities count. Second email to product_team@company.com with subject "Category Competitive Analysis - Detailed Findings" that goes through each category with its price position, market share, and strategic priority, highlighting which categories need immediate attention.
