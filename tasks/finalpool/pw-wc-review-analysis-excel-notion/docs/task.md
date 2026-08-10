I need to analyze our product reviews by combining external review data with our internal store reviews. There is a review aggregation site at http://localhost:30307 with aggregated ratings from multiple platforms. Please visit and extract all the product review data.

The aggregation table lists exactly 8 products: Wireless Headphones, Smart Watch, Bluetooth Speaker, USB-C Hub, Webcam HD, Mechanical Keyboard, Gaming Mouse, and Laptop Stand. The External_Reviews sheet must contain one row for each of these 8 products.

Then check our online store (WooCommerce) for our internal product reviews, ratings, and sales data, and save them into internal_reviews.json.

Write and run a Python script called review_analyzer.py that reads external_reviews.json and internal_reviews.json (create both), compares ratings, identifies discrepancies, and outputs review_insights.json.

Create an Excel file called Review_Analysis_Report.xlsx with two sheets. The first sheet External_Reviews should have columns Product, External_Rating, Review_Count, Sentiment, and Common_Complaint. The second sheet Review_Summary should have Metric and Value columns with Total_Products_Reviewed, Avg_External_Rating (round to 2 decimals), Positive_Products, and Needs_Attention (count of Negative or Mixed sentiment products). Write all numeric values in the Excel file as literal numbers, not Excel formulas.

Also create a page in the team knowledge base titled "Product Review Tracker" with a summary of products needing attention and their common complaints.
