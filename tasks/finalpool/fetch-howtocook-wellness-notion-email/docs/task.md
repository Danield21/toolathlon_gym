I need help with a recipe wellness analysis. There is external nutrition data available that I need you to fetch from http://localhost:30321/api/data.json and extract the relevant metrics.

Then pick five popular dishes from the recipe collection service and gather their calorie content for the analysis.

Write and run a Python script called cook_wellness_processor.py in the workspace that reads the collected data from JSON files you create, performs the analysis, and outputs cook_wellness_results.json.

Create an Excel file called Wellness_Report.xlsx with three sheets. The first sheet Data_Analysis should have columns Recipe, Calories, and Meets_Guidelines (Yes or No, where Yes means the recipe satisfies the calorie guideline). Include one row per selected recipe. The second sheet Metrics should have two columns Metric and Value with at least these metrics: Total_Recipes, Avg_Calories, Recipes_Meeting_Guidelines, and Avg_Protein. The third sheet Recommendations should have columns Priority and Action and list at least two actionable items based on the analysis.

Send an email to team-lead@company.com with subject "Analysis Report Complete" summarizing the key findings. Create a page in the team knowledge base titled "Cook Wellness Dashboard" with a summary of the analysis.
