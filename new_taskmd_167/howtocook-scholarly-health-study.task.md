I am a wellness researcher studying the nutritional value of traditional Chinese cooking. A nutrition reference file called nutrition_reference.json is available in the workspace, containing calorie and macronutrient data for common Chinese cooking ingredients.

Please browse the recipe database and select 5 dishes from at least 3 different categories. Then search for academic papers related to Chinese diet and health or traditional Asian cuisine nutrition. I expect to find about 4 relevant papers. For each selected dish, estimate its nutritional content using the ingredient data from the reference file.

Create an Excel spreadsheet called Health_Diet_Analysis.xlsx with the following three sheets. The first sheet should be named "Recipe Nutrition" with columns Dish_Name, Category, Estimated_Calories, Protein_g, Fat_g, Carbs_g, and Health_Rating. The Health_Rating should be Low, Medium, or High based on the calorie density of the dish. The second sheet should be named "Research Summary" with columns Paper_Title, Authors, Year, Citation_Count, and Key_Finding. The third sheet should be named "Combined Analysis" with columns Dish_Name, Health_Rating, and Supporting_Research, linking the dishes to relevant research findings.

Also create a Word document called Chinese_Cuisine_Health_Report.docx. The document should have the title "Nutritional Analysis of Traditional Chinese Cuisine" and include the following sections: Introduction, Recipe Analysis, Literature Review, and Conclusions.

Notes on expected output format:
- All numeric values in the spreadsheet (Estimated_Calories, Protein_g, Fat_g, Carbs_g, Citation_Count, Year) must be written as literal numbers, not as Excel formulas.
- "About 4 papers" is flexible: a Research Summary with 3 to 5 papers is acceptable. Only include papers that are genuinely related to Chinese diet and health or Asian cuisine nutrition.
- Assign Health_Rating as Low, Medium, or High based on the calorie density of each dish (e.g., Low for low-calorie dishes, Medium for moderate, High for calorie-dense dishes).

To solve this task efficiently, run the two genuinely independent research branches in one wave. Before dispatch, the main agent must read nutrition_reference.json and freeze its ingredient schema.

1. Wave 1 — dispatch exactly 2 explore sub-agents in parallel:

   1. One explore sub-agent selects exactly five suitable HowToCook dishes from at least three categories, retrieves their ingredients, and estimates calories, protein, fat, and carbohydrates from the provided nutrition reference. Return one compact five-row table with calculation notes and category/count controls; do not write artifacts.

   2. One explore sub-agent finds 3-5 genuinely relevant papers on Chinese diet and health or Asian cuisine nutrition, reads the available abstracts/full records, and returns exact title, authors, year, citation count, key finding, identifier, and relevance evidence for each paper; do not write artifacts.

After Wave 1, the main agent must reconcile both handoffs, assign supported Low/Medium/High calorie-density ratings, and directly create Health_Diet_Analysis.xlsx and Chinese_Cuisine_Health_Report.docx with the exact sheets, headings, literal numeric values, and cross-links. Read both files back; do not dispatch writer or verification agents.
