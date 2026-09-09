Our company is organizing a team-building cooking workshop next month and I need you to prepare the materials. We want to feature three Chinese dishes that are relatively easy to make (rated as easy, i.e. a low difficulty score, in the recipe database), and cover different categories so we get some variety. Please search through the available recipe database to find three suitable dishes that meet these criteria.

Note: write all document content in English (dish names, ingredients, cooking steps, tips, and the shopping list).

Once you have selected the three recipes, I need you to create three deliverables in the workspace.

First, create a Word document called Workshop_Handbook.docx. This should serve as a detailed step-by-step cooking guide for participants. Start with a title "Cooking Workshop Handbook" as a top-level heading. Then for each of the three dishes, include a second-level heading with the dish name, followed by a paragraph listing all the required ingredients with quantities, and then another paragraph (or multiple paragraphs) walking through the cooking steps in order. At the end of the document, add a section called "Tips and Notes" with any general cooking advice that would help beginners.

Second, create a PowerPoint presentation called Workshop_Slides.pptx for the instructor to use during the workshop. The first slide should be a title slide with "Team Cooking Workshop" as the title. Then for each dish, create two slides: one slide with the dish name as the title and the ingredient list as the content, and a second slide with the dish name followed by "Steps" as the title and the cooking steps as bullet points. Finally add a closing slide with "Enjoy Your Meal!" as the title.

Third, create a PDF file called Shopping_List.pdf in the workspace. This should be a consolidated shopping list that combines all ingredients from all three dishes, organized neatly so that someone from our procurement team can use it to go shopping. Group similar ingredients together if possible and include the total quantities needed. The document should have a clear title "Shopping List for Cooking Workshop" at the top.

To solve this task efficiently, the main agent must first query the exact `早餐`, `素菜`, and `主食` category listings in one native-parallel response, select one low-difficulty candidate ID from each category, and retrieve the three exact-ID recipe details in a second native-parallel response. Require three distinct IDs/categories and source-supported low difficulty, then freeze each recipe's English name, full translated ingredients with quantities, ordered translated steps, and beginner tips. The three independent deliverables can then be created in one wave.

1. Wave 1 — dispatch exactly 3 coder sub-agents in parallel:

   1. One coder sub-agent owns Workshop_Handbook.docx and writes the exact title, one complete section for each of the three frozen dishes, ordered ingredients/steps, and Tips and Notes. Read it back and return its path and three-dish coverage checks.

   2. One coder sub-agent owns Workshop_Slides.pptx and creates exactly the title slide, two slides for each of the three frozen dishes, and the Enjoy Your Meal closing slide. Read it back and return its path, presentation ID, eight-slide count, titles, and dish coverage.

   3. One coder sub-agent owns Shopping_List.pdf and consolidates all ingredients from the same three frozen recipes in English with combined quantities and clear grouping. Read it back and return its path, title, ingredient coverage, and quantity controls.

After Wave 1, the main agent must cross-check the same three dishes and ingredients across all deliverables and finish without another wave.
