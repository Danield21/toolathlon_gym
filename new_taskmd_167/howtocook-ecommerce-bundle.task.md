We run an online store that sells kitchen appliances and we want to create recipe-based product bundles to boost sales. The idea is to pair our appliance products with popular recipes that would benefit from those appliances, then package them as promotional bundles.

Here is what I need you to do.

First, look up our online store's product catalog and pull out all the products in the Home Appliances category. I need each product's name, regular price, and stock quantity.

Second, search the recipe database for recipes in three different categories: one recipe from the drinks category, one from the staple food category, and one from the meat dishes category. For each recipe, get the full details including name, ingredients list, and difficulty level. Pick recipes that would naturally pair with a kitchen appliance (for example, a drink recipe that benefits from a blender, a staple food that uses a cooker, or a meat dish that works well with a sealing or steaming device). Use your judgment on which recipes make good pairings.

Third, check the latest gold futures price trend since gold prices affect our electronics component sourcing costs. Get the closing prices for the last 5 trading days using the gold futures symbol.

Once you have all the data, create an Excel file called Bundle_Pricing.xlsx in the workspace with three sheets.

The first sheet should be called Store Products. It should list all Home Appliances category products with these columns: Product_Name, Price, Stock_Quantity. Sort the rows by Price in ascending order.

The second sheet should be called Gold Trend. It should have these columns: Date, Close_Price. Include the last 5 trading days of gold futures closing prices, sorted by date in descending order (most recent first). The Close_Price values should be rounded to 2 decimal places.

The third sheet should be called Recipe Bundles. It should have these columns: Recipe_Name, Recipe_Category, Difficulty, Paired_Product, Bundle_Price. Each row represents one bundle pairing a recipe with an appliance product from our store. There should be exactly 3 rows, one for each recipe you selected. The Recipe_Name should be the Chinese name exactly as it appears in the recipe database. Recipe_Category should be the category name in Chinese. Difficulty should be the difficulty level from the recipe. Paired_Product should be the name of the store product you are pairing with that recipe. Bundle_Price should be the price of the paired product multiplied by 0.85, rounded to 2 decimal places, representing a 15 percent bundle discount.

To solve this task efficiently, complete it directly in the main agent; do not dispatch sub-agents. Freeze the task-visible agent-workspace absolute path from the Visible Boundary/current working directory. In the first native-parallel source response, call `woo_products_categories_list` and query exact recipe categories `饮品`, `主食`, and `荤菜` while retrieving daily history for `GC=F`. Consume a recipe category's runtime `output_path` when its response is overlong; do not repeat that category call. From the Woo category response require exactly one complete-name match for `Home Appliances` and freeze its literal category ID; from each recipe response select one distinct suitable exact recipe ID.

In the next native-parallel response, call `woo_products_list(category=<literal Home Appliances category ID as a string>, page=1, perPage=100)` and retrieve only the three selected exact recipe IDs. If the first product page is full, continue with increasing page numbers through a short or empty terminal page. Never pass the category name where the tool requires an ID, and never substitute sale/current price for `regular_price`.

Require complete unique appliance rows, three distinct source-backed recipe rows, and the five newest distinct gold dates. Freeze three explicit recipe-to-product pairings whose product names exactly match the catalog and compute `Bundle_Price = regular_price * 0.85` once from raw values. As sole writer create `<agent_workspace>/Bundle_Pricing.xlsx` with exactly `Store Products(Product_Name,Price,Stock_Quantity)`, `Gold Trend(Date,Close_Price)`, and three-row `Recipe Bundles(Recipe_Name,Recipe_Category,Difficulty,Paired_Product,Bundle_Price)` using the requested sorts, literal numeric cells, Chinese recipe names/categories, and rounding. Read each populated range back once and finish without a verification wave.
