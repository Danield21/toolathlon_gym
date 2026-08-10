I am a research librarian doing a cross-reference analysis between two academic databases. I need to compare the papers available in the Scholarly database (specifically the arxiv_papers collection) with those in the arxiv database to find overlapping papers and understand coverage differences.

Please create an Excel file called "Citation_Cross_Reference.xlsx" in my workspace with three sheets.

The first sheet, "Scholarly_Papers", should list all papers from the Scholarly arxiv_papers collection with columns: Paper_ID, Title, Authors, Published_Year, Journal_Ref. Published_Year is the four-digit year of the paper's published date. Authors should be listed as a single comma-separated string.

The second sheet, "Arxiv_Papers", should list all papers from the arxiv database with columns: Paper_ID, Title, Authors, Published_Year, Categories. Published_Year is the four-digit year of the paper's published date. Categories should be a comma-separated string.

The third sheet, "Overlap_Analysis", should contain every paper that appears in either database (the union), matched by paper ID, with columns: Paper_ID, Title, In_Scholarly (Yes/No), In_Arxiv (Yes/No). A paper present in both is marked Yes/Yes; a paper only in Scholarly is Yes/No; a paper only in arxiv is No/Yes.

After creating the spreadsheet, please send an email summarizing the findings to research-lead@university.edu with the subject "Cross-Reference Analysis: Scholarly vs Arxiv Paper Databases". The email body should mention the total number of papers in each database, the number of overlapping papers, and highlight any papers that are unique to one database.
