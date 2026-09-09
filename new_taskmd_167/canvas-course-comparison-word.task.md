The provost has asked for a year-over-year academic performance comparison between Fall 2013 and Fall 2014. We need to identify courses that were offered in both semesters and compare their enrollment numbers, assignment counts, and average submission scores.

Please query the learning management system data to find courses with course codes ending in "2013J" and "2014J". Match courses that share the same prefix (for example AAA-2013J and AAA-2014J are the same course offered in different years) — only include a course prefix when it exists in both semesters. For each matched pair, compare the enrollment (total_students), the number of assignments, and the average submission score. The average submission score is the mean of all non-null submission scores in the course (i.e. a simple average over every graded submission, not a per-student average).

Create an Excel file called Year_Over_Year_Comparison.xlsx in the workspace with two sheets. The first sheet should be named "Course Comparison" with columns Course_Name (the common course name without the year suffix), Fall_2013_Enrollment, Fall_2014_Enrollment, Enrollment_Change (2014 minus 2013), Fall_2013_Assignments, Fall_2014_Assignments, Fall_2013_Avg_Grade (rounded to 2 decimal places), Fall_2014_Avg_Grade (rounded to 2 decimal places), and Grade_Change (2014 avg minus 2013 avg, rounded to 2 decimal places). Sort rows alphabetically by Course_Name. The second sheet should be named "Summary" with two columns, Metric and Value, containing these rows: Courses_Compared (number of course pairs), Avg_Enrollment_Change (average of all enrollment changes, rounded to 1 decimal), Avg_Grade_Change (average of all grade changes, rounded to 2 decimal places).

Create a Word document called Academic_Year_Comparison.docx in the workspace. The document should have the title "Fall 2013 vs Fall 2014 Academic Performance Review" followed by a narrative paragraph for each course pair describing the enrollment and grade changes. Conclude with a summary paragraph.

Create a Google Sheet titled "Academic Year Comparison" with a sheet called "Comparison Data" containing the same data as the Excel Course Comparison sheet.

To solve this task efficiently,

Follow this exact orchestration plan. Canvas is read-only. The valid paired prefixes are AAA, BBB, DDD, EEE, FFF, and GGG; CCC has no Fall 2013 counterpart and is excluded. Before Wave 1, resolve the task-visible workspace directory from the runtime allowed-directory listing (or the literal accessible-workspace directory already supplied by the harness), freeze it as `<workspace_root>`, and reject bare `/workspace` and dumps/run-specific paths. After Wave 1, freeze one literal canonical six-pair payload; both writers consume exactly this payload without recomputing comparisons.

1. **Wave 1 — dispatch exactly 113 explore sub-agents concurrently:** render all 113 subtasks as parallel individual Agent calls in one orchestration response; do not use AgentSwarm because the wave mixes three instruction templates.

- **Explore assignment-summary template:** substitute `{course_id,course_code,assignment_id}`; call `canvas_list_assignment_submissions(course_id={course_id},assignment_id={assignment_id},page=1,per_page=1)` once. The returned `summary` covers the whole assignment; return `{course_id,course_code,assignment_id,total_count:summary.total_count,graded_count:summary.graded_count,avg_score:summary.avg_score}`; do not page raw rows.
- **Explore metadata template:** call `canvas_get_course(course_id={course_id})` once; verify code and return `{course_id,course_code,course_name,total_students}`.
- **Agent 113 — Explore identity verifier:** call `canvas_list_courses(include_ended=true)` once; verify exactly the 12 listed ID/code identities form the six pairs and no CCC pair exists; return `{paired_prefixes,missing,unexpected,identity_mismatches}`.

Exact assignment rows:

|agent_no|course_id|course_code|assignment_id|
|---|---|---|---|
|1|1|AAA-2013J|1|
|2|1|AAA-2013J|2|
|3|1|AAA-2013J|3|
|4|1|AAA-2013J|4|
|5|1|AAA-2013J|5|
|6|1|AAA-2013J|6|
|7|2|AAA-2014J|7|
|8|2|AAA-2014J|8|
|9|2|AAA-2014J|9|
|10|2|AAA-2014J|10|
|11|2|AAA-2014J|11|
|12|2|AAA-2014J|12|
|13|3|BBB-2013J|25|
|14|3|BBB-2013J|26|
|15|3|BBB-2013J|27|
|16|3|BBB-2013J|28|
|17|3|BBB-2013J|29|
|18|3|BBB-2013J|30|
|19|3|BBB-2013J|31|
|20|3|BBB-2013J|32|
|21|3|BBB-2013J|33|
|22|3|BBB-2013J|34|
|23|3|BBB-2013J|35|
|24|3|BBB-2013J|36|
|25|4|BBB-2014J|49|
|26|4|BBB-2014J|50|
|27|4|BBB-2014J|51|
|28|4|BBB-2014J|52|
|29|4|BBB-2014J|53|
|30|4|BBB-2014J|54|
|31|9|DDD-2013J|89|
|32|9|DDD-2013J|90|
|33|9|DDD-2013J|91|
|34|9|DDD-2013J|92|
|35|9|DDD-2013J|93|
|36|9|DDD-2013J|94|
|37|9|DDD-2013J|95|
|38|10|DDD-2014J|103|
|39|10|DDD-2014J|104|
|40|10|DDD-2014J|105|
|41|10|DDD-2014J|106|
|42|10|DDD-2014J|107|
|43|10|DDD-2014J|108|
|44|10|DDD-2014J|109|
|45|13|EEE-2013J|110|
|46|13|EEE-2013J|111|
|47|13|EEE-2013J|112|
|48|13|EEE-2013J|113|
|49|13|EEE-2013J|114|
|50|14|EEE-2014J|120|
|51|14|EEE-2014J|121|
|52|14|EEE-2014J|122|
|53|14|EEE-2014J|123|
|54|14|EEE-2014J|124|
|55|16|FFF-2013J|138|
|56|16|FFF-2013J|139|
|57|16|FFF-2013J|140|
|58|16|FFF-2013J|141|
|59|16|FFF-2013J|142|
|60|16|FFF-2013J|143|
|61|16|FFF-2013J|144|
|62|16|FFF-2013J|145|
|63|16|FFF-2013J|146|
|64|16|FFF-2013J|147|
|65|16|FFF-2013J|148|
|66|16|FFF-2013J|149|
|67|16|FFF-2013J|150|
|68|17|FFF-2014J|164|
|69|17|FFF-2014J|165|
|70|17|FFF-2014J|166|
|71|17|FFF-2014J|167|
|72|17|FFF-2014J|168|
|73|17|FFF-2014J|169|
|74|17|FFF-2014J|170|
|75|17|FFF-2014J|171|
|76|17|FFF-2014J|172|
|77|17|FFF-2014J|173|
|78|17|FFF-2014J|174|
|79|17|FFF-2014J|175|
|80|17|FFF-2014J|176|
|81|20|GGG-2013J|177|
|82|20|GGG-2013J|178|
|83|20|GGG-2013J|179|
|84|20|GGG-2013J|180|
|85|20|GGG-2013J|181|
|86|20|GGG-2013J|182|
|87|20|GGG-2013J|183|
|88|20|GGG-2013J|184|
|89|20|GGG-2013J|185|
|90|20|GGG-2013J|186|
|91|21|GGG-2014J|197|
|92|21|GGG-2014J|198|
|93|21|GGG-2014J|199|
|94|21|GGG-2014J|200|
|95|21|GGG-2014J|201|
|96|21|GGG-2014J|202|
|97|21|GGG-2014J|203|
|98|21|GGG-2014J|204|
|99|21|GGG-2014J|205|
|100|21|GGG-2014J|206|

Exact metadata rows:

|agent_no|course_id|course_code|
|---|---|---|
|101|1|AAA-2013J|
|102|2|AAA-2014J|
|103|3|BBB-2013J|
|104|4|BBB-2014J|
|105|9|DDD-2013J|
|106|10|DDD-2014J|
|107|13|EEE-2013J|
|108|14|EEE-2014J|
|109|16|FFF-2013J|
|110|17|FFF-2014J|
|111|20|GGG-2013J|
|112|21|GGG-2014J|

After Wave 1, per course compute the simple average over all graded submissions as `sum(assignment avg_score * graded_count)/sum(graded_count)`, ignoring zero-graded assignments; round course averages two decimals. Match six prefixes, calculate enrollment and grade changes and summary averages with specified rounding; sort common course name alphabetically.

2. **Wave 2 — dispatch exactly 3 coder sub-agents in parallel:**

   1. **Coder Excel writer:** create `<workspace_root>/Year_Over_Year_Comparison.xlsx` with only `Course Comparison` and vertical `Summary`; exact nine headers, six sorted rows, three metrics from the canonical payload; read back and return `{path,sheet_names,row_counts,readback_ok}`.
   2. **Coder Word writer:** create `<workspace_root>/Academic_Year_Comparison.docx`, exact title, one evidence-based paragraph for each of six canonical pairs, conclusion; read text/info back and return `{path,course_names,readback_ok}`.
   3. **Coder Google Sheets writer:** act as the only writer for exact spreadsheet title `Academic Year Comparison` and exact tab `Comparison Data`. Call `search_spreadsheets(query="Academic Year Comparison", max_results=100)` and compare the complete returned `name` field: if no exact-title spreadsheet exists, call `create_spreadsheet(title="Academic Year Comparison")` exactly once and retain its `spreadsheetId`; if exactly one exists, reuse its `id`; if more than one exists, return a duplicate blocker without writing. Call `list_sheets(spreadsheet_id=<retained id>)`; if `Comparison Data` occurs more than once, return a blocker. If it is absent, rename the sole empty default tab with `rename_sheet(spreadsheet=<retained id>, sheet=<literal default-tab name>, new_name="Comparison Data")`, or call `create_sheet(spreadsheet_id=<retained id>, title="Comparison Data")` exactly once when a non-default tab must be preserved. Write the exact nine headers and six rows with `update_cells(spreadsheet_id=<retained id>, sheet="Comparison Data", range="A1:I7", data=<literal 7-by-9 matrix>)`. Then call `list_sheets` again and `get_sheet_data(spreadsheet_id=<retained id>, sheet="Comparison Data", range="A1:I7")`; require one exact tab and exact cell equality, and return `{spreadsheet_id,sheet_name:"Comparison Data",row_count:6,duplicate_check,readback_ok}`. Never create a second spreadsheet or retry a create/write after an ambiguous failure.

Accept only after all three read-backs.
