The Office of Academic Affairs is conducting its annual review of course offerings and wants to compare our university's course metrics against national benchmarks. A benchmarking service provides national averages through a REST API endpoint at http://localhost:30203/api/benchmarks.json, which returns data including national average enrollment, assignment counts, and quiz counts for seven course discipline areas.

Our university offers 22 courses through our learning management system across seven distinct course types. We need to understand how our courses compare to national standards in terms of enrollment numbers, number of assignments, and number of quizzes.

Start by fetching the national benchmark data from the API endpoint. Then retrieve the full list of courses from the university learning management system. For each course, note the course name, total student enrollment, and count the number of assignments and quizzes associated with it. Use each course's reported total student count (`total_students`, i.e. student enrollments only — do not count teacher or TA enrollments) as the enrollment figure. Group the courses by their course type (the part of the course name before the semester and year in parentheses, for example "Foundations of Finance" from "Foundations of Finance (Fall 2013)") and compute the average enrollment, average assignment count, and average quiz count per course type.

Create an Excel file called "Course_Benchmark_Analysis.xlsx" in the workspace with three sheets.

The first sheet called "National Benchmarks" should have columns Course_Type, Discipline, National_Avg_Enrollment, National_Avg_Assignments, and National_Avg_Quizzes, populated with the benchmark API data for all seven course types.

The second sheet called "Our Courses" should list all 22 individual courses with columns Course_Name, Course_Type, Enrollment, Assignment_Count, and Quiz_Count.

The third sheet called "Comparison" should have columns Course_Type, Our_Avg_Enrollment, National_Avg_Enrollment, Enrollment_Diff, Our_Avg_Assignments, National_Avg_Assignments, Assignment_Diff, Our_Avg_Quizzes, National_Avg_Quizzes, and Quiz_Diff. The difference columns should be calculated as our average minus the national average. There should be 7 rows, one for each course type.

Next, create a PowerPoint presentation called "Academic_Benchmark_Presentation.pptx" in the workspace. The presentation should have a title slide with the title "Course Benchmark Analysis" and a subtitle referencing the academic year. Include one slide per course type (7 slides) showing the course type name as the title and a brief comparison of enrollment and assignment metrics versus national benchmarks. Add a final summary slide identifying which course types exceed national benchmarks in enrollment and which fall below.

Finally, send an email to dean@university.edu with the subject "Annual Course Benchmark Report". The sender address is fixed by the email environment; only the recipient and subject matter. The body should summarize which course types are above and below national enrollment benchmarks.

Save all output files to the workspace directory.

To solve this task efficiently, combine each course's assignment and quiz reads in one Explore agent and keep the two independent artifact writers parallel.

Before Wave 1, the main agent natively in parallel fetches http://localhost:30203/api/benchmarks.json and calls canvas_list_courses(include_ended=true). Validate exactly seven unique benchmark course types and unique Canvas IDs/codes/full names, retaining each course's reported total_students. Instantiate one Explore prompt per literal returned course.

1. Wave 1 — dispatch exactly 22 explore sub-agents, one for each literal course in the validated inventory, all in parallel:

   Each Explore agent receives one exact course_id, course_code and full course_name and calls canvas_list_assignments(course_id=<ID>, include_submissions=false) and canvas_list_quizzes(course_id=<ID>) natively in parallel. Return only {course_id,course_code,course_name,assignment_ids,assignment_count,quiz_ids,quiz_count,returned_counts,duplicate_ids}, using empty lists and zero counts when appropriate. Do not fetch submissions or modify Canvas.

After Wave 1, the main agent requires exactly one result per frozen course, derives Course_Type by removing only the trailing semester/year parenthetical, joins to the seven benchmark rows, and computes per-type averages and university-minus-national differences from unrounded inputs. Freeze canonical National Benchmarks, 22-row Our Courses, seven-row Comparison, presentation payload, above/below enrollment classifications and email facts. Writers may not recompute groupings or classifications.

2. Wave 2 — dispatch exactly 2 coder sub-agents in parallel:

   1. One Coder is the sole writer of Course_Benchmark_Analysis.xlsx in the dynamically resolved workspace. Create exactly National Benchmarks, Our Courses, and Comparison with exact headers, 7/22/7 row coverage, literal values and task ordering, then read every populated range back.
   2. One Coder is the sole writer of Academic_Benchmark_Presentation.pptx in the same workspace. Consume only the frozen payload and create exactly nine slides: title with source-supported academic-year subtitle, seven literal course-type slides, and one final above/below summary; read all slide titles/text back.

In the same response that launches Wave 2, the main agent searches Sent mail for complete To dean@university.edu and subject Annual Course Benchmark Report. Read candidates to verify recipients, send exactly once only if absent using the frozen above/below classifications, and read the retained message back. After Wave 2, reconcile workbook, presentation and email against the same canonical payload; do not dispatch a verification wave.
