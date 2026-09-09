I need to create a comprehensive grade report for our department heads.

Data source: pull course, enrollment, and grade data from the learning management system (Canvas LMS) for all courses. For each course, use the enrollment records and each student's overall final score — the `final_score` value inside the enrollment's `grades` data. Include every student who has a recorded final score; students without a recorded final score are excluded from that course's totals.

In the workspace, use the terminal to create a Python script called grade_reporter.py that reads grade_data.json (create it first — it should hold the course name and the list of final scores for every student in every course), calculates the grade distributions, pass rates, and course averages, and writes grade_report.json. Use the standard grade scale: A: 90+, B: 80-89, C: 70-79, D: 60-69, F: below 60. Pass means grade C or above. Round Pass_Rate_Pct and Course_Avg to 1 decimal place. The datasets are large (thousands of enrollments per course); it is fine to build grade_data.json incrementally, one course at a time, as you pull each course's enrollment data from the LMS rather than relying on a single large dump.

Create a Google Sheet titled "Department Grade Dashboard" with two sheets. The first sheet Grade_Distribution should have columns Course_Name, A_Count, B_Count, C_Count, D_Count, F_Count, Total_Students, Pass_Rate_Pct (round to 1 decimal, pass means C or above), and Course_Avg (round to 1 decimal), sorted by Course_Name. Use the full course name as shown in the LMS for the Course_Name column. The second sheet Department_Summary should have Metric and Value columns with Total_Courses, Total_Students, Overall_Pass_Rate (round to 1 decimal), Overall_Avg_Grade (round to 1 decimal), Highest_Avg_Course, and Lowest_Avg_Course.

Also read the Course_Policies.pdf in the workspace which has the grade scale and academic policies, and make sure the grade thresholds you use match those defined in the policy document.

Send an email to dept-heads@university.edu with subject "Q1 2026 Grade Distribution Report" including a summary of the overall pass rate, the highest and lowest performing courses, and a note that courses with pass rates below 70% require review.

To solve this task efficiently, use one course-owner source wave and let the main agent own the canonical script plus the small Google Sheet and email writes. Canvas is read-only. Resolve the runtime workspace dynamically and read Course_Policies.pdf and reporting_guide.md completely, freezing grade bands, C-or-above passing threshold, below-70 review rule and report requirements.

Call canvas_list_courses(include_ended=true) once and freeze all 22 exact course IDs/codes/full names.

Wave 1 — start one Coder owner per verified course in one parallel response. Each receives one literal course identity and unique JSON path, calls canvas_get_course_grades(course_id=<id>,page=1,per_page=100,type=["StudentEnrollment"]), then issues all remaining literal pages from total_pages in native-parallel batches and verifies every page identity. It retains only numeric non-null grades.final_score with literal enrollment row.id, rejects duplicates, writes one compact JSON with exact full course name, all final scores and page/count controls, reads it back, and returns only path/hash and controls. It must not use current_score or a reconstructed course label.

After the wave, require all 22 receipts, every advertised page and disjoint enrollment IDs. The main agent is sole owner of exact files grade_data.json, grade_reporter.py and grade_report.json. Write all exact full course names/final scores to grade_data.json; write and run the script once to compute A/B/C/D/F, Total_Students, C-or-above pass rate, one-decimal course averages, weighted overall pass rate and overall average; then read all three files back and reconcile 22 courses, totals, buckets, extrema and rounding.

The main agent starts the two independent external branches in one native-parallel response:

- Search/reuse-or-create the exact Google spreadsheet Department Grade Dashboard at most once, block duplicate exact titles/tabs, own exactly Grade_Distribution and Department_Summary, write 22 alphabetical rows and six exact summary metrics, then read both ranges back.
- Paginate Sent mail and read candidates for exact recipient dept-heads@university.edu and subject Q1 2026 Grade Distribution Report. Send at most once with canonical overall pass rate, full highest/lowest course names and below-70 review statement, then read the retained message back.

Do not dispatch page, aggregation, cloud-writer, email or audit agents. Final acceptance compares the exact three files, external sheet ranges and retained email to the course receipts.
