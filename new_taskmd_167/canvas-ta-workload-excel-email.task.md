The department chair wants to understand the workload distribution among Teaching Assistants across all courses to make informed staffing decisions for next semester. Pull TA enrollment data across all courses along with assignment and submission counts to assess the workload each TA is responsible for.

Create an Excel file called TA_Workload_Report.xlsx with two sheets. The first sheet should be called "Course Workload" and contain the following columns: Course_Name, Course_Code, TA_Count, Assignment_Count, Submission_Count, and Submissions_Per_TA. Submissions_Per_TA should be the Submission_Count divided by TA_Count, rounded to 1 decimal place. If a course has no TAs, set Submissions_Per_TA to 0. Sort the rows by Course_Name alphabetically. The second sheet should be called "Summary" with rows for Total_Courses, Total_TAs (sum of all TA_Count across courses), Avg_TAs_Per_Course (rounded to 1 decimal), Max_Assignment_Count (the highest assignment count among all courses), and Most_Loaded_Course (the course name with the highest Submissions_Per_TA).

Then create a Notion page called "TA Staffing Overview" that summarizes the key findings from the workload analysis including total courses, total TAs, and which course has the heaviest TA workload.

Finally, email the report to dept_chair@university.edu with subject "TA Workload Report" and a brief summary in the body.

To solve this task efficiently, use one compact Explore agent per concrete course. Each course has a true local dependency from its assignment list to its bulk submission-count call, but courses are independent.

The main agent calls canvas_list_courses(include_ended=true), validates unique returned IDs/codes/full names, and instantiates one literal prompt per course.

1. Wave 1 — dispatch exactly 22 explore sub-agents, one for each literal course in the validated inventory, all in parallel:

   Each Explore agent receives one exact course_id, course_code and course_name. First call canvas_list_assignments(course_id=<ID>, include_submissions=false), validate unique assignment IDs, and freeze that literal assignment_ids list. In the next response call natively in parallel:
   - canvas_get_course_grades(course_id=<ID>, page=1, per_page=1, type=["TaEnrollment"]); take TA_Count from pagination.total_count and do not fetch TA rows.
   - canvas_list_student_submissions(course_id=<ID>, assignment_ids=<literal returned assignment IDs>, student_ids=["all"], grouped=false, page=1, per_page=1); take Submission_Count from pagination.total_count and do not fetch later submission pages.
   Return only {course_id,course_code,course_name,assignment_ids,assignment_count,ta_count,submission_count,assignment_response_count,ta_pagination,submission_pagination}. Do not modify Canvas.

After Wave 1, the main agent requires exactly one result per course, reconciles all IDs, computes Submissions_Per_TA with zero for courses without TAs, sorts course names, and freezes the complete Course Workload table, Summary rows, deterministic Most_Loaded_Course, Notion paragraph, and email facts.

2. Wave 2 — dispatch exactly 2 coder sub-agents in parallel:

   1. One Coder is the sole writer of TA_Workload_Report.xlsx in the dynamically resolved workspace. Create exactly Course Workload and Summary with exact headers, all course rows, literal values and rounding, then read both populated ranges back.
   2. One Coder solely owns the Notion page TA Staffing Overview. Search and retrieve all exact-title candidates; stop without mutation on duplicates, create exactly once only if absent, and append the complete canonical paragraph exactly once only if absent. The paragraph must contain the literal total course count, total TA count, full most-loaded course name, and Submissions_Per_TA. Retrieve the page and child blocks after writing and leave unrelated content untouched.

In the same response that launches Wave 2, the main agent searches Sent mail for exact To dept_chair@university.edu and subject TA Workload Report, reading candidates to verify recipients. Send once only if absent, using the same canonical total courses, total TAs, average TAs and full heaviest-workload course, then read it back. After Wave 2, reconcile workbook, Notion and email against the canonical payload; do not add a verification wave.
