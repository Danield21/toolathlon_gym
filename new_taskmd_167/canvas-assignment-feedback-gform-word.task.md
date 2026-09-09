You are a course analytics assistant for the Foundations of Finance course offered in Fall 2013. This course has a course ID of 16 in the learning management system.

Your task has three parts.

Part 1: Retrieve all assignments for course 16 and gather submission statistics for each one. For each assignment, collect the assignment name, total number of submissions, average score (rounded to 2 decimal places), number of late submissions, and late submission rate (late submissions divided by total submissions, multiplied by 100, rounded to 2 decimal places).

Part 2: Create a Google Forms survey titled "Foundations of Finance - Assignment Feedback Survey" to collect student feedback on assignments. The survey must contain exactly 5 questions covering the following topics: overall assignment workload, clarity of assignment instructions, time required to complete assignments, availability of instructor support, and overall satisfaction with the assignments. Use appropriate question types for each.

Part 3: Create a Word document called Assignment_Analysis.docx in the workspace. The document must have the title "Foundations of Finance (Fall 2013) - Assignment Analysis" at the top. It must contain a table with the following columns in this exact order: Assignment_Name, Total_Submissions, Avg_Score, Late_Submissions, Late_Rate(%). Sort the rows alphabetically by Assignment_Name. Include all assignments found in the course that have at least one submission (assignments with no submissions may be excluded).

After creating the Word document, send a summary email to instructor@financeou.example.com from analytics@university.example.com with the subject "Assignment Analysis Report - Foundations of Finance (Fall 2013)". The email body must mention the total number of assignments analyzed, the assignment with the highest average score, and the assignment with the most late submissions. Also include the Google Forms survey link in the email body.

The initial workspace contains Assignment_Guidelines.pdf for reference on the grading criteria used in this course.

To solve this task efficiently,

This is a compact one-course task. The main agent completes it directly; do not dispatch any sub-agent. Canvas is read-only.

First resolve the task-visible workspace directory from the runtime allowed-directory listing (or the literal accessible-workspace directory already supplied by the harness), freeze it as `<workspace_root>`, and reject bare `/workspace` and dumps/run-specific paths. In one native-parallel response, call `canvas_list_courses(include_ended=true)`, extract all pages of `<workspace_root>/Assignment_Guidelines.pdf` read-only with `python_execute`, call `create_form(title="Foundations of Finance - Assignment Feedback Survey")` exactly once, and search Sent mail for the exact recipient/subject. Select exactly one returned course whose code is `FFF-2013J` and whose full name is `Foundations of Finance (Fall 2013)`; stop on zero or multiple matches, and never retry Form creation after an ambiguous result. Then call `canvas_list_assignments(course_id=<literal returned course ID>)` once.

From the returned assignment list, issue one independent `canvas_list_assignment_submissions(course_id=<literal returned course ID>, assignment_id=<returned assignment ID>, page=1, per_page=1)` call per assignment in one native-parallel response. Use the whole-assignment `summary` returned on each call; do not page raw submissions. For every assignment with `summary.total_count>0`, compute:
`Total_Submissions=total_count`, `Avg_Score=round(avg_score,2)`, `Late_Submissions=late_count`, and `Late_Rate(%)=round(100*late_count/total_count,2)`.
Sort rows by exact assignment name. Select highest average score and most late submissions with assignment-name ascending tie-breaks.

Freeze one literal canonical payload containing all sorted rows, analyzed count, the two named extrema, and the verified Form ID/responder link. The Word document and email must consume exactly this payload without recomputing statistics.

Complete the Form chain on the retained ID, adding exactly these five multiple-choice questions in order and then reading the Form back:
1. `How would you rate the overall assignment workload?` — `Very light, Light, About right, Heavy, Very heavy`
2. `How clear were the assignment instructions?` — `Very unclear, Unclear, Neutral, Clear, Very clear`
3. `How much time did you typically spend on each assignment?` — `Less than 1 hour, 1-3 hours, 4-6 hours, 7-10 hours, More than 10 hours`
4. `How available was instructor support for assignments?` — `Never available, Rarely available, Sometimes available, Usually available, Always available`
5. `What is your overall satisfaction with the assignments?` — `1, 2, 3, 4, 5`
Require exactly this five-question sequence on the newly retained Form ID; do not append duplicates.

Create `<workspace_root>/Assignment_Analysis.docx` exactly once, preferably atomically with `python-docx`, using title `Foundations of Finance (Fall 2013) - Assignment Analysis` and one table with exact columns `Assignment_Name, Total_Submissions, Avg_Score, Late_Submissions, Late_Rate(%)`. Read the document back and require every canonical row in alphabetical order with literal numeric values.

For email idempotency, paginate `search_emails(query="Assignment Analysis Report - Foundations of Finance (Fall 2013)", folder="Sent", page=<n>, page_size=50)`, read every candidate, and match complete To, From, and Subject. Send exactly once only when no exact match exists, from `analytics@university.example.com` to `instructor@financeou.example.com`. The body must include canonical analyzed count, named highest-average assignment, named most-late assignment, and verified Form responder link. Repeat search/read-back and require exactly one exact message.

Accept only after the Form, Word file, and email read back against the same canonical payload.
