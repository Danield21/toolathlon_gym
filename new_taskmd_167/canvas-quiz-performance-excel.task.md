I need an analysis of quiz performance for the course "Biochemistry & Bioinformatics" which is offered in Fall 2013 and has a course ID of 3 in the learning management system. (The course's exact name in the LMS is "Biochemistry & Bioinformatics (Fall 2013)"; use the course ID 3 to look it up.) Please look up all quizzes for this course and gather the submission data for each quiz.

Create an Excel file called Quiz_Performance.xlsx in the workspace with two sheets. The first sheet should be called "Quiz Stats" with columns Quiz_Title, Total_Submissions, Avg_Score (rounded to 2 decimal places), Min_Score, Max_Score, and Pass_Rate (the percentage of submissions scoring 60 or above, rounded to 2 decimal places). Sort the rows by quiz title.

The second sheet should be called "Summary" with two columns Metric and Value. Include the following rows: Total_Quizzes (the number of distinct quizzes), Overall_Avg_Score (the average score across all submissions in all quizzes for this course, rounded to 2 decimal places), Highest_Avg_Quiz (the title of the quiz with the highest average score), and Lowest_Avg_Quiz (the title of the quiz with the lowest average score).

After creating the Excel file, send a summary email to instructor@university.example.com from quiz-analytics@university.example.com with the subject "Quiz Performance Summary - Biochemistry & Bioinformatics (Fall 2013)". The email body should include the total number of quizzes analyzed, the overall average score, and the names of the highest and lowest performing quizzes.

To solve this task efficiently, the main agent must complete this small five-quiz task directly; do not dispatch any sub-agent. The five-row table supplies exact fields to the quiz procedure below.

Email-owner contract for every email sub-task below: call `search_emails(query=<literal subject>, folder="Sent", page=<n>, page_size=50)` from page 1 through the returned final page. Because search results do not expose recipients, call `read_email(email_id=<candidate id>)` for every candidate and match the complete literal `To`, `Subject`, and, whenever the Task specifies it, `From`; a subject-only hit is not a match. Call `send_email` exactly once only when no exact match exists. After sending, repeat the paginated search, call `read_email` on every candidate, require exactly one exact match, and return its ID plus exact recipient/sender/subject/body read-back. If search, send, or read-back fails, return a blocker and do not retry `send_email`.

**Quiz procedure.** Call `canvas_list_quiz_submissions(course_id=3, quiz_id=<quiz_id>)` exactly once; the response is complete and unpaginated. Treat this endpoint only as score evidence, and obtain `quiz_title` and `points_possible` only from the `canvas_list_quizzes` inventory. For numeric scores compute `score_sum`, `score_min`, `score_max`, and `pass_count_ge_60`; define every submission identity and duplicate check from literal `row.id`. Do not modify Canvas.

The main agent performs these seven independent read-only operations directly in one native-parallel tool-call response:

   1. Call `canvas_get_course(course_id=3)` exactly once. Return `{course_id,course_code,course_name,total_students,identity_ok,call_count,warnings}` and require exact name `Biochemistry & Bioinformatics (Fall 2013)`. Do not call quizzes or submissions.

   2. Call `canvas_list_quizzes(course_id=3)` exactly once. Return every quiz as `{quiz_id,title,points_possible}`, sorted by quiz ID, plus `returned_count`, unique IDs, and warnings. Require the exact manifest `6:CMA 15003, 7:CMA 15004, 8:CMA 15005, 9:CMA 15006, 10:CMA 15007`; do not call submissions.

   Apply the quiz procedure to these exact rows:

| row_no | course_id | quiz_id | quiz_title |
|---:|---:|---:|---|
| 1 | 3 | 6 | CMA 15003 |
| 2 | 3 | 7 | CMA 15004 |
| 3 | 3 | 8 | CMA 15005 |
| 4 | 3 | 9 | CMA 15006 |
| 5 | 3 | 10 | CMA 15007 |

The main agent verifies course identity, requires the five-row inventory exactly once, joins title and points exclusively from that inventory, and derives the task metrics with the stated denominators and deterministic tie-break.

It then starts the workbook-create/write branch and the exact sent-mail search in the same native-parallel response; neither branch waits for the other. Workbook read-back and the email's conditional send/read-back continue independently. It performs both outputs directly:

   1. Resolve the runtime workspace root and own `Quiz_Performance.xlsx` there exactly once. Create `Quiz Stats(Quiz_Title,Total_Submissions,Avg_Score,Min_Score,Max_Score,Pass_Rate)` with exactly five title-sorted rows and two-decimal averages/rates, and `Summary(Metric,Value)` with exactly `Total_Quizzes`, `Overall_Avg_Score`, `Highest_Avg_Quiz`, and `Lowest_Avg_Quiz`. Write literal values, not formulas. Read both sheets back and return `{path,hash,sheets:[{name,headers,row_count}],quiz_titles,stored_summary,arithmetic_checks,verification}`. Do not send email.

   2. Own only one email from `quiz-analytics@university.example.com` to `instructor@university.example.com` with subject `Quiz Performance Summary - Biochemistry & Bioinformatics (Fall 2013)`. Follow the Email-owner contract above. Search the exact sender/recipient/subject first; if absent, call `send_email(to="instructor@university.example.com", from_addr="quiz-analytics@university.example.com", subject="Quiz Performance Summary - Biochemistry & Bioinformatics (Fall 2013)", body=<verified total quiz count, overall average, complete highest title, complete lowest title>)` once; if present, do not duplicate it. Read back or search again and return `{message_id,from,to,subject,total_quizzes,overall_average,highest_title,lowest_title,action,duplicate_count,verification}`. Do not modify files.

Finally, reconcile the five source summaries, two workbook sheets, and one email. Verify quiz identity/order, denominators, rounding, literal workbook values, full highest/lowest titles, sender/recipient/subject/body values, read-back evidence, and absence of a duplicate message before claiming completion.
