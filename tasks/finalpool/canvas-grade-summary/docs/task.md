I need a summary of overall student grades across all courses. For each course, query the learning management system for that course's enrollments with grades, and use each enrolled student's **current overall grade** — the `current_score` field inside each enrollment's grade record. Ignore enrollments that have no grade or a null `current_score`.

Create an Excel file called Canvas_Grade_Summary.xlsx with two sheets.

The first sheet, named "Grade Summary", should have exactly one row per course, with these columns in this order: **Course** (the course name), **Students_Graded** (the number of students in the course whose `current_score` is present/non-null), **Avg_Score**, **Max_Score**, and **Min_Score**. Avg_Score, Max_Score, and Min_Score are the average, maximum, and minimum of those students' `current_score` values, each rounded to 2 decimals. Sort the rows by Avg_Score descending.

The second sheet, named "Summary", should contain three rows: **Total_Courses** (the total number of courses), **Highest_Avg_Course** (the name of the course with the highest Avg_Score), and **Overall_Avg_Score** (the unweighted average of the per-course Avg_Score values across all courses, rounded to 2 decimals).

Also record these three summary metrics (Total_Courses, Highest_Avg_Course, Overall_Avg_Score) in a Google Sheet titled "Grade Summary Report" for the academic team to review.
