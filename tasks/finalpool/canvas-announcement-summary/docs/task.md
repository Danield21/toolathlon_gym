I need a summary of course announcements from our learning management system. Pull all announcements and aggregate by course, showing count and date range.

Create an Excel file called Course_Announcements.xlsx with two sheets. The first sheet, named "Announcement Stats", should have columns Course_Code, Announcements, Earliest_Date, and Latest_Date with one row per course, sorted by Announcements descending. Use UTC dates (YYYY-MM-DD) for Earliest_Date and Latest_Date — do not apply any timezone conversion to the timestamps returned by the announcements API.

The second sheet, named "Summary", should be a two-column table with the header "Metric, Value" and exactly three rows: Total_Announcements (total announcements across all courses), Courses_With_Announcements (number of distinct courses that have at least one announcement), and Most_Active_Course (the course code with the highest announcement count).

Send an email to academic-affairs@openuniversity.ac.uk with subject "Course Announcement Activity Report" briefly summarizing the activity levels.
