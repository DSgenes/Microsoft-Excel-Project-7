# Microsoft-Excel-Project-7

# Case study

Adventure Works is preparing a series of advertising campaigns to be rolled out in several different regions. A colleague, Lucas, has asked you to update a spreadsheet that focuses on the launch dates for the USA campaign. The spreadsheet is called Advertising Campaign USA Dates.xlsx.

For each project, Lucas needs to know the following information:

⦁ The number of working days available between the start date and the deadline date.

⦁ The month and year when each campaign will launch.

⦁ The number of calendar days to the deadline date for each campaign.
______________________________________________________________________________________________________________________________________________________________________________________________________________________________________________

# Calculating the Number of Working Days Remaining in a Year in Excel

# Overview: 

In this exercise, I practiced using date functions in Microsoft Excel to calculate key timeline information for projects, such as the total number of calendar days, working days (excluding weekends and holidays), and extracting month and year components from a deadline date.

# Key Tasks Completed:

# 1. Calculate Total Calendar Days Between Two Dates:

   ⦁ Used subtraction to calculate the number of days between the start date and the deadline date.
   ⦁ Example Formula: =D5-$B$1
   ⦁ Result: 54 days.

# 2. Calculate Working Days Excluding Weekends and Holidays:

⦁ Used NETWORKDAYS to calculate the weekdays between the start date and the deadline date, excluding weekends and holidays (from a given range).
⦁ Example Formula: =NETWORKDAYS($B$1,D5,$J$5:$J$26)
⦁ Result: 37 weekdays.

# 3. Extract Month from Deadline Date:

⦁ Used MONTH to extract the month from a given date.
⦁ Example Formula: =MONTH(D5)
⦁ Result: 7 (July).

# 4. Extract Year from Deadline Date:

⦁ Used YEAR to extract the year from a given date.
⦁ Example Formula: =YEAR(D5)
⦁ Result: 2023.
______________________________________________________________________________________________________________________________________________________________________________________________________________________________________________
# Final Steps:

⦁ Autofill: Applied the double-click shortcut to copy the formulas down from row 5 to row 9.
⦁ Dynamic Date Calculation: Used the TODAY function in cell B1 to display the current date, which automatically updates daily.
______________________________________________________________________________________________________________________________________________________________________________________________________________________________________________
# Conclusion:

By using a variety of date functions, I successfully calculated the total calendar days, working days (excluding holidays and weekends), and extracted month and year information from the deadline date. This work can now be used for tracking project milestones.
______________________________________________________________________________________________________________________________________________________________________________________________________________________________________________
