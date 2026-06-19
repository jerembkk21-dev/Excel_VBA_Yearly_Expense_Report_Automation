# Excel_VBA_Yearly_Expense_Report_Automation
Automated Excel reporting tool built with VBA to process multiple worksheets and create a yearly expense summary.

# 1) Overview

This project automates the creation of a yearly expense report using Excel VBA.

The workbook processes multiple worksheets containing monthly expenses, formats each sheet, calculates totals, and consolidates all data into a single "Yearly Report" worksheet.

This project was developed while completing the Udemy course:

Project Based Course on Excel VBA (Visual Basic for Applications) and Excel Macros

# 2) Features

- Automatically loops through all worksheets
- Inserts report headers
- Formats tables and currency values
- Calculates total expenses
- Copies data from each sheet
- Appends all information into a master worksheet
- AutoFits columns for improved readability

# 3) VBA Procedures

# a) FinalReportLoop

Processes every worksheet and consolidates all information into the yearly report.

# b) AddHeaders

Creates column headers:

- Division
- Category
- Jan
- Feb
- Mar
- Total Expense

# c) FormatData:

Formats:
- Header row
- Borders
- Colors
- Currency style

# d) AutoSum

Calculates the total expense for each worksheet and formats the result.

# 4) Technologies Used

- Microsoft Excel
- VBA (Visual Basic for Applications)
- Excel Macros

# 5) Example Workflow

1. Open workbook.
2. Run FinalReportLoop.
3. Each worksheet is formatted automatically.
4. Totals are calculated.
5. Data from all worksheets is copied into the "Yearly Report" sheet.
6. Final report is generated automatically.

# 6) Skills Demonstrated

- Loops (Do While)
- Variables
- Worksheet manipulation
- Copy and paste automation
- Relative cell navigation
- Formatting with VBA
- Excel object model
- Macro development

# 7) Screenshots

Source worksheets

![Source Data](Screenshots/excel_vba_source.png)

Formatted data

![Source Data](Screenshots/excel_vba_formated_data.png)

VBA Script

![Source Data](Screenshots/excel_vba_script.png)

Final yearly report

![Source Data](Screenshots/excel_vba_final_yearly_report.png)

