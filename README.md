📊 Excel Data Validator (Nepali BS Date Aware)
A Python tool for validating Excel files (.xlsx, .xlsm, and .xls) used in Nepali financial data migration and quality assurance workflows. This script performs intelligent validations on key fields such as Member ID, Loan Dates, Maturity Dates, Share Amounts, and more — with support for Nepali calendar (BS) using the nepali-datetime module.

🧩 Key Features
✅ Validate multiple Excel files in one go
✅ Highlights:

Missing or duplicate Member IDs and Names

Invalid or future BS dates (LoanIssueDate, BirthDate, MaturityDate, etc.)

Improper PeriodType or DurationType values

Non-numeric, zero, or negative balances

Share amounts not divisible by 100
✅ Supports .xlsx, .xlsm, and .xls formats
✅ Uses both openpyxl and xlwings as appropriate
✅ Colored highlighting (Red, Green, Orange, Blue) for easy visual QA

🛠️ Requirements
Make sure you have Python 3.8+ and install the following packages:

bash
Copy
Edit
pip install pandas openpyxl xlwings nepali-datetime
Note:
xlwings requires Microsoft Excel installed on your system to work with .xls files.

🗃️ File Support
Format	Description	Library Used
.xlsx	Standard Excel file	openpyxl
.xlsm	Macro-enabled Excel	openpyxl
.xls	Legacy Excel format	xlwings

📂 How to Use
Run the script using Python:

bash
Copy
Edit
python validate_excel_files.py
A file dialog will appear — select one or more Excel files to validate.

The script will:

Open each file

Validate and highlight errors

Save the updated file

Automatically open the file in Excel for review

🎯 Validation Rules
Field Type	Validation Logic
Member ID/Name	Must not be blank. IDs must be unique and numeric.
Loan Dates (BS)	Must be in correct format and not in the future.
Maturity Date	Must not be before LoanIssueDate or AccountOpenDate
PeriodType/DurationType	Must be a single character
Balance/Amount	Must be a number > 0. Highlights blue if invalid.
Share Amount	Must be divisible by 100
Closing Balance	Must not be null

🎨 Color Codes
Color	Meaning
🔴 Red	Missing, invalid, or critical error
🟢 Green	Duplicate ID/account number
🔵 Blue	Invalid balance or amount
🟠 Orange	Suspicious but not critical (e.g., future birthdate)

📌 Notes
Script supports Nepali BS Dates using nepali_datetime.

Works offline with GUI file picker using tkinter.

Columns are detected by name (e.g., "LoanIssueDate BS", "MaturityDateBS", "ShareAmt", etc.).

Entirely non-destructive — original file is updated with highlights only.

