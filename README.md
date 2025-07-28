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





