Excel Data Validator (Nepali BS Date Aware)
A Python tool to validate Nepali financial Excel files (.xlsx, .xlsm, .xls) with support for Nepali calendar (BS) using nepali-datetime. It checks key fields like Member ID, Loan Dates, Maturity Dates, Share Amounts, and highlights errors with colors.

Features
Validate multiple Excel files at once

Highlights missing/duplicate IDs, invalid/future BS dates, incorrect Period/Duration types, invalid balances, and share amounts not divisible by 100

Supports .xlsx, .xlsm (via openpyxl) and .xls (via xlwings)

Color-coded highlights:
🔴 Red = critical errors
🟢 Green = duplicates
🔵 Blue = invalid balances
🟠 Orange = warnings (e.g., future birthdates)

Requirements
bash
Copy
Edit
pip install pandas openpyxl xlwings nepali-datetime
Note: xlwings requires Excel installed for .xls files.

Usage
Run the script:

bash
Copy
Edit
python validate_excel_files.py
Select Excel files via the file dialog. The script will validate, highlight issues, save, and open each file automatically.

Validation Rules Summary
Member ID/Name: not blank, unique, numeric

Loan Dates (BS): valid format, not future

Maturity Date: not before LoanIssueDate/AccountOpenDate

PeriodType/DurationType: single character

Balances/Amounts: numeric & > 0

Share Amount: divisible by 100

Closing Balance: must not be null

Notes
Works offline with GUI picker (tkinter)

Detects columns by name

Non-destructive highlighting in original files
