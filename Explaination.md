[ACC 201-Workbook.xlsx](https://github.com/user-attachments/files/17969734/ACC.201-Workbook.xlsx)

Here is a Accounting related assignment I had to do involving excel. I will explain in detail what excel hacks I used to accelerate my assignement process:

**Step 1: Record Financial Data (Journal Entries)**
1. Input Journal Entries (Accuracy and Completeness)
   - Used the "Journal Entries" tab to input transactions using the provided accounting data appendix.
   - Excel Tools/Shortcuts:
     - Used AUTOFILL (Ctrl+D
 or drag the fill handle) to copy debit/credit labels or repetitive data.
     - Ensured descriptions align using text alignment options or Alt+H+A+C
 (Center Align).
     - Applied number formatting using Ctrl+Shift+1
 to ensure consistent display for monetary values.

2. Transfer to T-Accounts (Ledger Posting)
   - Posted entries from the journal to ledger T-accounts in the "T-Accounts" tab.
   - Excel Tools/Shortcuts:
     - Used Filter Tool (Ctrl+Shift+L
) to quickly filter transactions for specific accounts.
     - Copy and paste data using Ctrl+C
 and Ctrl+V
 for efficiency.
     - Used SUM (Alt+=
) to total debits and credits for each account.


**Step 2: Prepare the Unadjusted Trial Balance**
1. Summarize T-Account Balances
   - Transfered the ending balances from T-Accounts to the "Trial Balance" tab in the Unadjusted Trial Balance section.
   - Excel Tools/Shortcuts:
     - Use =SUM()
 or Alt+=
 to compute totals for each T-Account automatically.
     - Referenced T-Accounts directly using formulas (e.g., =TAccounts!B10
) to reduce manual entry errors.
     - Used Conditional Formatting (Alt+H+L
) to highlight discrepancies between debits and credits.

2. Verify Total Debits = Total Credits
   - Checked the Unadjusted Trial Balance totals to ensure debits equal credits.
   - Excel Tools/Shortcuts:
     - Used =SUM(range)
 to calculate column totals.
     - Used Custom Formatting (Ctrl+1 > Number > Custom
) to format trial balance cells appropriately (e.g., $#,##0.00).


**Step 3: Created Financial Statements**
Prepared financial statements in the following order:

1. Income Statement
   - Used the "Income Statement" tab to calculate net income (revenues - expenses) from the adjusted trial balance.
   - Excel Tools/Shortcuts:
     - Referenced adjusted trial balance figures using formulas (=TrialBalance!B5
).
     - Used AutoSum (Alt+=
) to total revenues and expenses quickly.
     - Used Bold Formatting (Ctrl+B
) to highlight net income.

2. Statement of Owner's Equity
   - Added the beginning balance of owner's equity, net income from the Income Statement, and subtract withdrawals.
   - Excel Tools/Shortcuts:
     - Linked net income directly using cell references (e.g., =IncomeStatement!B10
).
     - Used Named Ranges (Ctrl+F3
) to define specific account balances for clarity.
     - Applied simple subtraction formulas (e.g., =B5+B6-B7
) for changes to equity.

3. Balance Sheet (Assets and Liabilities)
   - Prepared the balance sheet by listing assets, liabilities, and owner's equity from the adjusted trial balance.
   - Ensured assets = liabilities + equity.
   - Excel Tools/Shortcuts:
     - Used =SUM()
 to compute subtotals for assets and liabilities.
     - Applied Indentation (Alt+H+6
 to increase or Alt+H+5
 to decrease) for better formatting.
     - Used Borders (Alt+H+B
) to delineate sections visually.


**Step 4: Enter Closing Entries**
1. Close Temporary Accounts
   - In the "Closing Entries" tab, transfered all revenue and expense account balances to the income summary account, and then closed the income summary to owner's equity.
   - Excel Tools/Shortcuts:
     - Used Cell Referencing to pull balances from the adjusted trial balance (=TrialBalance!C10
).
     - Used Ctrl+R
 to quickly copy formulas to the right across rows.

2. Verify Closing Entries
   - Ensured all temporary accounts have zero balances.
   - Excel Tools/Shortcuts:
     - Applied filters to zero out balances using Conditional Formatting (Alt+H+L
).
     - Used =IF()
 formulas (e.g., =IF(B5=0,"Closed","Check")
) to verify correctness.


**Step 5: Finalized Workbook**
1. Reviewed for Errors
   - Use Excel’s Spell Check (F7
) to review text descriptions.
   - Validated formulas using Trace Precedents (Alt+M+P
) or Trace Dependents (Alt+M+D
) to ensure proper linking between cells.

2. Format Workbook for Submission
   - Applied consistent formatting using Cell Styles (Alt+H+J
) to highlight headers and totals.
   - Used Page Layout view (Alt+W+P
) to ensure proper alignment for printing or display.

3. Summary Report Preparation
   - Created a summary report in a separate Excel tab or Word file. Highlight the financial performance and accuracy of the work.
   - Excel Tools/Shortcuts:
     - Used Data Validation (Alt+D+L
) to ensure consistent date ranges or entry values in the summary.
     - Used a PivotTable (Alt+N+V
) for a quick summary of financial data if needed.
