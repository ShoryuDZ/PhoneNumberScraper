# Shoryu's PhoneNumberScraper
Get phone numbers for an excel sheet from Google

This script allows you to obtain phone numbers of a list of schools or businesses, given their name and suburb/town. 
The list must be in .xls or .xlsx, and formatted according to test.xlsx:
|   | A           | B      | C       | D       | E     |
|---|-------------|--------|---------|---------|-------|
| 1 | School Name | Suburb | Field A | Field B | Phone | 
| 2 | ...         | ...    |
| 3 | ...         | ...    |


Run in 3 steps:
1. Relocate your directory location in terminal to where your excel workbook has been saved: <br>
` cd G:\3. BIG DAY IN\BiG Day In Intern\BiG Day In Market Research\All Schools\ `
2. Run script in Python 3.x: <br>
` python3 script.py `
3. Enter exact filename and sheetname: <br>
```
  Workbook Name: test.xlsx
  Sheet Name: Sheet1
```

NOTE: Script will only be save spreadsheet once parsing and obtaining is complete, to prevent file corruption.
