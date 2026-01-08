# Excel-Hours-Worked-Insertor

This script reads employee work hours from a JSON file and updates an Excel sheet with entry and exit times for each day.

## Requirements

- Python 3.x
- `pywin32` library (for Excel automation)
- Windows OS (requires Excel installed)

Install `pywin32` if you don’t have it:

```bash
pip install pywin32
```

## Usage
1.Update the script with the correct file paths:

```python
EXCEL_FILE = "C:/path/to/your/excel.xls"
JSON_FILE = "C:/path/to/your/data.json"
SHEET_NAME = "SheetName"
```

2.Make sure your JSON file has the following structure:

```json
{
  "employees": [
    {
      "name": "Employee Name",
      "days": {
        "1": "09:00 - 17:00",
        "2": "09:30 - 17:30"
      }
    }
  ]
}
```

3.Run the script:
```bash
python your_script.py
```

4.The Excel sheet will be updated with entry and exit times for each employee.
Note: Employee names in the JSON must match exactly the names in the Excel sheet.

## Notes
-The script starts reading Excel rows from row 5 and assumes names are in column B.

-Entry and exit columns are calculated automatically based on the day number.

-Make sure Excel is closed before running the script to avoid conflicts.

This script automates one of my personal excel tasks,modify the values of
```python
def start_col_for_day(day):
  return 3 + (day - 1) * 7
def entry_col_for_day(day):
  return start_col_for_day(day) + 3
def exit_col_for_day(day):
  return start_col_for_day(day) + 5
```
so it is useful for your own sheet.

### Tip
Use an OCR either from Python like `pytesseract` (`pip install pytesseract`) or any AI with image reading capabilities (Claude works best from what I see) to extract data from hand-written notes to fully automate the process
