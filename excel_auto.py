import json
import win32com.client as win32

EXCEL_FILE = "C:/Users/petru/Downloads/PONTAJE Copie Test.xls"
JSON_FILE = "C:/Users/petru/Downloads/proba.json"
SHEET_NAME = "CONDICA"
START_ROW = 5
NAME_COLUMN = 2  # column B

def find_row(sheet, name):
    row = START_ROW
    while True:
        val = sheet.Cells(row, NAME_COLUMN).Value
        if val is None:
            return None
        if str(val).strip().upper() == name.upper():
            return row
        row += 1

def start_col_for_day(day):
    return 3 + (day - 1) * 7

def entry_col_for_day(day):
    return start_col_for_day(day) + 3

def exit_col_for_day(day):
    return start_col_for_day(day) + 5

excel = win32.Dispatch("Excel.Application")
#excel.Visible = False

wb = excel.Workbooks.Open(EXCEL_FILE)
sheet = wb.Worksheets(SHEET_NAME)

with open(JSON_FILE, "r", encoding="utf-8") as f:
    data = json.load(f)

for emp in data["employees"]:
    row = find_row(sheet, emp["name"])
    if not row:
        print("Name not found:", emp["name"])
        continue

    for day_str, hours in emp["days"].items():
        if "-" not in hours:
            continue

        day = int(day_str)
        entry, exit_time = hours.split("-")
        entry = entry.strip()
        exit_time = exit_time.strip()

        entry_col = entry_col_for_day(day)
        exit_col = exit_col_for_day(day)

        sheet.Cells(row, entry_col).Value = entry
        sheet.Cells(row, exit_col).Value = exit_time

wb.Save()
wb.Close()
excel.Quit()

print("Done (XLS updated correctly)")
