
from pathlib import Path
import openpyxl

cd = Path(".")
path_to_file = cd / "Automating Worksheets.xlsx"
wb:openpyxl.Workbook = openpyxl.load_workbook(path_to_file)

ws  = wb.active
print(ws.title)

# Capture a specific sheet
ws = wb["Sheet3"]

# rename a sheet

ws.title = "Third Sheet"

wb.save(path_to_file)



