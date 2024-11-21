import openpyxl
from pathlib import Path

cd  = Path(".")
path_to_file = cd / "Automating Worksheets.xlsx"

# Create a workbook

wb = openpyxl.load_workbook(path_to_file) # file already existing

# Create a new sheet

# wb.create_sheet(title="New Sheet")

# wb.save(path_to_file)

# delete a sheet 

for num in range (1,3):

    del wb[f"New Sheet{num}"]
    wb.save(path_to_file)

