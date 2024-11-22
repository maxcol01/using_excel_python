from pathlib import Path
import openpyxl

class Load:

    @staticmethod
    def load(file):
        cd = Path(".") / "data"
        file_to_path = cd / file

        # Loading workbook: here a single workbook containing the data (extension: consider different workbooks)
        try:
            wb = openpyxl.load_workbook(file_to_path)
            print("Workbook loaded with success !")
            # Possibilites for data comparison
            sheets = wb.sheetnames
            print(f"Here is the data you can pick up from: {"/".join(sheets)}")
            return wb
        except Exception as e:
            print(f"There is something wrong: {e}")

