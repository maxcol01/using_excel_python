import re

class Data:
    def __init__(self):
        self.ds = ""
        self.rgn_ds = ""

    def prompt_user(self, tag):
        self.ds = input(f"Please enter the {tag} dataset for comparison!: ")
        self.rgn_ds = input("Please enter a range of cells you want to compare! (e.g., A11:E11): ")
        return self.ds, self.rgn_ds
    
    
    def check_input(self,worbook):
        pattern = r"^[A-Z]\d{1,}:[A-Z]\d{1,}$"
        if self.ds in worbook.sheetnames and re.search(pattern, self.rgn_ds):
            return True
        else:
            print("Wrong input ! Please correct it !")
            return False
    
    def collect_size(self, workbook):
        ws = workbook[self.ds]
        rgn = ws[self.rgn_ds]
        n_row = len(rgn)
        n_col = len(rgn[0])
        return n_row, n_col