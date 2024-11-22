
from prepare_data import Data
from load_file import Load

# Initialization of variables
FILE = "Employee Sales.xlsx"
wb = Load.load(FILE)

# Prepare the data we want to compare and check for their correct introduction by the user
is_dataset1 = False
is_dataset2 = False
ds1 = Data()
ds2 = Data()

while not is_dataset1:
    tag = "first"
    data1, rgn1 = ds1.prompt_user(tag)
    if ds1.check_input(wb):
        is_dataset1 = True
        n_row1, n_col1 = ds1.collect_size(wb)

while not is_dataset2:
    tag = "second"
    data2, rgn2 = ds2.prompt_user(tag)
    if ds2.check_input(wb):
        n_row2, n_col2 = ds2.collect_size(wb)
        if n_row1 == n_row2 and n_col1 == n_col2:
            is_dataset2 = True
            print("Your are all set !")
        else:
            print("There may be something wrong with the size of the second dataset \nMake sure the sizes match")



# Look for differences (from now on we suppose the size of the search window, i.e., the n_row and n_col, is the same)
ws1 = wb[data1]
ws2 = wb[data2]
rng_1 = ws1[rgn1]
rng_2 = ws2[rgn2]
wb.create_sheet("Differences")
wb.save(FILE) # preserve original data so store a new file
ws3 = wb["Differences"]
count = 0
for i in range(0, len(rng_1)):
    row1 = rng_1[i]
    row2 = rng_2[i]
    for j in range(0, len(row1)):
        if row1[j].value != row2[j].value:
            count += 1
            print(f"There is a difference in cell {i+1},{j+1}")
            ws3.cell(count,1).value = (f"In the first dataset we have this value at cell ({i+1},{j+1}): {row1[j].value}\t" 
                                   f"and similarly in the second data set we have {row2[j].value}" )
            wb.save(FILE)
            
print("#############################################")
print("#!!!! Checking of differences completed !!! #")
print("#############################################")