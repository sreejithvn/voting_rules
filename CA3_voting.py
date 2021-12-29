import openpyxl
# import os

# Go to folder where excel file is located
# os.chdir('/Users/sreejith/Downloads')
# print(os.getcwd())

# loading the excel sheet
wb = openpyxl.load_workbook('/Users/sreejith/Downloads/voting.xlsx')
print(wb.sheetnames)

# Getting Sheet1
sheet = wb['Sheet1']

# Getting value in cell 'A1', but preferred method when using loops is: sheet.cell(row_num, col_num).value
# print(sheet['A1'].value)

# Initilizing the dictionary with agents as keys, based on max row in sheet
# preferences_dict = {key: [] for key in range(1, sheet.max_row+1)}

# Below method of initializing does not work, as each list references the other lists!! 
# (When using dict.fromkeys(..., default_value), make sure default_value is not of a type that requires a constructor (like list), 
# else you'll get references to the same object in all your buckets.)
# preferences_dict = dict.fromkeys(range(1, sheet.max_row+1), [])

preferences_dict = {}

# print('Initialized preferences dictionary:', preferences_dict)

# print('Max row: ', sheet.max_row, 'Max column: ', sheet.max_column)

for row in range(1, sheet.max_row+1):
    # print("\n\nEntered row :", row)

    key = row
    if key not in preferences_dict.keys():
        preferences_dict[key] = []

    # Looping through the entire column values for each row
    for col in range(1, sheet.max_column+1):
        # print('For row :', row, ', Entered column :', col)

        # Appending each column value to the list for each row
        preferences_dict[key].append(sheet.cell(row, col).value)
        
        # print(f'Appending value to preferences_dict[{key}]: with value in cell {row,col}')
        # print(f'preferences dict[{key}]: {preferences_dict[key]}')

        # print(sheet.cell(row, col).value)

print('\n\nPopulated preferences dictionary: ', preferences_dict)