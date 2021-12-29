import openpyxl
# import os

# Go to folder where excel file is located
# os.chdir('/Users/sreejith/Downloads')
# print(os.getcwd())

# loading the excel sheet
wb = openpyxl.load_workbook('/Users/sreejith/Downloads/voting.xlsx')
# print(wb.sheetnames)

# Getting Sheet1
sheet = wb['Sheet1']

# Initilizing the dictionary with agents as keys, based on max row in sheet
# preferences_dict = {key: [] for key in range(1, sheet.max_row+1)}

preferences_dict = {}

for row in range(1, sheet.max_row+1):
    key = row
    if key not in preferences_dict.keys():
        preferences_dict[key] = []

    # Looping through the entire column values for each row
    for col in range(1, sheet.max_column+1):
        # Appending each column value to the list for each row
        preferences_dict[key].append(sheet.cell(row, col).value)
        
        # print(sheet.cell(row, col).value)

# print('\n\nPopulated preferences dictionary: ', preferences_dict)

temp_dict = {}
for key, value in preferences_dict.items():
    # print(dict(enumerate(value)))

    for i, element in enumerate(value):
        # Creating a temp_dict, with index values of list as keys, and list elements as dictionary values
        temp_dict[i+1] = element

    # print(temp_dict)

    # To get sorted values from dictionary
    # print(sorted(temp_dict.values())[::-1])

    # To get only matching keys of sorted values from dictionary
    # print(sorted(temp_dict, key=temp_dict.get)[::-1])

    # reverse=True, for equal values gets the key with lowest value first ([::1] sorts that issue)
    # print(sorted(temp_dict, key=temp_dict.get, reverse=True))

    # Updating preferences dictionary with sorted indexes (keys) of temp_dict based on it's values
    preferences_dict[key] = sorted(temp_dict, key=temp_dict.get)[::-1]

    # To get sorted values and matching keys as a tuple from dictionary
    # print(sorted(temp_dict.items(), key=lambda x:x[1])[::-1])


print('\n\nSorted preferences dictionary: ', preferences_dict)
