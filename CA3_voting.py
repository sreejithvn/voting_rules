import openpyxl
import pprint
# import os

# Go to folder where excel file is located
# os.chdir('/Users/sreejith/Downloads')
# print(os.getcwd())

def generatePreferences(values):
    # loading the excel sheet
    wb = openpyxl.load_workbook(values)
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

    # For printing a dictionary properly
    pp = pprint.PrettyPrinter(indent=4, width = 120)

    print('\n\nPopulated preferences dictionary: ')
    pp.pprint(preferences_dict)
    print('\n\n')

    # Create a temporary dictionary to access the list index and values, as 
    temp_dict = {}
    for key, value in preferences_dict.items():
        # temp_dict = dict(enumerate(value, start=1))     # same output for temp_dict as below loop

        print(f"For agent {key}: get it's original list: \n{value}")
        for index, element in enumerate(value):
            # Creating a temp_dict, with index values of list as keys, and list elements as dictionary values
            temp_dict[index+1] = element    # index+1 to match the alternatives start count of 1

        print(f'From this original list create a temporary dictionary: \n{temp_dict} \n')

        # Updating preferences dictionary with sorted indexes (keys) of temp_dict based on it's values
        print(f'Get the sorted list of keys from temporary dictionary based on its values:\n {sorted(temp_dict, key=temp_dict.get)}')
        print(f'Reverse it to get descending order:\n {sorted(temp_dict, key=temp_dict.get)[::-1]}')
        preferences_dict[key] = sorted(temp_dict, key=temp_dict.get)[::-1]
        print("Overwrite preferences dictionary with these new index based values list for each agent(key)\n\n")

        # To get sorted values and matching keys as a tuple from dictionary
        # print(sorted(temp_dict.items(), key=lambda x:x[1])[::-1])

    print('\n\nSorted preferences dictionary: ')
    pprint.pprint(preferences_dict)
    return preferences_dict


generatePreferences('/Users/sreejith/Downloads/voting.xlsx')

# def dictatorship(preferenceProfile, agent) -> int
