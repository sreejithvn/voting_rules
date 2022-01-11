import openpyxl
import pprint

from openpyxl.descriptors.base import Length
# import os

# Go to folder where excel file is located
# os.chdir('/Users/sreejith/Downloads')
# print(os.getcwd())


def generatePreferences(values):

    # Initilizing the dictionary with agents as keys, based on max row in sheet
    # preferences_dict = {key: [] for key in range(1, sheet.max_row+1)}

    preferences_dict = {}

    for row in range(1, values.max_row+1):
        agent = row
        if agent not in preferences_dict.keys():
            preferences_dict[agent] = []

        # Looping through the entire column values for each row
        for col in range(1, values.max_column+1):
            # Appending each column value to the list for each row
            preferences_dict[agent].append(values.cell(row, col).value)

    

    # print('\n\nPopulated preferences dictionary: ')
    # pp.pprint(preferences_dict)

    # Create a temporary dictionary to access the list index and values
    temp_dict = {}
    for agent, alternatives_list in preferences_dict.items():
        # temp_dict = dict(enumerate(value, start=1))

        for index, element in enumerate(alternatives_list):
            # Creating a temp_dict, with index values of list as keys, and list elements as dictionary values
            temp_dict[index+1] = element


        # To get sorted values from dictionary
        # print(sorted(temp_dict.values())[::-1])

        # To get only matching keys of sorted values from dictionary
        # print(sorted(temp_dict, key=temp_dict.get)[::-1])

        # reverse=True, for equal values, gets the key with lowest value first ([::1], sorts that issue)
        # print(sorted(temp_dict, key=temp_dict.get, reverse=True))

        # Updating preferences dictionary with sorted indexes (keys) of temp_dict based on it's values
        preferences_dict[agent] = sorted(temp_dict, key=temp_dict.get)[::-1]

        # To get sorted values and matching keys as a tuple from dictionary
        # print(sorted(temp_dict.items(), key=lambda x:x[1])[::-1])

    # print('\n\nSorted preferences dictionary: ')
    # pprint.pprint(preferences_dict)
    return preferences_dict

if __name__ == '__main__':
    wb = openpyxl.load_workbook('/Users/sreejith/Downloads/voting.xlsx')
    sheet = wb['Sheet1']

    # For printing a dictionary properly
    pp = pprint.PrettyPrinter(indent=4, width = 120)

    # pprint.pprint(generatePreferences(sheet))

# pprint.pprint(generatePreferences(sheet))

def dictatorship(preferenceProfile, agent):
     return preferenceProfile[agent][0]

# for agent in range(1, 9):
#     print(f'For agent {agent}: {dictatorship(generatePreferences("/Users/sreejith/Downloads/voting.xlsx"), agent)}')

def get_max_list(temp_dict):
    list_of_keys = []
    max_value = max(temp_dict.values())
    for key, value in temp_dict.items():
        if value == max_value:
            list_of_keys.append(key)
    # print(list_of_keys)
    return list_of_keys     


def get_min_list(temp_dict):
    list_of_keys = []
    min_value = min(temp_dict.values())
    for key, value in temp_dict.items():
        if value == min_value:
            list_of_keys.append(key)
    # print(list_of_keys)
    return list_of_keys  


def tiebreak_output(list, tieBreak, preferences):
    print('Checkout preferences', preferences)
    if tieBreak == 'max':
        print(max(list))
    elif tieBreak == 'min':
        print(min(list))
    else:
        agent = tieBreak
        temp_dict = {}
        for item in list:
            temp_dict[item] = preferences[agent].index(item)
        print(preferences[agent])
        print(temp_dict)
        print(min(temp_dict, key=temp_dict.get))


def scoringRule(preferences, scoreVector, tieBreak):
    # Check if sortVector length is same as the number of alternatives
    if len(scoreVector) != len(preferences[1]):
        print("Incorrect input")
        return False

    # sort score vector
    scoreVector.sort(reverse=True)
    scoring_dict = {}
    for alternatives_list in preferences.values():
        for index, alternative in enumerate(alternatives_list):
            if alternative not in scoring_dict:
                scoring_dict[alternative] = 0
            else:
                scoring_dict[alternative] += scoreVector[index]

    print(scoring_dict)
    list_of_keys = get_max_list(scoring_dict)
    print(list_of_keys)
    return tiebreak_output(list_of_keys, tieBreak, preferences)


def plurality(preferences, tieBreak):
    temp_dict = {}
    for key, value in preferences.items():
        temp_dict[key] = value[0]

    print(temp_dict)

    temp_dict2 = {}
    for value in temp_dict.values():
        if value not in temp_dict2.keys():
            temp_dict2[value] = 1
        else:
            temp_dict2[value] += 1

    print(temp_dict2)
    # print(max(temp_dict2, key=temp_dict2.get))
    # print(max(temp_dict2, key=lambda key: temp_dict2[key]))

    list_of_keys = get_max_list(temp_dict2)
    return tiebreak_output(list_of_keys, tieBreak, preferences)
    

def veto(preferences, tieBreak):
    # initialise preferences_veto with keys as alternatives, and with values as 0
    preferences_veto = dict.fromkeys(preferences[1], 0)
    # print(preferences_veto)
    print(preferences.values())
    # loop through each list of alternatives, and add points based on veto rule
    for alternative_list in preferences.values():
        # Add one to the value for every 'alternative', except for the last one
        for item in alternative_list[:-1]:
            # if item in preferences_veto.keys():   # check can be removed as all keys are already initialised and the same
            preferences_veto[item] += 1
    print(preferences_veto)
    # get list of winners (with most points)
    list_of_keys = get_max_list(preferences_veto)
    print(list_of_keys)
    # return the winner based on the tieBreak rule
    return tiebreak_output(list_of_keys, tieBreak, preferences)
    

def borda(preferences, tieBreak):
    # initialise preferences_borda with keys as agents, and with values as 0
    preferences_borda = dict.fromkeys(preferences[1], 0)
    # print(preferences_borda)
    # get the number of alternatives
    num_alternatives = len(preferences[1])

    for value in preferences.values():
        # Add value for every preference item, based on their rank position
        for rank_position, item in enumerate(value, start=1):
            if item in preferences_borda.keys(): # check can be removed as all keys are already initialised and the same
                preferences_borda[item] += (num_alternatives - rank_position)
    print(preferences_borda)
    list_of_keys = get_max_list(preferences_borda)
    print(list_of_keys)
    return tiebreak_output(list_of_keys, tieBreak, preferences)
    

def harmonic(preferences, tieBreak):
    # initialise preferences_borda with keys as agents, and with values as 0
    preferences_harmonic = dict.fromkeys(preferences[1], 0)
    # print(preferences_harmonic)

    for value in preferences.values():
        # Add value for every preference item, based on their rank position
        for rank_position, item in enumerate(value, start=1):
            if item in preferences_harmonic.keys(): # check can be removed as all keys are already initialised and the same
                preferences_harmonic[item] += (1/rank_position)
    print(preferences_harmonic)
    list_of_keys = get_max_list(preferences_harmonic)
    print(list_of_keys)
    return tiebreak_output(list_of_keys, tieBreak, preferences)


def STV(preferences, tieBreak):
    temp_preferences = preferences
    # temp_dict = {key: 0 for key in range(1, len(temp_preferences[1]))}
    temp_dict = dict.fromkeys(temp_preferences[1], 0)
    print('Initial temp dict: ', temp_dict)
    while True:
        for value in temp_preferences.values():
            temp_dict[value[0]] += 1
            print('temp_dict updated: ', temp_dict)
        key_list = get_min_list(temp_dict)
        print('least frequency key list: ', key_list)
        if len(key_list) == len(temp_preferences[1]):
            print(f'len of key list: {len(key_list)} and len of temp preferences values: {len(temp_preferences[1])}')
            return tiebreak_output(key_list, tieBreak, preferences)
        else:
            # to update preferences values, by removing the least frequency keys(alternatives)
            for value in temp_preferences.values():
                for key in key_list:
                    value.remove(key)
            print(f'Updated temp preferences dict after removing key list {key_list}: \n{temp_preferences}')
            for key in key_list:
                temp_dict.pop(key, None)
            print(f'Updated temp dict after removing key list {key_list}: \n{temp_dict}')
            
        

# preferences_dictionary = {1: [4, 2, 1, 3], 
#                           2: [4, 3, 1, 2],
#                           3: [4, 3, 1, 2],
#                           4: [1, 3, 4, 2],
#                           5: [2, 3, 4, 1],
#                           6: [2, 1, 3, 4],
#                           7: [4, 2, 3, 1],
#                           8: [4, 2, 1, 3]}

# preferences_dictionary = {1: [4, 2, 1, 3], 
#                           2: [4, 3, 1, 2],
#                           3: [5, 3, 1, 4],
#                           4: [1, 3, 4, 2],
#                           5: [5, 3, 4, 1],
#                           6: [5, 1, 3, 4],
#                           7: [1, 2, 3, 4],
#                           8: [4, 2, 1, 3]}

preferences_dictionary = {1: [4, 2, 1, 3], 
                          2: [4, 3, 1, 2],
                          3: [2, 3, 1, 4],
                          4: [1, 3, 4, 2],
                          5: [2, 3, 4, 1],
                          6: [2, 1, 3, 4],
                          7: [1, 2, 3, 4],
                          8: [4, 2, 1, 3]}

# preferences_dictionary = {1: [2, 4, 1, 3], 
#                           2: [4, 3, 1, 2],
#                           3: [2, 3, 1, 4],
#                           }

# plurality(preferences_dictionary, 'max')
# plurality(preferences_dictionary, 'min')
# plurality(preferences_dictionary, 1)

# veto(preferences_dictionary, 'max')
# veto(preferences_dictionary, 'min')
# veto(preferences_dictionary, 1)

# borda(preferences_dictionary, 'max')
# borda(preferences_dictionary, 'min')
# borda(preferences_dictionary, 1)

# harmonic(preferences_dictionary, 'max')
# harmonic(preferences_dictionary, 'min')
# harmonic(preferences_dictionary, 1)

# score_vector = [0.123, 0.345, 0.543, 0.876]
# score_vector = [0.12, 0.34, 0.54, 0.87]
# score_vector = [1.51111, 2, 3.5, 3]

# scoringRule(preferences_dictionary, score_vector, 'max')
# scoringRule(preferences_dictionary, score_vector, 'min')
# scoringRule(preferences_dictionary, score_vector, 1)


# STV(preferences_dictionary, 'max')
# STV(preferences_dictionary, 'min')
# STV(preferences_dictionary, 1)

