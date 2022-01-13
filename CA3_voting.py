    
import copy


def generatePreferences(values):

    preferences_dict = {}

    # Looping through each row in sheet
    for row in range(1, values.max_row+1):
        agent = row

        # Initialise an empty dictionary for each agent, 
        # for storing the alternative and the agent's corresponding preference
        preferences_dict[agent] = {}

        # Looping through the entire column values for each row
        for column in range(1, values.max_column+1):
            alternative = column
            cell_value = values.cell(row, column).value

            # Adding each alternative and the column value as a key pair to each agent's value
            preferences_dict[agent][alternative] = cell_value
            # preferences_dict[agent].update({alternative: cell_value})

        # A dictionary storing each alternative's number and the current agent's preference
        alternatives_dict = preferences_dict[agent]

        # Getting the preference order of alternatives, by sorting the dictionary keys(alternatives), 
        # based on it's values(preferences) for each agent
        alternatives_sorted_list = sorted(alternatives_dict, key=alternatives_dict.get)[::-1]

        # Updating preferences dictionary with the list of alternatives ordered based on the agent's preferences
        preferences_dict[agent] = alternatives_sorted_list

    return preferences_dict


def dictatorship(preferenceProfile, agent):
     return preferenceProfile[agent][0]

# For testing dictatorship function
# for agent in range(1, 7):
#     print(f'For agent {agent}: {dictatorship(generatePreferences("/Users/sreejith/Downloads/voting.xlsx"), agent)}')

def get_max_list(temp_dict):
    possible_winners = []
    max_value = max(temp_dict.values())
    for key, value in temp_dict.items():
        if value == max_value:
            possible_winners.append(key)
    # print(possible_winners)
    return possible_winners     


def get_min_list(temp_dict):
    possible_winners = []
    min_value = min(temp_dict.values())
    for key, value in temp_dict.items():
        if value == min_value:
            possible_winners.append(key)
    # print(possible_winners)
    return possible_winners  


def tiebreak_output(list, tieBreak, preferences):

    if tieBreak == 'max':
        return max(list)
    elif tieBreak == 'min':
        return min(list)
    else:
        agent = tieBreak
        temp_dict = {}
        for alternative in list:
            temp_dict[alternative] = preferences[agent].index(alternative)
        # print(preferences[agent])
        # print(temp_dict)
        return min(temp_dict, key=temp_dict.get)


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
    possible_winners = get_max_list(scoring_dict)
    print(possible_winners)
    return tiebreak_output(possible_winners, tieBreak, preferences)


def plurality(preferences, tieBreak):

    # Create a temporary dictionary to store the first position count of each alternative
    first_position_count = {}
    # Looping through the values of preferences, to get each agents, list of alternatives
    for alternatives_list in preferences.values():
        # get the most preferred alternative from the agents list of alternatives
        alternative = alternatives_list[0]
        # initialise the temporary dictionary for each new alternative, with count 1
        if alternative not in first_position_count:
            first_position_count[alternative] = 1
        else:
            # update the count of alternatives for each repeated occurence
            first_position_count[alternative] += 1
    print(first_position_count)

    # get the list of possible winners, with same high scores
    possible_winners = get_max_list(first_position_count)

    # find rhe winner using the tiebreak rule
    return tiebreak_output(possible_winners, tieBreak, preferences)
    

def veto(preferences, tieBreak):
    # initialise preferences_veto with keys as alternatives, and with values as 0
    preferences_veto = dict.fromkeys(preferences[1], 0)
    # print(preferences_veto)
    print(preferences.values())
    # loop through each list of alternatives, and add points based on veto rule
    for alternatives_list in preferences.values():
        # Add one to the value for every 'alternative', except for the last one
        for alternative in alternatives_list[:-1]:
            # if item in preferences_veto.keys():   # check can be removed as all keys are already initialised and the same
            preferences_veto[alternative] += 1
    print(preferences_veto)
    # get list of possible winners (with most points)
    possible_winners = get_max_list(preferences_veto)
    print(possible_winners)
    # return the winner based on the tieBreak rule
    return tiebreak_output(possible_winners, tieBreak, preferences)
    

def borda(preferences, tieBreak):
    # initialise preferences_borda with keys as alternatives, and with values as 0
    preferences_borda = dict.fromkeys(preferences[1], 0)
    # print(preferences_borda)
    # get the number of alternatives
    num_alternatives = len(preferences[1])

    for alternatives_list in preferences.values():
        # Add value for every preference item, based on their rank position
        for rank_position, alternative in enumerate(alternatives_list, start=1):
            if alternative in preferences_borda.keys(): # check can be removed as all keys are already initialised and the same
                preferences_borda[alternative] += (num_alternatives - rank_position)
    print(preferences_borda)
    possible_winners = get_max_list(preferences_borda)
    # print(possible_winners)
    return tiebreak_output(possible_winners, tieBreak, preferences)
    

def harmonic(preferences, tieBreak):
    # initialise preferences_borda with keys as alternatives, and with values as 0
    preferences_harmonic = dict.fromkeys(preferences[1], 0)
    # print(preferences_harmonic)

    for alternatives_list in preferences.values():
        # Add value for every preference item, based on their rank position
        for rank_position, alternative in enumerate(alternatives_list, start=1):
            if alternative in preferences_harmonic.keys(): # check can be removed as all keys are already initialised and the same
                preferences_harmonic[alternative] += (1/rank_position)
    print(preferences_harmonic)
    possible_winners = get_max_list(preferences_harmonic)
    print(possible_winners)
    return tiebreak_output(possible_winners, tieBreak, preferences)


def STV(preferences, tieBreak):
    temp_preferences = copy.deepcopy(preferences)
    # temp_preferences = preferences

    while True:

        # temporary dictionary for storing the frequency count of alternatives which resets to 0 after each updation
        alternative_frequency = dict.fromkeys(temp_preferences[1], 0)
        print('Temporary dictionary (for each round): ', temp_preferences)
        for alternatives_list in temp_preferences.values():
            alternative_frequency[alternatives_list[0]] += 1
        print('Current alternative frequency (for each round):',alternative_frequency)

        # list of least frequency keys
        alternatives_to_remove = get_min_list(alternative_frequency)
        print('Least frequent alternatives to be removed: ', alternatives_to_remove, '\n')
        # get the final set of alternatives to be removed
        if len(alternatives_to_remove) == len(temp_preferences[1]):
            print("Current temp preferences (first value): ", temp_preferences[1])
            print("Don't remove this list of alternatives, as length equals the length of the current preferences list: ", alternatives_to_remove)
            print('\nWinner is:')
            return tiebreak_output(alternatives_to_remove, tieBreak, preferences)
        else:
            for alternative in alternatives_to_remove:
                alternative_frequency.pop(alternative, None)

            # to update preferences values, by removing the least frequency keys(alternatives)
            for value in temp_preferences.values():
                for alternative in alternatives_to_remove:
                    value.remove(alternative)
            

def rangeVoting(values, tieBreak):
    valuation_sum = {}
    
    # Looping through each column in sheet
    for column in range(1, values.max_column+1):
        alternative = column
        if alternative not in valuation_sum.keys():
            valuation_sum[alternative] = 0

        # Looping through the entire row values for each column
        for row in range(1, values.max_row+1):
            # Appending each column value to the list for each row
            valuation_sum[alternative] += values.cell(row, column).value

    possible_winners = get_max_list(valuation_sum)
    print(possible_winners)
    return tiebreak_output(possible_winners, tieBreak, generatePreferences(values))







if __name__ == '__main__':
    import openpyxl
    import pprint

    workbook = openpyxl.load_workbook('/Users/sreejith/Downloads/voting.xlsx')
    sheet = workbook.active

    # For printing a dictionary properly
    pp = pprint.PrettyPrinter(indent=4, width = 120)

    print("Preferences Dictionary:")
    pprint.pprint(generatePreferences(sheet))
    print('\n')
    # pprint.pprint(rangeVoting(sheet, 'max'))
    # pprint.pprint(rangeVoting(sheet, 'min'))
    # pprint.pprint(rangeVoting(sheet, 1))


    preferences_dictionary = {1: [4, 2, 1, 3], 
                          2: [4, 3, 1, 2],
                          3: [4, 3, 1, 2],
                          4: [1, 3, 4, 2],
                          5: [2, 3, 4, 1],
                          6: [2, 1, 3, 4],
                          7: [4, 2, 3, 1],
                          8: [4, 2, 1, 3]}

    preferences_dictionary = {1: [4, 2, 1, 3], 
                            2: [4, 3, 1, 2],
                            3: [5, 3, 1, 4],
                            4: [1, 3, 4, 2],
                            5: [5, 3, 4, 1],
                            6: [5, 1, 3, 4],
                            7: [1, 2, 3, 4],
                            8: [4, 2, 1, 3]}

    preferences_dictionary = {1: [4, 2, 1, 3], 
                            2: [4, 3, 1, 2],
                            3: [2, 3, 1, 4],
                            4: [1, 3, 4, 2],
                            5: [2, 3, 4, 1],
                            6: [2, 1, 3, 4],
                            7: [1, 2, 3, 4],
                            8: [4, 2, 1, 3]}

    # preferences_dictionary = {1: [2, 4, 1, 3], 
    #                         2: [4, 3, 1, 2],
    #                         3: [2, 3, 1, 4],
    #                         }

    
    # score_vector = [0.123, 0.345, 0.543, 0.876]
    # score_vector = [0.12, 0.34, 0.54, 0.87]
    # score_vector = [1.51111, 2, 3.5, 3]

    # print(scoringRule(preferences_dictionary, score_vector, 'max'))
    # print(scoringRule(preferences_dictionary, score_vector, 'min'))
    # print(scoringRule(preferences_dictionary, score_vector, 1))
    
    # print(plurality(preferences_dictionary, 'max'))
    # print(plurality(preferences_dictionary, 'min'))
    # print(plurality(preferences_dictionary, 1))

    # print(veto(preferences_dictionary, 'max'))
    # print(veto(preferences_dictionary, 'min'))
    # print(veto(preferences_dictionary, 1))

    # print(borda(preferences_dictionary, 'max'))
    # print(borda(preferences_dictionary, 'min'))
    # print(borda(preferences_dictionary, 1))

    # print(harmonic(preferences_dictionary, 'max'))
    # print(harmonic(preferences_dictionary, 'min'))
    # print(harmonic(preferences_dictionary, 1))

    # print(STV(preferences_dictionary, 'max'))
    # print(STV(preferences_dictionary, 'min'))
    # print(STV(preferences_dictionary, 1))

    # USING SHEET
    # score_vector = [0.123, 0.345, 0.543, 0.876]
    # score_vector = [0.12, 0.34, 0.54, 0.87]
    # score_vector = [1.51111, 2, 3.5, 3]

    # print(scoringRule(generatePreferences(sheet), score_vector, 'max'))
    # print(scoringRule(generatePreferences(sheet), score_vector, 'min'))
    # print(scoringRule(generatePreferences(sheet), score_vector, 1))
    
    # print(plurality(generatePreferences(sheet), 'max'))
    # print(plurality(generatePreferences(sheet), 'min'))
    # print(plurality(generatePreferences(sheet), 1))

    # print(veto(generatePreferences(sheet), 'max'))
    # print(veto(generatePreferences(sheet), 'min'))
    # print(veto(generatePreferences(sheet), 1))

    # print(borda(generatePreferences(sheet), 'max'))
    # print(borda(generatePreferences(sheet), 'min'))
    # print(borda(generatePreferences(sheet), 1))

    # print(harmonic(generatePreferences(sheet), 'max'))
    # print(harmonic(generatePreferences(sheet), 'min'))
    # print(harmonic(generatePreferences(sheet), 1))

    # print(STV(generatePreferences(sheet), 'max'))
    # print(STV(generatePreferences(sheet), 'min'))
    # print(STV(generatePreferences(sheet), 1))