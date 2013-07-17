# picks.py
# Ari Cohen

from picks_interface import *
import os

class PFunctions():
    QUIT = -1
    POPULATE = 1
    CHECK_WINNERS = 2
    CLEAR_DATA = 3

    FROM_MAIL = 4
    FROM_FOLDER = 5


def main():
    week_num = check_week()
    chosen_funct = what_doing()
    if (chosen_funct == PFunctions.POPULATE):
        method = populate_type()
        forms_folder = get_forms_folder()
        if (method == PFunctions.FROM_FOLDER):
            all_forms = extract_and_check_forms(forms_folder)
        elif (method == PFunctions.FROM_MAIL):
            export_from_mail(forms_folder, week_num)
        else:
            print "quitting"
            return
            #Quit Code Here
        print "populating"
        #Populate Cells Code
    elif (chosen_funct == PFunctions.CHECK_WINNERS):
        matchups = get_matchups(week_num)
        
    elif (chosen_funct == PFunctions.CLEAR_DATA):
        #Clear Data Code
        if (are_you_sure()):
            clear_data()
            print "clearing"
    else:
        print "quitting"
        return
        #Quit Code Here
