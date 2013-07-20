# picks_interface.py
# Ari Cohen



def what_doing():
    while (True):
        print "1) Populate Cells"
        print "2) Check Against Winners"
        print "3) Clear Data"
        selected_funct = raw_input("Select a Function ('q' for quit) ")
        if (selected_funct == "1"):
            return PFunctions.POPULATE
        elif (selected_funct == "2"):
            return PFunctions.CHECK_WINNERS
        elif (selected_funct == "3"):
            return PFunctions.CLEAR_DATA
        elif (selected_funct in "qQ"):
            return PFunctions.QUIT
        else:
            invalid_input()

def are_you_sure():
    while (True):
        print "Are you sure you want to continue?"
        print "This action cannot be undone"
        cont = raw_input("(y/n)")
        if (cont in "yY"):
            return True
        elif (cont in "nN"):
            return False
        else:
            invalid_input()

def populate_type():
    while (True):
        print "1) Choose Existing Forms Folder"
        print "2) Export ALL Forms From Mail"
        selected_type = raw_input("Select a Type ('q' for quit)")
        if (selected_type == "1"):
            return PFunctions.FROM_FOLDER
        elif (selected_type == "2"):
            return PFunctions.FROM_MAIL
        elif (selected_type in "qQ"):
            return PFunctions.QUIT
        else:
            invalid_input()
    
def get_forms_folder():
    return raw_input("Folder Location (you can drag the folder directly here):")

def invalid_input():
    print "Invalid Input, Try Again"
    print

def check_week():
    while (True):
        the_week = raw_input("What week is it?: ")
        try:
            return int(the_week)
        except:
            invalid_input()
            continue

def get_winners(matchups):
    entering = True
    while (entering):
        winners = []
        entering = False
        print "Enter 1 or 2 for the winner"
        print "To restart, enter 'r'"
        for matchup in matchups:
            spacer = [" ", " "]
            for m in range(2):
                if (len(matchup[m]) < 3): spacer[m] = "  "
            winner = raw_input(str(matchup[0])+spacer[0]+"vs."+spacer[1]+str(matchup[1])+": ")
            if (winner in "rR"):
                entering = True
                break
            else:
                winners.append(matchup[int(winner) - 1])              
    #winners = ['CAR', 'TEN', 'NOS', 'MIN', 'PHI', 'NEP', 'MIA', 'SFF', 'NYJ',
               #'BAL', 'SDC', 'ATL', 'GBP', 'ARZ', 'PIT', 'WAS']    
    return winners
        
        

def get_dir():
    return os.path.dirname(os.path.abspath(__file__))

def get_excel_file():
    #return raw_input("Drag Excel File Here: ")
    return '/Users/aricohen/Desktop/Picks/FP14_MASTER.xls'

def get_save_loc(current_path):
    path = raw_input("Path to save new version to (blank to writeover curent): ")
    if (path == ""):
        return current_path
    return path

from picks_computation import *
