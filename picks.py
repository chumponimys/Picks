# picks.py
# Ari Cohen

from picks_interface import *
from picks_computation import *
import os




def main():
    excel_path = get_excel_file()
    week_num = check_week()
    rb = xlrd.open_workbook(excel_path, formatting_info=True)
    wb = copy(rb)
    ws = wb.get_sheet(int(week_num) + 2)
    rs = rb.sheet_by_name("Week ("+str(week_num)+")")
    chosen_funct = what_doing()
    if (chosen_funct == PFunctions.POPULATE):
        method = populate_type()
        forms_folder = get_forms_folder()
        if (method == PFunctions.FROM_FOLDER):
            all_forms = extract_and_check_forms(forms_folder)                   
        elif (method == PFunctions.FROM_MAIL):
            export_from_mail(forms_folder, week_num, rs, ws)
        else:
            print "quitting"
            return
            #Quit Code Here
        print "populating"
        #Populate Cells Code
    elif (chosen_funct == PFunctions.CHECK_WINNERS):
        matchups = get_matchups(week_num)
        winners = get_winners(matchups)
        check_picks(week_num, winners, rs, ws, wb)
    elif (chosen_funct == PFunctions.CLEAR_DATA):
        #Clear Data Code
        if (are_you_sure()):
            clear_data()
            print "clearing"
    else:
        print "quitting"
        return
        #Quit Code Here
    wb.save(get_save_loc(excel_path))

main()
    
