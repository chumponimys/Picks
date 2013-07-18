# picks_computation.py
# Ari Cohen

import collections, os, getpass, re, urllib2
from picks_interface import *
from subprocess import Popen, PIPE

########### UPDATE THIS VALUE ##############
FINAL_ROW_NUMBER = 108 #####################
############################################


class PFunctions():
    QUIT = -1
    POPULATE = 1
    CHECK_WINNERS = 2
    CLEAR_DATA = 3

    FROM_MAIL = 4
    FROM_FOLDER = 5

def clear_data():
    pass

def split_list(lst, chunks):
    return [lst[i:i+chunks] for i in range(0, len(lst), chunks)]

def export_from_mail(folder_loc, week_num):
    global FINAL_ROW_NUMBER
    alphab = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O",
              "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
    game_num = len(get_matchups(week_num))
    data = run_applescript(ASFunctions.MAIL, week_num);
    id_nums = re.findall(r'\w*(\d\d\d\d)\w* \(postmaster', data)
    teams = split_list(re.findall(r'(\w+)-----', data), game_num)
    ws_ls = split_list(re.findall(r'\w\w: (\d+)', data), 2)
    blitz = re.findall(r'BLITZ: (\w+)', data)
    duplicates = [x for x, y in collections.Counter(id_nums).items() if y > 1]
    
    for y in xrange(3, FINAL_ROW_NUMBER): #Go down the line looking at ID's
        current_id = as_get_val(week_num, y)
        #If we found a matching ID that's not a duplicate
        if ((current_id in id_nums) and (not current_id in duplicates)):
            #Enter data into Excel!
            indx = id_nums.index(current_id)
            picks = teams[indx]
            current_blitz = blitz[indx]
            for game in xrange(game_num):
                cell = alphab[game]+str(y)
                as_write_val(week_num, cell, picks[game],
                             (current_blitz == picks[game]))
            as_write_val(week_num, alphab[game_num], ws_ls[indx][0], False)
            as_write_val(week_num, alphab[game_num + 1], ws_ls[indx][1], False)
            as_write_val(week_num, alphab[game_num + 6], current_blitz, False)

    print duplicates

def check_picks(week_num, winners):
    global FINAL_ROW_NUMBER
    active_rows = []
    alphab = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O",
              "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
    game_num = len(get_matchups(week_num))
    imp_cells = range(0, game_num)
    imp_cells.append(game_num + 6)
    for y in xrange(3, FINAL_ROW_NUMBER): #check for active rows to save time
        cell = "E"+str(y)
        if (as_get_val(week_num, cell) != ""):
            active_rows.append(y)

    for y in active_rows: #Check picks...
        num_right = 0
        for game in imp_cells:
            cell = alphab[game]+str(y)
            if (not (as_get_val(week_num, cell) in winners)):
                as_mark_wrong(week_num, cell)
            else:
                num_right += 1
        cell = alphab[game_num + 3]+str(y)
        as_write_val(week_num, cell, num_right, false)
        

def extract_and_check_forms(folder_loc):
    folder_loc = get_forms_folder()

def get_matchups(week_num):
    allInfo = urllib2.urlopen("http://www.nfl.com/schedules/2013/REG"+str(week_num))
    matchups = re.findall(r'\|(\w+) \: TBD_(\w+)', allInfo.read())
    return matchups

def populate_cells(week_num, forms_folder):
    pass

def as_get_messages(week_num):
    script = '/Users/'+getpass.getuser()+'/Desktop/Picks/Applescript/mail.scpt'
    (messages, stderr) = Popen(["osascript", script, str(week_num)], stdout=PIPE).communicate()
    return messages

def as_get_val(week_num, cell):
    cell = "A"+str(cell)
    script = '/Users/'+getpass.getuser()+'/Desktop/Picks/Applescript/excel_read.scpt'
    (cell_val, stderr) = Popen(["osascript", script, str(week_num), cell], stdout=PIPE).communicate()
    return cell_val

def as_write_val(week_num, cell, val, blitz):
    script = '/Users/'+getpass.getuser()+'/Desktop/Picks/Applescript/excel_read.scpt'
    Popen(["osascript", script, str(week_num), str(val), cell, str(blitz)], stdout=PIPE).communicate()

def as_mark_wrong(week_num, cell):
    script = '/Users/'+getpass.getuser()+'/Desktop/Picks/Applescript/excel_mark_wrong.scpt'
    Popen(["osascript", script, str(week_num), cell], stdout=PIPE).communicate()
