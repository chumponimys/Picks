# picks_computation.py
# Ari Cohen

import collections, os, getpass, re, urllib2, xlrd, xlwt
from xlutils.copy import copy
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

def export_from_mail(folder_loc, week_num, rs, ws):
    global FINAL_ROW_NUMBER
    alphab = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O",
              "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
    game_num = len(get_matchups(week_num))
    data = as_get_messages(week_num);
    id_nums = re.findall(r'\w*(\d\d\d\d)\w* \(postmaster', data)
    teams = split_list(re.findall(r'(\w+)-----', data), game_num)
    ws_ls = split_list(re.findall(r'\w\w: (\d+)', data), 2)
    blitz = re.findall(r'BLITZ: (\w+)', data)
    duplicates = [x for x, y in collections.Counter(id_nums).items() if y > 1]

    print "entering"
    print id_nums
    print
    excel_ids = rs.col_values(0, 2)
    for y in xrange(2, rs.nrows): #Go down the line looking at ID's
        if (excel_ids[y] == "stop"):
            break
        current_id = str(int(excel_ids[y]))
        #If we found a matching ID that's not a duplicate
        if ((current_id in id_nums) and (not current_id in duplicates)):
            #Enter data into Excel!
            print "Inputing: "+ str(current_id)
            indx = id_nums.index(current_id)
            id_nums.remove(current_id)
            picks = teams.pop(indx)
            current_blitz = blitz.pop(indx)
            current_ws_ls = ws_ls.pop(indx)
            for game in xrange(2, game_num + 2):
                setOutCell(ws, game, y + 2, picks[game - 2])
                (current_blitz == picks[game - 2])
            setOutCell(ws, game_num + 2, y + 2, current_ws_ls[0])
            setOutCell(ws, game_num + 3, y + 2, current_ws_ls[1])
            setOutCell(ws, game_num + 8, y + 2, current_blitz)

    print
    print "Bad ID's:"
    print id_nums

def check_picks(week_num, winners, rs, ws, wb):
    global FINAL_ROW_NUMBER
    form = xlwt.easyxf(
                 'font: name Gotham Narrow Book, height 140, color red;'
                 'borders: left thin, right thin, top thin, bottom thin;'
                 'pattern: pattern solid, pattern_fore_colour white, pattern_back_colour white'
                 )
    game_num = len(get_matchups(week_num))
    game_num = 16
    all_rows = rs.col_values(4, 2) #Get rows 2 and on...
    print all_rows
    for row in xrange(2, rs.nrows): #Go through them
        if (all_rows[row - 3] != ''):
            num_right = 0
            for col in (range(2, game_num + 2) + [game_num + 8]):
                val = str(rs.cell(row, col).value)
                if (val not in winners):
                    #style = xlwt.easyxf('pattern: pattern solid, fore_color white; font: color red;')
                    ws.write(row, col, val, form)
                    #setOutCell(ws, col, row, val)
                else:
                    num_right += 1
            ws.write(row, game_num + 5, num_right)
    #wb.save("/Users/aricohen/Desktop/test.xls")
        

def extract_and_check_forms(folder_loc):
    folder_loc = get_forms_folder()

def get_matchups(week_num):
    allInfo = urllib2.urlopen("http://www.nfl.com/schedules/2013/REG"+str(week_num))
    matchups = re.findall(r'\|(\w+) \: TBD_(\w+)', allInfo.read())
    matchups = [list(t) for t in matchups]
    for m in range(len(matchups)):
        for t in range(len(matchups[m])):
            team = matchups[m][t]
            if (team == "NE"):
                matchups[m][t] = "NEP"
            elif (team == "NO"):
                matchups[m][t] = "NOS"
            elif (team == "TB"):
                matchups[m][t] = "TBB"
            elif (team == "KC"):
                matchups[m][t] = "KCC"
            elif (team == "GB"):
                matchups[m][t] = "GBP"
            elif (team == "SF"):
                matchups[m][t] = "SFF"
            elif (team == "SD"):
                matchups[m][t] = "SDC"
    return matchups

def _getOutCell(outSheet, colIndex, rowIndex):
    """ HACK: Extract the internal xlwt cell representation. """
    row = outSheet._Worksheet__rows.get(rowIndex)
    if not row: return None

    cell = row._Row__cells.get(colIndex)
    return cell

def setOutCell(outSheet, col, row, value):
    """ Change cell value without changing formatting. """
    # HACK to retain cell style.
    previousCell = _getOutCell(outSheet, col, row)
    # END HACK, PART I

    outSheet.write(row, col, value)

    # HACK, PART II
    if previousCell:
        newCell = _getOutCell(outSheet, col, row)
        if newCell:
            idx = previousCell.xf_idx
            #idx.pattern_fore_colour = xlwt.Style.colour_map['dark_purple']
            newCell.xf_idx = idx
    # END HACK

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
    script = '/Users/'+getpass.getuser()+'/Desktop/Picks/Applescript/excel_write.scpt'
    Popen(["osascript", script, str(week_num), str(val), cell, str(blitz)], stdout=PIPE).communicate()

def as_mark_wrong(week_num, cell):
    script = '/Users/'+getpass.getuser()+'/Desktop/Picks/Applescript/excel_mark_wrong.scpt'
    Popen(["osascript", script, str(week_num), cell], stdout=PIPE).communicate()
