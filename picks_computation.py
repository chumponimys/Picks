# picks_computation.py
# Ari Cohen

import collections, os, getpass, re
from picks_interface import *
from subprocess import Popen, PIPE

def clear_data():
    pass

def export_from_mail(folder_loc, week_num):
    data = run_applescript(ASFunctions.MAIL, week_num);
    id_nums = re.findall(r'\w*(\d\d\d\d)\w* \(postmaster', data)
    teams = re.findall(r'(\w+)-----', data)
    ws_ls = re.findall(r'\w\w: (\d+)', data)
    blitz = re.findall(r'BLITZ: (\w+)', data)
    duplicates = [x for x, y in collections.Counter(id_nums).items() if y > 1]

    for z in range(len(id_nums)):
        #path = folder_loc+"/"+
        pass
    print data
    print
    print
    print id_nums
    print teams
    print ws_ls
    print blitz

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
