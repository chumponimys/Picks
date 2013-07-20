from xlutils.copy import copy
import xlrd, xlwt
import datetime

def get_active_rows(ws):
    all_rows = ws.col_values(4, 2)
    active_rows = []
    for row in xrange(ws.nrows - 2):
        if (all_rows[row] != ""):
            active_rows.append(row + 2)
    return active_rows


def main():
    startTime = datetime.datetime.now()
    winners = ['CAR', 'TEN', 'NOS', 'DET', 'NYG', 'BUF', 'CLE', 'SFF', 'OAK',
               'BAL', 'SDC', 'TBB', 'GBP', 'SEA', 'PIT', 'DAL']
    game_num = 16
    path = "/Users/ludo/Desktop/Ari_Dev/Python/Picks/Picks/FP14_MASTER.xls"
    wb = xlrd.open_workbook(path)
    wbw = copy(wb)
    wsw = wbw.get_sheet(5)
    #nms = wb.sheet_names()
    ws = wb.sheet_by_name("Week (3)")
    #id_nums = ws.col_values(0, 2)
    all_rows = ws.col_values(4, 2)
    for row in xrange(ws.nrows - 2):
        if (all_rows[row] != ""):
            for col in xrange(2, game_num + 2):
                cell = ws.cell(row, col)
                if (str(cell.value) not in winners):
                    wsw.write(row, col, "BAD")
            #picks = ws.row_values(row + 2, 2, 24)
            #picks.append(ws.cell_value(row + 2, 24))
    
    wbw.save("/Users/ludo/Desktop/test.xls")
    return (datetime.datetime.now()-startTime)
#0:00:07.064465, 0:00:03.250335 --Print
#0:00:00.483543 --No Print
#0:00:00.476995

def test_average():
    times = []
    for x in xrange(20):
        times.append(main())
        print x
        
    print sum(times, datetime.timedelta(0)) / len(times)

print main()
