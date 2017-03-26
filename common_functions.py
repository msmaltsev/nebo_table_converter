# -*- coding: cp1251 -*-
import xlrd
import xlwt
import os, re
import openpyxl as op
import codecs
import dbfread


def resultName(name):
    dot = name.index('.')
    before = name[:dot]
    after = name[dot:]
    return before + '_result'

def getWsData(wb_name, d=os.getcwd(), subdir='\\source\\'):
##    print(os.path.splitext(wb_name)[-1])
    if os.path.splitext(wb_name)[-1] == '.dbf' or os.path.splitext(wb_name)[-1] == '.DBF':
        ws = dbfread.DBF(d+subdir+wb_name)
    else:
        sb = xlrd.open_workbook(d+subdir+wb_name, logfile=open(os.devnull, 'w'))
        ws = sb.sheet_by_index(0)
    return ws

def findInWs(line, ws, regex=True, strt_col=0, strt_row=0, match=False):
    if regex == True:
        print('searchin by regexp')
    else:
        print('searching by substring')
    col = strt_col
    row = strt_row
    found = None
    regexp = re.compile(line)
    while not found:
        if col != 702 and row != 500:
            try:
                cell = ws.cell_value(row, col)
                if cell is not None:
                    if match:
                        cell = cell.lower()
                    if regex == True:
                        m = re.search(regexp, cell)
                        if m is not None:
    ##                        print('data found: %s'%cell)
                            found = True
                            return {'col':col, 'row': row}
                            break
                        else:
                            col += 1
                    else:
                        if line.lower() in cell.lower():
                            print('data found: %s'%cell)
                            found = True
                            return {'col':col, 'row': row}
                            break
                        else:
                            col += 1
                else:
                    col = 1
                    row += 1
            except:
                col = 1
                row += 1
        else:
            return found
            break
