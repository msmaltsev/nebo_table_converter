#!/usr/bin/env/python3
# -*- coding: utf8 -*-
import os
import openpyxl as op
import xlrd

import estet
import ferrum
import klyuvenkov
import magia
import neolog
import agold
from common_functions import *


# suppl_func_dict = {u'тд эстет':'estet',
#                    u'юк эстет':'estet',
#                    u'феррум':'ferrum',
#                    u'клювенков':'klyuvenkov',
#                    u'паршина':'magia',
#                    u'магкаева':'magia',
#                    u'рудакова':'magia',
#                    u'нео-лог':'neolog',
#                    u'а-голд':'agold',
#                    u'киселева':'magia',
#                    u'киселёва':'magia',
# 				   u'бижу трезор':'estet'}



# suppl_code_dict = {u'тд эстет':'b00000016',
#                    u'юк эстет':'b00000017',
#                    u'феррум':'b00009910',
#                    u'клювенков':'b00009923',
#                    u'паршина':'b00000005',
#                    u'магкаева':'b00009905',
#                    u'рудакова':'b00009716',
#                    u'нео-лог':'b00009920',
#                    u'а-голд':'b00009749',
#                    u'Индивидуальный предприниматель Рудакова Елена Вениаминовна':'b00009716',
#                    u'Индивидуальный предприниматель Магкаева Алина Ацамазовна':'b00009905',
#                    u'киселева':'b00010213',
#                    u'киселёва':'b00010213',
# 				   u'бижу трезор':'b00000014'}

# ctrct_code_dict = {u'тд эстет':'10378',
                #    u'юк эстет':'14',
                #    u'феррум':'10080',
                #    u'клювенков':'10158',
                #    u'паршина':'2',
                #    u'магкаева':'10074',
                #    u'рудакова':'10174',
                #    u'нео-лог':'10162',
                #    u'а-голд':'9866',
                #    u'Индивидуальный предприниматель Рудакова Елена Вениаминовна':'10174',
                #    u'Индивидуальный предприниматель Магкаева Алина Ацамазовна':'10074',
                #    u'киселева':'10519',
                #    u'киселёва':'10519',
				#    u'бижу трезор':'11'}



def getContragentsList(table='contragent.xlsx'):
    suppl_func_dict = {}
    suppl_code_dict = {}
    ctrct_code_dict = {}
    sb = xlrd.open_workbook(table)
    ws = sb.sheet_by_index(0)
    col = 0
    row = 1
    while row < ws.nrows:
        try:
##            print(ws.cell_value(row, col))
            suppl_func_dict[ws.cell_value(row, col)] = ws.cell_value(row, 1)
            suppl_code_dict[ws.cell_value(row, col)] = ws.cell_value(row, 2)
            ctrct_code_dict[ws.cell_value(row, col)] = ws.cell_value(row, 5)
            row += 1
        except:
            break
    return suppl_func_dict, suppl_code_dict, ctrct_code_dict

def defineContragent(ws, cta=getContragentsList()[0].keys()):
    if type(ws) == dbfread.dbf.DBF:
        return u'клювенков'

    try:
        m = re.search(u'Название', ws.cell_value(3,0))
        if m is not None:
            return u'магкаева'
        else:
            pass
    except:
        pass
        
    if u'п/п' in ws.cell_value(0,0):
        return u'а-голд'
    else:
        for ct in cta:
            # print(ct)
            if findInWs(ct, ws, match=True):
                return ct
        
def getTables(subdir, d=os.getcwd()):
    extensions = ['.xls',
                  '.xlsx',
                  '.XLS',
                  '.XLSX',
                  '.dbf',
                  '.DBF']
    d = '%s\\%s'%(d,subdir)
    tables = []
    for f in os.listdir(d):
        if os.path.splitext(f)[-1] in extensions:
            tables.append(f)
    return tables


def getVat(ws):
    print('VAT FUNCTION')
    vn = findInWs('НДС', ws)
    print('bez nds', vn)
    print('END VAT FUNCTION')


def makeXml(ws, name, vat):
    result = u'<xml>\r\n'
    result += u'\t<ndoc>%s</ndoc>\r\n'%ws.cell_value(0,1)
    result += u'\t<ddoc>%s</ddoc>\r\n'%ws.cell_value(1,1)
    result += u'\t<supplier_code>%s</supplier_code>\r\n'%suppl_code_dict[ws.cell_value(2,1)]
    result += u'\t<contract_code>%s</contract_code>\r\n'%ctrct_code_dict[ws.cell_value(2,1)]
    if vat is not None:
        result += u'\t<vat>%s</vat>\r\n'%vat
    else:
        result += u'\t<vat>None</vat>\r\n'
    row = 4
    while row < ws.nrows:
        result += u'\t<stock>\r\n'
        result += u'\t\t<article></article>\r\n'
        result += u'\t\t<supplier_article>%s</supplier_article>\r\n'%ws.cell_value(row, 0)
        result += u'\t\t<name>%s</name>\r\n'%ws.cell_value(row, 1)
        result += u'\t\t<weight>%s</weight>\r\n'%ws.cell_value(row, 2)
        result += u'\t\t<size>%s</size>\r\n'%ws.cell_value(row, 3)
        result += u'\t\t<strich_code>%s</strich_code>\r\n'%ws.cell_value(row, 4)
        result += u'\t\t<price>%s</price>\r\n'%ws.cell_value(row, 5)
        result += u'\t\t<amount>%s</amount>\r\n'%ws.cell_value(row, 6)
        result += u'\t</stock>\r\n'
        row += 1    
        
    result += u'</xml>'
    resfix = re.sub(u'\.0<', u'<', result, flags=re.U)
    result = resfix
    f = codecs.open(name, 'w', 'utf8')
    f.write(result)
    f.close()

if __name__ == '__main__':

    suppl_func_dict, suppl_code_dict, ctrct_code_dict = getContragentsList()
    # print(suppl_func_dict)
    # print(suppl_code_dict)
    # print(ctrct_code_dict)
    
    tblz = getTables('source')
    counter = 0
    for t in tblz:
        print(t)
        ws = getWsData(t)
        # print('vat', getVat(ws))
        ct = defineContragent(ws)
        print(ct)
        exec(u'%s.main(t, from_="%s")'%(suppl_func_dict[ct],ct))
        counter += 1
        print()
    print('%s files processed'%counter)

    for t in getTables('results'):
        try:
            ws = getWsData(t, subdir = '\\results\\')
            makeXml(ws, 'xml_results\\'+t+'.xml', vat='None')
        except Exception as e:
    ##        print('cant make xml due to %s'%e)
            pass
