#!/usr/bin/python
# -*- coding: utf-8 -*-
import ast, io
from collections import OrderedDict
import xlwt
import csv
import sys, os
import json
import logging, locale

from datetime import datetime, time, date, timedelta, tzinfo
from xlutils.copy import copy # http://pypi.python.org/pypi/xlutils
from xlrd import open_workbook  # http://pypi.python.org/pypi/xlrd
from xlwt import easyxf # http://pypi.python.org/pypi/xlwt
__author__ = 'martinh'

_logger = logging.getLogger(__name__)

csvfile = "C:/Users/wb15/Dropbox/Shitkram/Umsatzanzeige_5418294198_20170821(1).csv"
xlsfile = "C:/Users/wb15/Dropbox/MOTB/GEMHHBDB.xls"
#f = open(sys.argv[1], "rb")
f = open(csvfile, "rb")
reader = f.readlines()
reader = [x.strip() for x in reader]
reader = [x.replace("'", "") for x in reader]
reader = [x.replace("\"", "") for x in reader]
startsignal = False

style = xlwt.XFStyle()

# font
font = xlwt.Font()
font.bold = False
style.font = font

LIST_OF_FIX_WORDS =  ("Miete und Nebenkosten", "Gem.-Kto", "eprimo", "Miete, Nebenkosten und Handy")
START_SHEET = -1 # 0 based (subtract 1 from excel row number)
style0 = xlwt.easyxf('font: name Arial, color-index black, bold off',num_format_str=u'#,##0.00')
style2 = xlwt.easyxf('font: name Arial, color-index black, bold off', num_format_str=u'#,##0.00 €')
style3 = xlwt.easyxf('font: name Times New Roman, color-index green, bold on', num_format_str=u'#,##0.00 €')
style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
#date_format = S3XLS.dt_format_translate(settings.get_L10n_date_format())
style = xlwt.XFStyle()
font = xlwt.Font()
font.name = "Lucida Handwriting"
style.font = font
style.number = u'0.00 €'
locale.setlocale(locale.LC_ALL, "")
def prepare_sheets(wb):
    ri = {}
    for i in range(1,18,1):
        w_sheet = wb.get_sheet(START_SHEET+i)  # the sheet to write to within the writable copy
        w_sheet.col(0).width = 3500
        w_sheet.col(1).width = 10000
        if i >12:
            w_sheet.col(2).width = 20000
            ctr = 0
            for m in char_range("D", "O"):
                str = 'SUM(%s5:%s300)' % (m,m)
                w_sheet.write(3, 3+ctr, xlwt.Formula(str), style3)
                ctr +=1
        else:
            w_sheet.col(2).width = 13000
        w_sheet.col(3).width = 4000
        w_sheet.write(3, 3, xlwt.Formula('SUM(D5:D300)'), style3)
        ri[i] = 4

    return ri

def char_range(c1, c2):
    """Generates the characters from `c1` to `c2`, inclusive."""
    for c in xrange(ord(c1), ord(c2)+1):
        yield chr(c)

def test_4_wrd(list, teststring):
    for t in list:
        if t in teststring:
            return True
    return False

rb = open_workbook(xlsfile, formatting_info=True)
r_sheet = rb.sheet_by_index(0) # read only copy to introspect the file
wb = copy(rb) # a writable copy (I can't read values out of this, only write to it)
all_sheet =  wb.get_sheet(START_SHEET+12+1)
einn_sheet =  wb.get_sheet(START_SHEET+12+2)
ausg_sheet = wb.get_sheet(START_SHEET+12+3)
row_index = prepare_sheets(wb)


def writesheet(month, datatab, amount, rowmonth=0):
    w_sheet = wb.get_sheet(START_SHEET + month)  # the sheet to write to within the writable copy
    w_sheet.write(row_index[month], 0, datatab[0], style1)
    w_sheet.write(row_index[month], 1, datatab[2], style0)
    w_sheet.write(row_index[month], 2, datatab[4], style0)
    w_sheet.write(row_index[month], rowmonth+3, amount, style2)  # float(datatab[5].replace(",", "."))
    row_index[month] += 1

for row in reader:
    if startsignal:
        datatab = row.split(";")
        #print datetime.strptime(datatab[0], "%d.%m.%Y"), datatab[2], datatab[4], datatab[5]
        month = datetime.strptime(datatab[0], "%d.%m.%Y").month
        amount = locale.atof(datatab[5])
        writesheet(12 + 1, datatab, amount, month-1)
        if amount <=0 and not test_4_wrd(LIST_OF_FIX_WORDS, datatab[4]):
            amount *= -1
            if amount > 250:  # Sonderausgabe ab 250 ocken
                writesheet(12 + 3, datatab, amount, month - 1)
            else:
                writesheet(month, datatab, amount)
        elif amount >0 and not test_4_wrd(LIST_OF_FIX_WORDS, datatab[4]):
            writesheet(12+2,datatab, amount, month-1)

        if test_4_wrd(LIST_OF_FIX_WORDS, datatab[4]):
            if amount < 0:
                amount *= -1
                writesheet(12 + 4, datatab, amount, month - 1)
            else:
                writesheet(12 + 5, datatab, amount, month - 1)


    if "Buchung" in row:
        startsignal=True

f.close()
#wb.save(xlsfile  + '.out' + os.path.splitext(xlsfile)[-1])
wb.save(xlsfile)

