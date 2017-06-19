#!/usr/bin/env python2
# -*- coding: utf-8 -*-

import matplotlib.pyplot as plt
import xlrd

currency = str(raw_input("Please enter currency type: "))
if currency == "eur":
    cur_index = 0
elif currency == "xau":
    cur_index = 1
elif currency == "GBP":
    cur_index = 2
elif currency == "jpy":
    cur_index = 3
elif currency == "cad":
    cur_index = 4
else:
    print "Ops, you enter a wrong type"
    exit()

opt = str(raw_input("Do you want to set base line? (yes/no)\n"))
if opt == 'yes':
    bl = float(raw_input("Please enter base line: "))

bp_day = []
blbp_day = []
sp_day = []
blsp_day = []
date = []
index = []

data_day = xlrd.open_workbook('./data/data_day.xlsx')
table_day = data_day.sheets()[cur_index]
#table_day = data_day.sheet_by_name(u'currency') # load sheet by name
nrows_day = table_day.nrows

data_hour = xlrd.open_workbook('./data/data_hour.xlsx')
table_hour = data_hour.sheets()[cur_index]
nrows_hour = table_hour.nrows

def getdate(table, row):
    date = xlrd.xldate_as_tuple(table.cell(row, 2).value, 0)
    short_date = date[0:3]
    return short_date

def getbenchmark(row):
    Close = table_day.cell(row, 6).value
    Low = table_day.cell(row, 5).value
    High = table_day.cell(row, 4).value
    f1 = table_day.cell(row, 7).value
#    f2 = table_day.cell(row, 8).value
    f3 = table_day.cell(row, 9).value
    Bsetup = Low - f1 * (High - Close)
    Ssetup = High + f1 * (Close - Low)
#    Benter = (1 + f2) / 2 * (High + Low) - f2 * High
#    Senter = (1 + f2) / 2 * (High + Low) - f2 * Low
    Bbreak = Ssetup + f3 * (Ssetup - Bsetup)
    Sbreak = Bsetup - f3 * (Ssetup - Bsetup)
    return Bbreak, Sbreak

def autolabel(hists, local):
    for hist in hists:
        num = hist.get_height()
        if local == 1:
            plt.text(hist.get_x() + hist.get_width()/10.,  num + 0.1, '%s' %int(num))
        elif local == -1:
            plt.text(hist.get_x() + hist.get_width()/10.,  -(num + 1.3), '%s' %int(num))

i = 0
for r_day in range(nrows_day-1, 1, -1):
    bp = 0
    blbp = 0
    sp = 0
    blsp = 0
    short_date_day = getdate(table_day,r_day)

    for r_hour in range(nrows_hour-1-i, 1, -1):
        short_date_hour = getdate(table_hour, r_hour)

        if (short_date_day[0] > short_date_hour[0]):
            continue
        elif (short_date_day[0] == short_date_hour[0]) and \
        (short_date_day[1] > short_date_hour[1]):
            continue
        elif (short_date_day[0] == short_date_hour[0]) and \
        (short_date_day[1] == short_date_hour[1]) and \
        (short_date_day[2] > short_date_hour[2]):
            continue
        elif (short_date_day[0] == short_date_hour[0]) and \
        (short_date_day[1] == short_date_hour[1]) and \
        (short_date_day[2] == short_date_hour[2]):
            i += 1
            if r_day+1 >= nrows_day:
                continue
            else:
                Bbreak, Sbreak = getbenchmark(r_day+1)

                cur_high = table_hour.cell(r_hour, 4).value
                if (cur_high > Bbreak):
                    bp += 1
                if 'bl' in dir() and (cur_high - Bbreak > bl):
                    blbp += 1

                cur_low = table_hour.cell(r_hour, 5).value
                if (cur_low < Sbreak):
                    sp += 1
                if 'bl' in dir() and (Sbreak - cur_low > bl):
                    blsp += 1
        else:
            break
    if bp != 0 or sp != 0:
        bp_day.append(bp)
        sp_day.append(sp)
        date.append('%s/%s' %(short_date_day[1], short_date_day[2]))
        if 'bl' in dir():
            blbp_day.append(-blbp)
            blsp_day.append(-blsp)

for i in range(len(bp_day)):
    index.append(i)
bp_hist = plt.bar(tuple(index), tuple(bp_day), color = ('red'), \
                  label = ('bp'), width = 0.3)
sp_hist = plt.bar(tuple(index), tuple(sp_day), color = ('green'), \
                  label = ('sp'), width = -0.3)

autolabel(bp_hist, 1)
autolabel(sp_hist, 1)

if 'bl' in dir():
    blbp_hist = plt.bar(tuple(index), tuple(blbp_day), color = ('blue'), \
                            label = ('blbp'), width = 0.3)
    blsp_hist = plt.bar(tuple(index), tuple(blsp_day), color = ('yellow'), \
                            label = ('blsp'), width = -0.3)
    autolabel(blbp_hist, -1)
    autolabel(blsp_hist, -1)

plt.xticks(tuple(index), tuple(date))
plt.title('Arbitrage Points')
plt.legend()
plt.show()
