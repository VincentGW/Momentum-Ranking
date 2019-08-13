# functions to be imported to the main module

import xlwings as xw
import csv
import datetime
import os

thisyear = str(datetime.datetime.now().year)
thismonth = str(datetime.datetime.now().month)
thisday = str(datetime.datetime.now().day)
otheryear = str(datetime.datetime.now().year - 2)

if len(thismonth) == 1:
    thismonth = '0'+thismonth

if len(thisday) == 1:
    thisday = '0'+thisday

datefolder = "%s-%s-%s"%(thisyear,thismonth,thisday)

excel_file = xw.Book('MomentumCalc.xlsx')
excel_file_dos = xw.Book('MomentumRank.xlsx')
f_calc = excel_file.sheets['MomentumCalc']
f_rank = excel_file_dos.sheets['MomentumRank']

def manifest_pricelist(tickerarg):
    filename = tickerarg+"-"+datefolder
    with open(os.getcwd()+r"/"+filename+".csv",'r') as f_from:
        readCSV = csv.reader(f_from, delimiter=',')
        pricelist = []
        for row in readCSV:
            pricelist.append(row[1])

        return(pricelist)

def chuck_csv_data(arg):
    data = arg
    for x in data:
        myrange = 'C'+str((data.index(x)+1))
        f_calc.range(myrange).value = x

def transfer(num, ticker):
    num += 1
    tickerloc = 'B'+str(num)
    r90loc = 'C'+str(num)
    r125loc = 'D'+str(num)
    r250loc = 'E'+str(num)
    adjslope90 = f_calc.range('H509').value
    adjslope125 = f_calc.range('L509').value
    adjslope250 = f_calc.range('P509').value
    f_rank.range(tickerloc).value = ticker
    f_rank.range(r90loc).value = adjslope90
    f_rank.range(r125loc).value = adjslope125
    f_rank.range(r250loc).value = adjslope250

def load_series(stock):
    chuck_csv_data(manifest_pricelist(stock))