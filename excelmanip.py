# functions to be imported to the main module

import xlwings as xw
import csv
import datetime
import os

thisyear = str(datetime.datetime.now().year)
lastyear = str(datetime.datetime.now().year - 1)
otheryear = str(datetime.datetime.now().year - 2)
thismonth = str(datetime.datetime.now().month)
lastmonth = str(datetime.datetime.now().month - 1)
thisday = str(datetime.datetime.now().day)
datedone = False

if len(thismonth) == 1:
    thismonth = '0'+thismonth

if len(lastmonth) == 1:
    lastmonth = '0'+lastmonth

if len(thisday) == 1:
    thisday = '0'+thisday

datefolder = "%s-%s-%s"%(thisyear,thismonth,thisday)

excel_file = xw.Book('MomentumCalc.xlsx')
excel_file_dos = xw.Book('MomentumRank.xlsm')
sort_macro = excel_file_dos.macro('Sort')
average_macro = excel_file_dos.macro('Averages')
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
        print('Copying data point no. '+ str(data.index(x)+1) +' to MomentumCalc.xlsx in Excel...', end='\r')
        f_calc.range(myrange).value = x

def transfer(num, ticker, momentum):
    num += 2
    tickerloc = 'B'+str(num)
    r90loc = 'C'+str(num)
    r125loc = 'D'+str(num)
    r250loc = 'E'+str(num)
    genericloc = 'G'+str(num)
    adjslope90 = f_calc.range('H509').value
    adjslope125 = f_calc.range('L509').value
    adjslope250 = f_calc.range('P509').value
    f_rank.range(tickerloc).value = ticker
    f_rank.range(r90loc).value = adjslope90
    f_rank.range(r125loc).value = adjslope125
    f_rank.range(r250loc).value = adjslope250
    f_rank.range(genericloc).value = momentum
    sort_macro()

def import_date(tickerarg):
    global datedone
    if datedone == False:
        filename = tickerarg+"-"+datefolder
        with open(os.getcwd()+r"/"+filename+".csv",'r') as f_from:
            readCSV = csv.reader(f_from, delimiter=',')
            datelist = []
            for row in readCSV:
                datelist.append(row[0])
        for x in datelist:
            myrange = 'B'+str((datelist.index(x)+1))
            print('Creating date '+ str(datelist.index(x)+1) +' in MomentumCalc.xlsx in Excel...', end='\r')
            f_calc.range(myrange).value = x
        datedone = True
    else:
        pass

def load_series(stock):
    chuck_csv_data(manifest_pricelist(stock))