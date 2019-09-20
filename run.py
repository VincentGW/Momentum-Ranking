#    This tool takes two years of historical stock data and performs a "Momentum Ranking"
#    analysis on them, which involves calculating the slope of 90 to 250 logarithmic data points
#    and then ranking those values
#    The stock data is adjusted for dividends and stock splits and is provided by Koyfin.com
#    Libraries used include xlwings for handling the excel templates in which analysis is done
#    and BeautifulSoup for handling requests sent out to retrieve stock data
#    To learn more about the concept behind the Momentum Ranking technique and how it can aid in
#    securities trading, you should read Andreas Clenow's book: "Stocks on the Move" (2015)

import os
import json
import csv
import pickle
import time
import requests
import bs4
import excelmanip
from excelmanip import datefolder, thismonth, thisday, thisyear, lastmonth, lastyear, otheryear
from bs4 import BeautifulSoup as soup

list_of_stocks = []

if not os.path.exists(datefolder):
    print('Creating a new folder for today\'s ranking...')
    os.makedirs(datefolder)

def get_sp500_to_pickle():
    print('Gathering up-to-date list of S&P500 stock tickers...')
    wikiurl = requests.get('https://en.wikipedia.org/wiki/List_of_S%26P_500_companies')
    scrape = soup(wikiurl.text, "lxml")
    table = scrape.find('table', {'class':'wikitable sortable'})
    stocks = []
    for row in table.findAll('tr')[1:]:
        stock = row.findAll('td')[0].text
        stocks.append(stock)    
    stocks = list(map(lambda s: s.strip(), stocks))
    with open("sp500tickers.pickle",'wb') as f:
        pickle.dump(stocks, f)
    print('Tickers saved to binary file offline...')

def load_tickers_to_list():
    global list_of_stocks
    with open("sp500tickers.pickle",'rb') as f:
        list_of_stocks = pickle.load(f)

def get_prices_to_pickle():
    if os.path.basename(os.getcwd()) != datefolder:
        os.chdir(os.getcwd()+"\\"+datefolder)
    else:
        pass
    temp_dict = {}
    for ticker in list_of_stocks:
        print('Downloading two years of stock history for company ' + ticker + '...')
        time.sleep(.100)
        koyfin_url = 'https://api.koyfin.com/api/v2/commands/g/g.eq/%s?dateFrom=%s-%s-%s&dateTo=%s-%s-%s&period=daily'%(ticker,otheryear,thismonth,thisday,thisyear,thismonth,thisday)
        r = requests.get(koyfin_url)
        scrape = soup(r.text, "html.parser")
        priceDictionary = json.loads(str(scrape))
        table = priceDictionary["graph"]["data"]
        singleseries = []
        for i in range(0, len(table)):
            stup = (table[i][0], table[i][4])
            singleseries.append(stup)
        temp_dict.update({ticker:singleseries})
    with open("Series_Database.pickle",'wb') as f:
        pickle.dump(temp_dict,f)
    print('Stock history database saved to binary file offline...')

def twelve_month_pickle():
    if os.path.basename(os.getcwd()) != datefolder:
        os.chdir(os.getcwd()+"\\"+datefolder)
    else:
        pass
    momentumdict = {}
    for ticker in list_of_stocks:
        print('Downloading monthly stock history for company ' + ticker + '...')
        time.sleep(.100)
        koyfin_url = 'https://api.koyfin.com/api/v2/commands/g/g.eq/%s?dateFrom=%s-%s-%s&dateTo=%s-%s-%s&period=monthly'%(ticker,lastyear,lastmonth,thisday,thisyear,thismonth,thisday)
        r = requests.get(koyfin_url)
        scrape = soup(r.text, "html.parser")
        priceDictionary = json.loads(str(scrape))
        table = priceDictionary["graph"]["data"]
        singleseries = []
        for i in range(0, len(table)):
            stup = table[i][4]
            singleseries.append(stup)
        returnseries = []
        for i in range(0, len(singleseries) - 1):
            chain1 = int(i)
            chain2 = int(i + 1)
            x = (singleseries[chain2]-singleseries[chain1])/(singleseries[chain1])
            x += 1
            returnseries.append(x)
        gross = 1
        for n in returnseries:
            gross = gross * n
        momentum = format(gross, "2%")
        momentumdict[ticker] = momentum
    with open("Generic_Momentum_Database.pickle",'wb') as f:
        pickle.dump(momentumdict,f)
    print('Twelve month momentum database saved to binary file offline...')

def load_database_to_csv():
    if os.path.basename(os.getcwd()) != datefolder:
        os.chdir(os.getcwd()+"\\"+datefolder)
    else:
        pass
    with open("Series_Database.pickle",'rb') as f:
        series_database = pickle.load(f)
    for (key, values) in series_database.items():
        print('Writing data from binary file to csv for company ' + key + '...')
        myfile = '%s-%s.csv'%(key,datefolder)
        f = open(myfile,'w',newline='')
        writer = csv.writer(f)
        for value in values:
            writer.writerow(value)
        f.close

def main():
    print('Did you already download stock data today? y/n')
    if input() == 'n':
        get_sp500_to_pickle()
        load_tickers_to_list()
        get_prices_to_pickle()
        load_database_to_csv()
        twelve_month_pickle()
        pass
    else:
        load_tickers_to_list()
        os.chdir(os.getcwd()+"\\"+datefolder)
        pass
        excelmanip.average_macro()
    with open("Generic_Momentum_Database.pickle",'rb') as f:
        momentumdict = pickle.load(f)
    for stock in list_of_stocks:
        excelmanip.import_date(stock)
        excelmanip.load_series(stock)
        excelmanip.transfer(list_of_stocks.index(stock),stock,momentumdict[stock])
    excelmanip.average_macro()

main()