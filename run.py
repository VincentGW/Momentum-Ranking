#    This tool takes two years of historical stock data and performs a "Momentum Ranking"
#    analysis on them, which involves calculating the slope of 90 to 250 logarithmic data points
#    and then ranking those values
#    The stock data is adjusted for dividends and stock splits and is provided by Koyfin.com
#    Libraries used include xlwings for handling the excel templates in which analysis is done
#    and BeautifulSoup for handling requests sent out to retrieve stock data
#    To learn more about the concept behind the Momentum Ranking technique and how it can aid in
#    securities trading, it is recommended you read Andreas Clenow's book: "Stocks on the Move" (2015)

import os
import json
import csv
import requests
import bs4
import excelmanip
from excelmanip import datefolder, thismonth, thisday, thisyear, otheryear
from bs4 import BeautifulSoup as soup

if not os.path.exists(datefolder):
    print('Creating a new folder for today\'s ranking...')
    os.makedirs(datefolder)

list_of_stocks = []

def get_sp500():
    global list_of_stocks
    print('Gathering up-to-date list of S&P500 stock tickers...')
    wikiurl = requests.get('https://en.wikipedia.org/wiki/List_of_S%26P_500_companies')
    scrape = soup(wikiurl.text, "lxml")
    table = scrape.find('table', {'class':'wikitable sortable'})
    stocks = []
    for row in table.findAll('tr')[1:]:
        stock = row.findAll('td')[0].text
        stocks.append(stock)
    
    stocks = list(map(lambda s: s.strip(), stocks))
    print('Stocks successfully acquired')
    for x in stocks:
        list_of_stocks.append(x)

def savePrices(arg):
    print('Downloading two years of stock history for the company ' + arg + '...')
    ticker = arg
    koyfin_url = 'https://api.koyfin.com/api/v2/commands/g/g.eq/%s?dateFrom=%s-%s-%s&dateTo=%s-%s-%s&period=daily'%(ticker,otheryear,thismonth,thisday,thisyear,thismonth,thisday)
    r = requests.get(koyfin_url)
    scrape = soup(r.text, "html.parser")
    priceDictionary = json.loads(str(scrape))
    series = priceDictionary["graph"]["data"]
    iterable = []
    for i in range(0, len(series)):
        stup = (series[i][0], series[i][4])
        iterable.append(stup)

    if os.path.basename(os.getcwd()) != datefolder:
        os.chdir(os.getcwd()+"\\"+datefolder)
    else:
        pass
    myfile = '%s-%s.csv'%(ticker,datefolder)
    f = open(myfile,'w',newline='')
    writer = csv.writer(f)
    for tup in iterable:
        writer.writerow(tup)
    f.close

def main():
    get_sp500()
    for stock in list_of_stocks:
        savePrices(stock)
        excelmanip.load_series(stock)
        excelmanip.transfer(list_of_stocks.index(stock),stock)

main()
