# Mastering The CoinMarketCap API with Python3

import xlsxwriter
import requests
import json

start = 1
f = 1

crypto_workbook = xlsxwriter.Workbook('cryptocurrencies.xlsx')
crypto_sheet = crypto_workbook.add_worksheet()

crypto_sheet.write('A1', 'Name')
crypto_sheet.write('C1', 'Symbol')
crypto_sheet.write('E1', 'Market Cap')
crypto_sheet.write('G1', 'Price')
crypto_sheet.write('I1', '24H Volume')
crypto_sheet.write('K1', 'Hour Change')
crypto_sheet.write('M1', 'Day Change')
crypto_sheet.write('O1', 'Week Change')

for i in range(10):
    ticker_url = 'https://api.coinmarketcap.com/v2/ticker/?structure=array&start=' + str(start)

    request = requests.get(ticker_url)
    results = request.json()
    data = results['data']

    for currency in data:
        rank = currency['rank']
        name = currency['name']
        symbol = currency['symbol']
        quotes = currency['quotes']['USD']
        market_cap = quotes['market_cap']
        hour_change = quotes['percent_change_1h']
        day_change = quotes['percent_change_24h']
        week_change = quotes['percent_change_7d']
        price = quotes['price']
        volume = quotes['volume_24h']

        crypto_sheet.write(f,0,name)
        crypto_sheet.write(f,2,symbol)
        crypto_sheet.write(f,4,str(market_cap))
        crypto_sheet.write(f,6,str(price))
        crypto_sheet.write(f,8,str(volume))
        crypto_sheet.write(f,10,str(hour_change))
        crypto_sheet.write(f,12,str(day_change))
        crypto_sheet.write(f,14,str(week_change))

        f += 1

    start += 100

crypto_workbook.close()
