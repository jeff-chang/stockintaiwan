# -*- coding: UTF-8 -*-
import requests
import pandas as pd
import configparser
from datetime import date
API_URL = "https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date=%s&stockNo=%s"
CONFIG_PATH = 'config.ini'

class STOCK(object):
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config.read(CONFIG_PATH)

    def getStockInfo(self, section):
        items = self.config.items(section)
        name = items[0][1]
        interval = items[1][1]
        start = items[2][1]
        end = items[3][1]
        return name, interval, start, end

    def requestURLs(self):
        total_section = self.config.sections()
        urls = {}
        for section in total_section:
            name, interval, start, end = self.getStockInfo(section)
            urls[section] = {"name": name, "urls":[] }
            if interval:
                today = date.today()
                year = today.year
                month = today.month
                for i in range(0, int(interval)):
                    if i != 0:
                        month = month - 1
                    if month <= 0:
                        year = year - 1
                        month = 12
                    url = API_URL % ("%s%02d01" % (year, month), section)
                    urls[section]['urls'].append(url)
        return urls

    def excelWriter(self, rawData=None):
        filename = date.today().strftime('%Y-%m-%d') + '.xlsx'
        from openpyxl import Workbook
        import os
        if os.path.exists(filename):
            pass
        else:
            new_wb = Workbook()
            new_wb.save(filename)
        cols = {'date':[], 'deal_share_num':[], 'the_price':[], 'start':[], 'highest':[], 'lowest':[], 'end':[], 'diff':[], 'deal_amount':[]}
        stockName = rawData[0]['title'].split(' ')[2]
        for i in rawData:
            for c in i['data']:
                cols['date'].append(c[0])
                cols['deal_share_num'].append(c[1])
                cols['the_price'].append(c[2])
                cols['start'].append(c[3])
                cols['highest'].append(c[4])
                cols['lowest'].append(c[5])
                cols['end'].append(c[6])
                cols['diff'].append(c[7])
                cols['deal_amount'].append(c[8])

        df = pd.DataFrame(cols)
        with pd.ExcelWriter(filename, engine='openpyxl', mode="a") as writer:
            df.to_excel(writer, sheet_name=stockName)
            writer.save()

    def crawller(self):
        stocks = self.requestURLs()
        for i in stocks.keys():
            rawData = []
            for j in stocks[i]['urls']:
                rawData.append(requests.get(j).json())
            self.excelWriter(rawData)

STOCK().crawller()
