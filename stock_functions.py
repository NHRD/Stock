import gspread
import json
from urllib import request
#import requests
from bs4 import BeautifulSoup
#from lxml import html
import re

def ticker_imp(row, worksheet):
    pos = "E" + str(row)
    ticker = worksheet.acell(pos).value
    return ticker

def get_Gross_margin(ticker):
    try:
        reg_url = "https://finviz.com/quote.ashx?t=" + ticker
        headers = {
            "User-Agent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:47.0) Gecko/20100101 Firefox/47.0"
            }
        req = request.Request(url=reg_url, headers=headers)
        response = request.urlopen(req)
        soup = BeautifulSoup(response, 'html.parser')
        data = soup.select("b")
        #if data[51].text != "-":
        try:
            data = re.findall("\d+\.\d+", data[51].text)
            data = float(data[0]) / 100
        #else:
        except IndexError:
            data = "-"
        response.close()
    except:
        data = "-"
    return data

def get_psr_ttm(ticker):
    try:
        reg_url = "http://quotes.morningstar.com/stockq/c-header?&t=" + ticker
        headers = {
            "User-Agent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:47.0) Gecko/20100101 Firefox/47.0"
            }
        req = request.Request(url=reg_url, headers=headers)
        response = request.urlopen(req)
        soup = BeautifulSoup(response, 'html.parser')
        data = soup.select(".gr_text1")
        try:
            data = re.findall("\d+\.\d+", data[9].text)[0]
        except IndexError:
            data = "-"
        response.close()
    except:
        data = "-"
    return data

def get_quarterly_sales(ticker):
    try:
        reg_url = "https://stocks.finance.yahoo.co.jp/us/quarterly/" + ticker
        headers = {
            "User-Agent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:47.0) Gecko/20100101 Firefox/47.0"
            }
        req = request.Request(url=reg_url, headers=headers)
        response = request.urlopen(req)
        soup = BeautifulSoup(response, 'html.parser')
        data = soup.select(".yjMS > td")
        try:
            data1 = re.findall("\d+\,\d+", data[6].text)
            data2 = re.findall("\d+\,\d+", data[7].text)
            data1 = data1[0].replace("," , "")
            data2 = data2[0].replace("," , "")
            data1 = int(data1)
            data2 = int(data2)
            data = round((data1 / data2) - 1, 4)
            data1 = data1 / 1000
            data2 = data2 / 1000
        except IndexError:
            data1, data2, data = "-", "-", "-" 
        response.close()
    except:
        data1, data2, data = "-", "-", "-" 
    return data, data1, data2

def get_yy_growth(ticker):
    try:
        reg_url = "https://stocks.finance.yahoo.co.jp/us/annual/" + ticker
        headers = {
            "User-Agent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:47.0) Gecko/20100101 Firefox/47.0"
            }
        req = request.Request(url=reg_url, headers=headers)
        response = request.urlopen(req)
        soup = BeautifulSoup(response, 'html.parser')
        data = soup.select(".yjMS > td")
        try:
            data1 = re.findall("\d+\,\d+", data[6].text)
            data1 = int(data1[0].replace("," , ""))
            data2 = re.findall("\d+\,\d+", data[7].text)
            data2 = int(data2[0].replace("," , ""))
            data = round((data1 / data2) - 1,4)
        except IndexError:
            data1, data2, data = "-", "-", "-" 
        response.close()
    except:
        data1, data2, data = "-", "-", "-" 
    #return data
    return data1, data2, data

def get_industory(ticker):
    try:
        reg_url = "http://quotes.morningstar.com/stockq/c-company-profile?&t=" + ticker
        headers = {
            "User-Agent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:47.0) Gecko/20100101 Firefox/47.0"
            }
        req = request.Request(url=reg_url, headers=headers)
        response = request.urlopen(req)
        soup = BeautifulSoup(response, 'html.parser')
        try:
            data = soup.select(".gr_text7")[1].text
            data = data.replace("\n", "").replace("   ", "")
        except IndexError:
            data = "-"
        response.close()
    except:
        data = "-"
    return data

#result = get_quarterly_sales("CRWD")
#print(result)

print("Import done.")

