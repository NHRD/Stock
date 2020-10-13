import gspread
import json
from urllib import request
from lxml import html
from stock_functions import get_Gross_margin, get_yy_growth, get_industory, get_psr_ttm, get_quarterly_sales, ticker_imp
#ServiceAccountCredentials：Googleの各サービスへアクセスできるservice変数を生成します。
from oauth2client.service_account import ServiceAccountCredentials 

#2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

#認証情報設定
#ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
credentials = ServiceAccountCredentials.from_json_keyfile_name('秘密鍵ファイル名.json', scope)

#OAuth2の資格情報を使用してGoogle APIにログインします。
gc = gspread.authorize(credentials)

#共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
SPREADSHEET_KEY = 'シートキー'

#共有設定したスプレッドシートのシート1を開く
workbook = gc.open_by_key(SPREADSHEET_KEY)
worksheet = workbook.worksheet('Future buy')

def data_out(row, col, export_value, worksheet):
    worksheet.update_cell(row, col, export_value)

for i in range(3, 60):
    #try:
    ticker = ticker_imp(i, worksheet)
    if ticker == "":
        break
    else:
        #gm = get_Gross_margin(ticker)
        #data_out(i, 13, gm, worksheet)
        #y2y = get_yy_growth(ticker)[2]
        #data_out(i, 11, y2y, worksheet)
        qus = get_quarterly_sales(ticker)
        data_out(i, 8, qus[1], worksheet)
        data_out(i, 12, qus[0], worksheet)
        #psrttm = get_psr_ttm(ticker)
        #data_out(i, 10, psrttm, worksheet)
        #ind = get_industory(ticker)
        #data_out(i, 4, ind, worksheet)
        #except:
        #    print(ticker + " error.")
        print(ticker + " done.")

print("Completed.")
    
    

#Tickerを受け取る
#import_value = int(worksheet.acell('A1').value)

#A1セルの値に100加算した値をB1セルに表示させる
