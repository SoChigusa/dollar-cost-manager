#!/usr/local/bin/python3
# coding: utf-8
import yaml
import numpy as np
import datetime
from datetime import timedelta

def openGoogleForm(cf):
    from selenium import webdriver
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait

    chrome = webdriver.Chrome(cf['webdriver'])

    # Googleフォームを開く
    chrome.get(cf['URL']+'&entry.'+cf['entry-expected-cost']+'='+cf['expected-cost'])

    # タブが閉じられるのを待つ
    WebDriverWait(chrome, 60*60*24).until(lambda d: len(d.window_handles) == 0)

    # 終了処理
    chrome.quit()

def readFromSpread(cf):
    import gspread
    import json

    #ServiceAccountCredentials：Googleの各サービスへアクセスできるservice変数を生成します。
    from oauth2client.service_account import ServiceAccountCredentials

    #2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

    #認証情報設定
    #ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
    credentials = ServiceAccountCredentials.from_json_keyfile_name(cf['json'], scope)

    #OAuth2の資格情報を使用してGoogle APIにログインします。
    gc = gspread.authorize(credentials)

    #共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
    SPREADSHEET_KEY = cf['spreadsheet-key']

    #共有設定したスプレッドシートのシート1を開く
    worksheet = gc.open_by_key(SPREADSHEET_KEY).sheet1

    #セルの値を受け取る
    c2 = worksheet.col_values(2)
    c3 = worksheet.col_values(3)

    # 次の購入日付を4列目セルに書き込み
    td = np.floor(30 * float(c3[-1]) / float(c2[-1]))
    next = today + timedelta(days=td)
    strnext = next.strftime('%Y/%m/%d')
    worksheet.update_cell(len(c2), 4, strnext)
    print('next purchase date is set to be '+strnext+'!!')

# 設定ファイルの読み込み
with open('config.yml', 'r') as yml:
    config = yaml.load(yml, Loader=yaml.SafeLoader)

today = datetime.date.today() # 今日の日付
openGoogleForm(config)
readFromSpread(config)
