#!/usr/local/bin/python3
# coding: utf-8
import yaml
import numpy as np
import datetime
from datetime import datetime, timedelta

# for Google Calendar access
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

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
    c1 = worksheet.col_values(1)
    c2 = worksheet.col_values(2)
    c3 = worksheet.col_values(3)
    c4 = worksheet.col_values(4)

    # 購入日付の計算
    # td : 月当たりのコストを用いてリスケールした基準買い付け日幅
    # td2 : 実際の買い付け日と基準日とのずれを補正する因子
    td = np.floor(30 * float(c3[-1]) / float(c2[-1]))
    expect = datetime.strptime(c4[-1], '%Y/%m/%d')
    current = datetime.strptime(c1[-1], '%Y/%m/%d %H:%M:%S')
    td2 = (expect - current).days

    # 次の購入日付を4列目セルに書き込み
    next = current + timedelta(days=td+td2)
    strnext = next.strftime('%Y/%m/%d')
    worksheet.update_cell(len(c2), 4, strnext)
    print('Next purchase date is set to be '+strnext+'!!')

    return next

def writeToCalendar(cf, next):
    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/calendar']

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('calendar', 'v3', credentials=creds)

    event = {
      'summary': 'ETF新規購入',
      'location': 'Charles Schwab',
      'description': 'Automatically scheduled by Dollar Cost Manager App',
      'start': {
        'dateTime': next.strftime('%Y-%m-%dT09:30:00'),
        'timeZone': 'America/New_York',
      },
      'end': {
        'dateTime': next.strftime('%Y-%m-%dT16:00:00'),
        'timeZone': 'America/New_York',
      },
    }

    event = service.events().insert(calendarId=config['calendar-id'],
                                    body=event).execute()
    print('Added to Google Calendar with ID=', event['id'])

# 更新を確認
os.system('git pull')

# 設定ファイルの読み込み
with open('config.yml', 'r') as yml:
    config = yaml.load(yml, Loader=yaml.SafeLoader)

openGoogleForm(config)
next = readFromSpread(config)
writeToCalendar(config, next)
