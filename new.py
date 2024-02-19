#!/usr/local/bin/python3
# coding: utf-8
import os
import yaml
import numpy as np
import datetime
from datetime import datetime, timedelta


def openGoogleForm(cf):
    from selenium import webdriver
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait

    cService = webdriver.ChromeService(executable_path=cf['webdriver'])
    chrome = webdriver.Chrome(service=cService)

    # Googleフォームを開く
    chrome.get(cf['URL']+'&entry.'+cf['entry-expected-cost'] +
               '='+cf['expected-cost'])

    # タブが閉じられるのを待つ
    WebDriverWait(chrome, 60*60*24).until(lambda d: len(d.window_handles) == 0)

    # 終了処理
    chrome.quit()


def readFromSpread(cf):
    import gspread

    startCol = int(cf['start-col'])

    # ServiceAccountCredentials：Googleの各サービスへアクセスできるservice変数を生成します。
    from oauth2client.service_account import ServiceAccountCredentials

    # 2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']

    # 認証情報設定
    # ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        cf['json'], scope)

    # OAuth2の資格情報を使用してGoogle APIにログインします。
    gc = gspread.authorize(credentials)

    # 共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
    SPREADSHEET_KEY = cf['spreadsheet-key']

    # 共有設定したスプレッドシートのシート1を開く
    worksheet = gc.open_by_key(SPREADSHEET_KEY).sheet1

    # セルの値を受け取る
    values = worksheet.get_all_records()
    prices = list(map(lambda l: l["Today's cost"], values))

    # 次回購入日付の計算
    startDate = datetime.strptime(
        values[startCol-2]["タイムスタンプ"], '%Y/%m/%d %H:%M:%S')
    totalCurrent = np.sum(prices[startCol-2:])
    expectDays = 30 * totalCurrent / \
        float(values[-1]["Expected cost per month"])
    next = startDate + timedelta(days=expectDays)

    # 次の購入日付および総コストを書き込み
    strnext = next.strftime('%Y/%m/%d')
    worksheet.update_cell(len(values)+1, 4, strnext)
    worksheet.update_cell(len(
        values)+1, 5, str(float(values[-2]["Total cost"]) + float(values[-1]["Today's cost"])))
    print('Next purchase date is set to be '+strnext+'!!')

    # カレンダー更新の必要性
    if values[-1]["Calendar update"] == "TRUE":
        return None
    else:
        worksheet.update_cell(len(values)+1, 6, True)
        return next


def writeToCalendar(cf, next):
    from googleapiclient.discovery import build
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials

    # カレンダーの更新が不要な場合
    if next == None:
        return

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
