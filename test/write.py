from __future__ import print_function
import datetime
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import yaml

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar']


def main():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
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

    # 設定ファイルの読み込み
    with open('config.yml', 'r') as yml:
        config = yaml.load(yml, Loader=yaml.SafeLoader)

    service = build('calendar', 'v3', credentials=creds)

    event = {
      'summary': '予定1',
      'location': 'Shibuya Office',
      'description': 'サンプルの予定',
      'start': {
        'dateTime': '2021-06-11T09:00:00',
        'timeZone': 'Japan',
      },
      'end': {
        'dateTime': '2021-06-11T17:00:00',
        'timeZone': 'Japan',
      },
    }

    event = service.events().insert(calendarId=config['calendar-id'],
                                    body=event).execute()
    print (event['id'])


if __name__ == '__main__':
    main()
