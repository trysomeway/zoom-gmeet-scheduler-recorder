from __future__ import print_function
import datetime
import os.path
from itertools import islice

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from apscheduler.schedulers.blocking import BlockingScheduler
from bs4 import BeautifulSoup
import obsws_python as obs
import time
from playwright.sync_api import sync_playwright
import os
from config import obs_websocket_password

def shutdown_computer():
    print("Shutting down the computer...")
    # os.system("shutdown /s /t 10")
    scheduler.shutdown(wait=False)


def connect_to_google_meet(session_duration, event_link):
    os.system('start /d "C:/Program Files/Google/Chrome/Application" chrome.exe --remote-debugging-port=9222 --profile-directory="Profile 2" --start-fullscreen')
    time.sleep(3)
    with sync_playwright() as playwright:
        # Connect to an existing instance of Chrome using the connect_over_cdp method.
        # check debug mode in chrome http://localhost:9222/json/version
        browser = playwright.chromium.connect_over_cdp("http://localhost:9222")
        default_context = browser.contexts[0]
        page = default_context.pages[0]
        page.goto(event_link)
        page.locator('button:has-text("Надіслати запит на приєднання"), button:has-text("Приєднатися зараз")').click()
        os.system('start /d "C:/Program Files/obs-studio/bin/64bit" obs64.exe --minimize-to-tray')
        time.sleep(3)
        cl = obs.ReqClient(host='localhost', port=4455, password=obs_websocket_password, timeout=3)
        cl.start_record()
        time.sleep(session_duration)
        cl.stop_record()
        page.close()
        os.system("taskkill /f /im obs64.exe")


def connect_to_zoom(session_duration, event_link):
    os.system(
        'start /d "C:/Program Files/Google/Chrome/Application" chrome.exe --remote-debugging-port=9222 --profile-directory="Profile 2" --start-fullscreen')
    time.sleep(3)
    with sync_playwright() as playwright:
        # Connect to an existing instance of Chrome using the connect_over_cdp method.
        # check debug mode in chrome http://localhost:9222/json/version
        browser = playwright.chromium.connect_over_cdp("http://localhost:9222")
        default_context = browser.contexts[0]
        page = default_context.pages[0]

        page.goto(event_link, wait_until='networkidle')
        os.system('start /d "C:/Program Files/obs-studio/bin/64bit" obs64.exe --minimize-to-tray')
        time.sleep(3)
        # get obs_websocket_password from obs_credentials.json
        cl = obs.ReqClient(host='localhost', port=4455, password=obs_websocket_password, timeout=3)
        cl.start_record()
        time.sleep(session_duration)
        cl.stop_record()
        page.close()
        os.system("taskkill /f /im Zoom.exe")
        os.system("taskkill /f /im obs64.exe")


def extract_link(text):
    soup = BeautifulSoup(text, 'html.parser')
    links = []
    for link in soup.find_all('a'):
        links.append(link.get_text())
    return links[0]



def main():
    SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
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

    try:
        service = build('calendar', 'v3', credentials=creds)

        # Call the Calendar API
        now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        end_of_day = datetime.datetime.utcnow().replace(hour=23, minute=59, second=59).isoformat() + 'Z'
        events_result = service.events().list(calendarId='primary', timeMin=now, timeMax=end_of_day,
                                              maxResults=10, singleEvents=True,
                                              orderBy='startTime').execute()
        events = events_result.get('items', [])

        if not events:
            print('No upcoming events found.')
            return
        else:
            for index, event in enumerate(events):
                start_date_text = event['start'].get('dateTime', event['start'].get('date'))
                end_date_text = event['end'].get('dateTime', event['end'].get('date'))
                event_link = extract_link(event['description'])
                print("added job: ", start_date_text, end_date_text, event['summary'], event_link)
                start = datetime.datetime.strptime(start_date_text, '%Y-%m-%dT%H:%M:%S%z')
                end_of_event = datetime.datetime.strptime(end_date_text, '%Y-%m-%dT%H:%M:%S%z')
                session_duration = (end_of_event - start).total_seconds()
                if "meet.google.com" in event_link:
                    scheduler.add_job(connect_to_google_meet, 'date', next_run_time=start,
                                      args=[session_duration, event_link])
                else:
                    scheduler.add_job(connect_to_zoom, 'date', next_run_time=start,
                                      args=[session_duration, event_link])
                if index == len(events) - 1:
                    time_for_shutdown_computer = end_of_event + datetime.timedelta(minutes=1)
                    scheduler.add_job(shutdown_computer, 'date', next_run_time=time_for_shutdown_computer)
    except HttpError as error:
        print('An error occurred: %s' % error)


if __name__ == '__main__':
    scheduler = BlockingScheduler()
    main()
    scheduler.start()
