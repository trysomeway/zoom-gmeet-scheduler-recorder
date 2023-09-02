from __future__ import print_function
import datetime
import os.path
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
from config import obs_websocket_password, run_browser, close_zoom, localhost_browser_url, run_obs, shutdown_pc_command, \
    close_obs


def shutdown_schedulerr():
    scheduler.shutdown(wait=False)


def shutdown_computer():
    print("Shutting down the computer...")
    shutdown_schedulerr()
    os.system(shutdown_pc_command)


def record_by_obs(session_duration):
    os.system(run_obs)
    time.sleep(3)
    cl = obs.ReqClient(host='localhost', port=4455, password=obs_websocket_password, timeout=3)
    cl.start_record()
    time.sleep(session_duration)
    cl.stop_record()
    os.system(close_obs)

def connect_to_google_meet_or_zoom(session_duration, event_link):
    os.system(run_browser)
    time.sleep(3)
    with sync_playwright() as playwright:
        # Connect to an existing instance of Chrome using the connect_over_cdp method.
        browser = playwright.chromium.connect_over_cdp(localhost_browser_url)
        default_context = browser.contexts[0]
        page = default_context.pages[0]
        page.goto(event_link)
        if "meet.google.com" in event_link:
            page.locator('button:has-text("Ask to join"), button:has-text("Join now")').click()
            record_by_obs(session_duration)
            page.close()
        else:
            page.goto(event_link, wait_until='networkidle')
            record_by_obs(session_duration)
            os.system(close_zoom)
            page.close()




def extract_link(text):
    soup = BeautifulSoup(text, 'html.parser')
    links = []
    for link in soup.find_all('a'):
        links.append(link.get_text())
    return links[0]

def get_credentials():
    SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def get_events_from_g_calendar_for_today():
    creds = get_credentials()
    try:
        service = build('calendar', 'v3', credentials=creds)
        # Call the Calendar API
        now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        end_of_day = datetime.datetime.utcnow().replace(hour=23, minute=59, second=59).isoformat() + 'Z'
        events_result = service.events().list(calendarId='primary', timeMin=now, timeMax=end_of_day,
                                              maxResults=10, singleEvents=True,
                                              orderBy='startTime').execute()
        return events_result.get('items', [])

    except HttpError as error:
        print('An error occurred: %s' % error)


def schedule_events_record(events):
    if not events:
        print('No upcoming events found.')
        scheduler.add_job(shutdown_schedulerr, 'interval')
    else:
        for index, event in enumerate(events):
            start_date_text = event['start'].get('dateTime', event['start'].get('date'))
            end_date_text = event['end'].get('dateTime', event['end'].get('date'))
            event_link = extract_link(event['description'])
            print("added job: ", start_date_text, end_date_text, event['summary'], event_link)
            start = datetime.datetime.strptime(start_date_text, '%Y-%m-%dT%H:%M:%S%z')
            end_of_event = datetime.datetime.strptime(end_date_text, '%Y-%m-%dT%H:%M:%S%z')
            session_duration = (end_of_event - start).total_seconds()
            scheduler.add_job(connect_to_google_meet_or_zoom, 'date', next_run_time=start,
                              args=[session_duration, event_link])
            if index == len(events) - 1:
                time_for_shutdown_computer = end_of_event + datetime.timedelta(minutes=1)
                scheduler.add_job(shutdown_computer, 'date', next_run_time=time_for_shutdown_computer)


if __name__ == '__main__':
    scheduler = BlockingScheduler()
    events = get_events_from_g_calendar_for_today()
    schedule_events_record(events)
    scheduler.start()
