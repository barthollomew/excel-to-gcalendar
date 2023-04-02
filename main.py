import re
import datetime
import pytz
import os
import openpyxl
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# Constants
TIMEZONE = 'YOUR_TIMEZONE_HERE'
CREDENTIALS_FILE = 'YOUR_CREDENTIALS_FILE_HERE'
SCOPES = ['https://www.googleapis.com/auth/calendar']


# Authenticate with Google API
creds = service_account.Credentials.from_service_account_file(
    CREDENTIALS_FILE, scopes=SCOPES)


# Create a service instance
service = build('calendar', 'v3', credentials=creds)


# Read data from Excel file
workbook = openpyxl.load_workbook('YOUR_EXCEL_FILE_HERE')
worksheet = workbook.active


# Get input from user
name = input('Enter your name: ')


# Search for the row where the user's name is located
for row in worksheet.iter_rows(values_only=True):
    if row[0] == name:
        cell_value = row[1]
        break
else:
    print('Error: Could not find data for user in the spreadsheet')
    exit()


# Extract event name, start time, and end time from cell value using regex
match = re.search(r'^(.*)\s(\d{1,2}:\d{2}[ap]m)-(\d{1,2}:\d{2}[ap]m)$', cell_value)
if not match:
    print('Error: Could not extract event details from cell value')
    exit()

event_name = match.group(1)
start_time = datetime.datetime.strptime(match.group(2), '%I:%M%p')
end_time = datetime.datetime.strptime(match.group(3), '%I:%M%p')


# Create a datetime object for the current day and the specified time
now = datetime.datetime.now()
start_datetime = datetime.datetime(now.year, now.month, now.day, start_time.hour, start_time.minute)
end_datetime = datetime.datetime(now.year, now.month, now.day, end_time.hour, end_time.minute)


# Create a datetime object in the specified timezone
local_tz = pytz.timezone(TIMEZONE)
localized_start_datetime = local_tz.localize(start_datetime)
localized_end_datetime = local_tz.localize(end_datetime)


# Create event
event = {
    'summary': event_name,
    'location': 'Event Location',
    'description': 'Event Description',
    'start': {
        'dateTime': localized_start_datetime.isoformat(),
        'timeZone': TIMEZONE,
    },
    'end': {
        'dateTime': localized_end_datetime.isoformat(),
        'timeZone': TIMEZONE,
    },
}


# Insert event into calendar
try:
    event = service.events().insert(calendarId='primary', body=event).execute()
    print(f'Event created: {event.get("htmlLink")}')
except HttpError as error:
    print(f'An error occurred: {error}')
