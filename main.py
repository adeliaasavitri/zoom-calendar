from __future__ import print_function
from datetime import datetime, timedelta
import datetime
import calendar
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly', 'https://www.googleapis.com/auth/spreadsheets']
COLORS = {'4': 'Preface Coding', '11': 'Preface Coding 1', '5': 'Preface Coding 2', '10': 'Preface Coding 3', '7': 'Preface Coding 4', '3': 'Preface Coding 5', '6': 'Misc/ Google Meet', 'None': 'Misc/ Google Meet'}
SPREADSHEET_ID = '1FKBtgyKbCWJDLANAUqA8mGTytD4rDJe33UTeVLcJahY'

# Get starting date and end of month
date_today = datetime.datetime.now()
currentMonth = datetime.datetime.now().month
currentYear = datetime.datetime.now().year
obj = calendar.Calendar()
weeks = []
# Get how many weeks
for day in obj.monthdatescalendar(currentYear, currentMonth):
        weeks.append(len(day))
day_values = {"Monday": 'B', "Tuesday": 'C', "Wednesday": 'D', "Thursday": 'E', "Friday": 'F', "Saturday": 'G', "Sunday": 'H'}
results_date = []
results_summary = []
results_account = []

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
        creds = flow.run_local_server(port=8080)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

# Grab the information
def grab(start, end):
    try:
        service = build('calendar', 'v3', credentials=creds)


        # Call the Calendar API
        events_result = service.events().list(calendarId='c_u84l39q8omdas37t7csorl7mk8@group.calendar.google.com', timeMin=start.isoformat() + 'T00:00:00Z', timeMax=end.isoformat() + 'T00:00:00Z',
                                              maxResults=1000, singleEvents=True,
                                              orderBy='startTime').execute()
        events = events_result.get('items', [])

        if not events:
            return "not found"
        
        results_date.clear()
        results_summary.clear()
        results_account.clear()

        for event in events:
            start_time = event['start'].get('dateTime', event['start'].get('date'))
            start_time_obj = datetime.datetime.fromisoformat(start_time)
            start_time_pm = start_time_obj.strftime("%I:%M %p").replace("AM", "PM").replace("am", "pm")
            start_date = start_time_obj.strftime("%B %d, %Y")
            start_day = start_time_obj.strftime("%A")
            summary = event['summary']
            company_name = summary.split('Executive')
            company_name=company_name[0]
            color_id = event.get('colorId')

            
            results_date.append(f"{start_day} {start_date}")
            results_summary.append(f"{start_time_pm} {company_name[:-3]}") 
            results_account.append(COLORS[str(color_id)])

    except HttpError as error:
        print('An error occurred: %s' % error)

def clear_sheet():
    try:
        # Define the range of cells to preserve
        for i in range(5):
            preserve_range = f'Week{i + 1}!A2:A22'
                
            service = build('sheets', 'v4', credentials=creds)

            # The A1 notation of the values to clear.
            range_ = f'Week{i + 1}!B1:H29'  

            clear_values_request_body = {
                'range': f'Week{i + 1}!B1:H29'
            }

            request = service.spreadsheets().values().clear(spreadsheetId=SPREADSHEET_ID, range=range_, body=clear_values_request_body)
            response = request.execute()

    except HttpError as error:
        print(f"An error occurred: {error}")
        return error

# Update date and day as header
def update_header():
    try:
        service = build('sheets', 'v4', credentials=creds)
        update_main =[]

        for i in range(len(weeks)):
            for day in obj.monthdatescalendar(currentYear, currentMonth)[i]:
                start_date = day.strftime("%B %d, %Y")
                start_day = day.strftime("%A")
                update_main.append({
                'range': f"Week{i+1}!{day_values[start_day]}1",
                'majorDimension': 'ROWS',
                'values': [[f"{start_day} {start_date}"]]
                })

        request_body = {'value_input_option': 'RAW','data': update_main}
        service.spreadsheets().values().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=request_body).execute()  

    except HttpError as error:
        print(f"An error occurred: {error}")
        return error
    
# Update based on zoom account
def update_content():
    try:
        service = build('sheets', 'v4', credentials=creds)
        update_body = []
        cell_per_week = []

        for i in range(5):
            row_values = {'Preface Coding': 2, 'Preface Coding 1':6, 'Preface Coding 2':10, \
            'Preface Coding 3':14, 'Preface Coding 4':18, 'Preface Coding 5':22, \
            'Misc/ Google Meet': 26}
            grab(obj.monthdatescalendar(currentYear, currentMonth)[i][0], obj.monthdatescalendar(currentYear, currentMonth)[i][0]+timedelta(days=6))
             
            for y in range(len(results_summary)):
                sheet = i+1
                day = results_date[y].split(' ')[0]
                letter = day_values[day]
                number = row_values[results_account[y]]
                for j in range(4):
                    if f"{letter}{number}" in cell_per_week:
                        number += 1
                    else:
                        break
                cell_per_week.append(f"{letter}{number}")
                update_body.append({
                        'range': f"Week{sheet}!{letter}{number}",
                        'majorDimension': 'ROWS',
                        'values': [[f"{results_summary[y]}"]]
                })

            cell_per_week.clear()
        request_body = {'value_input_option': 'RAW','data': update_body}
        service.spreadsheets().values().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=request_body).execute()

    except HttpError as error:
        print(f"An error occurred: {error}")
        return error

if __name__ == '__main__':
    clear_sheet()
    update_header()
    update_content()



