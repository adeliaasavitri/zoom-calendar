from __future__ import print_function

import datetime
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
print("Input date (Max 7 days)")
while True:
    try:
        date_from_str = input("Enter the start date (YYYY-MM-DD): ")
        date_from = datetime.datetime.strptime(date_from_str, "%Y-%m-%d")
        date_to_str = input("Enter the end date (YYYY-MM-DD): ")
        date_to = datetime.datetime.strptime(date_to_str, "%Y-%m-%d") + datetime.timedelta(days=1)

        if (date_to - date_from).days > 7:
            raise ValueError("The date range cannot be more than 7 days.")
        
        break   # Exit the loop if there is no error

    except ValueError as e:
        print("Invalid input:", e)
        print("Please enter the dates in the format YYYY-MM-DD and make sure the date range is not more than 7 days.\n")

row_values = {'Preface Coding': 2, 'Preface Coding 1':5, 'Preface Coding 2':8, \
              'Preface Coding 3':11, 'Preface Coding 4':14, 'Preface Coding 5':17, \
                'Misc/ Google Meet': 20}
unique_date = []
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
def grab():
    try:
        service = build('calendar', 'v3', credentials=creds)

        # Call the Calendar API
        events_result = service.events().list(calendarId='c_u84l39q8omdas37t7csorl7mk8@group.calendar.google.com', timeMin=date_from.isoformat() + 'Z', timeMax=date_to.isoformat() + 'Z',
                                              maxResults=70, singleEvents=True,
                                              orderBy='startTime').execute()
        events = events_result.get('items', [])

        if not events:
            return "not found"

        # Prints the start and name of the next 10 events
        # Prints the start time, summary, and colorId of the next 10 events
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

# Create a Spreadsheet
def create(title):
    #creds, _ = google.auth.default()
    try:
        service = build('sheets', 'v4', credentials=creds)
        spreadsheet = {
            'properties': {
                'title': title
            }
        }
        spreadsheet = service.spreadsheets().create(body=spreadsheet,
                                                    fields='spreadsheetId') \
            .execute()
        print(f"Spreadsheet ID: {(spreadsheet.get('spreadsheetId'))}")
        return spreadsheet.get('spreadsheetId')
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error

def clear_sheet():
    try:
        # Define the range of cells to preserve
        preserve_range = 'Sheet1!A2:A22'
            
        service = build('sheets', 'v4', credentials=creds)

        # The A1 notation of the values to clear.
        range_ = 'B1:F22'  

        clear_values_request_body = {
            'range': 'Sheet1!B1:F22'
        }

        request = service.spreadsheets().values().clear(spreadsheetId=SPREADSHEET_ID, range=range_, body=clear_values_request_body)
        response = request.execute()

    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


# # Clear cell values
# def clear_header():
#     try:
#         service = build('sheets', 'v4', credentials=creds)
#         values = [
#            ['', '', '', '', '', '', ''],
#         ]
#         data = [
#             {
#                 'range': "B1:H1",
#                 'values': values
#             }
#         ]
#         body = {
#             'valueInputOption': "RAW",
#             'data': data
#         }
#         service.spreadsheets().values().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
#     except HttpError as error:
#         print(f"An error occurred: {error}")
#         return error

# Update date and day as header
def update_header():
    try:
        service = build('sheets', 'v4', credentials=creds)

        for date in results_date:
            if date not in unique_date:
                unique_date.append(date)
        values = [
           unique_date,
        ]
        data = [
            {
                'range': "B1:H1",
                'values': values
            }
        ]
        body = {
            'valueInputOption': "RAW",
            'data': data
        }
        service.spreadsheets().values().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()

    except HttpError as error:
        print(f"An error occurred: {error}")
        return error
    
# Update based on zoom account
# def clear_content():
#     try:
#         service = build('sheets', 'v4', credentials=creds)
#         for i in range(len(results_date)):
#             values = [
#                 [''],
#             ]
#             letter = chr(ord('A') + unique_date.index(results_date[i]) + 1)
#             number = row_values[results_account[i]]
#             body = {
#                 'values':values
#             }
#             # Retrieve the value of the cell
#             result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f"{letter}{number}").execute()
#             # Check whether the cell is filled
#             if 'values' in result and len(result['values']) > 0 and len(result['values'][0]) > 0:
#                 number += 1
#             service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=f"{letter}{str(number)}", valueInputOption='RAW', body=body).execute()
#     except HttpError as error:
#         print(f"An error occurred: {error}")
#         return error

# Update based on zoom account
def update_content():
    try:
        service = build('sheets', 'v4', credentials=creds)
    
        for i in range(len(results_date)):

            values = [
                [str(results_summary[i])],
            ]

            letter = chr(ord('A') + unique_date.index(results_date[i]) + 1)
            number = row_values[results_account[i]]

            body = {
                'values':values
            }

            # Retrieve the value of the cell
            result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f"{letter}{number}").execute()

            # Check whether the cell is filled
            if 'values' in result and len(result['values']) > 0 and len(result['values'][0]) > 0:
                result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=f"{letter}{number+1}").execute()
                if 'values' in result and len(result['values']) > 0 and len(result['values'][0]) > 0:
                    number += 2
                else:
                    number += 1

            service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=f"{letter}{str(number)}", valueInputOption='RAW', body=body).execute()
        print("Access the spreadsheets here: https://docs.google.com/spreadsheets/d/1FKBtgyKbCWJDLANAUqA8mGTytD4rDJe33UTeVLcJahY/edit#gid=0")
    except HttpError as error:
        print(f"An error occurred: {error}")
        return error

if __name__ == '__main__':
    if grab() == 'not found':
        print("No event found")
    # Run this to make a spreadsheet 
    # create("Test Automation")
    else:
        clear_sheet()
        # clear_header()
        update_header()
        # clear_content()
        update_content()



