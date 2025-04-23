import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


def write_to_sheet(spreadsheet_id, range_name, value):
    """Writes a value to a specified cell in a Google Sheet."""
    creds = None
    # The file token1.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token1.json'):
        creds = Credentials.from_authorized_user_file('token1.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credential\\credential.json", SCOPES
            )  # Replace 'path/to/your/credentials.json' with the actual path
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token1.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        values = [[value]]  # The value to write, in a list of lists format
        body = {'values': values}

        result = (
            service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption='USER_ENTERED',  # or 'RAW'
                body=body,
            )
            .execute()
        )
        print(f"{result.get('updatedCells')} cells updated.")
        return result
    except HttpError as err:
        print(f'An error occurred: {err}')
        return None


if __name__ == '__main__':
    # If modifying these scopes, delete the file token1.json.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']  # For read/write access

    SPREADSHEET_ID = '18IXJIwoPPkcUt8zGXie4u5J3U8LkLlMOD4L8N8CJhoE'  # Replace with your actual spreadsheet ID
    RANGE_NAME = '250323~250423!C24'  # Replace with the cell you want to write to (e.g., 'Sheet1!A1')
    NUMBER_TO_WRITE = 123  # Replace with the number you want to write

    write_to_sheet(SPREADSHEET_ID, RANGE_NAME, NUMBER_TO_WRITE)