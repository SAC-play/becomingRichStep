import os.path 

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


def batch_update_cells(spreadsheet_id, updates):
    """Updates multiple non-contiguous ranges of cells in a Google Sheet."""
    creds = None
    if os.path.exists('token1.json'):
        creds = Credentials.from_authorized_user_file('token1.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credential\\credential.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open('token1.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        body = {'value_input_option': 'USER_ENTERED', 'data': updates}

        result = (
            service.spreadsheets()
            .values()
            .batchUpdate(spreadsheetId=spreadsheet_id, body=body)
            .execute()
        )
        print(f"{result.get('totalUpdatedCells')} cells updated in batch.")
        return result
    except HttpError as err:
        print(f'An error occurred: {err}')
        return None

import hana_card_amount
import shinhan_card_amount
import samsung_card_amount

if __name__ == '__main__':
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    SPREADSHEET_ID = '18IXJIwoPPkcUt8zGXie4u5J3U8LkLlMOD4L8N8CJhoE'  # Replace with your ID

    hana_amount = hana_card_amount.get_hana_amount()
    shinhan_amount = shinhan_card_amount.get_shinhan_amount()
    samsung_amount = samsung_card_amount.get_samsung_amount()

    sheet_name = "250323~250423"
    ranges = []
    ranges.append(sheet_name + '!' + "D7")
    ranges.append(sheet_name + '!' + "D13")
    values = []
    values.append(samsung_amount)
    values.append(hana_amount + shinhan_amount)

    # Define the updates as a list of dictionaries. Each dictionary specifies
    # a range and the corresponding values to write to that range.
    UPDATES = [
        {
            'range': ranges[0],
            'values': [
                [values[0]]
            ],
        },
        {
            'range': ranges[1],
            'values': [
                [values[1]],
            ],
        }
    ]

    batch_update_cells(SPREADSHEET_ID, UPDATES)