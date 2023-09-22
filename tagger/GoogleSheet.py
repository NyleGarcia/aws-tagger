from __future__ import print_function

import os.path
import string
from itertools import combinations_with_replacement as cwr
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


class GoogleSheet:
    def __init__(self, name, id=None, baserange=None):
        SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
        self.name = name
        self.sheet_name = baserange
        creds = None
        # The file token.json stores the user's access
        # and refresh tokens, and is
        # created automatically when the authorization
        # flow completes for the first time.
        if os.path.exists("Files/token.json"):
            creds = Credentials.from_authorized_user_file("Files/token.json", SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "Files/credentials.json", SCOPES
                )
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open("Files/token.json", "w") as token:
                token.write(creds.to_json())

        self.service = build("sheets", "v4", credentials=creds)
        if id is None:
            self.make_wb()

        else:
            self.id = id

        if self.sheet_name is None:
            self.sheet_name = "sheet1"
        else:
            self.make_sheet()

    def make_wb(self):
        spreadsheet = {"properties": {"title": self.name}}
        spreadsheet = (
            self.service.spreadsheets()
            .create(body=spreadsheet, fields="spreadsheetId")
            .execute()
        )
        self.id = spreadsheet.get("spreadsheetId")

    def make_sheet(self):
        data = {"requests": [{"addSheet": {"properties": {"title": self.sheet_name}}}]}
        try:
            spreadsheet = (
                self.service.spreadsheets()
                .batchUpdate(spreadsheetId=self.id, body=data)
                .execute()
            )
        except Exception as e:
            pass

    def calc_range(self, headers):
        alphabet = string.ascii_lowercase
        length = 2
        sizer = ["".join(comb) for comb in cwr(alphabet, length)]

        self.range = f"{self.sheet_name}!A:{sizer[len(headers) + 1]}"

    def sheet_write(self, headers, values):
        self.headers = headers
        data = [headers]
        for value in values:
            data.append(value)
        sheet = self.service.spreadsheets()
        self.calc_range(headers)
        body = {"values": data}

        result = (
            self.service.spreadsheets()
            .values()
            .update(
                spreadsheetId=self.id,
                range=self.range,
                valueInputOption="USER_ENTERED",
                body=body,
            )
            .execute()
        )
        print("{0} cells updated.".format(result.get("updatedCells")))
        result = sheet.values().get(spreadsheetId=self.id, range=self.range).execute()


    def clear_sheet(self, headers, values):
        blankrow = []
        blankset = []

        for x in range(len(headers) + 2):
            blankrow.append("")
        for x in range(len(values * 2)):
            blankset.append(blankrow)
        self.sheet_write(headers, blankset)

    def sheet_read(self):
        sheet = self.service.spreadsheets()
        result = sheet.values().get(spreadsheetId=self.id, range=self.range).execute()
        values = result.get("values", [])
        return values

    def get_id(self):
        return self.id


