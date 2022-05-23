from __future__ import print_function

import os.path
import json


from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime


dirname = os.path.dirname(__file__)

# If modifying these scopes, delete the file token.json.
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive",
]

CREDENTIAL_PATH = os.path.join(dirname, "../credentials/gcredentials.json")
CONFIG_PATH = os.path.join(dirname, "../credentials/config.json")
TOKEN_PATH = os.path.join(dirname, "../credentials/token.json")

with open(CONFIG_PATH) as json_data_file:
    cfg = json.load(json_data_file)

SAMPLE_RANGE_NAME = "avvii 25.05.22!A:P"
FOLDER_NAME = "Schede avvio"
PARENT_FOLDER_ID = "1RSM9Fgq2u12PHclaSqVyHtRJU57i3VpL"


def main():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIAL_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(TOKEN_PATH, "w") as token:
            token.write(creds.to_json())
    try:
        sheet_service = build("sheets", "v4", credentials=creds)
        drive_service = build("drive", "v3", credentials=creds)
        docs_service = build("docs", "v1", credentials=creds)

        main_folder_id = create_folder(drive_service, PARENT_FOLDER_ID, FOLDER_NAME)

        data_rows = load_sheet_values(
            sheet_service, cfg["SPREADSHEET_ID"], SAMPLE_RANGE_NAME
        )

        for index, row in enumerate(data_rows):
            if index > 0:
                if index == 1:
                    progetto = row[4]
                    sede = row[5]
                    progetto_folder_id = create_folder(
                        drive_service, main_folder_id, progetto
                    )
                    sede_folder_id = create_folder(
                        drive_service, progetto_folder_id, sede
                    )
                if progetto != row[4]:
                    progetto = row[4]
                    progetto_folder_id = create_folder(
                        drive_service, main_folder_id, progetto
                    )
                if sede != row[5]:
                    sede = row[5]
                    sede_folder_id = create_folder(
                        drive_service, progetto_folder_id, sede
                    )
                print(progetto + " " + sede + ": " + row[1] + " " + row[2])
                doc_id = create_doc_from_template(
                    drive_service,
                    sede_folder_id,
                    cfg["TEMPLATE_DOCUMENT_ID"],
                    row[1] + "_" + row[2],
                )
                replace_text(
                    docs_service,
                    doc_id,
                    [
                        "nome",
                        "cognome",
                        "codice",
                        "titolo progetto",
                        "sede nome",
                        "sede indirizzo",
                        "sede comune",
                        "olp nome cognome",
                        "olp telefono",
                        "e-mail",
                        "elenco",
                    ],
                    [
                        row[1],
                        row[2],
                        row[3],
                        row[4],
                        row[5],
                        row[6],
                        row[7],
                        row[8],
                        row[9],
                        row[10],
                        row[11],
                    ],
                )

    except HttpError as err:
        print(err)


# Create a new root folder
def create_folder(drive_service, parent_folder, name):
    file_metadata = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_folder],
    }

    drive_response = (
        drive_service.files().create(body=file_metadata, fields="id").execute()
    )
    return drive_response.get("id")


# Call the Sheets API and load range values
def load_sheet_values(sheet_service, spreadsheet_id, range):
    sheet = sheet_service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range).execute()
    values = result.get("values", [])

    if not values:
        print("No data found.")
        return

    return values


# Create a new doc from a template
def create_doc_from_template(drive_service, folder_id, template_id, name):
    body = {"name": name, "parents": [folder_id]}
    drive_response = drive_service.files().copy(fileId=template_id, body=body).execute()
    return drive_response.get("id")


# Replace text in a document
def replace_text(docs_service, document_id, keys, contents):
    requests = []
    for count, key in enumerate(keys):
        requests.append(
            {
                "replaceAllText": {
                    "containsText": {
                        "text": "{{" + key + "}}",
                        "matchCase": "true",
                    },
                    "replaceText": contents[count],
                }
            }
        )

    docs_response = (
        docs_service.documents()
        .batchUpdate(documentId=document_id, body={"requests": requests})
        .execute()
    )


if __name__ == "__main__":
    main()
