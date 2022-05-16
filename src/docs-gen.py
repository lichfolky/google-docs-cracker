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

SAMPLE_RANGE_NAME = "ELENCO PER IMPEGNATIVE!A:D"
FOLDER_NAME = "Impegnative progettazione 2021-22"


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

        now = datetime.now()
        dt_string = now.strftime("%d/%m/%Y_%H:%M:%S")

        folder_id = create_folder(drive_service, FOLDER_NAME)
        data_rows = load_sheet_values(
            sheet_service, cfg["SPREADSHEET_ID"], SAMPLE_RANGE_NAME
        )

        name = "#noname"
        previous_name = "#noname"
        project_rows = []
        xfiles = 0

        # Create a doc for each row of a sheet in a new folder,
        # using a document as template and substituting keywords
        # with the row contents
        for count, row in enumerate(data_rows):
            if count >= 1:
                if row:
                    name = row[0]
                if count == 1:
                    previous_name = name
                if previous_name != name and count >= 2:
                    create_documents(
                        drive_service,
                        docs_service,
                        folder_id,
                        previous_name,
                        project_rows,
                    )
                    previous_name = name
                    project_rows = [row]
                    xfiles += 1
                    print(xfiles, "files created")
                else:
                    project_rows.append(row)
                    if count == len(data_rows) - 1:
                        create_documents(
                            drive_service,
                            docs_service,
                            folder_id,
                            previous_name,
                            project_rows,
                        )
                        xfiles += 1
    except HttpError as err:
        print(err)


def create_documents(drive_service, docs_service, folder_id, name, values):
    doc_id_1 = create_doc_from_sample(
        drive_service,
        folder_id,
        cfg["SAMPLE_DOCUMENT_ID_1"],
        name + "_mod2",
    )

    replace_text_1(docs_service, doc_id_1, values)

    doc_id_2 = create_doc_from_sample(
        drive_service,
        folder_id,
        cfg["SAMPLE_DOCUMENT_ID_2"],
        name + "_mod4",
    )

    replace_text_2(docs_service, doc_id_2, values)


# Call the Sheets API and load range values
def load_sheet_values(sheet_service, spreadsheet_id, range):
    sheet = sheet_service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range).execute()
    values = result.get("values", [])

    if not values:
        print("No data found.")
        return

    return values


# Using Drive API, copy a file with the selected name and in the folder
# identified by the id
def create_doc_from_sample(drive_service, folder_id, sample_id, name):
    body = {"name": name, "parents": [folder_id]}
    drive_response = drive_service.files().copy(fileId=sample_id, body=body).execute()
    return drive_response.get("id")


# Create a new root folder
def create_folder(drive_service, name):
    file_metadata = {"name": name, "mimeType": "application/vnd.google-apps.folder"}
    drive_response = (
        drive_service.files().create(body=file_metadata, fields="id").execute()
    )
    return drive_response.get("id")


# Replace text in a document
def replace_text_1(docs_service, document_id, project_rows):

    # ES:
    # - Nuovi intrecci urbani
    # - Sport per includere: giocare per l'integrazione delle culture

    text = ""
    for row in project_rows:
        text += "- " + row[1] + "\n"

    requests = [
        {
            "replaceAllText": {
                "containsText": {"text": "{{content}}", "matchCase": "true"},
                "replaceText": text,
            }
        }
    ]

    docs_response = (
        docs_service.documents()
        .batchUpdate(documentId=document_id, body={"requests": requests})
        .execute()
    )


def replace_text_2(docs_service, document_id, project_rows):
    # ES:
    # Sport per includere: giocare per l'integrazione delle culture, 2 posti,
    # di cui 1 posto riservato a giovani con minori opportunità economiche (ISEE ≤ 15.000 €)

    text = ""
    for row in project_rows:
        if len(row) > 2:
            if row[2] == "1":
                text += "- " + row[1] + ", " + row[2] + " posti"
            else:
                text += "- " + row[1] + ", " + row[2] + " posto"
            if len(row) > 3:
                if row[3] == "1":
                    text += (
                        ", di cui "
                        + row[3]
                        + " posto riservato a giovani con minori opportunità economiche"
                    )
                else:
                    text += (
                        ", di cui "
                        + row[3]
                        + " posti riservati a giovani con minori opportunità economiche"
                    )
            text += "\n"

    requests = [
        {
            "replaceAllText": {
                "containsText": {"text": "{{content}}", "matchCase": "true"},
                "replaceText": text,
            }
        }
    ]

    docs_response = (
        docs_service.documents()
        .batchUpdate(documentId=document_id, body={"requests": requests})
        .execute()
    )


if __name__ == "__main__":
    main()
