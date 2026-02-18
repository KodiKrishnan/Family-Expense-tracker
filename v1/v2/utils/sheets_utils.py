from googleapiclient.http import MediaFileUpload

def upload_sheet(drive):
    """
    Uploads temp.xlsx as a Google Sheet and returns spreadsheet_id
    """
    metadata = {
        "name": "Family Expense Tracker",
        "mimeType": "application/vnd.google-apps.spreadsheet"
    }

    media = MediaFileUpload(
        "temp.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    file = (
        drive.files()
        .create(body=metadata, media_body=media, fields="id")
        .execute()
    )

    return file["id"]


def get_sheet_id(service, spreadsheet_id, title):
    spreadsheet = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id
    ).execute()

    for s in spreadsheet["sheets"]:
        if s["properties"]["title"] == title:
            return s["properties"]["sheetId"]

    raise ValueError("Sheet not found")
