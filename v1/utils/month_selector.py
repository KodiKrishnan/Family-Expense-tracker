from utils.sheets_utils import get_sheet_id

def create_meta_sheet(service, spreadsheet_id):
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "requests": [{
                "addSheet": {
                    "properties": {
                        "title": "_Meta",
                        "hidden": True
                    }
                }
            }]
        }
    ).execute()

def populate_month_list(service, spreadsheet_id):
    formula = (
        '=SORT(UNIQUE(Expenses!B2:B), 1, FALSE)'
    )

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="_Meta!A2",
        valueInputOption="USER_ENTERED",
        body={"values": [[formula]]}
    ).execute()


def add_month_selector(service, spreadsheet_id):
    dashboard_id = get_sheet_id(service, spreadsheet_id, "Dashboard")

    # Label
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!O1",
        valueInputOption="USER_ENTERED",
        body={"values": [["Month"]] }
    ).execute()

    # Default = current month
    default_formula = (
        '=TEXT(TODAY(), "mmm-yyyy")'
    )

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!O2",
        valueInputOption="USER_ENTERED",
        body={"values": [[default_formula]]}
    ).execute()

    # Data validation
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "requests": [{
                "setDataValidation": {
                    "range": {
                        "sheetId": dashboard_id,
                        "startRowIndex": 1,
                        "endRowIndex": 2,
                        "startColumnIndex": 14,
                        "endColumnIndex": 15
                    },
                    "rule": {
                        "condition": {
                            "type": "ONE_OF_RANGE",
                            "values": [{
                                "userEnteredValue": "_Meta!A2:A"
                            }]
                        },
                        "showCustomUi": True
                    }
                }
            }]
        }
    ).execute()
