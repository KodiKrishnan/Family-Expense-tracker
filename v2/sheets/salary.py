from utils.sheets_utils import get_sheet_id

def create_individual_salary_sheet(service, spreadsheet_id):
    """
    Creates Individual_Salary sheet to track monthly salary per person.
    Safe to run once; will error if sheet already exists (expected behavior).
    """

    # 1️⃣ Create sheet
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "requests": [{
                "addSheet": {
                    "properties": {
                        "title": "Individual_Salary"
                    }
                }
            }]
        }
    ).execute()

    # 2️⃣ Header + starter data
    values = [
        ["Month", "Person", "Salary", "Actual Spent", "Remaining"],
        ["Jan-2026", "Karthi", 40000],
        ["Jan-2026", "Chandru", 50000],
        ["Jan-2026", "Deiva", 20000],
    ]

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Individual_Salary!A1",
        valueInputOption="USER_ENTERED",
        body={"values": values}
    ).execute()

    # 3️⃣ Freeze header row (UX polish)
    sheet_id = get_sheet_id(service, spreadsheet_id, "Individual_Salary")
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "requests": [{
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sheet_id,
                        "gridProperties": {
                            "frozenRowCount": 1
                        }
                    },
                    "fields": "gridProperties.frozenRowCount"
                }
            }]
        }
    ).execute()
    
def compute_individual_salary_actuals(service, spreadsheet_id):
    """
    Writes Actual Spent (D) and Remaining (E) formulas
    into Individual_Salary sheet.
    Uses robust MAP/LAMBDA/SUMIFS for accurate calculation.
    """

    # Actual Spent (D2)
    actual_formula = (
        "=MAP("
        "A2:A, B2:B,"
        "LAMBDA(month, person,"
        "IF(person=\"\",\"\","
        "SUMIFS("
        "Expenses!G:G,"
        "Expenses!J:J, person,"
        "Expenses!B:B, month"
        ")"
        ")"
        ")"
        ")"
    )

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Individual_Salary!D2",
        valueInputOption="USER_ENTERED",
        body={"values": [[actual_formula]]}
    ).execute()

    # Remaining (E2)
    remaining_formula = (
        '=ARRAYFORMULA('
        'IF(B2:B="","",'
        'C2:C - D2:D'
        ')'
        ')'
    )

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Individual_Salary!E2",
        valueInputOption="USER_ENTERED",
        body={"values": [[remaining_formula]]}
).execute()


def highlight_negative_remaining_individual_salary(service, spreadsheet_id):
    sheet_id = get_sheet_id(service, spreadsheet_id, "Individual_Salary")

    request = {
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [{
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "startColumnIndex": 4,
                    "endColumnIndex": 5
                }],
                "booleanRule": {
                    "condition": {
                        "type": "NUMBER_LESS",
                        "values": [{"userEnteredValue": "0"}]
                    },
                    "format": {
                        "backgroundColor": {
                            "red": 1.0,
                            "green": 0.85,
                            "blue": 0.85
                        },
                        "textFormat": {
                            "bold": True
                        }
                    }
                }
            },
            "index": 0
        }
    }

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [request]}
    ).execute()

def add_individual_salary_dashboard_table(service, spreadsheet_id):

    formula = (
        '=QUERY('
        'Individual_Salary!A1:E,'
        '"select A, B, C, D, E '
        'where B is not null '
        'label '
        'A \'Month\', '
        'B \'Person\', '
        'C \'Salary\', '
        'D \'Actual Spent\', '
        'E \'Remaining\'"'
        ')'
    )

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!S2",
        valueInputOption="USER_ENTERED",
        body={"values": [[formula]]}
    ).execute()
