
"""
Family Expense Tracker – QUERY-based, Production Ready
Author: Kodi Arasan M

This script:
- OAuth2 authentication (no service account)
- Uploads Excel to Google Sheets
- Dashboard + Monthly summaries via QUERY
- Prints Google Sheet link
"""

import os
import pickle
import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.http import MediaFileUpload

CLIENT_SECRET_FILE = "client_secret.json"
TOKEN_PICKLE = "token.pickle"

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]

GOOGLE_SHEET_NAME = "Family Expense Tracker"
DRIVE_FOLDER_ID = "1gB27vvJbdolhvkAp8h-LPRx8e5C0bO8i"

def get_credentials():
    creds = None
    if os.path.exists(TOKEN_PICKLE):
        with open(TOKEN_PICKLE, "rb") as f:
            creds = pickle.load(f)
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
        with open(TOKEN_PICKLE, "wb") as f:
            pickle.dump(creds, f)
    return creds

def create_test_data():
    expenses = pd.DataFrame([
    ["2026-01-04","Jan-2026",2026,"Loans","EMI","Apty Kalanchiam",5000,"Cash","Cash","Deiva","Mother","Loan","Monthly","Veni Anni Sangam","Yes","No","Family","Note3"],
    ["2026-01-05","Jan-2026",2026,"Loans","EMI","Kotak Due",4313,"Bank Transfer","Kotak811","Chandru","Anna","Loan","Monthly","Vendor1","Yes","No","Family","PAID"],
    ["2026-01-09","Jan-2026",2026,"Loans","EMI","Kalanchiam Kmpty",7000,"Cash","Cash","Chandru","Family","Loan","Monthly","Vendor2","Yes","No","Family","Note2"]
    ], columns=[
    "Date","Month","Year","Category","Sub-Category","Description","Amount",
    "Payment Mode","Account","Paid By","For Whom","Expense Type","Frequency",
    "Vendor","Bill?","Reimbursable","Tags","Notes"
    ])
    categories = pd.DataFrame({
        "Category":["Food","Food","Transport","Health","Loans"],
        "Sub-Category":["Groceries","Dining","Fuel","Medicines","Repayments"]
    })

    family = pd.DataFrame({
        "Member Name":["Chandru","Karthi","Appa","Amma","Pothu"],
        "Role":["Self","Brother","Father","Mother","Anni"]
    })

    payment = pd.DataFrame({
        "Payment Mode":["Cash","UPI","Card"],
        "Account":["Cash","GPay","Credit Card"]
    })

    budget = pd.DataFrame({
        "Month":["Jan-2026","Feb-2026"],
        "Category":["Loans","Health"],
        "Budget Amount":[12000,3000]
    })

    return expenses, categories, family, payment, budget

def export_excel(expenses, categories, family, payment, budget):
    with pd.ExcelWriter("temp.xlsx", engine="openpyxl") as writer:
        expenses.to_excel(writer, sheet_name="Expenses", index=False)
        categories.to_excel(writer, sheet_name="Categories", index=False)
        family.to_excel(writer, sheet_name="Family", index=False)
        payment.to_excel(writer, sheet_name="Payment_Modes", index=False)
        budget.to_excel(writer, sheet_name="Monthly_Budget", index=False)

def upload_sheet(drive):
    metadata = {
        "name": GOOGLE_SHEET_NAME,
        "parents": [DRIVE_FOLDER_ID],
        "mimeType": "application/vnd.google-apps.spreadsheet"
    }
    media = MediaFileUpload(
        "temp.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    file = drive.files().create(body=metadata, media_body=media, fields="id").execute()
    return file["id"]

def get_sheet_id(service, spreadsheet_id, title):
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for s in spreadsheet["sheets"]:
        if s["properties"]["title"] == title:
            return s["properties"]["sheetId"]
    raise ValueError("Sheet not found")

def apply_month_year_formula(service, spreadsheet_id):
    expenses_id = get_sheet_id(service, spreadsheet_id, "Expenses")

    requests = [
        # Ensure headers are correct
        {
            "updateCells": {
                "range": {
                    "sheetId": expenses_id,
                    "startRowIndex": 0,
                    "startColumnIndex": 1,
                    "endColumnIndex": 3
                },
                "rows": [{
                    "values": [
                        {"userEnteredValue": {"stringValue": "Month"}},
                        {"userEnteredValue": {"stringValue": "Year"}}
                    ]
                }],
                "fields": "userEnteredValue"
            }
        },

        # Month formula in B2 (NOT B1)
        {
            "updateCells": {
                "range": {
                    "sheetId": expenses_id,
                    "startRowIndex": 1,   # row 2
                    "startColumnIndex": 1,
                    "endColumnIndex": 2
                },
                "rows": [{
                    "values": [{
                        "userEnteredValue": {
                            "formulaValue":
                            '=ARRAYFORMULA(IF(A2:A="","",TEXT(A2:A,"mmm-yyyy")))'
                        }
                    }]
                }],
                "fields": "userEnteredValue"
            }
        },

        # Year formula in C2 (NOT C1)
        {
            "updateCells": {
                "range": {
                    "sheetId": expenses_id,
                    "startRowIndex": 1,   # row 2
                    "startColumnIndex": 2,
                    "endColumnIndex": 3
                },
                "rows": [{
                    "values": [{
                        "userEnteredValue": {
                            "formulaValue":
                            '=ARRAYFORMULA(IF(A2:A="","",YEAR(A2:A)))'
                        }
                    }]
                }],
                "fields": "userEnteredValue"
            }
        }
    ]

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()

    #service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()

def create_dashboard(service, spreadsheet_id):
    dashboard_id = None

    # Create Dashboard sheet
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [{"addSheet": {"properties": {"title": "Dashboard"}}}]}
    ).execute()

    # ----- KPI: Total Expense -----
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!A1",
        valueInputOption="USER_ENTERED",
        body={"values": [["Total Expense"]]}
    ).execute()

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!A2",
        valueInputOption="USER_ENTERED",
        body={"values": [["=SUM(Expenses!G:G)"]]}
    ).execute()

    # ----- KPI: Current Month Total (robust) -----
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!B1",
        valueInputOption="USER_ENTERED",
        body={"values": [["Current Month Total"]]}
    ).execute()

    current_month_formula = (
        '=SUM('
        'ARRAYFORMULA('
        'IF('
        'TEXT(IF(Expenses!A2:A="",,DATEVALUE(Expenses!A2:A)),"mmm-yyyy") = '
        'TEXT(MAX(IF(Expenses!A2:A="",,DATEVALUE(Expenses!A2:A))),"mmm-yyyy"),'
        'Expenses!G2:G*1,'
        '0'
        ')'
        ')'
        ')'
    )

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!B2",
        valueInputOption="USER_ENTERED",
        body={"values":[[current_month_formula]]}
    ).execute()

    # ----- KPI: Highest Expense -----
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!C1",
        valueInputOption="USER_ENTERED",
        body={"values": [["Highest Expense"]]}
    ).execute()

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!C2",
        valueInputOption="USER_ENTERED",
        body={"values": [["=MAX(Expenses!G:G)"]]}
    ).execute()

    # ----- Category Summary -----
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!A5",
        valueInputOption="USER_ENTERED",
        body={"values": [[
            '=QUERY(Expenses!A:R,'
            '"select D,sum(G) where A is not null '
            'group by D order by sum(G) desc '
            'label D \'Category\', sum(G) \'Amount\'")'

        ]]}
    ).execute()

    # ----- Payment Mode Summary -----
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!D5",
        valueInputOption="USER_ENTERED",
        body={"values": [[
            '=QUERY(Expenses!A:R,'
            '"select H,sum(G) where A is not null '
            'group by H order by sum(G) desc '
            'label H \'Payment Mode\', sum(G) \'Amount\'")'

        ]]}
    ).execute()



def create_monthly_sheet(service, spreadsheet_id, month):
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests":[{"addSheet":{"properties":{"title":month}}}]}
    ).execute()

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{month}!A1",
        valueInputOption="USER_ENTERED",
        body={"values":[[
        f'=QUERY(Expenses!A:R,"select D,sum(G) where B=\'{month}\' group by D label sum(G) \'Total Amount\'")'
        ]]}

    ).execute()

def highlight_highest_expense(service, spreadsheet_id):
    expenses_id = get_sheet_id(service, spreadsheet_id, "Expenses")

    request = {
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [{
                    "sheetId": expenses_id,
                    "startRowIndex": 1,
                    "endColumnIndex": 18
                }],
                "booleanRule": {
                    "condition": {
                        "type": "CUSTOM_FORMULA",
                        "values": [{
                            "userEnteredValue": '=G2=MAX($G$2:$G)'
                        }]
                    },
                    "format": {
                        "backgroundColor": {
                            "red": 1.0,
                            "green": 0.9,
                            "blue": 0.9
                        }
                    }
                }
            },
            "index": 0
        }
    }

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests":[request]}
    ).execute()


# ================= BUDGET =================
def add_budget_actual_helper(service, spreadsheet_id):
    formula = (
        '=QUERY({Expenses!D2:D, '
        'ARRAYFORMULA(IF(Expenses!B2:B = LOOKUP(2,1/(Expenses!B2:B<>""),Expenses!B2:B), '
        'Expenses!G2:G, 0))},'
        '"select Col1, sum(Col2) '
        'where Col2 > 0 '
        'group by Col1 '
        'label Col1 \'Category\', sum(Col2) \'Actual\'", 0)'
    )

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!J20",
        valueInputOption="USER_ENTERED",
        body={"values": [[formula]]}
    ).execute()


def add_budget_vs_actual(service, spreadsheet_id):
    """
    Adds a Budget vs Actual table in Dashboard!A20:E
    Automatically calculates Variance
    """
    formula = (
        '=ARRAYFORMULA(IF(LEN(Monthly_Budget!B2:B), '
        ' {Monthly_Budget!B2:B, Monthly_Budget!C2:C, '
        ' IFERROR(VLOOKUP(Monthly_Budget!B2:B, J20:I, 2, FALSE), 0), '
        ' Monthly_Budget!C2:C - IFERROR(VLOOKUP(Monthly_Budget!B2:B, J20:I, 2, FALSE), 0)}, '
        ' ""))'
    )

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!A20",
        valueInputOption="USER_ENTERED",
        body={"values": [[formula]]}
    ).execute()
    
def highlight_budget_overrun(service, spreadsheet_id):
    dashboard_id = get_sheet_id(service, spreadsheet_id, "Dashboard")

    rule = {
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [{
                    "sheetId": dashboard_id,
                    "startRowIndex": 11,
                    "startColumnIndex": 3,
                    "endColumnIndex": 4
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
                        }
                    }
                }
            },
            "index": 3
        }
    }

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests":[rule]}
    ).execute()

def add_dashboard_section_titles(service, spreadsheet_id):
    dashboard_id = get_sheet_id(service, spreadsheet_id, "Dashboard")

    requests = [

        # ===== Category Summary Title =====
        {
            "repeatCell": {
                "range": {
                    "sheetId": dashboard_id,
                    "startRowIndex": 3,   # A4
                    "endRowIndex": 4,
                    "startColumnIndex": 0,
                    "endColumnIndex": 2
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "bold": True,
                            "fontSize": 13
                        }
                    }
                },
                "fields": "userEnteredFormat.textFormat"
            }
        },

        # ===== Payment Mode Summary Title =====
        {
            "repeatCell": {
                "range": {
                    "sheetId": dashboard_id,
                    "startRowIndex": 3,   # D4
                    "endRowIndex": 4,
                    "startColumnIndex": 3,
                    "endColumnIndex": 5
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "bold": True,
                            "fontSize": 13
                        }
                    }
                },
                "fields": "userEnteredFormat.textFormat"
            }
        },

        # ===== Budget vs Actual Title =====
        {
            "repeatCell": {
                "range": {
                    "sheetId": dashboard_id,
                    "startRowIndex": 18,  # A19
                    "endRowIndex": 19,
                    "startColumnIndex": 0,
                    "endColumnIndex": 4
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "bold": True,
                            "fontSize": 13
                        }
                    }
                },
                "fields": "userEnteredFormat.textFormat"
            }
        }
    ]

    # Write the actual title text
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {
                    "range": "Dashboard!A4",
                    "values": [["Expense by Category"]]
                },
                {
                    "range": "Dashboard!D4",
                    "values": [["Expense by Payment Mode"]]
                },
                {
                    "range": "Dashboard!A19",
                    "values": [["Budget vs Actual (Current Month)"]]
                }
            ]
        }
    ).execute()

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()

def add_for_whom_summary(service, spreadsheet_id):
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!M20",
        valueInputOption="USER_ENTERED",
        body={"values":[[
            '=QUERY(Expenses!A:R,'
            '"select K, sum(G) where A is not null '
            'group by K order by sum(G) desc '
            'label K \'For Whom\', sum(G) \'Amount\'")'

        ]]}
    ).execute()

def add_dashboard_charts(service, spreadsheet_id):
    dashboard_id = get_sheet_id(service, spreadsheet_id, "Dashboard")

    requests = [

        # ================= CATEGORY PIE =================
        {
            "addChart": {
                "chart": {
                    "spec": {
                        "title": "Expenses by Category (%)",
                        "pieChart": {
                            "legendPosition": "RIGHT_LEGEND",
                            "threeDimensional": False,
                            "domain": {
                                "sourceRange": {
                                    "sources": [{
                                        "sheetId": dashboard_id,
                                        "startRowIndex": 4,
                                        "endRowIndex": 15,
                                        "startColumnIndex": 0,
                                        "endColumnIndex": 1
                                    }]
                                }
                            },
                            "series": {
                                "sourceRange": {
                                    "sources": [{
                                        "sheetId": dashboard_id,
                                        "startRowIndex": 4,
                                        "endRowIndex": 15,
                                        "startColumnIndex": 1,
                                        "endColumnIndex": 2
                                    }]
                                }
                            }
                        }
                    },
                    "position": {
                        "overlayPosition": {
                            "anchorCell": {
                                "sheetId": dashboard_id,
                                "rowIndex": 1,
                                "columnIndex": 7
                            }
                        }
                    }
                }
            }
        },

        # ================= PAYMENT MODE COLUMN =================
        {
            "addChart": {
                "chart": {
                    "spec": {
                        "title": "Expenses by Payment Mode",
                        "basicChart": {
                            "chartType": "COLUMN",
                            "legendPosition": "NO_LEGEND",
                            "domains": [{
                                "domain": {
                                    "sourceRange": {
                                        "sources": [{
                                            "sheetId": dashboard_id,
                                            "startRowIndex": 4,
                                            "endRowIndex": 15,
                                            "startColumnIndex": 3,
                                            "endColumnIndex": 4
                                        }]
                                    }
                                }
                            }],
                            "series": [{
                                "series": {
                                    "sourceRange": {
                                        "sources": [{
                                            "sheetId": dashboard_id,
                                            "startRowIndex": 4,
                                            "endRowIndex": 15,
                                            "startColumnIndex": 4,
                                            "endColumnIndex": 5
                                        }]
                                    }
                                }
                            }]
                        }
                    },
                    "position": {
                        "overlayPosition": {
                            "anchorCell": {
                                "sheetId": dashboard_id,
                                "rowIndex": 20,
                                "columnIndex": 7
                            }
                        }
                    }
                }
            }
        },

    
        # ================= FOR WHOM PIE =================
        {
            "addChart": {
                "chart": {
                    "spec": {
                        "title": "Expenses by For Whom (%)",
                        "pieChart": {
                            "legendPosition": "RIGHT_LEGEND",
                            "threeDimensional": False,
                            "domain": {
                            "sourceRange": {
                                "sources": [{
                                "sheetId": dashboard_id,
                                "startRowIndex": 20,
                                "endRowIndex": 35,
                                "startColumnIndex": 12,
                                "endColumnIndex": 13
                                }]
                            }
                            },
                            "series": {
                            "sourceRange": {
                                "sources": [{
                                "sheetId": dashboard_id,
                                "startRowIndex": 20,
                                "endRowIndex": 35,
                                "startColumnIndex": 13,
                                "endColumnIndex": 14
                                }]
                            }
                            }

                        }
                    },
                    "position": {
                        "overlayPosition": {
                            "anchorCell": {
                                "sheetId": dashboard_id,
                                "rowIndex": 35,
                                "columnIndex": 7
                            }
                        }
                    }
                }
            }
        }
    ]

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()

def add_highest_expense_value(service, spreadsheet_id):
    # Label
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!C1",
        valueInputOption="USER_ENTERED",
        body={"values":[["Highest Expense"]]}
    ).execute()

    # Value
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!C2",
        valueInputOption="USER_ENTERED",
        body={"values":[["=MAX(Expenses!G:G)"]]}
    ).execute()


def apply_conditional_formatting(service, spreadsheet_id):
    expenses_id = get_sheet_id(service, spreadsheet_id, "Expenses")

    requests = [
        # Overspend highlight (Amount > 5000)
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{
                        "sheetId": expenses_id,
                        "startRowIndex": 1,
                        "startColumnIndex": 6,
                        "endColumnIndex": 7
                    }],
                    "booleanRule": {
                        "condition": {
                            "type": "NUMBER_GREATER",
                            "values": [{"userEnteredValue": "5000"}]
                        },
                        "format": {
                            "backgroundColor": {
                                "red": 1.0,
                                "green": 0.85,
                                "blue": 0.85
                            }
                        }
                    }
                },
                "index": 0
            }
        },

        # Missing mandatory fields (Date / Category / Amount)
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [{
                        "sheetId": expenses_id,
                        "startRowIndex": 1,
                        "endColumnIndex": 7
                    }],
                    "booleanRule": {
                        "condition": {
                            "type": "CUSTOM_FORMULA",
                            "values": [{
                                "userEnteredValue":
                                '=OR($A2="", $D2="", $G2="")'
                            }]
                        },
                        "format": {
                            "backgroundColor": {
                                "red": 1.0,
                                "green": 0.95,
                                "blue": 0.8
                            }
                        }
                    }
                },
                "index": 1
            }
        }
    ]

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()

def add_dropdowns(service, spreadsheet_id):
    expenses_id = get_sheet_id(service, spreadsheet_id, "Expenses")

    def list_dropdown(col_index, values):
        return {
            "setDataValidation": {
                "range": {
                    "sheetId": expenses_id,
                    "startRowIndex": 1,
                    "startColumnIndex": col_index,
                    "endColumnIndex": col_index + 1
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": [{"userEnteredValue": v} for v in values]
                    },
                    "showCustomUi": True,
                    "strict": True
                }
            }
        }

    requests = [
        list_dropdown(3,  ["Food","Transport","Health","Utilities","Rent","Education","Loans","Shopping","Travel","Entertainment","Savings","Investment"]),
        list_dropdown(4,  ["Groceries","Dining","Fuel","Medicines","Electricity","Internet","EMI","Fees","Flight","Hotel","Shopping","Insurance"]),
        list_dropdown(7,  ["Cash","UPI","Credit Card","Debit Card","Bank Transfer"]),
        list_dropdown(8,  ["Cash","Navi","PhonePe","Paytm","SBI","Kotak811","CRED","Imobile"]),
        list_dropdown(9,  ["Chandru","Karthi","Appa","Amma","Pothu"]),
        list_dropdown(10, ["Self","Appa","Amma","Thambi","Anna","Anni","Family","Friends"]),
        list_dropdown(11, ["Essential","Discretionary","Savings","Investment","Loan"]),
        list_dropdown(12, ["One-time","Daily","Weekly","Monthly","Quarterly","Yearly"]),
        list_dropdown(13, ["Amazon","Flipkart","Uber","Ola","Local Store","Pharmacy","TNEB","Jio","Woman Self Help Group"]),
        list_dropdown(14, ["Yes","No"]),
        list_dropdown(15, ["Yes","No"]),
        list_dropdown(16, ["Personal","Family","Office","Medical","Travel","Emergency","Education","Tax"])
    ]

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()

def format_total_expense_card(service, spreadsheet_id):
    dashboard_id = get_sheet_id(service, spreadsheet_id, "Dashboard")

    requests = [

        # ===== TOTAL EXPENSE (A1:A2) =====
        {
            "repeatCell": {
                "range": {
                    "sheetId": dashboard_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "bold": True,
                            "fontSize": 14
                        }
                    }
                },
                "fields": "userEnteredFormat.textFormat"
            }
        },
        {
            "repeatCell": {
                "range": {
                    "sheetId": dashboard_id,
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "bold": True,
                            "fontSize": 18
                        },
                        "backgroundColor": {
                            "red": 0.90,
                            "green": 0.96,
                            "blue": 0.90
                        }
                    }
                },
                "fields": "userEnteredFormat(textFormat,backgroundColor)"
            }
        },

        # ===== CURRENT MONTH TOTAL (B1:B2) – STRONG HIGHLIGHT =====
        {
            "repeatCell": {
                "range": {
                    "sheetId": dashboard_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 1,
                    "endColumnIndex": 2
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "bold": True,
                            "fontSize": 15
                        }
                    }
                },
                "fields": "userEnteredFormat.textFormat"
            }
        },
        {
            "repeatCell": {
                "range": {
                    "sheetId": dashboard_id,
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 1,
                    "endColumnIndex": 2
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "bold": True,
                            "fontSize": 26
                        },
                        "backgroundColor": {
                            "red": 0.98,
                            "green": 0.90,
                            "blue": 0.80
                        }
                    }
                },
                "fields": "userEnteredFormat(textFormat,backgroundColor)"
            }
        },

        # ===== HIGHEST EXPENSE (C1:C2) – WARNING STYLE =====
        {
            "repeatCell": {
                "range": {
                    "sheetId": dashboard_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 2,
                    "endColumnIndex": 3
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "bold": True,
                            "fontSize": 14
                        }
                    }
                },
                "fields": "userEnteredFormat.textFormat"
            }
        },
        {
            "repeatCell": {
                "range": {
                    "sheetId": dashboard_id,
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": 2,
                    "endColumnIndex": 3
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "bold": True,
                            "fontSize": 18
                        },
                        "backgroundColor": {
                            "red": 0.98,
                            "green": 0.88,
                            "blue": 0.88
                        }
                    }
                },
                "fields": "userEnteredFormat(textFormat,backgroundColor)"
            }
        }
    ]

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()



def main():
    expenses, categories, family, payment, budget = create_test_data()
    export_excel(expenses, categories, family, payment, budget)

    creds = get_credentials()
    drive = build("drive","v3",credentials=creds)
    sheets = build("sheets","v4",credentials=creds)

    # 1️ Create Google Sheet
    spreadsheet_id = upload_sheet(drive)

    # 2️ Apply Month & Year formulas (already fixed)
    apply_month_year_formula(sheets, spreadsheet_id)

    create_dashboard(sheets, spreadsheet_id)

    add_highest_expense_value(sheets, spreadsheet_id)
    # add_current_month_total(sheets, spreadsheet_id)

    add_budget_actual_helper(sheets, spreadsheet_id)
    add_budget_vs_actual(sheets, spreadsheet_id)
    add_dashboard_section_titles(sheets, spreadsheet_id)
    
    add_for_whom_summary(sheets, spreadsheet_id)
    format_total_expense_card(sheets, spreadsheet_id)

    apply_conditional_formatting(sheets, spreadsheet_id)
    highlight_highest_expense(sheets, spreadsheet_id)
    highlight_budget_overrun(sheets, spreadsheet_id)

    add_dropdowns(sheets, spreadsheet_id)
    add_dashboard_charts(sheets, spreadsheet_id)

    # # 6️Monthly summary sheets (optional, already working)
    # for m in ["Jan-2026","Feb-2026"]:
    #     create_monthly_sheet(sheets, spreadsheet_id, m)

    print("SUCCESS")
    print("https://docs.google.com/spreadsheets/d/" + spreadsheet_id)


if __name__ == "__main__":
    main()
