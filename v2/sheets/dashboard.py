from utils.sheets_utils import get_sheet_id

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

    # ----- KPI: Selected Month Total -----
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!B1",
        valueInputOption="USER_ENTERED",
        body={"values": [["Current Month Total"]]}
    ).execute()

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!B2",
        valueInputOption="USER_ENTERED",
        body={"values": [['=SUMIFS(Expenses!G:G, Expenses!B:B, O2)']]}
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
            '=QUERY(Expenses!A:R, '
            '"select D,sum(G) '
            'where B = \'" & O2 & "\' '
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
           '=QUERY(Expenses!A:R, '
            '"select H,sum(G) '
            'where B = \'" & O2 & "\' '
            'group by H order by sum(G) desc '
            'label H \'Payment Mode\', sum(G) \'Amount\'")'

        ]]}
    ).execute()

    # ----- For Whom Summary -----
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!G5",
        valueInputOption="USER_ENTERED",
        body={"values": [[
            '=QUERY(Expenses!A:R, '
            '"select K,sum(G) '
            'where B = \'" & O2 & "\' '
            'group by K order by sum(G) desc '
            'label K \'For Whom\', sum(G) \'Amount\'")'

        ]]}
    ).execute()
    
    # ----- Paid By Summary -----
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!J5",
        valueInputOption="USER_ENTERED",
        body={"values": [[
                '=QUERY(Expenses!A:R, ' 
            '"select J,sum(G) '
            'where B = \'" & O2 & "\' '
            'group by J order by sum(G) desc '
            'label J \'Paid By\', sum(G) \'Amount\'")'
        ]]}
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

def add_individual_salary_dashboard_table(service, spreadsheet_id):
    formula = (
        '=QUERY(Individual_Salary!A:E,'
        '"select A,B,C,D,E '
        'where A = \'" & O2 & "\' '
        'label A \'Month\', B \'Person\', '
        'C \'Salary\', D \'Actual Spent\', E \'Remaining\'")'

    )

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!S2",
        valueInputOption="USER_ENTERED",
        body={"values": [[formula]]}
    ).execute()



def create_meta_month_list(service, spreadsheet_id):
    """
    Creates hidden _Meta sheet to store unique month list.
    """

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

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="_Meta!A2",
        valueInputOption="USER_ENTERED",
        body={
            "values": [[
                '=SORT(UNIQUE(FILTER(Expenses!B2:B, Expenses!B2:B<>"")))'
            ]]
        }
    ).execute()

def add_month_selector(service, spreadsheet_id):
    dashboard_id = get_sheet_id(service, spreadsheet_id, "Dashboard")

    # Label
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!O1",
        valueInputOption="USER_ENTERED",
        body={"values": [["Select Month"]]}
    ).execute()

    # CRITICAL: Format O2 and P column as plain text FIRST
    # This prevents Google Sheets from auto-converting "Jan-2026" → date 01/01/2026
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": dashboard_id,
                        "startRowIndex": 1,
                        "endRowIndex": 2,
                        "startColumnIndex": 14,
                        "endColumnIndex": 15
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "TEXT"
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat"
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": dashboard_id,
                        "startRowIndex": 1,
                        "endRowIndex": 101,
                        "startColumnIndex": 15,
                        "endColumnIndex": 16
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "numberFormat": {
                                "type": "TEXT"
                            }
                        }
                    },
                    "fields": "userEnteredFormat.numberFormat"
                }
            }
        ]}
    ).execute()

    # Helper list in P column (used as dropdown source)
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!P1",
        valueInputOption="USER_ENTERED",
        body={"values": [["Month List"]]}
    ).execute()

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!P2",
        valueInputOption="USER_ENTERED",
        body={"values": [[
            '=SORT(UNIQUE(FILTER(Expenses!B2:B, Expenses!B2:B<>"")))'
        ]]}
    ).execute()

    # Data Validation dropdown pointing to P column
    request = {
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
                        "userEnteredValue": "=Dashboard!P2:P100"
                    }]
                },
                "showCustomUi": True,
                "strict": False
            }
        }
    }

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [request]}
    ).execute()

    # Default month (latest = last non-empty in P column)
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range="Dashboard!O2",
        valueInputOption="USER_ENTERED",
        body={"values": [[
            '=INDEX(P2:P, COUNTA(P2:P))'
        ]]}
    ).execute()

