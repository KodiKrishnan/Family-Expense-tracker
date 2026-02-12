from utils.sheets_utils import get_sheet_id

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
            #Category
            list_dropdown(3, [
                "Food","Transport","Health","Utilities","Rent","Education","Loans",
                "Shopping","Entertainment","Savings","Investment","Others","Debt",
                "Insurance","Household Maintenance","Festivals/Celebrations"
            ]),

            #Sub-Category
            list_dropdown(4, [
                "Groceries","Mobile","Dining","Fuel","Medicines","Electricity","Internet","EMI",
                "Fees","Flight","Hotel","Shopping","Chit Funds","Others","Bus","Train",
                "Taxi","Snacks","Maintenance","Skill Development","Interest","Housing",
                "Water Bill","Gas Cylinder","Subscriptions","Festive Expenses"
            ]),

             #Payment Mode
            list_dropdown(7, [
                "Cash","UPI","Credit Card","Debit Card","Bank Transfer","Net Banking","Wallets"
            ]),

            # Account / Payment Platform
            list_dropdown(8, [
                "Cash","Navi","PhonePe","Paytm","PhonePay","Kotak811","CRED","Imobile",
                "UBI","Gpay","Amazon Pay","Others","PayPal"
            ]),

            # Paid By
            list_dropdown(9, [
                "Chandru","Karthi","Appa","Amma","Pothu","Deiva","Family","Self","Children"
            ]),

            # For Whom (Beneficiary)
            list_dropdown(10, [
                "Chandru","Karthi","Appa","Amma","Pothu","Deiva","Family","Self",
                "Friends","Office","Children","Relatives","Patti","Athai","Others"
            ]),

            # Expense Type (Purpose)
            list_dropdown(11, [
                "Essential","Discretionary","Repayment","Savings","Investment","Loan",
                "Medical","Emergency","Tax","Travel","Education","Others","Clothes",
                "Insurance Premiums","Charity/Donations","Household Maintenance","Festivals/Celebrations"
            ]),

            # Frequency
            list_dropdown(12, [
                "One-time","Daily","Weekly","Monthly","Quarterly","Yearly","Bi-Annual"
            ]),

            # Vendor / Provider
            list_dropdown(13, [
                "Amazon","Flipkart","Uber","Ola","Local Store","Pharmacy","TNEB","Jio",
                "Woman Self Help Group","Bank","Others","Bus Corporation","Finance Company",
                "House Owners","Insurance Company","School/College","Streaming Services","Government"
            ]),

            # Bill?
            list_dropdown(14, ["Yes","No"]),

            # Reimbursable?
            list_dropdown(15, ["Yes","No"]),

            # Tags (Category Type)
            list_dropdown(16, [
                "Personal","Family","Office","Medical","Travel","Emergency","Education",
                "Tax","Self","Chit Fund","Festivals/Celebrations","Insurance","Charity/Donation"
            ])
        ]

    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()