from utils.sheets_utils import get_sheet_id


def add_dashboard_charts(service, spreadsheet_id):
    dashboard_id = get_sheet_id(service, spreadsheet_id, "Dashboard")

    # ================= CLEAR OLD CHARTS =================
    spreadsheet = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id
    ).execute()

    delete_requests = []

    for sheet in spreadsheet["sheets"]:
        if sheet["properties"]["title"] == "Dashboard":
            for chart in sheet.get("charts", []):
                delete_requests.append({
                    "deleteEmbeddedObject": {
                        "objectId": chart["chartId"]
                    }
                })

    if delete_requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": delete_requests}
        ).execute()

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
                                        "endRowIndex": 20,
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
                                        "endRowIndex": 20,
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
                                "rowIndex": 0,
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
                                            "endRowIndex": 20,
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
                                            "endRowIndex": 20,
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
                                "rowIndex": 18,
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
                                "startRowIndex": 4,
                                "endRowIndex": 20,
                                "startColumnIndex": 6,
                                "endColumnIndex": 7
                                }]
                            }
                            },
                            "series": {
                            "sourceRange": {
                                "sources": [{
                                "sheetId": dashboard_id,
                                "startRowIndex": 4,
                                "endRowIndex": 20,
                                "startColumnIndex": 7,
                                "endColumnIndex": 8
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
        },

        # ================= PAID BY PIE =================
        {
            "addChart": {
                "chart": {
                    "spec": {
                        "title": "Individual Spending (%)",
                        "pieChart": {
                            "legendPosition": "RIGHT_LEGEND",
                            "threeDimensional": False,
                            "domain": {
                                "sourceRange": {
                                    "sources": [{
                                        "sheetId": dashboard_id,
                                        "startRowIndex": 4,
                                        "endRowIndex": 20,
                                        "startColumnIndex": 9,
                                        "endColumnIndex": 10
                                    }]
                                }
                            },
                            "series": {
                                "sourceRange": {
                                    "sources": [{
                                        "sheetId": dashboard_id,
                                        "startRowIndex": 4,
                                        "endRowIndex": 20,
                                        "startColumnIndex": 10,
                                        "endColumnIndex": 11
                                    }]
                                }
                            }
                        }
                    },
                    "position": {
                        "overlayPosition": {
                            "anchorCell": {
                                "sheetId": dashboard_id,
                                "rowIndex": 55,
                                "columnIndex": 7
                            }
                        }
                    }
                }
            }
        },

        # ================= SALARY VS SPENDING =================
        {
            "addChart": {
                "chart": {
                    "spec": {
                        "title": "Salary vs Spent vs Remaining",
                        "basicChart": {
                            "chartType": "COLUMN",
                            "legendPosition": "RIGHT_LEGEND",
                            "headerCount": 1,
                            "domains": [{
                                "domain": {
                                    "sourceRange": {
                                        "sources": [{
                                            "sheetId": dashboard_id,
                                            "startRowIndex": 1,
                                            "endRowIndex": 10,
                                            "startColumnIndex": 19,
                                            "endColumnIndex": 20
                                        }]
                                    }
                                }
                            }],
                            "series": [
                                # Salary
                                {
                                    "series": {
                                        "sourceRange": {
                                            "sources": [{
                                                "sheetId": dashboard_id,
                                                "startRowIndex": 1,
                                                "endRowIndex": 10,
                                                "startColumnIndex": 20,
                                                "endColumnIndex": 21
                                            }]
                                        }
                                    }
                                },
                                # Actual Spent
                                {
                                    "series": {
                                        "sourceRange": {
                                            "sources": [{
                                                "sheetId": dashboard_id,
                                                "startRowIndex": 1,
                                                "endRowIndex": 10,
                                                "startColumnIndex": 21,
                                                "endColumnIndex": 22
                                            }]
                                        }
                                    }
                                },
                                # Remaining
                                {
                                    "series": {
                                        "sourceRange": {
                                            "sources": [{
                                                "sheetId": dashboard_id,
                                                "startRowIndex": 1,
                                                "endRowIndex": 10,
                                                "startColumnIndex": 22,
                                                "endColumnIndex": 23
                                            }]
                                        }
                                    }
                                }
                            ]
                        }
                    },
                    "position": {
                        "overlayPosition": {
                            "anchorCell": {
                                "sheetId": dashboard_id,
                                "rowIndex": 75,
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
