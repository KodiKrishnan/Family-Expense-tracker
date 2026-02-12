import sys
from pathlib import Path

# Add the project root to Python path
sys.path.insert(0, str(Path(__file__).parent))

from googleapiclient.discovery import build

# ---- AUTH ----
from auth.google_auth import get_credentials

# ---- DATA ----
from data.seed_data import create_test_data, export_excel

# ---- UTILS ----
from utils.sheets_utils import upload_sheet

# ---- EXPENSES ----
from sheets.expenses import (
    apply_month_year_formula,
    add_dropdowns,
    apply_conditional_formatting,
    highlight_highest_expense
)

# ---- DASHBOARD ----
from sheets.dashboard import (
    create_dashboard,
    format_total_expense_card,
    add_month_selector,
)

# ---- CHARTS ----
from sheets.charts import add_dashboard_charts

# ---- SALARY ----
from sheets.salary import (
    create_individual_salary_sheet,
    compute_individual_salary_actuals,
    highlight_negative_remaining_individual_salary,
    add_individual_salary_dashboard_table
)


def main():
    # 1️⃣ Create base Excel
    expenses = create_test_data()
    export_excel(expenses)
    print("Excel created")

    # 2️⃣ Authenticate
    creds = get_credentials()
    drive = build("drive", "v3", credentials=creds)
    sheets = build("sheets", "v4", credentials=creds)
    print("Authenticated")

    # 3️⃣ Upload to Google Sheets
    spreadsheet_id = upload_sheet(drive)
    print("Spreadsheet created")

    # 4️⃣ Expenses fixes
    print("Applying expenses")
    apply_month_year_formula(sheets, spreadsheet_id)
    add_dropdowns(sheets, spreadsheet_id)
    apply_conditional_formatting(sheets, spreadsheet_id)
    highlight_highest_expense(sheets, spreadsheet_id)

    # 5️⃣ Dashboard
    print("Creating dashboard...")
    create_dashboard(sheets, spreadsheet_id)
    add_month_selector(sheets, spreadsheet_id)
    format_total_expense_card(sheets, spreadsheet_id)
    
    # 6️⃣ Individual Salary
    print("Creating salary sheet...")
    create_individual_salary_sheet(sheets, spreadsheet_id)
    compute_individual_salary_actuals(sheets, spreadsheet_id)
    highlight_negative_remaining_individual_salary(sheets, spreadsheet_id)
    add_individual_salary_dashboard_table(sheets, spreadsheet_id)

    # 7️⃣ Charts (LAST)
    print("Adding charts...")
    add_dashboard_charts(sheets, spreadsheet_id)
    # add_individual_salary_charts(sheets, spreadsheet_id)

    print("SUCCESS")
    print(f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}")


if __name__ == "__main__":
    main()
