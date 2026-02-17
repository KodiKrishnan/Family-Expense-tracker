import pandas as pd

def create_test_data():
    
    expenses = pd.read_csv("/media/kodikrishnan/HDD-linux/Kodi/Expense-tracker/v2/data/expenses.csv")

    return expenses

def export_excel(expenses):
    with pd.ExcelWriter("temp.xlsx", engine="openpyxl") as writer:
        expenses.to_excel(writer, sheet_name="Expenses", index=False)