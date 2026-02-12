import pandas as pd

def create_test_data():
    
    expenses = pd.DataFrame([
    ["2026-01-04","Jan-2026",2026,"Loans","EMI","Apty Kalanchiam",5000,"Cash","Cash","Deiva","Family","Loan","Monthly","Woman Self Help Group","Yes","No","Family"],
    ["2026-01-15","Jan-2026",2026,"Utilities","Internet","Act Wifi -BLR",883.82,"Credit Card","Imobile","Karthi","Karthi","Discretionary","Monthly","Others","Yes","No","Personal"],
    ["2026-01-05","Feb-2026",2026,"Rent","Housing","Karthi-House",8500,"UPI","GPay","Karthi","Karthi","Essential","Monthly","House Owners","No","No","Personal"]
    ], columns=[
        "Date","Month","Year","Category","Sub-Category","Description","Amount",
        "Payment Mode","Account","Paid By","For Whom","Expense Type","Frequency",
        "Vendor","Bill?","Reimbursable","Tags"
    ])

    return expenses

def export_excel(expenses):
    with pd.ExcelWriter("temp.xlsx", engine="openpyxl") as writer:
        expenses.to_excel(writer, sheet_name="Expenses", index=False)