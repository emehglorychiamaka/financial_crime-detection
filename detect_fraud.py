import pandas as pd
from openpyxl import load_workbook
from sklearn.ensemble import IsolationForest
from datetime import datetime
import pyodbc
import ctypes
import random
import platform
import socket
import time
import os
import getpass

# === 1. Load Excel Form ===
excel_path = r"C:\Users\emehg\Desktop\3MTT_Data_Analysis\FinancialCrime\FinancialCrime.xlsm"

# Ensure the file exists and give Excel time to save
if not os.path.exists(excel_path):
    ctypes.windll.user32.MessageBoxW(0, " Excel form not found.", "File Error", 1)
    exit()

# Wait briefly to let Excel flush changes
time.sleep(1)

try:
    wb = load_workbook(excel_path, data_only=True, keep_links=False)
    sheet = wb["TransferForm"]

    # Read values
    # Automatically detect sender from logged-in user
    sender = getpass.getuser()
    receiver = sheet["G7"].value
    amount_cell = sheet["G8"].value
    bank = sheet["G9"].value
    acct_number = sheet["G10"].value
    

    # Validate each input
    #if sender is None or str(sender).strip() == "":
     #   raise ValueError("Sender name (G6) is missing or empty.")
    if receiver is None or str(receiver).strip() == "":
        raise ValueError("Receiver name (G7) is missing or empty.")
    if amount_cell is None or str(amount_cell).strip() == "":
        raise ValueError("Amount (G8) is missing.")
    try:
        amount = float(amount_cell)
    except:
        raise ValueError("Amount must be a valid number.")
    if bank is None or str(bank).strip() == "":
        raise ValueError("Receiving bank (G9) is missing.")
    #if acct_number is None or str(acct_number).strip() == "":
        #raise ValueError("Account number (G10) is missing.")
    acct_number = str(acct_number).strip() if acct_number is not None else ""
    if not acct_number.isdigit() or len(acct_number) != 10:
        raise ValueError("Account number must be exactly 10 digits.")

    device_info = platform.node() or socket.gethostname()
    location = random.choice(["Lagos", "Abuja", "PH", "Enugu", "Kano"])

    data = {
        #"Sender": str(sender).strip(),
        "Sender": sender,
        "Receiver": str(receiver).strip(),
        "Amount": amount,
        "Bank": str(bank).strip(),
        "AccountNumber": str(acct_number).strip(),
        "Location": location,
        "Device": device_info,
        "Time": datetime.now()
    }

except Exception as e:
    ctypes.windll.user32.MessageBoxW(0, f" Error in form: {e}", "Form Input Error", 1)
    exit()

# === 2. Fraud Detection ===
train = pd.DataFrame({
    "Amount": [random.gauss(50000, 8000) for _ in range(300)]
})
model = IsolationForest(contamination=0.05, random_state=42)
model.fit(train[["Amount"]])

df = pd.DataFrame([data])
score = model.decision_function(df[["Amount"]])[0]
flag = model.predict(df[["Amount"]])[0]

is_fraud = int(flag == -1)
fraud_score = float(round(abs(score), 4))

# === 3. Connect to SQL Server ===
try:
    conn = pyodbc.connect(
        r"DRIVER={ODBC Driver 17 for SQL Server};"
        r"SERVER=Emeh_GLory_C\SQLEXPRESS;"
        r"DATABASE=FinancialCrime;"
        r"Trusted_Connection=yes;"
    )
    cursor = conn.cursor()
except Exception as e:
    ctypes.windll.user32.MessageBoxW(0, f" DB connection failed:\n{e}", "DB Error", 1)
    exit()

# === 4. Insert into Transactions Table ===
try:
    cursor.execute("""
        INSERT INTO Transactions (Sender, Receiver, Amount, Bank, AccountNumber, Location, Device, Time, IsFraud, FraudScore)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, data["Sender"], data["Receiver"], data["Amount"], data["Bank"], data["AccountNumber"],
         data["Location"], data["Device"], data["Time"], is_fraud, fraud_score)
except Exception as e:
    ctypes.windll.user32.MessageBoxW(0, f" Insert to Transactions failed:\n{e}", "Insert Error", 1)
    conn.close()
    exit()

# === 5. Insert into FraudAlerts Table ===
if is_fraud:
    try:
        reason = "Unusual transaction amount"
        cursor.execute("""
            INSERT INTO FraudAlerts (Sender, Receiver, Amount, Bank, AccountNumber, Reason, Time, FraudScore)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, data["Sender"], data["Receiver"], data["Amount"], data["Bank"], data["AccountNumber"],
             reason, data["Time"], fraud_score)
    except Exception as e:
        ctypes.windll.user32.MessageBoxW(0, f" Insert to FraudAlerts failed:\n{e}", "Insert Error", 1)

conn.commit()
cursor.close()
conn.close()

# === 6. Show Confirmation ===
formatted_amount = "₦{:,.2f}".format(data["Amount"])
if is_fraud:
    msg = f"🚨 FRAUD DETECTED!\nSender: {data['Sender']}\nAmount: {formatted_amount}\nScore: {fraud_score}"
else:
    msg = f" Transaction complete!\nAmount: {formatted_amount}\nScore: {fraud_score}"

ctypes.windll.user32.MessageBoxW(0, msg, "Transaction Result", 1)