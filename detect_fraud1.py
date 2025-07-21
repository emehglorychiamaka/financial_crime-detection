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
if not os.path.exists(excel_path):
    ctypes.windll.user32.MessageBoxW(0, " Excel form not found.", "File Error", 1)
    exit()

time.sleep(1)

try:
    wb = load_workbook(excel_path, data_only=True, keep_links=False)
    sheet = wb["TransferForm"]

    sender = getpass.getuser()
    receiver = sheet["G7"].value
    amount_cell = sheet["G8"].value
    bank = sheet["G9"].value
    acct_number = sheet["G10"].value

    # === Validation ===
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
    acct_number = str(acct_number).strip() if acct_number is not None else ""
    if not acct_number.isdigit() or len(acct_number) != 10:
        raise ValueError("Account number must be exactly 10 digits.")

    device_info = platform.node() or socket.gethostname()
    location = random.choice(["Lagos", "Abuja", "PH", "Enugu", "Kano", "Maiduguri"])  # Simulated
    now = datetime.now()
    hour = now.hour

    data = {
        "Sender": sender,
        "Receiver": str(receiver).strip(),
        "Amount": amount,
        "Bank": str(bank).strip(),
        "AccountNumber": acct_number,
        "Location": location,
        "Device": device_info,
        "Time": now,
        "Hour": hour
    }

except Exception as e:
    ctypes.windll.user32.MessageBoxW(0, f" Error in form: {e}", "Form Input Error", 1)
    exit()

# === 2. Fraud Detection ===
# Isolation Forest on 'Amount'
train = pd.DataFrame({"Amount": [random.gauss(50000, 8000) for _ in range(500)] + [random.uniform(9500, 9999) for _ in range(20)]})
model = IsolationForest(contamination=0.03, random_state=42)
model.fit(train[["Amount"]])
score = model.decision_function([[amount]])[0]
flag = model.predict([[amount]])[0]
is_anomaly = int(flag == -1)
fraud_score = float(round(abs(score), 4))

# === 3. Rule-based Checks for Specific Crimes ===
reasons = []

# a. Unauthorized Transfer (Odd hours)
offHourFlag = int(data["Hour"] < 6 or data["Hour"] > 22)
if offHourFlag:
    reasons.append("Unauthorized transfer during odd hours")

# b. Structuring (Smurfing)
structuringFlag = int(9800 <= amount <= 9999)
if structuringFlag:
    reasons.append("Structuring: Amount close to reporting threshold")

# c. Spoofing / Identity Theft (unknown device)
known_devices = ["Emeh-GLORY", "Emeh_GLory_C", "GL-PC", "HomePC"]
spoofingFlag = int(data["Device"] not in known_devices)
if spoofingFlag:
    reasons.append("Unrecognized device used")

# d. Account Takeover / Location laundering
known_locations = ["Lagos", "Abuja", "PH", "Abia", "Enugu"]
launderingFlag = int(location not in known_locations and amount >= 100000)
if launderingFlag:
    reasons.append(f"High-value transfer from suspicious location ({location})")

# e. Anomalous transaction pattern
if is_anomaly:
    reasons.append("Anomalous transaction pattern")

# === 4. Final Decision ===
is_fraud = int(len(reasons) > 0)

# === 5. Database Connection ===
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

# === 6. Insert into Transactions Table ===
try:
    cursor.execute("""
        INSERT INTO Transactions 
        (Sender, Receiver, Amount, Bank, AccountNumber, Location, Device, Time, IsFraud, FraudScore)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, data["Sender"], data["Receiver"], data["Amount"], data["Bank"], data["AccountNumber"],
         data["Location"], data["Device"], data["Time"], is_fraud, fraud_score)
except Exception as e:
    ctypes.windll.user32.MessageBoxW(0, f" Insert to Transactions failed:\n{e}", "Insert Error", 1)
    conn.close()
    exit()

# === 7. Insert into FraudAlerts Table ===
if is_fraud:
    try:
        reason_str = "; ".join(reasons)
        cursor.execute("""
            INSERT INTO FraudAlerts (Sender, Receiver, Amount, Bank, AccountNumber, Reason, Time, FraudScore)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, data["Sender"], data["Receiver"], data["Amount"], data["Bank"], data["AccountNumber"],
             reason_str, data["Time"], fraud_score)
    except Exception as e:
        ctypes.windll.user32.MessageBoxW(0, f" Insert to FraudAlerts failed:\n{e}", "Insert Error", 1)

conn.commit()
cursor.close()
conn.close()

# === 8. Show Result ===
formatted_amount = "₦{:,.2f}".format(data["Amount"])
if is_fraud:
    msg = f" FRAUD DETECTED!\nSender: {data['Sender']}\nAmount: {formatted_amount}\n\nReasons:\n" + "\n".join(reasons) + f"\n\nFraud Score: {fraud_score}"
else:
    msg = f" Transaction successful!\nAmount: {formatted_amount}\n\nFraud Score: {fraud_score}"

ctypes.windll.user32.MessageBoxW(0, msg, "Transaction Result", 1)
