#  AI-Powered Financial Crime Detection System

This project simulates a **real-life fraud detection system** using **Python, Excel, and SQL Server** — designed to flag suspicious financial transactions using AI and automation.


# Project Overview

The system replicates a bank transfer environment with:

-  An Excel-based Transfer Form
-  An AI model (Isolation Forest) that flags abnormal transaction amounts
-  A secure SQL Server database storing all transactions
-  A separate `FraudAlerts` table for flagged entries
-  A live dashboard (Excel or Power BI) refreshing in real time


# Key Features

- **Auto-detects sender identity** using system login
- **Validates 10-digit account numbers only**
- **Unsupervised anomaly detection** using Isolation Forest
- **SQL Server integration** to log transactions and alerts
- **Real-time fraud alerts** via Excel or Power BI dashboard
- **Excel form** mimics a real banking interface with Send button automation


# Project Structure

```bash
├── detect_fraud.py                # Main Python logic for processing and fraud detection
├── requirements.txt               # Python dependencies
├── sql/
│   └── create_tables.sql          # SQL scripts to create Transactions and FraudAlerts tables
├── excel/
│   └── FinancialCrime.xlsm        # Excel form with VBA button
└── README.md
