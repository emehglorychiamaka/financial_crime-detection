CREATE DATABASE FinancialCrime;
GO

USE FinancialCrime;
GO

-- Main transactions table
CREATE TABLE Transactions (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    Sender VARCHAR(100),
    Receiver VARCHAR(100),
    Amount FLOAT,
    Location VARCHAR(100),
    Device VARCHAR(100),
    Time DATETIME,
    IsFraud BIT,
    FraudScore FLOAT
);

-- Separate fraud alerts
CREATE TABLE FraudAlerts (
    AlertId INT IDENTITY(1,1) PRIMARY KEY,
    Sender VARCHAR(100),
    Receiver VARCHAR(100),
    Amount FLOAT,
    Reason VARCHAR(255),
    Time DATETIME,
    FraudScore FLOAT
);

-- Rename "Location" to "Bank"
EXEC sp_rename 'Transactions.Location', 'Bank', 'COLUMN';

-- Rename Device to Account Number
Exec sp_rename 'Transactions.Device', 'AccountNumber', 'COLUMN';


-- Add columns to Transactions table
ALTER TABLE Transactions
ADD Location NVARCHAR(100), 
Device NVARCHAR(100);


-- Add Location and Device to FraudAlerts 
ALTER TABLE FraudAlerts
ADD Location NVARCHAR(100),
Device NVARCHAR(100);
 
-- use the correct db

USE FinancialCrime;
GO

-- Drop existing tables if needed (optional, for fresh start)
DROP TABLE FraudAlerts;
DROP TABLE Transactions;

-- === Create Transactions table ===
CREATE TABLE Transactions (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    Sender VARCHAR(100),
    Receiver VARCHAR(100),
    Amount FLOAT,
    Bank VARCHAR(100),
    AccountNumber VARCHAR(100),
    Location VARCHAR(100),
    Device VARCHAR(100),
    Time DATETIME,
    IsFraud BIT,
    FraudScore FLOAT
);

-- === Create FraudAlerts table with complete columns ===
CREATE TABLE FraudAlerts (
    AlertId INT IDENTITY(1,1) PRIMARY KEY,
    Sender VARCHAR(100),
    Receiver VARCHAR(100),
    Amount FLOAT,
    Bank VARCHAR(100),
    AccountNumber VARCHAR(100),
    Reason VARCHAR(255),
    Time DATETIME,
    FraudScore FLOAT,
    Location NVARCHAR(100),
    Device NVARCHAR(100)
);


Select *
from Transactions;

SELECT *
FROM FraudAlerts;

Select @@servername as servername;
