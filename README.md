# MasterAccount-VB6-SQL-
# Busy VB6 Project

This project is a **Visual Basic 6 (VB6)** desktop application named **"Busy"**, designed to manage customer master data and handle billing-related information. It mimics features of popular inventory tools like **Tally** or **Busy Accounting Software**, with a focus on user input forms, data management, and SQL integration.

---

## 🧩 Features

- Add, edit, and delete customer records.
- Data grid displaying customer details.
- Search functionality by serial number or name.
- Compute opening balance credits and debits.
- Export customer data to CSV.
- Group-based logic for dealer types, tax structures, and location hierarchy.
- GUI-based navigation and keyboard shortcuts.
- Auto-increment and field validation in forms.
- SQL connectivity using **DSN** and **ADODB**.

---

## 🖥️ Forms & Modules

### 1. `Form1`
- Entry point of the application.
- Two buttons:
  - `Command1`: Opens **Add Master** form (`frmaddMaster`)
  - `Command2`: Opens **Grid View** (`Form2`)

### 2. `Form2`
- Displays customer records in a **Grid (grid1)**.
- Pulls data using ADODB from a MySQL database.
- Functions to display:
  - Serial Numbers
  - Names
  - Aliases
  - Group Names
  - Dealer Types
  - Opening balances (Cr/Dr)
  - GST numbers

### 3. `frmaddMaster`
- A detailed input form with 40+ fields including:
  - Personal details: name, alias, contact
  - Address and location
  - Banking info (Account No, IFSC, SWIFT)
  - Taxation details (GST, PAN, Aadhar, etc.)
- Conditional UI logic based on selected group
- CSV export functionality
- Insert/Update/Delete SQL queries
- Auto-suggestions and focus-based UX

---

## 💾 Database Connectivity

- Uses DSN: `form1`
- Username: `root`
- Password: `1234`
- Example query:
  ```sql
  SELECT * FROM emp WHERE srno = ...
