# Export Consumables Reports from Access to Excel

## Overview
This VBA solution automates the process of exporting **Consumables** data from Microsoft Access to Excel workbooks, applying formatting, inserting the company logo, generating PivotTables, and saving files into structured monthly folders.

It is designed to:
- Filter customer data based on specific `Segment4` values.
- Insert company logos stored as attachments in the `Company` table.
- Build and format **Data** and **Overview** sheets in Excel.
- Generate PivotTables from the exported data.
- Save Excel files in monthly subfolders.
- Log any problems during the process.

---

## Requirements
- Microsoft Access with VBA enabled.
- Microsoft Excel installed.
- Tables:
  - **Customer** – customer details and export paths.
  - **Company** – includes `CompanyLogo` attachment field.
  - **Consumables** – main data source for exports.

---

## Table Relationships
- One-to-many relationship between:
  - `Company.CompanyID` and `Customer.Entity`
- `Customer.Customer_Name` matches `Consumables.Customer`.

---

## Key Features
1. **Automatic Folder Creation**
   - Files are saved inside a monthly folder named as `YYYYMM`.
   - If the folder already exists, files are saved without errors.

2. **Company Logo Insertion**
   - Reads the `CompanyLogo` attachment from the `Company` table.
   - Saves it temporarily and embeds it in both **Data** and **Overview** sheets.

3. **PivotTable Generation**
   - PivotTable fields are configured automatically.
   - Summarizes `Amount` per `Posting Month` and other dimensions.

4. **Custom Ledger Display**
   - If multiple companies are found in the `Ledger` field, their last 3 characters are shown in the format:
     ```
     [GHD,SAB,CLA]
     ```
     in:
     - `Data!B11`
     - `Overview!B10`

5. **Formatting**
   - Fonts, borders, merged cells, and headers applied consistently.
   - AutoFilter with highlighted header row.

6. **Logging**
   - Any failed exports are logged in `ExportLog.txt` with a timestamp and customer details.

---

## Process Flow

```text
+------------------+
|  Start Export    |
+--------+---------+
         |
         v
+------------------+
| Read Customers   |
| from table       |
+--------+---------+
         |
         v
+------------------+
| Query            |
| Consumables data |
+--------+---------+
         |
         v
+------------------+
| If no data ->    |
| Skip to next     |
+--------+---------+
         |
         v
+------------------+
| Create Excel WB  |
| Add Data Sheet   |
+--------+---------+
         |
         v
+------------------+
| Insert Company   |
| Logo in Sheets   |
+--------+---------+
         |
         v
+------------------+
| Format Headers   |
| & AutoFilter     |
+--------+---------+
         |
         v
+------------------+
| Create Pivot     |
| in Overview      |
+--------+---------+
         |
         v
+------------------+
| Save File to     |
| YYYYMM Folder    |
+--------+---------+
         |
         v
+------------------+
| Next Customer    |
+------------------+

## Author
Developed by **Dumitru Purce**  
📅 Last updated: `2025-08-14`
