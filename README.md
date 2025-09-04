# VN01 Arrow Ship Mode & OH Merge Automation

Automates merging of VN01 On-Hand (OH) Parts data with Arrow Ship Mode (ASM) files and outputs a consolidated Excel file filtered by specific criteria.

This script is designed to improve reporting efficiency for the Europe & Other Asia Team at Jabil by reducing manual Excel handling and ensuring consistency in weekly Arrow shipment reporting.

---

## Features

- Reads VN01 OH Parts Excel file and Arrow Ship Mode Excel file for the current ISO week.
- Handles merged or missing values using forward fill (`ffill`) to ensure clean data.
- Ensures consistent invoice numbers by filling them down if CUSTOMER PO values match.
- Standardizes column names for compatibility between OH and ASM files.
- Merges OH and ASM data on CUSTOMER PART NUMBER.
- Filters merged data for PGr codes starting with 'U'.
- Selects only the required columns for reporting.
- Outputs a weekly Excel file named 'JB <file_name>.xlsx' in the corresponding Arrow Ship Mode folder.
- Prints a list of unique Buyer Names in the console.

---

## How It Works

1. **Determine Current Week & File Paths**
   - Automatically generates paths based on the current ISO week for ASM files.
   - Uses the current week’s Monday date to locate OH part files.

2. **Data Processing**
   - Reads Excel files using pandas.
   - Performs forward fill on key columns (MPN, CUSTOMER PART NUMBER, CUSTOMER PO, QUANTITY SHIPPED, UNIT PRICE, Buyer).
   - Copies invoice numbers down rows if CUSTOMER PO values are repeated.
   - Merges OH and ASM data on CUSTOMER PART NUMBER.
   - Filters merged data for specific PGr codes.
   - Selects desired columns for the final report.

3. **Output**
   - Saves the consolidated Excel report in the structured folder path:
     //sgsind0nsifsv01a/.../ARROW SHIP MODE/WWxx'yy/JB <file_name>.xlsx
   - Prints unique Buyer Names for quick reference.
  
  ---
