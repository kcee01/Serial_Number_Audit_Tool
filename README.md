# Serial Number Audit Tool

This tool automates the monthly auditing of printer serial numbers and related meter readings for Nashua devices. It compares data between the current and previous month to detect changes,
missing entries, and updates values accordingly.

---

## ✅ Features

- ✅ Adds **"Nashua Serial Number"** column from the previous month's file.
- ✅ Compares serial numbers across months and identifies:
  - Missing entries
  - Matching entries
- ✅ Outputs unmatched or new serial numbers in a separate sheet.
- ✅ Identifies serial numbers not listed in **"Nashua Serial Numbers"**.
- ✅ Writes changes directly to the **original spreadsheet file** (no new output file needed).
- ✅ Removes **duplicate serial numbers** from the "Serial Number" column.
- ✅ Automatically populates the **"B/W Start Meter"** column using the **"B/W End Meter"** column from the previous month.

Output logs are::in Excel fomat
bw_start_meter_log
matched_serials
unmatched_serials
duplicates
---

## 📁 File Requirements

- Excel files (`.xlsx`) for both current and previous months.
- Each file should include at least:
  - `Serial Number` column
  - `B/W End Meter` column
  - `Nashua Serial Number` column

---

## 🚀 Getting Started

1. Clone the repository:

   ```bash
   git clone https://github.com/kcee01/Serial_Number_Audit_Tool.git
   cd Serial_Number_Audit_Tool


🛠 Dependencies
pandas
openpyxl

📬 Contact
Developer: Cliff Keabetswe
Email: innocliffkeab@gmail.com

📝 License
This project is open-source and available under the MIT License.
