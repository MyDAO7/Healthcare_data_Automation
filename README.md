# Healthcare_data_Automation
Automate monthly healthcare data processing across multiple clinics  from messy Excel, CSV files to professional reports emailed automatically in 30 seconds
#  Healthcare Data Automation

## *Stop wasting Fridays on Excel. Let the code do the work.*

### The Real Story

A clinic manager once told me:

> *"I spend every Friday afternoon opening files from 3 different clinics, copying data, fixing column names, and creating reports. I hate Fridays."*

That stuck with me.

So I built a tool that turns **3 hours of manual work into 30 seconds of waiting.**

**This is that tool.**

---

### What It Does (In Plain English)

You have Excel files. Maybe from Clinic A, Clinic B, and Clinic C. Maybe some are CSV. Maybe the columns are named differently. Maybe the dates are in different formats.

This script:

- **Finds** every Excel and CSV file in your folder
- **Cleans** messy data (missing values, blank rows)
- **Fixes** column names (whether it says "Patient Name", "Name", or "Patient", it becomes "Patient_Name")
- **Standardizes** dates (turns "06/01/2024", "2024-06-01", "June 1, 2024" into the same format)
- **Combines** everything into one master sheet
- **Summarizes** your data by clinic (totals, counts, averages)
- **Formats** your report like a pro (blue headers, borders, currency)
- **Emails** it to your team automatically

And if one file is corrupt? The others still work. The error log tells you which file failed.

---

### Before vs After

#### Before (Manual Process)

1-Open Clinic_A_June_2024.xlsx
2-Copy data
3-Paste into master sheet
4-Repeat for Clinic_B, Clinic_C
5-Fix column names (one says "Patient Name", another says "Name")
6-Fix dates (some are MM/DD, some are DD/MM)
7-Remove duplicate rows
8-Fill missing values
9-Create summary by clinic
10-Format headers, add borders
11-Double-check everything
12-Email to manager
Time: 2-3 hours
Frustration: High
Chance of error: Very high

#### After (This Tool)

1-Put all files in Clients_data/
2-Run: python main.py
3-Get report in your email

Time: 30 seconds
Frustration: Zero
Chance of error: Zero


---

### What You Get

#### 1. Master Sheet (All Data Combined)

Every patient record from every clinic, cleaned and standardized.

| Date | Patient_Name | Doctor | Service_Type | Amount | Status | Location |
|------|--------------|--------|--------------|--------|--------|----------|
| 2024-06-01 | John Smith | Dr. Patel | Consultation | 150 | Paid | Clinic_A |
| 2024-06-03 | Maria Garcia | Dr. Lee | Lab | 90 | Pending | Clinic_B |
| 2024-06-04 | Alice Cooper | Dr. Shah | General | 175 | Paid | Clinic_C |

#### 2. Clinic Summary (What You Actually Care About)

No more pivot tables. Just the numbers you need.

| Location | Total_Amount | Total_Patient | Avg_Sale |
|----------|--------------|---------------|----------|
| Clinic_A | 4,850 | 10 | 485.00 |
| Clinic_B | 6,675 | 10 | 667.50 |
| Clinic_C | 4,950 | 10 | 495.00 |

#### 3. Error Log (If Something Goes Wrong)

If a file is corrupt or has the wrong format, it's listed here. The rest of your files still process.

| Error |
|-------|
| Clinic_C_June_2024.csv: Column 'Patient' not found |

---
#### 4. Interactive dashboard(live presentation of data)
Instead of just emailing the Excel report, you can now explore the data live.

### Run it

    streamlit run dashboard.py

Opens at http://localhost:8501
### Screenshots
<img width="1702" height="638" alt="image" src="https://github.com/user-attachments/assets/40556cf7-70a3-4642-943d-db796f869c94" />
<img width="1840" height="619" alt="image" src="https://github.com/user-attachments/assets/22799ead-dbd0-4f01-af4d-b184f699153e" />
<img width="1848" height="769" alt="image" src="https://github.com/user-attachments/assets/16011eb5-ef49-4782-8f60-cc3cb9b19120" />


### What you get

- **KPI cards** — Total revenue, patients, avg sale, paid %
- **Filters** — Clinic, doctor, service type, status, date range
- **Charts** — Revenue over time, revenue by service, payment status, top doctors
- **Raw data tab** — Filter and download the cleaned data as CSV
### Why This Matters

| Real World Problem | How This Tool Solves It |
|-------------------|------------------------|
| One file says "Patient Name", another says "Name" | Both become "Patient_Name" |
| One file says "Total", another says "Amount" | Both become "Amount" |
| Dates like "06/01/2024" and "2024-06-01" | All become "2024-06-01" |
| Missing patient name | Filled with "Unknown" |
| Missing amount | Filled with average from that clinic |
| Missing status | Set to "Pending" |
| Blank rows | Removed automatically |
| Duplicate rows | Removed automatically |
| File is open in Excel | Handled automatically (pywin32) |

---

### Quick Start (3 Minutes)

**Step 1: Get the code**
```bash
git clone https://github.com/MyDAO7/Healthcare_data_Automation.git
cd Healthcare_data_Automation
```
**Step 2: Install requirements**
```bash
pip install -r requirement.txt
```
**3. Set Up Environment Variables**
Create a .env file in the root directory:
email=your_email@gmail.com
password=your_app_password
recipient=client@example.com
**For Gmail, use an App Password — not your regular password**

**4. Place Your Files**
Add your Excel and CSV files to the Clients_data/ folder.

**5. Run the Automation**
```bash
python main.py
```
**6. Check Your Email**
The report will be sent to the recipient you specified in .env.

**Project Structure**

Healthcare_data_Automation/
│
├── main.py              ← The script (run this)
├── email_sender.py      ← Sends the report
├── excel_handler.py     ← Handles open Excel files
├── formatter.py         ← Makes the report look professional
├── requirement.txt 
Dashborad.py← shows live data 
├── README.md            ← This file
│
├── Clients_data/        ←  PUT YOUR FILES HERE
│   ├── Clinic_A_June_2024.xlsx
│   ├── Clinic_B_June_2024.xlsx
│   └── Clinic_C_June_2024.csv
│
└── output/              ← YOUR REPORT
    └── Healthcare_Data_Automation.xlsx
**Technology**
Python,Pandas,Openpyxl,pywin32,smtplib,python-dotenv
**About**
Process unlimited Excel and CSV files automatically: clean data, standardize columns, merge files, and generate professional reports with automatic email delivery.

**License**
**MIT** — free to use, modify, and share.
