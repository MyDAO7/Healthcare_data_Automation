import os
import pandas as pd
import numpy as np 
import openpyxl
import glob
import logging
# from pathlib import path
from excel_handler import process_excel_safely
from formatter import ExcelFormatter
from email_sender import send_email_report

#############~logging~#########

logger=logging.getLogger(__name__)
logger.setLevel("DEBUG")
consolehandler=logging.StreamHandler()
fileHandler=logging.FileHandler("logs.log",mode='w',encoding='utf-8')
formatter=logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
fileHandler.setFormatter(formatter)                         
fileHandler.setFormatter(formatter)
logger.addHandler(consolehandler)                         
logger.addHandler(fileHandler)     

#######~File reader~########

def file_reader(filepath):
     try:
        ext = os.path.splitext(filepath)[1].lower()
        if ext in ['.xlsx','.xls']:
            df=process_excel_safely(filepath,action='read')
            logger.info(f"Read Excel: {os.path.basename(filepath)} - ({len(df)} rows)")
        elif  ext=='.csv':
            df=pd.read_csv(filepath,encoding='utf-8')
            logger.info(f"Read CSV: {os.path.basename(filepath)} - ({len(df)} rows)")
        else:
             raise ValueError("unsupported file type")
        return df
     except Exception as e:
         logger.error(f" Failed to read {filepath}: {e}")
         return None
     
#######~File cleaner~########

def file_cleaner(df, filepath):
    if df is None or df.empty:
        logger.warning("Empty DataFrame passed to cleaner")
        return None

    rename_map = {
        'Name': 'Patient_Name',
        'Patient Name': 'Patient_Name',
        'Patient': 'Patient_Name',
        'Total': 'Amount',
        'Type': 'Service_Type',
        'Service': 'Service_Type',
        'Paid?': 'Status',
    }
    df = df.rename(columns=rename_map)

    # Create missing columns
    if 'Date' not in df.columns:
        month = os.path.basename(filepath).split('_')[1]
        df['Date'] = pd.Timestamp(f'2024-{month}-1')
        logger.debug("Created missing 'Date' column")

    if 'Patient_Name' not in df.columns:
        df['Patient_Name'] = 'Unknown'
        logger.debug("Created missing 'Patient_Name' column")

    if 'Doctor' not in df.columns:
        df['Doctor'] = 'Unknown'
        logger.debug("Created missing 'Doctor' column")

    if 'Amount' not in df.columns:
        df['Amount'] = 0.0
        logger.debug("Created missing 'Amount' column")

    if 'Service_Type' not in df.columns:
        df['Service_Type'] = 'General'
        logger.debug("Created missing 'Service_Type' column")

    if 'Status' not in df.columns:
        df['Status'] = 'Pending'
        logger.debug("Created missing 'Status' column")

    if 'Location' not in df.columns:
        location = os.path.basename(filepath).split('_')[1]
        df['Location'] = location
        logger.debug(f"Created 'Location' column: {location}")

    # Fill missing values
    df['Amount'] = df['Amount'].fillna(df['Amount'].mean())
    df['Patient_Name'] = df['Patient_Name'].fillna('Unknown')
    df['Status'] = df['Status'].fillna('Pending')
    df['Doctor'] = df['Doctor'].fillna('Unknown')
    df['Service_Type'] = df['Service_Type'].fillna('General')

    if 'Date' in df.columns:
   
         df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    
         df['Date'] = df['Date'].fillna(pd.Timestamp('2024-01-01'))
    
         df['Date'] = df['Date'].dt.strftime('%Y-%m-%d')
    
         df['Date'] = pd.to_datetime(df['Date'])

    return df

########~Summary Creation~#########

def summary(df):
    if df is None or len(df) == 0:
        logger.warning("Empty DataFrame passed to summary")
        return pd.DataFrame()
    
    summary = df.groupby('Location')['Amount'].agg(['sum', 'count', 'mean']).reset_index()
    
    summary.columns = ['Location', 'Total_Amount', 'Total_Patient', 'Avg_Sale']
    
    summary['Total_Amount'] = summary['Total_Amount'].round(2)
    summary['Avg_Sale'] = summary['Avg_Sale'].round(2)
    
    logger.info(f" Summary created: {len(summary)} locations")
    return summary

##########~E-mail sender~############

def get_email_stats(master_df, summary_df):
    """Generate statistics for email body."""
    stats = {}
    
    if master_df is not None and not master_df.empty:
        stats['total_patients'] = len(master_df)
        stats['total_visits'] = len(master_df)
        if 'Amount' in master_df.columns:
            stats['total_revenue'] = master_df['Amount'].sum()
    
    if summary_df is not None and not summary_df.empty:
        stats['locations'] = len(summary_df)
    else:
        stats['locations'] = 0
    
    return stats

############## (main) #############

files=[]
filepath=glob.glob("Clients_data/*.xlsx")
files.extend(filepath)
filepath=glob.glob("Clients_data/*.xls")
files.extend(filepath)
filepath=glob.glob("Clients_data/*.csv")
files.extend(filepath)
print(files)
all_data=[]
for file in files:
    logger.info(f"Processing ... {os.path.basename(file)}")
    df=file_reader(file)
    if df is not None:
        cleaned_df=file_cleaner(df,file)
        if cleaned_df is not None:           # ← NOW it's defined
            all_data.append(cleaned_df)
            logger.info(f"  {len(cleaned_df)} rows cleaned")
        else:
            logger.warning(f"  Cleaning returned None for {os.path.basename(file)}")
    else:
        logger.warning(f"  Skipping {os.path.basename(file)} - read failed")
if all_data:
    master_df = pd.concat(all_data, ignore_index=True)
    logger.info(f" Master dataset: {len(master_df)} rows from {len(all_data)} files")
    logger.info(f" Columns: {list(master_df.columns)}")
    if 'Date' in master_df.columns:
        master_df['Date'] = pd.to_datetime(master_df['Date'], errors='coerce')
    print(master_df)
    summary = summary(master_df)
    


output_file='Healthcare_data_Automaion.xlsx'
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    if not master_df.empty:
        master_df.to_excel(writer, sheet_name='All_data', index=False)
    if not summary.empty:
        summary.to_excel(writer, sheet_name='Monthly_Summary', index=False)

try:
    formatter = ExcelFormatter(output_file, header_color='366092', font_color='FFFFFF')
    formatter.format_sheet('All_data', number_format_cols=[3, 4])
    formatter.format_sheet('Monthly_Summary', number_format_cols=[2, 4])

    if not summary.empty:
        formatter.add_chart(
            data=summary,
            sheet_name='Monthly_Summary',
            chart_type='line',
            position='F2'
        )
        logger.info(" Formatting applied successfully")
except Exception as e:
    logger.error(f" Formatting failed: {e}")


if not master_df.empty:
    email_stats = get_email_stats(master_df, summary)
    
    # Get recipient from .env (or pass explicitly)
    recipient = os.getenv('recipient')
    
    # Send the report
    send_email_report(output_file, recipient, email_stats)
    
    logger.info(" Email process completed")
else:
    logger.warning(" No data to send in email")

