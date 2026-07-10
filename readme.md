## Healthcare Data Automation

## Overview
Automates data processing for a healthcare clinic with 3 locations.

## Features
- Reads Excel and CSV files
- Standardizes columns
- Handles missing data
- Merges all data into one master file
- Generates monthly report
- Emails report automatically

## Setup
1. Install dependencies: `pip install -r requirements.txt`
2. Create `.env` with your credentials
3. Place data files in `Clients_data/`
4. Run: `python Clients_data/main.py`

## Files
- `main.py` — Core automation script
- `formatter.py` — Report formatting
- `email_sender.py` — Email automation
- `excel_handler.py` — Excel operations