import os
import time
import pandas as pd
import win32com.client
from win32com.client import constants
import pythoncom

class SafeExcelHandler:
    """
    Handles Excel files safely — even if they're open.
    Works with both .xls and .xlsx.
    """

    def __init__(self, filepath, visible=False):
        """
        Initialize the handler.
        - filepath: Full path to Excel file
        - visible: Show Excel window? (False = hidden, True = visible)
        """
        self.filepath = os.path.abspath(filepath)
        self.visible = visible
        self.excel = None
        self.workbook = None

    def _connect_excel(self):
        """Connect to a running Excel instance or start a new one."""
        try:
            # Try to get an already running Excel
            self.excel = win32com.client.GetActiveObject("Excel.Application")
        except:
            # If no Excel is running, start a new one
            self.excel = win32com.client.Dispatch("Excel.Application")

        self.excel.Visible = self.visible
        return self.excel

    def _get_or_open_workbook(self):
        """Get the workbook if open, or open it if not."""
        if not self.excel:
            self._connect_excel()

        # Check if workbook is already open
        for wb in self.excel.Workbooks:
            if wb.FullName.lower() == self.filepath.lower():
                self.workbook = wb
                return True

        # If not open, open it
        try:
            self.workbook = self.excel.Workbooks.Open(self.filepath)
            return True
        except Exception as e:
            raise Exception(f"Could not open workbook: {e}")

    def read_data(self, sheet_name=None):
        """
        Read data from Excel.
        - sheet_name: Optional sheet name. If None, uses active sheet.
        Returns pandas DataFrame.
        """
        if not self.workbook:
            self._get_or_open_workbook()

        # Get the sheet
        if sheet_name:
            sheet = self.workbook.Sheets(sheet_name)
        else:
            sheet = self.workbook.ActiveSheet

        # Get used range
        used_range = sheet.UsedRange
        data = used_range.Value

        # Convert to DataFrame
        if data and len(data) > 0:
            # First row as headers
            df = pd.DataFrame(data[1:], columns=data[0])
            return df
        else:
            return pd.DataFrame()

    def write_data(self, df, sheet_name=None):
        """
        Write DataFrame to Excel.
        - df: pandas DataFrame
        - sheet_name: Optional sheet name.
        """
        if not self.workbook:
            self._get_or_open_workbook()

        # Get the sheet
        if sheet_name:
            sheet = self.workbook.Sheets(sheet_name)
        else:
            sheet = self.workbook.ActiveSheet

        # Clear existing content
        sheet.Cells.ClearContents()

        # Write headers
        for col_idx, header in enumerate(df.columns, 1):
            sheet.Cells(1, col_idx).Value = str(header)

        # Write data
        for row_idx, row in enumerate(df.values, 2):
            for col_idx, value in enumerate(row, 1):
                sheet.Cells(row_idx, col_idx).Value = value

        # Auto-fit columns
        sheet.Columns.AutoFit()

    def save(self):
        """Save and close the workbook."""
        if self.workbook:
            self.workbook.Save()
            self.workbook.Close()
            self.workbook = None

    def close(self):
        """Close workbook without saving."""
        if self.workbook:
            self.workbook.Close(SaveChanges=False)
            self.workbook = None

    def quit_excel(self):
        """Quit Excel completely."""
        if self.excel:
            self.excel.Quit()
            self.excel = None


# ==================== THE CONVENIENCE FUNCTION ====================

def process_excel_safely(filepath, df=None, sheet_name=None, action='read'):
    """
    One function to handle Excel safely.
    Works whether file is open or not.

    Parameters:
    - filepath: Path to Excel file
    - df: DataFrame to write (if action='write')
    - sheet_name: Optional sheet name
    - action: 'read' or 'write' or 'update'

    Returns:
    - DataFrame (if action='read')
    - True (if action='write' or 'update')
    """
    handler = SafeExcelHandler(filepath, visible=False)

    try:
        handler._get_or_open_workbook()

        if action == 'read':
            result = handler.read_data(sheet_name)
            handler.save()
            return result

        elif action == 'write' or action == 'update':
            if df is None:
                raise ValueError("df required for write/update action")
            handler.write_data(df, sheet_name)
            handler.save()
            return True

        else:
            raise ValueError("action must be 'read', 'write', or 'update'")

    except Exception as e:
        raise Exception(f"Excel processing failed: {e}")
    finally:
        # Clean up
        handler.close()
        handler.quit_excel()


# ==================== HOW TO USE ====================

if __name__ == "__main__":
    # --- READING (works even if file is open) ---
    filepath = "client_data.xlsx"
    df = process_excel_safely(filepath, action='read')
    print("Read data:")
    print(df.head())

    # --- WRITING (works even if file is open) ---
    new_df = pd.DataFrame({
        'Name': ['Alice', 'Bob'],
        'Score': [95, 87]
    })
    process_excel_safely(filepath, df=new_df, action='write')
    print("Data written successfully.")

    # --- UPDATE (read, process, write) ---
    df = process_excel_safely(filepath, action='read')
    df['Processed'] = 'Done'
    process_excel_safely(filepath, df=df, action='write')
    print("Data updated successfully.")