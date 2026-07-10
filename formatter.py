# excel_formatter.py
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
import matplotlib.pyplot as plt
import pandas as pd
import os


class ExcelFormatter:
    """
    Apply professional formatting to any Excel report.
    Customizable for different clients, colors, and chart types.
    """

    def __init__(self, filepath, header_color='366092', font_color='FFFFFF'):
        """
        Initialize the formatter.

        Parameters:
        - filepath: Path to Excel file to format
        - header_color: Hex color for headers (default: blue)
        - font_color: Hex color for header text (default: white)
        """
        self.filepath = filepath
        self.header_color = header_color
        self.font_color = font_color

    def _get_styles(self):
        """Return style objects (font, fill, border, alignment)."""
        header_font = Font(bold=True, color=self.font_color, size=11)
        header_fill = PatternFill(
            start_color=self.header_color,
            end_color=self.header_color,
            fill_type='solid'
        )
        header_align = Alignment(horizontal='center', vertical='center')
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        return {
            'font': header_font,
            'fill': header_fill,
            'align': header_align,
            'border': thin_border
        }

    def format_sheet(self, sheet_name, number_format_cols=None):
        """
        Apply formatting to a specific sheet.

        Parameters:
        - sheet_name: Name of sheet to format
        - number_format_cols: List of column indices to format as numbers
        """
        wb = load_workbook(self.filepath)

        if sheet_name not in wb.sheetnames:
            print(f" Sheet '{sheet_name}' not found")
            return False

        ws = wb[sheet_name]
        styles = self._get_styles()

        # Format headers
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = styles['font']
            cell.fill = styles['fill']
            cell.alignment = styles['align']

        # Add borders to all cells
        for row in ws.iter_rows(
            min_row=1,
            max_row=ws.max_row,
            min_col=1,
            max_col=ws.max_column
        ):
            for cell in row:
                cell.border = styles['border']

        # Auto-fit columns
        for col in range(1, ws.max_column + 1):
            max_length = 0
            col_letter = get_column_letter(col)
            for row in range(1, ws.max_row + 1):
                cell_value = ws.cell(row=row, column=col).value
                if cell_value:
                    max_length = max(max_length, len(str(cell_value)))
            ws.column_dimensions[col_letter].width = min(max_length + 2, 30)

        # Format number columns
        if number_format_cols:
            for row in range(2, ws.max_row + 1):
                for col in number_format_cols:
                    cell = ws.cell(row=row, column=col)
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'

        wb.save(self.filepath)
        print(f" Formatted: {sheet_name}")
        return True

    def add_chart(self, data, sheet_name, chart_type='line', position='F2'):
        """
        Add a chart to a sheet.

        Parameters:
        - data: DataFrame with columns for chart
        - sheet_name: Where to place the chart
        - chart_type: 'line', 'bar', or 'pie'
        - position: Cell where chart starts (e.g., 'F2')
        """
        if data is None or data.empty:
            print(" No data for chart")
            return

        # Create chart
        plt.figure(figsize=(8, 5))

        if chart_type == 'line':
            plt.plot(data.iloc[:, 0], data.iloc[:, 1],
                     marker='o', linewidth=2, color='#1f77b4')
            plt.title('Monthly Trend', fontsize=14, fontweight='bold')
            plt.xlabel(data.columns[0], fontsize=12)
            plt.ylabel(data.columns[1], fontsize=12)
            plt.grid(True, alpha=0.3)

        elif chart_type == 'bar':
            plt.bar(data.iloc[:, 0], data.iloc[:, 1],
                    color='#1f77b4', edgecolor='black')
            plt.title('Monthly Comparison', fontsize=14, fontweight='bold')
            plt.xlabel(data.columns[0], fontsize=12)
            plt.ylabel(data.columns[1], fontsize=12)
            plt.xticks(rotation=45)

        elif chart_type == 'pie':
            plt.pie(data.iloc[:, 1], labels=data.iloc[:, 0],
                    autopct='%1.1f%%', startangle=90)
            plt.title('Distribution', fontsize=14, fontweight='bold')

        plt.xticks(rotation=45)
        plt.tight_layout()

        # Save and add to Excel
        temp_chart = 'temp_chart.png'
        plt.savefig(temp_chart, dpi=100)
        plt.close()

        wb = load_workbook(self.filepath)

        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            img = Image(temp_chart)
            img.width = 500
            img.height = 350
            ws.add_image(img, position)
            wb.save(self.filepath)
            print(f" Chart added to: {sheet_name}")

        # Clean up
        if os.path.exists(temp_chart):
            os.remove(temp_chart)

    def format_report(self, sheets_config, chart_data=None, chart_sheet=None, chart_type='line'):
        """
        One function to format entire report.

        Parameters:
        - sheets_config: Dict {sheet_name: [number_format_cols]}
        - chart_data: DataFrame for chart
        - chart_sheet: Sheet name for chart placement
        - chart_type: 'line', 'bar', or 'pie'
        """
        for sheet_name, number_cols in sheets_config.items():
            self.format_sheet(sheet_name, number_cols)

        if chart_data is not None and chart_sheet:
            self.add_chart(chart_data, chart_sheet, chart_type)

        print("Report formatting complete!")


# ==================== CONVENIENCE FUNCTION ====================

def format_excel_report(filepath, sheets, header_color='366092', font_color='FFFFFF'):
    """
    Quick one-liner to format an Excel report.

    Parameters:
    - filepath: Path to Excel file
    - sheets: List of sheet names to format
    - header_color: Hex color for headers
    - font_color: Hex color for header text

    Returns: True if successful
    """
    formatter = ExcelFormatter(filepath, header_color, font_color)

    for sheet in sheets:
        formatter.format_sheet(sheet)

    return True


# ==================== HOW TO USE ====================

if __name__ == "__main__":
    # Example 1: Basic formatting
    format_excel_report(
        'Sales_report_Monthly.xlsx',
        sheets=['All_data', 'Monthly_Summary'],
        header_color='366092',
        font_color='FFFFFF'
    )

    # Example 2: With custom formatting
    formatter = ExcelFormatter('Sales_report_Monthly.xlsx', header_color='1a5276')
    formatter.format_sheet('All_data', number_format_cols=[3, 4])
    formatter.format_sheet('Monthly_Summary', number_format_cols=[2, 4])

    # Example 3: With chart
    summary = pd.read_excel('Sales_report_Monthly.xlsx', sheet_name='Monthly_Summary')
    formatter.add_chart(summary, 'Monthly_Summary', chart_type='line', position='F2')