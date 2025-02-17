import os
import time
import datetime
import win32com.client
import xlwings as xw
import shutil

def update_chart_data_to_latest_date(excel_path, calc_sheet_name, chart_sheet_name, chart_name):
    today = datetime.datetime.now()

    # Open the Excel file
    app = xw.App(visible=False)
    wb = xw.Book(excel_path)
    
    try:
        calc_sheet = wb.sheets[calc_sheet_name]
        chart_sheet = wb.sheets[chart_sheet_name]
        chart = None

        # Find the chart by name in the chart_sheet
        for c in chart_sheet.charts:
            if c.name == chart_name:
                chart = c
                break

        if chart is not None:
            if today.day == 1:
                # If today is the 1st of the month, use the full range of the previous month
                previous_month = today.replace(day=1) - datetime.timedelta(days=1)
                start_row = 5
                end_row = start_row + previous_month.day - 1
                print(f"Updating chart '{chart_name}' in sheet '{chart_sheet_name}' to include data for the full range of the previous month: {previous_month.strftime('%B %Y')}")
            else:
                # Define the starting row for the data (row 5, assuming headers are on row 4)
                start_row = 5
                end_row = start_row + today.day - 1
                print(f"Updating chart '{chart_name}' in sheet '{chart_sheet_name}' to include data up to {today.strftime('%d-%m-%Y')}")

            # Update the data range to include all columns from C to I, and rows from 43 to end_row
            data_range = f"C4:I{end_row}"
            chart.set_source_data(calc_sheet.range(data_range))
            print(f"Chart data range updated to: {data_range}")
        else:
            print(f"Chart '{chart_name}' not found in sheet '{chart_sheet_name}'")

    finally:
        # Add a delay before saving
        time.sleep(2)
        # Save and close the workbook
        try:
            wb.save()
            print("Workbook saved successfully.")
        except Exception as e:
            print(f"An error occurred while saving the workbook: {e}")
        finally:
            wb.close()
            app.quit()

def save_excel_as_pdf(excel_path, sheet_name, pdf_folder):
    today = datetime.datetime.now().strftime("%d-%m-%Y")
    new_pdf_name = f"{today} KMR CRUSHING REPORT.pdf"
    pdf_path = os.path.join(pdf_folder, new_pdf_name)
    
    # Convert Excel sheet to PDF using win32com.client
    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False
    
    try:
        workbook = excel_app.Workbooks.Open(excel_path)
        sheet = workbook.Sheets[sheet_name]
        
        # Export to PDF
        sheet.ExportAsFixedFormat(0, pdf_path)
        
        print(f"PDF saved to: {pdf_path}")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Clean up
        workbook.Close(False)
        excel_app.Quit()

def find_charts_in_excel(file_path, sheet_name, chart_name):
    # Open the Excel file
    app = xw.App(visible=False)
    wb = xw.Book(file_path)
    
    chart_found = False
    try:
        sheet = wb.sheets[sheet_name]
        for chart in sheet.charts:
            if chart.name == chart_name:
                chart_found = True
                print(f"Chart found in sheet: {sheet_name} with name: '{chart_name}'")
                break

        if not chart_found:
            print(f"No chart named '{chart_name}' found in sheet: {sheet_name}")

    finally:
        # Close the workbook without saving
        wb.close()
        app.quit()
    
    return chart_found

# Define paths
excel_path = r"C:\Users\EbenOlivier\Desktop\Feb 2025 KMR CRUSHING REPORT - Concat.xlsm"
pdf_folder = r"C:\Users\EbenOlivier\OneDrive - 4 Arrows Mining\Eben - Frik\Report PDF's\2025"
copy_folder = r"C:\Users\EbenOlivier\OneDrive - 4 Arrows Mining\Eben - Frik"
chart_sheet_name = 'KMR CRUS PROD REPORT'
calc_sheet_name = 'Calculations2'
chart_name = 'Chart 1'

# Check for the specific chart and update it if found
if find_charts_in_excel(excel_path, chart_sheet_name, chart_name):
    update_chart_data_to_latest_date(excel_path, calc_sheet_name, chart_sheet_name, chart_name)
    save_excel_as_pdf(excel_path, chart_sheet_name, pdf_folder)
else:
    print("No chart found or updated. PDF will not be created.")

shutil.copy(excel_path, copy_folder)
