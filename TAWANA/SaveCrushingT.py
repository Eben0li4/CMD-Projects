import os
import time
import datetime
import win32com.client
import pythoncom

def update_chart_data_to_latest_date(excel_path, calc_sheet_name, chart_sheet_name, chart_name, desired_plot_by):
    if not os.path.exists(excel_path):
        print(f"Error: The file '{excel_path}' does not exist.")
        return
    
    today = datetime.datetime.now()
    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False
    workbook = excel_app.Workbooks.Open(excel_path)
    
    try:
        calc_sheet = workbook.Sheets(calc_sheet_name)
        chart_sheet = workbook.Sheets(chart_sheet_name)

        pythoncom.PumpWaitingMessages()
        
        chart = None
        for c in chart_sheet.ChartObjects():
            if c.Name == chart_name:
                chart = c.Chart
                break

        if chart is not None:
            print(f"Updating chart '{chart_name}' in sheet '{chart_sheet_name}' to include data up to {today.strftime('%d-%m-%Y')}")
            start_row = 44
            if today.day >= 21:
                day_adjustment = today.day - 22
            else:
                day_adjustment = today.day + 9
            end_row = start_row + day_adjustment
            data_range = calc_sheet.Range(f"C43:I{end_row}")
            chart.SetSourceData(data_range)
            
            if chart.PlotBy != desired_plot_by:
                chart.PlotBy = desired_plot_by
                print(f"Chart plotting adjusted to {'rows' if desired_plot_by == 1 else 'columns'}.")

            print(f"Chart data range updated to: C43:I{end_row}")
        else:
            print(f"Chart '{chart_name}' not found in sheet '{chart_sheet_name}'")

        for attempt in range(5):
            try:
                workbook.Save()
                print("Workbook saved successfully.")
                break
            except Exception as e:
                print(f"Attempt {attempt + 1}: Error occurred while saving the workbook: {e}")
                time.sleep(2)
        else:
            print("Failed to save the workbook after several attempts.")

    except Exception as e:
        print(f"An error occurred while updating chart data: {e}")
    
    finally:
        for attempt in range(5):
            try:
                workbook.Close(False)
                print("Workbook closed successfully.")
                break
            except Exception as e:
                print(f"Attempt {attempt + 1}: Error occurred while closing the workbook: {e}")
                time.sleep(2)
        else:
            print("Failed to close the workbook after several attempts.")
        
        excel_app.Quit()

def save_excel_as_pdf(excel_path, sheet_name, pdf_folder):
    if not os.path.exists(excel_path):
        print(f"Error: The file '{excel_path}' does not exist.")
        return
    
    if not os.path.exists(pdf_folder):
        print(f"Error: The folder '{pdf_folder}' does not exist.")
        return
    
    today = datetime.datetime.now().strftime("%d-%m-%Y")
    new_pdf_name = f"{today} TAWANA CRUSHING REPORT.pdf"
    pdf_path = os.path.join(pdf_folder, new_pdf_name)
    
    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False
    
    try:
        workbook = excel_app.Workbooks.Open(excel_path)
        sheet = workbook.Sheets(sheet_name)
        sheet.Activate()
        
        pythoncom.PumpWaitingMessages()
        
        sheet.ExportAsFixedFormat(0, pdf_path)
        print(f"PDF saved to: {pdf_path}")

    except Exception as e:
        print(f"An error occurred while saving the PDF: {e}")

    finally:
        for attempt in range(5):
            try:
                workbook.Close(False)
                print("Workbook closed successfully.")
                break
            except Exception as e:
                print(f"Attempt {attempt + 1}: Error occurred while closing the workbook: {e}")
                time.sleep(2)
        else:
            print("Failed to close the workbook after several attempts.")
        
        excel_app.Quit()

def find_charts_in_excel(file_path, sheet_name, chart_name):
    if not os.path.exists(file_path):
        print(f"Error: The file '{file_path}' does not exist.")
        return False
    
    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False
    workbook = excel_app.Workbooks.Open(file_path)
    
    chart_found = False
    try:
        sheet = workbook.Sheets(sheet_name)
        for c in sheet.ChartObjects():
            if c.Name == chart_name:
                chart_found = True
                print(f"Chart found in sheet: {sheet_name} with name: '{chart_name}'")
                break

        if not chart_found:
            print(f"No chart named '{chart_name}' found in sheet: {sheet_name}")

    finally:
        for attempt in range(5):
            try:
                workbook.Close(False)
                print("Workbook closed successfully.")
                break
            except Exception as e:
                print(f"Attempt {attempt + 1}: Error occurred while closing the workbook: {e}")
                time.sleep(2)
        else:
            print("Failed to close the workbook after several attempts.")
        
        excel_app.Quit()
    
    return chart_found

# Define paths and variables
excel_path = r"C:\Users\EbenOlivier\Desktop\Jan - Feb 2024 TAWANA CRUSHING REPORT.xlsm"
pdf_folder = r"C:\Users\EbenOlivier\OneDrive - 4 Arrows Mining\Sebilo Eben\TAWANA DAILY\PDFS"
chart_sheet_name = 'TAWANA CRUS PROD REPORT'
calc_sheet_name = 'Calculations'
chart_name = 'Chart 1'
desired_plot_by = 2  # Assuming you want columns

if find_charts_in_excel(excel_path, chart_sheet_name, chart_name):
    update_chart_data_to_latest_date(excel_path, calc_sheet_name, chart_sheet_name, chart_name, desired_plot_by)
    save_excel_as_pdf(excel_path, chart_sheet_name, pdf_folder)
else:
    print("No chart found or updated. PDF will not be created.")
