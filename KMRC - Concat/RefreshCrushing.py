import win32com.client as win32
import os

# Close any open workbooks and quit Excel
def terminate_excel():
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.DisplayAlerts = False  # Suppress any Excel alerts
        excel.Quit()  # Quit Excel application
    except Exception as e:
        print(f"Error while closing Excel: {e}")

    # Ensure Excel is terminated from Task Manager
    os.system("taskkill /f /im excel.exe")

terminate_excel()
print("Excel terminated successfully.")

def refresh_excel_workbook(file_path):
    # Create an instance of the Excel application
    excel_app = win32.Dispatch("Excel.Application")

    # Open the workbook
    workbook = excel_app.Workbooks.Open(file_path)

    # Refresh all connections and queries
    workbook.RefreshAll()

    # Wait until all refreshes are completed
    excel_app.CalculateUntilAsyncQueriesDone()

    # Save and close the workbook
    workbook.Close(SaveChanges=True)

    # Quit the Excel application
    excel_app.Quit()

if __name__ == "__main__":
    # List of Excel files to refresh in order
    excel_files = [
        r"C:\Users\EbenOlivier\OneDrive - 4 Arrows Mining\Eben - Frik\Info\KMRC Engineering & Production Delays Compared.xlsx",
        r"C:\Users\EbenOlivier\OneDrive - 4 Arrows Mining\Eben - Frik\Info\Crushing Checkup KMRC.xlsx",
        r"C:\Users\EbenOlivier\Desktop\Jan 2025 KMR CRUSHING REPORT - Concat.xlsm"
    ]

    # Refresh each Excel workbook
    for file_path in excel_files:
        refresh_excel_workbook(file_path)
        print(f"Refreshed: {file_path}")
