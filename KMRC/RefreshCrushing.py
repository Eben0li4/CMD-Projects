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
    file_path = r"C:\Users\EbenOlivier\Desktop\Oct 2024 KMR CRUSHING REPORT.xlsm"
    refresh_excel_workbook(file_path)
