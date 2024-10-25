import win32com.client
import os

def refresh_excel_workbook(file_path):
    if not os.path.exists(file_path):
        print(f"Error: The file '{file_path}' does not exist.")
        return

    try:
        # Create an instance of the Excel application
        excel_app = win32com.client.Dispatch("Excel.Application")

        # Open the workbook
        workbook = excel_app.Workbooks.Open(file_path)

        # Refresh all connections and queries
        workbook.RefreshAll()

        # Wait until all refreshes are completed (only if supported)
        try:
            excel_app.CalculateUntilAsyncQueriesDone()
        except AttributeError:
            print("Warning: CalculateUntilAsyncQueriesDone() not supported. Skipping.")

        # Save and close the workbook
        workbook.Close(SaveChanges=True)

    except Exception as e:
        print(f"Error: {e}")

    finally:
        # Quit the Excel application
        excel_app.Quit()

if __name__ == "__main__":
    file_path = r"C:\Users\EbenOlivier\Desktop\Oct - Nov 2024 TAWANA CRUSHING REPORT.xlsm"
    refresh_excel_workbook(file_path)
