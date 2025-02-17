import win32com.client
import os

def refresh_excel_workbooks(file_paths):
    excel_app = win32com.client.Dispatch("Excel.Application")
    try:
        for file_path in file_paths:
            if not os.path.exists(file_path):
                print(f"Error: The file '{file_path}' does not exist.")
                continue

            try:
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
                print(f"Error while processing '{file_path}': {e}")

    finally:
        # Quit the Excel application
        excel_app.Quit()

if __name__ == "__main__":
    file_paths = [
        r"C:\Users\EbenOlivier\OneDrive - 4 Arrows Mining\Sebilo Shared Folder\Sebilo\021. Sebilo Onsite Work Folder\15. 4AC REPORTS\03. LOGBOOK\25\02 2025 TAWANA CRUSHING LOGBOOK.xlsx",
        r"C:\Users\EbenOlivier\Desktop\Jan - Feb 2024 TAWANA CRUSHING REPORT.xlsm"
    ]
    refresh_excel_workbooks(file_paths)
