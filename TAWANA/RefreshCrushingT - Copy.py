import win32com.client
import os
from multiprocessing import Process

def refresh_excel_workbook(file_path):
    # Create an instance of the Excel application
    excel_app = win32com.client.Dispatch("Excel.Application")
    try:
        if not os.path.exists(file_path):
            print(f"Error: The file '{file_path}' does not exist.")
            return

        opened_workbooks = [wb.FullName.lower() for wb in excel_app.Workbooks]
        target_workbook_path = file_path.lower()

        # Only open the workbook if it's not already opened
        if target_workbook_path not in opened_workbooks:
            workbook = excel_app.Workbooks.Open(file_path)
        else:
            # Refer to the already opened workbook
            workbook = excel_app.Workbooks.Item(os.path.basename(file_path))

        # Refresh all connections and queries
        workbook.RefreshAll()

        # Wait until all refreshes are completed (only if supported)
        try:
            excel_app.CalculateUntilAsyncQueriesDone()
        except AttributeError:
            print("Warning: CalculateUntilAsyncQueriesDone() not supported. Skipping.")

        # Save and close the workbook only if it was opened by this script
        if target_workbook_path not in opened_workbooks:
            workbook.Close(SaveChanges=True)

    except Exception as e:
        print(f"Error while processing '{file_path}': {e}")

    finally:
        # Reduce Excel application interference
        if len(excel_app.Workbooks) == 0:
            excel_app.Quit()
        print(f"Workbook at {file_path} refreshed successfully.")

def run_process(file_path):
    refresh_excel_workbook(file_path)

if __name__ == "__main__":
    file_paths = [
        r"C:\Users\EbenOlivier\OneDrive - 4 Arrows Mining\Sebilo Shared Folder\Sebilo\021. Sebilo Onsite Work Folder\15. 4AC REPORTS\03. LOGBOOK\24\12 2024 TAWANA CRUSHING LOGBOOK.xlsx",
        r"C:\Users\EbenOlivier\Desktop\Nov - Dec 2024 TAWANA CRUSHING REPORT.xlsm"
        # Add more file paths as needed
    ]

    # Create and start a separate process for each file path
    processes = []
    for file_path in file_paths:
        p = Process(target=run_process, args=(file_path,))
        p.start()
        processes.append(p)

    # Wait for all processes to complete
    for process in processes:
        process.join()
