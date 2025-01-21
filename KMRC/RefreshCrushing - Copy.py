import win32com.client as win32
from multiprocessing import Process

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
    print(f"Workbook at {file_path} refreshed successfully.")

def run_process(file_path):
    refresh_excel_workbook(file_path)

if __name__ == "__main__":
    # List of Excel files to be processed (example paths)
    file_paths = [
        r"C:\Users\EbenOlivier\Desktop\Dec 2024 KMR CRUSHING REPORT.xlsm",
        #r"C:\Users\EbenOlivier\Desktop\Jan 2025 KMR CRUSHING REPORT.xlsm",
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
