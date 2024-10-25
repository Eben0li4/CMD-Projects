import win32com.client

def run_macro(excel_path, macro_name):
    # Create an instance of the Excel application
    excel_app = win32com.client.Dispatch('Excel.Application')
    excel_app.Visible = False  # Run in background

    try:
        # Open the Excel workbook
        workbook = excel_app.Workbooks.Open(excel_path)

        # Run the macro
        excel_app.Application.Run(f'{workbook.Name}!{macro_name}')

        # Optionally, save the workbook if the macro makes changes
        workbook.Save()

        # Close the workbook
        workbook.Close(SaveChanges=0)
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Quit the Excel application
        excel_app.Application.Quit()

    print("Macro executed successfully.")

# Path to your Excel file
excel_path = r"C:\Users\Eben Olivier\OneDrive - 4 Arrows Mining\Report Runner ii.xlsm"
# Name of the macro to run
macro_name = 'Modules.KMRC_Secondary_Report_Scheduler'

run_macro(excel_path, macro_name)
