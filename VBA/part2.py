import openpyxl
import time

def kmrc_secondary_report_scheduler():
    # Load the workbook and select the active sheet
    file_path = r"C:\Users\EbenOlivier\OneDrive - 4 Arrows Mining\Reporting\Data\4AM Current Plantlist.xlsx"
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    
    # Convert the worksheet to a list of lists
    data = list(ws.values)
    
    # Filter the data: keep rows where the third column is not 'OFFSITE'
    filtered_data = [data[0]]  # Keep the header
    for row in data[1:]:
        if row[2] != "OFFSITE":
            filtered_data.append(row)
    
    # Clear the worksheet
    ws.delete_rows(1, ws.max_row)
    
    # Write the filtered data back to the worksheet
    for row in filtered_data:
        ws.append(row)
    
    # Save the workbook
    wb.save(file_path)
    
    # Wait for 3 seconds
    time.sleep(3)

# Call the function
kmrc_secondary_report_scheduler()
