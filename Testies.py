import xlwings as xw
import time

def process_current_plantlist():
    # Start Excel (invisible)
    app = xw.App(visible=False)
    app.api.DisplayAlerts = False  # Disable alerts

    try:
        wb = xw.Book(r"C:\Users\EbenOlivier\OneDrive - 4 Arrows Mining\Reporting\Data2(Speed up OneDrive)\4AM Current Plantlist.xlsx")
        ws = wb.sheets.active

        # Remove any existing filter
        try:
            ws.api.ShowAllData()
        except Exception:
            pass

        # Apply filter on range A1:N5000, filtering column C (Field 3) for "OFFSITE"
        ws.range("A1:N5000").api.AutoFilter(Field=3, Criteria1="OFFSITE")

        # Delete visible cells (xlCellTypeVisible = 12) in A2:N5000
        try:
            ws.range("A2:N5000").api.SpecialCells(12).Delete()
        except Exception:
            # If no visible cells are found, ignore the error.
            pass

        # Remove filter
        try:
            ws.api.ShowAllData()
        except Exception:
            pass

        wb.save()
        wb.close()
    finally:
        app.api.DisplayAlerts = True
        time.sleep(3)  # Wait for 3 seconds
        app.quit()

if __name__ == "__main__":
    process_current_plantlist()
