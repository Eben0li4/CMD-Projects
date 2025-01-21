import os
import win32com.client as win32

def create_email():
    outlook = win32.Dispatch('Outlook.Application')
    mail_item = outlook.CreateItem(0)
    
    # Open Excel workbook and worksheet
    xl_app = win32.Dispatch("Excel.Application")
    wb = xl_app.Workbooks.Open(r"C:\Users\EbenOlivier\OneDrive - 4 Arrows Mining\Eben - Frik\Report Builds.xlsm")
    ws = wb.Worksheets("Dirk")
    
    try:
        # Read recipients and subject from Excel
        to_address = ws.Range("A2").Value
        cc_address = ws.Range("A3").Value
        
        # Compose email
        # mail_item.SentOnBehalfOfName = "reporting@concat.co.za"
        mail_item.To = to_address
        mail_item.CC = cc_address
        mail_item.Subject = "KMR Crushing Report"
        mail_item.BodyFormat = 2  # HTML format
        
        # Construct HTML body with image
        html_body = (
            "<p>Good day<br><br>"
            "Find attached the <strong>KMR Crushing Report</strong><br><br>"
            "Concat Systems<br><br>"
            '<img src="C:\Dropbox\Intellicode Production Reports\CONCAT logo.jpg">'
        )
        mail_item.HTMLBody = html_body
        
        # Attach the latest Excel file
        xl_file = r"C:\Users\EbenOlivier\Desktop\Jan 2025 KMR CRUSHING REPORT - Concat.xlsm"
        mail_item.Attachments.Add(xl_file)
        
        # Display the email
        mail_item.Display()
    finally:
        # Ensure workbook is closed and Excel application is quit
        wb.Close(SaveChanges=False)
        xl_app.Quit()

if __name__ == "__main__":
    create_email()
