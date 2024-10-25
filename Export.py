import re
import pandas as pd
from PyPDF2 import PdfReader

def extract_pdf_to_excel(pdf_path, output_path):
    # Read the PDF file
    reader = PdfReader(pdf_path)
    pdf_text = []

    # Extract text from each page
    for page in reader.pages:
        pdf_text.append(page.extract_text())

    # Combine text from all pages into a single string
    full_text = "\n".join(pdf_text)

    # Regular expression pattern to capture all columns
    pattern = re.compile(
        r"(\d{3,4})\s+(\w+)\s+(.+?)\s+(\d+EA|\d+PAC|\d+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)"
    )

    # Find all matches in the text
    matches = pattern.findall(full_text)

    # Create a DataFrame to store the extracted data
    columns = ["Item", "Part No.", "Description", "Qty", "Unit Price", "Disc Perc", "Discount Value", "Other Discount", "Ext Nett Price (ZAR)"]
    data = []

    for match in matches:
        # Strip whitespace from each element in the match
        data.append([element.strip() for element in match])

    df = pd.DataFrame(data, columns=columns)

    # Export to Excel
    df.to_excel(output_path, index=False)
    print(f"Data has been exported to {output_path}")

# Example usage
pdf_path = r"C:\Users\EbenOlivier\Downloads\214662141.pdf"  # Replace with the path to the PDF file
output_path = r"C:\Users\EbenOlivier\Documents\Yas.xlsx"  # Replace with the desired output path
extract_pdf_to_excel(pdf_path, output_path)
