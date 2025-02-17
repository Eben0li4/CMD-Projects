import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException

# File paths and chromedriver path (use raw strings for Windows paths)
INPUT_FILE = r"C:\Users\EbenOlivier\Music\INPUT AN.xlsx"
OUTPUT_FILE = r"C:\Users\EbenOlivier\Music\OUTPUT AN.xlsx"
CHROMEDRIVER_PATH = r"C:\Users\EbenOlivier\Downloads\chromedriver-win64\chromedriver-win64\chromedriver.exe"

# URL of the website
BASE_URL = "https://www.filtrationtransmission.co.za/"

def setup_driver():
    service = Service(executable_path=CHROMEDRIVER_PATH)
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def search_and_get_price(driver, part_number):
    price = None
    wait = WebDriverWait(driver, 10)
    
    try:
        # Navigate to homepage
        driver.get(BASE_URL)
        
        # Wait for and click the search button
        search_button = wait.until(EC.element_to_be_clickable((
            By.XPATH, "/html[1]/body[1]/div[3]/div[2]/header[1]/div[1]/div[2]/div[1]/button[1]/*[name()='svg'][1]"
        )))
        search_button.click()
        
        # Wait for the search input field to be clickable and click to focus
        input_field = wait.until(EC.element_to_be_clickable((
            By.XPATH, "/html[1]/body[1]/div[3]/div[1]/div[1]/div[1]/form[1]/div[1]/input[1]"
        )))
        input_field.click()
        input_field.clear()
        
        # Enter the part number
        input_field.send_keys(str(part_number))
        
        # Allow time for the suggestions to load
        time.sleep(2)
        
        try:
            # Wait for the first suggestion to be clickable
            suggestion = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((
                By.XPATH, "/html[1]/body[1]/div[3]/div[1]/div[1]/div[1]/form[1]/div[1]/div[1]/div[1]/ul[1]/li[1]/a[1]/div[2]/span[1]/span[1]"
            )))
            suggestion.click()
        except (TimeoutException, ElementClickInterceptedException):
            # No suggestion available; skip this part number
            print(f"No clickable suggestion found for part: {part_number}")
            return None
        
        # Wait for the price element to be present on the new page
        price_element = wait.until(EC.presence_of_element_located((
            By.XPATH, "/html[1]/body[1]/div[4]/main[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/dl[1]/div[2]/div[1]/dd[1]/span[1]"
        )))
        price = price_element.text
        print(f"Found price for part {part_number}: {price}")
    except Exception as e:
        print(f"Error processing part {part_number}: {e}")
    
    return price

def main():
    # Read the Excel file
    try:
        df_input = pd.read_excel(INPUT_FILE)
    except Exception as e:
        print(f"Error reading input Excel file: {e}")
        return

    # Check that the expected column exists
    if "Part Num" not in df_input.columns:
        print("The column 'Part Num' was not found in the input Excel file.")
        return

    # List to store results
    results = []

    # Setup the Chrome webdriver
    driver = setup_driver()
    
    try:
        # Iterate over each part number
        for idx, row in df_input.iterrows():
            part_number = row["Part Num"]
            print(f"Processing part number: {part_number}")
            price = search_and_get_price(driver, part_number)
            results.append({"Part Num": part_number, "Price": price})
            # Optional: wait a bit before processing the next part
            time.sleep(1)
    finally:
        # Quit the driver after processing
        driver.quit()

    # Save results to Excel
    df_output = pd.DataFrame(results)
    try:
        df_output.to_excel(OUTPUT_FILE, index=False)
        print(f"Results saved to {OUTPUT_FILE}")
    except Exception as e:
        print(f"Error saving output Excel file: {e}")

if __name__ == '__main__':
    main()
