import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# File paths
input_file = "C:\\Users\\EbenOlivier\\Music\\inputx.xlsx"
output_file = "C:\\Users\\EbenOlivier\\Music\\outputx.xlsx"

# Start timer
start_time = time.time()

# Read part numbers from the input Excel file
input_data = pd.read_excel(input_file)
part_numbers = input_data["Part Number"].tolist()

# Configure Chrome options
chrome_options = Options()
#chrome_options.add_argument("--headless")  # Enable headless mode
chrome_options.add_argument("--disable-gpu")  # Disable GPU for better performance
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

# Initialize the Chrome driver
driver = webdriver.Chrome(options=chrome_options)

# Prepare to store results
results = []

try:
    # Iterate over each part number
    for part_number in part_numbers:
        # Navigate to the website
        driver.get("https://parts.cat.com/en/barloworldequipment")

        # Wait for the popup and click it to close
        popup_button_xpath = "/html/body/div[3]/div[2]/div/div[1]/div/div[2]/div/button[2]"
        wait = WebDriverWait(driver, 6)
        try:
            popup_button = wait.until(EC.element_to_be_clickable((By.XPATH, popup_button_xpath)))
            popup_button.click()
        except TimeoutException:
            print("Popup not found or already closed.")

        # Find the search input inside the shadow DOM
        shadow_host = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'cat-search-form[data-testid="cat-header-search-bar"]')))
        shadow_root = driver.execute_script('return arguments[0].shadowRoot', shadow_host)
        search_input = shadow_root.find_element(By.CSS_SELECTOR, 'input.cat-c-search-form__input')

        # Enter the search term and submit
        search_input.send_keys(part_number)

        # Wait for the search suggestion to be clickable and click it
        try:
            suggestion = wait.until(EC.element_to_be_clickable((By.XPATH, '/html[1]/body[1]/div[1]/div[4]/main[1]/div[1]/section[1]/section[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/cat-search-form[1]/cat-list[1]/cat-list-item[2]/span[1]/span[1]')))
            suggestion.click()
        except TimeoutException:
            print(f"No suggestions found for part number {part_number}.")
            results.append({"Part Number": part_number, "Result": "No results found"})
            continue

        # After the next page loads, set zoom to 25%
        driver.execute_script("document.body.style.zoom='25%'")

        time.sleep(4)

        # Scroll to the bottom of the page
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        # Attempt to click the button in section[4] or section[5]
        button_clicked = False
        for section_index in [4, 5, 6]:
            try:
                button_xpath = f'/html[1]/body[1]/div[1]/div[5]/div[1]/div[3]/section[{section_index}]/div[1]/div[2]/button[1]'
                button = wait.until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
                button.click()
                button_clicked = True
                break  # Exit the loop as the button is clicked
            except TimeoutException:
                continue

        if not button_clicked:
            print("Button not found or not clickable in both section[4] and section[5], continuing without clicking.")

        # Scroll to the bottom of the page
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        # Wait a moment to ensure dynamic content has loaded
        time.sleep(2)

        # Extract H4 elements from section[4] or section[5]
        part_results = []
        for section_index in [4, 5]:  # Try section[4] first, then section[5]
            index = 1
            while True:
                try:
                    xpath = f'/html[1]/body[1]/div[1]/div[5]/div[1]/div[3]/section[{section_index}]/div[1]/div[1]/div[{index}]/h4[1]'
                    h4_text = wait.until(EC.visibility_of_element_located((By.XPATH, xpath))).text
                    # Replace spaces with commas
                    h4_text = h4_text.replace(" ", ",")
                    part_results.append(h4_text)
                    index += 1
                except TimeoutException:
                    # If no more elements are found, break the loop
                    break
            if part_results:  # Stop checking other sections if results are found
                break

        # Combine results and store them
        combined_results = ','.join(part_results) if part_results else "No results found"
        results.append({"Part Number": part_number, "Result": combined_results})

finally:
    # Convert results to DataFrame
    results_df = pd.DataFrame(results)

    # Save the DataFrame to Excel
    results_df.to_excel(output_file, index=False)

    # Calculate elapsed time and print
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Script runtime: {elapsed_time:.2f} seconds")

    # Close the browser
    driver.quit()
