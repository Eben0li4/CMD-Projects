import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# File paths
input_file = r"C:\Users\EbenOlivier\Desktop\input.xlsx"
output_file = r"C:\Users\EbenOlivier\Desktop\output.xlsx"

# Read part numbers from the spreadsheet
data = pd.read_excel(input_file)
part_numbers = data['Part Numbers']  # Assuming your column is named 'Part Numbers'

# Set up WebDriver with fullscreen mode
options = Options()
driver = webdriver.Chrome(options=options)
driver.maximize_window()  # Start browser in fullscreen mode

# List to store results
results = []

try:
    for part_number in part_numbers:
        print(f"Processing: {part_number}")
        # Start timer
        start_time = time.time()

        # Open the webpage
        driver.get("https://parts.cat.com/en/barloworldequipment")

        # Handle cookie popup
        try:
            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[2]/div/div[1]/div/div[2]/div/button[2]'))
            ).click()
            print("Cookie popup dismissed.")
        except Exception as e:
            print("No cookie popup found or unable to dismiss it:", e)

        # Wait for the shadow host to be present in the DOM
        wait = WebDriverWait(driver, 10)
        shadow_host = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, 'cat-search-form[data-testid="cat-header-search-bar"]')
        ))

        # Access the shadow root
        shadow_root = driver.execute_script('return arguments[0].shadowRoot', shadow_host)

        # Wait for the search input inside the shadow DOM
        search_input = WebDriverWait(driver, 5).until(
            lambda d: shadow_root.find_element(By.CSS_SELECTOR, 'input.cat-c-search-form__input')
        )

        # Type the part number into the search bar
        search_input.send_keys(part_number)

        # Wait for the suggestion to appear and click it
        suggestion = wait.until(EC.element_to_be_clickable(
            (By.XPATH, '/html[1]/body[1]/div[1]/div[4]/main[1]/div[1]/section[1]/section[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/cat-search-form[1]/cat-list[1]/cat-list-item[2]/span[1]/span[1]')
        ))
        driver.execute_script("arguments[0].scrollIntoView(true);", suggestion)
        suggestion.click()

        # Set zoom level to 25% after navigating to the price page
        driver.execute_script("document.body.style.zoom='25%'")
        print("Page zoom set to 25%.")

        # Try to retrieve the price on the page
        try:
            price_element = wait.until(
                EC.presence_of_element_located((By.XPATH, '/html[1]/body[1]/div[1]/div[5]/div[1]/section[1]/div[2]/div[2]/div[1]/p[1]'))
            )
            final_text = price_element.text
        except:
            try:
                replacement_message = wait.until(
                    EC.presence_of_element_located((By.XPATH, '/html[1]/body[1]/div[1]/div[5]/div[1]/section[1]/div[2]/div[2]/div[1]/div[2]/div[1]/p[1]'))
                )
                if "This part has been replaced by the item below" in replacement_message.text:
                    xpaths = [
                        '/html[1]/body[1]/div[1]/div[5]/div[1]/section[1]/div[2]/div[3]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/a[1]',
                        '/html[1]/body[1]/div[1]/div[5]/div[1]/section[1]/div[2]/div[4]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/a[1]'
                    ]
                    new_part_link = None
                    for xpath in xpaths:
                        try:
                            new_part_link = wait.until(
                                EC.presence_of_element_located((By.XPATH, xpath))
                            )
                            driver.execute_script("arguments[0].scrollIntoView(true);", new_part_link)
                            new_part_link.click()
                            break
                        except:
                            continue

                    if new_part_link:
                        # Set zoom level on the new page
                        driver.execute_script("document.body.style.zoom='25%'")
                        print("Page zoom set to 25% for replacement part.")
                        new_price_element = wait.until(
                            EC.presence_of_element_located((By.XPATH, '/html[1]/body[1]/div[1]/div[5]/div[1]/section[1]/div[2]/div[2]/div[1]/p[1]'))
                        )
                        final_text = new_price_element.text
                    else:
                        final_text = "Replacement part link not found"
                else:
                    final_text = "Replacement message found but no valid link"
            except Exception as e:
                print(f"Error: {e}")
                final_text = "No price or replacement information available"

        # Stop timer and record elapsed time
        end_time = time.time()
        elapsed_time = end_time - start_time

        results.append({
            'Part Number': part_number,
            'Result': final_text,
            'Time Taken (seconds)': round(elapsed_time, 2)
        })

        print(f"Result for {part_number}: {final_text}")

except Exception as outer_exception:
    print("An error occurred during processing:", outer_exception)

finally:
    driver.quit()
    # Save results to the output file even if an error occurred
    output_df = pd.DataFrame(results)
    output_df.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")
