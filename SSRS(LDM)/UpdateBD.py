import logging
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Initialize the webdriver
chrome_path = r"C:\Users\EbenOlivier\Downloads\chromedriver-win64\chromedriver-win64\chromedriver.exe"
chrome_service = Service(chrome_path)
chrome_options = Options()
chrome_options.add_argument("--headless")  # Run in headless mode
chrome_options.add_argument("--disable-gpu")  # Disable GPU acceleration (recommended for headless)
chrome_options.add_argument("--no-sandbox")  # Disable sandboxing (recommended for headless)
chrome_options.add_argument("--window-size=1920x1080")  # Optional: Set a specific window size

driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

def log_checkbox_state(checkbox):
    # Log checkbox attributes
    checkbox_attributes = driver.execute_script(
        'var items = {}; for (index = 0; index < arguments[0].attributes.length; ++index) { items[arguments[0].attributes[index].name] = arguments[0].attributes[index].value }; return items;',
        checkbox)
    logging.info("Checkbox attributes: %s", checkbox_attributes)

def process_page(url, checkbox_xpath, button_xpath):
    try:
        # Open the webpage with username and password embedded in URL
        username = "SSRS"
        password = "Zeti@Reports123!"
        url_with_auth = f"http://{username}:{password}@{url}"
        logging.info("Navigating to URL: %s", url_with_auth)
        driver.get(url_with_auth)

        # Wait for the checkbox to be clickable
        logging.info("Waiting for checkbox to be clickable")
        checkbox = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, checkbox_xpath)))
        logging.info("Checkbox found")

        # Log the initial state of the checkbox
        log_checkbox_state(checkbox)
        checkbox_selected_initial = checkbox.get_attribute("checked")
        logging.info("Initial checkbox state: %s", checkbox_selected_initial)

        # Take a screenshot before clicking the checkbox
        # driver.save_screenshot(f"before_click_{url.split('/')[-1]}.png")

        # Check the checkbox if it's not already checked
        if not checkbox_selected_initial:
            logging.info("Checkbox not selected, clicking to select")
            driver.execute_script("arguments[0].click();", checkbox)
            time.sleep(2)  # Small delay to ensure the click is registered

        # Take a screenshot after clicking the checkbox
        # driver.save_screenshot(f"after_click_{url.split('/')[-1]}.png")

        # Verify the checkbox state by checking the 'checked' attribute
        log_checkbox_state(checkbox)
        checkbox_selected_final = checkbox.get_attribute("checked")
        logging.info("Final checkbox state: %s", checkbox_selected_final)

        if checkbox_selected_final:
            logging.info("Checkbox selected successfully")
        else:
            logging.error("Checkbox was not selected successfully")

        # Find and click the button
        logging.info("Waiting for button to be clickable")
        button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
        logging.info("Button found, clicking button")
        button.click()

        # Log button click
        logging.info("Button clicked successfully")

    except NoSuchElementException as e:
        logging.error("Element not found: %s", e)

    except TimeoutException as e:
        logging.error("Timeout waiting for element: %s", e)

    except Exception as e:
        logging.error("An error occurred: %s", e)
        # Save the page source for debugging
        with open(f"page_source_{url.split('/')[-1]}.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)

# Process the second page
process_page(
    "102.221.36.221/Reports/manage/catalogitem/subscriptions/Matekane/Data%20Dumps/Breakdown%20Dash_LDM",
    "/html[1]/body[1]/div[1]/section[2]/div[1]/section[1]/div[1]/section[2]/subscriptions[1]/div[1]/fieldset[1]/ng-transclude[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/div[2]/label[1]/span[1]",
    "/html[1]/body[1]/div[1]/section[2]/div[1]/section[1]/div[1]/section[2]/subscriptions[1]/div[1]/fieldset[1]/ng-transclude[1]/div[1]/ul[1]/li[4]/fieldset[1]/ng-transclude[1]/a[1]/span[2]"
)

# Process the same page
process_page(
    "102.221.36.221/Reports/manage/catalogitem/subscriptions/Matekane/Data%20Dumps/Breakdown%20Dash_LDM",
    "/html[1]/body[1]/div[1]/section[2]/div[1]/section[1]/div[1]/section[2]/subscriptions[1]/div[1]/fieldset[1]/ng-transclude[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/div[2]/label[1]/span[1]",
    "/html[1]/body[1]/div[1]/section[2]/div[1]/section[1]/div[1]/section[2]/subscriptions[1]/div[1]/fieldset[1]/ng-transclude[1]/div[1]/ul[1]/li[4]/fieldset[1]/ng-transclude[1]/a[1]/span[2]"
)

# Close the browser window
driver.quit()
