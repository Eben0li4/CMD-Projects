import logging
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
chrome_options.add_argument("--disable-gpu")  # Disable GPU acceleration
chrome_options.add_argument("--no-sandbox")  # Disable sandboxing (recommended for headless)
chrome_options.add_argument("--window-size=1920x1080")  # Optional: Set a specific window size

driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

def log_element_attributes(element, element_description="Element"):
    # Log element attributes
    element_attributes = driver.execute_script(
        'var items = {}; for (index = 0; index < arguments[0].attributes.length; ++index) { items[arguments[0].attributes[index].name] = arguments[0].attributes[index].value }; return items;',
        element)
    logging.info("%s attributes: %s", element_description, element_attributes)

def process_page(url, checkbox_xpath, button_xpath):
    try:
        # Open the webpage with username and password embedded in URL
        username = "SSRS"
        password = "Zeti@Reports123!"
        url_with_auth = f"http://{username}:{password}@{url}"
        logging.info("Navigating to URL: %s", url_with_auth)
        driver.get(url_with_auth)

        # Wait for the checkbox to be clickable and log its state
        logging.info("Waiting for checkbox to be clickable")
        checkbox = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, checkbox_xpath)))
        log_element_attributes(checkbox, "Checkbox")

        # Click the checkbox if it's not already selected
        if not checkbox.is_selected():
            logging.info("Checkbox not selected, clicking to select")
            checkbox.click()
            time.sleep(2)  # Small delay to ensure the click is registered
            log_element_attributes(checkbox, "Checkbox after click")

        # Find and click the button
        logging.info("Waiting for button to be clickable")
        button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
        button.click()
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

# Process the Availability page
process_page(
    "102.221.36.221/Reports/manage/catalogitem/subscriptions/4AM/Data%20Dumps/Availability%20per%20Day/UMK%20Availability%20Per%20Day",
    "/html[1]/body[1]/div[1]/section[2]/div[1]/section[1]/div[1]/section[2]/subscriptions[1]/div[1]/fieldset[1]/ng-transclude[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/div[2]/label[1]/span[1]",
    "/html[1]/body[1]/div[1]/section[2]/div[1]/section[1]/div[1]/section[2]/subscriptions[1]/div[1]/fieldset[1]/ng-transclude[1]/div[1]/ul[1]/li[4]/fieldset[1]/ng-transclude[1]/a[1]/span[2]"
)

# Close the browser window
driver.quit()
