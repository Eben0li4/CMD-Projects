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
chrome_path = r"C:\Users\Eben Olivier\Downloads\chromedriver-win64\chromedriver-win64\chromedriver.exe"
chrome_service = Service(chrome_path)
chrome_options = Options()
chrome_options.add_argument("--headless")  # Run in headless mode
chrome_options.add_argument("--disable-gpu")  # Disable GPU acceleration (recommended for headless)
chrome_options.add_argument("--no-sandbox")  # Disable sandboxing (recommended for headless)
chrome_options.add_argument("--window-size=1920x1080")  # Optional: Set a specific window size

driver = webdriver.Chrome(service=chrome_service, options=chrome_options)

try:
    # Open the webpage with username and password embedded in URL
    username = "SSRS"
    password = "Zeti@Reports123!"
    url_with_auth = f"http://{username}:{password}@102.221.36.221/Reports/manage/catalogitem/subscriptions/4AM/Data%20Dumps/Availability%20per%20Day/KMRC%20Availability%20Per%20Day"
    driver.get(url_with_auth)

    # Wait for the checkbox to be clickable
    checkbox_xpath = "/html[1]/body[1]/div[1]/section[2]/div[1]/section[1]/div[1]/section[2]/subscriptions[1]/div[1]/fieldset[1]/ng-transclude[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/div[2]/label[1]/span[1]"
    checkbox = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, checkbox_xpath)))

    # Check the checkbox if it's not already checked
    if not checkbox.is_selected():
        checkbox.click()

    # Wait for the checkbox to be selected with a timeout
    try:
        WebDriverWait(driver, 2).until(EC.element_located_to_be_selected((By.XPATH, checkbox_xpath)))
    except TimeoutException as e:
        logging.error("Timeout waiting for checkbox to be selected: %s", e)

    # Find and click the button
    button_xpath = "/html[1]/body[1]/div[1]/section[2]/div[1]/section[1]/div[1]/section[2]/subscriptions[1]/div[1]/fieldset[1]/ng-transclude[1]/div[1]/ul[1]/li[4]/fieldset[1]/ng-transclude[1]/a[1]/span[2]"
    button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
    button.click()

    # Log button click
    logging.info("Button clicked successfully")

except NoSuchElementException as e:
    logging.error("Element not found: %s", e)

except TimeoutException as e:
    logging.error("Timeout waiting for element: %s", e)

except Exception as e:
    logging.error("An error occurred: %s", e)

finally:
    # Close the browser window
    driver.quit()
