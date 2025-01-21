import logging
import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def init_webdriver():
    chrome_path = os.getenv('CHROME_DRIVER_PATH', 'C:\\Users\\EbenOlivier\\Downloads\\chromedriver-win64\\chromedriver.exe')
    chrome_service = Service(chrome_path)
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Run in headless mode
    chrome_options.add_argument("--no-sandbox")  # Disable sandboxing (recommended for headless)
    chrome_options.add_argument("--disable-dev-shm-usage")  # Overcome limited resource problems
    chrome_options.add_argument("--window-size=1920x1080")  # Set a specific window size
    return webdriver.Chrome(service=chrome_service, options=chrome_options)

def log_checkbox_state(driver, checkbox):
    checkbox_attributes = driver.execute_script(
        'var items = {}; for (index = 0; index < arguments[0].attributes.length; ++index) { items[arguments[0].attributes[index].name] = arguments[0].attributes[index].value }; return items;', 
        checkbox)
    logging.info("Checkbox attributes: %s", checkbox_attributes)

def process_page(driver, url, checkbox_xpath, button_xpath):
    try:
        username = "SSRS"
        password = "Zeti@Reports123!"
        url_with_auth = f"http://{username}:{password}@{url}"
        logging.info("Navigating to URL: %s", url_with_auth)
        driver.get(url_with_auth)

        checkbox = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, checkbox_xpath)))
        log_checkbox_state(driver, checkbox)
        if not checkbox.get_attribute("checked"):
            checkbox.click()
            time.sleep(2)  # Small delay to ensure the click is registered

        button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
        button.click()
        logging.info("Button clicked successfully")

    except NoSuchElementException as e:
        logging.error("Element not found: %s", e)
    except TimeoutException as e:
        logging.error("Timeout waiting for element: %s", e)
    except Exception as e:
        logging.error("An error occurred: %s", e)
        with open(f"error_page_{url.split('/')[-1]}.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)

def main():
    driver = init_webdriver()
    try:
        # Processing multiple pages
        process_page(
            driver,
            "102.221.36.221/Reports/manage/catalogitem/subscriptions/4AM/Data%20Dumps/Availability%20per%20Day/KMRC%20Availability%20Per%20Day",
            "/html/body/div/section[2]/div/section/div/section[2]/subscriptions/div/fieldset/ng-transclude/div[2]/div/table/tbody/tr/td/div/div[2]/label/span",
            "/html/body/div/section[2]/div/section/div/section[2]/subscriptions/div/fieldset/ng-transclude/div[1]/ul/li[4]/fieldset/ng-transclude/a/span[2]"
        )
        process_page(
            driver,
            "102.221.36.221/Reports/manage/catalogitem/subscriptions/4AM/Data%20Dumps/BD%20Dash/KMRC%20Breakdown%20Dash",
            "/html/body/div/section[2]/div/section/div/section[2]/subscriptions/div/fieldset/ng-transclude/div[2]/div/table/tbody/tr/td/div/div[2]/label/span",
            "/html/body/div/section[2]/div/section/div/section[2]/subscriptions/div/fieldset/ng-transclude/div[1]/ul/li[4]/fieldset/ng-transclude/a/span[2]"
        )
    finally:
        driver.quit()

if __name__ == '__main__':
    main()
