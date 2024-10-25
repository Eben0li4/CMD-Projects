from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException
import time
import os

driver_path = 'C:\\Users\\EbenOlivier\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe'  # Make sure to set the correct path to the WebDriver
download_dir = "C:\cmd"  # Save exported file here

# Chrome in background
options = webdriver.ChromeOptions()
options.add_argument("--headless")  # Chrome Headless mode
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)

service = Service(driver_path)
driver = webdriver.Chrome(service=service, options=options)

try:
    # Open the login page
    driver.get("https://fleetmatic.bellequipment.com/Login.aspx")
    time.sleep(3)  # Wait for page to load

    # Handle cookie popup
    try:
        cookie_consent_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[@class='cc-btn cc-accept-all cc-btn-primary']"))
        )
        cookie_consent_button.click()
    except TimeoutException:
        print("No cookie consent popup found")

    # Click on the first button
    first_button = driver.find_element(By.XPATH, "//input[@id='MainContent_btnSignIn']")
    first_button.click()
    time.sleep(1)  # Wait

    # Fill in the email
    email_input = driver.find_element(By.XPATH, "/html[1]/body[1]/div[4]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/form[1]/fieldset[1]/div[1]/div[1]/div[1]/input[1]")
    email_input.send_keys("byron@4am.co.za")
    
    # Click the login button
    login_button = driver.find_element(By.XPATH, "/html[1]/body[1]/div[4]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/form[1]/fieldset[1]/div[2]/div[1]/button[1]")
    try:
        login_button.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].scrollIntoView(true);", login_button)
        login_button.click()
    time.sleep(3)  # Wait

    # Fill in the password
    password_input = driver.find_element(By.XPATH, "/html[1]/body[1]/div[4]/div[1]/div[3]/div[2]/div[1]/div[1]/div[2]/form[1]/fieldset[1]/div[2]/div[1]/div[1]/input[1]")
    password_input.send_keys("Operator1!@")
    
    # Click the submit button
    submit_button = driver.find_element(By.XPATH, "/html[1]/body[1]/div[4]/div[1]/div[3]/div[2]/div[1]/div[1]/div[2]/form[1]/fieldset[1]/div[4]/div[1]/button[1]")
    try:
        submit_button.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
        submit_button.click()
    time.sleep(3)  # Wait

    # Click on the initial notice button again
    initial_notice_button = driver.find_element(By.XPATH, "/html[1]/body[1]/div[3]/form[1]/div[4]/div[1]/div[1]/div[1]/div[1]/div[1]/input[1]")
    initial_notice_button.click()
    time.sleep(1)  # Wait

    # Click on the specific navigation link
    nav_link = driver.find_element(By.XPATH, "/html[1]/body[1]/div[4]/div[3]/div[1]/ul[1]/li[7]/a[1]/span[1]")
    try:
        nav_link.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].scrollIntoView(true);", nav_link)
        nav_link.click()
    time.sleep(3)  # Wait for the page to load

    # Enter "UMK" into the textbox
    textbox = driver.find_element(By.XPATH, "/html[1]/body[1]/div[4]/form[1]/div[4]/div[1]/div[2]/div[3]/div[1]/input[1]")
    textbox.send_keys("UMK")
    time.sleep(1)  # Wait for the input to be processed

    # Click on the search result
    search_result = driver.find_element(By.XPATH, "/html[1]/body[1]/div[4]/form[1]/div[4]/div[1]/div[2]/div[6]/div[2]/div[2]/div[1]/div[1]/div[1]/div[1]/ul[1]/li[1]/div[1]/div[1]/a[1]")
    try:
        search_result.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].scrollIntoView(true);", search_result)
        search_result.click()
    time.sleep(3)  # Wait for the search result to be processed

    # Click on the specified button to initiate the download
    alert_button = driver.find_element(By.XPATH, "/html[1]/body[1]/div[4]/form[1]/div[4]/div[1]/div[5]/div[3]/input[1]")
    try:
        alert_button.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].scrollIntoView(true);", alert_button)
        alert_button.click()
    time.sleep(1)  # Wait

    # Click on the export button
    export_button = driver.find_element(By.XPATH, "/html[1]/body[1]/div[4]/form[1]/div[4]/div[1]/div[5]/div[3]/div[1]/ul[1]/li[2]/div[1]/div[1]/div[2]")
    try:
        export_button.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].scrollIntoView(true);", export_button)
        export_button.click()
    time.sleep(5)  # Wait for the download to complete

finally:
    # Close the browser
    driver.quit()
