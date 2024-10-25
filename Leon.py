from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time

# Path to your WebDriver (e.g., ChromeDriver)
webdriver_path = "C:\\Users\\EbenOlivier\\Downloads\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe"

# Start a new browser session
service = Service(webdriver_path)
driver = webdriver.Chrome(service=service)

# Step 1: Open the URL
driver.get("https://web.concat.co.za/Account/Account/LoginUser")

# Add sufficient wait times
wait = WebDriverWait(driver, 10)

# Step 2: Add username
username_xpath = "/html[1]/body[1]/div[2]/div[1]/main[1]/form[1]/div[1]/div[1]/div[1]/input[1]"
username_field = wait.until(EC.presence_of_element_located((By.XPATH, username_xpath)))
username_field.send_keys("leon@4am.co.za")

# Step 3: Add password
password_xpath = "/html[1]/body[1]/div[2]/div[1]/main[1]/form[1]/div[1]/div[1]/div[2]/input[1]"
password_field = wait.until(EC.presence_of_element_located((By.XPATH, password_xpath)))
password_field.send_keys("dc6b8504")

# Step 4: Click login button
login_button_xpath = "/html[1]/body[1]/div[2]/div[1]/main[1]/form[1]/div[1]/div[1]/div[8]/button[1]/span[1]"
login_button = wait.until(EC.element_to_be_clickable((By.XPATH, login_button_xpath)))
login_button.click()

# Step 5: Interact with dropdown
dropdown_xpath = "/html[1]/body[1]/div[2]/div[1]/main[1]/form[1]/div[1]/div[1]/div[3]/select[1]"
dropdown = wait.until(EC.presence_of_element_located((By.XPATH, dropdown_xpath)))

# Use JavaScript to scroll into view and click the dropdown
driver.execute_script("arguments[0].scrollIntoView(true);", dropdown)
time.sleep(1)  # Add a short delay to ensure smooth interaction
dropdown.click()
dropdown.send_keys("4")
dropdown.send_keys(Keys.ENTER)

# Step 6: Click another button
next_button_xpath = "/html[1]/body[1]/div[2]/div[1]/main[1]/form[1]/div[1]/div[1]/div[8]/button[1]/span[1]"
next_button = wait.until(EC.element_to_be_clickable((By.XPATH, next_button_xpath)))
next_button.click()

# Step 7: Click a link/button
first_button_xpath = "/html[1]/body[1]/div[2]/div[1]/main[1]/div[2]/div[1]/div[1]/div[1]/div[1]/a[1]/span[1]"
first_button = wait.until(EC.element_to_be_clickable((By.XPATH, first_button_xpath)))
first_button.click()
time.sleep(6)

# Step 8: Click another button
second_button_xpath = "/html[1]/body[1]/div[2]/div[1]/main[1]/div[1]/div[1]/div[1]/button[1]"
second_button = wait.until(EC.element_to_be_clickable((By.XPATH, second_button_xpath)))
second_button.click()

# Step 9: Enter text "CRC"
crc_input_xpath = "/html[1]/body[1]/div[2]/div[1]/main[1]/div[1]/div[3]/div[1]/div[1]/form[1]/fieldset[1]/div[1]/div[1]/div[1]/div[1]/div[1]/input[1]"
crc_input = wait.until(EC.presence_of_element_located((By.XPATH, crc_input_xpath)))
driver.execute_script("arguments[0].scrollIntoView(true);", crc_input)
time.sleep(8)
crc_input.click()
crc_input.send_keys("CRC")
crc_input.send_keys(Keys.ENTER)

# Step 10: Enter text "EMT"
emt_input_xpath = "/html[1]/body[1]/div[2]/div[1]/main[1]/div[1]/div[3]/div[1]/div[1]/form[1]/fieldset[1]/div[1]/div[2]/div[1]/div[1]/div[1]/input[1]"
emt_input = wait.until(EC.presence_of_element_located((By.XPATH, emt_input_xpath)))
driver.execute_script("arguments[0].scrollIntoView(true);", emt_input)
time.sleep(1)
emt_input.click()
emt_input.send_keys("EMT")
emt_input.send_keys(Keys.ENTER)

# Step 11: Type "1024432" and press Enter
last_item_xpath = "/html[1]/body[1]/div[2]/div[1]/main[1]/div[1]/div[3]/div[1]/div[1]/form[1]/fieldset[1]/div[1]/div[3]/div[1]/div[1]/div[1]/input[1]"
last_item_input = wait.until(EC.presence_of_element_located((By.XPATH, last_item_xpath)))

# Scroll into view and interact with the input
driver.execute_script("arguments[0].scrollIntoView(true);", last_item_input)
time.sleep(1)
last_item_input.click()
last_item_input.send_keys("1024490")
time.sleep(0.5)  # Wait for the dropdown to update
last_item_input.send_keys(Keys.ENTER)

# Step 12: Click final button
final_button_xpath = "/html[1]/body[1]/div[2]/div[1]/main[1]/div[1]/div[3]/div[1]/div[1]/form[1]/fieldset[1]/div[1]/div[4]/div[1]/a[1]/span[1]"
final_button = wait.until(EC.element_to_be_clickable((By.XPATH, final_button_xpath)))
driver.execute_script("arguments[0].scrollIntoView(true);", final_button)
time.sleep(1)
final_button.click()

# Step 13: Print text to console
text_xpath = "/html[1]/body[1]/div[2]/div[1]/main[1]/div[1]/div[4]/div[1]/div[1]/div[2]/div[2]/table[1]/tbody[1]/tr[2]/td[6]"
text_element = wait.until(EC.presence_of_element_located((By.XPATH, text_xpath)))
print(text_element.text)

# Close the browser session
driver.quit()
