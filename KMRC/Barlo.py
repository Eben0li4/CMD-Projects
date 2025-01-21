import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#
start_time = time.time()

#
driver = webdriver.Chrome()

try:
    #
    driver.get("https://parts.cat.com/en/barloworldequipment")

    #
    wait = WebDriverWait(driver, 10)
    shadow_host = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'cat-search-form[data-testid="cat-header-search-bar"]')))

    #
    shadow_root = driver.execute_script('return arguments[0].shadowRoot', shadow_host)

    #
    search_input = WebDriverWait(driver, 10).until(
        lambda d: shadow_root.find_element(By.CSS_SELECTOR, 'input.cat-c-search-form__input')
    )

    #
    search_input.send_keys("2336072")

    #
    suggestion = wait.until(EC.element_to_be_clickable((By.XPATH, '/html[1]/body[1]/div[1]/div[4]/main[1]/div[1]/section[1]/section[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/cat-search-form[1]/cat-list[1]/cat-list-item[2]/span[1]/span[1]')))

    #
    suggestion.click()

    #
    final_text = wait.until(
        EC.presence_of_element_located((By.XPATH, '/html[1]/body[1]/div[1]/div[5]/div[1]/div[3]/section[4]/div[1]/div[1]/div[3]/h4[1]'))
    ).text

    #
    print("Final Text:", final_text)

finally:
    #
    end_time = time.time()
    elapsed_time = end_time - start_time

    #
    print(f"Script runtime: {elapsed_time:.2f} seconds")

    #
    driver.quit()
