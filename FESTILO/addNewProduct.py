import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Specify the path to the ChromeDriver
service_object = Service(r"D:\Web driver\chromedriver-win64\chromedriver.exe")
driver = webdriver.Chrome(service=service_object)
driver.implicitly_wait(10)  # Implicit wait for all elements

# Load the Excel file
data = pd.read_excel('product_data.xlsx')  # Ensure you specify the correct path

# Open the desired URL
driver.get("https://estilo-admin-dev.dreamonlinelimited.xyz/")

# Send keys to the username and password fields
driver.find_element(By.XPATH, "//input[@id='email']").send_keys("admin@estilo.com")
driver.find_element(By.XPATH, "//input[@id='password']").send_keys("password")

# Click the login button
driver.find_element(By.XPATH, "//button[@type='submit']").click()

# Navigate to the products page
driver.get("https://estilo-admin-dev.dreamonlinelimited.xyz/products")

# Iterate through each product in the Excel sheet
for index, row in data.iterrows():
    product_name = row['Product_Name']

    # Check if the product already exists by searching for it on the product listing page
    search_box = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, "//input[@type='search']"))
    )
    search_box.clear()
    search_box.send_keys(product_name)

    time.sleep(2)  # Wait for the search results to update (adjust as needed)

    # Check if any search results are displayed
    try:
        product_exists = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, f"//td[contains(text(), '{product_name}')]"))
        )
        if product_exists:
            print(f"Product '{product_name}' already exists. Skipping...")
            continue  # Skip to the next product in the loop

    except Exception as e:
        print(f"Product '{product_name}' does not exist. Proceeding to add...")

    # Click 'Add Product' button
    WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']"))
    ).click()

    # Fill product details using row data
    driver.find_element(By.XPATH, "//input[@placeholder='Ex - Shoes']").send_keys(product_name)

    # Wait and select Category
    category_element = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//span[@id='select2-category_id-container']"))
    )
    category_element.click()
    driver.find_element(By.XPATH, f"//li[contains(text(),'{row['Category']}')]").click()

    # Select Series
    series_dropdown = driver.find_element(By.ID, "series_id")
    select_series = Select(series_dropdown)
    select_series.select_by_visible_text(row['Series'])

    # Select Type and Subtype
    select_type = Select(driver.find_element(By.ID, "type_id"))
    select_type.select_by_visible_text(row['Type'])

    select_subtype = Select(driver.find_element(By.ID, "subtype_id"))
    select_subtype.select_by_visible_text(row['Subtype'])

    # Select the color checkbox
    try:
        color_text = row['Color_Value']
        color_checkbox = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located(
                (By.XPATH, f"//div[@class='col-6 col-md-4'][contains(.,'{color_text}')]/input[@type='checkbox']"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", color_checkbox)
        driver.execute_script("arguments[0].click();", color_checkbox)

    except Exception as e:
        print(f"Error while selecting the color checkbox for '{color_text}': {e}")

    # Set price
    driver.find_element(By.XPATH, "//input[@placeholder='Ex - 999.99']").send_keys(str(row['Price']))

    # Select radio button for "Is New Arrival"
    try:
        is_new_arrival = row['Is_New_Arrival']  # This should be either 'YES' or 'NO'
        if is_new_arrival == 'YES':
            new_arrival_radio = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@name='is_new_arrival'][@value='YES']"))
            )
            driver.execute_script("arguments[0].click();", new_arrival_radio)
        else:
            new_arrival_radio_no = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@name='is_new_arrival'][@value='NO']"))
            )
            driver.execute_script("arguments[0].click();", new_arrival_radio_no)

    except Exception as e:
        print(f"Error while selecting 'Is New Arrival': {e}")

    # Select radio button for "Is On Sale"
    try:
        is_on_sale = row['Is_On_Sale']  # This should be either 'YES' or 'NO'
        if is_on_sale == 'YES':
            on_sale_radio = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@name='is_on_sale'][@value='YES']"))
            )
            driver.execute_script("arguments[0].click();", on_sale_radio)
        else:
            on_sale_radio_no = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@name='is_on_sale'][@value='NO']"))
            )
            driver.execute_script("arguments[0].click();", on_sale_radio_no)

    except Exception as e:
        print(f"Error while selecting 'Is On Sale': {e}")

    # Enter description text
    driver.find_element(By.CLASS_NAME, "trumbowyg-editor").send_keys(row['Description'])

    # Wait for the submit button to be clickable
    submit_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and @class='btn btn-success']"))
    )

    # Scroll to the button and click it using JavaScript
    driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
    driver.execute_script("arguments[0].click();", submit_button)

    # Optionally, wait a moment to allow for processing
    time.sleep(2)

# Maximize window to see everything clearly (optional)
driver.maximize_window()

# Wait indefinitely for manual closing
input("Press Enter to close the browser...")

# Close the browser
driver.quit()
