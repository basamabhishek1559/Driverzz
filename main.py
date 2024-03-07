import time

import ttime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl

# Provide your credentials and URL
email = "abhishek@nextdrive.co.in"
password = "Today12#"
url = "https://app.driverzz.com/admin/login"
email_xpath = "/html/body/div[2]/form/div[1]/div/input"
password_xpath = "/html/body/div[2]/form/div[2]/div/input"
submit_button_xpath = "/html/body/div[2]/form/div[3]/div/input"
search_input_xpath = "/html/body/nav/nav/div[1]/div[2]/div/input"
search_button_xpath = "/html/body/nav/nav/div[1]/div[2]/div/button"
name_xpath = "/html/body/div[2]/div[4]/div/div[2]/div[1]/div[2]/table/tbody/tr[2]/td[2]"
mobile_xpath = "/html/body/div[2]/div[4]/div/div[2]/div[1]/div[2]/table/tbody/tr[6]/td[2]"
email_id_xpath = "/html/body/div[2]/div[4]/div/div[2]/div[1]/div[2]/table/tbody/tr[4]/td[2]"
status_xpath = "/html/body/div[2]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[8]/td[2]/input"
original_cost_xpath = "/html/body/div[2]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[32]/td[2]"
cgst_xpath = "/html/body/div[2]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[34]/td[2]"
sgst_xpath = "/html/body/div[2]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[35]/td[2]"
actual_cost_xpath = "/html/body/div[2]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[37]/td[2]"
from_address_xpath = "/html/body/div[2]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[10]/td[2]"
trip_type_xpath = "/html/body/div[2]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[18]/td[2]"
driver_phone_number_xpath = "/html/body/div[2]/div[4]/div/div[2]/div[2]/div[2]/table/tbody/tr[4]/td[2]"
driver_name_xpath = "/html/body/div[2]/div[4]/div/div[2]/div[2]/div[2]/table/tbody/tr[2]/td[2]"
booking_date_xpath = "/html/body/div[2]/div[4]/div/div[3]/div/div[2]/table/tbody/tr[1]/td[4]"
customer_request_start_time_xpath = "/html/body/div[2]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[26]/td[2]"
customer_request_end_time_xpath = "/html/body/div[2]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[27]/td[2]"
driver_assigned_at_xpath = "/html/body/div[2]/div[4]/div/div[3]/div/div[2]/table/tbody/tr[3]/td[4]"
start_trip_xpath = "/html/body/div[2]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[28]/td[2]"
end_trip_xpath = "/html/body/div[2]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[29]/td[2]"
route_type_xpath = "/html/body/div[2]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[17]/td[2]"
ride_time_xpath = "/html/body/div[2]/div[4]/div/div[1]/div/div[2]/table/tbody/tr[31]/td[2]"

# Load Excel file
wb = openpyxl.load_workbook('order_sheet.xlsx')
sheet = wb.active

# Open Safari browser
driver = webdriver.Safari()
driver.maximize_window()

# Navigate to the URL
driver.get(url)

# Login
driver.find_element(By.XPATH, email_xpath).send_keys(email)
driver.find_element(By.XPATH, password_xpath).send_keys(password)
driver.find_element(By.XPATH, submit_button_xpath).click()

time.sleep(10)

# Wait for page to load
WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, search_input_xpath)))

# Function to perform search and update Excel
def perform_search_update_excel(data):
    driver.find_element(By.XPATH, search_input_xpath).clear()
    driver.find_element(By.XPATH, search_input_xpath).send_keys(data)
    driver.find_element(By.XPATH, search_button_xpath).click()
    time.sleep(120)  # Adjust as needed
    name = driver.find_element(By.XPATH, name_xpath).text
    mobile = driver.find_element(By.XPATH, mobile_xpath).text
    email_id = driver.find_element(By.XPATH, email_id_xpath).text
    status = driver.find_element(By.XPATH, status_xpath).get_attribute('value')
    original_cost = driver.find_element(By.XPATH, original_cost_xpath).text
    cgst = driver.find_element(By.XPATH, cgst_xpath).text
    sgst = driver.find_element(By.XPATH, sgst_xpath).text
    actual_cost = driver.find_element(By.XPATH, actual_cost_xpath).text
    from_address = driver.find_element(By.XPATH, from_address_xpath).text
    trip_type = driver.find_element(By.XPATH, trip_type_xpath).text
    driver_phone_number = driver.find_element(By.XPATH, driver_phone_number_xpath).text
    driver_name = driver.find_element(By.XPATH, driver_name_xpath).text
    booking_date = driver.find_element(By.XPATH, booking_date_xpath).text
    customer_request_start_time = driver.find_element(By.XPATH, customer_request_start_time_xpath).text
    customer_request_end_time = driver.find_element(By.XPATH, customer_request_end_time_xpath).text
    driver_assigned_at = driver.find_element(By.XPATH, driver_assigned_at_xpath).text
    start_trip = driver.find_element(By.XPATH, start_trip_xpath).text
    end_trip = driver.find_element(By.XPATH, end_trip_xpath).text
    route_type = driver.find_element(By.XPATH, route_type_xpath).text
    ride_time = driver.find_element(By.XPATH, ride_time_xpath).text

    return name, mobile, email_id, status, original_cost, cgst, sgst, actual_cost, from_address, trip_type, driver_phone_number, driver_name, booking_date, customer_request_start_time, customer_request_end_time, driver_assigned_at, start_trip, end_trip, route_type, ride_time

# Iterate through Excel data
for row in range(2, sheet.max_row + 1):
    data = sheet[f'AJ{row}'].value
    if data:
        name, mobile, email_id, status, original_cost, cgst, sgst, actual_cost, from_address, trip_type, driver_phone_number, driver_name, booking_date, customer_request_start_time, customer_request_end_time, driver_assigned_at, start_trip, end_trip, route_type, ride_time = perform_search_update_excel(data)
        sheet[f'A{row}'] = name
        sheet[f'B{row}'] = mobile
        sheet[f'C{row}'] = email_id
        sheet[f'H{row}'] = status
        sheet[f'L{row}'] = original_cost
        sheet[f'O{row}'] = cgst
        sheet[f'P{row}'] = sgst
        sheet[f'Q{row}'] = actual_cost
        sheet[f'S{row}'] = from_address
        sheet[f'T{row}'] = trip_type
        sheet[f'U{row}'] = driver_phone_number
        sheet[f'W{row}'] = driver_name
        sheet[f'X{row}'] = booking_date
        sheet[f'Y{row}'] = customer_request_start_time
        sheet[f'Z{row}'] = customer_request_end_time
        sheet[f'AA{row}'] = driver_assigned_at
        sheet[f'AB{row}'] = start_trip
        sheet[f'AC{row}'] = end_trip
        sheet[f'AK{row}'] = route_type
        sheet[f'AL{row}'] = ride_time
        wb.save('order_sheet.xlsx')

# Close the browser
driver.quit()

print("Task completed successfully.")

