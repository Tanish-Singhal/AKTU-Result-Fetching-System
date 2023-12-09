from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# TODO: Load the Excel file and the website
df = pd.read_excel(r'C:\\Users\\Tanish Singhal\\Desktop\\AKTU result Mini Project.xlsx')

driver = webdriver.Chrome()
initial_url = 'https://erp.aktu.ac.in/webpages/oneview/oneview.aspx'
driver.get(initial_url)

wait = WebDriverWait(driver, 5)

for index, row in df.iterrows():
    # TODO: Roll No part
    roll_number_input = wait.until(EC.presence_of_element_located((By.ID, 'txtRollNo')))
    roll_number_input.clear()
    roll_number_input.send_keys(str(row['rollno']))

    submit_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnProceed')))
    submit_button.click()

    # TODO: Date of Birth Part
    dob_input = wait.until(EC.presence_of_element_located((By.ID, 'txtDOB')))
    formatted_dob = row['dob'].strftime('%Y-%m-%d')
    dob_input.clear()
    dob_input.send_keys(formatted_dob)

    # TODO: Captcha Manually Solve
    captcha_input = input("Please solve the captcha manually and press Enter when done: ")

    submit_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnSearch')))
    submit_button.click()

    # TODO: Second page start
    # TODO: Main div data extracted
    second_page_elements = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//tbody/tr[2]/td/div')))

    # Create new columns for each div
    for i, div_element in enumerate(second_page_elements, start=1):
        div_data = div_element.text
        column_name = f'div_{i}'  # Create a unique column name for each div
        df.loc[index, column_name] = div_data
 
    # TODO: Save the data in the Excel file
    df.to_excel(r'C:\Users\Tanish Singhal\Desktop\AKTU result Mini Project.xlsx', index=False)

    driver.get(initial_url)

time.sleep(5)
driver.quit()