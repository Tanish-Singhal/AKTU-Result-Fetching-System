from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

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

    # FIXME: Second page start
    # Move to the targeted div
    second_page_elements = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//tbody/tr[2]/td/div')))
    
    # Extract the data from the last div and update it
    last_div_data = second_page_elements[-1].text
    df.loc[index, 'Result'] = last_div_data

    # save the data in the excel file
    df.to_excel(r'C:\\Users\\Tanish Singhal\\Desktop\\AKTU result Mini Project.xlsx', index=False)

    print(f"Data for Roll No {row['rollno']} from the last div in the second page:", last_div_data)

    # Run the main URL again, to fetch the data of the next student
    driver.get(initial_url)

driver.quit()