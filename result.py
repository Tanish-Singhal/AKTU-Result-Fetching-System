from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

df = pd.read_excel(r'C:\\Users\\Tanish Singhal\\Desktop\\AKTU result Mini Project.xlsx')

driver = webdriver.Chrome()
driver.get('https://erp.aktu.ac.in/webpages/oneview/oneview.aspx')

wait = WebDriverWait(driver, 10)
roll_number_input = wait.until(EC.presence_of_element_located((By.ID, 'txtRollNo')))
time.sleep(5)

roll_number_input.send_keys(str(df['rollno'].iloc[0]))

submit_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnProceed')))
submit_button.click()

dob_input = wait.until(EC.presence_of_element_located((By.ID, 'dob_input_id')))
time.sleep(2)

# # Example: Fill the 'Date of Birth' and click on submit
# dob_input.send_keys(str(df['Date of Birth'].iloc[0]))  # Assuming 'Date of Birth' is a column in your Excel file

submit_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnProceed')))
submit_button.click()

driver.quit()
