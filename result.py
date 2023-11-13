from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import time

df = pd.read_excel(r'C:\\Users\\Tanish Singhal\\Desktop\\AKTU result Mini Project.xlsx')

driver = webdriver.Chrome()
driver.get('https://erp.aktu.ac.in/webpages/oneview/oneview.aspx')

wait = WebDriverWait(driver, 10)
roll_number_input = wait.until(EC.presence_of_element_located((By.ID, 'txtRollNo')))
time.sleep(3)

roll_number_input.send_keys(str(df['rollno'].iloc[0]))

submit_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnProceed')))
submit_button.click()

dob_input = wait.until(EC.presence_of_element_located((By.ID, 'txtDOB')))
time.sleep(4)

formatted_dob = df['dob'].iloc[0].strftime('%Y-%m-%d')
dob_input.send_keys(formatted_dob)

submit_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnProceed')))
submit_button.click()

captcha_checkbox = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'recaptcha-checkbox-border')))

specific_part_location = captcha_checkbox.location
specific_part_x = specific_part_location['x'] + 10
specific_part_y = specific_part_location['y'] + 10

action_chains = ActionChains(driver)
action_chains.move_to_element_with_offset(captcha_checkbox, specific_part_x, specific_part_y)
action_chains.click()
action_chains.perform()


driver.quit()
