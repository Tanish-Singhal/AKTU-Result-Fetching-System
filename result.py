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

# TODO: entering the roll no from the excel file
wait = WebDriverWait(driver, 10)
roll_number_input = wait.until(EC.presence_of_element_located((By.ID, 'txtRollNo')))
time.sleep(3)
roll_number_input.send_keys(str(df['rollno'].iloc[0]))

# TODO: click on proceed button
submit_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnProceed')))
submit_button.click()

# TODO: entering the date of birth from the excel file
dob_input = wait.until(EC.presence_of_element_located((By.ID, 'txtDOB')))
time.sleep(4)

formatted_dob = df['dob'].iloc[0].strftime('%Y-%m-%d')
dob_input.send_keys(formatted_dob)

# TODO: Captcha Part (ongoing)
# Wait for the reCAPTCHA iframe to be present
captcha_frame = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, ".//iframe[@title='reCAPTCHA']"))
)

driver.switch_to.frame(captcha_frame)

captcha_checkbox = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, 'recaptcha-anchor-label'))
)

captcha_checkbox.click()

driver.switch_to.default_content()

# TODO: click on proceed button
submit_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnProceed')))
submit_button.click()


driver.quit()
