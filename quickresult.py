from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# Load the Excel file and the website
df = pd.read_excel(r'C:\\Users\\Tanish Singhal\\Desktop\\AKTU result Mini Project.xlsx')

driver = webdriver.Chrome()
initial_url = 'https://erp.aktu.ac.in/webpages/oneview/oneview.aspx'
driver.get(initial_url)

wait = WebDriverWait(driver, 5)

for index, row in df.iterrows():
    # Roll No part
    roll_number_input = wait.until(EC.presence_of_element_located((By.ID, 'txtRollNo')))
    roll_number_input.clear()
    roll_number_input.send_keys(str(row['rollno']))

    submit_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnProceed')))
    submit_button.click()

    # Date of Birth Part
    dob_input = wait.until(EC.presence_of_element_located((By.ID, 'txtDOB')))
    formatted_dob = row['dob'].strftime('%Y-%m-%d')
    dob_input.clear()
    dob_input.send_keys(formatted_dob)

    # Captcha Manually Solve
    captcha_input = input("Please solve the captcha manually and press Enter when done: ")

    submit_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnSearch')))
    submit_button.click()

    # Second page start
    # TODO: Main div data extracted
    second_page_elements = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//tbody/tr[2]/td/div')))

    # Initialize dictionaries for each div
    div_data_dicts = []

    # Extract data for each div
    for div_element in second_page_elements:
        div_data = div_element.text

        # Check if the expected data exists before trying to access it
        session = div_data.split('Session : ')[1].split(' ')[0] if 'Session' in div_data else ''
        semesters = div_data.split('Semesters : ')[1].split(' ')[0] if 'Semesters' in div_data else ''
        result = div_data.split('Result : ')[1].split(' ')[0] if 'Result' in div_data else ''
        marks = div_data.split('Marks : ')[1].split(' ')[0] if 'Marks' in div_data else ''
        cop = div_data.split('COP : ')[1].split(' ')[0] if 'COP' in div_data else ''

        div_dict = {
            'Session': session,
            'Semesters': semesters,
            'Result': result,
            'Marks': marks,
            'COP': cop
        }
        div_data_dicts.append(div_dict)

    # Update DataFrame with div data
    for i, div_dict in enumerate(div_data_dicts, start=1):
        df.loc[index, f'Session_div{i}'] = div_dict.get('Session', '')
        df.loc[index, f'Semesters_div{i}'] = div_dict.get('Semesters', '')
        df.loc[index, f'Result_div{i}'] = div_dict.get('Result', '')
        df.loc[index, f'Marks_div{i}'] = div_dict.get('Marks', '')
        df.loc[index, f'COP_div{i}'] = div_dict.get('COP', '')

    # TODO: Save the data in the Excel file
    df.to_excel(r'C:\Users\Tanish Singhal\Desktop\AKTU result Mini Project.xlsx', index=False)

    driver.get(initial_url)

time.sleep(5)
driver.quit()
