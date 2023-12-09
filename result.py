from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import random

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

    random_delay = random.uniform(0, 5)
    time.sleep(random_delay)

    # TODO: Date of Birth Part
    dob_input = wait.until(EC.presence_of_element_located((By.ID, 'txtDOB')))
    formatted_dob = row['dob'].strftime('%Y-%m-%d')
    dob_input.clear()
    dob_input.send_keys(formatted_dob)

    # TODO: Captcha Manually Solve
    captcha_input = input("Please solve the captcha manually and press Enter when done: ")

    submit_button = wait.until(EC.element_to_be_clickable((By.ID, 'btnSearch')))

    random_delay = random.uniform(0, 5)
    time.sleep(random_delay)

    submit_button.click()

    # FIXME: Second page start
    # TODO: Main div data extracted
    second_page_elements = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//tbody/tr[2]/td/div')))

    last_div_data = second_page_elements[-1].text
    df.loc[index, 'overallResult'] = last_div_data
    last_div = second_page_elements[-1]
    
    random_delay = random.uniform(0, 5)
    time.sleep(random_delay)

    last_div.click()

    # TODO: Small Table part
    semester_info_div = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@class="col-md-6"]')))
    table_rows = semester_info_div.find_elements(By.XPATH, '//table//tr')

    for row in table_rows:
        cells = row.find_elements(By.XPATH, './/td')
        if len(cells) <= 6:  
            try:
                label_element = cells[0].find_element(By.XPATH, './/span')
                label = label_element.text.strip() if label_element else None
            except Exception as e:
                print(f"Error extracting label: {e}")
                continue

            try:
                value_element = cells[2].find_element(By.XPATH, './/span')
                value = value_element.text.strip() if value_element else None
            except Exception as e:
                print(f"Error extracting value: {e}")
                continue

            try:
                additional_label_element = cells[3].find_element(By.XPATH, './/span')
                additional_label = additional_label_element.text.strip() if additional_label_element else None
            except Exception as e:
                print(f"Error extracting additional label: {e}")
                additional_label = None

            try:
                additional_value_element = cells[5].find_element(By.XPATH, './/span')
                additional_value = additional_value_element.text.strip() if additional_value_element else None
            except Exception as e:
                print(f"Error extracting additional value: {e}")
                additional_value = None

            if label is not None and value is not None:
                df.loc[index, label] = value

            if additional_label is not None and additional_value is not None:
                df.loc[index, additional_label] = additional_value

    # TODO: Subject wise marks Table part
    # Extract subject marks from the table
    subject_marks_table = wait.until(EC.presence_of_element_located((By.ID, 'ctl06_ctl01_ctl00_grdViewSubjectMarksheet')))
    subject_rows = subject_marks_table.find_elements(By.XPATH, './/tr[position()>1]')  # Skip the header row

    # Create lists to store subject details
    subject_codes = []
    internal_marks = []
    external_marks = []

    # Extract subject details and append to lists
    for subject_row in subject_rows:
        subject_cells = subject_row.find_elements(By.XPATH, './/td')
        if len(subject_cells) == 7:
            subject_codes.append(subject_cells[0].text.strip())
            internal_marks.append(subject_cells[3].text.strip())
            external_marks.append(subject_cells[4].text.strip())

    # Update the main DataFrame with subject details
    for i in range(len(subject_codes)):
        col_prefix = f"Subject_{i + 1}"
        df.loc[index, f"{col_prefix}_Code"] = subject_codes[i]
        df.loc[index, f"{col_prefix}_Internal"] = internal_marks[i]
        df.loc[index, f"{col_prefix}_External"] = external_marks[i]

    # TODO: Save the data in the Excel file
    df.to_excel(r'C:\Users\Tanish Singhal\Desktop\AKTU result Mini Project.xlsx', index=False)

    driver.get(initial_url)

time.sleep(5)
driver.quit()