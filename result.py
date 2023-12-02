from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# Load the Excel file
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
    df.loc[index, 'overallResult'] = last_div_data

    # Click on the last div
    last_div = second_page_elements[-1]
    last_div.click()

    # Extract and update the specific data
    label_to_column = {
        'Semester': 'semester',
        'Even/Odd': 'even_odd',
        'Total Subjects': 'totalSubjects',
        'Theory Subjects': 'theorySubjects',
        'Practical Subjects': 'practicalSubjects',
        'Total Marks Obt.': 'totalMarks',
        'Result Status': 'resultStatus',
        'SGPA': 'sgpa',
        'Date of Declaration': 'declarationDate'
    }

    for row in last_div.find_elements(By.TAG_NAME, 'tr'):
        cells = row.find_elements(By.TAG_NAME, 'td')

        # Check if the row has at least 2 cells (label and value)
        if len(cells) >= 2:
            label = cells[0].text.strip()
            value = cells[-1].text.strip()

            # Map label to corresponding column name in the Excel file
            column_name = label_to_column.get(label)

            # If the label is in the mapping, update the DataFrame with the extracted value
            if column_name:
                # Explicitly cast the value to float if needed
                if column_name in ['totalMarks', 'sgpa']:
                    df.loc[index, column_name] = float(value)
                else:
                    df.loc[index, column_name] = value
        else:
            print(f"Skipping row: {row.text}")

    # Save the data in the Excel file
    df.to_excel(r'C:\Users\Tanish Singhal\Desktop\AKTU result Mini Project.xlsx', index=False)

    # Run the main URL again, to fetch the data of the next student
    driver.get(initial_url)

# Close the browser with a delay to allow time for the Excel file to save
time.sleep(5)
driver.quit()
