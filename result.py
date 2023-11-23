from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

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
    
    # Click on the last div to reveal the additional information
    last_div_data.click()

    # Find all elements with class 'col-md-6'
    col_md_6_elements = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'col-md-6')))
    
    # Assume the last 'col-md-6' element contains the desired data
    last_col_md_6 = col_md_6_elements[-1]

    # Find the table within the last 'col-md-6' element
    table = last_col_md_6.find_element(By.TAG_NAME, 'table')

    # Find all rows (tr) within the table
    rows = table.find_elements(By.TAG_NAME, 'tr')

    # Iterate through rows and print the data
    for row_num, row_element in enumerate(rows):
        # Find all columns (td) within the row
        columns = row_element.find_elements(By.TAG_NAME, 'td')

        # Assuming the third and sixth columns contain the desired data
        if len(columns) >= 6:
            semester = columns[2].text
            even_odd = columns[5].text

            # Print the data
            print(f"Semester: {semester}, Even/Odd: {even_odd}")

            # Update the DataFrame
            df.loc[index, 'Semester'] = semester
            df.loc[index, 'Even/Odd'] = even_odd

    # Save the data in the excel file after processing each student
    df.to_excel(r'C:\\Users\\Tanish Singhal\\Desktop\\AKTU result Mini Project.xlsx', index=False)

    # Run the main URL again, to fetch the data of the next student
    driver.get(initial_url)

# Close the browser with a delay to allow time for the Excel file to save
time.sleep(3)
driver.quit()
