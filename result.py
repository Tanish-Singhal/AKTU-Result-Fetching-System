from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# Read the initial DataFrame
df = pd.read_excel(r'C:\\Users\\Tanish Singhal\\Desktop\\AKTU result Mini Project.xlsx')

# Initialize the webdriver
driver = webdriver.Chrome()
initial_url = 'https://erp.aktu.ac.in/webpages/oneview/oneview.aspx'
driver.get(initial_url)

wait = WebDriverWait(driver, 5)

# Iterate through the DataFrame
for index, row in df.iterrows():
    # Create lists to store data from the third and sixth td tags for each student
    data_third_td = []
    data_sixth_td = []

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
    # Move to the targeted div
    second_page_elements = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//tbody/tr[2]/td/div')))
    
    # Extract the data from the last div and update it
    last_div_data = second_page_elements[-1].text
    df.loc[index, 'overallResult'] = last_div_data

    # Click on the last div
    last_div = second_page_elements[-1]
    last_div.click()
    
    # Wait for the new div to appear (you may need to adjust the XPath)
    new_div_xpath = '//div/div[@class="container-fluid"]/div[@class="row"]/div[@class="col-md-6"][last()]'
    new_last_div = wait.until(EC.presence_of_element_located((By.XPATH, new_div_xpath)))

    # Locate the table inside the last div
    table_xpath = '//div[@class="col-md-6"][last()]//table'
    table = new_last_div.find_element(By.XPATH, table_xpath)

    # Find all rows in the table
    rows = table.find_elements(By.TAG_NAME, 'tr')

    # Iterate over the first 6 rows
    for row in rows[:6]:
        # Find all columns in the row
        columns = row.find_elements(By.TAG_NAME, 'td')
        
        # Extract and store data from the third and sixth td tags
        for i, column in enumerate(columns):
            if i in [2]:  # Index 2 for the third td tag
                data_third_td.append(column.text)
            elif i in [5]:  # Index 5 for the sixth td tag
                data_sixth_td.append(column.text)

     # Check the lengths of the lists
    print(len(data_third_td))  # Should be 6
    print(len(data_sixth_td))  # Should be 6

    # Add the lists to your DataFrame if lengths match
    if len(data_third_td) == len(data_sixth_td):
        df.loc[index, 'data_third_td'] = data_third_td
        df.loc[index, 'data_sixth_td'] = data_sixth_td
    else:
        print(f"Lengths of data_third_td and data_sixth_td do not match for Roll No {row['rollno']}")

    # Save the updated data to the Excel file
    df.to_excel(r'C:\Users\Tanish Singhal\Desktop\AKTU result Mini Project.xlsx', index=False)

    # Run the main URL again, to fetch the data of the next student
    driver.get(initial_url)

# Close the browser with a delay to allow time for the Excel file to save
time.sleep(5)
driver.quit()
