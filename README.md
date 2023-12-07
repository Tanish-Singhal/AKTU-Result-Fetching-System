
# AKTU Result Fetching System

This Python script is designed to automate the process of fetching AKTU (Dr. A.P.J. Abdul Kalam Technical University) results for a list of students provided in an Excel file. The script uses Selenium, a web testing framework, to navigate through the AKTU result portal, input student details, and extract result data. The results are then appended to the same Excel file.
## Features

- **Excel Integration:** The script reads student details (roll number and date of birth) from an Excel file.

- **Web Automation:** Utilizes Selenium for automating interactions with the AKTU result portal.

- **Data Extraction:** Extracts overall results, semester-wise information, and subject-wise marks for each student.

- **Excel Update:** Appends the extracted data to the same Excel file for easy tracking and analysis.


## Prerequisites

To run this project, you need to have the following installed:

- VS Code
- Selenium: 
```
    pip install selenium
```
- Pandas: 
```
    pip install pandas
```
- ChromeDriver
## Usage/Examples

- **Excel File Preparation:** Create an Excel file with columns 'rollno' and 'dob' for each student.

- **Install Dependencies:** 
```
    pip install selenium pandas
```
- **ChromeDriver:** Download ChromeDriver and ensure its path is in your system's PATH variable.
- **Run the Script:** Execute the Python script, providing the correct path to your Excel file.
```
    python akut_result_fetch.py
```
- **Captcha Interaction:** During script execution, manually solve the captcha when prompted in the console.
- **Results:** The script will update the Excel file with the extracted results, including overall results, semester-wise information, and subject-wise marks.

## Important Section

- The script assumes that the structure of the AKTU result portal remains consistent. Any changes to the website structure may require updates to the script.
- Manually solve the captcha when prompted during script execution.
- Ensure that ChromeDriver is compatible with your Chrome browser version.
## Disclaimer

- This script is intended for educational purposes only. Ensure compliance with AKTU's terms of service and policies.
- Use responsibly and avoid excessive requests to the AKTU server to prevent any potential issues.
## Contributors

@Tanish-Singhal: Tanish Singhal
## License

This project is licensed under the [MIT](https://choosealicense.com/licenses/mit/) License - see the LICENSE.md file for details.


## Acknowledgements

- The script was developed as a mini project for educational purposes.
- Thanks to the Selenium and Pandas developers for their contributions to open-source software.
