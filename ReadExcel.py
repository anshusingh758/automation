import xlrd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import time

driver = webdriver.Chrome(ChromeDriverManager().install())
driver.set_window_size(1024, 600)
driver.maximize_window()
driver.implicitly_wait(10)

driver.get('https://www.orangehrm.com/orangehrm-30-day-trial/')

url = driver.find_element_by_id('Form_submitForm_subdomain')
first_name = driver.find_element_by_id('Form_submitForm_FirstName')
last_name = driver.find_element_by_id('Form_submitForm_LastName')
email = driver.find_element_by_id('Form_submitForm_Email')
job_title = driver.find_element_by_id('Form_submitForm_JobTitle')
no_of_employee = driver.find_element_by_id('Form_submitForm_NoOfEmployees')
company_name = driver.find_element_by_id('Form_submitForm_CompanyName')
industry = driver.find_element_by_id('Form_submitForm_Industry')
contact = driver.find_element_by_id('Form_submitForm_Contact')
country = driver.find_element_by_id('Form_submitForm_Country')

workbook = xlrd.open_workbook("DataFile.xlsx")
sheet = workbook.sheet_by_name("registration")

rowCount = sheet.nrows
colCount = sheet.ncols

for curr_row in range(1, rowCount):
    urlValue = sheet.cell_value(curr_row, 0)
    firstname = sheet.cell_value(curr_row, 1)
    lastname = sheet.cell_value(curr_row, 2)
    emailid = sheet.cell_value(curr_row, 3)
    jobtitle = sheet.cell_value(curr_row, 4)
    company = sheet.cell_value(curr_row, 5)
    phone = sheet.cell_value(curr_row, 6)
    totalemp = sheet.cell_value(curr_row, 7)
    select_industry = sheet.cell_value(curr_row, 8)
    select_country = sheet.cell_value(curr_row, 9)

    url.clear()
    url.send_keys(urlValue)
    first_name.clear()
    first_name.send_keys(firstname)
    last_name.clear()
    last_name.send_keys(lastname)
    email.clear()
    email.send_keys(emailid)
    job_title.clear()
    job_title.send_keys(jobtitle)
    company_name.clear()
    company_name.send_keys(company)
    contact.clear()
    contact.send_keys(phone)
    # no_of_employee.clear()
    no_of_employee.send_keys(totalemp)
    industry.send_keys(select_industry)
    country.send_keys(select_country)
    time.sleep(10)