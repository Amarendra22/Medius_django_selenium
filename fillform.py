from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time

def fill_form_from_excel(excel_file):
    driver = webdriver.Chrome()
    url = "https://forms.gle/WT68aV5UnPajeoSc8"
    driver.get(url)
    driver.maximize_window()
    excel_data = pd.read_excel(excel_file)
    for index, row in excel_data.iterrows():
        fullname_val = row['Full Name']
        contactno_val = str(row['Contact No'])
        emailid_val = row['Email Id']
        fulladdress_val = row['Full Address']
        pincode_val = str(row['Pin Code'])
        dob_val = row['DOB'].strftime('%d-%m-%Y')
        gender_val = row['Gender']

        fullname = driver.find_element(By.XPATH,"//input[@class='whsOnd zHQkBf' and @jsname='YPqjbf' and @aria-labelledby='i1']")
        contactno = driver.find_element(By.XPATH, "//input[@class='whsOnd zHQkBf' and @jsname='YPqjbf' and @aria-labelledby='i5']")
        emailid = driver.find_element(By.XPATH,"//input[@class='whsOnd zHQkBf' and @jsname='YPqjbf' and @aria-labelledby='i9']")
        fulladdress = driver.find_element(By.XPATH,"//textarea[@class='KHxj8b tL9Q4c']")
        pincode = driver.find_element(By.XPATH,"//input[@class='whsOnd zHQkBf' and @jsname='YPqjbf' and @aria-labelledby='i17']")
        dob = driver.find_element(By.XPATH, "//input[@class='whsOnd zHQkBf' and @jsname='YPqjbf' and @aria-labelledby='i25']")
        gender = driver.find_element(By.XPATH, "//input[@class='whsOnd zHQkBf' and @jsname='YPqjbf' and @aria-labelledby='i26']")

    
        fullname.send_keys(fullname_val)
        contactno.send_keys(contactno_val)
        emailid.send_keys(emailid_val)
        fulladdress.send_keys(fulladdress_val)
        pincode.send_keys(pincode_val)

        dob.clear()
        dob.send_keys(dob_val)
        gender.send_keys(gender_val)

        code_element = driver.find_element(By.XPATH, "//div[@id='i30' and @class='HoXoMd D1wxyf RjsPE' and @aria-level='3']//span[@class='M7eMe']//b")
        code = code_element.text
        code_box = driver.find_element(By.XPATH,"//input[@class='whsOnd zHQkBf' and @jsname='YPqjbf' and @aria-labelledby='i30']")
        code_box.send_keys(code)
        time.sleep(1)
    
        submit_button = driver.find_element(By.XPATH, "//span[text()='Submit']")
        submit_button.click()
        time.sleep(1)

        screenshot_name = f"screenshot_{index+1}.png"
        driver.save_screenshot(screenshot_name)
        time.sleep(1)
        driver.refresh()
        time.sleep(1)
    
fill_form_from_excel('data.xlsx')