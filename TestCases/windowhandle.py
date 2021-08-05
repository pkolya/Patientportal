import time
from selenium import webdriver
import openpyxl

driver = webdriver.Chrome(executable_path='..//Drivers//chromedriver.exe')
url='https://www.w3schools.com/html/html_tables.asp'
driver.get(url)
driver.maximize_window()
driver.find_element_by_id('w3loginbtn').click()
handles=driver.window_handles
driver.cu
for handle in handles:
    driver.switch_to_window(handle)
    print(driver.title)
driver.quit()
driver.coo