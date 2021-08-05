import time
from selenium import webdriver
import openpyxl

driver = webdriver.Chrome(executable_path='..//Drivers//chromedriver.exe')
url='https://www.w3schools.com/html/html_tables.asp'
driver.get(url)
driver.maximize_window()
#element=driver.find_element_by_xpath('//*[@id="post-41633"]/div[1]/div/div[2]/div/div/div/div/p[8]')
#driver.execute_script("arguments[0].scrollIntoView();", element)
webtablexpath_rows=len(driver.find_elements_by_xpath('//*[@id="customers"]/tbody/tr'))

webtablexpath_cols = len(driver.find_elements_by_xpath('//*[@id="customers"]/tbody/tr/th'))
print(webtablexpath_rows)
print(webtablexpath_cols)
wb = openpyxl.load_workbook('D:/OrthoASC/source/pythondemo.xlsx')
ws = wb.worksheets[0]

for x in range (2,webtablexpath_rows+1):
    for y in range(1,webtablexpath_cols+1):
        webtablexpath_row=driver.find_element_by_xpath('//*[@id="customers"]/tbody/tr['+str(x)+']/td['+str(y)+']')
        print(webtablexpath_row.text,end="     ")
        ws.cell(x,y).value=webtablexpath_row.text
        wb.save('D:/OrthoASC/source/pythondemo.xlsx')
    print()
