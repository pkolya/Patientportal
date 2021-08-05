import xlrd
import pyodbc
from selenium import webdriver
import sys

sys.path.append('D:/Back up/Back up/pyprojects/PATIENTPORTAL')
PATH ='..//Drivers//chromedriver.exe'
driver=webdriver.Chrome(PATH)
driver.maximize_window()
driver.get('https://www.countries-ofthe-world.com/flags-of-the-world.html')
title1 = driver.title
print(title1)
#assert "countries" in tit
##first way
#driver.execute_script("window.scrollBy(0,500)","")

##second
#ele=driver.find_element_by_xpath('//*[@id="content"]/div[2]/div[2]/table[2]/tbody/tr[46]/td[1]/img')
#driver.execute_script('arguments[0].scrollIntoView();', ele)

##scroll3 window.scrollBy(0,document.body.scrollHeight

driver.execute_script("window.scrollBy(0,1000)")
assert "Country flags of the world with images and names"==title1
print("pass")