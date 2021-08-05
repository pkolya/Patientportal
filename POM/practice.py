from selenium import webdriver
import time
from selenium.webdriver.support.ui import Select

driver = webdriver.Chrome(executable_path='D:/Back up/Back up/pyprojects/PATIENTPORTAL/Drivers/chromedriver.exe')

url='https://courses.letskodeit.com/practice'
bmwradio_xpath='//*[@id="bmwradio"]'
benz_select_xpath='//*[@id="carselect"]'
multiselect_xpath='//*[@id="multiple-select-example"]'
hondacheckbox_xpath='//*[@id="hondacheck"]'
switchbutton_xpath='//*[@id="openwindow"]'
opentab_xpath='//*[@id="opentab"]'
textbox='//*[@id="name"]'
alert_button='//*[@id="alertbtn"]'
tab_link='//*[@id="opentab"]'
name_text_xpath='//*[@id="name"]'


driver.maximize_window()
driver.get(url)
time.sleep(2)
driver.find_element_by_xpath(bmwradio_xpath).click()
sel=Select(driver.find_element_by_xpath(benz_select_xpath))
sel.select_by_visible_text('BMW')
sel2=Select(driver.find_element_by_xpath(multiselect_xpath))
sel2.select_by_visible_text('Apple')
sel2.select_by_visible_text('Orange')
driver.find_element_by_xpath(hondacheckbox_xpath).click()
time.sleep(2)
#driver.find_element_by_link_text('Open Tab').click()
#time.sleep(3)
#driver.switch_to.window(driver.window_handles[1])
#driver.close()
'''driver.find_element_by_xpath(switchbutton_xpath).click()
x=driver.current_window_handle
driver.switch_to.window(driver.window_handles[1])
driver.maximize_window()
driver.close()
driver.switch_to.window(driver.window_handles[0])
time.sleep(2)'''
#driver.find_element_by_xpath(tab_link).click()
#driver.find_element_by_link_text('Open Tab').click()
#time.sleep(3)
#driver.switch_to.window(driver.window_handles[1])
driver.find_element_by_xpath(name_text_xpath).send_keys('pkolya')
driver.find_element_by_xpath('//*[@id="alertbtn"]').click()
time.sleep(3)
driver.switch_to.alert.dismiss()
td=driver.find_elements_by_xpath('//*[@id="product"]/tbody')
for i in td:
    print(i.text)



driver.close()
        





