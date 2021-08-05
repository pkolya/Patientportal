from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
driver = webdriver.Chrome(executable_path="D:/Back up/Back up/pyprojects/PATIENTPORTAL/Drivers/chromedriver.exe")
url = "https://opensource-demo.orangehrmlive.com/"
username = "Admin"
password = "admin123"
txt_username_name = "txtUsername"
txt_password_name = "txtPassword"
btn_login_id = "btnLogin"
tab_Admin_id = "menu_admin_viewAdminModule"
submenu_UserManagement_id = "menu_admin_UserManagement"
submenu_Users_id = "menu_admin_viewSystemUsers"
driver.maximize_window()
driver.get(url)
driver.find_element_by_name(txt_username_name).send_keys(username)
driver.find_element_by_name(txt_password_name).send_keys(password)
driver.find_element_by_id(btn_login_id).click()
pim = driver.find_element_by_id("menu_pim_viewPimModule")
add_employee = driver.find_element_by_id("menu_pim_addEmployee")
actions = ActionChains(driver)
actions.move_to_element(pim).move_to_element(add_employee).click().perform()
driver.find_element_by_id("firstName").send_keys("Ashok")
driver.find_element_by_id("middleName").send_keys("Kumar")
driver.find_element_by_id("lastName").send_keys("Reddymasi")
driver.find_element_by_xpath('//*[@id="chkLogin"]').click()
time.sleep(3)
# wait=WebDriverWait(driver,10)
# wait.until(EC.element_to_be_clickable((By.ID, 'photofile')))
#tc.click()
photopath='D:\photos\dermatology.PNG'
#driver.switch_to_frame('frmAddEmp')
driver.find_element_by_id("photofile").send_keys(photopath)
for x in driver.window_handles:
    print(x.index(0))
#driver.find_element_by_xpath("//input[@id='photofile']").send_keys("‪C://Users//areddymasi//Desktop//Spy.PNG")