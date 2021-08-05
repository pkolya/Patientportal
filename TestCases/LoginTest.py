from selenium import webdriver
from selenium.webdriver.common import keys
import time
import sys
import unittest
import HtmlTestRunner
sys.path.append("D:\Back up\Back up\pyprojects\NextMD")
#sys.path.append("C:/Users/pkolya/PycharmProjects/HRMSOrange/Drivers/chromedriver.exe")
from POM.LoginPage_PatientPortal import Login_Page
from POM.AddUser import enter_user_details


class TestLogin(unittest.TestCase):
    driver = webdriver.Chrome(executable_path='D:\Back up\Drivers\chromedriver.exe')
    @classmethod
    def setUpClass(cls):
        #cls.driver = webdriver.Chrome(executable_path='..//Drivers//chromedriver.exe')
        cls.driver.get("https://www.nextmd.com/ud2/Login/Login.aspx")
        cls.driver.maximize_window()

    def test_01login(self):
        Lp=Login_Page(self.driver)
        Lp.set_username("admin")
        Lp.set_password("admin123")
        Lp.click_login_btn()
        print("Login successful")
        print(self.driver.title)
        #Lp.click_admintab()

    def test_02createuser(self):
        lp1 = Login_Page(self.driver)
        lp1.click_admintab()
        lp1.click_admin_add()
        print("user navigated to add user page")

    def test_03enter_user_details(self):
        AP = enter_user_details(self.driver)
        AP.selectuserrole("Admin")
        AP.enter_empolyeename("Hannah Flores")
        AP.enter_username("AdminHOld23")
        AP.select_status("Enabled")
        AP.enter_Password("AdminNew")
        AP.enter_conf_password("AdminNew")
        AP.click_save()
        AP.take_screenshot()
        print("success")
        time.sleep(2)

    @classmethod
    def tearDownClass(cls):
        cls.driver.get_screenshot_as_file('..//Screenshots/12.png')
        #cls.driver.close()
if __name__ == '__main__':
    unittest.main(testRunner=HtmlTestRunner.HTMLTestRunner(output='..//Reports'))



