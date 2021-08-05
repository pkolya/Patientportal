from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import unittest
import HtmlTestRunner
import time
from selenium.webdriver import ActionChains
import sys
from selenium.webdriver.support.ui import WebDriverWait
import unittest
sys.path.append("D://Back up/Back up/pyprojects/PATIENTPORTAL")
from POM.LoginPage_PatientPortal import Login_Page
from POM.HomePage_PatientPortal import HomePage
from POM.RequestAppointment import requestAppointment
import time


class TestL(unittest.TestCase):
    un="Roh@prodgr8"
    pwd="Roh@prodgr8"
    driver = webdriver.Chrome(executable_path='..//Drivers//chromedriver.exe')
    Provider="Test_PP3, Provider1"
    Practice="PP#59'ProdEnt1-P2"
    Category="Test_PP3_RTM1"
    Location="PP#59'ProdEnt1-P1_L1"
    Reason="Pain in leg"
    Priority="Normal"
    MakeApptFor="Next Week"
    Startdate="20012021"
    Enddate="20022021"
    Timefrom="6:15 AM"
    TimeTo="6:15 PM"


    @classmethod
    def setUpClass(cls):
        cls.driver.maximize_window()
        cls.driver.get("https://www.nextmd.com/ud2/Login/Login.aspx")

    def test_Login(self):
        Hp=HomePage(self.driver)
        Lp = Login_Page(self.driver)
        # Lp.set_username
        Lp.set_username(self.un)
        Lp.set_password(self.pwd)
        Lp.click_login_btn()
        self.driver.implicitly_wait(10)
        Hp.request_appointment()
        RA=requestAppointment(self.driver)
        RA.practiceselect(self.Practice)
        RA.providerselect(self.Provider)
        time.sleep(3)
        RA.categoryselect(self.Category)
        time.sleep(3)
        RA.locationselect(self.Location)
        time.sleep(3)
        RA.enterreasonforvisit(self.Reason)
        time.sleep(3)
        RA.priorityselect(self.Priority)
        time.sleep(3)
        RA.makeappointmentforselect(self.MakeApptFor)
        time.sleep(3)
        RA.enterstartdate(self.Startdate)
        time.sleep(3)
        RA.enterendate(self.Enddate)
        time.sleep(3)
        RA.selectpreferedtimefrom(self.Timefrom)
        time.sleep(3)
        RA.selectpreferedtimeto(self.TimeTo)
        time.sleep(3)
        RA.scrollPagefull()
        time.sleep(3)
        RA.clicksubmitbutton()


    @classmethod
    def tearDownClass(cls):
        cls.driver.get_screenshot_as_file('..//Screenshots/12.png')
        #cls.driver.close()

if __name__== '__main__':
           unittest.main(testRunner=HtmlTestRunner.HTMLTestRunner(output='..//Reports'))
        # driver.quit()
