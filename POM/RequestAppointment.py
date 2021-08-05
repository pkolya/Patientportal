from selenium import webdriver
from selenium.webdriver.support.ui import Select
import time
import sys

class requestAppointment():

    Schedule_id = "ctl00_ucHeader_lnkScheduleMenuItem"
    Request_appt = "ctl00_ucHeader_lnkRequest_Appointment"
    practice_select_id = "ctl00_ContentPlaceHolder1_PracticePersonSelector1_ddlPractices"
    selectProvider_id = "ctl00_ContentPlaceHolder1_ddlProviders"
    select_category_id = "ctl00_ContentPlaceHolder1_ddlCategories"
    select_location_id = "ctl00_ContentPlaceHolder1_ddlLocation"
    reason_for_appointment_id = "ctl00_ContentPlaceHolder1_txtReason"
    select_priority_id = "ctl00_ContentPlaceHolder1_ddlUrgency"
    select_make_appointment_for_id = "ctl00_ContentPlaceHolder1_ddlAppointmentFor"
    select_preferred_time_from_id = "ctl00_ContentPlaceHolder1_ddlStartTimes"
    select_preferred_time_to_id = "ctl00_ContentPlaceHolder1_ddlEndTimes"
    select_alternate_time_from_id = "ctl00_ContentPlaceHolder1_ddl2ndStartTimes"
    select_alternate_time_to_id = "ctl00_ContentPlaceHolder1_ddl2ndEndTimes"
    Submit_button_xpath = '//*[@id="ctl00_ContentPlaceHolder1_btnSubmit"]'
    start_date_id="ctl00_ContentPlaceHolder1_txtStartDate"
    end_date_id="ctl00_ContentPlaceHolder1_txtEndDate"
    # cls.driver=driver
#try:

    def __init__(self, driver):
        self.driver = driver

    def practiceselect(self,practicevalue):
        try:
            Sele = Select(self.driver.find_element_by_id(self.practice_select_id))
            Sele.select_by_visible_text(practicevalue)
            self.driver.implicitly_wait(10)
        except Exception as e:
            print(e)
            print("Failed to select value")
            print()


    def providerselect(self, providerValue):
        Sele = Select(self.driver.find_element_by_id(self.selectProvider_id))
        Sele.select_by_visible_text(providerValue)
        time.sleep(3)
        self.driver.implicitly_wait(10)

    def categoryselect(self, categoryvalue):
        Sele = Select(self.driver.find_element_by_id(self.select_category_id))
        Sele.select_by_visible_text(categoryvalue)
        self.driver.implicitly_wait(10)

    def locationselect(self, locationvalue):
        Sele = Select(self.driver.find_element_by_id(self.select_location_id))
        Sele.select_by_visible_text(locationvalue)
        self.driver.implicitly_wait(10)

    def enterreasonforvisit(self,reason):
        self.driver.find_element_by_id(self.reason_for_appointment_id).send_keys(reason)
        self.driver.implicitly_wait(10)

    def priorityselect(self,priority):
        Sele=Select(self.driver.find_element_by_id(self.select_priority_id))
        Sele.select_by_visible_text(priority)
        self.driver.implicitly_wait(10)

    def makeappointmentforselect(self, apptfor):
        Sele = Select(self.driver.find_element_by_id(self.select_make_appointment_for_id))
        Sele.select_by_visible_text(apptfor)
        self.driver.implicitly_wait(10)

    def enterstartdate(self, startdate):
        self.driver.find_element_by_id(self.start_date_id).send_keys(startdate)
        self.driver.implicitly_wait(10)

    def enterendate(self, enddate):
        self.driver.find_element_by_id(self.end_date_id).send_keys(enddate)
        self.driver.implicitly_wait(10)

    def selectpreferedtimefrom(self, timefrom):
        Sele = Select(self.driver.find_element_by_id(self.select_preferred_time_from_id))
        Sele.select_by_visible_text(timefrom)
        self.driver.implicitly_wait(10)

    def selectpreferedtimeto(self, timeto):
        Sele = Select(self.driver.find_element_by_id(self.select_preferred_time_to_id))
        Sele.select_by_visible_text(timeto)
        self.driver.implicitly_wait(10)

    def clicksubmitbutton(self):
        self.driver.find_element_by_xpath(self.Submit_button_xpath).click()

    def scrollPagefull(self):
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# except:
#     print("error except")

# finally:
#     print("done")
