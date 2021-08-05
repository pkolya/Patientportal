from selenium.webdriver import ActionChains
from selenium import webdriver
from selenium.webdriver.support.ui import Select
class HomePage():
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

    def __init__(self,driver):
        self.driver=driver

    def request_appointment(self):
        scedule_id_element=self.driver.find_element_by_id(self.Schedule_id)
        request_appt_element=self.driver.find_element_by_id(self.Request_appt)
        action=ActionChains(self.driver)
        action.move_to_element(scedule_id_element).move_to_element(request_appt_element).click().perform()
       # Assert self.driver.title()

    def create_new_appointment(self):
        Sele=Select(self.driver.find_element_by_id(self.practice_select_id))
        Sele.select_by_index(0)
       # Sele.(self.driver.find_element_by_id(self.practice_select_id)).select_by_index(1)
        Sele=Select(self.driver.find_element_by_id(self.selectProvider_id))
        Sele.select_by_index(2)
        self.driver.implicitly_wait(10)
        Sele=Select(self.driver.find_element_by_id(self.select_category_id))
        Sele.select_by_index(1)





       #""" self.driver.find_element_by_id(self.reason_for_appointment_id).send_keys("testing")
        # Select(self.driver.find_element_by_id(self.select_priority_id)).select_by_visible_text("Low")
        # self.driver.find_element_by_id(self.select_make_appointment_for_id).select_by_index(1)
        # self.driver.find_element_by_id(self.select_preferred_time_from_id).select_by_index(1)
        # self.driver.find_element_by_id(self.select_preferred_time_to_id).select_by_index(2)
        # self.driver.find_element_by_id(self.select_alternate_time_from_id).select_by_index(3)
        # self.driver.find_element_by_id(self.select_alternate_time_to_id).select_by_index(4)
        # self.driver.find_element_by_xpath(self.Submit_button_xpath).click() """
