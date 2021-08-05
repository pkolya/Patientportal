import openpyxl


class Login_Page():
    txt_username_xpath='//*[@id="ctl00_ContentPlaceHolder1_Login2_txtUserName"]'
    txt_password_id="ctl00_ContentPlaceHolder1_Login2_txtPassword"
    btn_login_xpath='//*[@id="ctl00_ContentPlaceHolder1_Login2_btnLogin1"]'
    demoxl = openpyxl.load_workbook('D:/automation/pythondemo.xlsx')
    demosheet=demoxl.active
    un=demosheet['A2'].value
    pswrd=demosheet['B2'].value

    def __init__(self, driver):
        self.driver = driver

    def set_username(self,username):
        self.driver.find_element_by_xpath(self.txt_username_xpath).send_keys(self.un)

    def set_password(self,password):
        self.driver.find_element_by_id(self.txt_password_id).send_keys(self.pswrd)

    def click_login_btn(self):
        self.driver.find_element_by_xpath(self.btn_login_xpath).click()