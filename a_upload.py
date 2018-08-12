from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait  
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from accessexcel import accessExcel


# browser = webdriver.Chrome('D:\\chromedriver')
# browser.maximize_window()

browser = webdriver.Chrome('D:\\chromedriver')


class browserAccess:
    
    def __init__(self, url, uname, pwd):
        
        self.url = url
        browser.minimize_window
        browser.get(self.url)

        self.uname = uname
        self.pwd = pwd

# Login the site
    def siteLogin(self):
        browser.get(self.url)
        self.elementAccess({"findtype": "id", "findText":"j_username", "clear" : False, "sendkeys": True, "keyText": self.uname, "enterRequired" : False, "click" : False})
        self.elementAccess({"findtype": "id", "findText":"j_password", "clear" : False, "sendkeys": True, "keyText": self.pwd, "enterRequired" : True, "click" : False})

# To access Elements
    def elementAccess(self, elementDetails):
        elementvar = ""
        browser.implicitly_wait(25)
        try:
            if elementDetails["findtype"] == "id":
                elementvar = browser.find_element_by_id(elementDetails['findText'])
            elif elementDetails["findtype"] == "xpath":
                elementvar = browser.find_element_by_xpath(elementDetails['findText'])
            elif elementDetails["findtype"] == "name":
                elementvar = browser.find_element_by_name(elementDetails['findText'])
            elif elementDetails["findtype"] == "title":
                elementvar = browser.find_element_by_title(elementDetails['findText'])
            elif elementDetails["findtype"] == "linktext":
                elementvar = browser.find_element_by_link_text(elementDetails['findText'])

        except Exception as xpt:
            print("Element Not found")
            self.elementAccess(elementDetails)

        if elementDetails["clear"]:
            elementvar.clear()

        if elementDetails["sendkeys"]:
            elementvar.send_keys(elementDetails["keyText"])
            if elementDetails["enterRequired"]:
                elementvar.send_keys(Keys.ENTER)

        if elementDetails["click"]:
            elementvar.click()
