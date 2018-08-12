
# --------- EXTRACT DATA FROM EXCEL FILE ------------------------
import openpyxl
# import os

book = openpyxl.load_workbook("email_data.xlsx")
sheet = book.get_sheet_by_name("Sheet1")

# Data collection ------------------
# dmentType = sheet['K2'].value
uploadFilePath = sheet['L2'].value
sesName = sheet['C2'].value
sesTitle = sheet['D2'].value
dmentType = sheet['K2'].value
sesCountry = sheet['E2'].value
sesAccount = sheet['F2'].value
sesProduct = sheet['G2'].value
imageFilePath = sheet['J2'].value
assetFilePath = sheet['L2'].value
efNumber = sheet['I2'].value
uit = sheet['H2'].value
print(uit)


uit = uit.replace('XX-XXXXX', efNumber)
uit = uit.replace('_', '%5F')
uit = uit.replace('-', '%2D')

sheet['N2'] = uit

book.save("email_data.xlsx")
gdType = "Reference Document › Data on File"
print(uit)


delay = 5
# -------- TASKS IN INDEGENE SANDBOX ------------------------


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait  
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException





# Open Chrome driver

browser = webdriver.Chrome('D:\\chromedriver')
print("Browser opened")
browser.maximize_window()


# METHODS --------------------------------------------
# Login Function
def siteLogin(url, uname, pwd):
    browser.get(url)
    username = browser.find_element_by_id('j_username')
    username.send_keys(uname)
    password = browser.find_element_by_id('j_password')
    password.send_keys(pwd)
    password.send_keys(Keys.ENTER)


# Choose upload file
def chooseFileFunc(dType):
    gdType=dType
    
    if dType == "Email Fragment":
        uploadFilePath = sheet['B2'].value
    elif dType == "Reference Document › Data on File":
        uploadFilePath = sheet['L2'].value
    
    try:
        chooseFile = browser.find_element_by_id('inboxFileChooserHTML5')
        chooseFile.send_keys(uploadFilePath)
        print("Successful")
    except Exception as axp:
        print("Click ok Again")
        chooseFileFunc("Email Fragment")

# Click Next and Save Button
def clicknextBtn():
    try:
        nextBtn = browser.find_element_by_id('inboxUploadNext')
        nextBtn.click()
        print("Able to find NEXT BUTTON")
        clicknextBtn()
        
    except Exception as axp:
        print("unable to find NEXT BUTTON")


def clicksaveBtn():
    try:
        saveBtn = browser.find_element_by_id('inboxUploadSave')
        saveBtn.click()
        print('Save clicked')
    except Exception as axpn:
        print("Unable to click save")
        clicknextBtn()


# Choose Document Type

def documentType(dType):
    docType = browser.find_element_by_xpath('//*[@id="inboxUploadPageContent"]/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/input')
    docType.clear()
    docType.send_keys(dType)
    docType.click()
    docType1 = browser.find_element_by_xpath('//*[@id="veevaBasePage"]/ul')

    for li in docType1.find_elements_by_tag_name('li'):
        print(li.text)
        if li.text == dType:
            li.click() 
            break


# Add Asset
def addAssets():
    try:
        assetBtn = WebDriverWait(browser, 100).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="di3Form"]/div[2]/div[4]/h3/a')))
        assetBtn.click()
    except Exception as xp:
        print("Waiting asset dialog box")
        addAssets()

# Choose image zip
def chooseImgFile():
    
    try:
        imageFile = browser.find_element_by_xpath('//*[@id="ui-id-1"]/form/div[1]/div/div[3]/label/input')
        imageFile.send_keys(imageFilePath)
        print("Successful")
    except Exception as axp:
        print("Waiting...")
        chooseImgFile()

# Update data elements
def updateData():
    
    # Document name
    eleName=browser.find_element_by_name('name')
    eleName.clear()
    eleName.send_keys(sesName)

    # Document Title
    eleTitle=browser.find_element_by_name('title')
    eleTitle.send_keys(sesTitle)
    browser.implicitly_wait(delay)

    # Country
    try:
        countryName = browser.find_element_by_xpath('//*[@id="di3Form"]/div[2]/div[1]/div/div[1]/div[9]/div/div[2]/div/div[1]/input')
        countryName.clear()
        countryName.send_keys(sesCountry)
        countryName.click()
        country = browser.find_element_by_link_text(sesCountry).click()
    except Exception as sp:
        print("Element Not Found")

    # Account
    try:
        accountName = browser.find_element_by_xpath('//*[@id="di3Form"]/div[2]/div[1]/div/div[1]/div[12]/div/div[2]/div/div[1]/input')
        accountName.clear()
        accountName.send_keys(sesAccount)
        accountName.click()
        accountText = browser.find_element_by_link_text(sesAccount).click()
    except Exception as sp:
        print("Element Not Found")

    # Product
    try:
        productName = browser.find_element_by_xpath('//*[@id="di3Form"]/div[2]/div[2]/div/div[1]/div/div/div[2]/div/div[1]/input')
        productName.clear()
        productName.send_keys(sesProduct)
        productName.click()
        linkText = browser.find_element_by_link_text(sesProduct).click()
    except Exception as sp:
        print("Element Not Found")

    oldUrl = browser.current_url
    browser.find_element_by_link_text('Save').click()

    if gdType == "Reference Document › Data on File":
        try:
            WebDriverWait(browser, 100).until(EC.url_changes(oldUrl))
            activeLink = browser.current_url
            subStr = activeLink[-8:]
            docInfoNumber = subStr[0:4]
            sheet['M2'] = docInfoNumber
            book.save("email_data.xlsx")        
        except Exception as pn:
            subStr = activeLink[-8:]
            docInfoNumber = subStr[0:4]
            sheet['M2'] = docInfoNumber
            book.save("email_data.xlsx")

# ---------------------------------------------------        


# COMMON TASKS


siteLogin('https://login.veevavault.com/auth/login',
          'girish@vv-agency.com', 'V@ult123')


print("Logged in")
# -------------

# UPLOAD SECTION

browser.get('https://vv-agency-indegene.veevavault.com/ui/#inbox/upload')

print("Entered upload page")

# WAIT
browser.implicitly_wait(delay)
chooseFileFunc("Reference Document › Data on File")
documentType("Reference Document › Data on File")
# ---------------------------------------------------

# Click NEXT button
clicknextBtn()


browser.implicitly_wait(delay)
updateData()


def frgUpload():
    print("FRGUPLOAD CALLED")
    addAssets()
    chooseImgFile()
    try:
        uploadImgBtn = browser.find_element_by_link_text('Upload')
        uploadImgBtn.click()
        print('Asset uploaded')
    except Exception as xp:
        print('Unable to upoad the asset')



if gdType == "Email Fragment":
    frgUpload()



# siteLogin('https://login.veevavault.com/auth/login',
#           'girish@vv-agency.com', 'V@ult123')