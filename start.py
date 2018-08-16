from tkinter import *
# import tkinter.messagebox
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Self defined classes
from accessexcel import accessExcel
from a_upload import browserAccess
from a_upload import browser

uploadType = ""
oldUrl = ""
StatusText = "Loading..."
loggedin=""
# rownum = 5
# print(type(rownum))
dname = browserAccess("https://login.veevavault.com/auth/login", "girish@vv-agency.com", "V@ult123")

def updateExcel(rowNum):
    print(rowNum)
    exlFilePath = excelPathVar.get()
    excelFile = accessExcel(exlFilePath, "Sheet1")
    excelFile.selectSheet()
    excelData = excelFile.accessFields(int(rowNum))

    if(excelData["Document Type"] == "Email Fragment"):
        excelFile.modifySheet(excelData["AUT"], 'XX-XXXXX', excelData["EF Number"], "E5" )
        excelFile.modifySheet(excelData["Tracking Id"], '_', '%5F', "E5" )
        excelFile.modifySheet(excelData["Tracking Id"], '-', '%2D', "E5" )

    return excelData


def uploadTemplate():
    rnum=templateRow.get()
    excelData=updateExcel(rnum)
    commonTasks(excelData)

def uploadAsset():
    rnum=assetRow.get()
    # updateExcel()
    excelData=updateExcel(rnum)
    commonTasks(excelData)

    oldUrl = browser.current_url
    print("OLD URL " + oldUrl)
    try:
        WebDriverWait(browser, 100).until(EC.url_changes(oldUrl))
        activeLink = browser.current_url
        print("ACTIVE LINK " + activeLink)
        subStr = activeLink[-8:]
        docInfoNumber = subStr[0:4]
        print("Doc No: "+ docInfoNumber)

    except Exception as pn:
        print("Not found")


def uploadFragment():
    rnum=fragmentRow.get()
    excelData=updateExcel(rnum)
    commonTasks(excelData)


    # Open web Driver
def commonTasks(excelData):
    # dname = browserAccess("https://login.veevavault.com/auth/login", "girish@vv-agency.com", "V@ult123")
    print("LOGGEDIN  "+ str(dname.loggedin))
    if(dname.loggedin==False):
        dname.siteLogin()
        dname.elementAccess({"findtype": "id", "findText":"search_main_box", "clear" : False, "sendkeys": False, "enterRequired" : False, "click" : False})
        browser.get('https://vv-agency-indegene.veevavault.com/ui/#inbox/upload')
        dname.loggedin=True
    else:
        browser.get('https://vv-agency-indegene.veevavault.com/ui/#inbox/upload')


    # statusVar.set("Entering upload page...")

    # Select file to upload
    dname.elementAccess({"findtype": "id", "findText":"inboxFileChooserHTML5", "clear" : False, "sendkeys": True, "keyText": excelData["Path"], "enterRequired" : False, "click" : False})
    # Select Document type
    dname.elementAccess({"findtype": "xpath", "findText":'//*[@id="inboxUploadPageContent"]/div[1]/div[2]/div[2]/div/div[1]/div/div[2]/input', "clear" : True, "sendkeys": True, "keyText": excelData["Document Type"], "enterRequired" : True, "click" : True})
    dname.elementAccess({"findtype": "linktext", "findText":excelData["Document Type"], "clear" : False, "sendkeys": False, "enterRequired" : False, "click" : True})
    # Click next button
    dname.elementAccess({"findtype": "id", "findText":'inboxUploadNext', "clear" : False, "sendkeys": False, "enterRequired" : False, "click" : True})
    # Document name
    dname.elementAccess({"findtype": "name", "findText":'name', "clear" : True, "sendkeys": True, "keyText": excelData["Name"], "enterRequired" : False, "click" : False})
    # Document Title
    dname.elementAccess({"findtype": "name", "findText":'title', "clear" : True, "sendkeys": True, "keyText": excelData["Title"], "enterRequired" : False, "click" : False})
    # Email template Type
    if(excelData["Document Type"] == "Email Template"):
        dname.elementAccess({"findtype": "xpath", "findText":'//*[@id="di3Form"]/div[2]/div[1]/div/div[1]/div[6]/div/div[2]/div/div[1]/input', "clear" : False, "sendkeys": True, "keyText": excelData["Email Template Type"], "enterRequired" : False, "click" : True})
        dname.elementAccess({"findtype": "linktext", "findText": excelData["Email Template Type"], "clear" : False, "sendkeys": False, "enterRequired" : False, "click" : True})
        # Country for Email Template
        dname.elementAccess({"findtype": "xpath", "findText":'//*[@id="di3Form"]/div[2]/div[1]/div/div[1]/div[10]/div/div[2]/div/div[1]/input', "clear" : True, "sendkeys": True, "keyText": excelData["Country"], "enterRequired" : False, "click" : True})
    else:
    # Country
        dname.elementAccess({"findtype": "xpath", "findText":'//*[@id="di3Form"]/div[2]/div[1]/div/div[1]/div[9]/div/div[2]/div/div[1]/input', "clear" : True, "sendkeys": True, "keyText": excelData["Country"], "enterRequired" : False, "click" : True})

    dname.elementAccess({"findtype": "linktext", "findText": excelData["Country"], "clear" : False, "sendkeys": False, "enterRequired" : False, "click" : True})


    # Account
    if(excelData["Document Type"] == "Email Fragment"):
        dname.elementAccess({"findtype": "xpath", "findText":'//*[@id="di3Form"]/div[2]/div[1]/div/div[1]/div[12]/div/div[2]/div/div[1]/input', "clear" : True, "sendkeys": True, "keyText": excelData["Account"], "enterRequired" : False, "click" : True})
        dname.elementAccess({"findtype": "linktext", "findText": excelData["Account"], "clear" : False, "sendkeys": False, "enterRequired" : False, "click" : True})
    elif (excelData["Document Type"] == "Email Template"):
        dname.elementAccess({"findtype": "xpath", "findText":'//*[@id="di3Form"]/div[2]/div[1]/div/div[1]/div[15]/div/div[2]/div/div[1]/input', "clear" : True, "sendkeys": True, "keyText": excelData["Account"], "enterRequired" : False, "click" : True})
        dname.elementAccess({"findtype": "linktext", "findText": excelData["Account"], "clear" : False, "sendkeys": False, "enterRequired" : False, "click" : True})
    # Product
    dname.elementAccess({"findtype": "xpath", "findText":'//*[@id="di3Form"]/div[2]/div[2]/div/div[1]/div/div/div[2]/div/div[1]/input', "clear" : True, "sendkeys": True, "keyText": excelData["Product"], "enterRequired" : False, "click" : True})
    dname.elementAccess({"findtype": "linktext", "findText": excelData["Product"], "clear" : False, "sendkeys": False, "enterRequired" : False, "click" : True})

    if(excelData["Document Type"] == "Email Template"):
        dname.elementAccess({"findtype": "name", "findText":'subject_b', "clear" : True, "sendkeys": True, "keyText": excelData["Subject"], "enterRequired" : False, "click" : False})

        # From address
        dname.elementAccess({"findtype": "name", "findText":'fromAddress_b', "clear" : True, "sendkeys": True, "keyText": excelData["From Address"], "enterRequired" : False, "click" : False})
        dname.elementAccess({"findtype": "name", "findText":'fromName_b', "clear" : True, "sendkeys": True, "keyText": excelData["From Name"], "enterRequired" : False, "click" : False})
        dname.elementAccess({"findtype": "name", "findText":'replyToAddress_b', "clear" : True, "sendkeys": True, "keyText": excelData["Reply To Address"], "enterRequired" : False, "click" : False})
        dname.elementAccess({"findtype": "name", "findText":'replyToName_b', "clear" : True, "sendkeys": True, "keyText": excelData["Reply To Name"], "enterRequired" : False, "click" : False})
        dname.elementAccess({"findtype": "xpath", "findText": '//*[@id="di3Form"]/div[3]/div/div/a[1]/span', "clear" : False, "sendkeys": False, "enterRequired" : False, "click" : True})

    # Click save Button
    oldUrl = browser.current_url
    print(type(oldUrl))
    if(excelData["Document Type"] != "Email Template"):
        dname.elementAccess({"findtype": "linktext", "findText": "Save", "clear" : False, "sendkeys": False, "enterRequired" : False, "click" : True})


    if(excelData["Document Type"] == "Email Fragment"):
        addAssetsFragments()
        dname.elementAccess({"findtype": "xpath", "findText":'//*[@id="ui-id-1"]/form/div[1]/div/div[3]/label/input', "clear" : False, "sendkeys": True, "keyText": excelData["Image path"], "enterRequired" : False, "click" : False})
        dname.elementAccess({"findtype": "linktext", "findText": "Upload", "clear" : False, "sendkeys": False, "enterRequired" : False, "click" : True})

    if(excelData["Document Type"] == "Email Template" ):
        addAssetsTemplate()
        print("Post add template")

        dname.elementAccess({"findtype": "xpath", "findText":'//*[@id="ui-id-1"]/form/div[1]/div/div[3]/label/input', "clear" : False, "sendkeys": True, "keyText": excelData["Image path"], "enterRequired" : False, "click" : False})
        dname.elementAccess({"findtype": "linktext", "findText": "Upload", "clear" : False, "sendkeys": False, "enterRequired" : False, "click" : True})


def addAssetsFragments():
        try:
            assetBtn = WebDriverWait(browser, 100).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="di3Form"]/div[2]/div[4]/h3/a')))
            assetBtn.click()
        except Exception as xp:
            print("Waiting asset dialog box" + xp)
            addAssetsFragments()

def addAssetsTemplate():
        print("Inside add template")
        try:
            assetBtn = WebDriverWait(browser, 100).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="di3Form"]/div[2]/div[5]/h3/a')))
            assetBtn.click()
        except Exception as xp:
            print("Waiting asset dialog box")
            addAssetsTemplate()

def closeWindow():
    root.destroy()


# GUI-------------------------------------------------
root = Tk()
root.geometry("500x220")
root.title("Indegene Email Uploader")
Label(root, text = "INDEGENE - Email Upload App", fg="Gray",font=("verdana", 11), relief=RIDGE).grid(row=0, column=0, columnspan=4, pady=2, padx=0, sticky=W)

# Data file path input box
Label(root, text="Paste the data file path",width=25,font=("verdana", 10), anchor=E).grid(row=3, column=0, columnspan=3, padx=0)
excelPathVar = StringVar()
Entry(root, width=30,font=("verdana", 9), textvariable=excelPathVar).grid(row=3, column=4, columnspan=4, pady=15, sticky=W)

# Template upload
Label(root, text="Template row number",width=25,font=("verdana", 10), anchor=E).grid(row=4, column=0, columnspan=3)
templateRow = Spinbox(root, from_=3, to=100, width=5,font=("verdana", 10))
templateRow.grid(row=4, column=4,sticky=W)
Button(root, text="Upload Template",width=20,font=("verdana", 10), command=uploadTemplate).grid(row=4, column=6)

# Asset upload
Label(root, text="Assets row number",width=25,font=("verdana", 10), anchor=E).grid(row=5, column=0, columnspan=3)
assetRow = Spinbox(root, from_=3, to=100, width=5,font=("verdana", 10))
assetRow.grid(row=5, column=4,sticky=W)
Button(root, text="Upload Assets",width=20,font=("verdana", 10),  command=uploadAsset).grid(row=5, column=6)

# Fragment upload
Label(root, text="Fragment row number",width=25,font=("verdana", 10), anchor=E).grid(row=6, column=0, columnspan=3)
fragmentRow = Spinbox(root, from_=3, to=100, width=5,font=("verdana", 10))
fragmentRow.grid(row=6, column=4,sticky=W)

Button(root, text="Upload Fragment",width=20,font=("verdana", 10),  command=uploadFragment).grid(row=6, column=6)

# Quit button
Button(root, text="Close",command=closeWindow,font=("verdana", 10)).grid(row=8, column=6, sticky=E, pady=10)

# Status
statusVar=StringVar()
Label(root, text="Loading...",font=("verdana", 8), textvariable=statusVar).grid(row=8, column=0, columnspan=5, sticky=W, padx=5)

root.mainloop()