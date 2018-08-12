from tkinter import *
import tkinter.messagebox
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
oldurl = ""
StatusText = "Loading..."


def updateExcel():
    exlFilePath = xlPath.get()
    excelFile = accessExcel(exlFilePath, "Sheet1")
    excelFile.selectSheet()
    excelData = excelFile.accessFields()
    excelFile.modifySheet(excelData["AUT"], 'XX-XXXXX', excelData["EF Number"], "E5" )
    excelFile.modifySheet(excelData["Tracking Id"], '_', '%5F', "E5" )
    excelFile.modifySheet(excelData["Tracking Id"], '-', '%2D', "E5" )    
    return excelData


def uploadTemplate():
    # xlPath.pack_forget()
    rowNumber.pack()
    excelData=updateExcel()
    commonTasks(excelData) 
    


def uploadAsset():
    
    # Access/Update Excel file
    updateExcel()
    excelData=updateExcel()
    commonTasks(excelData)    
   
    try:
        WebDriverWait(browser, 100).until(EC.url_changes(oldUrl))
        activeLink = browser.current_url
        subStr = activeLink[-8:]
        docInfoNumber = subStr[0:4]
        sheet['F5'] = docInfoNumber
        excelFile.book.save(exlFilePath)        
    except Exception as pn:
        print("Not found")


def uploadFragment():
    updateExcel()
    excelData=updateExcel()
    commonTasks(excelData)


    # Open web Driver
def commonTasks(excelData):
    dname = browserAccess("https://login.veevavault.com/auth/login", "girish@vv-agency.com", "V@ult123")
    dname.siteLogin()
    status.config(text="Logging in...")
    
    # StatusText = "Logged in..."
    browser.get('https://vv-agency-indegene.veevavault.com/ui/#inbox/upload')
    status.config(text="Upload page...")
    # Select file to upload
    dname.elementAccess({"findtype": "id", "findText":"inboxFileChooserHTML5", "clear" : False, "sendkeys": True, "keyText": excelData["Path"], "enterRequired" : False, "click" : False})
    status.config(text="Field update")
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
    dname.elementAccess({"findtype": "linktext", "findText": "Save", "clear" : False, "sendkeys": False, "enterRequired" : False, "click" : True})
    

    if(excelData["Document Type"] == "Email Fragment"):
        addAssetsFragments()
        dname.elementAccess({"findtype": "xpath", "findText":'//*[@id="ui-id-1"]/form/div[1]/div/div[3]/label/input', "clear" : False, "sendkeys": True, "keyText": excelData["Image path"], "enterRequired" : False, "click" : False})
        dname.elementAccess({"findtype": "linktext", "findText": "Upload", "clear" : False, "sendkeys": False, "enterRequired" : False, "click" : True})

    if(excelData["Document Type"] == "Email Template" ):
        addAssetsTemplate()
        dname.elementAccess({"findtype": "xpath", "findText":'//*[@id="ui-id-11"]/form/div[1]/div/div[3]/label/input', "clear" : False, "sendkeys": True, "keyText": excelData["Image path"], "enterRequired" : False, "click" : False})
        dname.elementAccess({"findtype": "linktext", "findText": "Upload", "clear" : False, "sendkeys": False, "enterRequired" : False, "click" : True})


def addAssetsFragments():
        try:
            assetBtn = WebDriverWait(browser, 100).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="di3Form"]/div[2]/div[4]/h3/a')))
            assetBtn.click()
        except Exception as xp:
            print("Waiting asset dialog box")
            addAssetsFragments()

def addAssetsTemplate():
        try:
            assetBtn = WebDriverWait(browser, 100).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="di3Form"]/div[2]/div[5]/h3/a')))
            assetBtn.click()
        except Exception as xp:
            print("Waiting asset dialog box")
            addAssetsTemplate()



# GUI-------------------------------------------------
root = Tk()

topFrame = Frame(root)
topFrame.pack()
root.geometry("650x192")
root.title("Email Uplaoder")
label = Label(topFrame, text = "INDEGENE - Email Upload Automation", fg="blue")
label.pack()

fileLabel = Label(topFrame, text="Paste the data file path")
fileLabel.pack()

# Data file path input box
xlPath = Entry(topFrame, text="Testing")
xlPath.insert(0, 'Data file path')
xlPath.pack()

templateBtn = Button(topFrame, text="Upload Template", command=uploadTemplate)
templateBtn.pack(side=LEFT,)

assetBtn = Button(topFrame, text="Upload Assets", command=uploadAsset)
assetBtn.pack(side=LEFT)

fragmentBtn = Button(topFrame, text="Upload Fragment", command=uploadFragment)
fragmentBtn.pack(side=LEFT)

rowNumber=Entry(topFrame)
rowNumber.insert(0, "Enter the row number")
rowNumber.pack_forget

status = Label(root, text="Loading...", relief=SUNKEN, anchor = W, bd=1)
status.pack(side=BOTTOM, fill=X)



root.mainloop()