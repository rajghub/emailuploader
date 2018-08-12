# --------- EXTRACT DATA FROM EXCEL FILE ------------------------
import openpyxl

class accessExcel:
    book = ""
    sheet = ""
    fieldDetails = {}
    cellCoordinates = {}

# Initialisation
    def __init__(self, exlPath, sheetName):
        self.exlPath = exlPath
        self.sheetName = sheetName      
        # print(exlPath + " - " + sheetName)

# Open Excel and select sheet
    def selectSheet(self):
        self.book = openpyxl.load_workbook(self.exlPath)
        self.sheet = self.book[self.sheetName] 
        
# Access and collect required fields
    def accessFields(self, valuerow):

        # cols = self.sheet.max_column
        self.headingrow=2
        self.valuerow = valuerow
        self.fieldDetails={}
        cellDetails={}

        for i in range(1, self.sheet.max_column+1):
            
            hr=self.sheet.cell(self.headingrow, i)
            vr=self.sheet.cell(self.valuerow, i)
            self.fieldDetails[hr.value] = vr.value
            if(hr.value == "Ref No"):
                print("COORDINATE : " + hr.coordinate)
                print("COORDINATE : " + vr.coordinate)   
            if(hr.value == "Tracking Id"):
                print("COORDINATE : " + hr.coordinate)
                print("COORDINATE : " + vr.coordinate)
                
        print(self.fieldDetails)     
        print("TESTING   :   "+ fieldDetails)                         
        
        return self.fieldDetails
             
#   To modify the sheet
    def modifySheet(self, originText, textToReplace, replaceWith, cellNo):
        self.originText = originText
        self.originText = self.originText.replace(textToReplace, replaceWith)
        self.sheet[cellNo] = self.originText
        self.fieldDetails["Tracking Id"] = self.originText
        self.book.save(self.exlPath)

    def updateDocInfo(docInfoNumber):
        self.docInfoNumber=docInfoNumber
        print("Its coming here")
        print(docInfoNumber)
        self.sheet['F5'] = self.docInfoNumber
        self.book.save(self.exlPath)
