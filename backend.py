# coding=utf-8

from Code import Status
from fileIO import *
import time


class Account():
    def __init__(self,dict):
        self.dict = {}
        self.dict["ID"] = dict["ID"]
        self.dict["name"] = dict["name"]
        self.dict["card"] = dict["card"]
        self.dict["checkInTime"] = ""
        self.dict["checkOutTime"] = ""

    def __getitem__(self, key):
        return self.dict[key]

    def __setitem__(self, key, value):
        self.dict[key] = value

    def showDetail(self):
        outputStr = ""
        for keyName in self.dict.keys():
            outputStr += "%s:%20s\t\t"%(keyName,self[keyName])
        print(outputStr)

    def checkIn(self):
        if(self["checkInTime"] == ""):
            self["checkInTime"] = time.strftime('%Y-%m-%d %H:%M:%S')
            return {
                "code":Status["success"],
                "info":self.dict
            }
        else:
            return {
                "code":Status["checked"],
                "info":self.dict
            }

    def checkOut(self):
        if(self["checkOutTime"] == ""):
            self["checkOutTime"] = time.strftime('%Y-%m-%d %H:%M:%S')
            return {
                "code":Status["success"],
                "info":self.dict
            }
        else:
            return {
                "code":Status["checked"],
                "info":self.dict
            }

class System():
    def __init__(self,recordFileName = ""):
        self.FileInfo = {
            "ACCOUNT" : {
                "fileName": "accounts.xlsx",
                "sheetName": "account",
                "tableHeadRow": "ID,name,card".split(",")
            },
            "RECORD": {
                "fileName":recordFileName,
                "sheetName": "record",
                "tableHeadRow": "ID,name,checkInTime,checkOutTime".split(",")
            }
        }

        self.accountArr = self.loadAccount()
        if recordFileName != "":
            self.FileInfo["RECORD"]["fileName"] = "%s.xlsx"%recordFileName
            self.accountArr = self.loadRecord(self.accountArr)

    def loadAccount(self):
        def creatAccountArr(rowData):
            accountData = {}
            keyArr = self.FileInfo["ACCOUNT"]["tableHeadRow"]
            for index, value in enumerate(rowData):
                keyName = keyArr[index]
                if type(value).__name__ != 'str':
                    value = int(value)
                    value = str(value)
                accountData[keyName] = str(value)
            return Account(accountData)
        fileName = self.FileInfo["ACCOUNT"]["fileName"]
        sheetName = self.FileInfo["ACCOUNT"]["sheetName"]
        accountRowDataArr = get_data(fileName, sheetName)
        accountRowDataArr.pop(0) #removeTableHead
        accountArr = list(map(creatAccountArr, accountRowDataArr))
        return accountArr

    def loadRecord(self,accountArr):
        def creatRecordArr(rowData):
            record = {}
            keyArr = self.FileInfo["RECORD"]["tableHeadRow"]
            for index, value in enumerate(rowData):
                if type(value).__name__ != 'str':
                    value = int(value)
                    value = str(value)
                keyName = keyArr[index]
                record[keyName] = value
            return record

        def load(account):
            record = list(filter(lambda record:record["ID"]==account["ID"],recordArr))[0]
            account["checkInTime"] = record["checkInTime"]
            account["checkOutTime"] = record["checkOutTime"]
            return account
        fileName = self.FileInfo["RECORD"]["fileName"]
        sheetName = self.FileInfo["RECORD"]["sheetName"]
        rocordRowDataArr = get_data(fileName, sheetName)
        rocordRowDataArr.pop(0) #removeTableHead
        recordArr = list(map(creatRecordArr,rocordRowDataArr))
        accountArr = list(map(load,accountArr))
        return accountArr

    def createEmptyRecordFile(self,recordFileName):
        self.FileInfo["RECORD"]["fileName"] = "%s.xlsx"%recordFileName
        self.saveRecord(self.accountArr)



    def saveAccount(self,accountArr):
        def getAccountRowData(account):
            dict = account.dict
            rowData = []
            keyArr = self.FileInfo["ACCOUNT"]["tableHeadRow"]
            for key in keyArr:
                value = dict[key]
                rowData.append(value)
            return rowData

        rowDataArr = list(map(getAccountRowData,accountArr))
        keyArr = self.FileInfo["ACCOUNT"]["tableHeadRow"]
        rowDataArr.insert(0,keyArr)

        fileName = self.FileInfo["ACCOUNT"]["fileName"]
        sheetName = self.FileInfo["ACCOUNT"]["sheetName"]
        save_sheet(rowDataArr,fileName,sheetName)


    def saveRecord(self,accountArr):
        def getRcardRowData(account):
            dict = account.dict
            rowData = []
            keyArr = self.FileInfo["RECORD"]["tableHeadRow"]
            for key in keyArr:
                value = dict[key]
                rowData.append(value)
            return rowData

        rowDataArr = list(map(getRcardRowData,accountArr))
        keyArr = self.FileInfo["RECORD"]["tableHeadRow"]
        rowDataArr.insert(0,keyArr)

        fileName = self.FileInfo["RECORD"]["fileName"]
        sheetName = self.FileInfo["RECORD"]["sheetName"]
        save_sheet(rowDataArr,fileName,sheetName)

    def checkIn(self,checkInCardNumber):
        findAccountArr = list(filter(lambda account:account["card"]==checkInCardNumber,self.accountArr))
        if not findAccountArr:
            return {
                "code":Status["fail"]
            }
        else:
            account = findAccountArr[0]
            res = account.checkIn()
            self.saveRecord(self.accountArr)
            return res

    def checkOut(self,checkOutCardNumber):
        findAccountArr = list(filter(lambda account:account["card"]==checkOutCardNumber,self.accountArr))
        if not findAccountArr:
            return {
                "code":Status["fail"]
            }
        else:
            account = findAccountArr[0]
            res = account.checkOut()
            self.saveRecord(self.accountArr)
            return res

    def isExist(self,keyName,checkValue):
        findAccountArr = list(filter(lambda account:account[keyName]==checkValue,self.accountArr))
        if findAccountArr:
            return True
        else:
            return False

    def updateCardNumber(self,ID,newCardNumber):
        findAccountArr = list(filter(lambda account: account["ID"] == ID, self.accountArr))
        if findAccountArr:
            account = findAccountArr[0]
            if account["card"] == newCardNumber :
                return {
                    "code":Status["success"],
                    "msg":"變更的卡後與原本相同",
                    "info":account.dict
                }
            elif self.isExist("card",newCardNumber):
                return {
                    "code":Status["fail"],
                    "msg":"此卡號已經被其他用戶使用過"
                }
            else:
                account["card"] = newCardNumber
                self.saveAccount(self.accountArr)
                return {
                    "code":Status["success"],
                    "msg":"卡號更新成功",
                    "info":account.dict
                }
        else:
            return {
                "code":Status["fail"],
                "msg":"找不到此人員資訊"
            }

        return -1
    def showAllCount(self):
        list(map(lambda account: account.showDetail(), self.accountArr))

