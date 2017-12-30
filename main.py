# coding=utf-8

from backend import *
from Code import Status
MODE = {
    "BREAK": -1,
    "SHOW": 1,
    "CHECKIN": 2,
    "CHECKOUT": 3,
    "RESETCARD": 4
}

def printSeperateLine(n=65):
    print('\n%s' % ("-" * n))


def print_order_list():
    print('輸入 %d => 停止程式' % MODE["BREAK"])
    print('輸入 %d => 列出表格' % MODE["SHOW"])
    print('輸入 %d => 簽到模式' % MODE["CHECKIN"])
    print('輸入 %d => 簽退模式' % MODE["CHECKOUT"])
    print('輸入 %d => 更改卡號' % MODE["RESETCARD"])

def clearWindow():
    print("\n"*30)


def showAccountArr(system):
    clearWindow()
    printSeperateLine()
    system.showAllCount()

def pause():
    printSeperateLine()
    input(" 按下 Enter鍵 回到主介面 ")

if __name__ == '__main__':
    while True:
        try :
            haveCheckingFile = input("有簽到中的檔案嗎(y/n)：")
            if haveCheckingFile == "y":
                recordFileName = input("輸入已有的檔案名稱（不用包含.xlsx）：")
                mySystem = System(recordFileName)
            else:
                recordFileName = input("準備創立新的簽到，輸入檔案名稱（不用包含.xlsx）：")
                mySystem = System()
                mySystem.createEmptyRecordFile(recordFileName)
            break
        except Exception as e:
            print(e.__str__())

    while True:
        clearWindow()
        printSeperateLine()
        print_order_list()
        try:
            order = int(input('輸入指令：'))
        except Exception:
            print("系統訊息：無法辨識的輸入")
            continue

        if order == MODE["BREAK"]:
            break
        elif order == MODE["SHOW"]:
            showAccountArr(mySystem)
            pause()
        elif order == MODE["CHECKIN"]:
            while (True):
                printSeperateLine()
                checkInNumber = input('報到模式( 輸入 0 跳出 )：')
                if checkInNumber == '0' or checkInNumber == '': break
                printSeperateLine()
                res = mySystem.checkIn(checkInNumber)
                if res['code'] == Status["success"] :
                    info = res["info"]
                    print("系統訊息：報到成功")
                    print("報到資訊：%s\t%s"%(info["ID"],info["name"]))
                    print("報到時間：%s" %info["checkInTime"])
                elif res['code']== Status["checked"]:
                    info = res["info"]
                    print("系統訊息：報到失敗（已經簽到過）")
                    print("報到資訊：%s\t%s"%(info["ID"],info["name"]))
                    print("報到時間：%s" %info["checkInTime"])
                else:
                    print("系統訊息：報到失敗（無此卡號資訊）")
        elif order == MODE["CHECKOUT"]:
            while True:
                printSeperateLine()
                checkOutNumber = input('報到模式( 輸入 0 跳出 )：')
                if checkOutNumber == '0' or checkOutNumber == '': break
                res = mySystem.checkOut(checkOutNumber)
                if res["code"] == Status["success"] :
                    info = res["info"]
                    print("系統訊息：成功報到")
                    print("報到資訊：%s\t%s"%(info["ID"],info["name"]))
                    print("報到時間：%s" %info["checkOutTime"])
                elif res['code'] == Status["checked"]:
                    info = res["info"]
                    print("系統訊息：報到失敗（已經簽到過）")
                    print("報到資訊：%s\t%s"%(info["ID"],info["name"]))
                    print("報到時間：%s" %info["checkOutTime"])
                else:
                    print("系統訊息：報到失敗（無此卡號資訊）")
        elif order == MODE["RESETCARD"]:
            showAccountArr(mySystem)
            while True:
                printSeperateLine()
                appointAccountID = input("輸入要變更卡號的人員 ID（ 如想放棄更改，請輸入 0 ）：")
                if appointAccountID == '0' or appointAccountID == '': break
                if mySystem.isExist("ID",appointAccountID):
                    newCardNumber = input("輸入新的卡號（ 如想放棄更改，請輸入 0 ）：")
                    if newCardNumber == '0' or appointAccountID == '': continue
                    res = mySystem.updateCardNumber(appointAccountID,newCardNumber)
                    showAccountArr(mySystem)
                    printSeperateLine()
                    if(res["code"]==Status["success"]):
                        info = res["info"]
                        print("系統資訊：%s\t%s" % (info["ID"], info["name"]))
                    print("系統訊息：%s" % res["msg"])
                else:
                    showAccountArr(mySystem)
                    print("無此ID資訊，請再輸入一次")
        else:
            print('無此指令')