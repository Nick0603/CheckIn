# coding=utf-8
import xlsxwriter
import xlrd
import time

MODE = {
    "PRINT" : 1,
    "CHECKIN" : 2,
    "CHECKOUT" : 3
}

SHEET = {
    "ACCOUNT":"account"
}
def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception:
        print ('無法開啟檔案')

def get_date(sheet,r1,c1,r2,c2):
    date = [[sheet.cell_value(r1+r,c1+c) for c in range(c2-c1)] for r in range(r2 - r1)]
    return date

def read_sheet(r_workbook,sheet_name):
     sheet = r_workbook.sheet_by_name(sheet_name)
     data = get_date(sheet,0,0,sheet.nrows,sheet.ncols)
     return data

def save_sheet(file,data1,sheet_name1,data2,sheet_name2):
    workbook = xlsxwriter.Workbook(file)

    worksheet = workbook.add_worksheet(sheet_name1)
    row = 0
    for person in data1:
        col = 0
        for item in person:
            worksheet.write(row,col,item)
            col+=1
        row+=1

    worksheet = workbook.add_worksheet(sheet_name2)
    row = 0
    for person in data2:
        col = 0
        for item in person:
            worksheet.write(row,col,item)
            col+=1
        row+=1

    workbook.close()

def print_content(sheet):
    def alignment(str, spaceN, align='left'):
        length = 0;
        for char in str:
            encodeLen = len(char.encode('utf-8'))
            if encodeLen > 2 :
                length += 2
            else:
                length += encodeLen
        spaceN = spaceN - length if spaceN >= length else 0
        if align == 'left':
            str = str + ' ' * spaceN
        elif align == 'right':
            str = ' ' * spaceN + str
        elif align == 'center':
            str = ' ' * (spaceN // 2) + str + ' ' * (spaceN - spaceN // 2)
        return str

    def toString(data):
        if isinstance(data,float):
            data = "%s"%data
        return alignment(data,10,"center")

    def makeSeparateLine(data):
        return alignment("----------",10,"center")

    ouputString = ""
    for row in sheet:
        rowData = map(toString,row)
        dataString = "|".join(rowData)
        separateLineArr =  map(makeSeparateLine,row)
        separateLine = " ".join(separateLineArr)

        ouputString += dataString
        ouputString += "\n"
        ouputString += separateLine
        ouputString += "\n"
    print(ouputString)

def Check_In(Card_Number,accountData,checkInData,mode):
    status = '找不到此卡號' #預設
    nameIndex = 1
    cardIndex = 2
    card2Index = 0
    if(mode == MODE["CHECKIN"]):
        checkMode = "簽到"
        statusIndex = 2
        timeIndex = 3
    else:
        checkMode = "簽退"
        statusIndex = 4
        timeIndex = 5
    row = 0
    for person in accountData:
        card1 = person[cardIndex]
        card2 = person[card2Index]
        if  card1 == Card_Number or str(card1) == Card_Number or str(card2) == Card_Number or card2 == Card_Number  :
            if(checkInData[row][statusIndex] == ''):
                if( card1 == Card_Number or str(card1) == Card_Number  ):
                    checkInData[row][statusIndex] = checkMode + '成功'
                else:
                    if(card1 == ""):
                        newCardNumber = input("設定卡號：")
                        if(newCardNumber != ""):
                            accountData[row][cardIndex] = newCardNumber;
                            status = "卡號更新成功，請重新報到"
                            break;
                    checkInData[row][statusIndex] = checkMode + '成功（透過備用卡號）'
                checkInData[row][timeIndex] = time.strftime('%Y-%m-%d %H:%M:%S')
                status = checkMode + '成功'
            else:
                status = '重複' + checkMode + "\n" + "已於 " + checkInData[row][timeIndex]  + " " + checkMode
            break
        row = row + 1
    print('系統時間：', time.strftime('%Y-%m-%d %H:%M:%S'))
    if(status == "找不到此卡號"):
        print('系統訊息：',status)
    else:
        print('系統訊息：',person[nameIndex],' ',status)

    return {"checkInData":checkInData , "accountData":accountData}

def printSeperateLine(n):
    print('\n%s' % ("-"*n) )

def print_order_list():
    print('輸入 %d => 列出表格'%MODE["PRINT"])
    print('輸入 %d => 簽到模式'%MODE["CHECKIN"])
    print('輸入 %d => 簽退模式'%MODE["CHECKOUT"])

def verifyInput(fileName,sheetName):
    try:
        spreadsheets = xlrd.open_workbook(fileName)
    except Exception:
        print ('無法開啟檔案')
        return False
    try:
        info = read_sheet(spreadsheets,SHEET["ACCOUNT"]);
    except Exception:
        print ('已開啟檔案，找不到名為 %s 的資料表'%SHEET["ACCOUNT"])
        return False
    try:
        checkInData = read_sheet(spreadsheets,sheetName);
    except Exception:
        print ('已開啟檔案，找不到名為%s的資料表',sheetName);
        return False
    return True

while True:
    fileName = input('輸入檔案名稱(檔名結尾為 .xlsx ):')
    sheetName = input("輸入簽到工作表名稱:");
    if verifyInput(fileName,sheetName):break


while True :
    printSeperateLine(65)
    print_order_list()
    try:
        order = int(input('輸入指令：'))
    except Exception:
        print("系統訊息：無法辨識的輸入")
        continue
    if(order == MODE["PRINT"]):
        spreadsheets = open_excel(fileName)
        checkInData = read_sheet(spreadsheets, sheetName)
        print_content(checkInData)
    elif(order == MODE["CHECKIN"]):
        while(True):
            Check_In_Number=input('報到模式( 輸入 0 跳出 )：')
            if Check_In_Number == '0':break
            printSeperateLine(65)
            spreadsheets = open_excel(fileName)
            accountData = read_sheet(spreadsheets, SHEET["ACCOUNT"])
            checkInData = read_sheet(spreadsheets, sheetName)
            result  = Check_In(Check_In_Number,accountData,checkInData,MODE["CHECKIN"])
            checkInData = result["checkInData"];
            accountData = result["accountData"];
            save_sheet(fileName,accountData,SHEET["ACCOUNT"],checkInData,sheetName)
            printSeperateLine(65)
    elif(order == MODE["CHECKOUT"]):
        while (True):
            Check_In_Number = input('報到模式( 輸入 0 跳出 )：')
            if Check_In_Number == '0': break
            printSeperateLine(65)
            spreadsheets = open_excel(fileName)
            accountData = read_sheet(spreadsheets, SHEET["ACCOUNT"])
            checkInData = read_sheet(spreadsheets, sheetName)
            result = checkInData = Check_In(Check_In_Number,accountData,checkInData,MODE["CHECKIN"])
            checkInData = result["checkInData"];
            accountData = result["accountData"];
            save_sheet(fileName,accountData,SHEET["ACCOUNT"],checkInData,sheetName)
            printSeperateLine(65)
    else:
        print('無此指令')

