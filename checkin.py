import xlsxwriter
import xlrd


def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception:
        print ('無法開啟檔案')

def get_date(sheet,r1,c1,r2,c2):
    date = [[sheet.cell_value(r1+r,c1+c) for c in range(c2-c1)] for r in range(r2 - r1)]
    return date

def read_sheet(file,sheet_name):
     r_workbook = open_excel(file)
     sheet = r_workbook.sheet_by_name(sheet_name)
     data = get_date(sheet,0,0,sheet.nrows,sheet.ncols)
     return data

def save_sheet(data,file,sheet_name):
     workbook = xlsxwriter.Workbook(file)
     worksheet = workbook.add_worksheet(sheet_name)
     row = 0
     for person in data:
          col = 0
          for item in person:
               worksheet.write(row,col,item)
               col = col + 1
          row+=1
     workbook.close()
     
def print_content(data):
     
     row = 0
     for person in data:
         for item in person:
              if(isinstance(item, float)):
                  print('%d'%item,end='\t')
              else:
                  print('%s'%item,end='\t')
                  
         print('')
         for count in range(len(person)):
              print('------------',end='')
         print('')
def Check_In(Card_Number,data,file,sheet_name):
    print('\n－－－－－－－－－－－－－－－－－－－－－－－－－－')
    status = '找不到此卡號' #預設
    row = 0
    for person in data:
        if(person[2] == Card_Number):
            print('姓名：' + person[1])
            
            if(person[len(person)-1] == ''):
                data[row][len(person)-1] = '報到成功'
                status = '報到成功'
            else:
                status = '重複報到'
            break
        row = row + 1
    print('系統訊息：',status)
    save_sheet(data,file,sheet_name);

def reset_chekin(col,data,file,sheet_name):
    for row in range(1,len(data)):
        data[row][col] = ''
    save_sheet(data,file,sheet_name);
        
def print_order_list():
    print('\n－－－－－－－－－－－－－－－－－－－－－－－－－－')
    print('輸入 1 => 列出表格')
    print('輸入 2 => 報到模式')
    print('輸入 3 => 清空最近的簽到')







data = read_sheet('test.xlsx','簽到狀況');
while(True):
    print_order_list()
    order = input('輸入指令：')
    
    if(order == '1'):
        print('\n－－－－－－－－－－－－－－－－－－－－－－－－－－')
        print_content(data)
        
    elif(order == '2'):
        while(True):
            print('\n－－－－－－－－－－－－－－－－－－－－－－－－－－')
            Check_In_Number=input('報到模式( 輸入 0 跳出 )：')
            if(Check_In_Number == '0'):
                break
            Check_In_Number = float(Check_In_Number)
            Check_In(Check_In_Number,data,'test.xlsx','簽到狀況')
            data = read_sheet('test.xlsx','簽到狀況')
        
    elif(order == '3'):
        print('\n－－－－－－－－－－－－－－－－－－－－－－－－－－')
        reset_chekin(3,data,'test.xlsx','簽到狀況')
        data = read_sheet('test.xlsx','簽到狀況')
        print('系統訊息：重置完成')
        
        
    else:
        print('無此指令')

