import xlsxwriter
import xlrd



def get_date(sheet, r1, c1, r2, c2):
    date = [[sheet.cell_value(r1 + r, c1 + c) for c in range(c2 - c1)] for r in range(r2 - r1)]
    return date

def open_excel(fileName):
    r_workbook = xlrd.open_workbook(fileName)
    return r_workbook

def read_sheet(r_workbook, sheet_name):
    sheet = r_workbook.sheet_by_name(sheet_name)
    rowDataArr = get_date(sheet,0, 0, sheet.nrows, sheet.ncols)
    return rowDataArr

def save_sheet(data,fileName, sheet_name):
    workbook = xlsxwriter.Workbook(fileName)
    worksheet = workbook.add_worksheet(sheet_name)
    for rowInx,rowData in enumerate(data):
        for colInx,value in enumerate(rowData):
            worksheet.write(rowInx, colInx, value)
    workbook.close()


def get_data(fileName,sheetName):
    workbook = open_excel(fileName)
    rowDataArr = read_sheet(workbook,sheetName)
    return rowDataArr