import xlrd
import openpyxl

MAX = 7180

class ExtractData:
    def __init__(self):
        self.mailBody = None
        
    def readExcel(self, sheet, row):
        self.mailBody = sheet.cell_value(row, 2)

    def getFirstMail(self):
        fromLastIndex = self.mailBody.rfind('From:')
        if fromLastIndex == -1:
            return self.mailBody
        return self.mailBody[fromLastIndex:]
    

    def writeOnExcel(self, sheet, row1):
        # print(self.getFirstMail())
        c1 = sheet.cell(row = row1+1, column = 4)
        c1.value = self.getFirstMail()

    def compareData(self):
        loc = r'C:\Workspace02\GetOriginalMail\New email body.xlsx'
        wb = xlrd.open_workbook(loc)
        wb1 = openpyxl.load_workbook(loc)
        sheet_rd = wb.sheet_by_index(0)
        sheet_wrt = wb1.get_sheet_by_name('Dataset')
        for index in range(1, MAX):
            self.__init__()
            self.readExcel(sheet_rd, index)
            self.writeOnExcel(sheet_wrt, index)

        wb1.save(loc)

obj = ExtractData()
obj.compareData()