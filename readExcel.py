#run pip install openpyxl before
import openpyxl
from enum import Enum

class Table:
    wb = 0
    def __init__(self, arq):
        Table.wb = openpyxl.load_workbook(arq)
        self.sheet = Table.wb.active
        self.col = {}
        self.hydrateCols()

    def getSheet(self):
        return self.sheet

    def getColumns(self):
        return self.col

    def getLine(self, line):
        values = []
        for x in self.sheet[line]:
            values.append(x.value)
        return values
    
    def getColumn(self, col):
        values = []
        #tirar a primeira linha
        it = iter(self.sheet[col])
        next(it) 
        for x in it:
            values.append(x.value)
        return values

    def getAttr(self, attr):
        return self.getColumn(self.col[attr])

    def hydrateCols(self):
        cells = self.sheet[1]
        for x in cells:
            self.col[x.value.lower()] = x.column   

def main():
    ws = Table('teste.xlsx')
    print(ws.getAttr("email"))
    return


if __name__ == "__main__":
    main()

#change sheet:
#wb[sheet]

