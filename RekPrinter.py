import os
# Import Workbook
from openpyxl import *
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
# Import GUI
from tkinter import *
from tkinter import filedialog

def getPath():
    root.filename = filedialog.askopenfilename(title = 'Select file', filetypes = (('Excel files','*.xlsx'),('All files','*.*')))
    return root.filename


def setPath(path):
    i = 0
    while os.path.isfile(path):
        if i == 0:
            path = path[:-5] + ' (1).xlsx'
        else:
            parenthesis = path.find('(')
            path = path[:parenthesis+1] + str(i) + ').xlsx'
        i += 1

    return path


def copyRow(row, startCol, endCol, sheet):
    rowSelected = []
    for j in range(startCol,endCol+1):
        rowSelected.append(sheet.cell(row = row, column = j).value)
 
    return rowSelected


def pasteRow(row, startCol, endCol, sheetReceiving, copiedData):
    countCol = 0
    for j in range(startCol,endCol+1):
        sheetReceiving.cell(row = row, column = j).value = copiedData[countCol]
        countCol += 1

        sheetReceiving.cell(row = row, column = j).border = Border(left=Side(border_style='thin', color='00000000'),
                                                                    right=Side(border_style='thin', color='00000000'),
                                                                    top=Side(border_style='thin', color='00000000'),
                                                                    bottom=Side(border_style='thin', color='00000000'))

    sheetReceiving.column_dimensions['A'].width = 6
    sheetReceiving.column_dimensions['B'].width = 6
    sheetReceiving.column_dimensions['C'].width = 15
    sheetReceiving.column_dimensions['D'].width = 39
    sheetReceiving.column_dimensions['E'].width = 9
    sheetReceiving.column_dimensions['F'].width = 6


def formatFile(path):
    wb = load_workbook(filename=path)

    orgWS = wb['Blad1']
    orgWS.delete_cols(5)
    orgWS.delete_cols(2)

    previousRow = ['Radnr', 'Företagskod', 'Företag', 'Benämning', 'Återstår antal', 'Enhet']
    i = 1
    j = 0
    for row in orgWS.rows:
        rowlist = list()
        for cell in row:
            rowlist.append(str(cell.value))
        
        if rowlist[2] != previousRow[2]:
            wb.create_sheet(rowlist[2].replace(':', ''))
            i = 1
        else:
            i += 1
        j += 1
        
        if rowlist[2] != 'Företag':
            activeWS = wb[rowlist[2].replace(':', '')]

            activeRow = copyRow(j, 1, len(rowlist), orgWS)
            pasteRow(i, 1, len(activeRow), activeWS, activeRow)

        previousRow = rowlist
    
    wb.remove(orgWS)
    wb.save(setPath(path))


def main():
    path = getPath()
    formatFile(path)


root = Tk()
root.withdraw()

if __name__ == '__main__':
    main()

