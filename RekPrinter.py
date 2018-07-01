import os
import re
import datetime
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
    now = datetime.datetime.fromtimestamp(os.path.getctime(path)).strftime('%Y%m%d')
    path = path.split('/')
    output = ''
    for i in range(len(path) - 1):
        output += path[i] + '/'
    output += ('Kvällsbeställningar ' + now[2:] + '.xlsx')

    i = 0
    while os.path.isfile(output):
        if i == 0:
            output = output[:-5] + ' (1).xlsx'
        else:
            parenthesis = output.find('(')
            output = output[:parenthesis+1] + str(i) + ').xlsx'
        i += 1

    return output


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


def checkReq(wb):
    placeMap = [['551', 'Slushbaren', 9505],
                ['552', 'Ben & Jerry\'s', 9506],
                ['553', 'Pop 3an', 9507],
                ['554', 'Pop 2an', 9508],
                ['555', 'Boardwalk Café', 9509],
                ['556', 'Glasskammaren', 9510],
                ['557', 'Coffee and Donuts', 9511],
                ['558', 'Langos', 9512],
                ['559', 'Fish & Chips', 9513],
                ['560', 'Godisfabriken', 9514],
                ['561', 'Coffeebar', 9515],
                ['562', 'Mexican Corner', 9516],
                ['563', 'Matvraket', 9517],
                ['564', 'Ham 1an', 9518],
                ['565', 'Korv 2an', 9519],
                ['566', 'Hamburger 3an', 9520],
                ['567', 'Glass & Pop 1an', 9521],
                ['568', 'Glass 2an', 9522],
                ['569', 'Gyros', 9523],
                ['570', 'Grädderiet', 9524],
                ['571', 'Honeycomb', 9525],
                ['572', 'Pizzan', 9526],
                ['573', 'Poké Bowl', 9557],
                ['574', 'Tivolisnacks', 9557],
                ['575', 'Remvagn 1', 9557],
                ['576', 'Kebaben', 9557],
                ['577', 'Kvastenkiosken', 9557],
                ['578', 'Remvagn 2', 9557],
                ['579', 'Remvagn 3', 9557],
                ['580', 'Popcorn & Cotton Candy', 9557],
                ['581', 'Korvvagn 1', 9557],
                ['582', 'Milkshakebaren', 9557],
                ['583', 'Korvvagn 3', 9557],
                ['584', 'Coca cola store', 9557],
                ['612', '1883-butiken', 9557],
                ['613', 'Tivolibutiken', 9557],
                ['623', 'Fotobutik Lustiga huset', 9557],
                ['626', 'Fotobutik Twister', 9557],
                ['700', 'Testlocation', 9557]]

    reqs = wb.sheetnames
    missing = list()
    for row in placeMap:
        place = row[1]
        if place not in reqs:
            if re.match('.* [0-9]an', place):
                row[1] = place[:-2] + ':' + place[-2:]
                missing.append(row)
            else:
                missing.append(row)
    
    wb.create_sheet('Saknade rekar', 0)
    activeWS = wb['Saknade rekar']
    activeWS['C1'] = 'Saknade rekar'
    activeWS['C1'].alignment = Alignment(horizontal='center')
    activeWS.merge_cells('C1:E1')

    pasteRow(3, 3, 5, activeWS, ['Kostnadsställe', 'Enhet', 'Telefon'])
    i = 5
    for row in missing:
        pasteRow(i, 3, 5, activeWS, row)
        i += 1


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
    
    checkReq(wb)

    output = setPath(path)
    wb.save(output)
    os.startfile(output)


def main():
    path = getPath()
    formatFile(path)


root = Tk()
root.withdraw()

if __name__ == '__main__':
    main()

