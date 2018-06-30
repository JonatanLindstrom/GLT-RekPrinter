import os
import re
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
    path = path.split('/')
    output = ''
    for i in range(len(path) - 1):
        output += path[i] + '/'
    output += 'Kvällsbeställningar.xlsx'

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
    placedict = {'551': 'Slushbaren',
                 '552': 'Ben & Jerry\'s',
                 '553': 'Pop 3an',
                 '554': 'Pop 2an',
                 '555': 'Boardwalk Café',
                 '556': 'Glasskammaren',
                 '557': 'Coffee and Donuts',
                 '558': 'Langos',
                 '559': 'Fish & Chips',
                 '560': 'Godisfabriken',
                 '561': 'Coffeebar',
                 '562': 'Mexican Corner',
                 '563': 'Matvraket',
                 '564': 'Ham 1an',
                 '565': 'Korv 2an',
                 '566': 'Hamburger 3an',
                 '567': 'Glass & Pop 1an',
                 '568': 'Glass 2an',
                 '569': 'Gyros',
                 '570': 'Grädderiet',
                 '571': 'Honeycomb',
                 '572': 'Pizzan',
                 '573': 'Poké Bowl',
                 '574': 'Tivolisnacks',
                 '575': 'Remvagn 1',
                 '576': 'Kebaben',
                 '577': 'Kvastenkiosken',
                 '578': 'Remvagn 2',
                 '579': 'Remvagn 3',
                 '580': 'Popcorn & Cotton Candy',
                 '581': 'Korvvagn 1',
                 '582': 'Milkshakebaren',
                 '583': 'Korvvagn 3',
                 '584': 'Coca cola store',
                 '612': '1883-butiken',
                 '613': 'Tivolibutiken',
                 '623': 'Fotobutik Lustiga huset',
                 '626': 'Fotobutik Twister',
                 '700': 'Testlocation'}

    reqs = wb.sheetnames
    missing = list()
    for cc, place in placedict.items():
        if place not in reqs:
            if re.match('.* [0-9]an', place):
                missing.append([cc, place[:-2] + ':' + place[-2:]])
            else:
                missing.append([cc, place])
    
    wb.create_sheet('Saknade rekar')
    activeWS = wb['Saknade rekar']
    pasteRow(1, 4, 4, activeWS, ['Saknade rekar'])
    pasteRow(3, 3, 5, activeWS, ['Kostnadsställe', 'Enhet', 'Telefon'])
    i = 4
    for row in missing:
        pasteRow(i, 3, 4, activeWS, row)
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

