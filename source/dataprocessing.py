import openpyxl
import numpy as np

# Nhap du lieu
wb = openpyxl.load_workbook("Data.xlsx")

def importData(sheet):
    return([[cell.value for cell in row] for row in sheet])


Properties = wb['Properties']
P = importData(Properties['A2:E48'])

#print(P)

def inputData():
    #Nhập file dữ liệu.
    Data = wb['Data']
    D = importData(Data['B2:O200'])
    return D



def fuzzy(Xi, input):

    if Xi == 13:
        if input.upper() == "TSG NẶNG":
            return 2
        if input.upper() == "TSG":
            return 1
        if input.upper() == "THAI BÌNH THƯỜNG":
            return 0
    else:
        for X in P:
            if Xi != X[0] :
                continue
            else:
                if input is None:
                    return None
                if float(input) >= float(X[2]) and float(input) <= float(X[3]):
                    return X[4]
                else:
                    continue

def fuzzyList(X):
    list = []
    for i in range(len(X)):
        list.append(fuzzy(i,X[i]))

    return list


def fuzzyData(D):
    newD = D
    for i in range(len(D)):
        for j in range(len(D[0])):
            X = fuzzy(j, D[i][j])
            newD[i][j] = X

    return newD



def writeToExcel(sheet, table, row):
    for x in range(row):
        for y in range(len(table[x])):
            sheet.cell(row=x + 1, column=y + 1, value=table[x][y])

def writeFuzzyDataToExcel():
    sheet = wb['fuzzyData']
    writeToExcel(sheet,fuzzyData(D),rowD)
    wb.save("Data.xlsx")
    newD = importData(sheet['A1:N199'])