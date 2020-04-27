import matplotlib.pyplot as plt
import xlrd
import numpy as np




# introducing excel file to program

class FileInput:
    #fileInput = input("Please enter excel file path")
    path = '/Users/Rishi/Desktop/module 11/Carlist.xlsx'

    excelFile = path
    file1 = xlrd.open_workbook(excelFile)
    sheet = file1.sheet_by_name('Data')
    sheet1 = file1.sheet_by_index(0)


class FileTranslate:
    sheetCol = FileInput.sheet1.ncols
    sheetRow = FileInput.sheet1.nrows

    sheetInfo = FileInput.sheet1
    excelData = [[FileInput.sheet1.cell_value(r,c) for c in range(FileInput.sheet1.ncols)] for r in range(FileInput.sheet1.nrows)]
    print("\n")
    print(excelData)



class SumVals:
    countB = 0
    billList = []
    countC = 0
    cathyList = []
    countJ = 0
    joeList = []
    countS = 0
    susanList = []

    for people in FileTranslate.excelData:
        for person in people:
            if person == 'Bill':
                countB = countB + 1
                billList.append(people[1])
            if person == 'Cathy':
                countC = countC + 1
                cathyList.append(people[1])
            if person == 'Joe':
                countJ = countJ + 1
                joeList.append(people[1])
            if person == 'Susan':
                countS = countS + 1
                susanList.append(people[1])

    billTotal = sum(billList)
    cathyTotal = sum(cathyList)
    joeTotal = sum(joeList)
    susanTotal = sum(susanList)

    print("Bills total sales:{0} \n Cathys total sales: {1} \n Joes total sales: {2} \n Susans Total sales: {3}".format(
        billTotal, cathyTotal, joeTotal, susanTotal))


class Plot:

    labels =['Bill','Cathy','Joe','Susan']
    sums = [SumVals.billTotal,SumVals.cathyTotal,SumVals.joeTotal,SumVals.susanTotal]

    yPos = np.arange(len(labels))

    plt.xticks(yPos,labels)
    plt.ylabel("Money earned (millions)")
    plt.title("Total money earned")
    plt.xlabel("Salesperson")
    plt.bar(yPos,sums,label= "$")
    plt.legend()
    plt.show()








