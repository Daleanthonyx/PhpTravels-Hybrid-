import openpyxl
import ExcelFunc

Path1 = "C:\\Users\\cheqws133-user\\PyCharmProjects\\Data\\PhpTravels Test Case v3.xlsx"
Path2 = "C:\\Users\\cheqws133-user\\PyCharmProjects\\Data\\PhpTravels Objects v3.xlsx"

def FindObject(objectName, sheetName):
    PathExcelFile = Path2
    rowCount = ExcelFunc.getRowCount(PathExcelFile, sheetName)+1
    for i in range (2, rowCount):
        objectName2 = ExcelFunc.readData(PathExcelFile, sheetName, i, 1)

        if objectName == objectName2:
            objectPath = ExcelFunc.readData(PathExcelFile, sheetName, i, 2)
            return objectPath

def CountCase(sheetName, rowCount, TCID):
    PathExcelFile = Path1
    rowCount = ExcelFunc.getRowCount(PathExcelFile, sheetName) + 1
    counter = 1
    for i in range (2, rowCount):
        temp_TCID = ExcelFunc.readData(PathExcelFile, sheetName, i, 1)
        if temp_TCID == TCID:
            counter += 1

    return counter

def FindRow(sheetName, rowCount, TCID):
    PathExcelFile = Path1
    rowCount = ExcelFunc.getRowCount(PathExcelFile, sheetName) + 1
    counter = 1
    for i in range (1, rowCount):
        temp_TCID = ExcelFunc.readData(PathExcelFile, sheetName, i, 1)
        if temp_TCID == TCID:
            return counter

        counter += 1