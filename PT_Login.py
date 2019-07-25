from selenium.common.exceptions import *
from selenium import webdriver
import Perform_Keyword
import Locator
import ExcelFunc
import openpyxl
import time

PathExcelFile = "C:\\Users\\cheqws133-user\\PyCharmProjects\\Data\\PhpTravels Test Case v3.xlsx"
sheetName = "Login"

def Execute(start, testCount, TCID):
    for i in range (start, testCount+start):
        if ExcelFunc.readData(PathExcelFile, sheetName, i, 1) == TCID:
            keyword = ExcelFunc.readData(PathExcelFile, sheetName, i, 4)
            objectName = ExcelFunc.readData(PathExcelFile, sheetName, i, 5)
            object = Locator.FindObject(objectName, sheetName)
            data = ExcelFunc.readData(PathExcelFile, sheetName, i, 6)

            print(keyword, objectName)
            Perform_Keyword.Execute(keyword, object, data)

