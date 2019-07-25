from selenium.common.exceptions import *
from selenium import webdriver
import PT_Login
import PT_Booking
import Perform_Keyword
import Locator
import ExcelFunc
import openpyxl
import time

PathExcelFile = "C:\\Users\\cheqws133-user\\PyCharmProjects\\Data\\PhpTravels Test Case v3.xlsx"
sheetName = "Suites"
suiteCount = ExcelFunc.getRowCount(PathExcelFile, sheetName)+1

for i in range (2, suiteCount):
    if ExcelFunc.readData(PathExcelFile, sheetName, i, 5) == "Yes":
        TCID = ExcelFunc.readData(PathExcelFile, sheetName, i, 2)
        sheetName = ExcelFunc.readData(PathExcelFile, sheetName, i, 4)
        rowCount = ExcelFunc.getRowCount(PathExcelFile, sheetName)
        testCount = Locator.CountCase(sheetName, rowCount, TCID)
        start = Locator.FindRow(sheetName, rowCount, TCID)

        if sheetName == "Login":
            PT_Login.Execute(start, testCount, TCID)

        if sheetName == "Booking":
            PT_Booking.Execute(start, testCount, TCID)

    sheetName = "Suites"



