from selenium.common.exceptions import *
from selenium.webdriver.support.ui import *
from selenium import webdriver
import time

driver = webdriver.Chrome("C:\\Users\\cheqws133-user\\PyCharmProjects\\Drivers\\chromedriver.exe")

def Execute(keyword, object, data):
    if keyword == "Navigate":
        driver.maximize_window()
        driver.get(data)

    elif keyword == "Input":
        driver.find_element_by_xpath(object).send_keys(data)

    elif keyword == "Click":
        driver.find_element_by_xpath(object).click()


def Logout():
    driver.get("https://www.phptravels.net/account/logout/")