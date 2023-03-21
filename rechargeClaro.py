import time
import pandas
import numpy as np
from xml.etree.ElementPath import xpath_tokenizer
from selenium import webdriver
from openpyxl import load_workbook
from selenium.common.exceptions import NoSuchElementException, NoSuchFrameException

link = "https://atiendo.claro.com.co/pretups/"
link2 = "https://atiendo.claro.com.co/pretups/c2sRechargeAction.do?method=c2sRechargeAuthorize&amp;moduleCode=C2STRF"
user_key = "cotalvaro"
password_key = "Team08+"
pin = "2345"

def openChrome(link):
    options =  webdriver.ChromeOptions()
    browser = webdriver.Chrome(chrome_options= options)
    browser.get(link)
    return browser

def excel():
    file = "lineas.xlsx"
    fileExcel = pandas.read_excel(file)
    numbers = np.asarray(fileExcel)
    return numbers

def deleteCeldExcel(celd):
    file = "lineas.xlsx"
    workbook = load_workbook(file)
    sheet = workbook.active
    sheet['A'+str(celd)] = ""
    sheet['B'+str(celd)] = ""
    workbook.save(filename=file) 

def insert(by, str, text, browser):
    if by == "xpath": find = browser.find_element_by_xpath(str)
    elif by == "id": find = browser.find_element_by_id(str)
    elif by == "name": find = browser.find_element_by_name(str)
    else: find =None
    if find is not None:
        find.send_keys(text)
    

def click(by, str, browser):
    if by == "xpath": find = browser.find_element_by_xpath(str)
    elif by == "id": find = browser.find_element_by_id(str)
    elif by == "name": find = browser.find_element_by_name(str)
    else: find =None
    if find is not None:
        find.click()

def login(user_key, password_key, browser):    
    insert("id", "loginID", user_key, browser)
    insert("id", "password", password_key, browser)
    time.sleep(1)
    click("name", "submit1", browser)

def recharge(number,amount,pin, browser):
    insert("name", "subscriberMsisdn", number, browser)
    insert("name", "amount", amount, browser)
    insert("name", "pin", pin, browser)
    click("name", "btnSubmit", browser)
    time.sleep(0.5)
    click("name", "btnSubmit", browser)
    time.sleep(0.5)
    click("name", "btnBack", browser)
    time.sleep(0.5)

def checklink(browser):
    try:
        browser.switch_to.frame("mainFrame")
        elementExist = browser.find_element_by_name("subscriberMsisdn")
    except NoSuchElementException:
        browser.quit()
        openChrome(link)
        login("ooo", "Team09+", browser)
        time.sleep(2)
    except NoSuchFrameException:
        browser.quit()
        openChrome(link)
        login(user_key, "Team09+", browser)
        time.sleep(2)


def run():
    browser = openChrome(link)
    time.sleep(1)
    try:
        login(user_key, password_key, browser)
    except NoSuchElementException:
        browser.get(link)
        time.sleep(1)
    time.sleep(3)
    numbers = excel()
    celd = 2
    try:
        browser.switch_to.frame("mainFrame")
    except NoSuchFrameException:
        browser.get(link)
        login(user_key, password_key, browser)
        time.sleep(1)
    for i in numbers:
        number = str(i[0])
        amount = str(i[1])
        try:
            recharge(number, amount, pin, browser)
            print(number, amount)
            print("----------------------------------------")
            deleteCeldExcel(celd)
            celd+=1
            time.sleep(0.5)
        except NoSuchElementException:
            pass
        
    


run()