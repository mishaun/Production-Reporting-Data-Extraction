# -*- coding: utf-8 -*-
"""
Created on Fri Feb 21 11:15:12 2020

@author: -
"""
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

opIDs = {
        "MPLP": 521516,
        "MEC": 521539,
        "MOP": 521542
        }

monthDict={'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6, 'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12}
filepath = os.path.dirname(__file__)

def getPreviousMonthPR(operatorName, month, year):
    #driver will be used based on operating system - windows or mac
    try:
        driver = webdriver.Chrome(filepath + "/chromedriver.exe")
    except:
        driver = webdriver.Chrome(filepath + "/chromedriver")
    
    driver.get("http://webapps.rrc.texas.gov/PR/initializePublicQueriesMenuAction.do")
    
    #radio button clicking oil and gas
    driver.find_element_by_xpath("/html/body/table[4]/tbody/tr/td[3]/table/tbody/tr/td/form/table[2]/tbody/tr[3]/td[1]/input[3]").click()
    opnum = driver.find_element_by_name("operatorNo")
    opnum.click()
    opnum.send_keys(opIDs[operatorName])
    
    repSelect = driver.find_element_by_name("month")
    repSelect.click()
    
    #checking to see if report month passed in is january - if it is, we need to do 
    #a special block of code to set the previous month to december and subract report year 1 year back
    if month == "Jan":
        previousMonth = "Dec"
        year = year - 1
    #if current report month is not january, then get the previous month's value, and find the month abbreviation
    else:
        #takes current month's  numberical value from dictionary, and subracts by 1 to get previous month numerican value
        previousMonthVal = monthDict[month] - 1
        
        #The previous month value will be found in month dictionary and return its key to get previous month string name
        #converting result of filtered month dictionary to a list and grabbing the first element in the tuple - which contains the month
        #evaluating the second element in tuple - which is why it is x[1]
        previousMonth = list(filter(lambda x: x[1] == previousMonthVal, monthDict.items()))[0][0]
    
    #rrc website will then be inputed the previous month's name
    repSelect.send_keys(previousMonth)
    repSelect.click()
    
    repSelect = driver.find_element_by_name("year")
    repSelect.click()
    repSelect.send_keys(year)
    repSelect.click()
    
    #submit query
    driver.find_element_by_xpath("/html/body/table[4]/tbody/tr/td[3]/table/tbody/tr/td/form/table[2]/tbody/tr[3]/td[5]/input").click()
    
    #prevents from download button not being found
    time.sleep(3)
    
    #download button for csv of previous month
    driver.find_element_by_xpath("/html/body/table[4]/tbody/tr[1]/td[3]/table/tbody/tr/td/div/table/tbody/tr/td/form/table[3]/tbody/tr[2]/td[2]/input").click()
