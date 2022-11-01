import time, sys, unittest, random, json, requests, openpyxl, testlink
from datetime import datetime
from selenium import webdriver
from random import randint
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.alert import Alert

from common_functions import bcolors, objects, data, Logging, ValidateFailResultAndSystem, TestLinkResult_Pass, TestLinkResult_Fail, Groupware_CheckWorkingTimeInSidebar

def ValidateUserWorkingTime(driver):
    try:
        driver.find_element_by_xpath("//button[contains(@class, 'close-punch-btn')]").click()
        Logging("Close timecard popup")
    except WebDriverException:
        Logging("Cannot find timecard popup")

    Groupware_CheckWorkingTimeInSidebar(driver)

def ValidateUserSettings(driver):
    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//a[@title='Settings']")))
        driver.find_element_by_xpath("//a[@title='Settings']").click()
        Logging("Access User Settings page")
    except WebDriverException:
        driver.find_element_by_xpath("//a[contains(@class, 'open-sidebar')]").click()
        Logging("Open Right Sidebar")

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@title='Settings']")))

        time.sleep(1)
        
        driver.find_element_by_xpath("//a[@title='Settings']").click()
        Logging("Access User Settings page")

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[@class='menu-text' and contains(., 'General Settings')]")))

    driver.find_element_by_xpath("//span[@class='menu-text' and contains(., 'General Settings')]").click()
    Logging("Access User Setting page")

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[@ng-show='activeTab.general']")))

    selected_language = driver.find_element_by_xpath("//select[@ng-model='formData4.lang']/option[@selected='selected']")
    if selected_language.text == "English":
        Logging("English is being used")
    else:
        Select(driver.find_element_by_xpath("//select[@ng-model='formData4.lang']")).select_by_visible_text("English")
        Logging("Select language English")

    try:
        timezone_selected = driver.find_element_by_xpath("//div[@ng-bind='$select.selected.name']")
        Logging("current timezone" + timezone_selected.text)
        if "(GMT +0700)" in str(timezone_selected.text):
            Logging("Timezone can be used")
        else:
            driver.find_element_by_xpath("//a[contains(@class, 'select2-choice')]").click()
            Logging("Click timezone selectbox")

            timezone_no = int(len(driver.find_elements_by_xpath("//ul[@id='ui-select-choices-0']/li[starts-with(@id, 'ui-select-choices-row')]")))
            x = 0
            for x in range(1, timezone_no):
                x += 1
                timezone = driver.find_element_by_xpath("//ul[@id='ui-select-choices-0']/li[starts-with(@id, 'ui-select-choices-row')]" + "[" + str(x) + "]/div/span")
                if "(GMT +0700)" in str(timezone.text):
                    timezone.click()
                    Logging("Select timezone")
                    break
    except WebDriverException:
        Select(driver.find_element_by_xpath("//select[@ng-model='formData4.timeZone']")).select_by_value("string:Asia/Saigon")
        Logging("Select timezone")

    Select(driver.find_element_by_xpath("//select[@ng-model='formData4.dateFormat']")).select_by_value("string:Y/m/d H:i:s")
    Logging("Select date format")
    
    driver.find_element_by_xpath("//button[@ng-click='generalProcess()']").click()
    Logging("Save user general settings")

    time.sleep(2)

def UserLogIn(driver, domain_name, userid, userpw):
    Logging(bcolors.WARNING + "-------------------------------------------" + bcolors.ENDC)
    Logging(bcolors.WARNING + "[TEST CASE] User LogIn" + bcolors.ENDC)

    driver.get(domain_name + "/sign")
    Logging("Log in - Access login page")

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "iframeLoginPassword")))

    input_id = driver.find_element_by_xpath("//input[@id='log-userid']")
    input_id.send_keys(userid)
    if input_id.get_attribute("value") != userid:
        input_id.clear()
        input_id.send_keys(userid)
    Logging("Log in - Input valid ID")

    frame_element = driver.find_element_by_id("iframeLoginPassword")
    driver.switch_to.frame(frame_element)
    input_pw = driver.find_element_by_id("p")
    input_pw.send_keys(userpw)
    if input_pw.get_attribute("value") != userpw:
        input_pw.clear()
        input_pw.send_keys(userpw)
    driver.switch_to.default_content()
    Logging("Log in - Input valid password")

    driver.find_element_by_id("btn-log").send_keys(Keys.RETURN)
    Logging("Log in - Click Submit button")

    login_status = []

    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'alert-warning')]")))
        print("Fail to log in" + objects.testcase_fail)
        login_status.append(False)
    except WebDriverException:
        print("Log in successfully")
        print("Log in successfully" + objects.testcase_pass)
        login_status.append(True)

    #print(login_status[0])
    if login_status[0] == True:
        try:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//span[@class='infobox-text']")))
            
            try:
                WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.XPATH, data["common"]["loading_dialog"])))
            except WebDriverException:
                pass
            
            time.sleep(3)

            ValidateUserWorkingTime(driver)

            ValidateUserSettings(driver)
        except WebDriverException:
            print("Fail to access Groupware after log in")
            login_status[0] = False

    return login_status

def UserLogIn_Quick(driver, domain_name, userid, userpw):
    Logging(bcolors.WARNING + "-------------------------------------------" + bcolors.ENDC)
    Logging(bcolors.WARNING + "[TEST CASE] User LogIn" + bcolors.ENDC)

    driver.get(domain_name + "/sign")
    Logging("Log in - Access login page")

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "iframeLoginPassword")))

    input_id = driver.find_element_by_xpath("//input[@id='log-userid']")
    input_id.send_keys(userid)
    if input_id.get_attribute("value") != userid:
        input_id.clear()
        input_id.send_keys(userid)
    Logging("Log in - Input valid ID")

    frame_element = driver.find_element_by_id("iframeLoginPassword")
    driver.switch_to.frame(frame_element)
    input_pw = driver.find_element_by_id("p")
    input_pw.send_keys(userpw)
    if input_pw.get_attribute("value") != userpw:
        input_pw.clear()
        input_pw.send_keys(userpw)
    driver.switch_to.default_content()
    Logging("Log in - Input valid password")

    driver.find_element_by_id("btn-log").send_keys(Keys.RETURN)
    Logging("Log in - Click Submit button")

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//span[@class='infobox-text']")))
            
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, data["common"]["loading_dialog"])))
    except WebDriverException:
        pass
