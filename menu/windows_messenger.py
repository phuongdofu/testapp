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


# Start the web driver
service = webdriver.chrome.service.Service("D:\\PhuongDofu\\chromedriver_talk.exe")
service.start()

# start the app
driver = webdriver.remote.webdriver.WebDriver(
    command_executor=service.service_url,
    desired_capabilities={
        'browserName': 'chrome',
        'goog:chromeOptions': {
            'args': ['develop_mode'],
            'binary': 'C:\\Users\\Hanbiro\\AppData\\Local\\Programs\\hanbiro-talk\\HanbiroTalk2.exe',
            'extensions': [],
            'windowTypes': ['webview']},
        'platform': 'ANY',
        'version': ''},
    browser_profile=None,
    proxy=None,
    keep_alive=False)


WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//input[@id='domain']")))

domain = driver.find_element_by_xpath("//input[@id='domain']")
if bool(domain.get_attribute("value")) == True:
    domain.clear()
    time.sleep(1)
domain.send_keys("qa.hanbiro.net")

time.sleep(1)

user_id = driver.find_element_by_xpath("//input[@id='userid']")
if bool(user_id.get_attribute("value")) == True:
    user_id.clear()
    time.sleep(1)
user_id.send_keys("automationtest")

time.sleep(1)

driver.find_element_by_xpath("//input[@id='password']").send_keys("automationtest1!")

time.sleep(1)

driver.find_element_by_xpath("//span[text()='Sign In']").click()
