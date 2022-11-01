import unittest
from appium import webdriver
import unittest, time
from appium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.alert import Alert

# Get sender address from previous mail execution
sender_address = "automationtest@groupware57.hanbiro.net"

def StartWinAppDriver():
    global win_driver
    desired_caps = {}
    desired_caps["app"] = "C:\\Program Files\\Microsoft Office\\Office16\\OUTLOOK.EXE"
    win_driver = webdriver.Remote(
        command_executor='http://127.0.0.1:4723',
        desired_capabilities= desired_caps)

def AccessOutlook():
    try:
        WebDriverWait(win_driver, 10).until(EC.presence_of_element_located((By.NAME, "Inbox")))
        print("Access Outlook successfully")
        access_outlook = True
    except WebDriverException:
        print("Fail to access Outlook")
        access_outlook = False
    
    return access_outlook

def ViewMail(mail_subject):
    '''
        Wait for the email which was sent from web automation test
        Check this incoming mail if it's displayed in Inbox of Outlook
    '''

    try:
        WebDriverWait(win_driver, 60).until(EC.presence_of_element_located((By.NAME, mail_subject)))
        receive_mail = True
        print("Receive mail successfully")
    except WebDriverException:
        receive_mail = False
        print("Fail to receive mail")
    
    return receive_mail

def SendMail():
    import win32com.client as win32
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = sender_address
    mail.Subject = "mail_title"
    mail.body = "mail_content"
    mail.send

'''AccessOutlook()
ViewMail(mail_subject="test mail subject")
SendMail()'''




