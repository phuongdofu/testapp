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
from common_functions import *

def Dictionaries():
    global circular_dict, circular_tc
    circular_dict = dict(data["circular"])
    circular_tc = dict(data["testcase_result"]["circular"])

def Circular_AccessMenu():
    PrintYellow("-------------------------------------------")
    PrintYellow("[MENU CIRCULAR]")

    access_menu = AccessGroupwareMenu(name="circular,Circular", page_xpath=circular_dict["list_footer"])
    
    return access_menu

def Circular_WriteNewCircular(user_id):
    PrintYellow("WRITE NEW CIRCULAR")

    Waits.Wait10s_ElementLoaded(circular_dict["list_footer"])

    CommonWriteItem(circular_dict["pen_button"], circular_dict["subject"], objects.hanbiro_content)
    Logging("Write circular - Click Create button / Input subject and content")

    circular_name = Functions.GetInputValue(circular_dict["subject"])
    Logging("Write circular - Circular title is: " + circular_name)

    Commands.ExecuteScript("window.scrollTo(0,-300)")

    Autocomplete_SelectRecipient(user_id , circular_dict["address_holder_recipient"])
    Logging("Write circular - Select circular recipient")

    Commands.ExecuteScript("window.scrollTo(0,-300)")

    FindPushNoti()
    Commands.ClickElement(circular_dict["send_button"])
    Logging("Write circular - Click Send button")

    try:
        Waits.Wait10s_ElementLoaded(circular_dict["list_footer"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="circular", testcase="write", msg="click save circular")
        TestCase_LogResult(**circular_tc["write"]["fail"])
    
    return circular_name

def Circular_ViewByNotification(title, counter1):
    circular_notification = Commands.FindElement(circular_dict["circular_linktext"] % title)

    counter2 = Counter_CheckCounterNumber(circular_dict["top_counter"], circular_dict["left_counter"])    
    if counter2 != counter1:
        TestCase_LogResult(**circular_tc["unread_counter"]["pass"])
    else:
        TestCase_LogResult(**circular_tc["unread_counter"]["fail"])

    circular_notification.click()
    Logging("Notification - Click to view incoming notification")

    Waits.Wait10s_ElementLoaded(circular_dict["circular_view"])
    FindPushNoti()
    Logging("Notification - Close push notification")

    if title not in Functions.GetPageSource:
        view_tc_update = circular_tc["view"]["fail"].update({"description": "Title is not displayed in view page"})
        TestCase_LogResult(**view_tc_update)

    counter3 = Counter_CheckCounterNumber(circular_dict["top_counter"], circular_dict["left_counter"])
    if counter3 != counter2:
        TestCase_LogResult(**circular_tc["read_counter"]["pass"])
    else:
        TestCase_LogResult(**circular_tc["read_counter"]["fail"])     

    Commands.Wait10s_ClickElement(circular_dict["receive_check"])
    Logging("Cicrular View - Check Received check box")

    time.sleep(1)
    comment = InputContent()
    Logging("Cicrular View - Input content for Circular Receive: " + comment)

    Commands.ClickElement(circular_dict["comment_save"])
    Logging("Circular View - Save content for Circular Receive")

    Waits.Wait10s_ElementInvisibility(circular_dict["confirm_receipt"])

    circular_content = Commands.FindElement(circular_dict["circular_content"])
    Logging("Circular View - Circular content: " + circular_content.text)
    if circular_content.text == comment:
        TestCase_LogResult(**circular_tc["comment"]["pass"])
    else:
        TestCase_LogResult(**circular_tc["comment"]["fail"])

    Commands.ClickElement(circular_dict["back_button"])
    Logging("Circular View - Back from content")

    Waits.WaitUntilPageIsLoaded(circular_dict["list_footer"])

def Circular_ViewWithoutNotification(title, counter1):
    new_circular = Waits.Wait10s_ElementLoaded(circular_dict["list_circular"] % title)
    Logging("Circular View - Click to view new unread circular")

    counter2 = Counter_CheckCounterNumber(circular_dict["top_counter"], circular_dict["left_counter"])
    if counter2[1] != counter1[1]:
        TestCase_LogResult(**circular_tc["unread_counter"]["pass"])
    else:
        TestCase_LogResult(**circular_tc["unread_counter"]["fail"])  
    
    new_circular.click()

    Waits.Wait10s_ElementLoaded(circular_dict["circular_view"])
    if title not in Functions.GetPageSource:
        circular_tc["view"]["pass"].update({"description": "Title is not displayed in view page"})
        TestCase_LogResult(**circular_tc["view"]["fail"])

    counter3 = Counter_CheckCounterNumber(circular_dict["top_counter"], circular_dict["left_counter"])
    if counter3[1] != counter2[1]:
        TestCase_LogResult(**circular_tc["read_counter"]["pass"]) 
    else:
        TestCase_LogResult(**circular_tc["read_counter"]["fail"]) 

    Commands.Wait10s_ClickElement(circular_dict["receive_check"])
    Logging("Cicrular View - Check Received check box")

    time.sleep(1)
    InputContent()
    Logging("Cicrular View - Input content for Circular Receive")

    Commands.ClickElement(circular_dict["comment_save"])
    Logging("Circular View - Save content for Circular Receive")

    Waits.Wait10s_ElementLoaded(circular_dict["circular_content"])
    circular_content = Functions.GetElementText(circular_dict["circular_content"])
    Logging("Circular View - Circular content: " + circular_content)
    
    if bool(circular_content) == True:
        TestCase_LogResult(**circular_tc["comment"]["pass"])
    else:
        TestCase_LogResult(**circular_tc["comment"]["fail"])

    Commands.ClickElement(circular_dict["back_button"])
    Logging("Circular View - Back from content")

    Waits.Wait10s_ElementLoaded(circular_dict["list_footer"])

def Circular_ViewCircular(circular_name):
    PrintYellow("CHECK NOTIFICATION AND VIEW CIRCULAR")
    '''Define if push notification can be checked and check if circular can be viewed with or without notification'''
    try:
        Waits.Wait10s_ElementLoaded(circular_dict["circular_linktext"] % circular_name)
        notification = True
    except WebDriverException:
        notification = False
    
    if notification == True:
        Logging("Circular - Notification is valid")
        TestCase_LogResult(**circular_tc["notification"]["pass"])
        Commands.ClickElement(circular_dict["circular_linktext"] % circular_name)
        Logging("View notification via notification")
    else:
        TestCase_LogResult(**circular_tc["notification"]["fail"])
        try:
            circular = Commands.ClickElement(circular_dict["circular_item"] % circular_name)
            Logging("View circular without notification")
        except WebDriverException:
            circular = Commands.FindElements(circular_dict["circular_title"])
            if int(len(circular)) > 1:
                circular_name = circular[0].text
                circular[0].click()
                Logging("Cannot find new created circular \n Click to view circular " + circular_name)
            else:
                Logging("Cannot find circular to view")
    
    Waits.Wait10s_ElementLoaded(circular_dict["circular_view"])
    FindPushNoti()

    try:
        Waits.Wait10s_ElementLoaded(circular_dict["content_view"])
        TestCase_LogResult(**circular_tc["view"]["pass"])
    except:
        TestCase_LogResult(**circular_tc["view"]["fail"])
    
    try:
        Waits.Wait10s_ElementLoaded(circular_dict["receive_check"])
        Logging("Check Received check box")
        receipt = True
    except WebDriverException:
        receipt = None
        Logging("Cannot find Received checkbox")

    if receipt == True:
        try:
            Commands.ClickElement(circular_dict["receive_check"])
            Logging("Click Confirm circular receipt")
            Waits.WaitElementLoaded(5, circular_dict["confirm_receipt"])
            Logging("Confirm circular receipt successfully")
        except WebDriverException:
            dict(circular_tc["view"]["fail"]).update({"description": "Fail to confirm receipt"})
            TestCase_LogResult(**circular_tc["comment"]["fail"])
    
    time.sleep(1)
    comment = InputContent()
    Logging("Cicrular View - Input content for Circular Receive: " + comment)
    print("comment", comment)

    Commands.ClickElement(circular_dict["comment_save"])
    Logging("Circular View - Save content for Circular Receive")

    circular_content = Functions.GetElementText(circular_dict["circular_content"])
    Logging("Circular View - Circular content: " + circular_content)
    
    if circular_content == comment:
        TestCase_LogResult(**circular_tc["comment"]["pass"])
    else:
        TestCase_LogResult(**circular_tc["comment"]["fail"])

    Commands.Wait10s_ClickElement(circular_dict["back_button"])
    Logging("Circular View - Back from content")

    Waits.WaitUntilPageIsLoaded(circular_dict["list_footer"])
    time.sleep(1)

def Circular_SearchCircular():
    PrintYellow("SEARCH CIRCULAR")
    Waits.Wait10s_ElementLoaded(circular_dict["circular_div"])

    try:
        circular_name = Functions.GetElementText(circular_dict["normal_circular"])
    except WebDriverException:
        circular_name = Functions.GetElementText(circular_dict["secure_circular"])
    
    search_dict = {
        "title": {
            "key": circular_name,
            "value": "title"
        }
    }

    creator_list = []
    senders = Commands.FindElements(circular_dict["search_sender"])
    for creator in senders:
        creator_list.append(creator.text)
    creator_list.remove("Creator")
    creator_list = list(dict.fromkeys(creator_list))
    if len(creator_list) > 1:
        search_dict["creator"] = {"key": creator_list[0], "value": "sender"}

    search_circular_dict = dict(circular_dict["search_details"])
    search_circular_dict["search_dict"] = search_dict
    SearchDetailsBySelectBox(**search_circular_dict)
    
def Circular_ValidateNextPageList():
    PrintYellow("[MENU CIRCULAR] MOVE PAGE")
    List_ValidateListMovingPage(circular_dict["list_target"], 
                                circular_dict["item_suf"], 
                                circular_dict["page_total"], 
                                circular_dict["nextpage_icon"])

def Circular_CollectListData():
    Waits.Wait10s_ElementLoaded(circular_dict["list_footer"])
    Waits.Wait10s_ElementLoaded(circular_dict["page_total"])
    try:
        list_data = CollectListData(list_footer=circular_dict["list_footer"], page_total=circular_dict["page_total"])
        circulars = list_data["total_items"]
        pages = list_data["total_pages"]
    except:
        circulars = 0
        pages = 0
    print("circulars" + str(circulars))
    print("pages" + str(pages))
    circular_list = {
        "circulars": circulars,
        "pages": pages
    }

    return circular_list

def CircularExecution(user_id):
    access_circular = Circular_AccessMenu()
    circular_name = Circular_WriteNewCircular(user_id)
    Circular_ViewCircular(circular_name)
    circular_list = Circular_CollectListData()
    Circular_SearchCircular()
    Circular_ValidateNextPageList()
    ValidateUnexpectedModal()