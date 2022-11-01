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
    global resource_dict, resource_tc
    resource_dict = dict(data["resource"])
    resource_tc = dict(data["testcase_result"]["resource"])

def Resource_AccessMenu(domain_name):
    PrintYellow("-------------------------------------------")
    PrintYellow("[MENU RESOURCE]")

    reservation_links = {
        "qa.hanbiro.net": "/resource/list/476_0/",
        "qavn.hanbiro.net": "/resource/list/10_0/",
        "qa1.hanbiro.net": "/resource/list/207_0/",
        "dofu.hanbiro.net": "/resource/list/6_0/",
        "groupware57.hanbiro.net": "/resource/list/18_0/",
        "global3.hanbiro.com": "/resource/list/409_102/",
        "tg01.hanbiro.net": "/resource/list/3_0/"
    }
    
    request_resource = None
    resource_system = None

    current_url = DefineCurrentURL()
    for reservation_folder in reservation_links.keys():
        if str(reservation_folder) in current_url:
            link = str(reservation_links[reservation_folder])
            Commands.NavigateTo(domain_name + link)

            Waits.WaitElementLoaded(20, resource_dict["view_container"])
            time.sleep(2)
            Waits.WaitUntilPageIsLoaded(None)
            request_resource = True
            break
        else:
            link = None
    
    if link == None:
        AccessGroupwareMenu(name="resource,Resource", page_xpath=resource_dict["header_title"])

        Waits.WaitElementLoaded(20, resource_dict["view_container"])
        time.sleep(2)
        Waits.WaitUntilPageIsLoaded(None)
        
        Waits.Wait10s_ElementLoaded(resource_dict["resource_arrow"])
        main_tree = Functions.GetElementAttribute(resource_dict["resource_li"], "class")
        if "open" not in main_tree:
            Commands.ClickElement(resource_dict["resource_arrow"])
            Logging("Open Resource menu tree")
        
        Waits.Wait10s_ElementLoaded(resource_dict["resource_comp"])
        time.sleep(1)

        i=0
        added_tree_xpath = ""
        for i in range(1,5):
            i+=1
            print("added_tree_xpath: " + added_tree_xpath)
            try:
                Waits.Wait10s_ElementLoaded(resource_dict["category_name"] % added_tree_xpath)
                time.sleep(1)
                folder_attr = Functions.GetElementAttribute(resource_dict["category_li"] % added_tree_xpath, "class")
                if "open" not in folder_attr:
                    Commands.ClickElement(resource_dict["category_arrow"] % added_tree_xpath)
                added_tree_xpath = added_tree_xpath + "/ul/li"
                Waits.Wait10s_ElementLoaded(resource_dict["category_name"] % added_tree_xpath)
            except WebDriverException:
                Commands.ClickElement(resource_dict["category_name"] % added_tree_xpath)
                resource_type = Functions.GetElementAttribute(resource_dict["category_a"] % added_tree_xpath, "class")
                if "file_lock" in resource_type:
                    resource_system = "Permission"
                else:
                    resource_system = "Reservation"
                Waits.WaitElementLoaded(20, resource_dict["view_container"])
                Waits.WaitUntilPageIsLoaded(None)
                time.sleep(2)
                request_resource = True
                break
    
    request_data = {
        "request_resource": request_resource,
        "resource_system": resource_system
    }

    return request_data
            
def Resource_Reserve():
    PrintYellow("[RESOURCE] REQUEST RESOURCE IN RESERVATION SYSTEM")

    Commands.FindElement_ByCSS("thead .fc-today").click()
    Logging("Click today grid on calendar")

    Waits.WaitElementLoaded(20, data["editor"]["tox_iframe"])

    reservation = []
    resource_name = "Resource request is generated at %s" % objects.date_time
    Commands.InputElement(resource_dict["title"], resource_name)
    reservation_name = Functions.GetInputValue(resource_dict["title"])
    Logging("Input title: " + reservation_name)

    Commands.ClickElement(resource_dict["all_day"])
    Logging("Select All Day")

    Commands.ClickElement(resource_dict["calendar_start"])
    Logging("Click Calendar button to change start date")

    Commands.Wait10s_ClickElement(resource_dict["calendar_active"] + "/following-sibling::td")
    Logging("Select next day from today")

    time.sleep(1)

    start_date_text = Functions.GetInputValue(resource_dict["start_date"])
    Logging("Start date: " + start_date_text)

    end_date_text = start_date_text = Functions.GetInputValue(resource_dict["end_date"])
    Logging("End date: " +  end_date_text)

    if start_date_text == end_date_text:
        Logging(" Start date and end date are same")
    else:
        resource_tc["write"]["fail"].update({"description": "Start date and end date are not same"})
        TestCase_LogResult(**resource_tc["write"]["fail"])
    
    time.sleep(1)
    
    Commands.ClickElement("//button[contains(., 'Mail Content')]")
    Logging("Insert mail content")

    Waits.Wait10s_ElementLoaded("//button[contains(., 'Mail Content')]/preceding-sibling::han-editor[contains(@class, 'ng-not-empty')]")
    Logging("Mail content is inserted")

    time.sleep(1)

    Commands.ClickElement(resource_dict["save_reservation"])
    Waits.Wait10s_ElementLoaded(data["common"]["loading_dialog"])

    try:
        Waits.WaitElementLoaded(2, resource_dict["warning"])
        Commands.ClickElement(resource_dict["close_warning"])
        ValidateFailResultAndSystem("Fail to reserve resource")
        Logging(objects.testcase_fail)
    except WebDriverException:
        Waits.Wait10s_ElementLoaded(resource_dict["fc_title"])
        
        time.sleep(1)
        
        try:
            resource_name = "Resource request is generated at %s" % objects.date_time
            Commands.FindElement(resource_dict["resource_title"] % resource_name)
            TestCase_LogResult(**resource_tc["write"]["pass"])
        except WebDriverException:
            TestCase_LogResult(**resource_tc["write"]["fail"])

    reservation_data = [reservation_name, objects.hanbiro_content, start_date_text, end_date_text]
    reservation.extend(reservation_data)

    return reservation

def Resource_View(hanbiro_title, hanbiro_content):
    PrintYellow("[RESOURCE] VIEW RESERVATION")
    time.sleep(1)
    Waits.Wait10s_ElementLoaded(resource_dict["fc_title"])
    time.sleep(1)
    Commands.ClickElement(resource_dict["resource_title"] % hanbiro_title)
    try:
        Waits.Wait10s_ElementLoaded(resource_dict["content_title"])
        TestCase_LogResult(**resource_tc["view"]["pass"])

        content_title = Commands.FindElement(resource_dict["content_title"])
        if content_title.text != hanbiro_title:
            resource_tc["view"]["fail"].update({"description": "Cannot find title in view page"})
            TestCase_LogResult(**resource_tc["view"]["fail"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="resource", testcase="view", msg="click view resource")
        TestCase_LogResult(**resource_tc["view"]["fail"])

def Resource_Edit(start_date, end_date):
    PrintYellow("[RESOURCE] EDIT RESERVATION")
    
    Commands.ClickElement(resource_dict["modify_button"])
    Logging("Resource View - Click Modify button")

    Waits.Wait10s_ElementLoaded(data["editor"]["tox_iframe"])
    time.sleep(1)
    Commands.InputElement(resource_dict["title"], objects.title_edit)

    start_date1_text = Functions.GetInputValue(resource_dict["start_date"])
    Logging("Start date: " + start_date1_text)

    end_date1_text = Functions.GetInputValue(resource_dict["end_date"])
    Logging("End date: " +  end_date1_text)

    if start_date1_text != start_date:
        resource_tc["write"]["fail"].update({"testcase": "Edit Request"})
        resource_tc["write"]["fail"].update({"description": "Start date is different when edit"})
        TestCase_LogResult(**resource_tc["write"]["fail"])

    if end_date1_text != end_date:
        resource_tc["write"]["fail"].update({"testcase": "Edit Request"})
        resource_tc["write"]["fail"].update({"description": "End date is different when edit"})
        TestCase_LogResult(**resource_tc["write"]["fail"])

    Commands.ClickElement(resource_dict["save_reservation"])
    try:
        Waits.WaitElementLoaded(2, resource_dict["warning"])
        Commands.ClickElement(resource_dict["close_warning"])
        resource_tc["write"]["fail"].update({"testcase": "Edit Request"})
        TestCase_LogResult(**resource_tc["write"]["fail"])
    except WebDriverException:
        Waits.Wait10s_ElementLoaded(resource_dict["fc_title"])

    return objects.title_edit

def Resource_Cancel(title):
    PrintYellow("[RESOURCE] CANCEL RESERVATION")
    Waits.WaitUntilPageIsLoaded(None)

    current_url = DefineCurrentURL()
    if "/list/" in current_url:
        Waits.Wait10s_ElementLoaded(resource_dict["fc_title"])
        Commands.Wait10s_ClickElement(resource_dict["resource_title"] % objects.title_edit)
        Logging("Resource list - Click on resource with updated title")

        Waits.Wait10s_ElementLoaded(resource_dict["content_title"])
        Logging("Resource View - Wait until title is presented")

    Commands.Wait10s_ClickElement(resource_dict["cancel_button"])
    Logging("Resource View - Click Cancel button")

    Commands.Wait10s_InputElement(resource_dict["cancel_reason"], "Cancel Reason")
    Logging("Resource Cancel - Input cancel reason")

    Commands.ClickElement(resource_dict["confirm_button"])
    Logging("Resource Cancel - Confirm cancellation")
    
    try:
        Waits.Wait10s_ElementLoaded(resource_dict["fc_title"])
        Waits.Wait10s_ElementInvisibility(resource_dict["resource_attr"] % title)
        Logging("Resource List - Wait until reservation is invisible")
        TestCase_LogResult(**resource_tc["cancel"]["pass"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="resource", testcase="cancel", msg="click cancel")
        TestCase_LogResult(**resource_tc["cancel"]["fail"])  

    time.sleep(1)

def Resource_Comment():
    PrintYellow("[RESOURCE] COMMENT RESERVATION")
    Commands.ClickElement(resource_dict["memo"])
    Logging("Resource Comment - Click Memo button in content")

    comment = InputContent()
    Logging("Resource Comment - Input content for comment")

    Commands.ClickElement(resource_dict["ok_button"])
    Logging("Resource Comment - Click Save button")

    try:
        Waits.Wait10s_ElementLoaded(resource_dict["new_comment"] % comment)
        TestCase_LogResult(**resource_tc["comment"]["pass"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="resource", testcase="comment", msg="click save comment")
        TestCase_LogResult(**resource_tc["comment"]["fail"])

def Resource_Approve():
    PrintYellow("[RESOURCE] APPROVE RESERVATION")

    try:
        Commands.FindElement("//a[contains(@class, 'setting-admin')]")
        admin = True
    except WebDriverException:
        admin = False
    
    if admin == True:
        Waits.Wait10s_ElementLoaded(resource_dict["approve_button"])
        
        Commands.ClickElement(resource_dict["approve_button"])
        Logging("Click Approve button")

        Waits.WaitUntilPageIsLoaded(None)

        try:
            Waits.Wait10s_ElementInvisibility(resource_dict["approve_button"])
            TestCase_LogResult(**resource_tc["approval"]["pass"])
        except WebDriverException:
            TCResult_ValidateAlertMsg(menu="resource", testcase="approval", msg="click approve")
            TestCase_LogResult(**resource_tc["approval"]["fail"])
        
        time.sleep(1)

def Resource_AddNewResource():
    add_resource = None
    Commands.ClickElement(resource_dict["add_resource_submenu"])
    Logging("Click Add Resource sub-menu")
    
    try:
        Waits.WaitElementLoaded(3, resource_dict["add_resource_child"] % "display:block")
    except WebDriverException:
        Commands.ClickElement(resource_dict["add_resource_child"] % "display: none;")
        Logging("Click to open Resource Manager")
    finally:
        Commands.ClickElement(resource_dict["add_resource_child"] % "display:block")
        Logging("Access Add Resource page")

    Waits.Wait10s_ElementLoaded(resource_dict["add_category"])
    time.sleep(1)

    Commands.ClickElement(resource_dict["add_category"])
    Logging("Click Add Resource button")

    Commands.Wait10s_ClickElement(resource_dict["conference_room"])
    Logging("Select Conference Room")

    time.sleep(1)

    category_name = "Meeting Room"
    resource_name  = "Meeting Room 01"
    Commands.InputElement(resource_dict["new_category_name"], category_name)
    Logging("Input category name")

    Commands.ClickElement(resource_dict["save_category"])
    Logging("Click save category")

    try:
        Waits.Wait10s_ElementLoaded(resource_dict["resource_category"] % category_name)
        Logging("Resource category is added successfully")
        add_category = True
    except WebDriverException:
        Logging("Fail to add resource category")
        add_category = False
    
    if add_category == True:
        time.sleep(1)
        Commands.ClickElement(resource_dict["resource_category"])
        Logging("Select new category")

        Commands.Wait10s_ClickElement(resource_dict["add_resource"])
        Logging("Click Add resource button")

        Waits.Wait10s_ElementLoaded(resource_dict["permission_system"])
        time.sleep(1)

        Commands.InputElement(resource_dict["resource_name"], resource_name)
        Logging("Input Conference Room Name")

        Commands.ClickElement(resource_dict["permission_system"])
        Logging("Select Permission System option")
        Waits.Wait10s_ElementLoaded(resource_dict["permission_send_mail"])
        
        time.sleep(1)

        Commands.ClickElement(resource_dict["save_resource"])
        Logging("Click save resource")
        
        try:
            Waits.Wait10s_ElementLoaded(resource_dict["category_expander"] % category_name)
            Logging("New resource is added successfully")
            add_resource = True
        except WebDriverException:
            Logging("Cannot find new resource")
            add_resource = False
    
    return add_resource

def ResourceReservationExecution(domain_name):
    PrintYellow("-------------------------------------------")
    PrintYellow("[MENU RESOURCE]")
    request_data = Resource_AccessMenu(domain_name)
    if bool(request_data["request_resource"]) == True:
        reservation = Resource_Reserve()
        if bool(reservation) == True:
            Resource_View(reservation[0], reservation[1])
            if request_data["resource_system"] == "Permission":
                Resource_Approve()
            Resource_Cancel(reservation[0])
            ValidateUnexpectedModal()

