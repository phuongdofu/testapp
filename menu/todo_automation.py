from os import access
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
    global todo_dict, todo_tc
    todo_dict = dict(data["todo"])
    todo_tc = dict(data["testcase_result"]["todo"])

def ToDo_AccessToDoMenu():
    PrintYellow("-------------------------------------------")
    PrintYellow("[MENU TODO]")

    todo_menu = Commands.FindElement(todo_dict["todo_nav"])
    todo_display = todo_menu.get_attribute("style")
    if todo_display == 'display: none;':
        FindPushNoti()
        
        Commands.MoveToElement("//a[@class='lm--navDropdown']")
        Logging("Hover nav dropdown (more menu)")

        Waits.WaitElementLoaded(5, todo_dict["todo_menu"])
        Commands.ClickElement(todo_dict["todo_inline"])
        Logging("Access ToDo menu from dropdown menu")
    else:
        todo_menu.click()
        Logging("Access ToDo menu")

    try:
        Waits.Wait10s_ElementLoaded(todo_dict["list_footer"])
        access_menu = True
    except WebDriverException:
        access_menu = False

    if access_menu == True:
        Commands.Wait10s_ClickElement("//span[contains(.,' My To-Do')]").click()
        Waits.Wait10s_ElementLoaded(todo_dict["list_footer"])

        list_data = CollectListData(todo_dict["list_footer"], todo_dict["page_total"])
        todos = list_data["total_items"]
        pages = list_data["total_pages"]
    else:
        todos = 0
        pages = 0

    todo_data = {
        "access_menu": access_menu,
        "todos": todos,
        "pages": pages
    }

    return todo_data

def ToDo_WriteToDo():
    PrintYellow("[TODO] WRITE TODO")

    Waits.Wait10s_ElementLoaded(todo_dict["list_footer"])
    
    CommonWriteItem(todo_dict["red_pen"], todo_dict["subject"], objects.hanbiro_title)

    todo_name = Functions.GetInputValue(todo_dict["subject"])
    
    CloseAutosave()
    FindPushNoti()

    Commands.ClickElement(todo_dict["save_button"])
    Logging("Write ToDo - Click Save button")

    Logging(todo_name)

    try:
        Waits.Wait10s_ElementLoaded(todo_dict["list_footer"])
        TestCase_LogResult(**todo_tc["write"]["pass"])
    except WebDriverException:
        TestCase_LogResult(**todo_tc["write"]["fail"])

    return todo_name

def ViewToDo(hanbiro_title):
    PrintYellow("[TODO] VIEW TODO")
    try:
        Waits.Wait10s_ElementLoaded(todo_dict["todo_item"])
        Commands.ClickElement(todo_dict["todo_title"] % hanbiro_title)
    except WebDriverException:
        Commands.ClickElements(todo_dict["todo_item"], 0)
    finally:
        Logging("View ToDo - Click to view new todo")

    try:
        Waits.Wait10s_ElementLoaded(todo_dict["content_subject"])
        TestCase_LogResult(**todo_tc["view"]["pass"])
        access_view = True
    except WebDriverException:
        TestCase_LogResult(**todo_tc["view"]["fail"])
        access_view = False
    
    if access_view == True:
        try:
            Waits.Wait10s_ElementLoaded(data["editor"]["tox_editor_header"])

            Commands.Selectbox_ByVisibleText("//select[@id='todo-percent']", "Complete")
            Logging("View ToDo - Update ToDo progress with Complete option")

            Commands.InputElement(todo_dict["progress_memo"], objects.hanbiro_content)
            Logging("View ToDo - Input progress memo")

            Commands.ClickElement(todo_dict["save_progress"])
            Logging("View ToDo - Save progress")
            
            progress_result = []
            new_progress = Waits.Wait10s_ElementLoaded(todo_dict["new_progress"])
            if new_progress.is_displayed():
                progress_result.append(True)
            else:
                progress_result.append(False)
                todo_tc["update_progress"]["fail"].update({"description": "Fail to save ToDo progress"})
                TestCase_LogResult(**todo_tc["update_progress"]["fail"])
            
            new_memo = Commands.FindElement(todo_dict["new_memo"])
            if new_memo.text == objects.hanbiro_content:
                progress_result.append(True)
            else:
                progress_result.append(False)
                todo_tc["update_progress"]["fail"].update({"description": "Fail to save progress memo"})
                TestCase_LogResult(**todo_tc["update_progress"]["fail"])
            
            success_label = Commands.FindElement(todo_dict["success_label"])
            if success_label.is_displayed():
                progress_result.append(True)
            else:
                progress_result.append(False)
                todo_tc["update_progress"]["fail"].update({"description": "Fail to update ToDo progress status"})
                TestCase_LogResult(**todo_tc["update_progress"]["fail"])
            
            if False in progress_result:
                pass
            else:
                TestCase_LogResult(**todo_tc["update_progress"]["pass"])

            Commands.ClickElement(todo_dict["back_button"])
            Waits.Wait10s_ElementLoaded(todo_dict["list_footer"])
        except WebDriverException:
            todo_tc["update_progress"]["fail"].update({"description": "ToDo progress module is not displayed"})
            TestCase_LogResult(**todo_tc["update_progress"]["fail"])

def ToDo_SearchToDo():
    PrintYellow("[TODO] SEARCH TODO")

    Waits.Wait10s_ElementLoaded(todo_dict["list_todo_title"])
    title = Commands.FindElements(todo_dict["list_todo_title"])
    todo_name = title[0].text
    
    search_dict = {
        "title": {
            "key": todo_name,
            "value": "title"
        }
    }
    
    search_todo_dict = dict(todo_dict["search_details"])
    search_todo_dict["search_dict"] = search_dict
    SearchDetailsBySelectBox(**search_todo_dict)

    time.sleep(1)

def ToDo_ValidateNextPageList():
    PrintYellow("[MENU TASK TODO] MOVE PAGE")
    List_ValidateListMovingPage(todo_dict["list_target"], todo_dict["item_suf"], todo_dict["page_total"], todo_dict["nextpage_icon"])

def ToDoExecution():
    todo_data = ToDo_AccessToDoMenu()
    todo_subject = ToDo_WriteToDo()
    ViewToDo(todo_subject)
    ToDo_SearchToDo()
    ToDo_ValidateNextPageList()
    ValidateUnexpectedModal()

    
