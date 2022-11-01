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
    global diary_dict, diary_tc
    diary_dict = dict(data["diary"])
    diary_tc = dict(data["testcase_result"]["diary"])

    global report_dict, report_tc
    report_dict = dict(data["report"])
    report_tc = dict(data["testcase_result"]["report"])

def Task_AccessMenu(domain_name, name):
    PrintYellow("-------------------------------------------")
    PrintYellow("[MENU TASK]")
    if name == "work_diary":
        try:
            Commands.NavigateTo(domain_name + "/task/diary/write/pdefault/")
            Logging("Access work diary")
            
            Waits.WaitElementLoaded(15, data["editor"]["tox_iframe"])
            Logging("Find Editor")
            
            access_menu = True
        except WebDriverException:
            access_menu = False
    else:
        try:
            Commands.NavigateTo(domain_name + "/task/report/list/udefault/")
            Logging("Access task report")
            
            Waits.Wait10s_ElementLoaded(report_dict["list_footer"])
            Logging("Find list foooter")
            
            access_menu = True
        except WebDriverException:
            access_menu = False

    Waits.WaitUntilPageIsLoaded(None)

    return access_menu

def Task_CollectListData(name):
    time.sleep(1)
    
    try:
        list_data = CollectListData(list_footer=data[name]["list_footer"], page_total=data[name]["page_total"])
        tasks = list_data["total_items"]
        pages = list_data["total_pages"]
    except:
        tasks = 0
        pages = 0

    task_list = {
        "tasks": tasks,
        "pages": pages
    }

    return task_list

def Task_SendReport(domain_name, recipient_id):
    PrintYellow("[MENU TASK] TASK REPORT - SEND TASK REPORT")
    
    CommonWriteItem(report_dict["pen_button"],report_dict["subject"], objects.hanbiro_content)
    Logging("Send Task Report - Click Create button / Input subject and content")
    
    if recipient_id == data["tooltip"]["recipient"]:
        recipient_id = False

    if bool(recipient_id) == True:
        selected_user = Org_SelectUser(
            org_tree = report_dict["org_select"][0],
            org_input = report_dict["org_select"][1],
            org_plus = report_dict["org_select"][2],
            org_save = report_dict["org_select"][4],
            recipient_id = recipient_id
        )
        select_org = True
    else:
        select_org = False
    
    report = None
    if select_org == True:
        report_name = Functions.GetInputValue(report_dict["subject"])
        if bool(selected_user) == True:
            FindPushNoti()
            Commands.Wait10s_ClickElement(report_dict["save_button"])
            Logging("Send Task Report - Save Task Report")

            Commands.Wait10s_ClickElement(diary_dict["ok_button"])
            Logging("Send Task Report - Confirm sending")

            try:
                Waits.Wait10s_ElementLoaded(report_dict["content_subject"])
                TestCase_LogResult(**report_tc["write"]["pass"])

                report_subject = Functions.GetElementText(report_dict["content_subject"])
                if report_subject == report_name:
                    TestCase_LogResult(**report_tc["view"]["pass"])
                    report = report_name
                    
                    Commands.Wait10s_ClickElement(report_dict["recipient"])
                    try:
                        Waits.Wait10s_ElementLoaded(report_dict["report_recipient"])    
                        Logging("Report recipient is not empty")
                    except WebDriverException:
                        report_tc["view"]["fail"].update({"description": "Report recipient is empty"})
                        TestCase_LogResult(**report_tc["view"]["fail"])     
                else:
                    TestCase_LogResult(**report_tc["view"]["fail"])
                    report = False
            except WebDriverException:
                TCResult_ValidateAlertMsg(menu="task_report", testcase="write", msg="click save task report")
                TestCase_LogResult(**report_tc["write"]["fail"])
                
                Logging("Cannot save task report \n Back to report list")
                Commands.ClickElement(report_dict["back_write"])
                Waits.Wait10s_ElementLoaded(report_dict["list_footer"])
                report = None
    else:
        Logging("Task recipient is not selected \n Back to task report list")

        Commands.ClickElement(report_dict["close_org"])
        Logging("Close org tree")
        
        time.sleep(1)
        Waits.WaitUntilPageIsLoaded(None)
        Waits.Wait10s_ElementLoaded(report_dict["back_write"])
        
        Commands.ClickElement(report_dict["back_write"])
        Logging("Click Back button")
        Waits.Wait10s_ElementLoaded(report_dict["list_footer"])
        report = None

    return report

def Task_EditWorkDiary():
    PrintYellow("[MENU TASK] EDIT WORK DIARY")
    Commands.Wait10s_ClickElement(diary_dict["content_more"])
    Logging("Edit Work Diary - Click More button")

    Commands.Wait10s_ClickElement(diary_dict["modify_button"])
    Logging("Edit Work Diary - Click Modify button")
    
    Waits.WaitElementLoaded(20, data["editor"]["tox_iframe"])
    Commands.SwitchToFrame(data["editor"]["tox_iframe"])
    time.sleep(1)
    Commands.InputElement(data["editor"]["input_p"], objects.content_edit)
    Logging("Edit Work diary - Update content")
    Commands.SwitchToDefaultContent()
    
    CloseAutosave()

    FindPushNoti()
    Commands.ClickElement(diary_dict["save_button"])
    Logging("Edit Work Diary - Click Save button")

    Waits.Wait10s_ElementLoaded(diary_dict["task_view_content"])
    if objects.content_edit in Functions.GetPageSource():
        Logging("Edit Work Diary - Content is updated successfully")
    else:
        ValidateFailResultAndSystem("Edit Work Diary - Fail to update work diary")

    Commands.ClickElement(diary_dict["back_button"])
    Logging("View Work Diary - Back to work diary list")

    Waits.WaitUntilPageIsLoaded(diary_dict["list_footer"])

def Task_WriteWorkDiary(domain_name):
    PrintYellow("[MENU TASK] WRITE WORK DIARY")
    work_diary_name = "[Task]" + objects.hanbiro_title
    
    Commands.InputElement(diary_dict["subject"], work_diary_name)
    Logging("Write - Input title/ subject")
    work_diary_name =  Functions.GetInputValue(diary_dict["subject"])
    Logging(">>> Title: [" + work_diary_name + "] is input")
    Logging("Write work diary - Click Create button / Input subject and content")
    
    Commands.ExecuteScript("window.scrollTo(0,301)") # Scroll down to Editor

    Commands.SwitchToFrame(data["editor"]["tox_iframe"])
    Logging("Write - Switch to Editor iframe")
    Commands.InputElement(data["editor"]["input_p"], objects.hanbiro_content)
    Logging("Write - Input content")
    Commands.SwitchToDefaultContent()
    
    CloseAutosave()
    FindPushNoti()

    Commands.ClickElement(diary_dict["save_button"])
    Logging("Click Save button")
    try:
        Waits.Wait10s_ElementLoaded(diary_dict["content_subject_text"] % objects.hanbiro_title)
        TestCase_LogResult(**diary_tc["write"]["pass"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="work_diary", testcase="write", msg="click save work diary")
        TestCase_LogResult(**diary_tc["write"]["fail"])
        Commands.NavigateTo(domain_name + "/task/diary/write/pdefault/") 

    return work_diary_name

def Task_ResgiterNewComment(name):
    PrintYellow("[MENU TASK] REGISTER COMMENT")
    Commands.ClickElement(diary_dict["comment_button"])
    
    time.sleep(1)

    comment = InputContent()

    Commands.ClickElement(diary_dict["confirm_button"])
    try:
        Waits.Wait10s_ElementLoaded(diary_dict["new_comment"] % comment)
        TestCase_LogResult(**data["testcase_result"]["" + name + ""]["comment"]["pass"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu=name, testcase="comment", msg="click save comment")
        TestCase_LogResult(**data["testcase_result"]["" + name + ""]["comment"]["fail"])
    
    if name == "task_report":
        Commands.ClickElement(report_dict["back_button"])
        Logging("View Task Report - Back to Report list")

        Waits.Wait10s_ElementLoaded(report_dict["list_footer"])

def WorkDiary_CopyArchive(domain_name):
    Waits.Wait10s_ElementLoaded(diary_dict["task_view_content"])

    work_title = Functions.GetElementText(diary_dict["task_view_title"])
    Logging("work diary content - collect work title: " + work_title)

    work_content = Functions.GetElementText(diary_dict["task_view_content"])
    Logging("work diary content - collect work content: " + work_content)

    Commands.ClickElement(diary_dict["task_view_more"])
    Logging("work diary content - click more button")

    Commands.Wait10s_ClickElement(diary_dict["copy_archive"])
    Logging("work diary content - Click Copy to Archive button")

    time.sleep(1)

    archive_folder = CopyArchive_SelectArchiveFolder()

    archive_data = {
        "archived_name": work_title,
        "archive_folder": archive_folder
    }

    return archive_data

def TaskReport_CopyArchive(domain_name):
    current_url = DefineCurrentURL()
    if "/task/report/view" not in current_url:
        Waits.Wait10s_ElementLoaded(report_dict["list_report_title"])
        time.sleep(1)
        Commands.ClickElement(report_dict["list_report_title"])
        Logging("Select task report to view")
        Waits.Wait10s_ElementLoaded(report_dict["view_report_content"])

    report_title = Functions.GetElementText(report_dict["view_report_title"])
    Logging("task report - select task title: " + report_title)

    report_content = Functions.GetElementText(report_dict["view_report_content"])
    Logging("task report - select task content: " + report_content)

    Commands.ClickElement(report_dict["view_report_more"])
    Logging("task report content - click more button")

    Commands.Wait10s_ClickElement(diary_dict["copy_archive"])
    Logging("task report content - Click Copy to Archive button")

    archive_folder = CopyArchive_SelectArchiveFolder()

    archive_data = {
        "archived_name": report_title,
        "archive_folder": archive_folder
    }

    return archive_data

def Task_SearchWorkDiary(name):
    PrintYellow("[TEST CASE] SEARCH")
    time.sleep(1)
    search_result = wrapper(searchInput, data["" + name + ""]["search_input"])
    if search_result == True:
        TestCase_LogResult(**data["testcase_result"]["" + name + ""]["search"]["pass"])
    else:
        TestCase_LogResult(**data["testcase_result"]["" + name + ""]["search"]["fail"])

    time.sleep(1)

def Task_ValidateNextPageList(name):
    PrintYellow("[MENU WORK DIARY] MOVE PAGE")
    '''Check if the function moving to next page is working
        and the list changes after moving '''
    menu_dict = dict(data["" + name + ""])
    
    List_ValidateListMovingPage(menu_dict["list_target"], menu_dict["item_suf"], menu_dict["page_total"], menu_dict["nextpage_icon"])

def Task_AccessViewPage(name):
    if name == "work_diary":
        task = Commands.FindElements(diary_dict["list_diary_title"])
    else:
        task = Commands.FindElements(report_dict["list_report_gotitle"])
    
    if int(len(task)) > 1:
        task[0].click()
        Logging("Click " + name + "to access view page")
    Waits.Wait10s_ElementLoaded(diary_dict["content_view"])

def Task_SearchDetails(task_type):
    if task_type == "task":
        dict_name = "work_diary"
        list_footer_xpath = diary_dict["list_footer"]
        item_xpath = diary_dict["diary_item"] 
    else:
        dict_name = "task_report"
        list_footer_xpath = report_dict["list_footer"]
        item_xpath = report_dict["report_item"]
    
    current_url = DefineCurrentURL()
    if "/view/" in current_url:
        Commands.ClickElement(report_dict["back_button_2"])
        Logging("Back to list from detail page")

    Waits.Wait10s_ElementLoaded(list_footer_xpath)

    list1 = DefineListLength(item_xpath)
    if list1 > 0:
        Commands.ClickElement(report_dict["search_details_b"])
        Logging("Open search box")

        Waits.Wait10s_ElementLoaded(report_dict["search_subject"])
        
        task_data = {
            "task": {
                "div": {
                    "writer": diary_dict["item_div"]["writer"],
                    "subject": diary_dict["item_div"]["subject"]
                },
                "search": {
                    "writer":  diary_dict["search_details"]["writer"],
                    "subject":  diary_dict["search_details"]["subject"]
                }   
            },
            "report": {
                "div": {
                    "writer": report_dict["item_div"]["writer"],
                    "subject": report_dict["item_div"]["subject"]
                },
                "search": {
                    "writer":  report_dict["search_details"]["writer"],
                    "subject":  report_dict["search_details"]["subject"]
                }   
            }
        }

        time.sleep(1)

        for search_data in task_data[task_type]["div"].keys():
            element_xpath = task_data[task_type]["div"][search_data]
            text = Functions.GetElementText(element_xpath)
            Logging("Key word for " + search_data + " -> " + text)
            
            search_input_xpath = task_data[task_type]["search"][search_data]
            Commands.InputElement(search_input_xpath, text)
            Logging("-> Input key word for " + search_data)
            #time.sleep(1)

        Commands.ClickElement(report_dict["search_button"])
        Logging("Click Search button")

        if list1 > 1:
            i=0
            for i in range(0,10):
                i+=1
                time.sleep(1)
                list2 = DefineListLength(item_xpath)
                if list2 != list1:
                    search_result = True
                    break
                else:
                    search_result = False
        else:
            try:
                Commands.FindElement(report_dict["no_data"])
                Logging("List is empty")
                search_result = False
            except:
                search_result = True
        try:
            Commands.FindElement(report_dict["content_wrap"])
            Logging("Page is error")
            TestCase_LogResult(**data["testcase_result"][dict_name]["search"]["fail"])
        except WebDriverException:
            if search_result == True:
                TestCase_LogResult(**data["testcase_result"][dict_name]["search"]["pass"])
            else:
                TestCase_LogResult(**data["testcase_result"][dict_name]["search"]["fail"])

        time.sleep(1)

        list3 = DefineListLength(item_xpath)
        Commands.ClickElement(report_dict["search_reset"])
        Logging("Click Reset button")
        i=0
        for i in range(0,10):
            i+=1
            time.sleep(1)
            list4 = DefineListLength(item_xpath)
            if list4 != list3:
                reset = True
                break
            else:
                reset = False
        if reset == True:
            Logging("Reset search result successfully")
        else:
            Logging("Fail to reset list")

def TaskWorkDiaryExecution(domain_name):
    access_menu = Task_AccessMenu(domain_name=domain_name, name="work_diary")
    work_diary_name = Task_WriteWorkDiary(domain_name)
    Task_ResgiterNewComment(name="work_diary")
    archive_data = WorkDiary_CopyArchive(domain_name)
    Task_EditWorkDiary()
    Task_SearchDetails("task")
    Task_ValidateNextPageList(name="work_diary")
    ValidateUnexpectedModal()

def TaskReportExcution(domain_name, recipient_id):
    archive_data = None

    Task_AccessMenu(domain_name=domain_name, name="task_report")
    
    if recipient_id != data["tooltip"]["recipient"]:
        report_name = Task_SendReport(domain_name, recipient_id)
        Task_ResgiterNewComment(name="task_report")
        archive_data = ""
        #TaskReport_CopyArchive(domain_name)
    
    Task_CollectListData(name="task_report")
    Task_SearchDetails("report")
    Task_ValidateNextPageList(name="task_report")
    ValidateUnexpectedModal()
