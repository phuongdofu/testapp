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
    global project_dict, project_tc
    project_dict = dict(data["project"])
    project_tc = dict(data["testcase_result"]["project"])

def Project_AccessProject(user_id):
    PrintYellow("[PROJECT] ACCESS MENU")
    access_result = AccessGroupwareMenu(name="project,Project", page_xpath=project_dict["project_list_footer"])
    access_project = None
    time.sleep(1)
    if access_result == True:
        TestCase_LogResult(**project_tc["access_menu"]["pass"])
        
        try:
            Commands.ClickElement(project_dict["project"])
            Logging("Click existing project")
            Waits.Wait10s_ElementLoaded(project_dict["work_list"])
            current_url = DefineCurrentURL()
            access_project = True
        except WebDriverException:
            try:
                Commands.FindElement(project_dict["no_project"])
                Commands.ClickElement(project_dict["create_project"])
                Logging("No project available - Click to create new prject")
                create_project = True
            except WebDriverException:
                project_tc["access_menu"]["fail"].update({"description": "No project available to access / User has no permission to create"})
                TestCase_LogResult(**project_tc["access_menu"]["fail"])
                create_project = False
            
            if create_project == True:
                try:
                    Waits.WaitUntilPageIsLoaded(None)
                    Waits.WaitElementLoaded(15, data["editor"]["tox_iframe"])
                    
                    Commands.InputElement(project_dict["code_name"], "projectcode")
                    Logging("Input project code")

                    Commands.InputElement(project_dict["project_name_input"], "Project Test")
                    Logging("Input project name")
                    
                    project_users = ["Leader", "Participant(s)"]
                    for user in project_users:
                        Commands.ClickElement(project_dict["user_div"] % user)
                        Logging("Focus Leader autocomplete")
                        
                        Commands.InputElement(project_dict["select_user"] % user, user_id)
                        Logging("Input user_id in autocomplete input box")
                        
                        Waits.Wait10s_ElementLoaded(project_dict["user_is.selected"])
                        time.sleep(0.5)
                        
                        Commands.ClickElement(project_dict["user_is.selected"])
                        Logging("Select " + user + "")
                        time.sleep(0.5)
                    
                    Commands.ClickElement(project_dict["save_project"])
                    print("Save project")
                    
                    Waits.Wait10s_ElementLoaded(project_dict["project_linktext"] % "Project Test")
                    Logging("Create new project successfully")

                    Commands.ClickElement(project_dict["project_linktext"] % "Project Test")
                    Logging("Access project")
                    access_project = True
                except WebDriverException:
                    pass
        
        if bool(access_project) == True:
            current_url = DefineCurrentURL()
        else:
            current_url = None            
    else:
        TestCase_LogResult(**project_tc["access_menu"]["fail"])
        current_url = None
        access_project = None

    project_menu = {
        "access_project": access_project,
        "current_url": current_url
    }

    return project_menu
    
def Project_AddWork():
    PrintYellow("[PROJECT] - ADD WORK")
    CommonWriteItem(project_dict["work_add"], project_dict["subject"], objects.hanbiro_content)
    # Click to add work in project, input subject and content

    work_name = Functions.GetInputValue(project_dict["subject"])
    start_day =  Functions.GetInputValue(project_dict["start_day"]).split("/")[2]
    
    if int(start_day) < 10:
        start_day = start_day.replace("0", "")
    Logging("start_day: " + start_day)

    end_day = int(start_day) + 3
    Logging("end_day :" + str(end_day))

    FindPushNoti()
    Commands.ClickElement(project_dict["work_save"])
    Logging("Click Save button")
    
    try:
        Waits.Wait10s_ElementLoaded(project_dict["list_work_name"] % work_name)
        TestCase_LogResult(**project_tc["write_work"]["pass"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="project", testcase="write_work", msg="click save work")
        TestCase_LogResult(**project_tc["write_work"]["fail"])

    return work_name

def Project_ViewWork(work_name):
    PrintYellow("[PROJECT] - VIEW WORK")
    Waits.Wait10s_ElementLoaded(project_dict["work_list_item"])
    
    work_xpath = project_dict["list_work_name"] % work_name
    Commands.ClickElement(work_xpath)
    Logging("Select work to view")
    
    try:
        Waits.Wait10s_ElementLoaded(project_dict["work_content"])
        TestCase_LogResult(**project_tc["view_work"]["pass"])
        time.sleep(1)
        access_work = True
    except WebDriverException:
        TestCase_LogResult(**project_tc["view_work"]["fail"])
        access_work = False 

    if access_work == True:
        work_content = Functions.GetElementText(project_dict["work_content"])
        Logging(work_content)
        if work_content == "":
            project_tc["view_work"]["fail"].update({"description": "Work content is empty"})
            TestCase_LogResult(**project_tc["view_work"]["fail"])

    return access_work

def Project_EditWork():
    PrintYellow("[PROJECT] - EDIT WORK")
    Commands.ClickElement(project_dict["more_dropdown"])
    Logging("Project - Edit Work - Click More dropdown in content")

    Commands.ClickElement(project_dict["modify_button"])
    Logging("Project - Edit Work - Select Modify button in dropdown")

    Waits.WaitUntilPageIsLoaded(None)

    Commands.SwitchToFrame(data["editor"]["tox_iframe"])

    Waits.Wait10s_ElementLoaded(data["editor"]["input_p"])
    time.sleep(1)
    comment = Functions.GetElementText(data["editor"]["input_p"])
    Commands.InputElement(data["editor"]["input_p"], comment + Keys.ENTER + objects.content_edit)
    
    Commands.SwitchToDefaultContent()

    Commands.ClickElement(project_dict["work_save"])
    Logging("Project - Edit Work - Save and update work")

    Waits.Wait10s_ElementLoaded(project_dict["work_content"])

    if objects.content_edit in Functions.GetPageSource():
        Logging("Project - Edit Work - New work subject is updated successfully")
        Logging(objects.testcase_pass)
    else:
        ValidateFailResultAndSystem("Project - Edit Work - Fail to update new work subject")
        Logging(objects.testcase_fail)

def Project_ChangeStatus():
    PrintYellow("[PROJECT] - UPDATE STATUS")
    Commands.ClickElement(project_dict["change_status"])

    Waits.Wait10s_ElementLoaded(data["editor"]["tox_iframe"])
    Waits.Wait10s_ElementLoaded(project_dict["status_selectbox"])

    Commands.Selectbox_ByVisibleText(project_dict["status_selectbox"], "Resolved")
    Logging("Change Status - Update with status 'Resolved'")

    status_data = []
    try:
        Waits.Wait10s_ElementLoaded(project_dict["update_done"])
        status_data.append(True)
    except WebDriverException:
        status_data.append(False)
        TestCase_LogResult(**project_tc["update_status"]["fail"])

    Commands.ClickElement(project_dict["ok_button"])
    Waits.Wait10s_ElementLoaded(project_dict["content_view_wrap"])

    try:
        Waits.Wait10s_ElementLoaded(project_dict["status_update"])
        status_data.append(True)
    except WebDriverException:
        status_data.append(False)
        project_tc["update_status"]["fail"].update({"description": "Fail to update status"})
        TestCase_LogResult(**project_tc["update_status"]["fail"])

    try:
        Commands.FindElement(project_dict["percent_update"])
        status_data.append(True)
    except WebDriverException:
        status_data.append(False)
        project_tc["update_status"]["fail"].update({"description": "Fail to update percent"})
        TestCase_LogResult(**project_tc["update_status"]["fail"])

    if False in status_data:
        pass
    else:
        TestCase_LogResult(**project_tc["update_status"]["pass"])

def Project_AddTicket():
    PrintYellow("[PROJECT] - ADD TICKET")
    Commands.ClickElement(project_dict["write_ticket"])
    Logging("Add Ticket - Click button 'Write Ticket' in project content page")

    Waits.Wait10s_ElementLoaded(data["editor"]["tox_iframe"])
    time.sleep(1)

    ticket_name = "Ticket is added at " + objects.date_time
    Commands.InputElement(project_dict["project_name_input"], ticket_name)
    Logging("Add Ticket - Input ticket name")

    InputContent()
    Logging("Add Ticket - Input ticket content")

    Commands.ClickElement(project_dict["save_span"])
    
    try:
        Waits.Wait10s_ElementLoaded(project_dict["list_ticket_name"] % ticket_name)
        TestCase_LogResult(**project_tc["write_ticket"]["pass"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="project", testcase="write_ticket", msg="click save ticket")
        TestCase_LogResult(**project_tc["write_ticket"]["fail"])

    return ticket_name

def Project_ViewTicket(ticket_name):
    PrintYellow("[PROJECT] - VIEW TICKET")
    Waits.Wait10s_ElementLoaded(project_dict["ticket_item"])

    try:
        ticket = Commands.FindElement(project_dict["list_ticket_name"] % ticket_name)
    except WebDriverException:
        ticket = Commands.FindElement(project_dict["ticket_item"].replace("/tr/", "/tr[1]/"))
    finally:
        ticket.click()
        Logging("Select ticket to view")
    
    try:
        Waits.Wait10s_ElementLoaded(project_dict["ticket_content"])
        TestCase_LogResult(**project_tc["view_ticket"]["pass"])
    except WebDriverException:
        TestCase_LogResult(**project_tc["view_ticket"]["fail"])

    Commands.Wait10s_ClickElement(project_dict["back_button"])
    Waits.Wait10s_ElementLoaded(project_dict["work_list_item"])

def Project_InsertTicket(ticket_name):
    PrintYellow("[PROJECT] - INSERT TICKET")
    Commands.ClickElement(project_dict["related_ticket"])

    Commands.Wait10s_ClickElement(project_dict["add_ticket"])
    Commands.Wait10s_ClickElement(project_dict["ticket_table"])

    Commands.ClickElement(project_dict["insert_ticket"])
    try:
        Waits.WaitElementLoaded(5, project_dict["selected_ticket"])
        ticket = Functions.GetElementText(project_dict["selected_ticket"])
        if ticket == ticket_name:
            TestCase_LogResult(**project_tc["insert_ticket"]["pass"])
        else:
            TestCase_LogResult(**project_tc["insert_ticket"]["fail"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="project", testcase="insert_ticket", msg="click save ticket")
        TestCase_LogResult(**project_tc["insert_ticket"]["fail"])

def Project_AddReference():
    PrintYellow("[PROJECT] - ADD REFERENCE")

    Commands.Wait10s_ClickElement(data["common"]["linktext"] % "Reference List")
    Logging("Add Reference - Access Reference tab")

    Commands.Wait10s_ClickElement(project_dict["add_reference"])
    Logging("Add Reference - Click to add reference")

    Waits.Wait10s_ElementLoaded(data["editor"]["tox_iframe"])

    reference_name = "Reference is written at " + objects.date_time
    Commands.InputElement(project_dict["project_name_input"], reference_name)
    Logging("Add Reference - Input reference title")

    InputContent()
    Logging("Add Reference - Input reference content")

    Commands.ClickElement(project_dict["save_span"])

    Waits.Wait10s_ElementLoaded(project_dict["work_list_item"])
    Logging("Add Reference - Access work list")

    Commands.Wait10s_ClickElement(data["common"]["linktext"] % "Reference List")
    Logging("Add Reference - Access Reference tab")

    try:
        Waits.Wait10s_ElementLoaded(project_dict["new_reference"] % reference_name)
        TestCase_LogResult(**project_tc["insert_reference"]["pass"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="project", testcase="insert_reference", msg="click insert reference")
        TestCase_LogResult(**project_tc["insert_reference"]["fail"])

    return reference_name

def Project_ViewReference(reference_name):
    PrintYellow("[PROJECT] - VIEW REFERENCE")
    Commands.ClickElement(project_dict["reference_tab"])
    Commands.Wait10s_ClickElement(project_dict["list_reference"])
    Logging("Reference List - Click to view reference")
    try:
        Waits.Wait10s_ElementLoaded(project_dict["content_reference_title"])
        TestCase_LogResult(**project_tc["view_reference"]["pass"])
    except WebDriverException:
        TestCase_LogResult(**project_tc["view_reference"]["fail"])

    Commands.ClickElement(project_dict["reference_back"])
    Waits.Wait10s_ElementLoaded(project_dict["work_list_item"])

def Project_EditReference():
    Commands.ClickElement(project_dict["reference_more"])
    Logging("View Reference - Click More button in reference view page")

    Commands.Wait10s_ClickElement(data["common"]["linktext"] % "Modify")
    Logging("View Reference - Click Modify button")

    EditContent()
    Logging("Edit Reference - Update content for reference")

    Commands.ClickElement(project_dict["save_button"])
    Logging("Edit Reference - Click to save new content for reference")

    Waits.Wait10s_ElementLoaded(project_dict["work_content"])

    if objects.content_edit in Functions.GetPageSource():
        Logging("Edit Reference - Reference content is updated successfully")
        Logging(objects.testcase_pass)
    else:
        ValidateFailResultAndSystem("Project - Edit Reference - Fail to update new reference")
        Logging(objects.testcase_fail)

def Project_InsertReference(reference_name):
    PrintYellow("[PROJECT] - INSERT REFERENCE")
    Waits.Wait10s_ElementLoaded(project_dict["work_add_reference"])
    
    Commands.ClickElement(project_dict["work_add_reference"])
    Logging("Work content - Click to add reference")

    Commands.Wait10s_ClickElement(project_dict["ticket_table"])
    Logging("Insert reference - Select reference [1]")

    Commands.ClickElement(project_dict["insert_reference"])
    Logging("Insert reference - Save selected reference")
    try:
        Waits.Wait10s_ElementLoaded(project_dict["selected_reference"])
        reference = Functions.GetElementText(project_dict["selected_reference"])
        if reference == reference_name:
            TestCase_LogResult(**project_tc["insert_reference"]["pass"])
        else:
            TestCase_LogResult(**project_tc["insert_reference"]["fail"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="project", testcase="insert_reference", msg="click save reference")
        TestCase_LogResult(**project_tc["insert_reference"]["fail"])

def Project_WorkComment():
    PrintYellow("[PROJECT] - WRITE COMMENT")

    Commands.ClickElement(project_dict["comment_button"])
    Logging("Click comment button")

    comment = InputContent()
    Logging("Input comment content")

    Commands.ClickElement(project_dict["ok_button"])
    Logging("Save comment")

    try:
        Waits.Wait10s_ElementLoaded(project_dict["new_comment"] % comment)
        TestCase_LogResult(**project_tc["work_comment"]["pass"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="project", testcase="work_comment", msg="click save comment")
        TestCase_LogResult(**project_tc["work_comment"]["fail"])

def Project_NavigateToProjectContent(project_url):
    current_url = DefineCurrentURL()
    if current_url != project_url:
        Commands.NavigateTo(project_url)
        try:
            Waits.Wait10s_ElementLoaded(project_dict["work_list_footer"])
            access_list = True
            Logging("Access list successfully")
        except WebDriverException:
            access_list = False
            Logging("Fail to access list")
    else:
        access_list = None
    
    return access_list

def ProjectExecution(user_id):
    PrintYellow("-------------------------------------------")
    PrintYellow("[MENU PROJECT]")
    
    project_menu = Project_AccessProject(user_id)
    work_name = Project_AddWork()
    ticket_name = Project_AddTicket()
    reference_name = Project_AddReference()
    access_work = Project_ViewWork(work_name)
    Project_EditWork()
    Project_InsertReference(reference_name)
    Project_InsertTicket(ticket_name)
    Project_ChangeStatus()
    Project_WorkComment()
    access_list = Project_NavigateToProjectContent(project_menu["current_url"])
    Project_ViewTicket(ticket_name)
    Project_ViewReference(reference_name)
    ValidateUnexpectedModal()

