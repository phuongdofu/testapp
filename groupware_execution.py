import time, sys, unittest, random, json, requests, openpyxl, testlink, platform, os
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
from menu.approval_automation import *
from menu.archive_automation import *
from menu.asset_automation import *
from menu.board_automation import *
from menu.calendar_automation import *
from menu.circular_automation import *
from menu.clouddisk_automation import *
from menu.contact_automation import *
from menu.expense_automation import *
from menu.login import *
from menu.mail_automation import *
from menu.project_automation import *
from menu.resource_automation import *
from menu.task_automation import *
from menu.todo_automation import *
from menu.whisper_automation import *

from socket_client_config import SendTestCaseFile

def GroupwareExecution(**domain_config):
    domain_name = domain_config["domain_name"]
    user_id = domain_config["user_id"]
    user_pw = domain_config["user_pw"]
    recipient_id = domain_config["recipient_id"]
    menu_dict = dict(data["excel_menu_list"])

    UserLogIn(driver, domain_name, user_id, user_pw)
    #UserLogIn_Quick(driver, domain_name, user_id, user_pw)

    if Logs.Value_CollectMenu(menu_dict["Mail"], 2):
        try:
            MailExecution(domain_name)
            Logs.UpdateSuccess_ColectMenu("Mail")
        except (WebDriverException, TimeoutException) as error:
            pass
    
    if Logs.Value_CollectMenu(menu_dict["Board"], 2):
        try:
            BoardExecution()
            Logs.UpdateSuccess_ColectMenu("Board")
        except (WebDriverException, TimeoutException) as error:
            pass
    
    if Logs.Value_CollectMenu(menu_dict["Contact"], 2):
        try:
            ContactExecution(domain_name)
            Logs.UpdateSuccess_ColectMenu("Contact")
        except (WebDriverException, TimeoutException) as error:
            pass
    
    if Logs.Value_CollectMenu(menu_dict["Task Work Diary"], 2):
        try:
            diary_data = TaskWorkDiaryExecution(domain_name)
            Logs.UpdateSuccess_ColectMenu("Task Work Diary")
        except (WebDriverException, TimeoutException) as error:
            pass
    
    if Logs.Value_CollectMenu(menu_dict["Task Report"], 2):
        try:
            report_data = TaskReportExcution(domain_name, recipient_id)
            Logs.UpdateSuccess_ColectMenu("Task Report")
        except (WebDriverException, TimeoutException) as error:
            pass
    
    if Logs.Value_CollectMenu(menu_dict["Calendar"], 2):
        try:
            CalendarExecution()
            Logs.UpdateSuccess_ColectMenu("Calendar")
        except (WebDriverException, TimeoutException) as error:
            pass
    
    if Logs.Value_CollectMenu(menu_dict["CloudDisk"], 2):
        try:
            CloudDiskExecution(domain_name)
            Logs.UpdateSuccess_ColectMenu("CloudDisk")
        except (WebDriverException, TimeoutException) as error:
            pass
    
    if Logs.Value_CollectMenu(menu_dict["Circular"], 2):
        try:
            CircularExecution(user_id)
            Logs.UpdateSuccess_ColectMenu("Circular")
        except (WebDriverException, TimeoutException) as error:
            pass
    
    if Logs.Value_CollectMenu(menu_dict["ToDo"], 2):
        try:
            ToDoExecution()
            Logs.UpdateSuccess_ColectMenu("ToDo")
        except (WebDriverException, TimeoutException) as error:
            pass
    
    if Logs.Value_CollectMenu(menu_dict["Project"], 2):
        try:
            ProjectExecution(user_id)
            Logs.UpdateSuccess_ColectMenu("Project")
        except (WebDriverException, TimeoutException) as error:
            pass
    
    if Logs.Value_CollectMenu(menu_dict["Archive"], 2):
        try:
            ArchiveExecution(user_pw)
            Logs.UpdateSuccess_ColectMenu("Archive")
        except (WebDriverException, TimeoutException) as error:
            pass
    
    if Logs.Value_CollectMenu(menu_dict["Asset"], 2):
        try:
            AssetExecution()
            Logs.UpdateSuccess_ColectMenu("Asset")
        except (WebDriverException, TimeoutException) as error:
            pass
    
    if Logs.Value_CollectMenu(menu_dict["Expense"], 2):
        try:
            ExpenseExecution(domain_name, recipient_id)
            Logs.UpdateSuccess_ColectMenu("Expense")
        except (WebDriverException, TimeoutException) as error:
            pass
    
    if Logs.Value_CollectMenu(menu_dict["Resource"], 2):
        try:
            ResourceReservationExecution(domain_name)
            Logs.UpdateSuccess_ColectMenu("Resource")
        except (WebDriverException, TimeoutException) as error:
            pass

    time.sleep(1)
    
    if Logs.Value_CollectMenu(menu_dict["Approval"], 2):
        try:
            approval_data = Approval_Execution(domain_name, recipient_id)
            Logs.UpdateSuccess_ColectMenu("Approval")
        except (WebDriverException, TimeoutException) as error:
            pass

    if Logs.Value_CollectMenu(menu_dict["Whisper"], 2):
        try:
            WhisperExecution_Driver1(domain_name, user_id)
            Logs.UpdateSuccess_ColectMenu("Whisper")
        except (WebDriverException, TimeoutException) as error:
            pass

    '''archive_list = [diary_data, report_data, approval_data]
    
    for archive_dict in archive_list:
        if bool(archive_dict) == True:
            try:
                CopyArchive_ValiateDocumentTransfer(
                    domain_name=domain_name, 
                    folder=archive_dict["archive_folder"][0], 
                    folder_name=archive_dict["archive_folder"][1], +
                    document_name=archive_dict["archived_name"]
                )
            except WebDriverException:
                pass'''
    
    shutil.copy(Files.testcase_log, Files.testcase_file)
    SendTestCaseFile()
                                                                                                                                                                                                                                                                                        
def Groupware_Execution(**domain_config):
    domain_name = domain_config["domain_name"]
    user_id = domain_config["user_id"]
    user_pw = domain_config["user_pw"]
    recipient_id = domain_config["recipient_id"]
    recipient_name = domain_config["recipient_name"]

    UserLogIn_Quick(driver, domain_name, user_id, user_pw)
    MailExecution(domain_name)
    BoardExecution()
    ContactExecution(domain_name)
    diary_data = TaskWorkDiaryExecution(domain_name)
    report_data = TaskReportExcution(domain_name, recipient_id)
    CalendarExecution()
    CloudDiskExecution(domain_name)
    CircularExecution(user_id)
    ToDoExecution()
    ProjectExecution(user_id)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               
    ArchiveExecution(user_pw)
    AssetExecution()
    ExpenseExecution(domain_name, recipient_id)
    ResourceReservationExecution(domain_name)
    time.sleep(1)
    approval_data = Approval_Execution(domain_name, recipient_id)
    #WhisperExecution_Driver1(domain_name, user_id)
    
    # archive_list = [diary_data, report_data, approval_data]
    # for archive_dict in archive_list:
    #     CopyArchive_ValidateDocumentTransfer(
    #         domain_name=domain_name, 
    #         folder=archive_dict["archive_folder"][0], 
    #         folder_name=archive_dict["archive_folder"][1], 
    #         document_name=archive_dict["archived_name"]
    #     )

def RunMainFeatures(**domain_config):
    global driver

    # Start webdriver
    driver = Driver.DefineDriver("web")
    driver.maximize_window()
    
    # Create logs file
    # logs = [Files.execution_log, Files.fail_log, Files.error_log, Files.testcase_log]
    Logs.CreateLogFiles()

    start_time = time.time()

    # Main Execution
    GroupwareExecution(**domain_config)

    # DEBUG Execution
    # Groupware_Execution(**data["domain_config"]["tg01"]) #debug
    
    end_time = time.time()
    duration = end_time - start_time
    print(duration)

    return driver

def Postmaster():
    import menu.hanbiro_postmaster as postmaster

    postmaster.NewOrganization_Execution("https://tg01.hanbiro.net/ngw/app/#", "o&YHK_+TJo?MtX&")
