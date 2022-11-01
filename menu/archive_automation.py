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
    global archive_dict, archive_tc
    archive_dict = dict(data["archive"])
    archive_tc = dict(data["testcase_result"]["archive"])

def AccessArchiveMeu(secure_pw):
    access_result = AccessGroupwareMenu(name="archive,Archive", page_xpath=archive_dict["my_archive"])
    
    access_folder = None
    my_folders = 0
    if bool(access_result) == True:
        Commands.ClickElement(archive_dict["my_archive"])
        Logging("Click My Archive sub-menu")
        
        try:
            Waits.WaitElementLoaded(3, archive_dict["my_archive_folder"])
            my_folders = Functions.GetListLength(archive_dict["folder_li"])
            Logging("My Archive Folder: [%s]" % my_folders)
        except WebDriverException:
            Logging("Cannot find archive folder")
    
    if bool(secure_pw) == True:
        if my_folders > 0:
            Commands.ClickElement(archive_dict["my_archive_folder"])
            Logging("Access archive folder")
            
            folder_name = Functions.GetElementText(archive_dict["archive_folder"])
            Logging("Access folder: " + folder_name)
        
        Waits.Wait10s_ElementLoaded(archive_dict["secure_pw"])
        Commands.InputElement_2Values(archive_dict["secure_pw"], secure_pw, Keys.ENTER)
        Logging("Input security password")
        time.sleep(2)

        # Check if security password is invalid
        alert_msg = ValidateModalDialog_ErrorAlert()
        if bool(alert_msg) == True:
            Logging("!!! MESSSAGE: Invalid password")
            access_folder = False
        
        if my_folders == 0 and bool(alert_msg) != True:
            new_folder_name = "Test Folder"
            Commands.Wait10s_ClickElement(archive_dict["newarchive_tree"])
            Logging("Click My Archive folder")
            Waits.WaitElementLoaded(3, archive_dict["newarchive_tree_active"])
            Logging("My Archive folder is being active")
            Commands.ClickElement(archive_dict["add_newarchive"])
            Logging("Click add new folder")
            Commands.Wait10s_InputElement(archive_dict["newarchive_folder"], new_folder_name)
            Logging("Input folder name")
            Commands.ClickElement(archive_dict["create_permission"])
            Logging("Click Enabled button (Create Permission)")
            time.sleep(1)
            Commands.ClickElement(archive_dict["newarchive_save"])
            Logging("Save new archive folder")
            try:
                Waits.Wait10s_ElementLoaded(archive_dict["archive_newfolder"] % new_folder_name)
                Logging("New archive folder is created successfully")
                Commands.Wait10s_ClickElement(archive_dict["my_archive"])
                Logging("Click Archive sub-menu")
                Commands.ClickElement(archive_dict["my_archive_folder"])
                Logging("Access archive folder")
                my_folders = 1
            except WebDriverException:
                Logging("Fail to create new archive folder")
            
        if my_folders > 0 and bool(alert_msg) != True:
            Waits.Wait10s_ElementLoaded(archive_dict["list_footer"])
            Logging("Access folder successfully")
            access_folder = True

    if access_folder == True:
        list_data = CollectListData(list_footer = archive_dict["list_footer"], 
                                    page_total = archive_dict["page_total"])
        archive_len = list_data["total_items"]
        pages = list_data["total_pages"]
    else:
        archive_len = 0
        pages = 0

    time.sleep(1)

    archive_menu = {
        "access_folder": access_folder,
        "archive": archive_len,
        "pages": pages
    }

    return archive_menu

def Archive_WriteArchiveDocument():
    '''Write new archive document with attached file from PC & CloudDisk '''
    PrintYellow("[ARCHIVE] WRITE NEW ARCHIVE DOCUMENT")

    try:
        Commands.FindElement(archive_dict["write"][0])
        create_permission = True
    except WebDriverException:
        create_permission = False
        archive_name = pc_attach_name = cloud_attach_name = None

    if create_permission == True:
        CommonWriteItem(archive_dict["write"][0], archive_dict["write"][1], objects.hanbiro_content)
        
        archive_name = objects.hanbiro_title
        Logging("Write Archive - Click Create button / Input subject and content")

        Commands.ScrollUp()
        Commands.ClickElement(archive_dict["archive_image"])
        Logging("Write Archive - Scroll up and click Upload Image Files button")

        pc_attach_name = Attachment_SelectPC()
        Logging("Write Archive - Upload file from PC")

        cloud_attach_name = Attachment_SelectImageCloudDisk()
        Logging("Write Archive - Select file from CloudDisk")

        FindPushNoti()
        Commands.ClickElement(archive_dict["save_button"])
        Logging("Write Archive - Save new archive")

        Waits.WaitUntilPageIsLoaded(None)

        try:
            Waits.Wait10s_ElementLoaded(data["archive"]["new_archive"] % archive_name)
            TestCase_LogResult(**archive_tc["write"]["pass"])
        except WebDriverException:
            TestCase_LogResult(**archive_tc["write"]["fail"])

    archive_data = {
        "archive_name": archive_name,
        "pc_attach_name": pc_attach_name,
        "cloud_attach_name": cloud_attach_name
    }

    return archive_data

def Archive_ViewArchive(archive_name, pc_attach_name, cloud_attach_name):
    PrintYellow("[ARCHIVE] VIEW ARCHIVE DOCUMENT")
    ''' View archive content and check content '''

    new_archive_xpath = archive_dict["archive_item"] % archive_name
    
    Waits.Wait10s_ElementLoaded(new_archive_xpath)
    archive_doc = Commands.FindElements(new_archive_xpath)
    archive_range = Functions.GetListLength(new_archive_xpath)
    
    i=0
    for i in range(0, archive_range):
        try:
            archive_doc[i].click()
            break
        except WebDriverException:
            pass
        i+=1
    
    Logging("View Archive - Open content")

    view_fail = dict(data["testcase_result"]["archive"]["view"]["fail"])
    try:
        Commands.SwitchToFrame("//*[@id='viewDetail']")
        Waits.Wait10s_ElementLoaded(archive_dict["content_attachment"])
        Logging("View content modal is opened")
        view_archive = True
    except WebDriverException:
        Logging("Fail to open view modal")
        view_archive = False
        TestCase_LogResult(**view_fail)
    
    if view_archive == True:
        try:
            Commands.FindElement("//a[contains(., '%s')]" % pc_attach_name)
            Logging("PC attached file is saved successfully")
            Logging(objects.testcase_pass)
        except WebDriverException:
            view_fail.update({"description": "PC attached file is not saved"})
            TestCase_LogResult(**view_fail)
        
        if bool(cloud_attach_name) == True:
            try:
                Commands.FindElement("//a[contains(., '%s')]" % cloud_attach_name)
                Logging("CloudDisk attached file is saved successfully")
                Logging(objects.testcase_pass)
            except WebDriverException:
                view_fail.update({"description": "CloudDisk attached file is not saved"})
                TestCase_LogResult(**view_fail)
        
        Commands.SwitchToDefaultContent()
        
        try:
            Commands.Wait10s_ClickElement(archive_dict["view_close"])
            Logging("View Archive - Close view mode")
        except WebDriverException:
            pass

        time.sleep(1)

def Archive_ValidateNextPageList():
    PrintYellow("[MENU ARCHIVE] MOVE PAGE")
    
    List_ValidateListMovingPage(archive_dict["list_target"],
                                archive_dict["item_suf"],
                                archive_dict["page_total"],
                                archive_dict["nextpage_icon"])

def Archive_AccessCompanyArchive():
    PrintYellow("ACCESS COMPANY ARCHIVE")
    access_company_archive = None

    company_archive = Commands.FindElement(archive_dict["company_submenu"])
    company_attr = Functions.GetElementAttribute(archive_dict["company_submenu"], "class")
    if "open" not in company_attr:
        company_archive.click()
        Logging("Open Company Archive submenu")
    
    try:
        Waits.WaitElementLoaded(5, archive_dict["company_archive_submenu"])
        common_folders_xpath = str(archive_dict["company_archive_folder"]).replace("[position]", "")
        company_folders = Functions.GetListLength(common_folders_xpath)
    except WebDriverException:
        company_folders = 0
        ValidateUnexpectedModal()
    
    if company_folders > 0:
        i=0
        for i in range(0, company_folders):
            i+=1
            folder_xpath = str(archive_dict["company_archive_folder"]).replace("position", str(i)).replace("/i[contains(@class, 'fa-upload')]", "")
            folder = Commands.FindElement(folder_xpath)
            folder_name = folder.text
            upload_icon = "/i[contains(@class, 'fa-upload')]"
            share_icon = "/i[contains(@class, 'fa-share')]"
            try:
                Commands.FindElement(folder_xpath + upload_icon)
                Logging("Folder includes upload icon")
                access_company_archive = True
            except WebDriverException:
                try:
                    Commands.FindElement(folder_xpath + share_icon)
                    Logging("Folder includes share icon")
                    access_company_archive = True
                except WebDriverException:
                    Logging("Folder has no permission")
            
            if access_company_archive == True:
                folder.click()
                Logging("Access folder -> " + folder_name)

            Waits.WaitUntilPageIsLoaded(archive_dict["list_footer"])

            try:
                Waits.WaitElementLoaded(3, archive_dict["company_archive_item"])
            except WebDriverException:
                PrintYellow("No archive document found in the folder")
                access_company_archive = False
    
    return access_company_archive

def Archive_SearchCompanyArchive():
    PrintYellow("SEARCH COMPANY ARCHIVE")
    
    items = str(Functions.GetElementText(archive_dict["list_footer"])).replace(",", "").split(" ")[1]
    if int(items) > 1:
        creators = Commands.FindElements(archive_dict["creator"])
        creator_len = int(len(creators))
        
        creator_list = []
        i=1
        for i in range(1, creator_len):
            creator_name = creators[i].text
            if bool(creator_name) == True:
                creator_list.append(creator_name)
            i+=1

        creator_list = Functions.RemoveDuplicate_fromList(creator_list)
        if "Creator" in creator_list:
            creator_list.remove("Creator")
        
        if int(len(creator_list)) > 1:
            creator_search = True
        else:
            creator_search = False
            Logging("Cannot do search with creator in company archive")
            Logging(objects.testcase_fail)
    
        if creator_search == True:
            search_dict = {
                "creator": {
                    "key": creator_list[0],
                    "value": "username"
                }
            }
            search_archive_dict = dict(archive_dict["search_details"])
            search_archive_dict["search_dict"] = search_dict
            SearchDetailsBySelectBox(**search_archive_dict)

def Archive_ViewCompanyArchive():
    PrintYellow("[VIEW COMPANY ARCHIVE]")

    Waits.Wait10s_ElementLoaded(archive_dict["company_archive_item"])
    archives = Commands.FindElements(archive_dict["company_archive_item"])
    
    random_list = []
    i=0
    list_range = len(archives)
    for i in range(0, list_range):
        random_number = Functions.getRandomNumber_fromSpecificRange(0, list_range-1)
        random_list.append(random_number)
        random_list = Functions.RemoveDuplicate_fromList(random_list)
        if len(random_list) == 3:
            break
        i+=1

    for position_number in random_list:
        archive_title = archives[position_number].text
        print("--> archive_title" + str(archive_title))
        archives[position_number].click()

        Commands.SwitchToFrame("//*[@id='viewDetail']")
        
        try:
            Waits.Wait10s_ElementLoaded(archive_dict["view_archive_title"] % archive_title)
            Logging("Archive content of [" + str(archive_title) + "] is correct")
            Logging(objects.testcase_pass)
        except:
            Logging("Archive content of [" + str(archive_title) + "] is incorrect")
            Logging(objects.testcase_fail)
        
        Commands.SwitchToDefaultContent()

        Waits.Wait10s_ElementLoaded(archive_dict["view_close"])
        Commands.ClickElement(archive_dict["view_close"])
        time.sleep(1)
        
        try:
            Commands.ClickElement(archive_dict["view_close"])
            Logging("View Archive - Close view mode")
            time.sleep(1)
        except WebDriverException:
            Logging("View mode is closed successfully")

        time.sleep(1)

def ArchiveExecution(secure_pw):
    archive_data = AccessArchiveMeu(secure_pw)
    if bool(archive_data["access_folder"]) == True:
        new_archive = Archive_WriteArchiveDocument()
        print("---> new_archive: %s" % new_archive)
        Archive_ViewArchive(**new_archive)
    access_company_archive = Archive_AccessCompanyArchive()
    if bool(access_company_archive) == True:
        Archive_SearchCompanyArchive()
        Archive_ViewCompanyArchive()
        ValidateUnexpectedModal()