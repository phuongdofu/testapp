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
    global clouddisk_dict, clouddisk_tc
    clouddisk_dict = dict(data["clouddisk"])
    clouddisk_tc = dict(data["testcase_result"]["clouddisk"])

def CloudDisk_AccessMyFolder(domain_name):
    PrintYellow("-------------------------------------------")
    PrintYellow("[CLOUDDISK - ACCESS MY FOLDER]")

    Commands.NavigateTo(domain_name + "/clouddisk/list/my/my/")

    time.sleep(1)

    try:
        Waits.Wait10s_ElementLoaded(clouddisk_dict["list_footer"])
        access_clouddisk = True
    except WebDriverException:
        access_clouddisk = False

    Waits.WaitUntilPageIsLoaded(None)

    return access_clouddisk

def CloudDisk_CollectListData():
    list_data = CollectListData(list_footer=clouddisk_dict["list_footer"], page_total=clouddisk_dict["page_total"])
    clouddisks = list_data["total_items"]
    pages = list_data["total_pages"]
    clouddisk_list = {
        "clouddisks": clouddisks,
        "pages": pages
    }

    return clouddisk_list

def CloudDisk_GenerateFileImage():
    '''Create an image by taking screenshot and save it to local folder'''
    
    screenshot_name = "screenshot_{}.png".format(objects.date_id)
    Logging("screenshot_name:", screenshot_name)
    screenshot_location = Files.clouddisk_folder + screenshot_name
    Commands.SaveScreenShot(screenshot_location)

    clouddisk_file = {
        "name": screenshot_name,
        "location": screenshot_location
        }

    return clouddisk_file

def CloudDisk_SelectFile():
    time.sleep(1)
    
    list_length = Functions.GetListLength(clouddisk_dict["active_checkbox"])
    clouddisk_file = {}
    
    for file_position in range(3, list_length+1):
        file_position+=1
        Logging("location of the selected file:",str(file_position))
        try:
            folder_checkbox_xpath = clouddisk_dict["folder_checkbox"] % str(file_position)
            Commands.FindElement(folder_checkbox_xpath)
            Logging("CloudDisk folder has been found")
        except WebDriverException:
            selected_file_xpath = clouddisk_dict["select_file_checkbox"] % str(file_position)
            Commands.ClickElement(selected_file_xpath)
            Logging("Select file")
            
            selected_checkbox = Functions.GetElementAttribute(selected_file_xpath + "/input", "class")
            time.sleep(1)
            if "ng-not-empty" in selected_checkbox:
                Logging("File is selected successfully")  
            else:
                Commands.ClickElement(selected_file_xpath)
                Logging("File is not selected - Select file again")
            
            clouddisk_file["position"] = file_position
            file_name = Functions.GetElementText(selected_file_xpath + "/following-sibling::span[@class='summary']/a")
            clouddisk_file["name"] = file_name
            
            break
    
    time.sleep(1)
    
    return clouddisk_file

def CloudDisk_UnSelectFile(file_position):
    selected_file_xpath = clouddisk_dict["select_file_checkbox"] % str(file_position)
    selected_file = Functions.GetElementAttribute(selected_file_xpath + "/input", "class")
    if "ng-not-empty" in selected_file:
        Commands.ClickElement(selected_file_xpath)
        Logging("Un-select file")
    else:
        Logging("File is already un-selected")

def CloudDisk_UploadMyFolder():
    '''Upload new file to CloudDisk My Folder'''

    PrintYellow("[UPLOAD CLOUDDISK]")
    
    Logging("Upload My CloudDisk - Define list total number")
    list_total1 = CloudDisk_CountAllFiles()

    FindPushNoti()

    clouddisk_file = CloudDisk_GenerateFileImage()
    Logging("Create image file by taking a screenshot")

    Commands.Wait10s_ClickElement(clouddisk_dict["upload"])
    Logging("Upload My CloudDisk - Click Upload button in My Folder")

    Commands.Wait10s_InputElement(clouddisk_dict["file_container"], clouddisk_file["location"])
    Logging("Upload My CloudDisk - Collect file from local folder")

    try:
        Commands.FindElement(clouddisk_dict["attach_name"])
        Logging("Upload My CloudDisk - Local file is collected successfully")

        time.sleep(1)

        Commands.ClickElement(clouddisk_dict["save_button"])
        Logging("Upload My CloudDisk - Click Upload button (Send file)")

        time.sleep(1)
        try:
            Waits.WaitUntilPageIsLoaded(clouddisk_dict["complete_load"])   

            Waits.Wait10s_ElementClickable(clouddisk_dict["close_upload"])
            Commands.ClickElement(clouddisk_dict["close_upload"])
            
            Waits.Wait10s_ElementLoaded(clouddisk_dict["list_footer"])
            
            time.sleep(1)
            Commands.Wait10s_ClickElement(clouddisk_dict["reload"])
            Logging("Upload My CloudDisk - Reload list")

            Waits.WaitUntilPageIsLoaded(None)
            time.sleep(2)

            list_total2 = CloudDisk_CountAllFiles()
            Logging("list_total2", str(list_total2))
            list_result = ValidateListNumberUpdate(list_total1, list_total2)
            if list_result == True:
                TestCase_LogResult(**clouddisk_tc["upload"]["pass"])
            else:
                TestCase_LogResult(**clouddisk_tc["upload"]["fail"])
        except WebDriverException:
            TCResult_ValidateAlertMsg(menu="clouddisk", testcase="upload", msg="click upload button")
            TestCase_LogResult(**clouddisk_tc["upload"]["fail"])
    except WebDriverException:
        clouddisk_tc["upload"]["fail"].update({"description": "Fail to collect local file"})
        TestCase_LogResult(**clouddisk_tc["upload"]["fail"])

        Commands.ClickElement(clouddisk_dict["close_upload"])
        Logging("Upload My CloudDisk - Close Upload modal")
        Waits.Wait10s_ElementLoaded(clouddisk_dict["list_footer"])
    
    return clouddisk_file

def CloudDisk_SearchInput():
    PrintYellow("SEARCH INPUT")
    '''Search file from the screenshot image'''

    list_total1 = ValidateListTotal(clouddisk_dict["list_footer"])
    
    selected_file = CloudDisk_SelectFile()
    Commands.InputElement_2Values(clouddisk_dict["search_input"], selected_file["name"], Keys.ENTER)
    Logging("CloudDisk - Search Folder - Enter key words")

    time.sleep(2)

    list_total2 = ValidateListTotal(clouddisk_dict["list_footer"])
    if list_total1 != list_total2:
        TestCase_LogResult(**clouddisk_tc["search"]["pass"])
    else:
        TestCase_LogResult(**clouddisk_tc["search"]["fail"])

    Waits.WaitUntilPageIsLoaded(None)

    Commands.ClickElement(clouddisk_dict["my_folder_span"])

    Waits.WaitUntilPageIsLoaded(None)
    time.sleep(1)

def CloudDisk_CountAllFiles():
    '''Count the number of all files in the list '''

    list_total = ValidateListTotal(clouddisk_dict["list_footer"])
    list_number = int(list_total) + 1

    x=1
    version = []
    for x in range(1, list_number):
        try:
            version_number = int(Functions.GetElementText(clouddisk_dict["manage_version"] % (str(x))))
            version.append(version_number)
            Logging("file_version has been found with [x]", str(x))
            break
        except WebDriverException:
            pass
        
        x+=1
    
    file_total = sum(version) + list_number 
    Logging("file_total: ", str(file_total))

    return file_total

def CloudDisk_CreateWeblink():
    PrintYellow("CREATE WEBLINK")
    '''Create Weblink from the selected file '''

    selected_file = CloudDisk_SelectFile()

    Commands.Wait10s_ClickElement(clouddisk_dict["more_button"])
    Logging("Click More button")

    Commands.Wait10s_ClickElement(clouddisk_dict["weblink"])
    Logging("Select Weblink from dropdown")
    
    Commands.Wait10s_ClickElement(clouddisk_dict["create_weblink"])
    Logging("Click Create weblink button")

    try:
        Waits.WaitUntilPageIsLoaded(None)
        Waits.Wait10s_ElementLoaded(clouddisk_dict["new_weblink"])
        weblink = Commands.FindElement(clouddisk_dict["new_weblink"])
        TestCase_LogResult(**clouddisk_tc["weblink"]["pass"])
        close_weblink = True
    except WebDriverException:
        close_weblink = None
        TCResult_ValidateAlertMsg(menu="clouddisk", testcase="weblink", msg="click create weblink")
        TestCase_LogResult(**clouddisk_tc["weblink"]["fail"])

    if close_weblink == True:
        try:
            Commands.ClickElement(clouddisk_dict["close_modal"])
            Logging("Close weblink pop-up")
            Waits.Wait10s_ElementLoaded(clouddisk_dict["list_footer"])
        except WebDriverException:
            Logging("Cannot close weblink modal")
        
        CloudDisk_UnSelectFile(selected_file["position"])

        time.sleep(1)
    
    Waits.WaitUntilPageIsLoaded(clouddisk_dict["reload"])
    Commands.Wait10s_ClickElement(clouddisk_dict["reload"])
    Logging("Reload list")

    Waits.WaitUntilPageIsLoaded(None)
    time.sleep(1)
    
def CloudDisk_PreviewFile():
    PrintYellow("PREVIEW FILE")
    '''Preview the selected file from dropdown menu '''
    time.sleep(1)
    selected_file = CloudDisk_SelectFile()

    Commands.Wait10s_ClickElement(clouddisk_dict["more_button"])
    Logging("Click More button")

    Commands.Wait10s_ClickElement(clouddisk_dict["preview"])
    Logging("Select Weblink from dropdown")

    Waits.Wait10s_ElementLoaded(clouddisk_dict["preview_frame"])

    Commands.SwitchToFrame(clouddisk_dict["preview_frame"])

    Waits.Wait10s_ElementLoaded("//img")
    time.sleep(1)

    image = Commands.FindElement("//img")
    dimension = image.size
    width = str(dimension["width"])
    height = str(dimension["height"])
    Logging("width: {}/height: {}".format(width,height))
    #error image size: width = height = 16
    if (width + height) != 32:
        TestCase_LogResult(**clouddisk_tc["preview"]["pass"])
    else:
        TestCase_LogResult(**clouddisk_tc["preview"]["fail"])

    Commands.SwitchToDefaultContent()

    Commands.ClickElement(clouddisk_dict["close_modal"])
    Logging("Close Preview modal")

    Waits.Wait10s_ElementLoaded(clouddisk_dict["list_footer"])

    CloudDisk_UnSelectFile(selected_file["position"])

    time.sleep(1)
    Commands.Wait10s_ClickElement(clouddisk_dict["reload"])
    Logging("Reload list")

    Waits.WaitUntilPageIsLoaded(None)
    time.sleep(1)

def CloudDisk_CopyFile():
    PrintYellow("COPY FILE")
    Waits.WaitUntilPageIsLoaded(None)
    
    CloudDisk_SelectFile()
    Logging("Select file from list")

    try:
        folder_name = Functions.GetElementText(clouddisk_dict["clouddisk_folder"])
        Logging("folder_name: ",folder_name)

        folder_id = Functions.GetElementAttribute(clouddisk_dict["clouddisk_folder"], "id")
        Logging("folder_id: ",folder_id)
    except WebDriverException:
        try:
            Commands.ClickElement(clouddisk_dict["create_folder"])
        except WebDriverException:
            list_length = Functions.GetListLength(clouddisk_dict["clouddisk_items"])
            for i in range(0, list_length):
                i+=1
                try:
                    Commands.ClickElement(clouddisk_dict["selected_item"] % str(i))
                    unselect = True
                    break
                except WebDriverException:
                    unselect = False
            
            if unselect == True:
                time.sleep(1)
                Commands.ClickElement(clouddisk_dict["create_folder"])

        Logging("Click folder")
        Waits.Wait10s_ElementLoaded(clouddisk_dict["folder_name"])

        folder_name = "Copy Folder"
        Commands.InputElement(clouddisk_dict["folder_name"], folder_name)
        Logging("Input folder name")

        Commands.ClickElement(clouddisk_dict["save_folder"])
        Logging("Click save folder")

        Waits.Wait10s_ElementLoaded(clouddisk_dict["folder_li"] % folder_name)
        Logging("Copy folder is created successfully")

        folder_id = Functions.GetElementAttribute(clouddisk_dict["folder_id"] % folder_name, "id")
        time.sleep(1)
        
        CloudDisk_SelectFile()
        Logging("Select file from list")

    try:
        size1 = Functions.GetElementText(clouddisk_dict["folder_size"] % folder_id).strip()
        Logging("folder size before copy", size1)
    except WebDriverException:
        size1 = "0"
        Logging("folder is empty")
    
    Commands.Wait10s_ClickElement(clouddisk_dict["more_button"])
    Logging("Click More button")

    Commands.Wait10s_ClickElement(clouddisk_dict["copy"])
    Logging("Select Copy from dropdown")

    Waits.Wait10s_ElementLoaded(clouddisk_dict["folder_dropdown"])

    file_copy = Functions.GetElementText(clouddisk_dict["copied_file"]).strip().split(".")[0]
    Logging("file name for copy",file_copy)

    Commands.ClickElement(clouddisk_dict["folder_dropdown"])
    Logging("Click folder dropdown")

    Waits.Wait10s_ElementLoaded(clouddisk_dict["my_folder"])
    time.sleep(1)

    Commands.ClickElement(clouddisk_dict["my_folder"])
    Logging("Click My Folder right narrow")

    Commands.Wait10s_ClickElement(clouddisk_dict["dynatree_folder"] % (folder_name))
    Logging("Select folder '{}' for copy".format(folder_name))

    Waits.Wait10s_ElementInvisibility(clouddisk_dict["copy_disable"])
    time.sleep(1)

    Commands.ClickElement(clouddisk_dict["copy_button"])
    Logging("Click Copy button")

    Waits.WaitUntilPageIsLoaded(None)
    
    Commands.ReloadBrowser(None)
    time.sleep(2)
    # Waits.Wait10s_ElementLoaded(clouddisk_dict["reload"])
    # Commands.ClickElement(clouddisk_dict["reload"])
    Logging("Reload list")
    
    time.sleep(2)
    Waits.WaitUntilPageIsLoaded(None)

    size2 = Functions.GetElementText(clouddisk_dict["folder_size"] % (folder_id)).strip()
    Logging("folder size after copy",size2)
    if size2 != size1:
        TestCase_LogResult(**clouddisk_tc["copy"]["pass"])
    else:
        TestCase_LogResult(**clouddisk_tc["copy"]["fail"])

    Waits.WaitUntilPageIsLoaded(None)
    time.sleep(1)

def CloudDisk_LockFile():
    PrintYellow("LOCK FILE")
    '''Select file lock/unlock and check if file is locked/unlocked successfully '''

    selected_file = CloudDisk_SelectFile()
    Logging("Select file from list")

    Commands.Wait10s_ClickElement(clouddisk_dict["more_button"])
    Logging("Click More button")

    Commands.Wait10s_ClickElement(clouddisk_dict["lock"])
    Logging("Select File Lock from dropdown")

    Waits.WaitElementLoaded(20, data["common"]["loading_dialog"])
    time.sleep(2)

    file_lock_xpath = clouddisk_dict["select_file_icon"] % str(selected_file["position"])
    define_lock = Functions.GetElementAttribute(file_lock_xpath, "class")
    if "fa-ban" in define_lock:
        TestCase_LogResult(**clouddisk_tc["lock"]["pass"])
    else:
        TestCase_LogResult(**clouddisk_tc["lock"]["fail"])
    
    time.sleep(1)

    Commands.Wait10s_ClickElement(clouddisk_dict["more_button"])
    Logging("Click More button")

    Commands.Wait10s_ClickElement(clouddisk_dict["unlock"])
    Logging("Select File Unlock from dropdown")

    Waits.WaitUntilPageIsLoaded(None)
    time.sleep(1)

    file_unlock_xpath = clouddisk_dict["select_file_icon"] % str(selected_file["position"])
    define_unlock = Functions.GetElementAttribute(file_unlock_xpath, "class")
    if "fa-file-image-o" in define_unlock:
        TestCase_LogResult(**clouddisk_tc["unlock"]["pass"])
    else:
        TestCase_LogResult(**clouddisk_tc["unlock"]["fail"])

    CloudDisk_UnSelectFile(selected_file["position"])
    time.sleep(1)

def CloudDisk_ReloadPage():
    Commands.ReloadBrowser(clouddisk_dict["list_footer"])
    Waits.WaitUntilPageIsLoaded(clouddisk_dict["list_footer"])
    time.sleep(1)

def CloudDiskExecution(domain_name):
    access_clouddisk = CloudDisk_AccessMyFolder(domain_name)    
    CloudDisk_UploadMyFolder()
    if "qavn.hanbiro.net" not in domain_name:
        clouddisk_list = CloudDisk_CollectListData()
        CloudDisk_CreateWeblink()
        CloudDisk_PreviewFile()
    CloudDisk_SearchInput()
    if "qavn.hanbiro.net" not in domain_name:
        CloudDisk_CopyFile()
        CloudDisk_LockFile()
        CloudDisk_ReloadPage()
    ValidateUnexpectedModal()
