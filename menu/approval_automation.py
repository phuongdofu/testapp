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
    global approval_dict, approval_tc
    approval_dict = dict(data["approval"])
    approval_tc = dict(data["testcase_result"]["approval"])

def Approval_WriteApproval(domain_name, recipient_id):
    PrintYellow("[MENU APPROVAL] WRITE APPROVAL")

    Waits.Wait10s_ElementLoaded(approval_dict["list_footer"])
    Waits.Wait10s_ElementLoaded(approval_dict["write_button"])
    
    time.sleep(1)
    FindPushNoti()
    
    Commands.ClickElement(approval_dict["write_button"])
    Logging("Write new approval - Click Write approval button")

    Waits.WaitElementLoaded(20, approval_dict["approval_route"])
    time.sleep(1)

    ValidateUnexpectedModal()

    Waits.WaitUntilPageIsLoaded(page_xpath=approval_dict["approval_route"])

    approval_type = Functions.GetElementText(approval_dict["approval_type_selected"])
    Logging("approval_type: " +  approval_type)
    if "Implementation Route" in approval_type:
        Commands.ClickElement(approval_dict["agreement_checkbox"])
        Logging("Change to agreement route")
    else:
        pass
    
    Waits.WaitUntilPageIsLoaded(approval_dict["approver_org"])
    Commands.ClickElement(approval_dict["approver_org"])
    Logging("Write new approval - Click org tree button from approval route")

    time.sleep(1)

    CloseAutosave()
    Waits.Wait10s_ElementLoaded(approval_dict["org_input"])
    
    try:
        Commands.ClickElement(data["approval"]["org_input"])
        Commands.InputElement_2Values(data["approval"]["org_input"], recipient_id, Keys.RETURN)
        Logging("Input recipient key in search box")
        
        Waits.WaitUntilPageIsLoaded(None)
        Waits.Wait10s_ElementLoaded(expense_dict["user_org_xpath"])
        time.sleep(1)
        try:
            Commands.FindElement(expense_dict["recipient_xpath"])
            recipients = Commands.FindElements(expense_dict["recipient_xpath"])
        except WebDriverException:
            recipients = Commands.FindElements(expense_dict["recipient_block_xpath"])
        recipients[0].click()
        Logging("Select recipient")

        Commands.ClickElement(data["approval"]["org_plus"])
        Logging("Add recipient")

        Waits.Wait10s_ElementLoaded(expense_dict["2nd_approver"])
        Logging("Write new approval - Selected approver is visible in approval route")

        Commands.MoveToElement(data["approval"]["save_org_tree"])
        Commands.ClickElement(data["approval"]["save_org_tree"])
        Logging("Write new approval - Click Save button from Org tree - save selected approver")

        selected = True
    except WebDriverException:
        selected = False

    if bool(selected) == True:
        approval_subject = Commands.FindElement(approval_dict["subject"])
        approval_subject.clear()
        Commands.InputElement_2Values(approval_dict["subject"], objects.hanbiro_title, Keys.ENTER)
        Logging("Write new approval - Clear form name from approval subject - Input new approval subject")
        
        approval_name = Functions.GetInputValue(approval_dict["subject"])
        if approval_name != objects.hanbiro_title:
            approval_subject.clear()
            Commands.InputElement_2Values(approval_dict["subject"], objects.hanbiro_title, Keys.ENTER)

        FindPushNoti()

        Commands.ClickElement(approval_dict["submit_button"])
        Logging("Write new approval - Submit approval")

        Waits.Wait10s_ElementLoaded(approval_dict["apply_button"])
        time.sleep(1)
        Commands.ClickElement(approval_dict["apply_button"])
        Logging("Write new approval - Apply approval options")

        try:
            Waits.WaitElementLoaded(20, approval_dict["list_footer"])
            Waits.Wait10s_ElementLoaded(approval_dict["approval_name"] % objects.hanbiro_title)
            TestCase_LogResult(**approval_tc["write"]["pass"])
        except WebDriverException:
            TCResult_ValidateAlertMsg(menu="approval", testcase="write", msg="click save approval")
            TestCase_LogResult(**approval_tc["write"]["fail"])
            Commands.NavigateTo(domain_name + "/approval/list/progress/iall/")
    else:
        TestCase_LogResult(**approval_tc["write"]["fail"])
        approval_name = None
        
        try:
            Commands.ClickElement(approval_dict["close_org_tree"])
            Logging("Close org tree")
            Commands.ReloadBrowser(approval_dict["approval_route"])
            Waits.WaitElementLoaded(20, approval_dict["approval_route"])
            Waits.WaitUntilPageIsLoaded(None)
        except WebDriverException:
            pass

        Waits.Wait10s_ElementClickable(approval_dict["back_to_list"])
        time.sleep(1)

        Commands.ClickElement(approval_dict["back_to_list"])
        Logging("Cannot continue writing approval \n Back to approval list")

        Waits.WaitUntilPageIsLoaded(approval_dict["list_item"])

    time.sleep(1)
    
    return approval_name

def Approval_FindApprovalToView():
    Waits.WaitUntilPageIsLoaded(approval_dict["list_approval"])
    Waits.Wait10s_ElementLoaded(approval_dict["list_footer"])
    
    try:
        Waits.WaitElementLoaded(2, data["common"]["list_nodata"])
        view = False
        approval_tc["view"]["fail"].update({"description": "Cannot locate any approval in list"})
        TestCase_LogResult(**approval_tc["view"]["fail"])
    except WebDriverException:
        view = True
        print("view = True")
    
    if view == True:
        time.sleep(1)
        current_url = DefineCurrentURL()
        Logging(current_url)
        approval_urls = ["/official/", "/complete/"]
        for approval_url in approval_urls:
            if approval_url in current_url:
                public_submenu = True
            else:
                public_submenu = False

        if public_submenu == True:
            defined_item_xpath = approval_dict["defined_approval"] % objects.hanbiro_title
            undefined_item_xpath = approval_dict["list_item"].replace("/span[6]", "/span[5]")
        else:
            defined_item_xpath = approval_dict["approval_name"] % objects.hanbiro_title
            undefined_item_xpath = approval_dict["list_item"]
        
        Waits.WaitUntilPageIsLoaded(approval_dict["list_footer"])

        try:
            approval = Commands.FindElement(defined_item_xpath)
            approval_to_view = objects.hanbiro_title
        except WebDriverException:
            approval = Commands.FindElement(undefined_item_xpath)
            approval_to_view = Functions.GetElementText(undefined_item_xpath)

        approval.click()
        Logging("Click view approval: " + approval_to_view)

        try:
            Waits.Wait10s_ElementLoaded(approval_dict["document_frame"])
            TestCase_LogResult(**approval_tc["view"]["pass"])
            access_approval_view = True
        except WebDriverException:
            access_approval_view = False
            TCResult_ValidateAlertMsg(menu="approval", testcase="view", msg="click view approval")
            TestCase_LogResult(**approval_tc["view"]["fail"])
    else:
        access_approval_view = False

    time.sleep(1)

    return access_approval_view

def Approval_ViewDetails(domain_name):
    PrintYellow("[MENU APPROVAL] VIEW APPROVAL")
    
    time.sleep(2)

    access_approval_view = Approval_FindApprovalToView()

    if access_approval_view == True:
        info_results = []

        Commands.SwitchToFrame(approval_dict["document_frame"])
        
        try:
            try:
                Waits.Wait10s_ElementLoaded("//div[@class='annie-editor']")
                editor_xpath = "//div[@class='annie-editor']"
            except WebDriverException:
                editor_xpath = "//div[@id='HTML_SRC']"
                Commands.FindElement("//div[@id='HTML_SRC']")
            finally:
                try:
                    Commands.FindElement(editor_xpath + "/div")
                    Logging("[HTML] Approval Details - Approval content is displayed normally")
                    info_results.append(True)
                except WebDriverException:
                    Commands.FindElement(editor_xpath + "/p")
                    Logging("[Excel Form] Approval Details - Approval content is displayed normally")
                    info_results.append(True)
        except WebDriverException:
            info_results.append(False)
            approval_tc["view"]["fail"].update({"description": "Approval content is not displayed"})
            TestCase_LogResult(**approval_tc["view"]["fail"])
        
        Commands.SwitchToDefaultContent()

        time.sleep(1)

        approval_data = {
            "category": approval_dict["content_category"],
            "doc_no": approval_dict["content_doc"],
            "title": approval_dict["content_title"]
        }

        for approval_info in approval_data.keys():
            xpath = approval_data["" + approval_info + ""]
            info = Commands.FindElement(xpath)
            if bool(info.text) == True:
                info_results.append(True)
            else:
                info_results.append(False)
                approval_tc["view"]["fail"].update({"description": str(approval_info) + " is empty"})
                TestCase_LogResult(**approval_tc["view"]["fail"])

        if False not in info_results:
            approval_tc["view"]["fail"].update({"description": "All approval details are displayed"})
            TestCase_LogResult(**approval_tc["view"]["pass"])
    
    time.sleep(1)
    archive_folder = Approval_CopyArchive(domain_name)

    Commands.ClickElement(approval_dict["back_button"])
    Logging("Back to list")

    Waits.WaitUntilPageIsLoaded(approval_dict["list_footer"])

def Approval_Approve(domain_name, approval_name):
    PrintYellow("DRIVER2 - APPROVE APPROVAL")
    
    Commands.NavigateTo(domain_name + "/approval/list/progress/ireq/")
    Logging("Approval List - Access Received Approval list")

    access_approval_view = Approval_FindApprovalToView()
    if access_approval_view == True:
        time.sleep(1)
        
        try:
            Commands.ClickElement(approval_dict["approve_button"])
            Logging("Approval Approve - Click Decide button")

            Waits.Wait10s_ElementLoaded(approval_dict["send_mail"])
            Logging("Approval Approve - Wait until check box 'Send Email' is loaded")

            Commands.ClickElement(approval_dict["apply_button"])
            Logging("Approval Approve - Confirm approval with 'Approve option'")

            Waits.WaitUntilPageIsLoaded(None)
            Waits.Wait10s_ElementInvisibility(approval_dict["approve_button"])
        
            TestCase_LogResult(**approval_tc["approve"]["pass"])         
        except WebDriverException:
            TCResult_ValidateAlertMsg(menu="approval", testcase="approve", msg="click approve button")
            TestCase_LogResult(**approval_tc["approve"]["fail"])
        finally:
            Commands.ClickElement(approval_dict["back_button"])
            Waits.WaitUntilPageIsLoaded(None)
            time.sleep(1)

def Approval_OfficialDocumentation():
    PrintYellow("DRIVER1 - SAVE OFFICIAL DOCUMENTATION")
    time.sleep(1)
    
    Commands.ClickElement(approval_dict["completed_submenu"])
    Logging("Open Completed submenu")
    
    access_approval_view = Approval_FindApprovalToView()
    if access_approval_view == True:
        Commands.ClickElement("//button[contains(.,' Other')]")

        Waits.WaitElementLoaded(5, approval_dict["official_doc_href"])
        Commands.ClickElement(approval_dict["official_doc_href"])
        Logging("Approval Content - Select Official Documentation")

        Waits.Wait10s_ElementLoaded(approval_dict["officialform_select"])

        time.sleep(1)

        Commands.Selectbox_ByVisibleText(approval_dict["officialform_select"], "Default")
        Logging("Select form for official documentation")

        Waits.WaitElementLoaded(5, "//*[@id='executeConfig-header']")

        time.sleep(1)

        FindPushNoti()
        Commands.ClickElement(approval_dict["officialform_save"])
        Logging("Click to save documentation")
        
        Commands.SwitchToFrame(approval_dict["loading_doc_iframe"])

        doc_data = {
            "company": "Company Header",
            "info": "Approval Information",
            "main": "Form Content"
        }

        doc_results = []
        dict_view_pass = dict(approval_tc["view"]["pass"])
        dict_view_fail = dict(approval_tc["view"]["fail"])
        
        try:
            Waits.Wait10s_ElementLoaded(approval_dict["document_id"])
            doc_results.append(True)
        except WebDriverException:
            doc_results.append(False)
            dict_view_fail.update({"testcase": "View Official Documentation"})
            dict_view_fail.update({"description": "Documentation content is not displayed"})
            TestCase_LogResult(**dict_view_fail)
        
        for doc_item in doc_data.keys():
            item_name = doc_data["" + doc_item + ""]
            doc_item_xpath = "//*[@id='%s']" % doc_item
            try:
                Commands.FindElement(doc_item_xpath)
                doc_results.append(True)
            except WebDriverException:
                doc_results.append(False)
                dict_view_fail.update({"testcase": "View Official Documentation"})
                dict_view_fail.update({"description": str(item_name) + " is not displayed in official doc content"})
                TestCase_LogResult(**dict_view_fail)

        if False not in doc_results:
            dict_view_pass.update({"testcase": "View Official Documentation"})
            dict_view_pass.update({"description": "All doc data is displayed in official doc content"})
            TestCase_LogResult(**dict_view_pass)

        Commands.SwitchToDefaultContent()

        Commands.ClickElement(approval_dict["back_button"])
        Waits.Wait10s_ElementLoaded(approval_dict["list_footer"])

def Approval_CopyArchive(domain_name):
    Waits.Wait10s_ElementLoaded(approval_dict["content_title"])
    
    approval_name = str(Functions.GetElementText(approval_dict["content_title"])).strip()
    approval_doc_name = str(Functions.GetElementText(approval_dict["content_doc"])).strip()
    archived_name = "[" + approval_doc_name + "] " + approval_name
    Logging("archived_name " + archived_name)

    time.sleep(1)

    Waits.Wait10s_ElementLoaded(approval_dict["more_button"])
    Commands.ClickElement(approval_dict["more_button"])
    Logging("Approval details - Click More button")

    time.sleep(1)

    Waits.Wait10s_ElementLoaded(approval_dict["copy_archive_href"])
    Commands.ClickElement(approval_dict["copy_archive_href"])
    Logging("Approval details - Click Copy to Archive button")

    time.sleep(1)

    archive_folder = CopyArchive_SelectArchiveFolder()

    time.sleep(1)

    archive_data = {
        "archived_name": archived_name,
        "archive_folder": archive_folder
    }

    return archive_data

def Approval_SearchOfficialDocumentation(domain_name):
    PrintYellow("[TEST CASE] SEARCH APPROVAL (OFFICIAL DOCUMENTATION)")

    Commands.NavigateTo(domain_name + "/approval/list/official/default/")
    search_result = wrapper(searchInput, approval_dict["search_input"])
    if search_result == True:
        TestCase_LogResult(**approval_tc["search"]["pass"])
    else:
        TestCase_LogResult(**approval_tc["search"]["fail"])

def Approval_ValidateNextPageList():
    PrintYellow("[MENU APPROVAL] MOVE PAGE")
    List_ValidateListMovingPage(list_target = approval_dict["list_target"], 
                                item_suf = approval_dict["item_suf"], 
                                page_total_xpath = approval_dict["page_total"], 
                                nextpage_icon = approval_dict["nextpage_icon"])

def Approval_NavigateToList(domain_name):
    current_url = DefineCurrentURL()

    if "/approval/list/" not in current_url:
        Commands.NavigateTo(domain_name + "/approval/list/progress/ireq/")

    Waits.WaitUntilPageIsLoaded(approval_dict["list_item"])
    list = CollectListData(approval_dict["list_footer"], approval_dict["page_total"])
    list_data = {
        "approvals": list["total_items"],
        "pages": list["total_pages"]
    }

    return list_data

def Approval_SearchDetails():
    Waits.Wait10s_ElementLoaded(approval_dict["list_footer"])

    current_url = DefineCurrentURL()
    if "official/default" not in current_url:
        Commands.ClickElement(approval_dict["official_submenu"])
        Logging("Access Official Documentation")
        Waits.Wait10s_ElementLoaded(approval_dict["list_footer"])

    try:
        Waits.WaitElementLoaded(2, data["common"]["list_nodata"])
        list1 = 0
    except WebDriverException:
        list1 = DefineListLength(approval_dict["approval_div"]["item"])

    #list1 = DefineListLength(approval_dict["approval_div"]["item"])
    if list1 > 0:
        Waits.Wait10s_ElementLoaded(approval_dict["documentation_header"])
        details_xpath = approval_dict["search_details_button"]
        Commands.ClickElement(details_xpath + "/parent::a")
        Logging("Open search box")

        caret_up = "/following-sibling::i[contains(@class, 'fa-caret-up')]"
        Waits.Wait10s_ElementLoaded(details_xpath + caret_up)
        
        time.sleep(1)

        doc_data = {
            "approval": {
                "doc_no": approval_dict["approval_div"]["doc_no"],
                #"title": approval_dict["approval_div"]["title"],
                "drafter": approval_dict["approval_div"]["drafter"]
                
            },
            "search": {
                "doc_no": approval_dict["search_details"]["doc_no"],
                #"title": approval_dict["search_details"]["title"],
                "drafter": approval_dict["search_details"]["drafter"]
                
            }   
        }

        for approval_data in doc_data["approval"].keys():
            approval_xpath = doc_data["approval"][approval_data]
            text = Functions.GetElementText(approval_xpath)
            Logging("Key word for " + approval_data + " -> " + text)
            
            search_input_xpath = doc_data["search"][approval_data]
            Commands.Wait10s_InputElement(search_input_xpath, text)
            Logging("-> Input key word for " + approval_data)
            time.sleep(1)

        Commands.ClickElement(approval_dict["search_button"])
        Logging("Click Search button")

        if list1 > 1:
            i=0
            for i in range(0,10):
                i+=1
                time.sleep(1)
                list2 = DefineListLength(approval_dict["approval_div"]["item"])
                if list2 != list1:
                    search_result = True
                    break
                else:
                    search_result = False
        else:
            try:
                if "No Data" in Functions.GetPageSource():
                    Logging("List is empty")
                    search_result = False
            except:
                pass
        try:
            Commands.FindElement(approval_dict["error_page"])
            Logging("Page is error")
            TestCase_LogResult(**approval_tc["search"]["fail"])
        except WebDriverException:
            if search_result == True:
                TestCase_LogResult(**approval_tc["search"]["pass"])
            else:
                TestCase_LogResult(**approval_tc["search"]["fail"])

        list3 = DefineListLength(approval_dict["approval_div"]["item"])
        Commands.ClickElement(approval_dict["reload_button"])
        Logging("Click Reload button")
        i=0
        for i in range(0,10):
            i+=1
            time.sleep(1)
            list4 = DefineListLength(approval_dict["approval_div"]["item"])
            if list4 != list3:
                reset = True
                break
            else:
                reset = False
        if reset == True:
            Logging("Reset search result successfully")
        else:
            Logging("Fail to reset list")
    

def ApprovalExecution_Driver1(domain_name, recipient_id):
    PrintYellow("-------------------------------------------")
    PrintYellow("[MENU APPROVAL]")
    approval_name = Approval_WriteApproval(domain_name, recipient_id)
    Approval_ViewDetails(domain_name, approval_name)
    Approval_SearchOfficialDocumentation(domain_name)
    Approval_ValidateNextPageList
    
    return approval_name

def ApprovalExecution(domain_name, recipient_id):
    PrintYellow("-------------------------------------------")
    PrintYellow("[MENU APPROVAL]")
    approval_name = ApprovalExecution_Driver1(domain_name, recipient_id)
    '''
    try:
        approval_name = Approval_WriteApproval(domain_name, recipient_id)
    except:
        approval_name = None
    
    if bool(approval_name) == True:
        Approval_ViewDetails(domain_name, approval_name)

    Approval_SearchOfficialDocumentation(domain_name)

    Approval_ValidateNextPageList()
    '''
    Approval_Approve(domain_name, approval_name)

    Approval_OfficialDocumentation()

def Approval_Execution(domain_name, recipient_id):
    AccessGroupwareMenu(name = "approval,Approval", page_xpath = approval_dict["list_footer"])
    
    if recipient_id != data["tooltip"]["recipient"]:
        Approval_WriteApproval(domain_name, recipient_id)
    
    Approval_ViewDetails(domain_name)
    Approval_SearchDetails()
    #Approval_ValidateNextPageList()
    Approval_OfficialDocumentation()
    ValidateUnexpectedModal()