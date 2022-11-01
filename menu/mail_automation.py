import time, sys, unittest, random, json, requests, openpyxl, testlink
from selenium.webdriver.remote.webelement import WebElement
from datetime import datetime
from selenium import webdriver
from appium import webdriver
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
import testlink

def Dictionaries():
    global mail_dict, mail_tc
    mail_dict = dict(data["mail"])
    mail_tc = dict(data["testcase_result"]["mail"])

def Mail_WriteNewMail(domain_name):
    PrintYellow("[TEST CASE] WRITE MAIL")

    #Commands.NavigateTo(domain_name + "/mail/list/all/")
    AccessGroupwareMenu(name="mail,Mail", page_xpath=mail_dict["pen_button"])

    CommonWriteItem(mail_dict["pen_button"], mail_dict["title"], objects.hanbiro_content)
    Logging("Write mail - Click Create button \n Input subject and content")

    mail_subject = Functions.GetInputValue(mail_dict["title"])
    if mail_subject != objects.hanbiro_title:
        Commands.InputElement(mail_dict["title"], objects.hanbiro_title)

    Commands.ExecuteScript("window.scrollTo(-501,0)")
    Logging("Scroll up")

    FindPushNoti()
    Commands.ClickElement(mail_dict["compose_to_me"])
    Logging("Write mail - Click 'Compose to me' button")

    time.sleep(1)

    # After clicking Compose to me button
    # continue checking if recipient is selected (mail recipient = current user's mail address)
    Waits.Wait10s_ElementLoaded(mail_dict["recipient_tag"])
    selected_recipients = Functions.GetListLength(mail_dict["recipient_tag"])

    if selected_recipients == 1:
        global selected_recipient
        selected_recipient = Functions.GetElementText(mail_dict["recipient_tag"])
        Logging("Write mail - Select mail address: Compose to me")

        Commands.ClickElement(mail_dict["send_button"])
        Logging("Write mail - Click Send mail button")
        
        try:
            Waits.Wait10s_ElementLoaded(mail_dict["list_footer"])
            TestCase_LogResult(**mail_tc["send"]["pass"])
            # <<<< TESTLINK REPORT >>>
            #TestLinkReport(testcase_id="GW-1", executed_status="Pass", domain_name=domain_name)
        except WebDriverException:
            TCResult_ValidateAlertMsg(menu="mail", testcase="send", msg="click send mail")
            TestCase_LogResult(**mail_tc["send"]["fail"])
            # <<<< TESTLINK REPORT >>>
            #TestLinkReport(testcase_id="GW-1", executed_status="Fail", domain_name=domain_name)
    else:
        compose_to_me_fail = mail_tc["send"]["fail"].update({"description": "Fail to selected mail address: Compose to me"})
        TestCase_LogResult(**compose_to_me_fail)
        # <<<< TESTLINK REPORT >>>
        #TestLinkReport(testcase_id="GW-1", executed_status="Fai", domain_name=domain_name)

def Mail_SendFailLog(**faillog_config):
    Commands.NavigateTo(faillog_config["domain_name"] + "/mail/list/all/")
    try:
        CommonWriteItem(mail_dict["pen_button"], mail_dict["title"], faillog_config["mail_content"])
    except WebDriverException:
        Commands.NavigateTo(faillog_config["domain_name"] + "/mail/list/all/")
        CommonWriteItem(mail_dict["pen_button"], mail_dict["title"], faillog_config["mail_content"])
    Logging("Write mail - Click Create button \n Input subject and content")

    mail_subject = Commands.FindElement(mail_dict["title"])
    mail_subject.clear()
    time.sleep(1)
    mail_subject.send_keys(faillog_config["mail_title"])

    Commands.ExecuteScript("window.scrollTo(-501,0)")
    Logging("Scroll up")

    time.sleep(1)

    Commands.ClickElement(mail_dict["recipient_span"])
    Logging("Focus auto-complete")

    Commands.InputElement_2Values(mail_dict["recipient_input"], faillog_config["recipients"], Keys.RETURN)
    Logging("Input recipient email address")

    Waits.Wait10s_ElementLoaded(mail_dict["recipient_tag"])
    
    Commands.ClickElement(data["attachment"]["attach_button"])
    Logging("PC Attachment - Click Attach file button")

    for attachment in faillog_config["files"]:
        Commands.InputElement(mail_dict["file_uploader"], attachment)
        Logging("attachment: " + str(attachment))
        Logging("PC Attachment - Collect file from local folder")

    Commands.ClickElement(mail_dict["send_button"])
    Logging("Write mail - Click Send mail button")

    Waits.Wait10s_ElementLoaded(mail_dict["list_footer"])

def Mail_ViewMail():
    PrintYellow("[MENU MAIL] CHECK NOTIFICATION AND VIEW MAIL CONTENT")
    
    try:
        Notification_CheckReceive(objects.hanbiro_title, mail_dict["leftside_badge"], mail_dict["unread_badge"])
        Logging("Notification - Mail notification is delivered")
        TestCase_LogResult(**mail_tc["notification"]["pass"])
        
        Commands.ClickElement(data["common"]["linktext"] % objects.hanbiro_title) 
        Logging("Notification - Click to view incoming notification")
        
        Waits.Wait10s_ElementLoaded(mail_dict["content_subject"])
        Logging("View Mail - Access mail content")

        content_subject = Functions.GetElementText(mail_dict["content_subject"])
        if content_subject == objects.hanbiro_title:
            mail_tc["view"]["pass"].update({"description": "View content from notification successfully"})
            TestCase_LogResult(**mail_tc["view"]["pass"])
        elif content_subject == "None":
            mail_tc["view"]["fail"].update({"description": "Fail to view content from notification"})
            TestCase_LogResult(**mail_tc["view"]["fail"])
            # if API is error -> Mail subject is recognized as None    
    except WebDriverException:
        TestCase_LogResult(**mail_tc["notification"]["fail"])
        Waits.Wait10s_ElementLoaded(mail_dict["list_item"])
    
        item_position = DefineItemPosition(mail_dict["search_input"][1])
        Logging("View Content - Define item position")

        item_text = Functions.GetElementText(mail_dict["list_mail"] % str(item_position))
        Logging(item_text)

        Commands.ClickElement(mail_dict["list_mail"] % str(item_position))
        Logging("View Content - Click on item")
        
        Waits.Wait10s_ElementLoaded(mail_dict["content_subject"])
        if item_text in Functions.GetPageSource():
            mail_tc["view"]["pass"].update({"description": "View content from mail list successfully"})
            TestCase_LogResult(**mail_tc["view"]["pass"])
        else:
            mail_tc["view"]["fail"].update({"description": "Fail to view content from  maillist"})
            TestCase_LogResult(**mail_tc["view"]["fail"])

    FindPushNoti()
    
def Mail_ReplyMail():
    PrintYellow("[MENU MAIL] REPLY MAIL")
    try:
        Waits.Wait10s_ElementLoaded(mail_dict["content_view_body"])
        list = False
    except WebDriverException:
        Commands.FindElement(mail_dict["list_mail_title"])
        list = True

    if list == True:
        mail_subject = Functions.GetElementText(mail_dict["list_mail_title"])
        
        Commands.ClickElement(mail_dict["list_mail_title"])
        Logging("Click mail:", mail_subject)

        Waits.Wait10s_ElementLoaded(mail_dict["content_subject"])
    
    Commands.Wait10s_ClickElement(mail_dict["reply_button"])
    Logging("Click Reply button")

    Waits.WaitElementLoaded(15, data["editor"]["tox_iframe"])
    time.sleep(2)

    Waits.Wait10s_ElementLoaded(mail_dict["title"])
    subject_reply = Functions.GetInputValue(mail_dict["title"])
    Logging("subject_reply", subject_reply)
    if "RE:" in subject_reply:
        Logging("Reply Mail - Mail subject is appended with prefix RE:")
    else:
        mail_tc["reply"]["fail"].update({"description": "Prefix of Reply is not appended to mail subject"})
        TestCase_LogResult(**mail_tc["reply"]["fail"])

    recipient_tag = Functions.GetListLength(mail_dict["recipient_tag"])
    Logging("recipient_tag", str(recipient_tag))
    if recipient_tag == 1:
        Logging("Reply Mail - Recipient for reply is collected successfully")
    else:
        mail_tc["reply"]["fail"].update({"description": "Fail to collect recipient for reply"})
        TestCase_LogResult(**mail_tc["reply"]["fail"])

    Commands.SwitchToFrame(data["editor"]["tox_iframe"])
    try:
        Waits.Wait10s_ElementLoaded(mail_dict["block_quote"])
        Logging("Reply Mail - Reply Content is appended successfully")
    except WebDriverException:
        mail_tc["reply"]["fail"].update({"description": "Fail to append reply content"})
        TestCase_LogResult(**mail_tc["reply"]["fail"])
    Commands.SwitchToDefaultContent()

    Commands.ClickElement(mail_dict["send_button"])
    Logging("Write mail - Click Send mail button")

    try:
        Waits.Wait10s_ElementLoaded(mail_dict["content_subject"])
        TestCase_LogResult(**mail_tc["reply"]["pass"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="mail", testcase="reply", msg="click send reply")
        TestCase_LogResult(**mail_tc["reply"]["fail"])

    Commands.ClickElement(mail_dict["back_button_2"])
    Logging("Back to mail list")

    Waits.Wait10s_ElementLoaded(mail_dict["list_item"])

def Mail_ForwardMail():
    PrintYellow("[MENU MAIL] FORWARD MAIL")
    try:
        Waits.Wait10s_ElementLoaded(mail_dict["list_mail_title"])

        mail = Commands.FindElement(mail_dict["list_mail_title"])
        Logging("Click mail:", str(mail.text))
        mail.click()
        
        Waits.Wait10s_ElementLoaded(mail_dict["content_subject"])
    except WebDriverException:
        pass
    finally:
        Commands.Wait10s_ClickElement(mail_dict["fwd_button"])

    Waits.WaitElementLoaded(15, data["editor"]["tox_iframe"])
    time.sleep(1)
    
    Waits.Wait10s_ElementLoaded(mail_dict["title"])
    subject_forward = Functions.GetInputValue(mail_dict["title"])
    Logging("subject_forward", subject_forward)
    
    if "FW:" in subject_forward:
        Logging("Forward Mail - Mail subject is appended with prefix FW:")
    else:
        Logging("Forward Mail - Prefix of Forward is not appended to mail subject")

    recipient_tag = Functions.GetListLength(mail_dict["replied_recipients"])
    if recipient_tag == 0:
        Logging("Forward Mail - Recipient for forward is empty")
    else:
        mail_tc["forward"]["fail"].update({"description": "Recipient for forward is not empty"})
        TestCase_LogResult(**mail_tc["forward"]["fail"])

    Commands.SwitchToFrame(data["editor"]["tox_iframe"])
    try:
        Waits.Wait10s_ElementLoaded(mail_dict["block_quote"])
        Logging("Forward Mail - Forward Content is appended successfully")
    except WebDriverException:
        mail_tc["forward"]["fail"].update({"description": "Fail to append forward content"})
        TestCase_LogResult(**mail_tc["forward"]["fail"])
    Commands.SwitchToDefaultContent()

    Commands.ClickElement(mail_dict["compose_to_me"])
    Logging("Write mail - Click 'Compose to me' button")

    Commands.ClickElement(mail_dict["send_button"])
    Logging("Write mail - Click Send mail button")

    Waits.Wait10s_ElementLoaded(mail_dict["content_subject"])
    time.sleep(1)
    Waits.WaitUntilPageIsLoaded(None)
    
    Commands.ClickElement(mail_dict["back_button2"])
    Logging("Write mail - Back to mail list")

    try:
        Waits.Wait10s_ElementLoaded(mail_dict["list_footer"])
        TestCase_LogResult(**mail_tc["forward"]["pass"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="mail", testcase="forward", msg="click send forward")
        TestCase_LogResult(**mail_tc["forward"]["fail"])

def Mail_SearchMail():
    PrintYellow("[MENU MAIL] SEARCH MAIL")
    search_result = wrapper(searchInput, mail_dict["search_input"])
    if search_result == True:
        TestCase_LogResult(**mail_tc["search"]["pass"])
    else:
        TestCase_LogResult(**mail_tc["search"]["fail"])

def Mail_ValidateNextPageList():
    PrintYellow("[MENU MAIL] MOVE PAGE")
    time.sleep(1)
    List_ValidateListMovingPage(mail_dict["list_item"], mail_dict["item_suf"], mail_dict["page_total"], mail_dict["nextpage_icon"])

def Mail_NavigateToList(domain_name):
    Commands.NavigateTo(domain_name + "/mail/list/all/")

    Waits.WaitUntilPageIsLoaded(mail_dict["list_item"])

    mail_total = int(Functions.GetElementText(mail_dict["list_footer"]).split(" ")[1].replace(",", ""))
    Logging("mail_total:", str(mail_total))

    page_total = int(Functions.GetElementText(mail_dict["page_total"]))
    Logging("page_total:", str(page_total))
    
    list_data = {
        "mail_total": mail_total,
        "page_total": page_total
    }

    return list_data

def Mail_ReportSpam(domain_name, spam_category):
    mail_spam = {
        "report_spam": "/mail/list/all",
        "report_not_spam": "/mail/list/Spam"
    }
    
    current_url = DefineCurrentURL()
    if mail_spam[spam_category] not in current_url:
        if spam_category == "report_spam":
            submenu_xpath = "//span[@class='menu-text' and text()='All Mail']"
            submenu_name = "All Mail"
        else:
            submenu_xpath = mail_dict["spam_submenu"]
            submenu_name = "Spam"

        Commands.ClickElement(submenu_xpath)
        print("Access sub menu " + submenu_name)
    else:
        submenu_name = spam_category

    current_url = DefineCurrentURL()
    if "/list/all/" in current_url:
        mail_xpath = mail_dict["list_mail_span"] % "all"
    else:
        mail_xpath = mail_dict["list_mail_span"] % "Spam"
    list_length1 = DefineListLength(xpath=mail_xpath)
    print("list_length1: " + str(list_length1))
    report_spam = None
    if list_length1 > 0:
        Waits.Wait10s_ElementLoaded(mail_dict["list_mail_div"])

        mail_text = Functions.GetElementText(mail_dict["list_mail_div"])
        Logging("Mail [" + mail_text + "] is selected")
        
        Commands.ClickElement(mail_dict["list_mail_checkbox"])
        Logging("Click on checkbox of target mail")

        if spam_category == "report_spam":
            button_xpath = "Report Spam"
        else:
            button_xpath = "Not Spam"

        try:
            Waits.WaitElementLoaded(3, data["common"]["span"] % button_xpath)
            report_spam = True
        except WebDriverException:
            report_spam = False

        if report_spam == True:
            '''try:
                Commands.ClickElement(data["common"]["span"] % button_xpath)
                Logging("Click button", spam_category)

                Waits.WaitElementLoaded(3, data["common"]["ok_button"])
                Commands.ClickElement(data["common"]["ok_button"])
                Logging("-> Confirm the function", spam_category)

                time.sleep(2)
                
                if spam_category == "report_spam":
                    Commands.ClickElement(mail_dict["spam_submenu"])
                    list_footer_all = str(mail_dict["list_footer"]).replace("list.mail", "list.mail_Spam")
                    Waits.Wait10s_ElementLoaded(list_footer_all)
                else:
                    list_footer_spam = str(mail_dict["list_footer"]).replace("list.mail", "list.mail_all")
                    Commands.ClickElement(mail_dict["allmail_submenu"])
                    Waits.Wait10s_ElementLoaded(list_footer_spam)
                print("Access sub menu " + submenu_name)
                
                list_length2 = DefineListLength(xpath=mail_xpath)
                print("list_length2: " + str(list_length2))
                
                if list_length2 < list_length1:
                    TestCase_LogResult(**mail_tc[spam_category]["pass"])
                    Logging(objects.testcase_pass)
                    report_spam = True
                else:
                    report_spam = False
                    TestCase_LogResult(**mail_tc[spam_category]["fail"])
                    Logging(objects.testcase_fail)
            except WebDriverException:
                dict(mail_tc[spam_category]["fail"]).update({"description": "Fail to click button " + spam_category})
                TestCase_LogResult(**mail_tc[spam_category]["fail"])
                Logging(objects.testcase_fail)
                report_spam = False'''
            
            Commands.ClickElement(data["common"]["span"] % button_xpath)
            Logging("Click button", spam_category)

            try:
                Waits.WaitElementLoaded(3, data["common"]["ok_button"])
                Commands.ClickElement(data["common"]["ok_button"])
                Logging("-> Confirm the function", spam_category)
            except WebDriverException:
                pass

            time.sleep(2)
            
            if spam_category == "report_spam":
                Commands.ClickElement(mail_dict["spam_submenu"])
                list_footer_all = str(mail_dict["list_footer"]).replace("list.mail", "list.mail_Spam")
                Waits.Wait10s_ElementLoaded(list_footer_all)
            else:
                list_footer_spam = str(mail_dict["list_footer"]).replace("list.mail", "list.mail_all")
                Commands.ClickElement(mail_dict["allmail_submenu"])
                Waits.Wait10s_ElementLoaded(list_footer_spam)
            print("Access sub menu " + submenu_name)
            
            list_length2 = DefineListLength(xpath=mail_xpath)
            print("list_length2: " + str(list_length2))
            
            if list_length2 < list_length1:
                TestCase_LogResult(**mail_tc[spam_category]["pass"])
                Logging(objects.testcase_pass)
                report_spam = True
            else:
                report_spam = False
                TestCase_LogResult(**mail_tc[spam_category]["fail"])
                Logging(objects.testcase_fail)

    time.sleep(1)

    return report_spam      

def Mail_SearchDetails():
    Waits.Wait10s_ElementLoaded(mail_dict["list_footer"])
    list_footer_all = str(mail_dict["list_footer"]).replace("list.mail", "list.mail_all")
    current_url = DefineCurrentURL()
    if "/mail/list/all/" not in current_url:
        Commands.ClickElement(mail_dict["allmail_submenu"])
        Logging("Access All Mail sub-menu")
        Waits.Wait10s_ElementLoaded(list_footer_all)

    list1 = DefineListLength(mail_dict["mail_div"]["row_item"])
    search_result = None
    if list1 > 0:
        Commands.ClickElement(mail_dict["search_details"])
        Logging("Open search box")

        Waits.Wait10s_ElementLoaded(mail_dict["search"]["sender"])

        sender_title = Functions.GetElementAttribute(mail_dict["mail_div"]["sender"], "title")
        Logging("Target Sender address:", sender_title)

        Commands.InputElement(mail_dict["search"]["sender"], sender_title)
        Logging("Input sender address in search box")

        Commands.Selectbox_ByValue(mail_dict["search"]["content_select"], "string:s")
        Logging("Select label Subject for search")

        mail_subject = Functions.GetElementText(mail_dict["mail_div"]["subject"])
        Logging("Target Mail subject:", mail_subject)

        Commands.InputElement(mail_dict["search"]["content_input"], mail_subject)
        Logging("Input mail subject in search box")

        try:
            Commands.FindElement(mail_dict["mail_div"]["attachment_icon"])
            attachment = True
        except WebDriverException:
            attachment = False

        if attachment == True:
            Commands.ClickElement(mail_dict["search"]["attachment"])
            Logging("Click to search with attachment")
        
        mail = Functions.GetElementAttribute(mail_dict["mail_div"]["row_item"] + "[1]", "class")
        if "message-unread" in mail:
            Commands.Selectbox_ByVisibleText(mail_dict["search"]["status_select"], "New")
            Logging("Search as unread mail")
        
        mail_box_label = Functions.GetElementText(mail_dict["mail_div"]["mail_box"])
        Logging("Mail box for search " + mail_box_label)
        
        Commands.ClickElement(mail_dict["search"]["mail_box_open"])
        Logging("Open folder list")
        
        mail_box_xpath = str(mail_dict["search"]["mail_box_name"]).replace("folder_name", mail_box_label)
        Commands.Wait10s_ClickElement(mail_box_xpath)
        Logging("Search mail with folder")
        
        i=0
        for i in range(0,5):
            i+=1
            time.sleep(1)
            folder_checkbox = Functions.GetElementAttribute(mail_box_xpath + "/parent::span", "class")
            if "dynatree-selected" in folder_checkbox:
                select_folder = True
                break
            else:
                select_folder = False
        if select_folder == True:
            Logging("Folder is selected successfully")
        else:
            Logging("Fail to select folder for searching")
        
        Commands.ClickElement(mail_dict["search"]["mail_box_open"])
        print("Close folder tree")

        Waits.Wait10s_ElementClickable(mail_dict["search"]["search_button"])
        Commands.ClickElement(mail_dict["search"]["search_button"])
        Logging("Click Search button")

        if list1 > 1:
            i=0
            for i in range(1,10):
                i+=1
                time.sleep(1)
                list2 = DefineListLength(mail_dict["mail_div"]["row_item"])
                if list2 != list1:
                    search_result = True
                    Logging("Search successfully")
                    break
                else:
                    Logging("List does not change")
                    search_result = False
        else: #list1 == 1
            try:
                "No Data" in Functions.GetPageSource()
                Logging("List is empty")
                search_result = False
            except:
                pass
    
    try:
        Commands.FindElement(data["common"]["page_error"])
        Logging("Page is error")
        TestCase_LogResult(**mail_tc["search"]["fail"])
    except WebDriverException:
        if search_result == True:
            TestCase_LogResult(**mail_tc["search"]["pass"])
            print(objects.testcase_pass)
        else:
            Logging("Cannot define search result " + objects.testcase_fail)
            TestCase_LogResult(**mail_tc["search"]["fail"])
        
        list3 = DefineListLength(mail_dict["mail_div"]["row_item"])
        Commands.ClickElement(mail_dict["search"]["reset_button"])
        Logging("Click Reset search box")
        
        i=0
        for i in range(0,10):
            i+=1
            time.sleep(1)
            list4 = DefineListLength(mail_dict["mail_div"]["row_item"])
            if list4 != list3:
                reset = True
                break
            else:
                reset = False
        if reset == True:
            Logging("Reset search result successfully")
        else:
            Logging("Fail to reset list")

def Mail_ValidateInboxCounter():
    Waits.Wait10s_ElementLoaded(mail_dict["inbox_counter"])
    counter_number = int(Functions.GetElementText(mail_dict["inbox_counter"]).replace(",", ""))

    return counter_number

def Outlook_StartWinAppDriver():
    global win_driver
    desired_caps = {}
    desired_caps["app"] = "C:\\Program Files\\Microsoft Office\\Office16\\OUTLOOK.EXE"
    win_driver = webdriver.Remote(
        command_executor='http://127.0.0.1:4723',
        desired_capabilities=desired_caps)

def Outlook_AccessOutlook():
    try:
        WebDriverWait(win_driver, 10).until(EC.presence_of_element_located((By.NAME, "Inbox")))
        print("Access Outlook successfully")
        access_outlook = True
    except WebDriverException:
        print("Fail to access Outlook")
        access_outlook = False
    
    return access_outlook

def Outlook_ViewWebMail():
    '''
        Wait for the email which was sent from web automation test
        Check this incoming mail if it's displayed in Inbox of Outlook
    '''
    mail_subject = objects.hanbiro_title
    try:
        WebDriverWait(win_driver, 60).until(EC.presence_of_element_located((By.NAME, mail_subject)))
        TestCase_LogResult(**mail_tc["outlook_receive"]["pass"])
    except WebDriverException:
        TestCase_LogResult(**mail_tc["outlook_receive"]["fail"])

def Outlook_SendOutlookMail():
    import win32com.client as win32
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    
    sender_address = selected_recipient
    mail_title = objects.hanbiro_title
    mail_content = objects.hanbiro_content

    mail.To = sender_address
    mail.Subject = mail_title
    mail.body = mail_content
    mail.send

def Outlook_ViewOutlookMail(previous_counter_numer):
    counter_update_xpath = mail_dict["inbox_counter_number"].replace("counter_number", str(previous_counter_numer + 1))
    
    try:
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, counter_update_xpath)))
        TestCase_LogResult(**mail_tc["outlook_send"]["pass"])
    except WebDriverException:
        TestCase_LogResult(**mail_tc["outlook_send"]["fail"])

def MailExecution(domain_name):
    Mail_WriteNewMail(domain_name)
    Mail_ViewMail()
    Mail_ReplyMail()
    Mail_ForwardMail()
    Mail_ReportSpam(domain_name=domain_name, spam_category="report_spam")
    Mail_ReportSpam(domain_name=domain_name, spam_category="report_not_spam")
    Mail_SearchDetails()
    Mail_ValidateNextPageList()
    ValidateUnexpectedModal()
    check_counter = Mail_ValidateInboxCounter()
    '''Outlook_StartWinAppDriver()
    access_outlook = Outlook_AccessOutlook()
    Outlook_ViewWebMail()
    Outlook_SendOutlookMail()
    Outlook_ViewOutlookMail(check_counter)'''
