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
    global whisper_dict, whisper_tc
    whisper_dict = dict(data["whisper"])
    whisper_tc = dict(data["testcase_result"]["whisper"])

def Whisper_SendWhisperWithRecipients(org_key, user1, user2, user3):
    Waits.Wait10s_ElementLoaded("//div[@id='div-sender']/label/a")
    Commands.ClickElement("//div[@id='div-sender']/label/a")
    Logging("Write Whisper - Click Org button")

    Waits.Wait10s_ElementLoaded(whisper_dict["org_input"])
    
    org_input = Commands.FindElement(whisper_dict["org_input"])
    time.sleep(1)
    org_input.click()
    org_input.send_keys(org_key)
    org_input.send_keys(Keys.RETURN)
    
    Commands.Wait10s_ClickElement("//a[text()='" + user1 + "']")
    Commands.Wait10s_ClickElement("//a[text()='" + user2 + "']")
    Commands.Wait10s_ClickElement("//a[text()='" + user3 + "']")
    Logging("Write Whisper - Select recipients from organization")

    Commands.ClickElement(whisper_dict["add_user"])
    Logging("Write Whisper - Select selected recipients")

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".listbox")))

    listbox = Functions.GetListLength(whisper_dict["listbox"])

    if listbox == 3:
        Logging("Write Whisper - 3 recipients are selected")
        Logging(objects.testcase_pass)
    
    Commands.ClickElement(whisper_dict["save_org"])
    Logging("Write Whisper - Save organization tree")

    Waits.Wait10s_ElementLoaded(whisper_dict["selected_org"])

    whisper_recipients = Functions.GetListLength(whisper_dict["selected_org"] + "[contains(@class, 'tag')]")
    if whisper_recipients == 3:
        Logging("Write Whisper - Whisper recipients are selected")
        Logging(objects.testcase_pass)

def Whisper_SendNewWhisper(domain_name, user_id):
    PrintYellow("[WRITE WHISPER]")
    try:
        Waits.WaitElementLoaded(5, whisper_dict["list_footer"])
    except WebDriverException:
        Commands.ClickElement("//a[@data-name='whisper,Whisper']")
        Waits.Wait10s_ElementLoaded(whisper_dict["list_footer"])

    counter1 = Counter_CheckCounterNumber(whisper_dict["top_counter"], whisper_dict["left_counter"])

    Waits.WaitUntilPageIsLoaded(None)

    FindPushNoti()

    Commands.ClickElement(whisper_dict["pen_button"])
    Logging("Write Whisper - Click Create button")

    Waits.WaitUntilPageIsLoaded(None)
    Waits.WaitElementLoaded(15, "//*[@class='tox-edit-area__iframe']")

    '''current_url = DefineCurrentURL()
    domain_name = current_url.split("/ngw/")[0].split("//")[1].split(".")[0]
    test_domain_list = ["groupware57", "dofu", "qavn", "global3"]
    if domain_name in test_domain_list:
        Whisper_SendWhisperWithRecipients(org_key, user1, user2, user3)
    else:
        Autocomplete_SelectRecipient(user_id, whisper_dict["search_placeholder"])'''
    
    Autocomplete_SelectRecipient(user_id, whisper_dict["search_placeholder"])

    time.sleep(1)

    Commands.SwitchToFrame("//*[@class='tox-edit-area__iframe']")
    
    content = Commands.FindElement("//*[@id='tinymce']/p")
    content.clear()
    time.sleep(1)
    content.send_keys(objects.hanbiro_content)
    
    Commands.SwitchToDefaultContent()
    Logging("Write Whisper - Input content")

    CloseAutosave()
    FindPushNoti()

    Commands.ClickElement(whisper_dict["send"])
    Logging("Write - Send whisper")
    
    try:
        Waits.WaitElementLoaded(5, "//a[text()='" + objects.hanbiro_content + "']")
        TestCase_LogResult(**whisper_tc["write"]["pass"])
        
        time.sleep(1)

        counter2 = Counter_CheckCounterNumber(whisper_dict["top_counter"], whisper_dict["left_counter"])
        if counter2 != counter1:
            Logging("Receive counter - Receive counter is updated successfully")
        else:
            whisper_tc["write"]["fail"].update({"description": "Unread counter is not updated"})
            TestCase_LogResult(**whisper_tc["write"]["fail"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="whisper", testcase="write", msg="send whisper")
        TestCase_LogResult(**whisper_tc["write"]["fail"])
        Commands.NavigateTo(domain_name + "/whisper/list/inbox/")
    
    return objects.hanbiro_content

def Whisper_ViewDetails(hanbiro_content):
    Waits.Wait10s_ElementLoaded(whisper_dict["list_item"])

    counter1 = Counter_CheckCounterNumber(whisper_dict["top_counter"], whisper_dict["left_counter"])
    try:
        Commands.FindElement(whisper_dict["top_counter"])
        Commands.Wait10s_ClickElement("//a[text()='" + objects.hanbiro_content + "']")
        Logging("Notification - Click to view incoming notification")
    except WebDriverException:
        whisper_tc["write"]["fail"].update({"description": "Fail to receive notification"})
        TestCase_LogResult(**whisper_tc["write"]["fail"])

        Commands.ClickElement(whisper_dict["list_item"])
        Logging("Whisper List - Click to view item in list")

    try:
        Waits.Wait10s_ElementLoaded(whisper_dict["content"] + "[contains(.,'" + objects.hanbiro_content + "')]")
        FindPushNoti()
        TestCase_LogResult(**whisper_tc["view"]["pass"])
    except WebDriverException:
        try:
            Logging("View whisper again")
            Commands.ClickElement(whisper_dict["back_button"])
            Logging("Whisper View - Escape view page")

            Waits.Wait10s_ElementLoaded(whisper_dict["list_footer"])
            Logging("Whisper View - Back to Whisper List")

            Commands.ClickElement(whisper_dict["list_item"])
            Logging("Whisper List - Click to view item in list")

            Waits.Wait10s_ElementLoaded(whisper_dict["content"] + "[contains(.,'" + objects.hanbiro_content + "')]")
            TestCase_LogResult(**whisper_tc["view"]["pass"])
        except WebDriverException:
            TestCase_LogResult(**whisper_tc["view"]["fail"])
    
    # < Check counter >
    counter2 = Counter_CheckCounterNumber(whisper_dict["top_counter"], whisper_dict["left_counter"])
    if counter2 != counter1:
        Logging("Read Counter - Read counter is updated successfully")
    else:
        whisper_tc["view"]["fail"].update({"description": "Read counter is not updated"})
        TestCase_LogResult(**whisper_tc["view"]["fail"])

    Commands.ClickElement(whisper_dict["back_button"])
    Logging("Whisper View - Escape view page")
    
    Waits.Wait10s_ElementLoaded(whisper_dict["list_footer"])
    Logging("Whisper View - Back to Whisper List")     

def Whisper_ReplyForward(driver, whisper_msg, link_text, recipient_number, org_key, user):
    Commands.Wait10s_ClickElement(whisper_dict["list_item"] + "[contains(.,'" + whisper_msg + "')]")

    Waits.Wait10s_ElementLoaded(whisper_dict["content"])
    
    content_p = Functions.GetListLength(whisper_dict["content"])
    Logging("whisper content" + str(content_p))
    Commands.ClickElement("//a[text()='" + link_text + "']")

    Waits.WaitElementLoaded(15, "//*[@class='tox-edit-area__iframe']")
    

    whisper_recipients = Functions.GetListLength(whisper_dict["selected_org"] + "[contains(@class, 'tag')]")
    if whisper_recipients == recipient_number:
        Logging("Reply Whisper - Recipient is selected for Reply")
        Logging(objects.testcase_pass)
    if link_text == "Forward":
        Commands.Wait10s_ClickElement("//div[@id='div-sender']/label/a")
        Logging("Write Whisper - Click Org button")

        Waits.Wait10s_ElementLoaded(whisper_dict["org_input"])
        org_input = Commands.FindElement(whisper_dict["org_input"])
        time.sleep(1)
        org_input.click()
        Commands.InputElement_2Values(whisper_dict["org_input"], org_key, Keys.RETURN)

        Commands.Wait10s_ClickElement("//a[text()='" + user + "']")
        Logging("Write Whisper - Select recipients from organization")

        Commands.ClickElement(whisper_dict["add_user"])
        Logging("Write Whisper - Select selected recipients")

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".listbox")))

        listbox = Functions.GetListLength(whisper_dict["listbox"])

        if listbox == 1:
            Logging("Write Whisper - 1 recipient is selected")
            Logging(objects.testcase_pass)
        
        Commands.ClickElement(whisper_dict["save_org"])
        Logging("Write Whisper - Save organization tree")

        Waits.Wait10s_ElementLoaded(whisper_dict["selected_org"])

        whisper_recipients = Functions.GetListLength(whisper_dict["selected_org"] + "[contains(@class, 'tag')]")
        if whisper_recipients == 1:
            Logging("Write Whisper - Whisper recipients are selected")
            Logging(objects.testcase_pass)

    time.sleep(1)

    Commands.SwitchToFrame("//*[@class='tox-edit-area__iframe']")
    
    reply_p = Functions.GetListLength("//*[@id='tinymce']/p")
    Logging("reply content" + str(reply_p))
    if reply_p > content_p:
        Logging("Reply All Whisper - Reply content is inserted successfully")
        Logging(objects.testcase_pass)
    else:
        Logging("Reply All Whisper - Fail to insert whisper content")
        Logging(objects.testcase_fail)
    
    Commands.SwitchToDefaultContent()

    CloseAutosave()

    Commands.ClickElement(whisper_dict["send"])
    Logging("Write - Send whisper")

    whisper_group = ["Reply", "Reply All", "Forward"]

    try:
        Waits.Wait10s_ElementLoaded(whisper_dict["list_footer"])
        Logging("Whisper is saved successfully")

        for whisper in whisper_group:
            if link_text == whisper:
                whisper_tc["write"]["pass"].update({"testcase": whisper + " Whisper"})
                whisper_tc["write"]["pass"].update({"description": whisper + " whisper successfully"})
        
        TestCase_LogResult(**whisper_tc["write"]["pass"])
    except WebDriverException:
        for whisper in whisper_group:
            if link_text == whisper:
                whisper_tc["write"]["fail"].update({"testcase": whisper + " Whisper"})
                whisper_tc["write"]["fail"].update({"description": "Fail to " + whisper + " whisper"})
        
        TCResult_ValidateAlertMsg(menu="whisper", testcase="write", msg="click send whisper")
        TestCase_LogResult(**whisper_tc["write"]["fail"])

def Whisper_Search():
    Waits.Wait10s_ElementLoaded(whisper_dict["list_item"])

    list1 = ValidateListTotal(whisper_dict["list_footer"])
    Waits.Wait10s_ElementLoaded(whisper_dict["search_input"][4])
    
    Commands.InputElement_2Values(whisper_dict["search_input"][4], "ruiemnnvf", Keys.ENTER)
    Logging("Search Input - Input key words")
    
    Waits.WaitUntilPageIsLoaded(None)

    i=0
    for i in range(10):
        i+=1
        time.sleep(1)
        list2 = ValidateListTotal(whisper_dict["list_footer"])
        if list2 != list1:
            search = True
            break
        else:
            search = False
    
    if search == True:
        TestCase_LogResult(**whisper_tc["search"]["pass"])
        Commands.InputElement_2Values(whisper_dict["search_input"][4], "", Keys.ENTER)
        Logging("Reset search result")
        time.sleep(1)
        Waits.WaitUntilPageIsLoaded(None)
    else:
        TestCase_LogResult(**whisper_tc["search"]["fail"])

def WhisperExecution_Driver1(domain_name, user_id):
    PrintYellow("-------------------------------------------")
    PrintYellow("[MENU WHISPER]")

    AccessGroupwareMenu(name = "whisper,Whisper", page_xpath = "//div[@id='ngw.whisper.container']/split-screen-view/list-view/div/div[2]/div/div/whisper-list/div[2]/div")

    time.sleep(1)

    PrintYellow("[SEND NEW WHISPER]")
    whisper_msg = Whisper_SendNewWhisper(domain_name, user_id)
    
    PrintYellow("[VIEW WHISPER]")
    Whisper_ViewDetails(whisper_msg)
    
    PrintYellow("[SEARCH WHISPER]")
    Whisper_Search()

    time.sleep(1)
    
    PrintYellow("[MENU TASK WHISPER] MOVE PAGE")
    List_ValidateListMovingPage(whisper_dict["list_target"], whisper_dict["item_suf"], whisper_dict["page_total"], whisper_dict["nextpage_icon"])
    
    ValidateUnexpectedModal()

    return whisper_msg
    
def WhisperExecution(driver2, domain_name, user_id, org_key, user1, user2, user3):
    PrintYellow("DRIVER1")

    # send whisper
    whisper_msg = WhisperExecution_Driver1(domain_name, user_id, org_key, user1, user2, user3)
    
    PrintYellow("DRIVER2")
    
    driver2.get(domain_name + "/whisper/list/inbox/")
    try:
        Waits.Wait10s_ElementLoaded(whisper_dict["item"])
    except WebDriverException:
        Commands.ReloadBrowser(whisper_dict["item"])
        Waits.Wait10s_ElementLoaded(whisper_dict["item"])
    
    # reply / forward received whisper
    Whisper_ReplyForward(driver2, whisper_msg, "Reply", 1, org_key, user1)
    Whisper_ReplyForward(driver2, whisper_msg, "Reply All", 2, org_key, user2)
    Whisper_ReplyForward(driver2, whisper_msg, "Forward", 0, org_key, user3)