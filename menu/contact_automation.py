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
    global contact_dict, contact_tc
    contact_dict = dict(data["contact"])
    contact_tc = dict(data["testcase_result"]["contact"])

def Contact_AccessMyContact(domain_name):
    PrintYellow("-------------------------------------------")
    PrintYellow("[MENU CONTACT]")

    Commands.NavigateTo(domain_name + "/addrbook/favourite")
    try:
        Waits.Wait10s_ElementLoaded(contact_dict["my_contacts"])
        access_contact = True
    except WebDriverException:
        access_contact = False
    
    if access_contact == True:
        Commands.ClickElement(contact_dict["my_contacts"])
        Logging("Access My Contacts")
        
        Commands.Wait10s_ClickElement(contact_dict["mycontacts_folder"])
        Logging("Access My Contact Folder")
        
        time.sleep(1)

        list_data = CollectListData(list_footer=contact_dict["list_footer"], page_total=contact_dict["page_total"])
        contacts = list_data["total_items"]
        pages = list_data["total_pages"]

        current_url = DefineCurrentURL()
    else:
        current_url = None
        contacts = 0
        pages = 0

    contact_data = {
        "access_menu": access_contact,
        "current_url": current_url,
        "contacts": contacts,
        "pages": pages
    }

    return contact_data
    
def Contact_AddMyContacts():
    PrintYellow("[ADD CONTACTS]")
    
    list_total1 = ValidateListTotal(contact_dict["list_footer"])
    # validate list counter before adding new contact
    
    FindPushNoti()

    Commands.ClickElement(contact_dict["write_button"])
    Logging("My Contacts - Click write contact button")

    CloseAutosave()
    # find an close autosave
    
    contact_number = str(Functions.getRandomNumber_fromSpecificRange(1, 100))
    contact_name = "Contact " + objects.date_id
    phone_number = str(Functions.getRandomNumber_fromSpecificRange(100000000, 999999999))
    email_addr = "email" + contact_number + "@test.com"
    company_name = "Company" + contact_number
    position_name = "Position" + contact_number
    contact_info = {
        "name": contact_name,
        "phone": phone_number,
        "email": email_addr,
        "position": company_name,
        "company": position_name
    }

    Waits.Wait10s_ElementLoaded(data["common"]["loading_dialog"])

    Commands.InputElement(contact_dict["name"], contact_name)
    Logging("My Contacts - Add - Input contact name:", contact_name)

    Commands.InputElement(contact_dict["company"], company_name)
    Logging("My Contacts - Add - Input company")

    Commands.InputElement(contact_dict["position"],position_name)
    Logging("My Contacts - Add - Input Position")

    Commands.InputElement(contact_dict["email0"], email_addr)
    Logging("My Contacts - Add - Input contact email")
    
    Commands.InputElement(contact_dict["phone0"], phone_number)
    Logging("My Contacts - Add - Input contact number")

    phone_label = Commands.FindElement(contact_dict["phone_label"])
    contact_info["phone_label"] = phone_label.text
    
    CloseAutosave()

    FindPushNoti()
    Commands.ClickElement(contact_dict["save_button"])
    Logging("My Contacts - Add - Click Save contact")

    try:
        Waits.Wait10s_ElementLoaded(contact_dict["contact_name"])
        TestCase_LogResult(**contact_tc["write"]["pass"])
        #Access contact details after save new contact

        time.sleep(2)

        Commands.ClickElement(contact_dict["back_button"])
        Logging("My Contacts - Content - Back to contact list")

        Waits.Wait10s_ElementLoaded(contact_dict["list_footer"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="contact", testcase="write", msg="click save contact")
        TestCase_LogResult(**contact_tc["write"]["fail"])

    time.sleep(1)

    return contact_info 

def Contact_Import():
    PrintYellow("[IMPORT CONTACTS]")
    list_total1 = ValidateListTotal(contact_dict["list_footer"])

    Commands.Wait10s_ClickElement(contact_dict["more_button"])
    Commands.Wait10s_ClickElement(contact_dict["import_button"])

    # store the object of Workbook class in a variable
    wrkbk = openpyxl.Workbook()
    # to create a new sheet
    sh = wrkbk.create_sheet("Details", 1)

    contact_no = str(Functions.getRandomNumber_fromSpecificRange(1, 999))
    random_number = str(Functions.getRandomNumber_fromSpecificRange(100000000, 999999999))
    contact_name = "Contact " + objects.date_time.replace("/", "").replace(":", "").replace(", ", "")
    Logging("contact name for import:",contact_name)
    
    import_data = dict(contact_dict["import_data"])
    header = list(import_data.keys())
    column_no = 0
    
    while column_no < int(len(import_data.keys())):
        column_no += 1
        header_value = str(header[column_no-1])
        contact_value = str(import_data[header_value])
        
        if "contact_name" in contact_value:
            contact_value_import = contact_value.replace("contact_name", contact_name)
        elif "[contact_no]" in contact_value:
            contact_value_import = contact_value.replace("[contact_no]", contact_no)
        elif "random_number" in contact_value:
            contact_value_import = contact_value.replace("random_number", random_number)
        else:
            contact_value_import = contact_value

        sh.cell(row=1, column=column_no).value = header_value
        sh.cell(row=2, column=column_no).value = contact_value_import

    wrkbk.get_sheet_names()
    removed_sheet = wrkbk.get_sheet_by_name('Sheet')
    wrkbk.remove_sheet(removed_sheet)
    wrkbk.get_sheet_names()

    # to save the workbook
    import_file = Files.contact_import
    wrkbk.save(import_file)

    Commands.Wait10s_InputElement(contact_dict["file_uploader"], import_file)
    Waits.Wait10s_ElementLoaded(contact_dict["htcontainer"])

    time.sleep(1)

    Commands.ClickElement(contact_dict["save_button"])
    Commands.Wait10s_ClickElement(data["resource"]["confirm_button"])

    try:
        Waits.Wait10s_ElementLoaded(contact_dict["importing"])
        Waits.Wait10s_ElementInvisibility(contact_dict["importing"])
    except WebDriverException:
        pass

    time.sleep(2)

    list_total2 = ValidateListTotal(contact_dict["list_footer"])

    list_result = ValidateListNumberUpdate(list_total1, list_total2)
    if list_result == True:
        TestCase_LogResult(**contact_tc["import"]["pass"])
        new_contact = True
    else:
        contact_tc["import"]["fail"].update({"description": "List counter is not updated after contact is imported"})
        TestCase_LogResult(**contact_tc["import"]["fail"])
        new_contact = None

    if new_contact == None:
        TCResult_ValidateAlertMsg(menu="contact", testcase="import", msg="click save import")
        TestCase_LogResult(**contact_tc["import"]["fail"])
    else: 
        Waits.Wait10s_ElementLoaded(contact_dict["contact_item"])
        
        Waits.Wait10s_ElementLoaded(contact_dict["search_input"][4])
        Commands.InputElement_2Values(contact_dict["search_input"][4], contact_name, Keys.ENTER)
        Logging("Search Input - Enter key word of new imported contact")

        Waits.Wait10s_ElementLoaded(contact_dict["contact_title_name"] % (contact_name))
        import_contact = Commands.FindElement(contact_dict["contact_title_name"] % (contact_name))
        time.sleep(1)
        import_contact.click()

        Waits.Wait10s_ElementLoaded(contact_dict["addr_item_list"])
        import_details = Functions.GetListLength(contact_dict["addr_item_list"])
        if import_details == 11:
            contact_tc["view"]["pass"].update({"testcase": "View imported contact"})
            contact_tc["view"]["pass"].update({"description": "All details are saved successfully"})
            TestCase_LogResult(**contact_tc["view"]["pass"])
        else:
            Logging("number of imported fields " + str(import_details))
            contact_tc["view"]["fail"].update({"testcase": "View imported contact"})
            contact_tc["view"]["fail"].update({"description": "All details are not saved successfully"})
            TestCase_LogResult(**contact_tc["view"]["fail"])
        
        time.sleep(1)
        
        import_elements = {
            "Contact Import": contact_dict["import_group"],
            "Dept.": contact_dict["import_dept"],
            "Position": contact_dict["import_position"],
            "Memo": contact_dict["import_memo"] 
        }

        for import_element in import_elements:
            element = Commands.FindElement(import_elements["" + import_element + ""])
            if str(element.text) in import_elements.keys():
                contact_tc["import"]["pass"].update({"testcase": "View Import contact [%s]" % str(element.text)})
                contact_tc["import"]["pass"].update({"description": "View import contact field [%s] successfully" % str(element.text)})
                TestCase_LogResult(**contact_tc["import"]["pass"])
            else:
                contact_tc["import"]["fail"].update({"testcase": "View Import contact [%s]" % str(element.text)})
                contact_tc["import"]["fail"].update({"description": "Fail to view import contact field [%s]" % str(element.text)})
                TestCase_LogResult(**contact_tc["import"]["fail"])
        
        Commands.ReloadBrowser(contact_dict["back_button"])
        Commands.Wait10s_ClickElement(contact_dict["back_button"])
        Logging("My Contacts - Content - Back to contact list")
        
        Commands.Wait10s_ClickElement(contact_dict["search_expander"])
        Commands.Wait10s_InputElement(contact_dict["search_content"], Keys.ENTER)
        Logging("Search Input - Enter empty key word")

        Waits.WaitUntilPageIsLoaded(contact_dict["contact_item"])

    time.sleep(1)

def Contact_VaidateListExposure(contact_name):
    PrintYellow("[MENU CONTACT] CHECK LIST EXPOSURE")

    Waits.WaitUntilPageIsLoaded(None)
    
    Waits.Wait10s_ElementLoaded(contact_dict["search_input"][4])
    Commands.InputElement_2Values(contact_dict["search_input"][4], contact_name, Keys.ENTER)
    Logging("Search contacts - Enter key words")

    time.sleep(2)
    Waits.WaitUntilPageIsLoaded(None)
    
    list_data = ["homepage", "memo", "position", "department", "company", "group", "mobile", "phone", "email"]
    Commands.Wait10s_ClickElement(contact_dict["list_dropdown"])
    Logging("Contact List - Open list dropdown")

    exposure_value = {}
    for exposure_name in list_data:
        dropdown_exposure_text = Commands.FindElement(contact_dict["dropdown_exposure"].replace("column_name", exposure_name))
        if "ng-not-empty" in dropdown_exposure_text.get_attribute("class"):
            Logging(exposure_name + " is activated")
            exposure_value[exposure_name] = True
        else:
            Logging(exposure_name + " is not activated")
            exposure_value[exposure_name] = False

    Commands.ClickElement(contact_dict["list_dropdown"])
    Logging("Contact List - Close list dropdown")
    time.sleep(1)
    
    exposure_header_text = []
    exposure_header_url = []
    for exposure_name in exposure_value:
        exposure_column_text = str(contact_dict["exposure_column"] % (exposure_name))
        exposure_column_url = exposure_column_text.replace("span[", "div[")
        if exposure_value[exposure_name] == True:
            try:
                try:
                    Commands.FindElement(exposure_column_text)
                    # variable in list_data_text
                    exposure_header_text.append(exposure_name)
                except WebDriverException:
                    Commands.FindElement(exposure_column_url)
                    # variable in list_data_url
                    exposure_header_url.append(exposure_name)
                Logging(exposure_name + " column is displayed in list while activated")
                Logging(objects.testcase_pass)
            except WebDriverException:
                Logging(exposure_name + " column is not displayed while activated")
                Logging(objects.testcase_fail)
        else:
            try:
                try:
                    Commands.FindElement(exposure_column_text)
                    # variable in list_data_text
                except WebDriverException:
                    Commands.FindElement(exposure_column_url)
                    # variable in list_data_url
                Logging(objects.testcase_fail)
            except WebDriverException:
                Logging(exposure_name + " column is not displayed while activated")
                Logging(objects.testcase_pass)
    
    exposure_header = exposure_header_text + exposure_header_url
    for exposure in exposure_header:
        try:
            ele_exposure = Functions.GetElementText(contact_dict["exposure"].replace("div[%s]", "div") % exposure)
        except WebDriverException:
            ele_exposure = Functions.GetElementText(contact_dict["exposure_phone"].replace("div[%s]", "div") % exposure)
        
        if bool(ele_exposure.strip()) == True:
            Logging(exposure + " data is displayed in activated column of list exposure")
            Logging(objects.testcase_pass)
        else:
            Logging(exposure + " data is not displayed in activated column of list exposure")
            Logging(objects.testcase_fail)

    Commands.InputElement(contact_dict["search_input"][4], Keys.ENTER)
    Waits.WaitUntilPageIsLoaded(None)

def Contact_CopyCompanyContact():
    PrintYellow("[MENU CONTACT] COPY CONTACT")
    list1 = ValidateListTotal(contact_dict["list_footer"])
    
    Commands.ClickElement(contact_dict["mycompany_contacts"])

    Commands.Wait10s_ClickElement(contact_dict["company_checkbox"])
    Logging("Select company contacts")

    time.sleep(1)

    Commands.ClickElement(contact_dict["more_button"])
    Logging("Click More button")

    Commands.Wait10s_ClickElement(contact_dict["copy_button"])

    Waits.Wait10s_ElementLoaded(contact_dict["copy_org_li"])
    time.sleep(1)

    Commands.ClickElement(contact_dict["mycontact_link"])
    Logging("Select My Contacts folder")

    Commands.ClickElement(contact_dict["save_copy"])
    Logging("Click Save button")

    Waits.WaitUntilPageIsLoaded(None)
    Commands.ReloadBrowser(contact_dict["mycontacts_downnarrow"])
    Waits.WaitUntilPageIsLoaded(contact_dict["mycontacts_downnarrow"])

    try:
        Commands.ClickElement(contact_dict["mycontacts_downnarrow"])
    except WebDriverException:
        Commands.ScrollIntoView(contact_dict["mycontacts_downnarrow"])
        Commands.ClickElement(contact_dict["mycontacts_downnarrow"])
    finally:
        Logging("Access My Contacts")

    time.sleep(1)
    
    Waits.Wait10s_ElementLoaded(contact_dict["my_contacts_span"])
    try:
        Commands.ClickElement(contact_dict["my_contacts_span"])
    except WebDriverException:
        Commands.ScrollIntoView(contact_dict["my_contacts_span"])
        Commands.ClickElement(contact_dict["my_contacts_span"])
    finally:
        Logging("Access My Contact Folder")

    Waits.Wait10s_ElementLoaded(contact_dict["list_footer"])

    time.sleep(2)

    list2 = ValidateListTotal(contact_dict["list_footer"])
    if list2 > list1:
        TestCase_LogResult(**contact_tc["copy"]["pass"])
    else:
        TestCase_LogResult(**contact_tc["copy"]["fail"])

def Contact_SearchContact():
    PrintYellow("[TEST CASE] SEARCH CONTACT")
    wrapper(searchInput, contact_dict["search_input"])

    time.sleep(2)

def Contact_ValidateNextPageList():
    PrintYellow("[MENU CONTACT] MOVE PAGE")
    List_ValidateListMovingPage(contact_dict["list_target"], contact_dict["item_suf"], contact_dict["page_total"], contact_dict["nextpage_icon"])

def ContactExecution(domain_name):
    list_data = Contact_AccessMyContact(domain_name)
    contact_info = Contact_AddMyContacts()
    Contact_SearchContact()
    Contact_CopyCompanyContact()
    Contact_VaidateListExposure(contact_info["name"])
    Contact_ValidateNextPageList()
    Contact_Import()
    ValidateUnexpectedModal()