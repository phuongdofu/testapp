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
    global expense_dict, expense_tc
    expense_dict = dict(data["expense"])
    expense_tc = dict(data["testcase_result"]["expense"])

def Expense_AccessMenu(domain_name):
    Commands.NavigateTo(domain_name + "/expense/list/my/all/")
    
    try:
        Waits.Wait10s_ElementLoaded(expense_dict["list_footer"])
        Logging("Access expense menu successfully")
        access_menu = True
    except WebDriverException:
        Logging("Fail to access expense menu")
        access_menu = False

    list_footer = Functions.GetElementText(expense_dict["list_footer"])
    expenses = int(list_footer.replace(",", "").split(" ")[1])

    try:
        pages = int(Functions.GetElementText(expense_dict["page_total"]))
    except WebDriverException:
        pages = 0

    expense_menu = {
        "access_menu": access_menu,
        "expenses": expenses,
        "pages": pages
    }
    
    return expense_menu

def Expense_ValidateRequiredData():
    currency = payment = None

    try:
        Waits.WaitElementLoaded(5, expense_dict["admin_submenu"])
        expense_admin = True
    except WebDriverException:
        expense_admin = False

    if expense_admin == True:
        Waits.WaitElementLoaded(5, expense_dict["admin_submenu"])
        Commands.ClickElement(expense_dict["admin_submenu"])
        Logging("Access Admin sub-menu")

        Logging("[Add Currency]")
        Commands.Wait10s_ClickElement(expense_dict["currency_submenu"])
        Logging("Access Curreny sub-menu")

        Waits.WaitUntilPageIsLoaded(expense_dict["admin_list_footer"])

        try:
            Commands.FindElement(expense_dict["list_empty"])
            Logging("Curreny list is empty")
        except WebDriverException:
            currency = True
            Logging("Currency is valid")
        
        if bool(currency) == False:
            try:
                Commands.ClickElement(expense_dict["add_payment"])
                Logging("Click Add Currency")

                Waits.Wait10s_ElementLoaded(expense_dict["select_currency"])
                Commands.Selectbox_ByVisibleText(expense_dict["select_currency"], "United States Dollar")
                Logging("Select VND Label")

                Commands.ClickElement(expense_dict["save_payment"])
                Logging("Click Save currency")

                Waits.WaitElementLoaded(3, expense_dict["confirm_success"])
                Commands.ClickElement(expense_dict["confirm_success"])
                Logging("Confirm Success modal")

                Waits.Wait10s_ElementLoaded(expense_dict["actions_button"])
                Logging("New currency is added successfully")
                currency = True
            except WebDriverException:
                pass

        time.sleep(1)

        Logging("[Add Payment Method]")
        Commands.ClickElement(expense_dict["payment_submenu"])
        Logging("Access Payment Method")

        Waits.WaitUntilPageIsLoaded(expense_dict["admin_list_footer"])

        try:
            Commands.FindElement(expense_dict["list_empty"])
            Logging("Payment method list is empty")
        except WebDriverException:
            Logging("Payment method is valid")
            payment = True
        
        if bool(payment) == False:
            try:
                Commands.ClickElement(expense_dict["add_payment"])
                Logging("Click Add Payment method")

                payment_name = "New Cash"
                Commands.Wait10s_InputElement(expense_dict["payment_name_input"], payment_name)
                Logging("Input Payment method name")

                Commands.ClickElement(expense_dict["cash_payment"])
                Logging("Define Payment method as Cash")

                try:
                    Commands.FindElement(expense_dict["active_payment"])
                    Logging("Payment method is being active")
                except WebDriverException:
                    Logging("Payement method is not active")
                    Commands.ClickElement(expense_dict["activate_payment"])
                    Logging("Activate payment method")

                    Waits.WaitElementLoaded(3, expense_dict["active_payment"])
                
                Commands.ClickElement(expense_dict["save_payment"])
                Logging("Click Save payment method")

                Waits.WaitElementLoaded(3, expense_dict["confirm_success"])
                Commands.ClickElement(expense_dict["confirm_success"])
                Logging("Confirm Success modal")

                Waits.Wait10s_ElementLoaded(expense_dict["actions_button"])
                Logging("New payment method is added successfully")
                
                payment = True
            except WebDriverException:
                pass
        
        time.sleep(1)

        Commands.ClickElement(expense_dict["myexpense_submenu"])
        Logging("Access My Expenses submenu")
        Waits.WaitUntilPageIsLoaded(expense_dict["list_footer"])
        Commands.ReloadBrowser(expense_dict["list_footer"])
        Logging("Reload page")
        Waits.WaitUntilPageIsLoaded(expense_dict["list_footer"])

    expense_data = (currency, payment)
    
    return expense_data

def Expense_AddNewItem():
    PrintYellow("[MENU EXPENSE - ADD NEW EXPENSE ITEM]")

    Commands.ClickElement(expense_dict["item_plus"])
    Logging("Add Item - Click add item button")

    Waits.WaitUntilPageIsLoaded(expense_dict["item_subject"])
    Commands.InputElement(expense_dict["item_subject"], "Item 1")
    Logging("Add Item - Input title")
    item_subject = Functions.GetInputValue(expense_dict["item_subject"])
    
    Commands.InputElement(expense_dict["item_amount"], "1000")
    Logging("Add Item - Input item amount")

    try:
        Commands.FindElement(expense_dict["invalid_card"])
        Logging("Credit card is empty (Cannot save item)")
        
        payment_len = Functions.GetListLength(expense_dict["payment_length"])

        for i in range(1, payment_len):
            i+=1
            Commands.Selectbox_ByIndex(expense_dict["payment_method"], str(i))
            time.sleep(1)
            try:
                Waits.WaitUntilPageIsLoaded(None)
                Commands.FindElement(expense_dict["credit_card_select"])
            except WebDriverException:
                Logging("Payment method with Cash is selected")
                break
    except WebDriverException:
        pass

    Waits.Wait10s_ElementLoaded(expense_dict["editor_header"])
    Commands.SwitchToFrame(expense_dict["editor_item"])
    Commands.InputElement(data["editor"]["input_p"], objects.hanbiro_content)
    Logging("Add Item - Input content")
    Commands.SwitchToDefaultContent()

    Commands.ClickElement(expense_dict["save_item_button"])
    Logging("Add Item - Click Save button")

    try:
        Waits.WaitElementLoaded(5, expense_dict["item_list_footer"])
        TestCase_LogResult(**expense_tc["add_item"]["pass"])
    except WebDriverException:
        TestCase_LogResult(**expense_tc["add_item"]["fail"])

    return item_subject

def Expense_ImportItem():
    Commands.ClickElement(expense_dict["import_button"])

    # store the object of Workbook class in a variable
    wrkbk = openpyxl.Workbook()
    # to create a new sheet
    sh = wrkbk.create_sheet("Details", 2)

    date_import = "{}/{}/{}".format(objects.year, objects.month, objects.day)
    header_group = ["Item Name", "Date", "Purpose", "Purpose2", "Purpose3", "Payment Method","Credit Card", "Currency", "Amount", "Memo"]
    item_group = ["Import Item", date_import, "Purpose 1", "Purpose 2", "Purpose 3", "Transfer", "SCB Bank", "Vietnamese Dong", "10000", "Automation test import"]
    column_no = 0
    for item in header_group and item_group:
        column_no += 1
        sh.cell(row=1, column=column_no).value = item
        sh.cell(row=2, column=column_no).value = item

    wrkbk.get_sheet_names()
    removed_sheet = wrkbk.get_sheet_by_name('Sheet')
    wrkbk.remove_sheet(removed_sheet)
    wrkbk.get_sheet_names()

    # to save the workbook
    import_file = Files.expense_import
    wrkbk.save(import_file)
    Waits.Wait10s_ElementLoaded(data["attachment"]["file_input"])
    Commands.Wait10s_InputElement(data["attachment"]["file_input"], import_file)

    Waits.Wait10s_ElementLoaded(expense_dict["ht_container"])
    Logging("Item Import - Wait until file is loaded successfully")

    time.sleep(1)

    Commands.ClickElement(expense_dict["ok_button"])

    try:
        Waits.WaitElementLoaded(5, expense_dict["import_warning"])
        Commands.ClickElement(expense_dict["import_confirm"])
    except WebDriverException:
        pass
    
    try:
        Waits.Wait10s_ElementLoaded(expense_dict["import_table"])
        Commands.FindElement(expense_dict["imported_item"])
        TestCase_LogResult(**expense_tc["import_item"]["pass"])
        return 'Import Item'
    except WebDriverException:
        TestCase_LogResult(**expense_tc["import_item"]["fail"])
        return None

def Expense_WriteExpense(domain_name):
    PrintYellow("[EXPENSE] WRITE EXPENSE")

    expense_title = new_item = import_item = None

    # Validate if currency and payment method are valid
    expense_data = Expense_ValidateRequiredData()
    currency = expense_data[0]
    payment = expense_data[1]

    if bool(currency) == True:
        CommonWriteItem(expense_dict["pen_button"], expense_dict["subject"], objects.hanbiro_content)

        expense_title = Functions.GetInputValue(expense_dict["subject"])
        currency_ele = Functions.GetElementText(expense_dict["currency_select"])
        start_date = Functions.GetInputValue(expense_dict["duration_start"])
        end_date = Functions.GetInputValue(expense_dict["duration_end"])

        required_fields = [currency_ele, start_date, end_date]
        for required_field in required_fields:
            if bool(required_field) == False:
                Logging("Cannot create expense because required field is empty")
                Logging(objects.testcase_fail)
                Commands.NavigateTo(domain_name + "/expense/list/my/all/")
                Waits.Wait10s_ElementLoaded(expense_dict["list_footer"])
                create_expense = False
            else:
                create_expense = True
                Logging("Write Expense - The required field is not empty")
        
        if create_expense == True and bool(payment):
            try:
                Commands.FindElement(expense_dict["unlimited_budget"])
                Logging("Expense budget is unlimited (default)")
            except WebDriverException:
                Commands.InputElement(expense_dict["budget_limit"], "100000")
                Logging("Expense budget is limited - Input amount")
            
            try:
                new_item = Expense_AddNewItem()
                time.sleep(1)
            except:
                new_item = None

            try:
                import_item = Expense_ImportItem()
                time.sleep(1)
            except:
                import_item = None  

        FindPushNoti()
        Commands.ClickElement(expense_dict["save_expense_button"])  
        Logging("Write Expense - Click Save button")

        Waits.Wait10s_ElementLoaded(expense_dict["list_footer"])
        try:
            Waits.Wait10s_ElementLoaded(expense_dict["expense_item"] % objects.hanbiro_title)
            TestCase_LogResult(**expense_tc["write_expense"]["pass"])
        except WebDriverException:
            TestCase_LogResult(**expense_tc["write_expense"]["fail"])
    
    expense_data = {
        "expense_title": expense_title,
        "new_item": new_item,
        "import_item": import_item
    }

    return expense_data

def Expense_ViewExpense(expense_title, new_item, import_item):
    PrintYellow("[EXPENSE] VIEW EXPENSE")
    Commands.Wait10s_ClickElement(expense_dict["expense_item"] % expense_title)
    print("Select expense to view")
    try:
        Waits.Wait10s_ElementLoaded(expense_dict["item_header"])
        expense_tc["view_expense"]["pass"].update({"description": "Access item view modal successfully"})
        TestCase_LogResult(**expense_tc["view_expense"]["pass"])

        created_items = [new_item, import_item]
        created_item_view = []
        for created_item in created_items:
            if bool(created_item) == True:
                try:
                    Commands.FindElement(expense_dict["td_item"] % created_item)
                    Logging("Created item is displayed in expense view page")
                    created_item_view.append(True)
                except WebDriverException:
                    created_item_view.append(False)

        if False in created_item_view:
            expense_tc["view_item"]["fail"].update({"description": "Created item / Imported item are not displayed in expense view page"})

        if created_item_view[0] and created_item_view[1] == True:
            existing_items = True
        else:
            existing_items = False

        empty_data = []
        item_data = {
            "Item name": expense_dict["list_item_name"],
            "Item amount": expense_dict["list_item_amount"]
        }
        for item in item_data.keys():
            Waits.WaitElementLoaded(5, item_data[item])
            item_ele = Functions.GetElementText(item_data[item])
            if bool(item_ele) == False:
                expense_tc["view_item"]["fail"].update({"description": str(item) + " is empty"})
                TestCase_LogResult(**expense_tc["view_item"]["fail"])
                empty_data.append(False)
            else:
                empty_data.append(True)
        
        if False in empty_data:
            expense_tc["view_item"]["fail"].update({"description": "Item name or amount is empty"})
            TestCase_LogResult(**expense_tc["view_item"]["fail"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="expense", testcase="view_expense", msg="click view")
        TestCase_LogResult(**expense_tc["view_expense"]["fail"])
        existing_items = None

    return existing_items

def Expense_ViewItem(created_item):
    PrintYellow("[EXPENSE] VIEW ITEM")
    
    try:
        Commands.FindElement(expense_dict["td_item"] % created_item)
        item = True
    except WebDriverException:
        expense_tc["view_item"]["fail"].update({"description": "Cannot find item to view " + created_item})
        TestCase_LogResult(**expense_tc["view_item"]["fail"])
        item = False
    
    if item == True:
        item_amount_number = Functions.GetElementText(expense_dict["item_amount_expense"])
        Commands.ClickElement(expense_dict["item_details"])
        Logging("View Item - Click Details button in expense content")
        try:
            Waits.Wait10s_ElementLoaded(expense_dict["item_wrapper"])
            TestCase_LogResult(**expense_tc["view_item"]["pass"])

            item_title2 = Functions.GetElementText(expense_dict["item_title_details"])
            Logging("item_amount_number: " + item_amount_number)
            if item_title2 != created_item:
                expense_tc["view_item"]["fail"].update({"description": "Fail to view Item name"})
                TestCase_LogResult(**expense_tc["view_item"]["fail"])

            item_amount_details = Functions.GetElementText(expense_dict["item_amount_details"])
            Logging(item_amount_details)
            if item_amount_details != item_amount_number:
                expense_tc["view_item"]["fail"].update({"description": "Fail to view Item amoun"})
                TestCase_LogResult(**expense_tc["view_item"]["fail"])
        except WebDriverException:
            TCResult_ValidateAlertMsg(menu="expense", testcase="view_item", msg="click view item")
            TestCase_LogResult(**expense_tc["view_item"]["fail"])

        time.sleep(1)

        Commands.ClickElement(expense_dict["close_view_item"])
        Logging("View Item - Close view modal")

        Waits.WaitUntilPageIsLoaded(expense_dict["item_header"])  

def Expense_SubmitApproval(recipient_id):
    PrintYellow("[MENU EXPENSE - SUBMIT EXPENSE AS APPROVAL]")

    if recipient_id != data["tooltip"]["recipient"]:
        Waits.WaitUntilPageIsLoaded(expense_dict["back_button"])
        time.sleep(1)
        Commands.ClickElement(expense_dict["back_button"])
        Waits.Wait10s_ElementLoaded(expense_dict["list_footer"])
    else:
        Commands.Wait10s_ClickElement(expense_dict["content_more"])
        Logging("View Expense - Click More button")

        Commands.Wait10s_ClickElement(expense_dict["content_approval"])
        Logging("View Expense - Click Approval button")

        Waits.Wait10s_ElementLoaded(data["select_form"]["form"])
        form_length = Functions.GetListLength(data["select_form"]["form"])

        i=0
        for i in range(1, form_length):
            i+=1
            
            form_xpath = data["select_form"]["form_order"] % str(i)
            form_type = Functions.GetElementAttribute(data["select_form"]["form_name"] % str(i), "class")
            
            if "html" in form_type:
                Commands.ClickElement(form_xpath)
                Logging("View Expense - Select html form")
                break
            elif "mimefile" in form_type:
                Commands.ClickElement(form_xpath)
                Logging("View Expense - Select mimefile form")
                break

        Commands.ClickElement(data["select_form"]["select_button"])
        Logging("View Expense - Select approval form")

        try:
            Waits.Wait10s_ElementLoaded(data["approval"]["approval_route"])
            submit_approval = True
        except WebDriverException:
            expense_tc["submit_approval"]["fail"].update({"description": "Fail to load approval route"})
            TestCase_LogResult(**expense_tc["submit_approval"]["fail"])
            submit_approval = False
        
        if submit_approval == True:
            Commands.ClickElement(data["approval"]["approver_org"])
            Logging("Write new approval - Click org tree button from approval route")

            Waits.Wait10s_ElementLoaded(data["approval"]["org_drafter"])
            Commands.ClickElement(data["approval"]["org_input"])
            Commands.InputElement_2Values(data["approval"]["org_input"], recipient_id, Keys.RETURN)
            Logging("Input recipient key in search box")
            
            Waits.WaitUntilPageIsLoaded(None)
            Waits.Wait10s_ElementLoaded(expense_dict["user_org_xpath"])

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
            
            try:
                Commands.SwitchToFrame(expense_dict["approval_frame"])
                Logging("Expense Approval - iFrame found")

                Waits.Wait10s_ElementLoaded(expense_dict["approval_expense_title"])
                Logging("Expense Approval - Expense title is displayed in approval content")

                Commands.SwitchToDefaultContent()
                
                save_approval = Waits.Wait10s_ElementLoaded(expense_dict["approval_save"])
                time.sleep(1)
                save_approval.click()
                Logging("Expense Approval - Click to submit exepense approval")

                Commands.Wait10s_ClickElement(expense_dict["apply_button"])
                Logging("Write new approval - Apply approval options")

                Waits.WaitUntilPageIsLoaded(None)
                TestCase_LogResult(**expense_tc["submit_approval"]["pass"])
            except WebDriverException:
                expense_tc["submit_approval"]["fail"].update({"description": "Fail to load approval content"})
                TestCase_LogResult(**expense_tc["submit_approval"]["fail"]) 
                ValidateUnexpectedModal()
            finally:
                Waits.WaitUntilPageIsLoaded(expense_dict["back_button"])
                time.sleep(1)
                Commands.ClickElement(expense_dict["back_button"])
                Waits.Wait10s_ElementLoaded(expense_dict["list_footer"])

def Expense_SearchExpense():
    PrintYellow("[EXPENSE] SEARCH EXPENSE")

    Waits.Wait10s_ElementLoaded(expense_dict["list_item"])
    '''expense_name = Functions.GetElementText(expense_dict["list_item"])
    search_dict = {
        "title":{
            "key": expense_name,
            "value": "number:0"
        }
    }
    search_expense_dict = dict(expense_dict["search_details"])
    search_expense_dict["search_dict"] = search_dict
    SearchDetailsBySelectBox(**search_expense_dict)'''
    searchInput(*expense_dict["search_input"])

def Expense_ValidateNextPageList():
    PrintYellow("[MENU EXPENSE] MOVE PAGE")
    List_ValidateListMovingPage(expense_dict["list_target"], expense_dict["item_suf"], expense_dict["page_total"], expense_dict["nextpage_icon"])

def ExpenseExecution(domain_name, recipient_id):
    PrintYellow("-------------------------------------------")
    PrintYellow("[MENU EXPENSE]")
    expense_menu = Expense_AccessMenu(domain_name)
    expense_data = Expense_WriteExpense(domain_name)
    existing_items = Expense_ViewExpense(expense_data["expense_title"], expense_data["new_item"], expense_data["import_item"])
    Expense_ViewItem(expense_data["new_item"])
    Expense_ViewItem(expense_data["import_item"])
    Expense_SubmitApproval(recipient_id)
    Expense_SearchExpense()
    Expense_ValidateNextPageList()
    ValidateUnexpectedModal()
