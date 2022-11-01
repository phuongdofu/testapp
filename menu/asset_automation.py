import time, sys, unittest, random, json, requests, openpyxl, testlink
from unicodedata import category
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
    global asset_dict, asset_tc
    asset_dict = dict(data["asset"])
    asset_tc = dict(data["testcase_result"]["asset"])

def Asset_ValidateFolderPermission():
    asset_folders = Functions.GetListLength(asset_dict["asset_folder_li"])
    li_number = 0
    folder_permission = []
    for li_number in range(1, asset_folders+1):
        asset_folder_span = asset_dict["asset_folder_span"] % str(li_number)
        asset_folder = Commands.FindElement(asset_folder_span)
        folder_name = str(asset_folder.text)
        asset_folder.click()
        Logging("Access asset folder")

        Waits.Wait10s_ElementLoaded(asset_dict["list_footer"])
        
        try:
            Waits.Wait10s_ElementLoaded(asset_dict["list_footer"])
        except WebDriverException:
            pass
        
        time.sleep(1)

        try:
            Commands.FindElement(asset_dict["pen_button"])
            Logging("User has permission to create asset in folder:", folder_name)
            folder_permission.append(True)
            break
        except WebDriverException:
            folder_permission.append(False)

    if False in folder_permission:
        Logging("User has no permission to create")
        return False
    else:
        return True

def Asset_AccessMenu():
    PrintYellow("-------------------------------------------")
    PrintYellow("[MENU ASSET]")
    
    global access_asset, asset_categories

    access_asset = AccessGroupwareMenu(name="asset,Asset", page_xpath=asset_dict["asset_report"])

    asset_admin = None
    add_category = None
    add_location = None
    asset_categories = 0
    asset_location = 0
    category_name = "Asset folder"
    location_name = "Location"

    if access_asset == True:
        try:
            Commands.FindElement(asset_dict["admin_submenu"])
            Logging("Current user is asset admin - Can add category and location")
            asset_admin = True
        except WebDriverException:
            asset_admin = False
            Logging("Current user is not asset admin - Cannot add category and location")
        
        existing_categories = Functions.GetListLength(asset_dict["report_cate_select"]) -1
        if existing_categories == 0:
            if bool(asset_admin) == True:
                Logging("User has permission to add category and location")
                add_category = True 
            else:
                Logging("User has no permission to add category and location")
        else:
            asset_categories = 1

            # Check if location is valid along with category
            Commands.ClickElement(asset_dict["admin_submenu_txt"])
            Logging("Open Admin sub-menu")
            Waits.WaitElementLoaded(3, asset_dict["admin_isActive"])

            Commands.ClickElement(asset_dict["location_submenu"])
            Logging("Access Location sub-menu")
            Waits.Wait10s_ElementLoaded(asset_dict["location_form"])

            try:
                Waits.WaitElementLoaded(3, asset_dict["location_li"])
                asset_location = 1
            except WebDriverException:
                pass    
    
    if add_category == True:
        try:
            Commands.FindElement(asset_dict["admin_isActive"])
        except WebDriverException:
            Commands.ClickElement(asset_dict["admin_submenu_txt"])
            Logging("Open Admin sub-menu")
            Waits.WaitElementLoaded(3, asset_dict["admin_isActive"])
        
        Commands.Wait10s_ClickElement(asset_dict["admin_category"])
        Logging("Access Categories sub-menu")
        Waits.Wait10s_ElementLoaded(asset_dict["category_form"])

        Commands.InputElement(asset_dict["category_name"], category_name)
        Logging("Input category name")

        time.sleep(1)

        Commands.ClickElement(asset_dict["save_category"])
        Logging("Click save category")
        
        Commands.Wait10s_ClickElement(asset_dict["confirm_success"])
        Logging("Confirm save category")

        try:
            Waits.Wait10s_ElementLoaded(asset_dict["new_category"] % category_name)
            Logging("New category found")
            asset_categories = 1
        except WebDriverException:
            Logging("Cannot found new category")
        
        time.sleep(1)
        
    if asset_categories > 0 and asset_location == 0:
        Commands.ClickElement(asset_dict["location_submenu"])
        Logging("Access Location sub-menu")
        Waits.Wait10s_ElementLoaded(asset_dict["location_form"])

        Commands.InputElement(asset_dict["location_input"], location_name)
        Logging("Input Location name")
        
        time.sleep(1)
        
        Commands.ClickElement(asset_dict["save_location"])
        Logging("Click save location")

        Commands.Wait10s_ClickElement(asset_dict["confirm_success"])
        Logging("Confirm save category")

        try:
            Waits.Wait10s_ElementLoaded(asset_dict["new_location"] % location_name)
            Logging("New location found")
            asset_location = 1
        except WebDriverException:
            Logging("Cannot found new location")
    
    print("asset_categories " + str(asset_categories))
    
    if access_asset == True and asset_categories > 0 and asset_location > 0:
        Commands.ClickElement(asset_dict["manage_items"])
        Logging("Access Company Asset folder")
        
        Waits.Wait10s_ElementLoaded(asset_dict["list_footer"])
        time.sleep(1)
        Waits.WaitUntilPageIsLoaded(None)
        try:
            Commands.FindElement(asset_dict["pen_button"])
            permission = True

            Commands.ClickElement(asset_dict["asset_folder_1"])
            Logging("Access asset folder")
        except WebDriverException:
            permission = Asset_ValidateFolderPermission()

        list_data = CollectListData(list_footer=asset_dict["list_footer"], page_total=asset_dict["page_total"])
        assets = list_data["total_items"]
        pages =  list_data["total_pages"]
    else:
        Logging("User has no permission to access asset category")
        permission = False
        assets = 0
        pages = 0

    asset_data = {
        "permission": permission,
        "assets": assets,
        "pages": pages
    }

    return asset_data

def Asset_AddAsset():
    PrintYellow("[TEST CASE] ADD ASSET")
    
    Waits.Wait10s_ElementLoaded(asset_dict["pen_button"])
    Waits.WaitUntilPageIsLoaded(None)

    FindPushNoti()
    
    Commands.ClickElement(asset_dict["pen_button"])
    Logging("Add new asset - Click Pen button")

    Waits.WaitElementLoaded(15, data["editor"]["tox_editor_header"])
    Waits.WaitUntilPageIsLoaded(None)

    asset_id = "asset" + objects.date_id
    subject = "Asset " + objects.date_id

    Commands.InputElement(asset_dict["asset_id"], asset_id)
    Logging("Input asset id")

    time.sleep(2)

    title = Commands.InputElement(asset_dict["asset_name"], subject)
    asset_name = Functions.GetInputValue(asset_dict["asset_name"])
    Logging("Input asset name")

    Commands.ClickElement(asset_dict["purchase_date_calendar"])
    Logging("Click calendar button from Purchase Date")

    random_number = str(Functions.getRandomNumber_fromSpecificRange(1, 30))
    Commands.Wait10s_ClickElement(asset_dict["purchased_date"] % random_number)
    Logging("Select date for Purchase Date")

    Commands.ClickElement(asset_dict["asset_location"])
    Commands.Selectbox_ByIndex(asset_dict["asset_location"], 1)
    Logging("Select asset location")

    Commands.InputElement(asset_dict["asset_cost"], "1000")
    Logging("Input asset cost")

    Commands.Selectbox_ByVisibleText(asset_dict["asset_curreny"], "Currency")
    Logging("Select Currency")

    Commands.SwitchToFrame(data["editor"]["tox_iframe"])
    Commands.InputElement(data["editor"]["input_p"], objects.hanbiro_content)
    Commands.SwitchToDefaultContent()
    Logging("Add new asset - Input asset description")

    Commands.ClickElement(asset_dict["category_tree"])
    Logging("Open category tree")
    
    Waits.Wait10s_ElementLoaded(asset_dict["category_li"])
    category = Commands.FindElement(asset_dict["category_li"])
    category_name = category.text
    category.click()
    Logging("Select assset catgory ->", category_name)
    Waits.Wait10s_ElementLoaded(asset_dict["selected_category"] % category_name)

    time.sleep(1)

    FindPushNoti()
    save_button = Commands.ClickElement(asset_dict["save_button"])
    Logging("Add new asset - Click to save asset")

    Waits.WaitUntilPageIsLoaded(None)
    try:
        try:
            Waits.Wait10s_ElementLoaded(asset_dict["list_item"])
        except WebDriverException:
            save_button.click()
            Waits.Wait10s_ElementLoaded(asset_dict["list_item"])
        
        Waits.WaitElementLoaded(15, data["common"]["loading_dialog"])
        
        Commands.FindElement(asset_dict["list_asset_title"] % asset_name)
        TestCase_LogResult(**asset_tc["write"]["pass"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="asset", testcase="write", msg="click save")
        TestCase_LogResult(**asset_tc["write"]["fail"])

    return asset_name

def Asset_ViewDetails(asset_name):
    PrintYellow("[TEST CASE] VIEW ASSET")
    
    '''Click view asset and check if details are displayed in content
            Asset name
            Purchase Date # 1970/01/01
            Create history is recognized'''
    
    Waits.Wait10s_ElementLoaded(asset_dict["list_item"])
    
    try:
        Commands.ClickElement(asset_dict["list_asset_title"] % asset_name)
    except WebDriverException:
        Commands.ClickElements(asset_dict["list_item"], 0)
    finally:    
        Logging("View asset - Click new asset to view details")

    try:
        Waits.Wait10s_ElementLoaded(asset_dict["content_name"])
        TestCase_LogResult(**asset_tc["view"]["pass"])
        
        asset_data = []
        asset_name = Functions.GetElementText(asset_dict["content_name"])
        if bool(asset_name) == True:
            asset_data.append(True)
        else:
            asset_data.append(False)
            asset_tc["view"]["fail"].update({"description": "Asset name is displayed abnormally in content"})
            TestCase_LogResult(**asset_tc["view"]["fail"])

        purchase_date = str(Functions.GetElementText(asset_dict["content_purchase_date"]))
        if purchase_date.strip() != "1970/01/01":
            asset_data.append(True)
        else:
            asset_data.append(False)
            asset_tc["view"]["fail"].update({"description": "Purchase Date is saved incorrectly"})
            TestCase_LogResult(**asset_tc["view"]["fail"])

        try:
            Commands.FindElement(asset_dict["create_history"])
            asset_data.append(True)
        except WebDriverException:
            asset_data.append(False)
            asset_tc["view"]["fail"].update({"description": "Fail to view Item Activity"})
            TestCase_LogResult(**asset_tc["view"]["fail"])
        
        if False in asset_data:
            pass
        else:
            asset_tc["view"]["pass"].update({"description": "Asset data is correct"})
            TestCase_LogResult(**asset_tc["view"]["pass"])
        
        Commands.ClickElement(asset_dict["back_button"])
        Logging("View asset - Back from content")

        Waits.Wait10s_ElementLoaded(asset_dict["list_item"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="asset", testcase="view", msg="click view")
        TestCase_LogResult(**asset_tc["view"]["fail"])

def Asset_ImportAsset(domain_name, asset_folder_link):
    ''' import asset with file excel '''

    PrintYellow("[TEST CASE] IMPORT ASSET")

    Commands.NavigateTo(domain_name + asset_folder_link)

    # store the object of Workbook class in a variable
    wrkbk = openpyxl.Workbook()
    # to create a new sheet
    sh = wrkbk.create_sheet("Details", 2)

    asset_number = str(Functions.getRandomNumber_fromSpecificRange(1,100))
    date_import = objects.day + "/" + objects.month + "/" + objects.year

    asset_id = "asset_import" + asset_number
    asset_name = "Asset Import " + date_import

    # to set the value in row 2 and column 3
    header_group = ["ID", "Name", "Brand", "Model", "Serial", "PO Number", "Purchase Date", "Purchased From", "Price", "Currency"]
    header_column = 0
    for header in header_group:
        header_column += 1
        sh.cell(row=1, column=1).value = header
    
    import_data = [asset_id, asset_name, "Brand Import", "Model Import", "Serial Import", "PO Number Import", date_import, "Place of Purchase", "1000", "USD"]
    data_column = 0
    for import_element in import_data:
        data_column += 1
        sh.cell(row=1, column=1).value = import_element

    wrkbk.get_sheet_names()
    removed_sheet = wrkbk.get_sheet_by_name('Sheet')
    wrkbk.remove_sheet(removed_sheet)
    wrkbk.get_sheet_names()

    import_file = Files.asset_import
    wrkbk.save(import_file)

    Waits.Wait10s_ElementLoaded(asset_dict["import_button"])
    Commands.ClickElement(asset_dict["import_button"])

    Waits.WaitUntilPageIsLoaded(None)

    Commands.InputElement(asset_dict["file_uploader"], import_file)
    Commands.Selectbox_ByIndex(asset_dict["import_location"], 2)
    Logging("Asset Import - Select Location")

    time.sleep(1)

    Commands.ClickElement(asset_dict["import_save"])
    Logging("Asset Import - Save uploaded file")

    Waits.Wait10s_ElementLoaded(asset_dict["list_item"])
    try:
        Commands.FindElement(asset_dict["list_asset_title"] % asset_name)
        Logging("Asset Import - New asset is imported successfully")
    except WebDriverException:
        ValidateFailResultAndSystem("Asset Import - Fail to import new asset")

def Asset_ValidateNextPageList():
    PrintYellow("[MENU ASSET] MOVE PAGE")
    List_ValidateListMovingPage(asset_dict["list_target"], asset_dict["item_suf"], asset_dict["page_total"], asset_dict["nextpage_icon"])

def Asset_SearchAsset():
    PrintYellow("[TEST CASE] SEARCH ASSET")
    Waits.Wait10s_ElementLoaded(asset_dict["list_asset_1"])

    Waits.WaitUntilPageIsLoaded(None)

    assets = Functions.GetListLength("//div[contains(@class, 'asset-list message-item')]") -1
    if assets > 1:
        search_data = {
            "ID": asset_dict["search_id"],
            "Name": asset_dict["search_name"],
            "Model": asset_dict["search_model"]
        }

        selectbox_dict = {
            "ID": {"value": "asset_id"},
            "Name": {"value": "name"},
            "Model": {"value": "model"}
        }
    else:
        search_data = {
            "name": asset_dict["search_name"]
        }

        selectbox_dict = {
            "name": {"value": "name"}
        }

    search_dict = {}

    for search_label in search_data.keys():
        element_xpath = search_data[search_label]
        elements = Commands.FindElements(element_xpath)
        
        element_key = []
        for element in elements:
            element_text = str(element.text)
            if bool(element_text.strip()) == True:
                element_key.append(element_text)
        
        element_key = list(dict.fromkeys(element_key))
        search_data[search_label] = element_key
        
        if len(search_data[search_label]) >= 1:
            search_dict[search_label] = selectbox_dict[search_label]
            search_dict[search_label]["key"] = element_key[0]

    search_asset_dict = dict(asset_dict["search_details"])
    search_asset_dict["search_dict"] = search_dict
    SearchDetailsBySelectBox(**search_asset_dict)

def Asset_DefineListData():
    try:
        list_data = CollectListData(list_footer=asset_dict["list_footer"], page_total=asset_dict["page_total"])
        assets = list_data["total_items"]
        pages =  list_data["total_pages"]
    except:
        assets = 0
        pages =  0
    
    list_data = {
        "assets": assets,
        "pages": pages
    }

    return list_data

def AssetExecution():
    asset_data = Asset_AccessMenu()
    if bool(asset_data["permission"]) == True:
        asset_name = Asset_AddAsset()
        list_data = Asset_DefineListData()
        Asset_ViewDetails(asset_name)
        #Asset_ImportAsset()
        Asset_SearchAsset()
        Asset_ValidateNextPageList()
    ValidateUnexpectedModal()  