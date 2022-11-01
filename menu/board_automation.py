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
    global board_dict, board_tc
    board_dict = dict(data["board"])
    board_tc = dict(data["testcase_result"]["board"])

def Board_ValidatePermission():
    Waits.WaitUntilPageIsLoaded(board_dict["list_footer"])
    write_board = None

    current_url = DefineCurrentURL()
    print(current_url)
    if "/comp_" in current_url:
        comp_li = Functions.GetElementAttribute(board_dict["comp_board"], "class")
        if "active" not in comp_li:
            Commands.ClickElement(board_dict["comp_dropdown"])
            Logging("Open Company Board")
        else:
            Logging("Company board is active")

        try:
            Commands.FindElement(board_dict["comp_folder"])
            Logging("Company board folder found")
            write_board = True
        except WebDriverException:
            try:
                Commands.ClickElement(board_dict["my_dropdown"])
                Logging("Click My Board dropdown button")
                Waits.WaitElementLoaded(3, board_dict["my_folder"])
                Commands.ClickElement(board_dict["my_folder"])
                Logging("Access My board folder")
                Waits.Wait10s_ElementLoaded(board_dict["list_footer"])
                write_board = True
            except WebDriverException:
                Logging("User has no permission to write in My board")
        
    return write_board

def Board_WriteNewBoard():
    PrintYellow("[TEST CASE] WRITE BOARD")
    try:
        Waits.Wait10s_ElementLoaded(data["common"]["loading_dialog"])
        FindPushNoti()
    except WebDriverException:
        pass
    
    time.sleep(1)

    Waits.WaitElementLoaded(15, board_dict["red_pen"])
    Commands.Wait10s_ClickElement(board_dict["red_pen"])
    Logging("Write - Click Create button")

    Waits.WaitElementLoaded(15, data["editor"]["tox_iframe"])
    Waits.WaitUntilPageIsLoaded(None)

    title = Commands.InputElement(board_dict["title"], objects.hanbiro_title)
    Logging("Write - Input title / subject")
    Logging(">>> Title: [" + title.get_attribute("value") + "] is input")
    
    Commands.ScrollDown()
    Waits.Wait10s_ElementLoaded(data["editor"]["tox_iframe"])
    time.sleep(1)
    
    Commands.SwitchToFrame(data["editor"]["tox_iframe"])
    time.sleep(1)

    try:
        Commands.InputElement(data["editor"]["input_p"], objects.hanbiro_content)
        Logging("Content is empty - Input content")
    except WebDriverException:
        try:
            Commands.InputElement(data["editor"]["empty_content"], objects.hanbiro_content)
            Logging("Content is empty - Input content")
        except WebDriverException:
                Commands.InputElement(data["editor"]["included_content"], objects.hanbiro_content)
                Logging("Input content")

    Commands.SwitchToDefaultContent()
    
    CloseAutosave()

    title = Commands.FindElement(board_dict["title"])
    board_title = title.get_attribute("value")
    if bool(board_title) == False:
        title.send_keys(objects.hanbiro_title)
        board_title = title.get_attribute("value")
        Logging("Input board title")

    Commands.ScrollUp()
    
    try:
        Commands.Selectbox_ByIndex(board_dict["select_box"], 1)
        Waits.WaitElementLoaded(3, board_dict["folder_selected"])
    except WebDriverException:
        Logging("Folder selectbox not found")

    FindPushNoti()
    Commands.ClickElement(board_dict["send_button"])
    Logging("Write - Save board")
    
    try:
        Waits.Wait10s_ElementLoaded(board_dict["board_content"])
        TestCase_LogResult(**board_tc["write"]["pass"])
        write_board = True
    except WebDriverException:
        write_board = False
        TCResult_ValidateAlertMsg(menu="board", testcase="write", msg="click save board")
        TestCase_LogResult(**board_tc["write"]["fail"])

        Commands.ClickElement(board_dict["company_board"])
        Logging("Access Company Board")

        Waits.WaitUntilPageIsLoaded(board_dict["list_footer"])
    
    return write_board

def AccessFolder(domain_name, folder_position, board_folder_link):
    Commands.Wait10s_ClickElement(board_dict["folder_target"] + "[" + str(folder_position) + "]/a")
    Logging("Access Board Folder - Click folder")

    try:
        subfolder = Waits.WaitElementLoaded(3, board_dict["folder_target"] + "[" + str(folder_position) + "]/ul/li/a/span/span")
        subfolder.click()
        Logging("Access Board Folder - Access Subfolder")
        Board_WriteNewBoard()
    except:
        Logging("Access Board Folder - Access Folder successfully")
        Waits.WaitElementLoaded(3, board_dict["list_header"])
        folder_li = Functions.GetElementText(board_dict["folder_target"] + "[" + str(folder_position) + "]/a/span/span")
        header = Functions.GetElementText(board_dict["list_header"])
        if header == folder_li:
            Board_WriteNewBoard()

def SearchBoard():
    try:
        Commands.FindElement(board_dict["search_input_list"][1])
        Logging("Search - Search posts in List Folder")
        wrapper(searchInput, board_dict["search_input_list"])
    except WebDriverException:
        Logging("Search - Search posts in Gallery Folder")
        wrapper(searchInput, board_dict["search_input_gallery"])

def Board_AccessFolder(domain_name, board_folder_link):
    folder_position = 1
    try:
        Logging("Folder can be accessed")
        AccessFolder(domain_name, folder_position, board_folder_link)
    except WebDriverException:
        Logging("Cannot access folder")
        folder_position = folder_position + 1
        AccessFolder(domain_name, folder_position, board_folder_link)

def Board_ContentMultipleView():
    Waits.Wait10s_ElementLoaded(board_dict["list_target"])
    
    item_number = Functions.GetListLength(board_dict["list_target"])
    Logging("---------- Define Item Position - List Number: " + str(item_number))

    item_position = Functions.getRandomNumber_fromSpecificRange(1, item_number)
    Logging("---------- Define Item Position - Item_position: " + str(item_position))

    view_result = []

    try:
        item_xpath = board_dict["list_target"] + "[" + str(item_position) + "]" + board_dict["item_suf"]
        item_text = Functions.GetElementText(item_xpath)
        Commands.ClickElement(item_xpath)
        Logging("View Normal Content - Click on item")

        try:
            Waits.Wait10s_ElementLoaded(data["editor"]["tox_iframe"])
        except WebDriverException:
            Waits.Wait10s_ElementLoaded(data["board"]["board_content_div"])
        
        if item_text.strip() in Functions.GetPageSource():
            view_result.append(True)
        else:
            view_result.append(False)
            board_tc["view"]["fail"].update({"description": "Board title is incorrect: " + str(item_text)})
            TestCase_LogResult(**board_tc["view"]["fail"])

        Commands.ClickElement(board_dict["back_button"])

        Waits.WaitUntilPageIsLoaded(board_dict["list_footer"])
        time.sleep(1)
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="board", testcase="view", msg="click view board")
        TestCase_LogResult(**board_tc["view"]["fail"])

    return view_result

def Board_WriteComment():
    PrintYellow("[MENU BOARD] - WRITE COMMENT")

    try:
        Waits.Wait10s_ElementLoaded(data["editor"]["tox_iframe"])
    except WebDriverException:
        pass

    try:
        Waits.WaitElementLoaded(2, data["board"]["importing"])
        Waits.Wait10s_ElementInvisibility(data["board"]["importing"])
    except WebDriverException:
        pass

    Commands.SwitchToFrame(data["editor"]["tox_iframe"])
    
    Waits.Wait10s_ElementLoaded(data["editor"]["input_p"])
    time.sleep(1)
    board_comment = "Comment is created at %s" % objects.date_time
    Commands.InputElement(data["editor"]["input_p"], board_comment)
    Logging("Write comment - Input comment") 
    
    Commands.SwitchToDefaultContent()

    Commands.ClickElement(board_dict["comment_save"])
    Logging("Write comment - Click Save comment")

    try:
        Waits.Wait10s_ElementLoaded(board_dict["board_comment"] % board_comment)
        TestCase_LogResult(**board_tc["comment"]["pass"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="board", testcase="comment", msg="click save comment")
        TestCase_LogResult(**board_tc["comment"]["fail"])

    Commands.Wait10s_ClickElement(board_dict["back_button"])

    time.sleep(1)
    Waits.Wait10s_ElementLoaded(board_dict["list_footer"])

def Board_CheckUnread():
    PrintYellow("[MENU BOARD] - CHECK UNREAD")

    Commands.Wait10s_ClickElement(board_dict["company_board"])
    Logging("Check unread - Access Company Board")

    Waits.Wait10s_ElementLoaded(board_dict["list_footer"])
    Logging("Check unread - Wait until the presence of list footer")

    Waits.WaitUntilPageIsLoaded(None)

    try:
        unread_counter = Commands.ClickElement(board_dict["unread_counter"])
        unread_number = int(unread_counter.text)
        Logging("unread_number: " + str(unread_number))
        Logging("Check unread - Click unread list number")
    except WebDriverException:
        unread_number = 0
        Logging("Check unread - unread counter is not displayed")

    Waits.WaitUntilPageIsLoaded(None)

    if unread_number > 0:
        list_footer_unread = Functions.GetElementText(board_dict["list_footer"])
        list_unread = int(list_footer_unread.split(" ")[1].replace(",", ""))
        Logging("list_unread: " + str(list_unread))
    else:
        try:
            company_unread = Functions.GetElementText("//span[@class='menu-text' and text()=' Company Board']/following-sibling::span[contains(@class, 'badge')]")
            Logging("Company unread counter and list unread counter are not same")
            TestCase_LogResult(**board_tc["check_unread"]["fail"])
        except WebDriverException:
            TestCase_LogResult(**board_tc["check_unread"]["pass"])

    # ---> Reset unread list
    Commands.ClickElement(board_dict["dropdown_toggle"])
    Logging("Board List - Click dropdown menu of board list")    
    
    Commands.Wait10s_ClickElement(board_dict["dropdown_view_all"])
    Logging("Click View All from dropdown menu")

    Waits.WaitUntilPageIsLoaded(None)
    
    time.sleep(1)

def Board_SearchCompanyBoard():
    PrintYellow("[MENU BOARD] - SEARCH BOARD")
    Waits.Wait10s_ElementLoaded(board_dict["list_target"])

    list1 = ValidateListTotal(board_dict["list_footer"])

    Waits.Wait10s_ElementLoaded(board_dict["search_input"][4])
    Commands.InputElement_2Values(board_dict["search_input"][4], objects.hanbiro_title, Keys.ENTER)
    Logging("Search Input - Enter key words")

    search_keys = Functions.GetInputValue(board_dict["search_input"][4])
    Logging("Search Input - Key word: " + search_keys)

    time.sleep(3)

    list2 = ValidateListTotal(board_dict["list_footer"])
    if list1 != list2:
        TestCase_LogResult(**board_tc["search"]["pass"])
    else:
        TestCase_LogResult(**board_tc["search"]["fail"])

    Commands.InputElement(board_dict["search_input"][4], Keys.RETURN)
    Waits.WaitUntilPageIsLoaded(board_dict["list_footer"])
    time.sleep(1)

def Board_ViewCompanyBoard():
    PrintYellow("[MENU BOARD] - VIEW COMPANY BOARD")

    Commands.ClickElement(data["board"]["company_board_span"])
    Waits.Wait10s_ElementLoaded(board_dict["list_footer"])

    total_posts = int(str(Functions.GetElementText(board_dict["list_footer"])).split(" ")[1].replace(",", ""))
    print("total_posts " + str(total_posts))
    if total_posts > 0:
        results = []
        for i in range(1,3):
            result = Board_ContentMultipleView()
            list(results).extend(result)
        
        if False not in results:
            dict(board_tc["check_unread"]["pass"]).update({"testcase": "View Company board"})
            dict(board_tc["check_unread"]["pass"]).update({"description": "View multiple posts in Company board successfully"})
            TestCase_LogResult(**board_tc["check_unread"]["pass"])

        time.sleep(1)
        try:
            Commands.ClickElement(board_dict["back_button"])
        except WebDriverException:
            pass
        Waits.Wait10s_ElementLoaded(board_dict["list_footer"])
        Commands.ReloadBrowser(board_dict["list_target"])

def Board_ValidateNextPageList():
    PrintYellow("[MENU BOARD] MOVE PAGE")
    total_posts = int(str(Functions.GetElementText(board_dict["list_footer"])).split(" ")[1].replace(",", ""))
    if total_posts > 0:
        List_ValidateListMovingPage(board_dict["list_target"], board_dict["item_suf"], board_dict["page_total"], board_dict["nextpage_icon"])

def Board_ValidateCurrentPage():
    Waits.WaitUntilPageIsLoaded(board_dict["list_footer"])

    current_url = DefineCurrentURL()

    try:
        list_footer = Functions.GetElementText(board_dict["list_footer"])
        board_total = int(list_footer.split(" ")[1].replace(",", ""))
    except WebDriverException:
        board_total = None

    try:
        page_total = int(Functions.GetElementText(board_dict["page_total"]))
    except WebDriverException:
        page_total = None

    current_page = {
        "url": current_url,
        "item_total": board_total,
        "page_total": page_total
    }

    return current_page

def Board_SearchDetails():
    current_list = Board_ValidateCurrentPage()
    total_posts = current_list["item_total"]
    current_page = current_list["url"]
    print("total_posts " + str(total_posts))
    
    if total_posts > 0:
        search = True
    else:
        
        Logging("Cannot do searching at current board list %s" % current_page)
        search = False

    if search == True:
        Waits.WaitUntilPageIsLoaded(board_dict["board_div"])
        div_xpath = board_dict["board_div"]
        title_xpath_normal = div_xpath + "/span[8]/span[1]/span"
        title_xpath_secure = div_xpath + "/span[8]/span/span[2]"

        try:
            title = Functions.GetElementText(title_xpath_secure)
        except WebDriverException:
            title = Functions.GetElementText(title_xpath_normal)

        search_dict = {
            "title": {
                "key": title,
                "value": "subject",
            } 
        }

        creator_list = []
        creators = Commands.FindElements(board_dict["list_creator"])
        i=0
        # ---> Collect creator name and append to creator_list
        for i in range(len(creators)):
            creator_name = creators[i].text
            
            if creator_name != "Writer" and creator_name not in creator_list:
                # ---> Ignore adding column header "Writer"
                # ---> Ignore adding duplicate creator name
                creator_list.append(creator_name)
            
            i+=1
        
        if len(creator_list) > 1:
            # ---> There are more than one different creator -> can do searching with creator
            creator = creator_list[0]
            search_dict.update({"creator":{"key": creator,"value": "name",}})

        search_board_dict = dict(board_dict["search_details"])
        search_board_dict["search_dict"] = search_dict
        SearchDetailsBySelectBox(**search_board_dict)

def BoardExecution():
    PrintYellow("-------------------------------------------")
    PrintYellow("[MENU BOARD] ACCESS MENU")
    access_result = AccessGroupwareMenu(name="board,Board", page_xpath=board_dict["list_footer"])
    write_permission = Board_ValidatePermission()
    if write_permission == True:
        create_board  = Board_WriteNewBoard()
        if create_board == True:
            Board_WriteComment()
    #current_list = Board_ValidateCurrentPage()
    Board_ViewCompanyBoard()
    Board_CheckUnread()
    Board_SearchDetails()
    Board_ValidateNextPageList()
    ValidateUnexpectedModal()


