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

CalendarIntepreter = {
    "01": "January",
    "02": "February",
    "03": "March",
    "04": "April",
    "05": "May",
    "06": "June",
    "07": "July",
    "08": "August",
    "09": "September",
    "10": "October",
    "11": "November",
    "12": "December"}

def Dictionaries():
    global calendar_dict, calendar_tc, CalendarIntepreter
    calendar_dict = dict(data["calendar"])
    calendar_tc = dict(data["testcase_result"]["calendar"])

    CalendarIntepreter = {
    "01": "January",
    "02": "February",
    "03": "March",
    "04": "April",
    "05": "May",
    "06": "June",
    "07": "July",
    "08": "August",
    "09": "September",
    "10": "October",
    "11": "November",
    "12": "December"}

def Calendar_AccessCalendarMenu():
    PrintYellow("-------------------------------------------")
    PrintYellow("[MENU CALENDAR] ACCESS MENU")

    access_menu = None
    
    time.sleep(1)
    AccessGroupwareMenu(name="calendar,Calendar", page_xpath="//*[@id='tuicalendar']/div/div[1]/div[2]/div[1]/div/div[2]/div[1]")
    Waits.WaitUntilPageIsLoaded(None)

    try:
        Waits.WaitElementLoaded(20, calendar_dict["weekday_grid"])
        Logging("Calendar Folder Type is being used")
        access_menu = True
    except WebDriverException:
        Logging("Calendar Category is being used")
        Commands.ClickElement("//a[@ng-click='goAdminSettings($event)']")
        Logging("Access Admin Settings of Calendar Category")

        Waits.WaitElementLoaded(20, "//button[@ng-click='changeMode()']")
        Waits.WaitUntilPageIsLoaded(None)

        Commands.ClickElement("//button[@ng-click='changeMode()']")
        Logging("Switch to Folder Type")
        Waits.WaitUntilPageIsLoaded(None)

        try:
            Waits.WaitElementLoaded(20, calendar_dict["my_calendar_sub"])
            access_menu = True
        except WebDriverException:
            access_menu = False
    
    if access_menu == True:
        time.sleep(1)
        Waits.WaitUntilPageIsLoaded(None)

        Commands.ReloadBrowser(calendar_dict["weekday_grid"])
        Waits.WaitElementLoaded(20, calendar_dict["weekday_grid"])
        Waits.WaitUntilPageIsLoaded(None)

        time.sleep(1)

        Commands.ClickElement(calendar_dict["my_calendar_sub"])
        Logging("Calendar Folder - Access My Calendar sub menu")

        Waits.WaitUntilPageIsLoaded(calendar_dict["calendar_folder"])
        calendar_folder = Commands.ClickElements(calendar_dict["calendar_folder"], 0)
        #time.sleep(1)
        #calendar_folder[0].click()
        Logging("Calendar Folder - Access My Calendar folder")

        Waits.WaitUntilPageIsLoaded(None)
        time.sleep(1)

    return access_menu

def Calendar_AddEvent(time_event, event_public, repeated):
    PrintYellow("[ADD EVENT - FOLDER TYPE]")
    
    Waits.Wait10s_ElementLoaded(calendar_dict["my_write_button"])
    FindPushNoti()

    Commands.ClickElement(calendar_dict["my_write_button"])
    Logging("Click Write button")

    current_time = DefineCurrentTime()
    Waits.Wait10s_ElementLoaded(data["resource"]["title"])
    title_input  = "Event is created at " + str(current_time)

    Commands.InputElement(calendar_dict["title_input"], title_input)
    event_title = Functions.GetInputValue(calendar_dict["title_input"])
    Logging("Calendar Folder - Write - Input event title")
    
    start_date = Functions.GetInputValue(calendar_dict["start_date"])
    Logging("Calendar Folder - Write - Start Date: " + start_date)
    
    end_date = Functions.GetInputValue(calendar_dict["end_date"])
    
    Logging("Calendar Folder - Write - End Date: " + end_date)

    time.sleep(1)
    
    if time_event == True:
        Commands.ClickElement(calendar_dict["fullday_checkbox"])
        Logging("Uncheck Full Day")

        Waits.Wait10s_ElementLoaded(calendar_dict["fullday_uncheck"])

        start_hour_text = Functions.GetInputValue(calendar_dict["start_hour"])
        Logging("start hour " + start_hour_text)
        if int(start_hour_text) < 10:
            start_hour_text = "0" + start_hour_text 
        
        start_min_text = Functions.GetInputValue(calendar_dict["start_min"])
        if start_min_text != "? string:59 ?":
            Logging("start minute " + start_min_text)
            if int(start_min_text) < 10:
                start_min_text = "0" + start_min_text
        else:
            Commands.Selectbox_ByVisibleText(calendar_dict["start_min"], "0")

        end_hour_text = Functions.GetInputValue(calendar_dict["end_hour"])
        Logging("end hour " + end_hour_text)
        if int(end_hour_text) < 10:
            end_hour_text = "0" + end_hour_text

        end_min_text = Functions.GetInputValue(calendar_dict["end_min"])
        if end_min_text != "? string:59 ?":
            Logging("end min " + end_min_text)
            if int(end_min_text) < 10:
                end_min_text = "0" + end_min_text
        else:
            Commands.Selectbox_ByVisibleText(calendar_dict["end_min"], "0")

        Commands.Selectbox_ByValue(calendar_dict["timezone"], "Asia/Saigon")
        Logging("Select Time Zone")
        event_time = True
    else:
        event_time = None

    if event_public == True:
        Commands.ClickElement(calendar_dict["public_button"])
        Logging("Click Public button")
        time.sleep(2)
        public_mode = True
    else:
        public_mode = False

    Commands.ClickElement(calendar_dict["details_button"])
    Logging("Calendar Folder - Click Details button")

    try:
        Commands.FindElement(calendar_dict["details_active"])
        Logging("Details tab is active")
    except WebDriverException:
        Commands.ClickElement(calendar_dict["details_button"])
        Logging("Click Details button")

    Commands.Wait10s_InputElement(calendar_dict["calendar_location"], "Event Location")

    Commands.Wait10s_InputElement(calendar_dict["calendar_content"], "Event Description")

    if repeated == True:
        Commands.Wait10s_ClickElement(calendar_dict["repeat_tab"])
        Commands.Wait10s_ClickElement(calendar_dict["active_repeat"])

        time.sleep(2)

        Commands.Selectbox_ByVisibleText(calendar_dict["repeat_type"], "Daily")
        Logging("Select repeat type Daily")

        Commands.Selectbox_ByVisibleText(calendar_dict["repetition"], "1")
        Logging("Select repeat every 1")

        try:
            Commands.FindElement(calendar_dict["unlimited_isactive"])
            Logging("Unlimted Repeat is active")
            
            Commands.ClickElement(calendar_dict["unlimited_toggle"])
            Logging("Activate limited repeat")
        except WebDriverException:
            Logging("Unlimited Repeat is not active")

        Waits.WaitElementLoaded(5, calendar_dict["unlimited_datepicker"])

        # Active datepicker
        Commands.ClickElement(calendar_dict["datepicker_button"])
        Waits.WaitElementLoaded(5, calendar_dict["datepicker_days"])
        
        # Asset name of month in datepicker and keep clicking prev button until datepicker displays current month
        # **** Default repetition: 2 months from the current date
        month = Commands.FindElement(calendar_dict["datepicker_switch"])
        i = 0
        for i in range(1,5):
            i += 1
            Commands.ClickElement(calendar_dict["datepicker_prev"])
            target_month = str(month.text).split(" ")[0]
            if target_month == CalendarIntepreter[objects.month]:
                Logging("Date picker - Access current month")
                
                # After the current month is found, select the last day of current month and break loop
                try:
                    end_day = Commands.FindElement(calendar_dict["end_month"])
                except WebDriverException:
                    end_day = Commands.FindElement(calendar_dict["end_month_newrow"])
                
                end_month = end_day.text
                Logging("end_month " + end_month)
                end_repetition = int(end_month)
                Logging("end date of repetition " + str(end_repetition))
                
                end_day.click()
                break

        current_day = int(objects.day.lstrip("0"))
        total_repetition = end_repetition - current_day + 1
        Logging("Total of repeated events: " + str(total_repetition))
        
        time.sleep(1)
        repetition = True
    else:
        total_repetition = 0
        repetition = False

    FindPushNoti()

    Commands.ClickElement(calendar_dict["save_button"])
    Logging("Calendar Folder - Write - Save event")

    try:
        Waits.WaitElementLoaded(3, calendar_dict["alert_select_category"])
        Commands.ClickElement(calendar_dict["close_warning"])
        Logging("Close warning modal 'Please add category'")
    except WebDriverException:
        pass

    Waits.WaitUntilPageIsLoaded(None)

    if event_title in Functions.GetPageSource():
        TestCase_LogResult(**calendar_tc["write"]["pass"])
        
        if repetition == True:
            repeated_events = Functions.GetListLength(calendar_dict["event_@title"] % event_title)
            Logging("Total of repeated event found on calendar: " + str(repeated_events))
            
            if repeated_events == total_repetition:
                calendar_tc["view"]["pass"].update({"testcase": "Add repeated event", "description": "All repeated events are added successfully"})
                TestCase_LogResult(**calendar_tc["view"]["pass"])
            else:
                calendar_tc["view"]["pass"].update({"testcase": "Add repeated event", "description": "Fail to save all add repeated events"})
                TestCase_LogResult(**calendar_tc["view"]["fail"])
    else:
        TCResult_ValidateAlertMsg(menu="calendar_folder", testcase="write", msg="click save event")
        TestCase_LogResult(**calendar_tc["write"]["fail"])

    event_data = {
        "title": event_title,
        "start_date": start_date,
        "end_date": end_date,
        "event_time": event_time,
        "public_mode": public_mode,
        "repetition": repetition
    }

    return event_data

def CalendarCategory_AddEvent():
    PrintYellow("[ADD EVENT - CATEGORY TYPE]")
    Commands.ClickElement(data["calendar_category"]["write_button"])
    Logging("Calendar Category - Click add event button")
    
    # Select Calendar Category - Working with quick event modal
    Commands.Wait10s_ClickElement("//*[@id='tui-full-calendar-schedule-calendar']")
    Commands.Wait10s_ClickElement("//li/span[contains(.,'My Calendar')]")
    Commands.Wait10s_ClickElement("//*[@id='tui-full-calendar-schedule-calendar']")
    Logging("Calendar Category - Quick Create - Select category My Calendar")
    
    Commands.InputElement("//*[@id='tui-full-calendar-schedule-title']", objects.hanbiro_title)
    event_title = Functions.GetInputValue("//*[@id='tui-full-calendar-schedule-title']")
    Logging("Calendar Category - Quick Create - Input event title")
    
    start_date = Functions.GetInputValue(calendar_dict["start_date"])
    Logging("Calendar Category - Quick Create - Start date: " + start_date)

    Commands.InputElement("//*[@id='tui-full-calendar-schedule-title']", "Test event location")
    Logging("Calendar Category - Quick Create - Input event location: Test event location")
    
    Commands.ClickElement(data["calendar_category"]["save_button"])
    Logging("Click Save event")

    try:
        Waits.WaitUntilPageIsLoaded(None)
        if event_title in Functions.GetPageSource():
            Logging("Calendar Category - Quick Create - Event is added successfully")
            Logging(objects.testcase_pass)
        
        Commands.ClickElement("//span[contains(text(),' " + objects.hanbiro_title + "')]")
        Logging("Calendar Category - View - Open View mode")

        Waits.Wait10s_ElementLoaded(calendar_dict["view_container"])

        event_date = Functions.GetElementText(calendar_dict["viewmode_date"]).strip().replace(".", "/")
        Logging("Calendar Category - View - View mode - Event date: " + event_date)
        
        if event_date == start_date:
            Logging("Calendar Category - View - Event is added with same date")
            Logging(objects.testcase_pass)
        else:
            Logging("Calendar Category - View - Event is not added with same date")
            Logging(objects.testcase_fail)

        Commands.ClickElement(calendar_dict["close_view"])
        Logging("Calendar Category - View - Close view mode")
    except WebDriverException:
        Logging("Calendar Category - Quick Crate - Failed to add event")
        Logging(objects.testcase_fail)
        Logging("Calendar Category - View - Close view mode")
    
    return event_title

def Calendar_EditEvent(event_title):
    PrintYellow("[MENU CALENDAR] EDIT EVENT")

    Commands.ClickElement(calendar_dict["event_span"] % event_title)
    Logging("Calendar Folder - View - Open View mode")

    Commands.Wait10s_ClickElement(calendar_dict["details_view"])
    Logging("Calendar View - Click Details button")

    Waits.Wait10s_ElementLoaded(calendar_dict["detailview_date"])
    event_date = Functions.GetElementAttribute(calendar_dict["detailview_date"], "title")
    Logging("date before edit: " + event_date)

    view_date = []
    
    if "~" in event_date:
        event_startdate = event_date.split("~")[0].strip()
        Logging("event start date: " + event_startdate)
        event_enddate = event_date.split("~")[1].strip()
        Logging("event end date: " + event_enddate)
    else:
        event_fulldate = event_startdate = event_enddate = event_date
        Logging("event full day date: " + event_fulldate)

    view_date.append(event_startdate)
    view_date.append(event_enddate)

    Commands.Wait10s_ClickElement(calendar_dict["edit_button"])
    Logging("Calendar View - Click Edit button")

    Waits.Wait10s_ElementLoaded(data["resource"]["title"])
    time.sleep(1)
    
    edit_date = []
    date_xpaths = [calendar_dict["start_date"], calendar_dict["end_date"]]
    for date_xpath in date_xpaths:
        date_value = Functions.GetInputValue(date_xpath)
        edit_date.append(date_value)
        Logging("Calendar Folder - Edit - Start Date: " + date_value)
    
    date = {
        "start_date": {
            "before_edit": view_date[0],
            "after_edit": edit_date[0]
        },
        "end_date": {
            "before_edit": view_date[1],
            "after_edit": edit_date[1]
        }
    }

    for date_type in date.keys():
        if date[date_type]["after_edit"] == date[date_type]["before_edit"]:
            calendar_tc["edit"]["pass"].update({"description": str(date_type) + " is same as created date"})
        else:
            calendar_tc["edit"]["fail"].update({"description": str(date_type) + " is different from created date"})

    FindPushNoti()
    Commands.Wait10s_ClickElement(calendar_dict["save_edit_button"])
    Logging("Calendar Edit - Click Save button")

    try:
        Waits.WaitElementLoaded(2, calendar_dict["alert_select_category"])
        Commands.ClickElement(calendar_dict["close_warning"])
        Logging("Close warning modal 'Please add category'")
    except WebDriverException:
        pass
    
    Waits.WaitUntilPageIsLoaded("//span[contains(., '" + event_title + "')]")
    time.sleep(1)

def Calendar_ViewEvent(event_title, repeated, start_date, end_date):
    PrintYellow("[MENU CALENDAR] VIEW EVENT")
    Waits.Wait10s_ElementLoaded(calendar_dict["event_@title"] % event_title)
    
    Commands.ClickElements(calendar_dict["event_@title"] % event_title, 0)
    Logging("Calendar Folder - View - Open View mode")

    try:
        Waits.Wait10s_ElementLoaded(calendar_dict["view_container"])
        time.sleep(1)
        view_container =  True
    except WebDriverException:
        calendar_tc["view"]["fail"].update({"description": "Cannot open view pop-up"})
        TestCase_LogResult(**calendar_tc["view"]["fail"])
        view_container = False

    if view_container == True:
        Waits.Wait10s_ElementLoaded(calendar_dict["viewmode_date"])
        if repeated == True:
            Logging("View Repeated Event")
        
            event_date = Functions.GetElementText(calendar_dict["event_date"]).strip()
            startdate_check = event_date.split(" ~ ")[0].split(" ")[0]
            Logging("start date when view " + startdate_check)
            
            enddate_check = event_date.split(" ~ ")[1].split(" ")[0]
            Logging("end date when view " + enddate_check)

            event_dates = {
                "start_date": {
                    "create": start_date,
                    "view": startdate_check
                },
                "end_date": {
                    "create": end_date,
                    "view": enddate_check
                }
            }
            for event_date in event_dates.keys():
                if event_dates[event_date]["view"] == event_dates[event_date]["create"]:
                    calendar_tc["view"]["pass"].update({"testcase": "View Repeated Events", "description": "Repeated event is saved with same " + str(event_date)})
                    TestCase_LogResult(**calendar_tc["view"]["pass"])
                else:
                    calendar_tc["view"]["fail"].update({"testcase": "View Repeated Events", "description": "Fail to save repeated event with " + str(event_date)})
                    TestCase_LogResult(**calendar_tc["view"]["fail"])
        else:
            Logging("View Single Event")

            event_date = Functions.GetElementText(calendar_dict["viewmode_date"]).strip()
            Logging(event_date)
            
            if event_date == start_date:
                calendar_tc["view"]["pass"].update({"description": "Event is added with same date"})
                TestCase_LogResult(**calendar_tc["view"]["pass"])
            else:
                calendar_tc["view"]["fail"].update({"description": "Event date is not saved with same date"})
                TestCase_LogResult(**calendar_tc["view"]["fail"])

        event_details = {
            "Location": calendar_dict["details_location"],
            "Description": calendar_dict["event_desc"]
        }

        for event_detail in event_details.keys():
            try:
                Commands.FindElement(event_details[event_detail])
                Logging("Calendar Folder - View - Event " + str(event_detail) + " is displayed in content")
            except WebDriverException:
                calendar_tc["view"]["fail"].update({"description": "Event " + str(event_detail) + " is not displayed in content"})
                TestCase_LogResult(**calendar_tc["view"]["fail"])

        Waits.WaitUntilPageIsLoaded(None)
        time.sleep(1)

        Commands.ClickElement(calendar_dict["close_view"])
        Logging("Calendar Folder - View - Close view mode")

        Waits.WaitUntilPageIsLoaded(calendar_dict["event"])
        time.sleep(1)

def Calendar_AccessPublicCalendar():
    time.sleep(1)

    Waits.Wait10s_ElementClickable(calendar_dict["public_calendar"])
    Commands.ClickElement(calendar_dict["public_calendar"])
    Logging("Acccess Public Calendar")

    Waits.WaitElementLoaded(20, calendar_dict["weekday_grid"])
    Waits.WaitUntilPageIsLoaded(None)
    time.sleep(1)

def Calendar_AccessMyCalendar():
    time.sleep(1)

    Commands.ClickElement(calendar_dict["my_calendar_sub"])
    Logging("Access My Calendar sub menu")
    
    Waits.Wait10s_ElementLoaded(calendar_dict["calendar_folder"])
    calendar_folder = Commands.ClickElements(calendar_dict["calendar_folder"], 0)
    #time.sleep(1)
    #calendar_folder[0].click()
    Logging("Calendar Folder - Access My Calendar folder")

    Waits.WaitUntilPageIsLoaded(None)
    time.sleep(1)

def Calendar_ImportEvent():
    PrintYellow("[MENU CALENDAR] IMPORT EVENT")
    
    Waits.WaitUntilPageIsLoaded(None)
    Waits.WaitElementLoaded(30, calendar_dict["weekday_grid"])
  
    event_title1 = Functions.GetListLength(calendar_dict["event"])
    print("event_title1 " + str(event_title1))

    Commands.Wait10s_ClickElement(calendar_dict["more_button"])
    Logging("Calendar Import - Click More button")

    Commands.ClickElement(calendar_dict["import_button"])
    Logging("Calendar Import - Click Import button")

    # store the object of Workbook class in a variable
    wrkbk = openpyxl.Workbook()

    # to create a new sheet
    sh = wrkbk.create_sheet("Details", 2)

    date_import = objects.year + "-" + objects.month + "-" + objects.day
    fullday_title = "Full Day Event: " + date_import
    clocked_title = "Time Event: " + date_import

    import_dict = {
        "row_header": ["Start Date", "End Date", "Time", "Title", "Memo", "Place"],
        "time_event": [date_import, date_import, "09:00 - 15:00", clocked_title, "Memo 1", "Place 1"],
        "full_day": [date_import, date_import, "Full Day", fullday_title, "Memo 2", "Place 2"]
    }
    
    column_number = 0
    while column_number < 6:
        column_number += 1
        cell_value = int(column_number-1)
        sh.cell(row=1, column=column_number).value = list(import_dict["row_header"])[cell_value]
        sh.cell(row=2, column=column_number).value = list(import_dict["time_event"])[cell_value]
        sh.cell(row=3, column=column_number).value = list(import_dict["full_day"])[cell_value]

    wrkbk.get_sheet_names()
    removed_sheet = wrkbk.get_sheet_by_name('Sheet')
    wrkbk.remove_sheet(removed_sheet)
    wrkbk.get_sheet_names()

    # to save the workbook
    import_file = Files.calendar_import
    wrkbk.save(import_file)

    Commands.Wait10s_InputElement(calendar_dict["file_uploader"], import_file)
    Waits.Wait10s_ElementLoaded(calendar_dict["import_table"])
    Logging("Calendar Import - Wait until file is loaded successfully")
    
    time.sleep(1)

    Commands.ClickElement(calendar_dict["save_import"])
    Waits.WaitUntilPageIsLoaded(None)
    Waits.WaitElementLoaded(30, calendar_dict["weekday_grid"])

    time.sleep(2)
    try:
        event_title2 = Functions.GetListLength(calendar_dict["event"])
        print("event_title2: " + str(event_title2))

        if event_title2 > event_title1:
            TestCase_LogResult(**calendar_tc["import"]["pass"])
        else:
            TestCase_LogResult(**calendar_tc["import"]["fail"])
    except WebDriverException:
        TCResult_ValidateAlertMsg(menu="calendar_folder", testcase="import", msg="click save import")
        TestCase_LogResult(**calendar_tc["import"]["fail"])

    return [fullday_title, clocked_title]

def CalendarFolder_CopyEvent(event_title):
    PrintYellow("[MENU CALENDAR] COPY EVENT")
    events_before = Functions.GetListLength("//span[contains(text(),'" + event_title + "')]")
    Logging("events_before: " + str(events_before))

    page_title = Functions.GetElementText(calendar_dict["page_header"])
    print("page_title_text: " + page_title)

    Commands.ClickElement(calendar_dict["event_span"] % event_title)
    Logging("Calendar Folder - View - Open View mode")

    Waits.Wait10s_ElementLoaded(calendar_dict["view_container"])
    Logging("Wait until view mode appears")

    Commands.ClickElement(calendar_dict["copy_button"])
    Logging("Click Copy button")

    Commands.Wait10s_ClickElement(calendar_dict["mycal_expander"])
    Logging("Select My Calendar")
    
    Commands.Wait10s_ClickElement(calendar_dict["mycal"] % page_title)
    Logging("Select My Calendar folder")

    time.sleep(1)

    Commands.ClickElement(calendar_dict["copy_save"])
    Logging("Click Save button")

    try:
        Waits.Wait10s_ElementLoaded(calendar_dict["alert_popup"])
        Logging("Wait until msg pop-up appears")

        alert_msg = Functions.GetElementText(calendar_dict["alert_popup"])
        Logging("alert_msg: " + str(alert_msg))
        if alert_msg == "Success in copying":
            Logging("Copy msg is displayed correctly")
        else:
            Logging("Copy msg is displayed incorrectly")

        Commands.ClickElement(calendar_dict["alert_close"])
        Logging("Close msg pop-up")

        Waits.Wait10s_ElementLoaded("//span[contains(.,'" + event_title + "')]")
        
        Commands.ClickElement(calendar_dict["reload_button"])
        Logging("Click reload button")
        time.sleep(1)
        Waits.WaitUntilPageIsLoaded(None)

        events_after = Functions.GetListLength("//span[contains(.,'" + event_title + "')]")
        Logging("events_after: " + str(events_after))
        
        if events_after > events_before:
            TestCase_LogResult(**calendar_tc["copy"]["pass"])
            copy_result = True
        else:
            TestCase_LogResult(**calendar_tc["copy"]["fail"])
            copy_result = False
    except WebDriverException:
        copy_result = False
        TCResult_ValidateAlertMsg(menu="calendar_folder", testcase="copy", msg="click copy event")
        TestCase_LogResult(**calendar_tc["copy"]["fail"])

    return copy_result

def CalendarFolder_DeleteEvents(event_title):
    PrintYellow("[MENU CALENDAR] DELETE EVENT")
    
    Waits.WaitUntilPageIsLoaded(None)
    time.sleep(1)

    Waits.Wait10s_ElementLoaded("//span[contains(.,'" + event_title + "')]")
    selected_event = Commands.FindElements("//span[contains(.,'" + event_title + "')]")
    events_before_delete = int(len(selected_event))
    Logging("Event name: " + event_title)
    Logging("Total of events before delete: " + str(events_before_delete))
    time.sleep(1)
    selected_event[0].click()
    Logging("View - Open View mode")

    Waits.Wait10s_ElementLoaded(calendar_dict["view_container"])

    Commands.ClickElement(calendar_dict["delete_button"])
    print("Click Delete button")

    Waits.Wait10s_ElementLoaded(calendar_dict["warning_modal"])

    try:
        Commands.ClickElement(calendar_dict["delete_all"])
        print("Delete all repeated events")
    except WebDriverException:
        print("Delete single event")

    Commands.ClickElement(calendar_dict["ok_button"])
    print("Confirm Delete events")

    time.sleep(1)
    Waits.WaitUntilPageIsLoaded(None)
    time.sleep(1)

    try:
        events_after_delete = Functions.GetListLength("//span[contains(text(),'" + event_title + "')]")
    except WebDriverException:
        events_after_delete = 0
    finally:
        Logging("Total events after delete: " + str(events_after_delete))
    
    if events_after_delete < events_before_delete:
        TestCase_LogResult(**calendar_tc["delete"]["pass"])
    else:
        TestCase_LogResult(**calendar_tc["delete"]["fail"])

def CalendarFolder_SearchDetails(event_name):
    page_header = Functions.GetElementText(calendar_dict["page_header"])
    if "My Calendar" not in page_header:
        mycalendar_ul = Functions.GetElementAttribute(calendar_dict["mycalendar_ul"], "style")
        if "display: none;" in mycalendar_ul:
            Commands.ClickElement(calendar_dict["mycalendar_b"])
            Logging("Open view calendar sub folders")
            Waits.Wait10s_ElementLoaded(calendar_dict["mycalendar_flagged"])
        
        Commands.ClickElement(calendar_dict["mycalendar_span"])
        Logging("Access My Calendar sub folder")
    
    Waits.WaitElementLoaded(30, calendar_dict["weekday_grid"])
    Waits.WaitUntilPageIsLoaded(None)
    
    Commands.ClickElement(calendar_dict["event_item0"] % event_name)
    Logging("Click view event")

    Waits.Wait10s_ElementLoaded(calendar_dict["calendar_event"])

    event_data = {
        "location": calendar_dict["view_location"],
        "creator": calendar_dict["view_creator"],
        "priority": calendar_dict["view_priority"]
    }
    
    date = Functions.GetElementText(calendar_dict["view_date"])
    if "~" in date:
        event_month = date.split(" ~ ")[0].split("/")[1]
    else:
        event_month = date.split("/")[1]

    for event_label in event_data.keys():
        xpath = event_data[event_label]
        try:
            text = Functions.GetElementText(xpath)
            if event_label == "creator":
                ele_text = text.split(" /")[0]
            else:
                ele_text = text
            Logging(" -> " + event_label + " = " + ele_text)
            event_data[event_label] = ele_text
        except WebDriverException:
            pass

    Commands.ClickElement(calendar_dict["view_details_b"])
    Logging("Access details page")
    Waits.Wait10s_ElementLoaded(calendar_dict["details_title"])

    event_title = Functions.GetElementText(calendar_dict["details_title"])
    event_data.update({"title": event_title})
    Logging("Update event name to search data dict")
    
    if event_title in event_data.values():
        Logging("Update event name successfully " + objects.testcase_pass)
    else:
        Logging("Fail to update event name " + objects.testcase_fail)

    time.sleep(1)

    Commands.ClickElement(calendar_dict["details_back_b"])
    Logging("Back to calendar view from details page")
    Waits.WaitUntilPageIsLoaded(calendar_dict["event_item"])

    monthly_title = Functions.GetElementText(calendar_dict["month_header"]).split("/")[1]
    if int(monthly_title) > int(event_month):
        Commands.ClickElement(calendar_dict["move_prev"])
        Logging("Move to previous month")
        move_calendar = True
    elif int(monthly_title) < int(event_month):
        Commands.ClickElement(calendar_dict["move_next"])
        Logging("Move to next month")
        move_calendar = True
    else:
        move_calendar = False
    
    if move_calendar == True:
        i=0
        for i in range(0,10):
            i+=1
            time.sleep(1)
            new_month = Functions.GetElementText(calendar_dict["month_header"])
            if new_month != monthly_title:
                move_calendar = False
                break
            else:
                move_calendar = None
        
    if move_calendar == False:
        time.sleep(1)
        
        Commands.ClickElement(calendar_dict["search_dropdown"])
        Logging("Open search dropdown")

        Waits.Wait10s_ElementLoaded(calendar_dict["search_title"])

        input_xpath = {
            "title": calendar_dict["search_title"],
            "location": calendar_dict["search_location"],
            "priority": calendar_dict["search_priority"],
            "creator": calendar_dict["search_creator"]
        }
        
        time.sleep(1)

        for search_data in input_xpath.keys():
            xpath = input_xpath[search_data]
            value = event_data[search_data]
            if search_data == "priority":
                Commands.Selectbox_ByVisibleText(xpath, value)
                Logging("Select priority by text " + value)
            else:
                Commands.InputElement(xpath, value)
                Logging("Send search key of " + search_data)

        try:
            Commands.ClickElement(calendar_dict["search_button"])
            Logging("Click Search button")
        except WebDriverException:
            pass

        Waits.Wait10s_ElementLoaded(calendar_dict["search_result"])

        Waits.Wait10s_ElementLoaded("//*[contains(., '%s')]" % event_name)
        Logging("Search event successfully" + objects.testcase_pass)

        Commands.ClickElement(calendar_dict["search_dropdown"])
        Logging("Open search dropdown")

        Waits.Wait10s_ElementLoaded(calendar_dict["reset_button"])
        try:
            Commands.ClickElement(calendar_dict["reset_button"])
            Logging("Click Reset search")
        except WebDriverException:
            pass

        Waits.Wait10s_ElementLoaded(calendar_dict["event_item"])
        Logging("Reset search successfully")
        search_result = True  
    else:
        search_result = False
    
    try:
        #Commands.FindElement("//div[@class='input-icon open']")
        Commands.ClickElement(calendar_dict["close_search"])
        Logging("Close search box")
        Waits.Wait10s_ElementInvisibility(calendar_dict["search_title"])
    except WebDriverException:
        pass
    
    return search_result

def CalendarExecution():
    access_menu = Calendar_AccessCalendarMenu()
    if access_menu == True:
        fullday_event = Calendar_AddEvent(time_event = False, event_public = False, repeated = False)
        Calendar_ViewEvent(event_title = fullday_event["title"], repeated = False, start_date = fullday_event["start_date"], end_date = fullday_event["end_date"])
        copy_event = CalendarFolder_CopyEvent(fullday_event["title"])
        Calendar_EditEvent(fullday_event["title"])
        CalendarFolder_DeleteEvents(fullday_event["title"])
        repeated_event = Calendar_AddEvent(time_event = True, event_public = True, repeated = True)
        import_events = Calendar_ImportEvent()
        Calendar_AccessPublicCalendar()
        Calendar_ViewEvent(event_title = repeated_event["title"], repeated = True, start_date = repeated_event["start_date"], end_date = repeated_event["end_date"])
        Calendar_AccessMyCalendar()
        CalendarFolder_DeleteEvents(repeated_event["title"])
        CalendarFolder_DeleteEvents(import_events[0])
        CalendarFolder_DeleteEvents(import_events[1])
        CalendarFolder_SearchDetails(fullday_event["title"])
        ValidateUnexpectedModal()
       