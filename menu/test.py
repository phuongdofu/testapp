'''
    API
        1. read excel file
        2. data format
            data = {
                "1": {
                    "no": "col1",
                    "menu": "col2",
                    "submenu": "col3",
                    "testcase": "col4",
                    "status": "col5",
                    "date": "col6",
                    "tester": "col7"
                },
                "2": {
                    "no": "col1",
                    "menu": "col2",
                    "submenu": "col3",
                    "testcase": "col4",
                    "status": "col5",
                    "date": "col6",
                    "tester": "col7"
                },
            }
        3. send data to server
        4. server read data (json.loads(json_file)) 
        5. server convert to excel file
'''

from openpyxl import load_workbook

api_data = {}

file = "D:\\PhuongDofu\\groupware-auto-test-2\\Attachment\\testcase_log_result.xlsx"
wb = load_workbook(file)
ws = wb.active

last_row = 101
for row_number in range(1,last_row):

    row_number+=1
    no = str(ws.cell(row=row_number, column=1).value)
    menu = ws.cell(row=row_number, column=2).value
    submenu = ws.cell(row=row_number, column=3).value
    testcase = ws.cell(row=row_number, column=4).value
    status = ws.cell(row=row_number, column=5).value
    date = ws.cell(row=row_number, column=6).value
    tester = ws.cell(row=row_number, column=7).value

    api_data[no] = {
        "no": no,
        "menu": menu,
        "submenu": submenu,
        "testcase": testcase,
        "status": status,
        "date": date,
        "tester": tester
    }

wb.save(file)

# Convert data to excel file
server_data_file = "D:\\PhuongDofu\\groupware-auto-test-2\\Attachment\\testcase.xlsx"
wb = load_workbook(file)
ws = wb.active

last_row = 101
for row_number in range(1,last_row):
    if row_number == 1:
        ws.cell(row=row_number, column=1).value = "No"
        ws.cell(row=row_number, column=2).value = "Menu"
        ws.cell(row=row_number, column=3).value = "Submenu"
        ws.cell(row=row_number, column=4).value = "Test Case"
        ws.cell(row=row_number, column=5).value = "Status"
        ws.cell(row=row_number, column=6).value = "Date"
        ws.cell(row=row_number, column=7).value = "Tester"
    else:
        tc_id = str(row_number)
        ws.cell(row=row_number, column=1).value = tc_id
        ws.cell(row=row_number, column=2).value = api_data[tc_id]["menu"]
        ws.cell(row=row_number, column=3).value = api_data[tc_id]["submenu"]
        ws.cell(row=row_number, column=4).value = api_data[tc_id]["testcase"]
        ws.cell(row=row_number, column=5).value = api_data[tc_id]["status"]
        ws.cell(row=row_number, column=6).value = api_data[tc_id]["date"]
        ws.cell(row=row_number, column=7).value = api_data[tc_id]["tester"]

    row_number+=1

wb.save(server_data_file)
 