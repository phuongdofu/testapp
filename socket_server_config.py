import socket, shutil, os, json, ast, time
from openpyxl import load_workbook, Workbook

# [Reference] https://www.digitalocean.com/community/tutorials/python-socket-programming-server-client
# [Reference] https://viblo.asia/p/lap-trinh-socket-bang-python-jvEla084Zkw

''' This file will be run at server '''

def get_maximum_rows(*, sheet_object):
    rows = 0
    for max_row, row in enumerate(sheet_object, 1):
        if not all(col.value is None for col in row):
            rows += 1
    return rows

def ReceiveInput():
    # get the host name
    host = socket.gethostname()
    port = 5000  # initiate port no above 1024

    server_socket = socket.socket()  # get instance
    # look closely. The bind() function takes tuple as argument
    server_socket.bind((host, port))  # bind host address and port together

    # configure how many client the server can listen simultaneously
    server_socket.listen(2)
    conn, address = server_socket.accept()  # accept new connection
    print("Connection from: " + str(address))

    while True:
        try:
            data = conn.recv(10485760).decode('utf-8')
            parsed_data = ast.literal_eval(str(data).replace("'", '"'))

            testcase_dict = dict(parsed_data)
            testplan_name = testcase_dict["1"]["tester"]
            section_id = testcase_dict["1"]["section_id"]
            testcase_filename = "%s_result_%s" % (testplan_name, section_id)

            current_path = os.path.dirname(os.path.realpath(__file__))
            if "\\" in current_path:
                file_name = current_path + '\\Log\\Test Log\\%s.xlsx' % testcase_filename
            else:
                file_name = current_path + '/Log/Test Log/%s.xlsx' % testcase_filename

            create_wb = Workbook()
            create_wb.save(file_name)
            
            time.sleep(1)
            
            wb = load_workbook(file_name)
            ws = wb.active

            last_row = 88
            for row_number in range(1, last_row):
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
                    ws.cell(row=row_number, column=2).value = testcase_dict[tc_id]["menu"]
                    ws.cell(row=row_number, column=3).value = testcase_dict[tc_id]["submenu"]
                    ws.cell(row=row_number, column=4).value = testcase_dict[tc_id]["testcase"]
                    ws.cell(row=row_number, column=5).value = testcase_dict[tc_id]["status"]
                    ws.cell(row=row_number, column=6).value = testcase_dict[tc_id]["date"]
                    ws.cell(row=row_number, column=7).value = testcase_dict[tc_id]["tester"]

                row_number+=1

            wb.save(file_name)
        except:
            pass


if __name__ == '__main__':
    ReceiveInput()