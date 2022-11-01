import socket, shutil, os

# [Reference] https://www.digitalocean.com/community/tutorials/python-socket-programming-server-client
# [Reference] https://viblo.asia/p/lap-trinh-socket-bang-python-jvEla084Zkw

''' This file will be run at server '''

def ReceiveInput():
    # get the hostname
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
        data = conn.recv(1024).decode()
        if str(data).endswith(".xlsx"):
            print("client send test case file result")
            # server: test case file name include the test plan name
            #         server recognize test plan by file name
            #         if test plan does not exist (folder): create new folder
            # client: submit test plan -> save test plan name in txt file
            #         test case file name is named by test plan name + date (id)

            file_name = str(data)
            print("file_name: " + str(file_name))

            if "\\Log\\Test Log\\" in file_name:
                testplan_name = file_name.split("\\Log\\Test Log\\")[1].split("_result_")[0]
                destination_path = os.path.dirname(os.path.realpath(__file__)) + '\\Log\\Test Log\\%s' % testplan_name
            else:
                testplan_name = file_name.split("/Log/Test Log/")[1].split("_result_")[0]
                destination_path = os.path.dirname(os.path.realpath(__file__)) + '/Log/Test Log/%s' % testplan_name
             
            shutil.copy(data, destination_path)

if __name__ == '__main__':
    ReceiveInput()