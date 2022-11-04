import socket
from common_functions import Files

# [Reference] https://www.digitalocean.com/community/tutorials/python-socket-programming-server-client
# [Reference] https://viblo.asia/p/lap-trinh-socket-bang-python-jvEla084Zkw

''' This file will run at client '''

def SendTestCaseFile():
    host = socket.gethostname()
    port = 5000 

    client_socket = socket.socket()  # instantiate
    client_socket.connect((host, port))  # connect to the server

    testcase_file = Files.testcase_file
    
    client_socket.send(testcase_file.encode())

    client_socket.close()

# if __name__ == '__main__':
#     SendTestCaseFile()