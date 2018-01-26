# -*- coding: utf-8 -*-
"""
Created on Sun Jan  7 15:53:26 2018
@author: Vadim Shkaberda
"""

from db_connect_sql import DBConnect
from os import path, getcwd
from pyodbc import Error as SQLError
from xl import update_file

import threading
import time

        
def control_thread():
    ''' Function that monitors user input.
        If 1 have been inputed:
        function is exiting, causing interruption of main()'''
    while True:
        output = input("Press 1 if you want to interrupt programm.\n")
        if output == '1':
            print('Programm is interrupted.')
            break


def main():
    with DBConnect() as dbconn:
        while thread.is_alive():
        #while True:
            filedata = dbconn.file_to_update()
            # if no files to update
            if filedata is None:
                print('No files to update. Waiting 30 seconds.')
                #break
                time.sleep(30)
                continue
            filename = filedata[0]
            root = filedata[1]
            reportID = filedata[2]
            #start_update = time.time()
            successful_update = update_file('\\' + root, filename)
            #update_duration = time.time() - start_update
            update_time = time.strftime("%d-%m-%Y %H:%M:%S", time.localtime())
            # Write in the db result of update
            if successful_update:
                dbconn.successful_update(reportID, update_time)
            else:
                dbconn.failed_update(reportID, update_time)
            time.sleep(5)


if __name__ == "__main__":
    #root = path.join(getcwd(), 'XL')
    
    # Start thread that monitors user input
    thread = threading.Thread(target=control_thread)
    # Python program exits when only daemon threads are left
    thread.daemon = True
    thread.start()
    
    connection_retry = 0
    while connection_retry < 3:
        try:
            main()        
        except SQLError as e:
            print(e)
            print('Retrying to connect in 5 minutes. \
                  Number of retries:', connection_retry)
            connection_retry += 1
            time.sleep(5)