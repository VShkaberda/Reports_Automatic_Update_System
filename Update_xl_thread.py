# -*- coding: utf-8 -*-
"""
Created on Sun Jan  7 15:53:26 2018
@author: Vadim Shkaberda
"""

from db_connect_sql import DBConnect
from pyodbc import Error as SQLError
from xl import update_file

import datetime
import threading
import time


class Main(object):
    def __init__(self, fileinfo):
        # Info about last file
        self.fileinfo = fileinfo
        self.sleep_duration = 30 # if no files to update


    def run(self):
        ''' Init cycle. Connects to database. Writes info about the last file
            to db, if the error has occured after file update.
        '''
        with DBConnect() as dbconn:
            # if info wasn't written to db after last file update
            if self.fileinfo['fname']:
                self.db_update()
            # run main cycle
            self.main_cycle(dbconn)
        print('Exiting run...')


    def main_cycle(self, dbconn):
        ''' Main cycle. Gets info about file, calls update_file function
            (that runs macro "Update" in Excel file) and saving result to db.
        '''
        # check if thread was user interrupted
        while thread.is_alive():
            # get file
            file_sql = dbconn.file_to_update()
            # if no files to update
            if file_sql is None:
                self.time_to_sleep()
                continue
            self.fileinfo['fname'] = file_sql[0]
            self.fileinfo['fpath'] = '\\' + file_sql[1]
            self.fileinfo['reportID'] = file_sql[2]
            # Calling function to work with Excel
            self.fileinfo['update_state'] = update_file(self.fileinfo['fpath'],
                                                         self.fileinfo['fname'])
            self.fileinfo['update_time'] = time.strftime("%d-%m-%Y %H:%M:%S",
                                                         time.localtime())
            # Write in the db result of update
            self.db_update(dbconn)
            time.sleep(5)
            self.sleep_duration = 30
        print('Exiting main cycle...')


    def db_update(self, dbconn):
        ''' Function loads info about file update to db.
        '''
        if self.fileinfo['update_state']:
            dbconn.successful_update(self.fileinfo['reportID'],
                                     self.fileinfo['update_time'])
        else:
            dbconn.failed_update(self.fileinfo['reportID'],
                                 self.fileinfo['update_time'])
        # clear info about last file
        for key in self.fileinfo.keys():
            self.fileinfo[key] = None


    def time_to_sleep(self):
        ''' Sleep before next cycle.
        '''
        now = datetime.datetime.today()
        if now.hour >= 19:  # sleep till 5 AM
            wakeup = datetime.datetime(now.year, now.month, now.day, 5, 0)
            wakeup += datetime.timedelta(days=1)
            print('Sleep till 5 AM.')
            time.sleep((wakeup-now).seconds)
            return
        print('No files to update. Waiting {} seconds.'.format(self.sleep_duration))
        time.sleep(self.sleep_duration)
        if self.sleep_duration < 900:
            self.sleep_duration *= 2


def control_thread():
    ''' Function that monitors user input.
        If 1 have been inputed:
        function is exiting, causing interruption of main()'''
    while True:
        output = input("Press 1 if you want to interrupt programm.\n")
        if output == '1':
            print('Programm is interrupted.')
            break
    print('Exiting daemon...')


if __name__ == "__main__":
    # Start thread that monitors user input
    thread = threading.Thread(target=control_thread)
    # Python program exits when only daemon threads are left
    thread.daemon = True
    thread.start()

    connection_retry = [0, time.time()]

    # Initialize with no info about file
    FileInfo = {'fname': None,
                'fpath': None,
                'reportID': None,
                'update_state': None,
                'update_time': None}
    main = Main(FileInfo)
    while connection_retry[0] < 3 and thread.is_alive():
        try:
            main.run()
        except SQLError as e:
            # reset connection retries counter after 24 hours
            if time.time() - connection_retry[1] > 86400:
                connection_retry[0] = 0
            print(e)
            print('Retrying to connect in 5 minutes. \
                  Number of retries:', connection_retry[0])
            connection_retry[0] += 1
            connection_retry[1] = time.time()
            time.sleep(300)
    print('Exiting program.')
