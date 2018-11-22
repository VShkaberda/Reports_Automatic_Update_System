# -*- coding: utf-8 -*-
"""
Created on Sun Jan  7 15:53:26 2018
@author: Vadim Shkaberda
"""

from db_connect_sql import DBConnect
from log_error import writelog
from os import path
from pyodbc import Error as SQLError
from send_mail import send_mail
from xl import copy_file, update_file

import sharepoint
import sys
import threading
import time


class Main(object):
    def __init__(self, fileinfo):
        # Info about last file
        self.fileinfo = fileinfo
        self.sleep_duration = 30 # if no files to update
        self.errors = {} # error description
        # keys for SQL parsing
        self.parse_keys = ('fname', 'fpath', 'reportID', 'reportName', 'Notifications', 'Attachments',
                'NotificationsWhom', 'NotificationsCopy', 'Notificationstext', 'SecondResourceLink',
                'GroupName')


    def db_update(self, dbconn):
        ''' Function loads info about file update to db.
        '''
        if self.fileinfo['update_error'] == 0:
            dbconn.successful_update(self.fileinfo['reportID'],
                                     self.fileinfo['update_time'])
        else:
            # 6 - problem with Outlook. Send mail using SQL Server.
            if self.fileinfo['update_error'] == 6:
                dbconn.send_emergency_mail(self.fileinfo['reportName'])
            else:
                send_mail(subject='(Ошибка обновления) ' + self.fileinfo['reportName'],
                      HTMLBody=('ID ошибки: ' + str(self.fileinfo['update_error']) +
                    '. ' + self.errors[self.fileinfo['update_error']] + '<br>' +
                    'Отчёт: ID ' + str(self.fileinfo['reportID']) + ' <a href="' +
                    path.join(self.fileinfo['fpath'], self.fileinfo['fname']) + '">' +
                    self.fileinfo['reportName'] + '</a>.'), rName=self.fileinfo['reportName'])
            dbconn.failed_update(self.fileinfo['reportID'],
                                 self.fileinfo['update_time'],
                                 self.fileinfo['update_error'])
        # clear info about last file
        for key in self.fileinfo.keys():
            self.fileinfo[key] = None


    def email_gen(self, dbconn=None):
        ''' Returns a dictionary with keywords for mail.
            "dbconn" parameter signalized that it is a group mail.
        '''
        # choose Group- or reportName and create path(-es) to attachment(-s)
        if dbconn:
            subj = self.fileinfo['GroupName']
            att = []
            for group_att in dbconn.group_attachments(self.fileinfo['GroupName']):
                att.append(path.join(group_att[0], group_att[1]))
        else:
            subj = self.fileinfo['reportName']
            att = [path.join(self.fileinfo['fpath'], self.fileinfo['fname'])] \
                        if self.fileinfo['Attachments'] else None
        return {'to': self.fileinfo['NotificationsWhom'],
                'copy': self.fileinfo['NotificationsCopy'],
                'subject': '(Автоотчет) ' +  subj + \
                ' (' + self.fileinfo['update_time'][:10] + ')', # date in brackets
                'HTMLBody': self.fileinfo['Notificationstext'],
                'att': att,
                'rName': self.fileinfo['reportName']
                }


    def parse_SQL(self, file_sql):
        ''' Turns result from SQL query into a dictionary.
        '''        
        for i in range(11):
            self.fileinfo[self.parse_keys[i]] = file_sql[i]


    def time_to_sleep(self):
        ''' Sleep before next cycle.
        '''
        now = time.localtime()
        if now.tm_hour >= 20 or now.tm_hour < 6:
            print('{}. No files to update.'
                  .format(time.strftime("%d-%m-%Y %H:%M:%S", now)))
            time.sleep(3600)
            return
        print('{}. No files to update. Waiting {} seconds.'
              .format(time.strftime("%d-%m-%Y %H:%M:%S", now), self.sleep_duration))
        time.sleep(self.sleep_duration)
        if self.sleep_duration < 450:
            self.sleep_duration *= 2

    ##### Working cycles. #####

    def run(self):
        ''' Init cycle. Connects to database. Writes info about the last file
            to db, if the error has occured after file update.
        '''
        with DBConnect() as dbconn:
            # download error description from db
            for err in dbconn.error_description():
                self.errors[err[0]] = err[1]
            # if info wasn't written to db after last file update
            if self.fileinfo['fname']:
                self.db_update(dbconn)
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
            self.parse_SQL(file_sql)
            # Calling function to work with Excel
            print('{}'.format(time.strftime("%d-%m-%Y %H:%M:%S", time.localtime())), end=' ')
            self.fileinfo['update_error'] = update_file(self.fileinfo['fpath'],
                                                         self.fileinfo['fname'])
            self.fileinfo['update_time'] = time.strftime("%d-%m-%Y %H:%M:%S",
                                                         time.localtime())
            # Copy file
            if self.fileinfo['update_error'] == 0 and self.fileinfo['SecondResourceLink']:
                self.fileinfo['update_error'] = copy_file(self.fileinfo['fpath'],
                                                          self.fileinfo['fname'],
                                                          self.fileinfo['SecondResourceLink'])
            # Send mail
            if self.fileinfo['update_error'] == 0 and self.fileinfo['Notifications'] == 1:
                # if we have no group - send mail
                if not self.fileinfo['GroupName']:
                    self.fileinfo['update_error'] = send_mail(**self.email_gen())
                # if we have GroupName - send mail if group_mail_check == 1
                elif dbconn.group_mail_check(self.fileinfo['GroupName']):
                    self.fileinfo['update_error'] = send_mail(**self.email_gen(dbconn))
            # Write in the db result of update and send mail in a case of a failure
            self.db_update(dbconn)
            time.sleep(3)
            self.sleep_duration = 30
        print('Exiting main cycle...')


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
                'reportName': None,
                'update_error': None,
                'update_time': None,
                'Notifications': None,
                'Attachments': None,
                'NotificationsWhom': None,
                'NotificationsCopy': None,
                'Notificationstext': None,
                'SecondResourceLink': None,
                'GroupName': None}
    main = Main(FileInfo)

    sharepoint.sharepoint_check() # check connection to sharepoint

    while connection_retry[0] < 10 and thread.is_alive():
        try:
            main.run()
        except SQLError as e:
            writelog(e)
            # reset connection retries counter after 1 hour
            if time.time() - connection_retry[1] > 3600:
                connection_retry = [0, time.time()]
            print(e)
            connection_retry[0] += 1
            # magnify sleep interval in case of repeatable connect failure
            t_sleep = 900 if connection_retry[0] == 9 else 20
            print('Retrying to connect in {} seconds. \
              Number of retries since the first fixed dicsonnect in 24 hours: {}'
              .format(t_sleep, connection_retry[0]))
            time.sleep(t_sleep)
        # in case of unexpected error - try to send email from SQL Server
        except Exception as e:
            writelog(e)
            try:
                with DBConnect() as dbconn:
                    dbconn.send_crash_mail()
            finally:
                sys.exit()

    print('Exiting program.')
