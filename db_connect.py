# -*- coding: utf-8 -*-
"""
Created on Sun Jan 14 13:27:45 2018
@author: Vadim Shkaberda
"""
import sqlite3


class DBConnect(object):
    ''' Provides connection to database and functions to work with server.
    '''
    def __enter__(self):
        self.__db = sqlite3.connect('reports.sqlite')
        self.__cursor = self.__db.cursor()
        return self
    
    def __exit__(self, type, value, traceback):
        self.__db.close()
    
    def file_to_update(self):
        ''' Fetching one file to be updated next.
        '''
        self.__cursor.execute('''SELECT report_name, refresh, done, priority
                              FROM Reports
                              WHERE refresh = 1 and Done = 0 and Error is NULL
                              ORDER BY priority ASC, report_name
                              LIMIT 1''')
        return self.__cursor.fetchone()
    
    def successful_update(self, f, update_dur):
        ''' Update data on server that file update was succeeded.
        '''
        self.__cursor.execute('UPDATE Reports \
                              SET Done = 1, updatetime = ? \
                              WHERE report_name = ?', (update_dur, f))
        self.__db.commit()
    
    def failed_update(self, f, update_dur):
        ''' Update data on server in case if update was failed.
        '''
        self.__cursor.execute('UPDATE Reports \
                              SET Error = 1, updatetime = ? \
                              WHERE report_name = ?', (update_dur, f))
        self.__db.commit()