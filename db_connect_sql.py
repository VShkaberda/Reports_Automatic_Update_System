# -*- coding: utf-8 -*-
"""
Created on Tue Jan 23 23:23:01 2018
@author: Vadim Shkaberda
"""
import pyodbc

class DBConnect(object):
    ''' Provides connection to database and functions to work with server.
    '''
    def __enter__(self):
        # Connection properties
        conn_str = (
            r'Driver={SQL Server};'
            r'Server=s-kv-center-s64;'
            r'Database=SILPOAnalitic;'
            r'Trusted_Connection=yes;'
            )
        self.__db = pyodbc.connect(conn_str)
        self.__cursor = self.__db.cursor()
        return self

    def __exit__(self, type, value, traceback):
        self.__db.close()

    def file_to_update(self):
        ''' Fetching one file to be updated next.
        '''
        self.__cursor.execute('''SELECT top 1 UploadFileName, FirstResourceLink, ReportID, ReportName,
            Notifications, Attachments, NotificationsWhom, NotificationsCopy, Notificationstext, SecondResourceLink
                  FROM [SILPOAnalitic].[dbo].[Hermes_Reports]
                  where 1=1
                	and
                	((datepart(weekday,ExecutedJob) = iif(DAY1=1, 1, 0) AND
                    datepart(weekday,ExecutedJob) =  datepart(weekday, getdate()))
                	OR
                	(datepart(weekday,ExecutedJob) = iif(DAY2=1, 2, 0) AND
                    datepart(weekday,ExecutedJob) =  datepart(weekday, getdate()))
                	OR
                	(datepart(weekday,ExecutedJob) = iif(DAY3=1, 3, 0) AND
                    datepart(weekday,ExecutedJob) =  datepart(weekday, getdate()))
                	OR
                	(datepart(weekday,ExecutedJob) = iif(DAY4=1, 4, 0) AND
                    datepart(weekday,ExecutedJob) =  datepart(weekday, getdate()))
                	OR
                	(datepart(weekday,ExecutedJob) = iif(DAY5=1, 5, 0) AND
                    datepart(weekday,ExecutedJob) =  datepart(weekday, getdate()))
                	OR
                	(datepart(weekday,ExecutedJob) = iif(DAY6=1, 6, 0) AND
                    datepart(weekday,ExecutedJob) =  datepart(weekday, getdate()))
                	OR
                	(datepart(weekday,ExecutedJob) = iif(DAY7=1, 7, 0) AND
                    datepart(weekday,ExecutedJob) =  datepart(weekday, getdate())))
                	AND
                	ExecutedJob > ISNULL(LastDateUpdate, 0)
                	and
                	StatusID = 1 --рабочий
                	and
                	ScheduleTypeID = 0 --прямой график
                	and
                	convert(time,getdate()) >= isnull(timefrom,'00:00')  -- обновлять позже назначенного
                and isnull(Error, 0) = 0
                order by [priority]''')
        return self.__cursor.fetchone()

    def successful_update(self, rID, update_time):
        ''' Update data on server that file update was succeeded.
        '''
        print('Update succeeded! Loading data to server.')
        self.__cursor.execute('UPDATE [SILPOAnalitic].[dbo].[Hermes_Reports] \
                              SET LastDateUpdate = cast(? as datetime) \
                              WHERE ReportID = ?', (update_time, rID))
        self.__db.commit()

    def failed_update(self, rID, update_time, update_error):
        ''' Update data on server in case if update was failed.
        '''
        print('Update failed!')
        self.__cursor.execute('UPDATE [SILPOAnalitic].[dbo].[Hermes_Reports] \
                              SET Error = ?, LastDateUpdate = cast(? as datetime) \
                              WHERE ReportID = ?', (update_error, update_time, rID))
        self.__db.commit()
