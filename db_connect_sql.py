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


    def error_description(self):
        ''' Generator of error description from db.
        '''
        self.__cursor.execute('SELECT [Error], [Error_name] \
            FROM [SILPOAnalitic].[dbo].[hermes_reports_error]')
        for row in self.__cursor.fetchall():
            if not row:
                break
            yield row


    def file_to_update(self):
        ''' Fetching one file to be updated next.
        '''
        self.__cursor.execute('''SELECT top 1 UploadFileName, FirstResourceLink, ReportID, ReportName,
            Notifications, Attachments, NotificationsWhom, NotificationsCopy, Notificationstext,
            SecondResourceLink, GroupName
                FROM [SILPOAnalitic].[dbo].[Hermes_Reports]
                WHERE (
                        ((ScheduleTypeID = 2 -- monthly
                            OR
                                (ScheduleTypeID = 0 -- direct schedule
                                AND -- instead of bit mask
                                    (
                                 (iif(DAY1=1, '1', '') +
                                 iif(DAY2=1, '2', '') +
                                 iif(DAY3=1, '3', '') +
                                 iif(DAY4=1, '4', '') +
                                 iif(DAY5=1, '5', '') +
                                 iif(DAY6=1, '6', '') +
                                 iif(DAY7=1, '7', '')) like '%' + cast(datepart(weekday, ExecutedJob) as char(1)) + '%'
                                    )
                                )
                            )
                            AND isnull(Error, 0) = 0 -- don't try to refresh report with error
                            AND ExecutedJob > ISNULL(LastDateUpdate, 0) -- didn't refresh after last job
                        )
                        OR
                        (
                            ScheduleTypeID = 5 -- refresh on time
                            AND -- instead of bit mask
                                (
                                (iif(DAY1=1, '1', '') +
                                iif(DAY2=1, '2', '') +
                                iif(DAY3=1, '3', '') +
                                iif(DAY4=1, '4', '') +
                                iif(DAY5=1, '5', '') +
                                iif(DAY6=1, '6', '') +
                                iif(DAY7=1, '7', '')) like '%' + cast(datepart(weekday, getdate()) as char(1)) + '%'
                                )
                            -- didn't refresh today
                            AND DATEDIFF(DAY, 0, GETDATE()) + cast(TimeFrom as datetime) > ISNULL(LastDateUpdate, 0)
                        )
                    )
                    AND StatusID = 1 -- working
                    AND convert(time, getdate()) >= isnull(timefrom, '00:00') -- refresh later timefrom
                ORDER BY [priority]''')
        return self.__cursor.fetchone()


    def group_attachments(self, groupname):
        ''' Generator of attachments for email after updating all of group reports.
        '''
        self.__cursor.execute('''SELECT FirstResourceLink, UploadFileName
                  FROM [SILPOAnalitic].[dbo].[Hermes_Reports]
                  where (ScheduleTypeID = 2 -- monthly
                         OR
                        (ScheduleTypeID = 0 -- direct schedule
                             AND
                             -- instead of bit mask
                            (
                             (iif(DAY1=1, '1', '') +
                             iif(DAY2=1, '2', '') +
                             iif(DAY3=1, '3', '') +
                             iif(DAY4=1, '4', '') +
                             iif(DAY5=1, '5', '') +
                             iif(DAY6=1, '6', '') +
                             iif(DAY7=1, '7', '')) like '%' + cast(datepart(weekday, ExecutedJob) as char(1)) + '%'
                            )
                        )
                    )
                	AND StatusID = 1
                	AND convert(time,getdate()) >= isnull(timefrom,'00:00')
                 AND Attachments = 1
                 AND GroupName = ?
                ''', (groupname))
        for row in self.__cursor.fetchall():
            if not row:
                break
            yield row


    def group_mail_check(self, groupname):
        ''' Return 1 if all files from group have been updated.
        '''
        self.__cursor.execute('''SELECT count(*) as Group_count,
                 sum(case when Error = 0 AND LastDateUpdate > ExecutedJob
                     then 1 else 0 end) + 1 as Updated
                  FROM [SILPOAnalitic].[dbo].[Hermes_Reports]
                  where (ScheduleTypeID = 2 -- monthly
                         OR
                        (ScheduleTypeID = 0 -- direct schedule
                             AND
                             -- instead of bit mask
                            (
                             (iif(DAY1=1, '1', '') +
                             iif(DAY2=1, '2', '') +
                             iif(DAY3=1, '3', '') +
                             iif(DAY4=1, '4', '') +
                             iif(DAY5=1, '5', '') +
                             iif(DAY6=1, '6', '') +
                             iif(DAY7=1, '7', '')) like '%' + cast(datepart(weekday, ExecutedJob) as char(1)) + '%'
                            )
                        )
                    )
                	AND StatusID = 1
                	AND convert(time,getdate()) >= isnull(timefrom,'00:00')
                    AND GroupName = ?
                ''', (groupname))
        # group_count[0] - number of files, that have to be updated
        # group_count[1] - number of files have been updated
        group_count = self.__cursor.fetchone()
        return group_count[0] == group_count[1]


    def send_crash_mail(self, to="silpo-sql-oper@fozzy.ua"):
        ''' Send mail using msdb.dbo.spsend_mail_db.
        '''
        self.__cursor.execute('''DECLARE @to nvarchar(max) = ?,
          @body nvarchar(max) = 'Работа была экстренно завершена. Подробности в фалйе log.txt'
          EXEC msdb.dbo.sp_send_dbmail
            @recipients = @to,
            @subject = '(Внимание) Прекращение работы программы автоматического обновления',
            @body = @body
            ''', (to,))
        self.__db.commit()


    def send_emergency_mail(self, reportName, to="silpo-sql-oper@fozzy.ua"):
        ''' Send mail using msdb.dbo.spsend_mail_db.
        '''
        self.__cursor.execute('''DECLARE @to nvarchar(max) = ?,
          @body nvarchar(max) = 'Обновлён последним: "'+ ? + '"'
          EXEC msdb.dbo.sp_send_dbmail
            @recipients = @to,
            @subject = '(Ошибка обновления) Не отправлено письмо после обновления отчёта',
            @body = @body
            ''', (to, reportName))
        self.__db.commit()


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
        print('Update failed! Error number: {}.'.format(update_error))
        self.__cursor.execute('UPDATE [SILPOAnalitic].[dbo].[Hermes_Reports] \
                              SET Error = ?, LastDateUpdate = cast(? as datetime) \
                              WHERE ReportID = ?', (update_error, update_time, rID))
        self.__db.commit()


if __name__ == '__main__':
    import getpass
    with DBConnect() as dbconn:
        assert dbconn.group_mail_check('Нулевой ЦТЗ') == 0, 'Group check failed.'
        dbconn.send_emergency_mail('Test', getpass.getuser() + '@fozzy.ua')
    print('Connected successfully.')
    input('Press Enter to exit...')