# -*- coding: utf-8 -*-
"""
Created on Sun Jan  7 15:53:26 2018
@author: Vadim Shkaberda
"""
import sqlite3

conn = sqlite3.connect('reports.sqlite')
cur = conn.cursor()

# Make some fresh tables using executescript()
cur.executescript('''
DROP TABLE IF EXISTS Reports;

CREATE TABLE Reports (
    report_name  VARCHAR(20) NOT NULL,
    refresh    INT NOT NULL,
    done    INT NOT NULL
);

''')


cur.execute('''INSERT INTO Reports
    (report_name, refresh, done) 
    VALUES ( '1.xlsx', 1 , 0),
        ( '1_NO.xlsx', 0 , 0),
        ( '2.xlsm', 1 , 0),
        ( '3.xlsb', 1 , 0)
    ''')

conn.commit()

conn.close()