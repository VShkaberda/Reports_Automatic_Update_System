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
    reportID INT NOT NULL,
    report_name  VARCHAR(20) NOT NULL,
    refresh    INT NOT NULL,
    done    INT NOT NULL,
    priority INT NOT NULL,
    error INT NULL,
    updatetime VARCHAR(20) NULL
);

INSERT INTO Reports (reportID, report_name, refresh, done, priority)
 VALUES ( 1, '1.xlsx',      0 , 0, 2),
        ( 2, '1_NO.xlsx',   0 , 0, 2),
        ( 3, '2.xlsm',      1 , 0, 2),
        ( 4, '3.xlsb',      1 , 0, 1)
    ''')

conn.commit()

conn.close()