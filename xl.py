# -*- coding: utf-8 -*-
"""
Created on Sun Jan  7 16:23:59 2018
@author: Vadim Shkaberda
"""
from os import path

import sqlite3
import win32com.client

def refresh(root, f):
    ''' Function to refresh excel files and write in db that file was refreshed.
        Input: root - path of folder where file is;
            f - excel file name.
    '''
    try:
        xl = win32com.client.DispatchEx("Excel.Application")
        wb = xl.Workbooks.Open(path.join(root, f))

        wb.RefreshAll()
        
        wb.Save()
        wb.Close(True)
        print(f, " updated.")
        
        conn = sqlite3.connect('reports.sqlite')
        cur = conn.cursor()
        
        # Write in the db that file have been updated
        try:
            cur.execute('UPDATE Reports SET Done = 1 WHERE report_name = ?', (f, ))
            conn.commit()
            
        except Exception as e:
            print( "Error: %s" % str(e) )
            
        finally:    
            conn.close()

    except Exception as e:
        print( "Error: %s" % str(e) )
        print("Failed to update ", f)

    finally:
        xl.Quit()