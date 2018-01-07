# -*- coding: utf-8 -*-
"""
Created on Sun Jan  7 15:53:26 2018
@author: Vadim Shkaberda
"""

from os import path, getcwd
from xl import refresh

import sqlite3
import threading
import time


def main():
    
    #root = r'XL'
    root = path.join(getcwd(), 'XL')    
    
    # Check if interrupted
    while thread.is_alive():
        files = []
        
        # Open database
        conn = sqlite3.connect('reports.sqlite')
        cur = conn.cursor()
        
        try:
            # List of reports
            cur.execute("SELECT report_name, refresh, done FROM Reports")
            while True:
                sql_data = cur.fetchone()
                if sql_data is None:
                    break
                # if need to update and NOT updated yet
                if sql_data[1] and not sql_data[2]:
                    files.append(sql_data)
            conn.commit()
        
        finally:    
            conn.close()
        
        # Checking for each file needed to update
        for f in files:
            # additional check for interruption
            if thread.is_alive():                
                refresh(root, f[0])
                time.sleep(5)
            
    print('Refreshing is interrupted.')
    
        
def test_thread():
    ''' Function that monitors user input.
        If 1 have been inputed:
        function is exiting, causing interruption of main()'''
    while True:
        output = input("Press 1 if you want to interrupt refreshing.\n")
        if output == '1':
            break


if __name__ == "__main__":
    # Start thread that monitors user input
    thread = threading.Thread(target=test_thread)
    # Python program exits when only daemon threads are left
    thread.daemon = False
    thread.start()
    main()
