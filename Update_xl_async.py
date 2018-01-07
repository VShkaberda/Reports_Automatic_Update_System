# -*- coding: utf-8
"""
Created on Sun Jan  7 15:53:26 2018
@author: Vadim Shkaberda
"""
from os import path, getcwd
from xl import refresh

import sys, msvcrt
import asyncio
import sqlite3
import time


async def main():
    
    #root = r'XL'
    root = path.join(getcwd(), 'XL')
    
    files = []
    
    # Open database
    conn = sqlite3.connect('reports.sqlite')
    cur = conn.cursor()
    
    try:    
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
        refresh(root, f[0])
        # Here we check for manual interruption.
        await asyncio.sleep(0)
        time.sleep(5)
        
    print(files)


async def read_input( caption, default=0, timeout = 5):
    ''' Function waits for user input for N=timeout seconds.'''
    while True:
        start_time = time.time()
        sys.stdout.write('%s(%s):'%(caption, default))
        sys.stdout.flush()
        output = ''
        while True:
            if msvcrt.kbhit():
                byte_arr = msvcrt.getche()
                if ord(byte_arr) == 13: # enter_key
                    break
                elif ord(byte_arr) >= 32: #space_char
                    output += "".join(map(chr,byte_arr))
            if len(output) == 0 and (time.time() - start_time) > timeout:
                print("timing out, using default value.")
                break

        if output == '1':
            print('Interrupting.')
            ioloop.stop()
            await asyncio.sleep(0)
            
        print('No interrupting.')
        await asyncio.sleep(0)


if __name__ == "__main__":
    ioloop = asyncio.get_event_loop()
    tasks = [ioloop.create_task(main()), ioloop.create_task(read_input('Press 1 to interrupt refreshing'))]
    wait_tasks = asyncio.wait(tasks)
#    ioloop.run_until_complete(wait_tasks)
    ioloop.run_forever()
    ioloop.close()
