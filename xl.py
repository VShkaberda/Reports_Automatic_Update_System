# -*- coding: utf-8 -*-
"""
Created on Sun Jan  7 16:23:59 2018
@author: Vadim Shkaberda
"""
from os import path

import pythoncom
import win32com.client

class ReadOnlyException(Exception):
    """Write access is not permitted on file. """
    def __init__(self, f, message='Write access is not permitted on file:', *args):
        self.message = message
        self.f = f
        # allow users initialize misc. arguments as any other builtin Error
        super(ReadOnlyException, self).__init__(f, message, *args)


def update_file(root, f):
    ''' Function to refresh excel files and write in db that file was refreshed.
        Input: root - path of folder where file is;
            f - excel file name.
        Return 0 if update was successful, otherwise error number.
    '''
    update_error = 1
    try:
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.DisplayAlerts = False

        wb = xl.Workbooks.Open(path.join(root, f))

        # check whether the file is read-only
        if xl.ActiveWorkbook.ReadOnly == True:
            raise ReadOnlyException(f)

        xl.Application.Run('\'' + f + '\'!Update')

        wb.Close(SaveChanges=1)
        print(f, " updated.")
        update_error = 0

    except pythoncom.com_error as e:
        print( "Excel Error: %s" % str(e) )

    except ReadOnlyException as e:
        print( "ReadOnly Error: %s" % str(e) )
        print(e.message, e.f)
        update_error = 3

    except Exception as e:
        print( "Common Error: %s" % str(e) )
        print(e)

    finally:
        xl.Quit()
        return update_error

if __name__ == '__main__':
    from os import getcwd
    update_file(path.join(getcwd(), 'XL'), '2.xlsm')