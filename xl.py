# -*- coding: utf-8 -*-
"""
Created on Sun Jan  7 16:23:59 2018
@author: Vadim Shkaberda
"""
from os import path
from shutil import copy2

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
        print("Updating ", f)
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.DisplayAlerts = False

        wb = xl.Workbooks.Open(path.join(root, f))

        # check whether the file is read-only
        if xl.ActiveWorkbook.ReadOnly == True:
            raise ReadOnlyException(f)

        xl.Application.Run('\'' + f + '\'!Update')

        wb.Close(SaveChanges=1)
        update_error = 0

    except pythoncom.com_error as e:
        print( "Excel Error: %s" % str(e) )
        # file hasn't been found
        if e.excepinfo[2].find('Не удалось найти', 0, 16) == 0:
            update_error = 2
        # macro hasn't been found
        elif e.excepinfo[2].find('Не удается выполнить макрос', 0, 27) == 0:
            update_error = 4

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


def copy_file(root, f, dst):
    ''' Copy file f to dst.
        Input: root - path of folder where file is;
            f - excel file name;
            dst - destination (either folder or filename).
        Return 0 if copy was successful, otherwise error number.
    '''
    update_error = 0
    try:
        copy2(path.join(root, f), dst)

    except Exception as e:
        print( "Error while copying: %s" % str(e) )
        print(e)
        update_error = 5

    finally:
        return update_error


if __name__ == '__main__':
    from os import getcwd
    update_file(path.join(getcwd(), 'XL'), '2.xlsm')
