# -*- coding: utf-8 -*-
"""
Created on Wed Aug  1 18:00:21 2018

@author: v.shkaberda
"""

import os

def sharepoint_check():
    ''' Check connection to sharepoint. If no - try to mount a drive.
    '''
    print('Проверка связи с sharepoint...')

    test_sharepoint = R'\\sharepoint.fozzy.lan\sites\logistics\vegetables_fruits'

    if not os.path.isdir(test_sharepoint):
        print('Монтирование диска sharepoint...')
        sharepoint_url = 'http://sharepoint.fozzy.lan'
        os.system('net use r: {}'.format(sharepoint_url))
        if not os.path.isdir(test_sharepoint):
            print('Доступ к sharepoint не установлен. \
                  Пожалуйста, зайдите на ресурс вручную.')
            return
        print('Диск смонтирован. Доступ к sharepoint установлен.')
        return

    print('Sharepoint доступен.')


if __name__ == '__main__':
    sharepoint_check()