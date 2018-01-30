# -*- coding: utf-8 -*-
"""
Created on Mon Jan 29 23:52:37 2018
@author: Vadim Shkaberda
"""
import sys
import win32com.client as win32

def send_mail(*, to='Фоззи|Логистика|Аналитики', copy=None, subject='Без темы',
              body=None, HTMLBody=None, att=None):
    ''' Function takes a list of named arguments and sends an e-mail
        via local Outlook account.
    '''
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to
    if copy:
        mail.CC = copy
    mail.Subject = subject
    mail.Body = body
    if HTMLBody:
        mail.HTMLBody = HTMLBody
    # In case you want to attach a file to the email
    if att:
        mail.Attachments.Add(att)
    mail.Send()


if __name__ == '__main__':
    mode = input('Mail sending module test.\n\
          Press 1 to send without attachment.\n\
          Press 2 to send with attachment\n')
    if not (mode == '1' or mode == '2'):
        print('Wrong input. Exiting.')
        sys.exit(1)
    to = input('Mail To (If None will use logist-analytics):')
    copy = input('Mail Copy:')
    if mode == '2':
        from os import path, getcwd
        root = getcwd()
        att = input('Attachment file name:')
        att  = path.join(root, att)
        send_mail(to, copy, att)
    else:
        send_mail(to, copy)