# -*- coding: utf-8 -*-
"""
Created on Mon Jan 29 23:52:37 2018
@author: Vadim Shkaberda
"""
from log_error import writelog
import sys
import win32com.client as win32

# example of subsription and default recipient
REPLY_TO = b'\xd0\xa4\xd0\xbe\xd0\xb7\xd0\xb7\xd0\xb8|\
\xd0\x9b\xd0\xbe\xd0\xb3\xd0\xb8\xd1\x81\xd1\x82\xd0\xb8\xd0\xba\xd0\xb0|\
\xd0\x90\xd0\xbd\xd0\xb0\xd0\xbb\xd0\xb8\xd1\x82\xd0\xb8\xd0\xba\xd0\xb8'.decode()

def send_mail(*, to=REPLY_TO, copy=None, subject='Без темы',
              body='', HTMLBody=None, att=None, rName=''):
    ''' Function takes a list of named arguments and sends an e-mail
        via local Outlook account.
        Return 0 if no errors occured, otherwise error number.
    '''
    errorID = 0
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = to
        if copy:
            mail.CC = copy
        mail.Subject = subject
        mail.Body = body
        if HTMLBody:
            mail.HTMLBody = HTMLBody
            
            mail.HTMLBody += '<p><font face="Calibri" color ="Black" size = 3>\
                                С уважением,<br>{}</font></p>'.format(REPLY_TO)
        # In case you want to attach a file to the email
        if att:
            for att_file in att:
                mail.Attachments.Add(att_file)
        mail.Send()
    except Exception as e:
        writelog(e, rName)
        print( "Common Error: %s" % str(e) )
        print(e)
        errorID = 6
    return errorID


if __name__ == '__main__':
    mode = input('Mail sending module test.\n\
          Press 1 to send without attachment.\n\
          Press 2 to send with attachment\n\
          Press 3 to send test mail to logist-analytics\n')
    if not (mode == '1' or mode == '2' or mode == '3'):
        print('Wrong input. Exiting.')
        sys.exit(1)
    if mode == '3':
        #create dict
        k = {'subject': '(Test) Test',
            'HTMLBody': 'Test <i>HTMLBody</i>',
            'rName': 'test rName'
            }
        send_mail(**k)
        sys.exit(1)
    to = input('Mail To (If None will use logist-analytics):')
    copy = input('Mail Copy:')
    subject = input('Subject:')
    body = input('Body:')
    if mode == '2':
        from os import path, getcwd
        root = getcwd()
        att = input('Attachment file name:')
        att  = [path.join(root, att)]
        if to:
            send_mail(to=to, copy=copy, subject=subject, body=body, att=att)
        else:
            send_mail(copy=copy, subject=subject, body=body, att=att)
    else:
        if to:
            send_mail(to=to, copy=copy, subject=subject, body=body)
        else:
            send_mail(copy=copy, subject=subject, body=body)