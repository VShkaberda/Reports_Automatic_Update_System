# Reports_Automatic_Update_System

<b>Update_xl_thread.py</b> is a script to automatically update reports.

<ul>
<li>createdb.py - creates database(SQLite);</li>
<li>db_connect.py - connects to SQLite database (for testing);</li>
<li>db_connect_sql.py - connects to SQL Server;</li>
<li>log_error.py - module to log errors;</li>
<li>send_mail.py - contains functions to work with Outlook;</li>
<li>sharepoint.py - checks connection to Sharepoint;</li>
<li>xl.py - contains functions to work with Excel;</li>
<li>Update_xl_async.py - script for automatic update (asynchronous) - <i>removed</i>;</li>
<li>Update_xl_thread.py - script for automatic update (synchronous).</li>
</ul>
<hr>

### Update_xl_thread.py

Update_xl_thread.py connects to the table [dbo].[VV_Hermes_Reports_Temp] on 64 Server and gets info about file needed to be updated. After providing an update, writes info about success/failure into server.

### db_connect_sql.py

Module to work with database.

class `DBConnect`() - establishes connection to database.

Uses methods:

`error_description`() - downloads error description from db.

`file_to_update`() - returns info about file from db.

`group_attachments`(<i>groupname</i>) - generator of attachments for email after updating all of <i>groupname</i> reports.

`group_mail_check`(<i>groupname</i>) - returns 1 if all files from group <i>groupname</i> have been updated.

`send_crash_mail`(<i>to</i>) - sends mail from server to recipient <i>to</i> using `msdb.dbo.sp_send_dbmail`. Contains message about crash of main program.

`send_emergency_mail`(<i>reportName, to</i>) - sends mail from server to recipient <i>to</i> using `msdb.dbo.sp_send_dbmail`. Contains message about failure with sending mail after <i>reportName</i> has been updated.

`successful_update`(<i>rID, update_time</i>) - writes <i>update_time</i> into field [LastDateUpdate] for row with [ReportID] = <i>rID</i> into db.

`failed_update`(<i>rID, update_time, update_error</i>) - writes <i>update_error</i> into field [Error] and <i>update_time</i> into field [LastDateUpdate] for row with [ReportID] = <i>rID</i> to db.

### log_error.py

function `writelog`() - writes info about error into file log.txt to the same direcotory where the main program is.

### sharepoint.py

Moodule to check connection to sharepoint.

function `sharepoint_check`() - checks existence of the specific folder on sharepoint. If folder is not found, maps drive R to sharepoint.

### send_mail.py

Module to send mail using local Outlook account.

function `send_mail`(<i>to, copy, subject, body, HTMLBody, att</i>) - sends e-mail to <i>to</i> and copy to <i>copy</i> with subject <i>subject</i>. The body of letter is <i>body</i> and optional <i>HTMLBody</i>. Attachment <i>att</i> can be added.

### xl.py

Module to work with Excel.

class `ReadOnlyException`() - class for catching read-only state of Excel files.

function `update_file`(<i>root, f</i>) - updates Excel file <i>f</i>. <i>root</i> - path to file.

function `cop_file`(<i>root, f, dst</i>) - copy file <i>f</i> to destination <i>dst</i>, where <i>dst</i> - either folder or filename.

<hr>

`Upd.manifest` - added to prevent admin privileges requirements after converting module into executable file. Tested on 32-bit Python with Pyinstaller. 

Usage: `pyinstaller -m Upd.manifest Update_xl_thread.py`

<hr>

## Requierments

* [pyodbc](https://github.com/mkleehammer/pyodbc)
 :warning: Works with `pyodbc` version < 4.0.22. `Pyodbc` 4.0.22 have an [issue with updating/inserting into field with type smalldatetime](https://github.com/mkleehammer/pyodbc/issues/334).

* [pythoncom](https://github.com/mkleehammer/pyodbc) (part of [pywin32](https://github.com/mhammond/pywin32); when have a problem installing pywin32, check [pypiwin32](https://stackoverflow.com/questions/49307303/installing-the-pypiwin32-module) and [postinstall](https://www.reddit.com/r/Python/comments/57h1pf/pywin32_not_installing_properly/))

* win32com (also part of pywin32)
