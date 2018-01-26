# Reports_Automatic_Update_System
<ul>
<li>createdb.py - file for creating database(SQLite);</li>
<li>db_connect.py - file for connecting to SQLite database (for testing);</li>
<li>db_connect_sql.py - file for connecting to SQL Server;</li>
<li>xl.py - contains functions to work with Excel;</li>
<li>Update_xl_async.py - script for automatic update(asynchronous);</li>
<li>Update_xl_thread.py - script for automatic update(synchronous).</li>
</ul>
<hr>

## Update_xl_thread.py

<p><b>Update_xl_thread.py</b> is a script to automatically update reports.</p>
<p>Update_xl_thread.py connects to table [dbo].[VV_Hermes_Reports_Temp] on 64 Server and gets info about file needed to be updated. After providing an update, writes info about success/failure into server.</p>

## db_connect_sql.py

<p><font size="6">Module to work with database.</font></p>
<p>class <b>DBConnect</b>() - establishes connection to database.</p>
<p>Uses methods:</p>
<p><b>file_to_update</b>() - returns info about file from db.</p>
<p><b>successful_update</b>(<i>rID, update_time</i>) - writes <i>update_time</i> into field [LastDateUpdate] for row with ReportID = <i>rID</i> into db.</p>
<p><b>failed_update</b>(<i>rID, update_time</i>) - writes 1 into field [Error] and <i>update_time</i> into field [LastDateUpdate] for row with ReportID = <i>rID</i> into db.</p>

## xl.py

<p>Module to work with Excel.</p>
<p>class <b>ReadOnlyException</b>() - class for catching read-only state of Excel files.</p>
<p>function <b>update_file</b>(<i>root, f</i>) - updates Excel file <i>f</i>. <i>root</i> - path to file.</p>
