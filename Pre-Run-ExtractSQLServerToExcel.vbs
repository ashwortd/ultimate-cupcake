'This script can help you to extract records from an SQL server table into current external Excel file
'Please note that you need to change at appropriate place before using this script as mentioned in the comment.

Dim wBook		'=== Excel workbook object
Dim wSheet		'=== Excel worksheet object
Dim wshShell		'=== 
Dim wshUsrEnv		'=== 

Dim mConnection                     '=== ADO connection object for database
Dim rs                              '=== ADO recordset object for table
Dim i, j                            '=== For Excel row and column numbers

Set wshShell = CreateObject("WScript.Shell")		'=== Get Shell object from WScript object
Set wshUsrEnv = wshShell.Environment("USER")		'=== Get current user environment variables
wshUsrEnv("PRERRORLEVEL") = 1                           '=== Set ProcessRunner error level environment variable to 1, when script runs without error reset it to 0 to mark success.

Set mConnection = CreateObject("ADODB.Connection")      '===Connection object initialization
Set rs = CreateObject("ADODB.Recordset")                '===Recordset object initialization

set wBook = GetObject(#CURXLFILE#) 'The currrent external Excel file
Set wSheet = wBook.Worksheets(#CURXLSHEET#) 'Cureent worksheet of current external Excel file

i = #XLDATASTARTROW#    '=== Put here the start row number from which data rows writing will start
k = 1                   '===Put here the start column number from which data columns writing will start

mConnection.Open "Provider=SQLOLEDB.1;Data Source=ip-address-of-your-database-machine;Initial Catalog=[yourdatabasename]","databaseusername","userpassword"              '===Change values here with your actual values like    '=====IPaddress, database name, username and password.

if mConnection.state = 1 then
	rs.Open "SELECT [Field1], [Field2], [Field3], [Field4], [Field5]  FROM  [TableName]" ,mConnection    '===Change SELECT query here stating table name and field list you want to extract
	while not rs.eof
		for j = 0 to rs.Fields.count - 1
			wSheet.Cells(i, k+j) = rs.Fields.Item(j) & "" 
		next	
		rs.movenext
		i = i + 1
	wend               
              rs.close
              mConnection.close
end if

Set cmd = Nothing
Set mConnection = Nothing

wshUsrEnv("PRERRORLEVEL") = 0  'Script ends without error here, reset environment variable to 0 to mark success.

wshShell = null
wshUsrEnv = null

wSheet = null
wBook = null