'**************************************************************************
'	Purpose - Update the ROP and ROQ values for all Optimized material.
'
'	Input - Excel file (\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\UpdateOptimized_ROP-ROQ_Data.xlsx)
'			ROP and ROQ values greater than zero (0) will be changed to 0 for all material where Status = "O"
'
'	Output - Updated ZSD_CONS_OA table
'
'	Variables: WERKS - Plant
'				LGORT - Storage Location
'				MATNR - Material number
'				ZZSSCODE - Stocking Status Code (aka. ABC Policy Code)
'
'	Created by: Danielle S. Thomas 02-14-2013
'
'	Modified (Please document any modifications using the template below):
'		Date	Name	Description of Modification
'
'**************************************************************************

If Not IsObject(application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject application, "on"
End If

Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim PolCode,NumOfRows,StatusText
Dim ROP,ROQ

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open("\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\UpdateOptimized_ROP-ROQ_Data.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets(1)

CurrDate = Date

'User is prompted to enter first row of Excel spreadsheet to be read
Row=InputBox("Row to start at")
Counter = 0

Session.findById("wnd[0]").maximize
Session.findById("wnd[0]/tbar[0]/okcd").text = "/nzsd_cons_oa"
session.findById("wnd[0]").sendVKey 0

' ** Do Until Loop will execute until the last row (a blank row) is found **
Do Until ExcelSheet.Cells(Row,1).Value = ""
	Session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = ExcelSheet.Cells(Row,2).Value  'Plant
	Session.findById("wnd[0]/usr/ctxtS_LGORT-LOW").text = ExcelSheet.Cells(Row,3).Value  'Storage loc
	Session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = ExcelSheet.Cells(Row,4).Value  'Material number
	Session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").setFocus
	session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").caretPosition = 11
	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	NumOfRows = Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").rowcount
	
	Do Until Counter = (NumofRows)
			PolCode = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(Counter,"ZZSSCODE")
			If PolCode = "O" Then 
				ROP = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(Counter,"ZZCREORDP")
				
				If ROP <> "0" Then
					session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").modifyCell Counter,"ZZCREORDP","0"
					Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").triggerModified
				End If
								
				ROQ = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(Counter,"ZZCREPLQ")
				If ROQ <> "0" Then
					session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").modifyCell Counter,"ZZCREPLQ","0"
					Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").triggerModified
				End If
											
				Counter = Counter + 1
			Else				
				Counter = Counter + 1	
			End If
	Loop

	Session.findById("wnd[0]/tbar[0]/btn[11]").press
	Session.findById("wnd[1]/usr/btnBUTTON_1").press
	StatusText = Session.findbyid("wnd[0]/sbar").text
	Session.findById("wnd[0]/tbar[0]/btn[3]").press
	ExcelSheet.Cells(Row,12).Value = StatusText
	
	row=row+1
	Counter = 0
Loop 


ExcelApp.Quit

Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing