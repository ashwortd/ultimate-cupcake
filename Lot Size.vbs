'*****************************************************************************************
'	Purpose: Add Max Qty per Period (Lot Size) to the ZSD_CONS_OA table
'
'	Input: \\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\Lot Size_data.xls
'
'	Variables - WERKS - Plant
'				LGORT - 
'				MATNR - Material number
'				MAX_QTY - Max Qty per Period
'				ZZSSCODE - Comment field
'
'	Created: 01-31-2013
'	Created by: Danielle S. Thomas
'
'	REVISION(S)			DATE			DESCRIPTION
'
'*****************************************************************************************

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim CurrDate

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open("\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\Lot Size_data.xls")
Set ExcelSheet = ExcelWorkbook.Worksheets(1)

CurrDate = Date

'User prompted to enter record in Excel file to read
Row=InputBox("Row to start at")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nzsd_cons_oa"
session.findById("wnd[0]").sendVKey 0

Do Until ExcelSheet.Cells(Row,1).Value = ""
	If ExcelSheet.cells(row,13).value <> "error" Then    'bypass records not found in ZSD_CONS_OA table. DST 1-31-13
		session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = ExcelSheet.Cells(Row,1).Value		'Plant
		session.findById("wnd[0]/usr/ctxtS_LGORT-LOW").text = ExcelSheet.Cells(Row,2).Value		'Storage location
		session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = ExcelSheet.Cells(Row,3).Value		'Material Number
		session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").setFocus
		session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").caretPosition = 0
		session.findById("wnd[0]").sendVKey 0		'Hit ENTER key
		
		session.findById("wnd[0]/tbar[1]/btn[8]").press
		
		'Enter the lot size into the Max Qty per Period field in the Z-table
		session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").modifyCell 0,"MAX_QTY",ExcelSheet.Cells(Row,4).Value
		session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "ZZSSCODE"
		session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").triggerModified
		
		'Enter the Stocking Status code in the Z-table
		session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").modifyCell 0,"ZZSSCODE",ExcelSheet.Cells(Row,5).Value
		session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = "ZZCOMMENT"
		session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleColumn = "ZZCMAXQ"
		session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").triggerModified
		
		'Enter a comment into the Comment Field
		session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").modifyCell 0,"ZZCOMMENT","Lot size added " & CurrDate
		session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").triggerModified
		
		session.findById("wnd[0]/tbar[0]/btn[11]").press	'Save button
		session.findById("wnd[1]/usr/btnBUTTON_1").press	'Confirm change
		session.findById("wnd[0]/tbar[0]/btn[3]").press		'Green arrow back
		
		ExcelSheet.Cells(Row,6).Value = "Lot size added"	'Status entered into the data file
	End If    'end if statement. DST 1-31-13
	
	Row = Row + 1
		
Loop	'End Do Loop
 
ExcelApp.Quit

Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing