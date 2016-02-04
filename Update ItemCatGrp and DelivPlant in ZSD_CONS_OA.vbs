'********************************************************************************************
'	Purpose: Update Item Category Group and Delivering Plant fields in the ZSD_CONS_OA table
'
'	Input: Excel file (\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\Update ItemCatGrp and DelivPlant in ZSD_CONS_OA_data.xlsx)
'
'		A		Sales Org
'		B		Plant
'		C		Material
'		D		ZSD Item Category Group
'		E		ZSD Delivering Plant
'		F		Material Master Item Category Group
'		G		Material Master Delivering Plant
'		H		PMx Status Message
'
'	Created on: 04-18-2013
'	Created by: Danielle S. Thomas
'	
'	Version:
'*********************************************************************************************

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
Dim Row,ChangeItemCatGrp,ChangeDelivPlant,StatusText
Dim NumOfRows,CountTableRows,NewItemCatGrpValue,NewDelivPlantValue

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open("\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\Update ItemCatGrp and DelivPlant in ZSD_CONS_OA_data.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets(1)


'User is prompted to enter first row of Excel spreadsheet to be read
Row = InputBox("Row to start at")
CountTableRows = 0

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nzsd_cons_oa"
Session.findById("wnd[0]").sendVKey 0

' ** Do Until Loop will execute until the last row (a blank row) is found **
Do Until ExcelSheet.Cells(Row,1).Value = ""

	'Compare Item Category Group values
	If ExcelSheet.Cells(Row,4).value <> ExcelSheet.cells(Row,6).value Then
		ChangeItemCatGrp = True
		NewItemCatGrpValue = excelSheet.cells(Row,6).value		'MM Item Cat Group value to be copied into Z-table
	End If
	
	'Compare Delivering plant values
	If ExcelSheet.cells(Row,5).value <> ExcelSheet.cells(Row,7).value Then
		ChangeDelivPlant = True
		NewDelivPlantValue = ExcelSheet.cells(Row,7).value		'MM Delivering Plant value to be copied into Z-table
	End if
	
	If ChangeItemCatGrp Or ChangeDelivPlant Then
		session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "" 	'Clear Plant value from selection screen
		Session.findbyid("wnd[0]/usr/ctxtS_MATNR-LOW").text = ""	'Clear Material value from selection screen
		
		'Enter selection criteria (Plant & Material)
		'session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = ExcelSheet.Cells(Row,2).Value 	'Plant
		Session.findbyid("wnd[0]/usr/ctxtS_MATNR-LOW").text = ExcelSheet.cells(Row,3).value		'Material
		Session.findbyid("wnd[0]/tbar[1]/btn[8]").press 										'Execute button
		NumOfRows = Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").rowcount
		
		Do Until CountTableRows = (NumOfRows)
			If ChangeItemCatGrp Then
				Session.findbyid("wnd[0]/usr/cntlGRID1/shellcont/shell").modifyCell CountTableRows,"MTPOS",""
				Session.findbyid("wnd[0]/usr/cntlGRID1/shellcont/shell").modifyCell CountTableRows,"MTPOS",NewItemCatGrpValue
				Session.findbyid("wnd[0]/usr/cntlGRID1/shellcont/shell").triggermodified
			End If
			
			If ChangeDelivPlant Then
				Session.findbyid("wnd[0]/usr/cntlGRID1/shellcont/shell").modifycell CountTableRows,"WRK02",""
				Session.findbyid("wnd[0]/usr/cntlGRID1/shellcont/shell").modifycell CountTableRows,"WRK02",NewDelivPlantValue
				Session.findbyid("wnd[0]/usr/cntlGRID1/shellcont/shell").triggermodified
			End If
			CountTableRows = CountTableRows + 1
		Loop 
		
		Session.findById("wnd[0]/tbar[0]/btn[11]").press	'Save button
		Session.findbyid("wnd[1]/usr/btnBUTTON_1").press	'Confirm change
		StatusText = session.findById("wnd[0]/sbar").text
		ExcelSheet.Cells(Row,8).Value = StatusText
	Else
		ExcelSheet.cells(Row,8).value = "No changes made"	
	End If
	
	Session.findbyid("wnd[0]/tbar[0]/btn[3]").press		'Green arrow back to entry screen

		
	Row = Row + 1
	CountTableRows = 0
	
Loop

ExcelApp.Quit

Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing
