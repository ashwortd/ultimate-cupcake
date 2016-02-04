'********************************************************************************
'	Purpose: Populate Shipping Plant into the Customer-Material Info Record (CMIR) via SAP transaction VD52.
'
'	Input: Excel file (\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\Populate Shipping Plant via VD52_data.xlsx)
'			 Column #		Field Info
'			    A			Material
'				B			Plant
'				C			Storage Location
'				D			Status Message
'
'	VARIABLES - KUNNR - Customer number
'				VKORG - Sales Org
'				VTWEG - Distribution Channel
'				MATNR - Material Number
'				WERKS - Plant			
'
'	Created on: 04-26-2013
'	Created by: Danielle S. Thomas
'	
'	REVISION(S)		DATE		DESCRIPTION
'********************************************************************************

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
Dim Row, StatusMsg

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open("\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\Populate Shipping Plant via VD52_data.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets(1)


'User is prompted to enter first row of Excel spreadsheet to be read
Row=InputBox("Row to start at")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nvd52"
session.findById("wnd[0]").sendVKey 0

Do Until ExcelSheet.Cells(Row,1).Value = ""

	Session.findById("wnd[0]/usr/ctxtKUNNR").text = ExcelSheet.Cells(Row,3).Value			'Customer number (i.e. 10080850)
	Session.findById("wnd[0]/usr/ctxtVKORG").text = ExcelSheet.Cells(Row,1).Value			'Sales Org
	session.findById("wnd[0]/usr/ctxtVTWEG").text = ExcelSheet.Cells(Row,2).Value			'Distribution Channel
	session.findById("wnd[0]/usr/ctxtMATNR_R-LOW").text = ExcelSheet.Cells(Row,4).Value		'Material number
	session.findById("wnd[0]/usr/ctxtMATNR_R-LOW").setFocus
	session.findById("wnd[0]/usr/ctxtMATNR_R-LOW").caretPosition = 0
	
	session.findById("wnd[0]/tbar[1]/btn[8]").press		'EXECUTE from Selection screen
	session.findById("wnd[0]/tbar[1]/btn[2]").press		'INFO RECORD DETAILS button

	'First checks to see if Z-table Plant value is already populate in VD52.
	'If the field value matches the Z-table value, nothing is done. Otherwise, the Shipping Plant value is updated in VD52.
	If session.findById("wnd[0]/usr/ctxtMV10A-WERKS").text = ExcelSheet.Cells(Row,5).Value Then
		session.findById("wnd[0]/tbar[0]/btn[3]").press				'GREEN ARROW BACK button to Overview Screen
		session.findById("wnd[0]/tbar[0]/btn[3]").press				'GREEN ARROW BACK button to Selection Screen
		Session.findById("wnd[0]/usr/ctxtKUNNR").text = ""			'Customer number (i.e. 10080850)
		Session.findById("wnd[0]/usr/ctxtVKORG").text = ""			'Sales Org
		session.findById("wnd[0]/usr/ctxtVTWEG").text = ""			'Distribution Channel
		session.findById("wnd[0]/usr/ctxtMATNR_R-LOW").text = ""	'Material number
		Row = Row + 1
	Else 	'If the VD52 Shipping Plant <> Z-table plant, a the z-table plant is entered as the VD52 Shipping Plant
		session.findById("wnd[0]/usr/ctxtMV10A-WERKS").text = ExcelSheet.Cells(Row,5).Value		'Plant field (i.e. 50DH)
		session.findById("wnd[0]/usr/ctxtMV10A-WERKS").setFocus
		session.findById("wnd[0]/usr/ctxtMV10A-WERKS").caretPosition = 4
		session.findById("wnd[0]/tbar[0]/btn[11]").press	'SAVE button
		
		StatusMsg = session.findById("wnd[0]/sbar").text
		ExcelSheet.Cells(Row,6).Value = StatusMsg
		
		Session.findById("wnd[0]/usr/ctxtKUNNR").text = ""			'Customer number (i.e. 10080850)
		Session.findById("wnd[0]/usr/ctxtVKORG").text = ""			'Sales Org
		session.findById("wnd[0]/usr/ctxtVTWEG").text = ""			'Distribution Channel
		session.findById("wnd[0]/usr/ctxtMATNR_R-LOW").text = ""	'Material number	
		Row = Row + 1
	End If 

Loop

ExcelApp.Quit

Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing