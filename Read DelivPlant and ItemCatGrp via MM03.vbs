'***************************************************************************
'	Purpose: Capture General Plant and Item Category Group from Material Master
'			via MM03
'
'	Input: Excel file (\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\Read DelivPlant and ItemCatGrp via MM03_data.xlsx)
'		A		Material
'		B		Plant
'		C		Sales Org
'		D		Item Category Group (captured from MM03)
'		E		Delivering Plant (captured from MM03)
'		F		Exception Message Captured
'
'	Created on: 04-17-2013
'	Created by: Danielle S. Thomas
'	
'	REVISIONS(S)		DATE			DESCRIPTION
'
'
'***************************************************************************

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
Dim Row, StatusMsg, PauseCounter, ErrorMsg1, ErrorMsg2
Dim ItmCGrp,DelivPlant 

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open("\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\Read DelivPlant and ItemCatGrp via MM03_data.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets(1)

PauseCounter = 0

'User is prompted to enter first row of Excel spreadsheet to be read
Row=InputBox("Row to start at")

Session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm03"
session.findById("wnd[0]").sendVKey 0

Do Until ExcelSheet.Cells(Row,1).Value = ""

	'Clear out any values from material input box
	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ""

	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = Excelsheet.cells(Row,1).value		'Material number (i.e. 00151-673-rebuilt)
	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 0
	session.findById("wnd[0]").sendVKey 0

	'session.findById("wnd[1]/tbar[0]/btn[19]").press	'Deselct any selections in Select View(s) screen
	'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(4).selected = True		'Select Sales:Sales Org. Data 2 option
	'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(5).selected = True		'Select Sales:General/Plant Data option
	'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,5]").setFocus
	'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,5]").caretPosition = 0
	session.findById("wnd[1]/tbar[0]/btn[0]").press		'Select GREEN CHECKMARK to continue to ORG LEVELS screen
	
	session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = Excelsheet.cells(Row,2).value		'Plant (ie. 50DC)
	session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = Excelsheet.cells(Row,3).value		'Sales Org (ie. 5013)
	session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").setFocus
	session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").caretPosition = 0
	session.findById("wnd[1]/tbar[0]/btn[0]").press		'Select GREEN CHECKMARK to continue to ORG LEVELS screen
	
	If Not Session.findbyid("wnd[1]", False) Is Nothing Then
		ErrorMsg1 = Session.findbyid("wnd[2]/usr/txtMESSTXT1").text
		ErrorMsg2 = Session.findbyid("wnd[2]/usr/txtMESSTXT2").text
		StatusMsg = ErrorMsg1 & " " & ErrorMsg2
		ExcelSheet.Cells(Row,6).value = StatusMsg
		session.findById("wnd[2]/tbar[0]/btn[0]").press		'Click GREEN CHECKMARK to exit from Error Message
		session.findById("wnd[1]").close	'Close the Org Levels pop-up screen
	Else	
		'Capture Item Category Group value
		ItmCGrp = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-MTPOS").text
		ExcelSheet.Cells(Row,4).Value = ItmCGrp
		'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-MTPOS").caretPosition = 4
		
		'Select Sales:General/Plant tab
		session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06").select
		
		'Capture General Plant value
		DelivPlant = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB1:SAPLMGD1:1001/ctxtRMMG1-WERKS").text
		ExcelSheet.Cells(Row,5).Value = DelivPlant
		
		'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB1:SAPLMGD1:1001/ctxtRMMG1-WERKS").caretPosition = 2
		'session.findById("wnd[0]").sendVKey 4
		'session.findById("wnd[1]/tbar[0]/btn[12]").press
		'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB1:SAPLMGD1:1001/ctxtRMMG1-WERKS").caretPosition = 4
		
		'Green arrow back to main screen
		Session.findById("wnd[0]/tbar[0]/btn[3]").press
	End If
	
	Row = Row + 1

Loop

ExcelApp.Quit

Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing