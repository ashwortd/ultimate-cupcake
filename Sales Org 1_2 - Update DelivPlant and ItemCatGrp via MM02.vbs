'********************************************************************************
'	Purpose: Update Material Master information via SAP transaction MM02.
'			 Sales: sales org. 1 tab - Updates Delivering Plant
'			 Sales: sales org. 2 tab - Updates Item Category Group
'
'	Input: Excel file (\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\Sales Org 1_2 - Update DelivPlant and ItemCatGrp via MM02_data.xlsx)
'			 Column #		Field Info
'			    A			Material
'				B			Plant
'				C			Sales Organization
'				D			Distribution Channel
'				E			Delivering Plant
'				F			Item Category Group
'				G			Status Message	
'
'	VARIABLES - MATNR - Material number
'				WERKS - Plant
'				VKORG - Sales Org
'				VTWEG - Distribution Channel
'				DWERK - Delivering Plant
'				MTPOS - Item Category Group
'
'	Created on: 04-10-2013
'	Created by: Danielle S. Thomas
'	
'	Version:
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
Dim Row, StatusMsg, PauseCounter

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open("\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\Sales Org 1_2 - Update DelivPlant and ItemCatGrp via MM02_data.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets(1)


'User is prompted to enter first row of Excel spreadsheet to be read
Row=InputBox("Row to start at")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0

Do Until ExcelSheet.Cells(Row,1).Value = ""

	'Clear Material number value
	Session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ""

	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = Excelsheet.cells(Row,1).value		'Material number (i.e. MPS-8800345)
	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 0
	session.findById("wnd[0]").sendVKey 0

	'Deselect any/all selections in the 'Select View(s)' pop-up screen
	'Session.findById("wnd[1]/tbar[0]/btn[19]").press

	'Selects the Sales: Sales Org. Data 1 & 2 view in the 'Select View(s)' pop-up screen
	'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = true
	'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(1).selected = True
	'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,11]").setFocus
	'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,11]").caretPosition = 0
	session.findById("wnd[1]/tbar[0]/btn[0]").press		'Select GREEN CHECKMARK to continue to Organizational Levels screen
	
	'Enters Organizational Levels parameters
	session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = Excelsheet.cells(Row,2).value		'Plant (i.e. 50xx)
	session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = Excelsheet.cells(Row,3).value		'Sales Org (i.e. 5013)
	session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = Excelsheet.cells(Row,4).value		'Distribution Channel (i.e. 01)
	session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").setFocus
	session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").caretPosition = 2
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	
	'Update Sales Org 1 Delivering Plant
	Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").text = Excelsheet.cells(Row,5).value		'Delivering Plant (i.e. 50xx)
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").setFocus
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").caretPosition = 4
	
	'Select Sales Org 2 tab
	Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05").select
	
	'Update Sales Org 2 Item Category Group
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-MTPOS").text = Excelsheet.cells(Row,6).value		'Item Category Group (i.e. ZVOR)
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-MTPOS").setFocus
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-MTPOS").caretPosition = 4
	
	session.findById("wnd[0]/tbar[0]/btn[11]").press	'SAVE button
	
	StatusMsg = session.findById("wnd[0]/sbar").text
	ExcelSheet.Cells(Row,7).Value = StatusMsg	'Populate status message in data file
	
	Row = Row + 1	'Move to next row of the Excel data file
	
Loop

ExcelApp.Quit

Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing