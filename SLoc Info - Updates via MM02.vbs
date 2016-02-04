'****************************************************************************
'	Purpose: Populate/Update Material Master Storage Location information via MM02
'			(SLoc MRP Indicator, SpecialProcurementType(SPT), ROP, and ROQ
'
'	Input: Excel file (\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\SLoc Info - Updates via MM02_data.xlsx)
'
'	Variables - 
'
'	Created on: 04-09-2013
'	Created by: Danielle S. Thomas
'	
'	Revision		Date		Description
'
'****************************************************************************

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

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open("\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\SLoc Info - Updates via MM02_data.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets(1)

PauseCounter = 0

'User is prompted to enter first row of Excel spreadsheet to be read
Row=InputBox("Row to start at")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0

Do Until ExcelSheet.Cells(Row,1).Value = ""

	'Clear Material number value
	Session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ""

	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = Excelsheet.cells(Row,1).value		'Material number (i.e. 00151-673-rebuilt)
	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 0
	session.findById("wnd[0]").sendVKey 0
	
	'Select MRP4 view from SELECT VIEW screen
	'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,11]").setFocus
	'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,11]").caretPosition = 0
	session.findById("wnd[1]/tbar[0]/btn[0]").press			'Select GREEN CHECKMARK to continue to ORG LEVELS screen
		
	'Populate Plant and Stor. Location into the Organizational Levels prompt
	session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,2).value		'Plant (i.e. 50DE)
	session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = ExcelSheet.Cells(Row,3).value		'Storage Location (i.e. D027)
	session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").setFocus
	session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").caretPosition = 4
	session.findById("wnd[1]/tbar[0]/btn[0]").press			'Select GREEN CHECKMARK to continue to MRP4 tab screen
	
	If Not Session.findbyid("wnd[1]", False) Is Nothing Then
		ErrorMsg1 = Session.findbyid("wnd[2]/usr/txtMESSTXT1").text
		ErrorMsg2 = Session.findbyid("wnd[2]/usr/txtMESSTXT2").text
		StatusMsg = ErrorMsg1 & " " & ErrorMsg2
		ExcelSheet.Cells(Row,8).value = StatusMsg
		session.findById("wnd[2]/tbar[0]/btn[0]").press		'Click GREEN CHECKMARK to exit from Error Message
		session.findById("wnd[1]").close	'Close the Org Levels pop-up screen
	else
		'Set SLOC MRP Indicator, Spec.Proc.Type:SLOC, Reorder point, and Replenishment qty to NULL
		session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB6:SAPLMGD1:2498/ctxtMARD-DISKZ").text = ExcelSheet.Cells(Row,4).value		'SLoc MRP Indicator
		session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB6:SAPLMGD1:2498/ctxtMARD-LSOBS").text = ExcelSheet.Cells(Row,5).value		'SpecProcType
		session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB6:SAPLMGD1:2498/txtMARD-LMINB").text = ExcelSheet.Cells(Row,6).value		'ROP
		Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB6:SAPLMGD1:2498/txtMARD-LBSTF").text = ExcelSheet.Cells(Row,7).value		'ROQ
		session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB6:SAPLMGD1:2498/txtMARD-LBSTF").setFocus
		session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB6:SAPLMGD1:2498/txtMARD-LBSTF").caretPosition = 0
		
		session.findById("wnd[0]/tbar[0]/btn[11]").press	'SAVE button
		
		StatusMsg = session.findById("wnd[0]/sbar").text
		ExcelSheet.Cells(Row,8).Value = StatusMsg
	End If
	
	Do Until PauseCounter = 5000
		PauseCounter = PauseCounter + 1
	Loop
	
	Row = Row + 1
	PauseCounter = 0

Loop

ExcelApp.Quit

Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing