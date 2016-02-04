'********************************************************************************
'	Purpose: Update Cycle Counting information via SAP transaction MM02.
'			 General Plant Data / Storage 1 tab - Update Cycle Counting and CC Fixed values
'
'	Input: Excel file (Cycle Counting - Update Info via MM02_data.xlsx)
'			 Column #		Field Info
'			    A			Material
'				B			Plant
'				C			Storage Location
'				D			Status Message				
'
'	Created on: 04-26-2013
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
Dim Row, StatusMsg

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open("\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\Cycle Counting - Update Info via MM02_data.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets(1)


'User is prompted to enter first row of Excel spreadsheet to be read
Row=InputBox("Row to start at")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0

Do Until ExcelSheet.Cells(Row,1).Value = ""

	'Clear Material number value
	Session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ""
	
	'Populate Material number into text box
	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = Excelsheet.cells(Row,1).value		'Material number (i.e. MPS-8800345)
	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 0
	
	'Click "Select view(s)" button
	session.findById("wnd[0]/tbar[1]/btn[5]").press
	
	'Deselect any/all selections in the 'Select View(s)' pop-up screen
	'session.findById("wnd[1]/tbar[0]/btn[19]").press

	'Selects the "General Plant Data / Storage 1" option in the 'Select View(s)' pop-up screen
	'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(12).selected = true
	'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,12]").setFocus
	'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,12]").caretPosition = 0
	session.findById("wnd[1]/tbar[0]/btn[0]").press		'Select GREEN CHECKMARK to continue to Organizational Levels screen
	
	'Enter Organizational Levels parameters	
	session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = Excelsheet.cells(Row,2).value		'Plant (i.e. 500B)
	session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = Excelsheet.cells(Row,3).value		'Storage Location (i.e. GV01)
	session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").setFocus
	session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").caretPosition = 0
	session.findById("wnd[1]/tbar[0]/btn[0]").press		'Select GREEN CHECKMARK
	
	If Not Session.findbyid("wnd[1]", False) Is Nothing Then
		ErrorMsg1 = Session.findbyid("wnd[2]/usr/txtMESSTXT1").text
		ErrorMsg2 = Session.findbyid("wnd[2]/usr/txtMESSTXT2").text
		StatusMsg = ErrorMsg1 & " " & ErrorMsg2
		ExcelSheet.Cells(Row,4).value = StatusMsg
		session.findById("wnd[2]/tbar[0]/btn[0]").press		'Click GREEN CHECKMARK to exit from Error Message
		session.findById("wnd[1]").close	'Close the Org Levels pop-up screen
	else
		'Checks if "CC Phys. inv. ind." field is visible
		                    'wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/lblMARA-BEHVO
		On Error Resume next
		'If Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/ctxtMARC-ABCIN").text = "" Then
			If Err.Number <> 0 then
				StatusMsg = session.findById("wnd[0]/sbar").text
				ExcelSheet.Cells(Row,4).Value = StatusMsg	'Populate status message in data file
				session.findById("wnd[0]/tbar[0]/btn[3]").press
			End if
	'	Else
			'Checks current value of the "CC phys. inv. ind." field
			If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/ctxtMARC-ABCIN").value <> "D" Then
				session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/ctxtMARC-ABCIN").text = "D"
			End If
			
			'Checks current value of the "CC fixed" field	
			If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/chkMARC-CCFIX").selected = False Then 
				session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/chkMARC-CCFIX").selected = True
			End If 
			
			session.findById("wnd[0]/tbar[0]/btn[11]").press	'SAVE button
			
			StatusMsg = session.findById("wnd[0]/sbar").text
			ExcelSheet.Cells(Row,4).Value = StatusMsg	'Populate status message in data file
		End If
	'End if
	
	Row = Row + 1	'Move to next row of the Excel data file
	
Loop

ExcelApp.Quit

Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing