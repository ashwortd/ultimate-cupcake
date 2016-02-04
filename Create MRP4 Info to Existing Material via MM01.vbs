'************************************************************************************************
' Purpose - Create (txn: MM01) Storage Location MRP Info (MRP4 tab) for existing material
' 
' Input - Excel file: \\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\Create MRP4 Info to Existing Material via MM01_data.xlsx
'
' Variables - VKORG - Sales Org
'			  WERKS - Plant
' 			  MATNR - Material
'             LGORT - Storage location (SLoc)
'             KUNNR - Customer
'             
'
' Created: 4-24-13
' Created by: Danielle Thomas
'
'	REVISION(S)		DATE			DESCRIPTION
'	
'************************************************************************************************

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
Dim Row,StatusMsg,ErrorMsg1,ErrorMsg2

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open("\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\Create MRP4 Info to Existing Material via MM01_data.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets(1)


'User is prompted to enter first row of Excel spreadsheet to be read
Row=InputBox("Row to start at")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
session.findById("wnd[0]").sendVKey 0

Do Until ExcelSheet.Cells(Row,1).Value = ""

	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value		'Material
	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").setFocus
	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 0
	session.findById("wnd[0]").sendVKey 0		'Hit enter key
	
'	If Not Session.findbyid("wnd[1]", False) Is Nothing Then
	If Session.findbyid("wnd[0]/sbar").text <> "" then
		session.findById("wnd[0]").sendVKey 0		'Hit enter key for warning message received
		
		'Session.findById("wnd[1]/tbar[0]/btn[19]").press		'Deselect all - Select View(s) pop-up
		'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(11).selected = True		'Select MRP 4 option
		'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,11]").setFocus
		'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,11]").caretPosition = 0
		session.findById("wnd[1]/tbar[0]/btn[0]").press		'Green check mark button
		session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,2).Value		'Plant
		session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = ExcelSheet.Cells(Row,3).Value		'Storage Location
		session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").setFocus
		session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").caretPosition = 0
		session.findById("wnd[1]/tbar[0]/btn[0]").press			'Green check mark button
		
		If Not Session.findbyid("wnd[1]", False) Is Nothing Then
			ErrorMsg1 = Session.findbyid("wnd[2]/usr/txtMESSTXT1").text
			ErrorMsg2 = Session.findbyid("wnd[2]/usr/txtMESSTXT2").text
			StatusMsg = ErrorMsg1 & " " & ErrorMsg2
			ExcelSheet.Cells(Row,4).value = StatusMsg
			Session.findbyid("wnd[2]/tbar[0]/btn[0]").press		'By-pass message that material already maintained
			session.findById("wnd[1]").close	'Close the Org Levels pop-up screen			
		Else 
			Session.findById("wnd[0]/tbar[0]/btn[11]").press		'SAVE button
		
			StatusMsg = session.findById("wnd[0]/sbar").text
			ExcelSheet.Cells(Row,4).Value = StatusMsg
		End If
			
'		Row = Row + 1
	Else
		'session.findById("wnd[1]/tbar[0]/btn[19]").press		'Deselect all - Select View(s) pop-up
		'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(11).selected = True		'Select MRP 4 option
		'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,11]").setFocus
		'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,11]").caretPosition = 0
		session.findById("wnd[1]/tbar[0]/btn[0]").press		'Green check mark button
		session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,2).Value		'Plant
		session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = ExcelSheet.Cells(Row,3).Value		'Storage Location
		session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").setFocus
		session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").caretPosition = 0
		session.findById("wnd[1]/tbar[0]/btn[0]").press			'Green check mark button
		
		If Not Session.findbyid("wnd[0]", False) Is Nothing Then
			ErrorMsg1 = Session.findbyid("wnd[2]/usr/txtMESSTXT1").text
			ErrorMsg2 = Session.findbyid("wnd[2]/usr/txtMESSTXT2").text
			StatusMsg = ErrorMsg1 & " " & ErrorMsg2
			ExcelSheet.Cells(Row,4).value = StatusMsg
			Session.findbyid("wnd[2]/tbar[0]/btn[0]").press		'By-pass message that material already maintained
			session.findById("wnd[1]").close	'Close the Org Levels pop-up screen			
		Else 
			Session.findById("wnd[0]/tbar[0]/btn[11]").press		'SAVE button
		
			StatusMsg = session.findById("wnd[0]/sbar").text
			ExcelSheet.Cells(Row,4).Value = StatusMsg
		End If
	End If
	
	Row = Row + 1

Loop

ExcelApp.Quit

Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing