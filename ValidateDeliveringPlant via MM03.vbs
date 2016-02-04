'****************************************************************************************
'
' Goal: Determines the Delivering Plant for a material number if the material number exists
' Actual function: Captures the status bar message for a specific part number
'
'****************************************************************************************

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
Dim Row,DelivPlant,StatusText


Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open("D:\ScriptData\ValidateDeliveringPlant_Data.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets(1)

'User is prompted to enter first row of Excel spreadsheet to be read
Row=InputBox("Row to start at")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm03"
session.findById("wnd[0]").sendVKey 0

' ** Do Until Loop will execute until the last row (a blank row) is found **
Do Until ExcelSheet.Cells(Row,1).Value = ""

	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,6).Value 'material number
	'session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 12
	session.findById("wnd[0]").sendVKey 0
	
'	If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").setFocus Then
'		DelivPlant = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").text
'		session.findById("wnd[1]/tbar[0]/btn[6]").press
'		session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,1).Value 'plant
'		session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = ExcelSheet.Cells(Row,2).Value 'storage loc
'		Session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = "5013"
'		session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "01"
'		session.findById("wnd[1]/tbar[0]/btn[0]").press
'		ExcelSheet.Cells(Row,9).Value = DelivPlant
'		session.findById("wnd[0]/tbar[0]/btn[3]").press  'green arrow back
'	Else
'		session.findById("wnd[0]").sendVKey 0
		StatusText = session.findById("wnd[0]/sbar").text
		ExcelSheet.Cells(Row,9).Value = StatusText
'	End If
	
	Row = Row + 1
	
Loop

ExcelApp.Quit

Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing
