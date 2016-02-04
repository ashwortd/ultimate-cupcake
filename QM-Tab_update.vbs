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
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim Row,status1
'************Ask for data file
Set objDialog = CreateObject("UserAccounts.CommonDialog")

objDialog.Filter = "VBScript Data Files |*.xls;*.xlsx;*.xlsm|All Files|*.*"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\Scripts"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
    Wscript.Quit
'Else
'    Wscript.Echo objDialog.FileName
End If
'****************
Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible=True
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
Row=InputBox("Starting Row")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0

Do
	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.cells(Row,5)
	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 9
	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[1]").sendVKey 0
	session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.cells(Row,4)
	session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").caretPosition = 4
	session.findById("wnd[1]").sendVKey 0
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/chkMARA-QMPUR").selected = true
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/chkMARA-QMPUR").setFocus
	session.findById("wnd[0]").sendVKey 0
	status1=session.findbyId("wnd[0]/sbar").text
		If status1="Plants exist in which you have not specified a control key" Then
			session.findById("wnd[0]").sendVKey 0
			status1="none"
		End if
	session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
	ExcelSheet.Cells(Row,10).Value = session.findById("wnd[0]/sbar").Text
	Row=Row+1
Loop Until ExcelSheet.cells(Row,1)=""

ExcelWorkbook.Close(True)
Set ExcelApp=Nothing
Set ExcelWorkbook=Nothing
Set ExcelSheet=Nothing
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit
