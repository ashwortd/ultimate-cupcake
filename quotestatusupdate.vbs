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
Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row
Set ExcelApp = CreateObject("Excel.Application")
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

'**********Open data file
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")'Specify tab in worksheet
'********
Row=InputBox("Row to start at")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nva22"
session.findById("wnd[0]").sendVKey 0

While ExcelSheet.Cells(Row,1).Value>""
Call updatequote
Wend
'close the script when it reaches a blank row
		WScript.Sleep(5000)'pause to let script catch up
		MsgBox("The end has come")
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.Quit
Sub updatequote		
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[13]").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR5").key = "WON"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR5").setFocus
session.findById("wnd[0]/tbar[0]/btn[11]").press
	If Not session.findById("wnd[1]/usr/txtSPOP-TEXTLINE1",false)Is Nothing Then '.text="Document Incomplete" Then
			Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
	End If
	If Not Session.findbyid("wnd[1]",False) Is Nothing Then
			session.findById("wnd[1]").close
			Session.findById("wnd[2]/usr/btnBUTTON_2").press
			'session.findById("wnd[1]").close
			'session.findById("wnd[2]/usr/btnBUTTON_2").press

	End If 
ExcelSheet.Cells(Row,12).Value = session.findById("wnd[0]/sbar").Text
Row=Row+1
End Sub


