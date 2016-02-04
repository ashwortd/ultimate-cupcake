If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   On Error Resume next
   Set connection = application.Children(0)
   If Err.Number<>0 Then
   	MsgBox("You are not connected to PMx, please connect and try again")
   	On Error Goto 0
   	WScript.Quit
   End If
Dim ExcelApp,ExcelSheet,ExcelWorkbook
Dim Row,WndTitle,StatBar
  	
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

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
Row=InputBox("Which row to start on?")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nva22"
session.findById("wnd[0]").sendVKey 0
WndTitle="none"
Do
Call Main
Row=Row+1
Loop Until ExcelSheet.Cells(Row,1).Value=""
Call CleanUp


Sub Main
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = ExcelSheet.Cells(Row,1).Value
'session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 8
session.findById("wnd[0]").sendVKey 0
On Error Resume next
WndTitle=session.findbyid("wnd[1]").text
On Error Goto 0
If WndTitle="Information" Then
	session.findById("wnd[1]").sendVKey 0
	WndTitle="none"
End If
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
'session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\10").select
'session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR5").key = "OCA"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR5").setFocus
session.findById("wnd[0]/tbar[0]/btn[11]").press
On Error Resume next
WndTitle=session.findbyid("wnd[1]").text
On Error Goto 0
If WndTitle="Information" Then
	session.findById("wnd[1]").sendVKey 0
	'session.findById("wnd[1]").close
	WndTitle="none"
End If
On Error Resume Next
WndTitle=session.findbyid("wnd[1]").text
On Error Goto 0
If WndTitle="Workflow Selection" Then
	session.findById("wnd[1]").close
	session.findById("wnd[2]/usr/btnBUTTON_2").press
	WndTitle="none"
End If
ExcelSheet.Cells(Row,2).Value = session.findById("wnd[0]/sbar").Text
StatBar=session.findById("wnd[0]/sbar").Text
If Left(StatBar,4)="Main" Then
	session.findbyid("wnd[0]/tbar[0]/btn[12]").press
	session.findbyid("wnd[1]/usr/btnSPOP-OPTION1").press
	StatBar="none"
End if
End Sub

Sub CleanUp
ExcelWorkbook.Close(True)
Set ExcelSheet = Nothing
Set ExcelWorkbook=Nothing
Set ExcelApp=Nothing
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit
End Sub

