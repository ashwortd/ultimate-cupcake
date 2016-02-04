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
Dim row
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets("No Spend - All Orgs")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nzmk05"
session.findById("wnd[0]").sendVKey 0
row=InputBox("Which row would you like to start on?")
Do
Call Main
row=row+1
Loop Until ExcelSheet.Cells(Row,1).Value=""
Call Cleanup

Sub Main
session.findById("wnd[0]/usr/ctxtP_LIFNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/ctxtP_EKORG").text = ExcelSheet.Cells(Row,3).Value
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/chkGW_COCKPIT-BLOCKPMX").selected = false
session.findById("wnd[0]/usr/chkGW_COCKPIT-BLOCKPMX").setFocus
session.findById("wnd[0]/tbar[0]/btn[11]").press
ExcelSheet.Cells(Row,27).Value = session.findById("wnd[0]/sbar").Text
End Sub

Sub Cleanup
	MsgBox("Script Complete")
	ExcelWorkbook.Close(True)
	ExcelApp.Quit
	Set ExcelApp=Nothing
	Set ExcelWoorkbook=Nothing
	Set ExcelSheet=Nothing
    WScript.ConnectObject session,     "off"
    WScript.ConnectObject application, "off"	
	WScript.Quit
End sub