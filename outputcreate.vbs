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

Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row,sapstatus,window1status
Set ExcelApp = CreateObject("Excel.Application")
'Next line sets the location of the excel spreadsheet
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
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets("ZSD_BACKLOG Export")
ExcelApp.Visible=True
Row=InputBox("Row to start at")
Session.findById("wnd[0]").maximize
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nva02"
session.findById("wnd[0]").sendVKey 0

Do While ExcelSheet.Cells(Row,1).Value <>""
Call CreateOutput
Loop

Sub CreateOutput
Session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = ExcelSheet.Cells(Row,4).Value
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
WScript.Sleep(500)
If session.findById("wnd[0]/sbar").Text <>"" Then
	ExcelSheet.cells(Row,8).Value = session.findById("wnd[0]/sbar").Text
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nva02"
	session.findById("wnd[0]").sendVKey 0
	Row=Row+1
	Exit Sub
End If
On Error Resume next
window1status= Session.findbyid("wnd[1]").text
If window1status="Help - Sales Order change: Input" Then
	Session.findbyid("wnd[1]/tbar[0]/btn[5]").press
	window1status="none"
	ExcelSheet.cells(Row,8).Value ="Can only Display"
	Row=Row+1
	Exit Sub
End If
On Error Goto 0
session.findById("wnd[0]/mbar/menu[3]/menu[11]/menu[0]/menu[0]").select
session.findById("wnd[0]/tbar[1]/btn[2]").press
ExcelSheet.cells(Row,8).Value = session.findById("wnd[0]/sbar").Text
sapstatus = session.findById("wnd[0]/sbar").Text
If session.findById("wnd[0]/sbar").Text<>"" Then
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nva02"
	session.findById("wnd[0]").sendVKey 0
	Row = Row+1
	Exit Sub
End if
session.findById("wnd[0]/usr/chkNAST-DIMME").selected = false
session.findById("wnd[0]/usr/chkNAST-DELET").selected = true
session.findById("wnd[0]/usr/chkNAST-DELET").setFocus
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[1]/btn[5]").press
session.findById("wnd[0]/usr/cmbNAST-VSZTP").key = "4"
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
On Error Resume Next
If Session.findbyid("wind[1]/titl").text="Workflow Selection" Then
	Session.findbyid("wnd[1]").close
	Session.findbyid("wnd[2]/usr/btnBUTTON_2").press
	ExcelSheet.cells(Row,8).Value = session.findById("wnd[0]/sbar").Text
End If
If Session.findbyid("wnd[1]/titl").text="Save Incomplete Document" Then
	Session.findbyid("wind[1]").close
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nva02"
	session.findById("wnd[0]").sendVKey 0
	Row=Row+1
	Exit Sub
End If

	
On Error Goto 0

WScript.Sleep(500)
ExcelSheet.cells(Row,8).Value = session.findById("wnd[0]/sbar").Text
sapstatus=0
Row=Row+1
End Sub
ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   	    WScript.ConnectObject application, "off"
	 	WScript.Quit

