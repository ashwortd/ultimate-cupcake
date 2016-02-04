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
Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row,dbchannel,salesorg
Set ExcelApp = CreateObject("Excel.Application")

'************Ask for data file
Set objDialog = CreateObject("UserAccounts.CommonDialog")

objDialog.Filter = "VBScript Data Files |*.xls;*.xlsx|All Files|*.*"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\Scripts"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
    Wscript.Quit
End If
'*************
'**********Open data file
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)	
Set ExcelSheet = ExcelWorkbook.Worksheets(1)
'*******
'-------Set Starting Row
Row=InputBox("Row to start at")
'-------
'------- Set distribution channel and sales org
dbchannel=InputBox("Please enter distribution channel (01 or 99)")
salesorg=InputBox("Please enter sales organization (5013,5063,etc)")
Call addreplby

Sub addreplby
If ExcelSheet.Cells(Row,2).Value = "" Then
	MsgBox("The end has come")
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.Quit
	End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nvb11"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtD000-KSCHL").text = "z501"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/ctxtKOMGD-VKORG").text = salesorg
session.findById("wnd[0]/usr/ctxtKOMGD-VTWEG").text = dbchannel
On Error Resume Next
Session.findById("wnd[0]/usr/ctxtKOMGD-MATWA").text = ExcelSheet.Cells(Row,2).Value
	If Err.Number<>0 Then
		ExcelSheet.Cells(Row,6).Value = session.findById("wnd[0]/sbar").Text
		Row=Row+1
		Call addreplby
	End If
On Error Goto 0
Session.findById("wnd[0]/usr/ctxtMV13D-SUGRV").text = "0002"
session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKOMGD-KUNNR[0,0]").text = ExcelSheet.Cells(Row,3).Value
session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKONDD-SMATN[2,0]").text = ExcelSheet.Cells(Row,4).Value
session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY").columns.elementAt(3).width = 27
session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKONDD-SUGRD[5,0]").text = "0002"
session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKONDD-SUGRD[5,0]").setFocus
session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKONDD-SUGRD[5,0]").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
If dbchannel=("01")Then
	ExcelSheet.Cells(Row,6).Value = session.findById("wnd[0]/sbar").Text
Else ExcelSheet.Cells(Row,7).Value = session.findById("wnd[0]/sbar").Text
End If
Row=Row+1
Call addreplby
End Sub
