If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
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
'Define global variables
Dim ex,wb,ws,windowtitl
Dim Row,SAPRow,success

Call Main
Do 
Call AddPurchOrg
Loop While ws.cells(Row,1).value <>""
		MsgBox("The end has come")
		ex.Quit
		Set ex=Nothing
		Set wb=Nothing
		Set ws=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"
		WScript.Quit
Sub Main

Set ex = WScript.CreateObject("Excel.Application")
ex.Visible=False

Set objDialog = CreateObject("UserAccounts.CommonDialog")
objdialog.Filter = "VBSctipt Data Files |*.xls;*.xlsx;*.xlsm;|All Files|*.*"
objdialog.FilterIndex=1
objdialog.InitialDir="C:\Scripts"
intResult = objdialog.ShowOpen
If intResult=0 Then
	WScript.Quit
End If

Set wb = ex.Workbooks.Open(objdialog.FileName)
Set ws = wb.Sheets(wb.ActiveSheet.Name)
Row = InputBox("Which row would you like to start on?","Starting Point")

session.findById("wnd[0]/tbar[0]/okcd").text = "/NXK01"
session.findById("wnd[0]").sendVKey 0
windowtitl=session.findbyid("wnd[0]/titl").text
End Sub

Sub AddPurchOrg


session.findById("wnd[0]/usr/ctxtRF02K-LIFNR").text = ws.cells(Row,1).value
session.findById("wnd[0]/usr/ctxtRF02K-BUKRS").text = "5000"
session.findById("wnd[0]/usr/ctxtRF02K-EKORG").text = "US63"
session.findById("wnd[0]/usr/ctxtRF02K-REF_LIFNR").text = ws.cells(Row,1).value
session.findById("wnd[0]/usr/ctxtRF02K-REF_BUKRS").text = "5000"
session.findById("wnd[0]/usr/ctxtRF02K-REF_EKORG").text = "US44"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
If session.findbyId("wnd[0]/titl").text= windowtitl then
	ws.Cells(Row,2).Value = session.findById("wnd[0]/sbar").Text
Else
	ws.Cells(Row,2).Value ="Error: "&session.findById("wnd[0]/sbar").Text
	Row = Row+1
	session.findById("wnd[0]/tbar[0]/okcd").text = "/NXK01"
	session.findById("wnd[0]").sendVKey 0
	Exit Sub
End If
Row=Row+1
End Sub

