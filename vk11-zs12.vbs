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
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim Row,StatusBar
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
Row=InputBox("Starting Row")
Do
Call Setup
	Do 
	Call Main
	Row=Row+1
	Loop While ExcelSheet.Cells(Row,1).Value <>""
Loop While ExcelSheet.Cells(Row,1).Value <>""
Call Cleanup

Sub Setup
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nvk11"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRV13A-KSCHL").text = "zs12"
session.findById("wnd[0]/tbar[0]/btn[0]").press
End Sub

Sub Main
session.findById("wnd[0]/usr/ctxtKOMG-BUKRS").text = "5000"
session.findById("wnd[0]/usr/ctxtKOMG-VKORG").text = "5013"
session.findById("wnd[0]/usr/ctxtKOMG-VTWEG").text = "01"
session.findById("wnd[0]/usr/ctxtKOMG-SPART").text = "00"
session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/ctxtKOMG-KUNWE[0,0]").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[4,0]").text = "5"
session.findById("wnd[0]").sendVKey 0
StatusBar=session.findById("wnd[0]/sbar").Text
StatusBar=Right(StatusBar,5)
If StatusBar ="exist" Then
	ExcelSheet.Cells(Row,3).Value = session.findById("wnd[0]/sbar").Text
	session.findbyid("wnd[0]/usr/btnFCODE_ENTF").press
	StatusBar="None"
	Exit Sub
End If
session.findById("wnd[0]/tbar[0]/btn[11]").press
On Error Resume next
If session.findbyid("wnd[1]").text <>"" Then
	session.findbyid("wnd[1]/tbar[0]/btn[5]").press
End If
On Error Goto 0		
ExcelSheet.Cells(Row,3).Value = session.findById("wnd[0]/sbar").Text
End Sub

Sub Cleanup
	ExcelWorkbook.Close(True)
	ExcelApp.Quit
	Set ExcelApp=Nothing
	Set ExcelWorkbook=Nothing
	Set ExcelSheet=Nothing
	MsgBox("Complete")
	WScript.ConnectObject session,     "off"
    WScript.ConnectObject application, "off"
	WScript.Quit
End Sub