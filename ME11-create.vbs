'*******ME11 Purchase Info Record Create
'*Created by Derek Ashworth
'*2/17/2014
'****************************************
'Check for Logon status and connect to GUI
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
'******************************************
'************Ask for data file
'Set objDialog = CreateObject("UserAccounts.CommonDialog")

'objDialog.Filter = "VBScript Data Files |*.xls;*.xlsx;*.xlsm|All Files|*.*"
'objDialog.FilterIndex = 1
'objDialog.InitialDir = "C:\Scripts"
'intResult = objDialog.ShowOpen
 
'If intResult = 0 Then
'    Wscript.Quit
'Else
'    Wscript.Echo objDialog.FileName
'End If
'*******************************************
'*Open Excel data file and set worksheet
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open ("O:\CustSvc\Parts\Pmx Scripting\Script Data\Purchase Info Record Datasheet.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
'*******************************************
'*Maximize PMx Window and enter TCode for Creat Purch info recorf
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nme11"
session.findById("wnd[0]").sendVKey 0
'*******************************************
Dim Row,StatusBar
Row=InputBox("Which row would you like to start with?","Create Purchase Info Records")
	If Row = False Then
		Call Endscript
	End If
Do	
Call Main
Row=Row+1
Loop While ExcelSheet.Cells(Row,1).Value <>""

Call Endscript

Sub Endscript
MsgBox("Script completed")
		ExcelWorkbook.Close(True)
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"
		WScript.Quit
End Sub

Sub Main
session.findById("wnd[0]/usr/ctxtEINA-LIFNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/ctxtEINA-MATNR").text = ExcelSheet.Cells(Row,2).Value
session.findById("wnd[0]/usr/ctxtEINE-EKORG").text = ExcelSheet.Cells(Row,3).Value
session.findById("wnd[0]/usr/ctxtEINE-WERKS").text = ExcelSheet.Cells(Row,4).Value
session.findbyid("wnd[0]/usr/radRM06I-NORMB").select
session.findById("wnd[0]/tbar[0]/btn[0]").press
StatusBar=session.findById("wnd[0]/sbar").Text
	If StatusBar <>"" Then
		Call ErrorSub
		StatusBar=""
		Exit Sub
	End If
session.findById("wnd[0]/usr/txtEINA-IDNLF").text = ExcelSheet.Cells(Row,5).Value
session.findById("wnd[0]/tbar[1]/btn[7]").press
session.findById("wnd[0]/usr/txtEINE-NORBM").text = ExcelSheet.Cells(Row,6).Value
session.findById("wnd[0]/usr/txtEINE-NETPR").text = ExcelSheet.Cells(Row,7).Value
session.findById("wnd[0]/tbar[0]/btn[11]").press
ExcelSheet.Cells(Row,8).Value = session.findById("wnd[0]/sbar").Text
End Sub

Sub ErrorSub
ExcelSheet.Cells(Row,8).Value= StatusBar
session.findById("wnd[0]/usr/ctxtEINA-LIFNR").text = ""
session.findById("wnd[0]/usr/ctxtEINA-MATNR").text = ""
session.findById("wnd[0]/usr/ctxtEINE-EKORG").text = ""
session.findById("wnd[0]/usr/ctxtEINE-WERKS").text = ""
session.findbyid("wnd[0]/usr/ctxtEINA-INFNR").text = ""
End Sub


