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
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim Row,Statusbar,wndstat
Row=InputBox("Which excel Row would you like to start with?","Starting Position")

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
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")

Call Main
Sub Main
Call MM02setup
session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM02"
session.findById("wnd[0]/tbar[0]/btn[0]").press
Do While ExcelSheet.cells(Row,1).value<>""
wndstat="none"
Call MM02
Row=Row+1
Loop
Call Cleanup

End Sub
Sub MM02setup
session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM02"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findbyid("wnd[0]/usr/ctxtRMMG1-MATNR").text="000-02-01-03-13"
session.findbyid("wnd[0]/tbar[0]/btn[0]").press
session.findbyid("wnd[1]/tbar[0]/btn[19]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(6).selected = true
session.findbyid("wnd[1]/tbar[0]/btn[14]").press
session.findbyid("wnd[1]/tbar[0]/btn[0]").press
End Sub
Sub MM02
session.findbyid("wnd[0]/usr/ctxtRMMG1-MATNR").text=ExcelSheet.cells(Row,1).value
session.findbyid("wnd[0]/tbar[0]/btn[0]").press
Statusbar=session.findbyid("wnd[0]/sbar").text
Statusbar=Left(statusbar,12)
If Statusbar="The material" Then
	ExcelSheet.Cells(Row,6).Value = session.findById("wnd[0]/sbar").Text
	Exit Sub
End if
If Statusbar<>"" Then
	WScript.Sleep(500)
	session.findbyid("wnd[0]/usr/ctxtRMMG1-MATNR").text=ExcelSheet.cells(Row,1).value
	session.findbyid("wnd[0]/tbar[0]/btn[0]").press
End if
	
session.findbyid("wnd[1]/tbar[0]/btn[0]").press
If session.findbyid("wnd[0]/sbar").text ="Select at least one view" Then
	ExcelSheet.Cells(Row,6).Value = session.findById("wnd[0]/sbar").Text
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM02"
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	Exit Sub
End If
session.findbyid("wnd[1]/usr/ctxtRMMG1-WERKS").text="500B"
session.findbyid("wnd[1]/tbar[0]/btn[0]").press
On Error Resume Next
wndstat=session.findbyid("wnd[2]").text
If wndstat="Error" Then
	ExcelSheet.Cells(Row,6).Value = session.findbyid("wnd[2]/usr/txtMESSTXT1").text &" "&session.findbyid("wnd[2]/usr/txtMESSTXT2").text
	session.findbyid("wnd[2]/tbar[0]/btn[0]").press
	session.findbyid("wnd[1]/tbar[0]/btn[12]").press
	Exit Sub
End If
On Error Goto 0
session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").text="ELM"
session.findbyid("wnd[0]/tbar[0]/btn[0]").press
session.findbyid("wnd[1]/usr/btnSPOP-OPTION1").press
statusbar3=session.findById("wnd[0]/sbar").Text
statusbar3=Left(statusbar3,8)

ExcelSheet.Cells(Row,6).Value = session.findById("wnd[0]/sbar").Text
End Sub
	
Sub Cleanup
Set ExcelApp = Nothing
Set ExcelWorkbook = Nothing
Set ExcelSheet = Nothing
MsgBox("The requested process has been completed." & chr(13) & chr(13) & "Thank you.")					
	WScript.ConnectObject session,     "off"
	WScript.ConnectObject application, "off"
	WScript.Quit
End Sub
