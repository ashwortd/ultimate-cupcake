'**********Created by Derek Ashworth
'this script will zero out the safety stock and minimum lot size 
'on material master MRP1 and MRP2
'*2/4/2014
'*****************
'Check for SAP login and make connection
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
'------------------------
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
'***define variables and open excel workbook and set sheet
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim Row
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
'--------------------
'*****get starting row
Row=InputBox("Which row would you like to start on?","Starting Row")
If Row =False Then
	MsgBox("You have cancelled the Script")
			ExcelWorkbook.Close(True)
			ExcelApp.Quit
			Set ExcelApp=Nothing
			Set ExcelWoorkbook=Nothing
			Set ExcelSheet=Nothing
			WScript.ConnectObject session,     "off"
   			WScript.ConnectObject application, "off"
			WScript.Quit
End If

'**Set loop to repeat subroutine until there are no more material numbers
Do
Call Main
Row=Row+1
Loop While ExcelSheet.Cells(Row,1).Value <>""
'****
'*******Close out connections and script
	MsgBox("You have completed the Script")
			ExcelWorkbook.Close(True)
			ExcelApp.Quit
			Set ExcelApp=Nothing
			Set ExcelWoorkbook=Nothing
			Set ExcelSheet=Nothing
			WScript.ConnectObject session,     "off"
   			WScript.ConnectObject application, "off"
			WScript.Quit

Sub Main
'*******Maximize SAP window and procede to Material Master Change Screen
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0
'**get part number from excel spreadsheet, based on starting row and column 1 or A
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]").sendVKey 0'**presses enter key
session.findById("wnd[1]/tbar[0]/btn[19]").press'***selects the clear all from the tab selection
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(9).selected = true'***Selects MRP 1
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(10).selected = true'***Selects MRP 2
session.findById("wnd[1]/tbar[0]/btn[0]").press'***selects the green check mark
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,2).Value'**selects the plant from column B in the excel sheet
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/txtMARC-BSTMI").text = "0"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2486/txtMARC-EISBE").text = "0"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press'**Selects the save button on the popup window
ExcelSheet.Cells(Row,5).Value = session.findById("wnd[0]/sbar").Text
End Sub

