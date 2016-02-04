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
Dim Row,statuserr,status2

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet2")
ExcelApp.Visible=True
Row=InputBox("Row?","Starting Point")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nxd02"
session.findById("wnd[0]").sendVKey 0
statuserr="none"
Do
Call Main
Row=Row+1
statuserr="none"
Loop While excelsheet.cells(row,1)<>""
If ExcelSheet.Cells(Row,1).Value=("") Then
		'Call endscript
		MsgBox("The end has come")
		excelworkbook.Close(True)
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
	    WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"
		WScript.Quit
	
	End If


Sub Main
session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = excelsheet.cells(row,1)
session.findById("wnd[1]/usr/ctxtRF02D-BUKRS").text = "5000"
session.findById("wnd[1]/usr/ctxtRF02D-VKORG").text = "5013"
session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = "01"
session.findById("wnd[1]/usr/ctxtRF02D-SPART").text = "00"
session.findById("wnd[1]/usr/ctxtRF02D-SPART").setFocus
session.findById("wnd[1]/usr/ctxtRF02D-SPART").caretPosition = 2
session.findById("wnd[1]").sendVKey 0
On Error Resume next
statuserr=session.findbyid("wnd[2]").text 
On Error Goto 0
status2="none"
If statuserr="Error" Then
   ExcelSheet.cells(row,4)="Error - Not in 5013 01 00"
   session.findbyid("wnd[2]/tbar[0]/btn[0]").press
   Exit Sub
End If
session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0313/txtSZA1_D0100-SMTP_ADDR").text=ExcelSheet.cells(row,2).value
session.findById("wnd[0]/tbar[0]/btn[11]").press
status2=session.findbyid("wnd[0]/sbar").text
If status2="The use of the last 25 characters in field STREET is restricted (52 of 60)" Then
session.findById("wnd[0]").sendVKey 0
End if
ExcelSheet.cells(row,3)=session.findbyid("wnd[0]/sbar").text
End sub	