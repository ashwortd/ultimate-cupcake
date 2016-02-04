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
Dim ApplicationPath,CredentialsPath,FilePath,bubba,WndTitle
Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row
Set ExcelApp = CreateObject("Excel.Application")
'Next line sets the location of the excel spreadsheet
Set ExcelWorkbook = ExcelApp.Workbooks.Open("D:\Documents and Settings\dma02\Desktop\Kelly\remaining.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
ExcelApp.Visible(True)
Row=InputBox("Row to start at")
If Row = False Then
		ExcelWorkbook.Close(True)
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"
		WScript.Quit
End If
Session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0
Do	
Call Main
Row=Row+1
Loop While ExcelSheet.Cells(Row,1).Value <>""

ExcelWorkbook.Close(True)
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"
		WScript.Quit


Sub Main
Session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0
bubba=ExcelSheet.Cells(Row,5).Value
Session.findbyid("wnd[0]/usr/ctxtRMMG1-MATNR").text = bubba


session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[19]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(10).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(11).selected = true
Session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,2).Value
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = ""
session.findById("wnd[1]").sendVKey 0
On Error Resume Next
WndTitle=Session.findbyid("wnd[2]").text
If WndTitle = "Error" Then
	ExcelSheet.Cells(Row,13).Value = Session.findbyid("wnd[2]/usr/txtMESSTXT1").text&" "&Session.findbyid("wnd[2]/usr/txtMESSTXT2").text
	WndTitle="null"
	Session.findbyid("wnd[2]/tbar[0]/btn[0]").press
	Session.findbyid("wnd[1]/tbar[0]/btn[12]").press
	Exit Sub
End If
On Error Goto 0 
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-PLIFZ").text = ExcelSheet.Cells(Row,9).Value
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").text = ExcelSheet.Cells(Row,9).Value+1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
ExcelSheet.Cells(Row,13).Value = session.findById("wnd[0]/sbar").Text
End Sub
