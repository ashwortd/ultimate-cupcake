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
Dim Row,window1

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet2")
Row=InputBox("Starting Row?")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nxd03"
session.findById("wnd[0]").sendVKey 0
Do 
Call Main
Row=Row+1
Loop While ExcelSheet.Cells(Row,1).Value<>""
		ExcelWorkbook.Close(True)
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"
		WScript.Quit
	
Sub Main
session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[1]/usr/ctxtRF02D-BUKRS").text = "5000"
session.findById("wnd[1]/usr/ctxtRF02D-VKORG").text = "5013"
session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = "99"
session.findById("wnd[1]/usr/ctxtRF02D-SPART").text = "00"
session.findById("wnd[1]/usr/ctxtRF02D-SPART").setFocus
session.findById("wnd[1]/usr/ctxtRF02D-SPART").caretPosition = 2
session.findById("wnd[1]/tbar[0]/btn[0]").press
On Error Resume Next
	window1=session.findById("wnd[2]").text
	'MsgBox window1
	If window1="Error" Then
		window1="none"
		session.findById("wnd[2]/tbar[0]/btn[0]").press
		Exit Sub
	End If
	On Error Goto 0
session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/ctxt*KNVP-PARVW").text = "z3"
session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/ctxt*KNVP-PARVW").setFocus
session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/ctxt*KNVP-PARVW").caretPosition = 2
session.findById("wnd[0]").sendVKey 0
	
ExcelSheet.Cells(Row,2).value=session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,0]").text
ExcelSheet.Cells(Row,3).value=session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/txt*TPART-VTEXT[1,0]").text
ExcelSheet.Cells(Row,4).value=session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtRF02D-KTONR[2,0]").text
ExcelSheet.Cells(Row,5).value=session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/txtRF02D-VTXTM[3,0]").text
session.findById("wnd[0]/tbar[0]/btn[3]").press
End Sub