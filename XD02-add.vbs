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
Dim Row, PFExist,Status1
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
Set ExcelSheet = ExcelWorkbook.Worksheets("DM_Assignment")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nxd02"
session.findById("wnd[0]").sendVKey 0
Row=InputBox("RowNumber?")
Status1="none"
Do
Call main
row=row+1
Loop until ExcelSheet.Cells(Row,1).Value=""
		MsgBox("The end has come")
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.Quit



Sub main
session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[1]/usr/ctxtRF02D-VKORG").text = "5013"
session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = "01"
session.findById("wnd[1]/usr/ctxtRF02D-SPART").text = "00"
session.findById("wnd[1]/tbar[0]/btn[0]").press
On Error Resume next
Status1=session.findbyid("wnd[2]").text
If Status1="Error" Then
	ExcelSheet.Cells(Row,4).Value = session.findbyid("wnd[2]/usr/txtMESSTXT1").text
	ExcelSheet.Cells(Row,5).Value = session.findbyid("wnd[2]/usr/txtMESSTXT2").text
	session.findbyid("wnd[2]").close
	Status1="none"
	Exit Sub
End If

session.findById("wnd[0]/tbar[1]/btn[25]").press
session.findById("wnd[0]/tbar[1]/btn[27]").press
session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05").select
On Error Goto 0
session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/ctxt*KNVP-PARVW").text="ZF"
session.findById("wnd[0]").sendVKey 0
PFexist = session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,0]").text
'MsgBox(PFexist)
If PFexist ="ZF" Then
	session.findbyid("/app/con[0]/ses[0]/wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtRF02D-KTONR[2,0]").text=ExcelSheet.Cells(Row,3).Value
	
elseif PFexist <>"ZF" Then
	session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN").verticalScrollbar.position = 19
	session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,4]").text="ZF"
	session.findbyid("/app/con[0]/ses[0]/wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtRF02D-KTONR[2,4]").text=ExcelSheet.Cells(Row,3).Value
End If
'session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/ctxt*KNVP-PARVW").text = "zg"
'session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/txt*RF02D-KTONR").text = "111200010"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
ExcelSheet.Cells(Row,4).Value = session.findById("wnd[0]/sbar").Text
End Sub
 