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
Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row
Set ExcelApp = CreateObject("Excel.Application")
'Next line sets the location of the excel spreadsheet
Set ExcelWorkbook = ExcelApp.Workbooks.Open("G:\ScriptData\ASM-Change.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
Row=InputBox("Row to start at")
Session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nxd02"
session.findById("wnd[0]").sendVKey 0
Call startread

Sub startread()
If ExcelSheet.Cells(Row,1).Value = "" Then
	Call endscript
End If
On Error Resume Next	
session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[1]/usr/ctxtRF02D-VKORG").text = "5013"
session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = "01"
session.findById("wnd[1]/usr/ctxtRF02D-SPART").text = "00"
session.findById("wnd[1]/usr/ctxtRF02D-SPART").setFocus
session.findById("wnd[1]/usr/ctxtRF02D-SPART").caretPosition = 2
session.findById("wnd[1]/tbar[0]/btn[0]").press
ExcelSheet.Cells(Row,2).Value = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,0]").text 
ExcelSheet.Cells(Row,3).Value = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,1]").text
ExcelSheet.Cells(Row,4).Value = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,2]").text
ExcelSheet.Cells(Row,5).Value = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,3]").text
ExcelSheet.Cells(Row,6).Value = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,4]").text
ExcelSheet.Cells(Row,7).Value = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,5]").text
ExcelSheet.Cells(Row,8).Value = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,6]").text
ExcelSheet.Cells(Row,9).Value = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,7]").text
ExcelSheet.Cells(Row,10).Value = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,8]").text
ExcelSheet.Cells(Row,11).Value = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,9]").text
ExcelSheet.Cells(Row,12).Value = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,10]").text
ExcelSheet.Cells(Row,13).Value = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,11]").text
ExcelSheet.Cells(Row,14).Value = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,12]").text
ExcelSheet.Cells(Row,15).Value = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,13]").text
ExcelSheet.Cells(Row,16).Value = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,14]").text
ExcelSheet.Cells(Row,17).Value = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,15]").text
Session.findById("wnd[0]/tbar[0]/btn[11]").press
ExcelSheet.Cells(Row,20).Value = session.findById("wnd[0]/sbar").Text
On Error Goto 0
Call addrow()
End Sub

Sub addrow()
Row=Row+1
Call startread
End Sub

Sub endscript()
ExcelApp.Quit
Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing
WScript.Quit
End Sub














Session.findById("wnd[0]").sendVKey 0
