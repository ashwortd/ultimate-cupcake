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
Set ExcelWorkbook = ExcelApp.Workbooks.Open("G:\ScriptData\Customer Industry Change.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
Row=InputBox("Row to start at")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nxd02"
session.findById("wnd[0]").sendVKey 0
Call modifyindustry

Sub modifyindustry()
If ExcelSheet.Cells(Row,1).Value = "" Then
	Call endscript
End If
session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").caretPosition = 10
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPMF02D:7123/ctxtKNA1-BRSCH").text = ExcelSheet.Cells(Row,7).Value
session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPMF02D:7123/ctxtKNA1-BRSCH").setFocus
session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:SAPLATAB:0200/subAREA2:SAPMF02D:7123/ctxtKNA1-BRSCH").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
ExcelSheet.Cells(Row,8).Value = session.findById("wnd[0]/sbar").Text
Call addrow
End Sub
Sub addrow()
Row=Row+1
Call modifyindustry
End Sub

Sub endscript()
ExcelApp.Quit
Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing
WScript.Quit
End Sub