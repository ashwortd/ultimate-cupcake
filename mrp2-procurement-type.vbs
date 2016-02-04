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
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim strStatus
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open("g:\ScriptData\MRP2-Procurement-type.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets(1)
Row=InputBox("Row to start at")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0
Call partnumber
Sub partnumber
if ExcelSheet.Cells(Row,1).Value = "" Then 
	Call nomore
End if	
strStatus = 0
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,3).Value
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 16
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[19]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(9).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,9]").setFocus
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,9]").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[14]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "500j"
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = ""
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").setFocus
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").caretPosition = 0
session.findById("wnd[1]").sendVKey 0
strStatus = session.findById("wnd[0]/sbar").Text
If strStatus = "Material not fully maintained for this transaction/event" Then
	Call Statuswrite
End If	
'On Error Resume next
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").text = "f"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
Call Statuswrite
End Sub

Sub Statuswrite
	ExcelSheet.Cells(Row,7).Value = session.findById("wnd[0]/sbar").Text
	row=row+1
	Call partnumber
End Sub

Sub nomore
ExcelApp.Quit
Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing
End sub