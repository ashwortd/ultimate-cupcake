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
Dim Row,Status2
Row=InputBox("Which Row to start on?")
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet4")
ExcelApp.Visible=True

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nxd03"
session.findById("wnd[0]").sendVKey 0

Do
session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = ExcelSheet.Cells(Row,7).Value
session.findById("wnd[1]/usr/ctxtRF02D-BUKRS").text = "5000"
session.findById("wnd[1]/usr/ctxtRF02D-VKORG").text = ""
session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = ""
session.findById("wnd[1]/usr/ctxtRF02D-SPART").text = ""
session.findById("wnd[1]/usr/ctxtRF02D-BUKRS").setFocus
session.findById("wnd[1]/usr/ctxtRF02D-BUKRS").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press
ExcelSheet.Cells(Row,2).Value=session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0313/ctxtADDR1_DATA-COUNTRY").text
session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7111/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0313/ctxtADDR1_DATA-COUNTRY").caretPosition = 2
session.findById("wnd[0]/tbar[0]/btn[3]").press
Row=Row+1
Status2=ExcelSheet.cells(Row,4)
Loop Until Status2=""
	WScript.ConnectObject session,     "off"
    WScript.ConnectObject application, "off"
    WScript.Quit
    