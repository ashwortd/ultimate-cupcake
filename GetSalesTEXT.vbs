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
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet2")
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
row=InputBox("Which row to start on?","Starting Point")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm03"
session.findById("wnd[0]").sendVKey 0
Do While ExcelSheet.Cells(Row,1).Value<>""
Call salestextread
Loop 

Sub salestextread
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[19]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(7).selected = true
'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,5]").setFocus
'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,5]").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
ExcelSheet.Cells(Row,2).Value=session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP08/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2121/cntlLONGTEXT_VERTRIEBS/shellcont/shell").text 
'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP08/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2121/cntlLONGTEXT_VERTRIEBS/shellcont/shell").setSelectionIndexes 41,41
session.findById("wnd[0]/tbar[0]/btn[3]").press
row=row+1
End Sub
excelworkbook.Close(True)
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
Set ExcelApp = Nothing
Set ExcelWorkbook = nothing
Set ExcelSheet = Nothing
