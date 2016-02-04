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
Dim xlRow,itmRow
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
Set ExcelSheet = ExcelWorkbook.Worksheets(1)
xlRow=InputBox("Starting Row in Excel Sheet?","Starting Point")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nva02"
session.findById("wnd[0]/tbar[0]/btn[0]").press
Do
Call Main
xlRow=xlRow+1
Loop While ExcelSheet.cells(xlRow,1)<>""

Sub Main
session.findbyid("wnd[0]/usr/ctxtVBAK-VBELN").text=ExcelSheet.cells(xlRow,3)
session.findbyid("wnd[0]/tbar[0]/btn[0]").press
On Error Resume Next
If session.findybyid("wnd[1]").text <>"" Then
	session.findbyid("wnd[1]/tbar[0]/btn[0]").press
End If
On Error Goto 0
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,1]").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,1]").caretPosition = 5
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KOAN").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,1]").text = "pr00"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,1]").text = "0"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,1]").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,1]").caretPosition = 16
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,2]").text = "ycmc"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").text = "0"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,2]").caretPosition = 16
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KONY").press
session.findById("wnd[1]/usr/lbl[6,5]").setFocus
session.findById("wnd[1]/usr/lbl[6,5]").caretPosition = 12
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[1]").close
session.findById("wnd[2]/usr/btnBUTTON_2").press

