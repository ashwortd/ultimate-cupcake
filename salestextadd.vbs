Dim row,status,statusbar
Dim ExcelApp,ExcelWorkbook,ExcelSheet

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
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")

row=InputBox("What row do you want to start on?","Starting Point")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0
Do While ExcelSheet.Cells(Row,1).Value <>""
	Call Main
	row =row +1
Loop
ExcelWorkbook.Close(True)
Set ExcelApp = Nothing
Set ExcelWorkbook = Nothing
Set ExcelSheet = Nothing

Sub main
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[19]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(5).selected = true
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ""
session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = "5013"
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "01"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP08/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2121/cntlLONGTEXT_VERTRIEBS/shellcont/shell").text = "Based on our abrasion testing, Alstom’s 90% Alumina DynawearTM  tile will have 2 to 3 times the life compared to ductile iron. Alstom’s DynawearTM coal pipe includes the following features:" + vbCr + vbCr + "1. High density 90% Alumina engineered tiles." + vbCr + "2. Mechanically interlocked tiles with staggered joint layout ensuring"+ vbCr +"   optimum wear and integrity." + vbCr + "3. Original ID of pipe is maintained." + vbCr + "4. Alstom supplies elbows with a maximum miter angle of 22.5 deg,"+ vbCr +"   our 90 deg elbows have 5 miter sections. This design reduces"+ vbCr +"   premature wear and minimizes pressure drop. " + vbCr + "5. Carbon steel casing manufactured with full penetration welds to"+ vbCr +"   meet structural and pressure requirements." + vbCr + vbCr + "For additional information please contact your Alstom "+ vbCr +"Customer Service Manager or Representative for product information"+ vbCr +"bulletin 212." + vbCr + "(ST03885 3/22/13 - MJS/CM)"+ vbCr + vbCr +"Please be aware that while the inside diameter of coal pipe bends are maintained, opting for 1” ceramic-lined pipe may result in a larger outside diameter than present with cast or un-lined fabricated steel piping.  This may affect existing clearance and support infrastructure.  Weight is comparable when going from cast to ceramic-lined, though it increases when going from fabricated un-lined to ceramic-lined.  Fit up and installation, including any adjustments necessary due to the inclusion of ceramics, are the sole responsibility of the purchaser."
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP08/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2121/cntlLONGTEXT_VERTRIEBS/shellcont/shell").setSelectionIndexes 0,0
session.findById("wnd[0]/tbar[0]/btn[11]").press
ExcelSheet.Cells(Row,5).Value = session.findById("wnd[0]/sbar").Text
End Sub
