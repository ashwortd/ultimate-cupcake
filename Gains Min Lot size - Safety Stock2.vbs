'Option Explicit()
Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim row,num

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
Set ExcelSheet = ExcelWorkbook.Worksheets("SS or ORQ")
row = InputBox("Which row would you like to start?")
Call Main
Call WrapUp
Sub Main
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
Session.findById("wnd[0]").sendVKey 0
Do While ExcelSheet.Cells(row,1).Value<>""	
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(row,1).Value
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[19]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(9).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(10).selected = true
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(row,2).Value
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = ""
session.findById("wnd[1]/tbar[0]/btn[0]").press
num = Right(Left(session.ActiveWindow.GuiFocus.ID,50),2)
If num <> 12 Then
	ExcelSheet.Cells(row,7).Value = "Part requires attention"
	row=row+1
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
	Session.findById("wnd[0]").sendVKey 0
	num = 0
	Exit Do
	End If
	
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/txtMARC-BSTMI").text = ExcelSheet.Cells(row,4).Value
session.findById("wnd[0]/tbar[0]/btn[0]").press
checkOne = session.findById("wnd[0]/sbar").Text
If checkOne ="Reorder point considered only with reorder point or time-phased planning" Then
	Session.findById("wnd[0]").sendVKey 0
End if
num = Right(Left(session.ActiveWindow.GuiFocus.ID,50),2)
If num <> 13 Then
	ExcelSheet.Cells(row,7).Value = "Part requires attention"
	row=row+1
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
	Session.findById("wnd[0]").sendVKey 0
	num = 0
	Exit Do
	End If
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2486/txtMARC-EISBE").text = ExcelSheet.Cells(row,3).Value
session.findById("wnd[0]/tbar[0]/btn[11]").press
ExcelSheet.Cells(row,5).Value = session.findById("wnd[0]/sbar").Text
row=row+1
Loop
End Sub

Sub WrapUP
MsgBox("Thanks for your Input")
Application.DisplayAlerts = False
ExcelWorkbook.Save
ExcelWorkbook.Close
ExcelApp.Quit
Application.DisplayAlerts = True
ExcelApp.ActiveWorkbook.Close
Set ExcelApp = Nothing
Set ExcelWorkbook=Nothing
Set ExcelSheet=Nothing
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit
End Sub
