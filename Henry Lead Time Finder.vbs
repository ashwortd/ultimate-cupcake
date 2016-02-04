Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row,SapStatus,window1status,SapWindow,LeadTime

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
ExcelApp.Visible=True
Row=InputBox("Row to start at")
Session.findById("wnd[0]").maximize
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm03"
session.findById("wnd[0]").sendVKey 0
Call Main

		MsgBox("The end has come")
		ExcelWorkbook.Close(True)
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"	
		WScript.Quit


Sub Main
Do
Session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text=ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/tbar[0]/btn[0]").press
SapStatus=session.findById("wnd[0]/sbar").Text
If SapStatus <> "" Then
	ExcelSheet.Cells(Row,3).Value=SapStatus
	Call reset
End If	
session.findById("wnd[1]/tbar[0]/btn[0]").press
SapStatus=session.findById("wnd[0]/sbar").Text
If SapStatus <> "" Then
	ExcelSheet.Cells(Row,3).Value=SapStatus
	Call reset2
End If
Session.findbyid("wnd[1]/usr/ctxtRMMG1-WERKS").text="500B"
Session.findbyid("wnd[1]/tbar[0]/btn[0]").press
On Error Resume next
SapWindow = Session.findbyid("wnd[2]").text
On Error Goto 0
If SapWindow="Error" Then
	'ExcelSheet.Cells(Row,3).Value=Session.findbyid("wnd[2]/usr/txtMESSTXT1").text &" "&Session.findbyid("wnd[2]/usr/txtMESSTXT2").text
	Session.findbyid("wnd[2]/tbar[0]/btn[0]").press
	Session.findbyid("wnd[1]/usr/ctxtRMMG1-WERKS").text="500J"
	Session.findbyid("wnd[1]/tbar[0]/btn[0]").press
	SapWindow="None"
End If
On Error Resume next
SapWindow = Session.findbyid("wnd[2]").text
On Error Goto 0
If SapWindow="Error" Then
	ExcelSheet.Cells(Row,3).Value=Session.findbyid("wnd[2]/usr/txtMESSTXT1").text &" "&Session.findbyid("wnd[2]/usr/txtMESSTXT2").text
	Session.findbyid("wnd[2]/tbar[0]/btn[0]").press
	Session.findbyid("wnd[1]/tbar[0]/btn[12]").press
	Call reset
End If
LeadTime=Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").text
ExcelSheet.Cells(Row,2).Value=LeadTime
LeadTime="Dude"
Session.findbyid("wnd[0]/tbar[0]/btn[3]").press
Row=Row+1
Loop Until ExcelSheet.Cells(Row,1).Value=""
End sub

Sub reset
	Row=Row+1
	Call main
End Sub

Sub reset2
Session.findbyid("wnd[1]/tbar[0]/btn[12]").press
Row=Row+1
Call Main
End Sub
