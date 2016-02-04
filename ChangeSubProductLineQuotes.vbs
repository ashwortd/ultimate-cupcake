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
Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row,QtOd,tcode
Set ExcelApp = CreateObject("Excel.Application")
'Next line sets the location of the excel spreadsheet
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
Set ExcelWorkbook = ExcelApp.Workbooks.Open(objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
Row=InputBox("Row to start at")
QtOd=InputBox("Change Sub Product line for Quotes or Orders? (q or o)")
	If QtOd="q" Then
		tcode="/nva22"
	ElseIf QtOd="o" Then
	   	tcode="/nva02"
	Else 
	    MsgBox ("Invalid Sales document type")
	   	ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.Quit
	End If

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = tcode
session.findById("wnd[0]").sendVKey 0

While ExcelSheet.Cells(Row,1).Value <>""
	Call Changesubpl
Wend
MsgBox("The end has come")
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.Quit


Sub Changesubpl
Session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = ExcelSheet.Cells(Row,1).Value
Session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 8
session.findById("wnd[0]/usr/btnBT_SUCH").press
On Error Resume Next
If Not session.findById("wnd[1]/usr/txtMESSTXT1",false)Is Nothing Then
		Session.findById("wnd[1]").sendVKey 0
	End If
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
If Err.Number <>0 Then
	ExcelSheet.Cells(Row,5).Value = session.findById("wnd[0]/sbar").Text
		Row=Row+1
	Err.Clear
	Exit Sub
End If
On Error Goto 0 
'Select tab additional data B
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12").select
'Change sub product line code
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_PRLINE2").text = ExcelSheet.Cells(Row,4).Value
session.findById("wnd[0]").sendVKey 0
'session.findById("wnd[1]/usr/lbl[1,4]").setFocus
'session.findById("wnd[1]/usr/lbl[1,4]").caretPosition = 5
'session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
ExcelSheet.Cells(Row,5).Value = session.findById("wnd[0]/sbar").Text
Row=Row+1
End Sub
