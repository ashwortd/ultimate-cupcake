'*********************************
'* Connect to SAP                *
'*********************************
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
'**********************************
'* Define Variables               *
'**********************************
Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim messtxt,x,Row
'************Ask for data file
Set objDialog = CreateObject("UserAccounts.CommonDialog")

objDialog.Filter = "VBScript Data Files |*.xls;*.xlsx;*.xlsm|All Files|*.*"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\Scripts"
intResult = objDialog.ShowOpen

'close the SAP Scripting if cancel is selected 
If intResult = 0 Then
	 WScript.ConnectObject session,     "off"
   	 WScript.ConnectObject application, "off"
	 WScript.Quit
End If
'****************
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets("LOCS TO BE ADDED IN MAT MSTR")'****This is the name of the sheet in the Excel Workbook****
'********************************
'* Request to Proceed           *
'********************************
x = MsgBox ("This script will add the Storage Bin to the Plant data",vbOKCancel,"Information")
If x=2 Then
	ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   	    WScript.ConnectObject application, "off"
		WScript.Quit
End If

Row=InputBox("Row to start at")'***********Ask which row of Excel Spreadsheet to start on - sets value to 'Row' Variable
'*****************************
'*Next few lines set Defaults*
'*****************************
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = "RIC31348"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[19]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(12).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,12]").setFocus
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,12]").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[14]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0

'*********************************************
'* Create loop to procede to next row unless *
'* column 'a' is empty                       *
'*********************************************
Do While ExcelSheet.Cells(Row,1).Value <>""
	Call UpdateSBin
Loop
If ExcelSheet.Cells(Row,1).Value=("0") Then
		MsgBox("The end of the list has been reached")
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   	    WScript.ConnectObject application, "off"
		WScript.Quit
	End If

Sub UpdateSBin
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "500N"'******Sets plant value
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "0005"'*****Sets Sloc Value
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").setFocus
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").caretPosition = 4

Session.findById("wnd[1]").sendVKey 0
On Error Resume next
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/txtMARD-LGPBE").text = ExcelSheet.Cells(Row,4).Value
If Err.Number<>0 Then
	Session.findById("wnd[2]").sendVKey 0
	Session.findById("wnd[1]").close
	ExcelSheet.Cells(Row,6).Value = "Material does not exist in Plant / Sloc combination"
	Row=Row+1
	Err.Clear
	Exit Sub
	End If
On Error Goto 0
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/txtMARD-LGPBE").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/txtMARD-LGPBE").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
Row=Row+1
End Sub
'******************************************
'*For questions and comments contact      *
'*Derek.m.Ashworth@alstom.com 860-285-9135*
'*MM02-plant data1.vbs ver .09a           *
'******************************************
