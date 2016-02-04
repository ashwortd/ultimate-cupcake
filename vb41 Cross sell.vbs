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
Dim messtxt,z,Row,intSlaveRow,strLeadPart,intFastEntryRow
Dim strStatus1,strStatus2,strAlreadyEntered1,strAlreadyEntered2

'************Ask for data file
Set objDialog = CreateObject("UserAccounts.CommonDialog")

objDialog.Filter = "VBScript Data Files |*.xls;*.xlsx;*.xlsm|All Files|*.*"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\Scripts"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
 		WScript.ConnectObject session,     "off"
   	    WScript.ConnectObject application, "off"
	 	WScript.Quit
  
End If
'****************
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)'<----- open excel file selected
Set ExcelSheet = ExcelWorkbook.Worksheets("Cross_Sell")'<----looks for tab named cross_sell in workbook

Row=InputBox("Row to start at")'<------dialog asking for starting row number
z=ExcelSheet.Cells(Row,1).Value'<-----Makes sure that the starting row, column A is not blank
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nvb41"'<-----cross sell maintenance tcode
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtD000-KSCHL").text = "ZCS1"'<-------cross sell determination type
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[1,0]").select'<----this selects setting cross sell by sales org/ distr channel
session.findById("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtKOMGD-VKORG").text = "5013"
session.findById("wnd[0]/usr/ctxtKOMGD-VTWEG").text = "01"
Do While z <> "" '<-----Start loop

session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKOMGD-MATNR[0,0]").text = ExcelSheet.Cells(Row,1).Value'<----column A value in excel sheet
Session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKONDD-SMATN[2,0]").text = ExcelSheet.Cells(Row,2).Value'<----column B value in excel sheet

session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKONDD-SMATN[2,0]").setFocus
'Session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKONDD-SMATN[2,0]").caretPosition = 9
session.findById("wnd[0]").sendVKey 0
strStatus1=session.findById("wnd[0]/sbar").Text
strStatus2=Left(strstatus1,8)
'MsgBox(strStatus1)
'MsgBox(strStatus2)
If strStatus2="Material" Then
	While strStatus2="Material"
	Row=Row+1
	session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKOMGD-MATNR[0,0]").text = ExcelSheet.Cells(Row,1).Value'<----column A value in excel sheet
	Session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKONDD-SMATN[2,0]").text = ExcelSheet.Cells(Row,2).Value'<----column B value in excel sheet
	session.findById("wnd[0]").sendVKey 0
	strStatus1=session.findById("wnd[0]/sbar").Text
	strStatus2=Left(strstatus1,8)
	Wend
End if
intSlaveRow = Row+1
strLeadPart = ExcelSheet.Cells(Row,1).Value
If strLeadPart = ExcelSheet.Cells(intSlaveRow,1).Value Then
	Session.findbyid("wnd[0]/tbar[1]/btn[2]").press
	intFastEntryRow=1
	
	Do While strLeadPart = ExcelSheet.cells(intSlaveRow,1)
	Session.findbyid("wnd[0]/usr/tblSAPMV13DTCTRL_D0200/ctxtKONDDP-SMATN[0,"&intFastEntryRow&"]").text = ExcelSheet.Cells(intSlaveRow,2).Value
	intFastEntryRow=intFastEntryRow+1
	intSlaveRow=intSlaveRow+1
	Row=Row+1
	Loop
End If
session.findById("wnd[0]/tbar[0]/btn[11]").press
strAlreadyEntered1 = session.findById("wnd[0]/sbar").Text
strAlreadyEntered2 = Right(strAlreadyEntered1,7)
If strAlreadyEntered2="entered" Then
	While strAlreadyEntered2="entered"
	Session.findbyid("wnd[0]/usr/btnFCODE_ENF1").press
	session.findById("wnd[0]/tbar[0]/btn[11]").press
	strAlreadyEntered1 = session.findById("wnd[0]/sbar").Text
	strAlreadyEntered2 = Right(strAlreadyEntered1,7)
	Wend
End if
ExcelSheet.Cells(Row,3).Value = session.findById("wnd[0]/sbar").Text'<-----Add status to C column in excel sheet
Row=Row+1
z=ExcelSheet.Cells(Row,1).Value'<-----Check next row to see if blank
Loop '<-----If next row blank, it ends loop snd stops the script.
	MsgBox("The end has come")
	session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
	session.findById("wnd[0]").sendVKey 0	
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   	    WScript.ConnectObject application, "off"
	 	WScript.Quit