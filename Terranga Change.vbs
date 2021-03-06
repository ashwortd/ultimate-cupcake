If Not IsObject(application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject application, "on"
End If

Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet,QtOd,tcode
Dim sapstatus

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
Row=InputBox("Row to start at")-1
Session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nva22"
session.findById("wnd[0]").sendVKey 0
QtOd=InputBox("Change terranga for Quotes or Orders? (q or o)")
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
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"		
		WScript.Quit
	End If
	
Call orderslct

Sub orderslct()
session.findById("wnd[0]/tbar[0]/okcd").text = tcode
session.findById("wnd[0]").sendVKey 0
	row=row+1
	If ExcelSheet.Cells(Row,1).Value=("") Then
		MsgBox("The end has come")
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"		
		WScript.Quit
	End If
	Call enterdata
End Sub

Sub enterdata()
	sapstatus=0
	Session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = ExcelSheet.Cells(Row,1).Value
	session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 5
	Session.findById("wnd[0]").sendVKey 0
	ExcelSheet.Cells(Row,4).Value = session.findById("wnd[0]/sbar").Text
	sapstatus=session.findById("wnd[0]/sbar").Text
	If sapstatus="Lock table overflow" Then
		WScript.Sleep(5000)
		Call enterdata
	End If
		
	On Error Resume Next
	If Not session.findById("wnd[1]/usr/txtMESSTXT1",false)Is Nothing Then
		Session.findById("wnd[1]").sendVKey 0
	End If
	session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[14]/menu[1]").select
	'session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[14]/menu[0]").select
	Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/txtVBAK-ZZCPNUM").text = "1"
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_TERANGA").text = ExcelSheet.Cells(Row,2).Value
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_TERANGA").setFocus
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_TERANGA").caretPosition = 4
	session.findById("wnd[0]").sendVKey 0
	If session.findById("wnd[0]/sbar").Text ="Please check the  Order Intake Date" Then
	Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_ORDINTDAT").text = date
	End If
	Session.findById("wnd[0]/tbar[0]/btn[11]").press
		If Not session.findById("wnd[1]/usr/txtSPOP-TEXTLINE1",false)Is Nothing Then '.text="Document Incomplete" Then
			Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
			ExcelSheet.Cells(Row,3).Value = session.findById("wnd[0]/sbar").Text
			'Call orderslct
		End If
		If Not Session.findbyid("wnd[1]",False) Is Nothing Then
			session.findById("wnd[1]").close
			Session.findById("wnd[2]/usr/btnBUTTON_2").press
			session.findById("wnd[1]").close
			session.findById("wnd[2]/usr/btnBUTTON_2").press

		End If 
	On Error Goto 0
	ExcelSheet.Cells(Row,3).Value = session.findById("wnd[0]/sbar").Text
	Call orderslct
End Sub

