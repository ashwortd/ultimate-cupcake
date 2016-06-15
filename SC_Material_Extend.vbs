'File: SC_Material_Extend.vbs
'Author: Derek Ashworth
'Edit Date: 05/17/2016
Option Explicit

'File Definitions
Const fileName="SC_Material_Data.xlsm"
Const fileDirectory="\\winfile02\data\CustSvc\Parts\Pmx Scripting\Script Data\"
Const showWindow = True
Dim excelFileLocation,excelApp,excelWorkbook,excelWorksheet
Dim intRow,userName,password,window

'Functions

Function strMsgBox(window)
	strMsgBox=session.findById("wnd["&window&"]/usr/txtMESSTXT1").text
End Function

Function sbarStatus()
	sbarStatus = Session.findbyid("wnd[0]/sbar").text
End Function

Function currentTab()
	currentTab=Session.activewindow.guifocus.ID
	currentTab= Left(currentTab,50)
	currentTab= Right(currentTab,8)
End Function
'File locations

intRow=InputBox("What is the starting row to extend?")
Set excelApp=CreateObject("Excel.Application")
excelFileLocation=fileDirectory&fileName
Set excelWorkbook=excelApp.workbooks.open(excelFileLocation)
excelApp.visible=True
OpenSAP()

sub openSAP
	' Open SAP
	Dim WshShell
	set WshShell = WScript.CreateObject("WScript.Shell")

	' Not yet completed
	If not(WshShell.AppActivate("SAP Logon")) then
		WshShell.Exec("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
		Wscript.Sleep 500
		
		Dim i : i = 0
		Do While not(WshShell.AppActivate("SAP Logon"))
			WScript.Sleep 250
			timeoutCheck i, 400, "SAP Logon Timeout"		' Loop a max of 10 seconds
		Loop
	End if
	
	' Run GUI Script
	Dim application, SapGuiAuto, connection, session, isNewConn
	If Not IsObject(application) Then
	   Set SapGuiAuto  = GetObject("SAPGUI")
	   Set application = SapGuiAuto.GetScriptingEngine
	End If
	If Not IsObject(connection) Then
		If application.Children.Count > 0 then				' If it has connections
			Set connection = application.Children(0)
			isNewConn = false
			If not connection.description = "1.1 PMx Production (PE1)" then
				Set connection = application.OpenConnection("1.1 PMx Production (PE1)", true)
				isNewConn = true
			End if
		Else
			Set connection = application.OpenConnection("1.1 PMx Production (PE1)", true)
			isNewConn = true
		End if
	End If
	If Not IsObject(session) Then
	   Set session = connection.Children(0)
	End If
	If IsObject(WScript) Then
	   WScript.ConnectObject session,     "on"
	   WScript.ConnectObject application, "on"
	End If
	session.findById("wnd[0]").maximize
	
	' Login
	If isNewConn Then
		userName=InputBox("SAP PE1 Username:")
		password=InputBox("SAP PE1 Password:")
		session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = userName
		session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = password
		session.findById("wnd[0]").sendVKey 0

		' If any messages come up clear them
		Dim messageCount, logonOption
		messageCount = 0
		Do while session.Children.Count > 1
			if messageCount > 5 then
				MsgBox "Error, too many message boxes detected"
				Wscript.quit
				exit do
			else
				Set logonOption = session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1", false)
				' Check for message to bump off another person logged on
				if TypeName(logonOption) <> "Nothing" then
					logonOption.select
				End if
				session.findById("wnd[1]/tbar[0]/btn[0]").press
			End if
			messageCount = messageCount + 1
		Loop
		
		
	Else
		Dim sessionCount
		sessionCount = connection.Children.Count
		
		session.CreateSession
		do while connection.Children.Count <= sessionCount
			WScript.Sleep 250
		loop
		Set session = connection.Children(connection.Children.Count - 1)
	End If
End Sub

Sub extendMaterial
	session.StartTransaction("MM01")
	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = excelWorksheet.Cells(Row,1).Value
	session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = "A"
	session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = "ZENG"
	session.findById("wnd[0]").sendVKey 0
		If sbarStatus ="Material type Project Materials copied from master record" Then
			Session.findById("wnd[0]").sendVKey 0
		End If
		If sbarStatus ="Material type Standard Components copied from master record" Then
			Session.findById("wnd[0]").sendVKey 0
		End if
	session.findById("wnd[1]/tbar[0]/btn[20]").press
End Sub
