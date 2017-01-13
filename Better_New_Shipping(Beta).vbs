Const userName = "dma02"
Dim ExcelApp,ExcelWorkbook,ExcelSheet,file,RowA,i
Dim Row,SOrder,FrCom(9),FrPro(9),InvSum(9),TxtCo,x,ExcelSheet2
Dim FrBill,z,FrBill2,Wnd1TTL,MessText,MainTtl,h,p,j
Dim PMx3,MtrLn,PMxRow,stat1,stat2,strContinue,mtlNum,strTrackChk,zebra
Dim costSum(9),costFreight
Const costMarkup =1.25 



Function password()
	password=InputBox("SAP PE1 Password")
End Function

Function partDuplicate(zebra)
	If PMxRow=-1 Then
		PMxRow=0
	End If
	mtlNum=session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,"&zebra&"]").text
	If mtlNum ="SHIP-HANDLING" Then
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,"&zebra&"]").setfocus
		session.findById("wnd[0]").sendVKey 2
		session.findbyId("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09").select
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem "0009","Column1"
		Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "0009","Column1"
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem "0009","Column1"
		strTrackChk=session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text
			If strTrackChk=TxtCo Then
				partDuplicate=True
			Else 
				partDuplicate=False
			End If
		session.findbyid("wnd[0]/tbar[0]/btn[3]").press
	End If 
End Function

	' Open SAP
	Dim WshShell
	set WshShell = WScript.CreateObject("WScript.Shell")

	' Not yet completed
	If not(WshShell.AppActivate("SAP Logon")) then
		WshShell.Exec("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
		Wscript.Sleep 500
		
		Dim y : i = 0
		Do While not(WshShell.AppActivate("SAP Logon"))
			WScript.Sleep 250
			timeoutCheck y, 400, "SAP Logon Timeout"		' Loop a max of 10 seconds
		Loop
	End if
	
	' Run GUI Script
	Dim application, SapGuiAuto, connection,isNewConn
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
	If isNewConn then
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
	'session.LockSessionUI
'******************************************
'Option Explicit


file = ChooseFile(defaultLocalDir)
'MsgBox file

Function ChooseFile (ByVal initialDir)
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

    Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    Dim winVersion

    ' This collection should contain just the one item
    For Each objItem in colItems
        'Caption e.g. Microsoft Windows 7 Professional
        'Name e.g. Microsoft Windows 7 Professional |C:\windows|...
        'OSType e.g. 18 / OSArchitecture e.g 64-bit
        'Version e.g 6.1.7601 / BuildNumber e.g 7601
        winVersion = CInt(Left(objItem.version, 1))
    Next
    Set objWMIService = Nothing
    Set colItems = Nothing

    If (winVersion <= 5) Then
        ' Then we are running XP and can use the original mechanism
        Set cd = CreateObject("UserAccounts.CommonDialog")
        cd.InitialDir = initialDir
        cd.Filter = "VBScript Data Files |*.xls;*.xlsx;*.xlsm|All Files|*.*"
        ' filter index 4 would show all files by default
        ' filter index 1 would show zip files by default
        cd.FilterIndex = 1
        If cd.ShowOpen = True Then
            ChooseFile = cd.FileName
        Else
            ChooseFile = ""
        End If
        Set cd = Nothing    

    Else
        ' We are running Windows 7 or later
        Set shell = CreateObject( "WScript.Shell" )
        Set ex = shell.Exec( "mshta.exe ""about: <input type=file id=X><script>X.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(X.value);close();resizeTo(0,0);</script>""" )
        ChooseFile = Replace( ex.StdOut.ReadAll, vbCRLF, "" )

        Set ex = Nothing
        Set shell = Nothing
    End If
End Function

Function statPrePay()
	session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press 
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12").select 
		If session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR3").key = "ZPP" Then
			statPrePay="Yes"
		 Else
		 	statPrePay="No"
		End If
	session.findById("wnd[0]/tbar[0]/btn[3]").press
End Function

Function sbarStatus()
	sbarStatus = Session.findbyid("wnd[0]/sbar").text
End Function
'*********************************
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
Set ExcelSheet2 = ExcelWorkbook.Worksheets("Sheet2")
ExcelApp.Visible=True
Row=InputBox("Row to start at")

RowA=Row
Do
Call MainLoop
Call PMxOut
Call RsetVar
Row=Row+1
Loop Until ExcelSheet.Cells(Row,8).Value=""

Sub MainLoop
For i=0 To 9
	SOrder=ExcelSheet.Cells(Row,8).Value
	If SOrder=ExcelSheet.Cells(Row+1,8).Value Then
		FrCom(i)= ExcelSheet.Cells(Row,4).Value
		FrPro(i)=ExcelSheet.Cells(Row,5).Value
		InvSum(i)=ExcelSheet.Cells(Row,10).Value
		'costSum(i)=ExcelSheet.Cells(Row,10).Value
		ExcelSheet.Cells(Row,15).Value="Added"
	 Else
	 	FrCom(i)= ExcelSheet.Cells(Row,4).Value
		FrPro(i)=ExcelSheet.Cells(Row,5).Value
		InvSum(i)=ExcelSheet.Cells(Row,10).Value
		'costSum(i)=ExcelSheet.Cells(Row,10).Value
		'Call PMxOut
		Exit For
	End If
	Row=Row+1
Next
End Sub

Sub PMxOut
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nva02"'<-----Change sales order
	session.findById("wnd[0]").sendVKey 0
	TxtCo=""
		
	session.findById("wnd[0]").maximize
	session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text =SOrder
	session.findById("wnd[0]").sendVKey 0

			
	On Error Resume Next
	'session.findById("wnd[1]").sendVKey 0
	Wnd1TTL=session.findbyid("wnd[1]").text
	MessText=session.findbyid("wnd[1]/usr/txtMESSTXT2").text
	If Left(Wnd1TTL,4)="Help" Then
		ExcelSheet.Cells(Row,14).Value = "Order Closed"
		session.findbyid("wnd[1]/tbar[0]/btn[5]").press
		Wnd1TTL="none"
		Exit Sub
 	ElseIf Left(MessText,4)="Over" Then
 		ExcelSheet.Cells(Row,16).Value = "Warning - Not Processed"
		session.findbyid("wnd[1]/tbar[0]/btn[5]").press
		MessText="none"
		Exit Sub
	ElseIf Left(sbarStatus,5)="No au" Then
		ExcelSheet.Cells(Row,14).Value = sbarStatus
		Exit Sub
	End If
	session.findById("wnd[1]").sendVKey 0
	MainTtl=session.findbyid("wnd[0]").text
	MainTtl=Left(MainTtl,21)
	MainTtl=Right(MainTtl,2)
	If MainTtl="BP" Then
		ExcelSheet.Cells(Row,16).Value = "Warning - BP Order freight not added"
		MainTtl="none"
		session.findbyid("wnd[0]/tbar[0]/btn[3]").press
		Exit Sub
	End If

On Error Goto 0
	If statPrePay="Yes" Then
	 	ExcelSheet.Cells(Row,14).Value = "Order is Prepay Non-billable"
	 	Exit Sub
	End If
	
	For x= 0 To 9
		If FrCom(x)="" Then 
		Exit For
		End If
	  TxtCo = TxtCo&"***"&FrCom(x)&"-"&FrPro(x)
	  'TxtCo=TxtCo&"-"
	  'TxtCo=TxtCo&FrPro(x)
	  ExcelSheet2.Cells(RowA,1).Value=TxtCo
	Next
	Call FindRow2
		If strContinue="No" Then
			session.findbyid("wnd[0]/tbar[0]/btn[3]").press
			ExcelSheet.cells(Row,14)="Already invoiced"
			Exit Sub
		End If
			
	'session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POAN").press
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,"&PMxRow&"]").text = "ship-handling"
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,"&PMxRow&"]").text = "1"
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,"&PMxRow&"]").setFocus
	'session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,"&PMxRow&"]").caretPosition = 9
	session.findById("wnd[0]").sendVKey 0
		On Error Resume Next
	stat2=session.findbyId("wnd[1]").text
	On Error Goto 0
	If Right(stat2,11) ="Information" Then
		ExcelSheet.cells(Row,16)=session.findbyid("wnd[1]/usr/txtMESSTXT1").text
		session.findbyid("wnd[1]/tbar[0]/btn[0]").press
		stat2="nope"
	End if
	If Left(sbarStatus,8)="No goods" Then
		session.findById("wnd[0]").sendVKey 0
	End if
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,"&PMxRow&"]").setFocus
	session.findById("wnd[0]").sendVKey 2
	stat1=session.findById("wnd[0]/sbar").Text
	
	If Left(stat1,8)="No goods" then
		session.findById("wnd[0]").sendVKey 0
	End If
	session.findbyId("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09").select
'	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-ARKTX[5,"&PMxRow&"]").text = TxtCo
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem "0009","Column1"
	Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "0009","Column1"
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem "0009","Column1"
	'session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem "0009", "Column1"
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = TxtCO
		
	FrBill=0
	For z=0 To 9
		If InvSum(z)="" Then 
			Exit For
		End If
	 FrBill= FrBill+InvSum(z)
	 'costFreight=costFreight+costSum(z)
	 Next
	 If FrBill<30 Then
	 	FrBill2=30
	  Else
	    FrBill2=FormatNumber(FrBill*costMarkup,2)
	 End If
	 session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05").select
	 session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KOAN").press
	 session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,1]").text = "pr00"
	 session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,1]").text = FrBill2
	 session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,1]").setFocus
	 session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,1]").caretPosition = 16
	 session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KOAN").press
	 session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,1]").text = "ycmc"
	 session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,1]").text = FrBill
	 session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,1]").setFocus
	 session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,1]").caretPosition = 16
	 session.findById("wnd[0]").sendVKey 0
	 session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KONY").press
	 session.findById("wnd[1]/usr/lbl[1,5]").setFocus
	 session.findById("wnd[1]/usr/lbl[1,5]").caretPosition = 8
	 session.findById("wnd[1]/tbar[0]/btn[0]").press
	 WScript.Sleep(2000)
	 session.findById("wnd[0]/tbar[0]/btn[11]").press
	 On Error Resume next
	 session.findById("wnd[1]").close
	 session.findById("wnd[2]/usr/btnBUTTON_2").press
	 On Error Goto 0
	 ExcelSheet.Cells(Row,14).Value = session.findById("wnd[0]/sbar").Text
		 ExcelSheet2.Cells(RowA,2).Value=FrBill
		 RowA=RowA+1
  End Sub
 
Sub RsetVar
 For p=0 To 9
	FrCom(p)=""
	FrPro(p)=""
	InvSum(p)=""
	strContinue=""
 Next
End Sub

Sub FindRow2
PMx3=True
h=1
i=0
j=session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").visiblerowcount
	Do Until PMx3=False
		MtrLn=session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,"&i&"]").text
		If i=(j-1) Then
			session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").verticalScrollbar.position = j*h
			h=h+1
			i=-1
		End If
		PMxRow=i
		If MtrLn="" Then
			If PMxRow=-1 Then
			PMxRow=0
			End If
			
		PMx3=False
		End If
	
	If partDuplicate(PMxRow) = True Then
		strContinue="No"
		Exit Do
	End If
		i=i+1
	Loop
End Sub
 
Sub timeoutCheck(loopVar, maxVal, title)
	loopVar = loopVar + 1		'Timeout var increase'
	if loopVar > maxVal then
		if isNull(title) then
			title = "Loop Timeout"
		End if
		OkCancelMsg "A loop has timed out. Press OK to continue or Cancel to exit", title
		loopVar = 0
	End if
End Sub	  
	  