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
Dim sapstatus,strChkBox,strWind1,skip

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
If file="" Then 
	WScript.Quit
End If
'****************
Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible= True
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
Set ExcelSheet = ExcelWorkbook.Worksheets(1)
Row=InputBox("Row to start at")
Session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nva22"
session.findById("wnd[0]").sendVKey 0
QtOd=InputBox("Update Contract DB and GE Contract code for Quotes or Orders? (q or o)")
	If QtOd="q" Then
		tcode="/nva22"
	ElseIf QtOd="o" Then
	   	tcode="/nva02"
	Else 
	    MsgBox ("Invalid Sales document type")
	   	ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"		
		WScript.Quit
	End If
Do
session.findById("wnd[0]/tbar[0]/okcd").text = tcode
session.findById("wnd[0]").sendVKey 0
On Error Resume next 
strWind1= Session.findbyid("wnd1").text
On Error Goto 0
' If Not strWind1 Is Nothing then
' End if
skip="No"
Call checkSO
Call enterdata
row=row+1
Loop Until ExcelSheet.Cells(Row,1).Value=""

	MsgBox("Thanks for shopping at Scripts 'R' Us")
		ExcelApp.close(True)
		Set ExcelApp=Nothing
		Set ExcelWorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"		
		WScript.Quit
		
Sub checkSO()
	If Left(strWind1,4)="Help" Then
		ExcelSheet.Cells(Row,4).Value="Order Closed"	
		Session.findbyid("wnd[1]/tbar[0]/btn[5]").press
		skip="Yes"
	End If
End Sub

Sub enterdata()
	If skip="yes" Then
		Exit Sub
	End if
	sapstatus=0
	Session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = ExcelSheet.Cells(Row,1).Value
	session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 5
	Session.findById("wnd[0]").sendVKey 0
	ExcelSheet.Cells(Row,4).Value = session.findById("wnd[0]/sbar").Text
	sapstatus=session.findById("wnd[0]/sbar").Text
	If sapstatus="Lock table overflow" Then
		WScript.Sleep(5000)
		Session.findById("wnd[0]").sendVKey 0
	End If
		
	On Error Resume Next
	If Not session.findById("wnd[1]/usr/txtMESSTXT1",false)Is Nothing Then
		Session.findById("wnd[1]").sendVKey 0
	End If
	session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[14]/menu[1]").select
	'session.findById("wnd[0]/mbar/menu[2]/menu[1]/menu[14]/menu[0]").select
	strChkBox=ExcelSheet.Cells(Row,5).Value
	Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_CCTYP").text="SAB 104"
	If strChkBox <>"" then
		Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZCPNUM").text = "CPPSP-PROD"
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_CNTRCTCODE").text = "PSP-PROD"
		session.findById("wnd[0]").sendVKey 0
			
	Else
		Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZCPNUM").text = "CPTRX-FLOW"
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_CNTRCTCODE").text = "TRX-FLOW"
		session.findById("wnd[0]").sendVKey 0
	End If
	
	If session.findById("wnd[0]/sbar").Text ="Please check the  Order Intake Date" Then
		Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_ORDINTDAT").text = date
	End If	
	
	Session.findById("wnd[0]/tbar[0]/btn[11]").press
		If Not session.findById("wnd[1]/usr/txtSPOP-TEXTLINE1",false)Is Nothing Then '.text="Document Incomplete" Then
			Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
			ExcelSheet.Cells(Row,3).Value = session.findById("wnd[0]/sbar").Text
			
		End If
		If Not Session.findbyid("wnd[1]",False) Is Nothing Then
			session.findById("wnd[1]").close
			Session.findById("wnd[2]/usr/btnBUTTON_2").press
			session.findById("wnd[1]").close
			session.findById("wnd[2]/usr/btnBUTTON_2").press

		End If 
	On Error Goto 0
	ExcelSheet.Cells(Row,3).Value = session.findById("wnd[0]/sbar").Text
	
End Sub

