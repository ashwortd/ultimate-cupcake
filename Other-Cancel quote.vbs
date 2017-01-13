Dim SapGuiApp,Connection,Session,FileObject,Counter
Dim ExcelApp,ExcelWorkbook,ExcelSheet,currentTab
Dim messtxt,z,Row,WshShell,File,strStatus,strWndName
file = ChooseFile(defaultLocalDir)
MsgBox file

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
'*********************
Function setFlow
	Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZCPNUM").text = "CPTRX-Flow"
	Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_CNTRCTCODE").text = "TRX-FLOW"
End Function
'****************************************
'Check for Logon status and connect to GUI
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
'******************************************
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (File)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet2")
ExcelApp.Visible = True
Row=InputBox("Starting row?")

session.findById("wnd[0]").maximize 
session.findById("wnd[0]/tbar[0]/okcd").text = "/nva22"
session.findById("wnd[0]").sendVKey 0

Do While ExcelSheet.Cells(Row,2).Value <> ""
 session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = ExcelSheet.Cells(Row,2).Value
 session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 8
 session.findById("wnd[0]").sendVKey 0
 strWndName = Session.findbyid("wnd[1]").text
If Left(strWndName,4)="Info" Then
	Call subWon

ElseIf Left(strWndName,4)="Canc" Then
	Call subCanc
Else
Call subClose
End if
	
strWndName ="none"
strStatus = session.findById("wnd[0]/sbar").Text
ExcelSheet.Cells(Row,3).Value = strStatus
If Left(strStatus,4) ="Main" Then
  Session.findbyid("wnd[0]/tbar[0]/btn[12]").press
  Session.findbyid("wnd[1]/usr/btnSPOP-OPTION1").press
End If
strStatus="none"
Row=Row+1
Loop
WScript.Quit
Sub subWon
Session.findbyid("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press 
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11").select 
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR5").key = "WON"
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR5").setFocus
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12").select
	Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/chkVBAK-ZZWFLW_IND").selected = False
	setFlow
	Session.findById("wnd[0]/tbar[0]/btn[11]").press
strWndName = Session.findbyid("wnd[1]").text
 If Left(strWndName,4)="Work" Then
	Session.findbyid("wnd[1]").close
	Session.findbyid("wnd[2]/usr/btnBUTTON_2").press
	
  ElseIf Left(strWndName,4)="Save" Then
    Session.findbyid("wnd[1]/usr/btnSPOP-VAROPTION1").press
    strWndName = Session.findbyid("wnd[1]").text
      If Left(strWndName,4)="Work" Then
	   Session.findbyid("wnd[1]").close
	   Session.findbyid("wnd[2]/usr/btnBUTTON_2").press
	   End If
	ElseIf Left(strWndName,4)="Save" Then
	  Session.findbyid("wnd[1]/tbar[0]/btn[0]").press
	
 End If
 If Right(strWndName,8)="Document" Then
	   Session.findbyid("wnd[1]/usr/btnSPOP-VAROPTION1").press
	  End If
	   
End Sub
Sub subCanc
  Session.findbyid("wnd[1]/tbar[0]/btn[0]").press
  ExcelSheet.Cells(Row,4).Value="No status object"
  session.findById("wnd[0]/tbar[0]/okcd").text = "/nva22"
  session.findById("wnd[0]").sendVKey 0
 End Sub
 Sub subClose
 Session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press 
	Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11").select 
	Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR5").key = "OCA"
	Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\11/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR5").setFocus 
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12").select
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/chkVBAK-ZZWFLW_IND").selected = False
	setFlow
	Session.findById("wnd[0]/tbar[0]/btn[11]").press
	strWndName = Session.findbyid("wnd[1]").text
 	If Left(strWndName,4)="Info" Then
	Session.findbyid("wnd[1]").close
 	ElseIf Left(strWndName,4)="Work" Then
	  Session.findbyid("wnd[1]").close
	  Session.findbyid("wnd[2]/usr/btnBUTTON_2").press
    ElseIf Left(strWndName,4)="Save" Then
      Session.findbyid("wnd[1]/usr/btnSPOP-VAROPTION1").press
      strWndName = Session.findbyid("wnd[1]").text
        If Left(strWndName,4)="Work" Then
	      Session.findbyid("wnd[1]").close
	      Session.findbyid("wnd[2]/usr/btnBUTTON_2").press
	    End If
	  End If
	  If Right(strWndName,8)="Document" Then
	   Session.findbyid("wnd[1]/usr/btnSPOP-VAROPTION1").press
	  End If
	   
    End Sub
    
  