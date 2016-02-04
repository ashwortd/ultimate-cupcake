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
Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row,QtOd,tcode,wndttl,status
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

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open(file)
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
		Set ExcelWorkbook=Nothing
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
		Set ExcelWorkbook=Nothing
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
Session.findById("wnd[0]/tbar[0]/btn[11]").press
On Error Resume next
wndttl=Session.findbyid("wnd[1]").text
On Error Goto 0
If wndttl="Information" Then
	Session.findbyid("wnd[1]").close
	Session.findbyid("wnd[1]/usr/btnSPOP-VAROPTION1").press
	wndttl="none"
End If
If wndttl="Workflow Selection" Then
	Session.findbyid("wnd[1]").close
	Session.findbyid("wnd[2]/usr/btnBUTTON_2").press
	wndttl="none"
End If

status=Session.findbyid("wnd[0]/sbar").Text
If Left(status,9)="The order" Then
	ExcelSheet.Cells(Row,6).Value = status
	Session.findbyid("wnd[0]/tbar[0]/btn[12]").press
	Session.findbyid("wnd[1]/usr/btnSPOP-OPTION1").press
End If
ExcelSheet.Cells(Row,5).Value = session.findById("wnd[0]/sbar").Text
Row=Row+1
End Sub
