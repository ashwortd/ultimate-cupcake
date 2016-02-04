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

Dim file
Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row,strBob


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
Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible=True
Set ExcelWorkbook = ExcelApp.Workbooks.Open(file)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet2")
Row=InputBox("Row to start at")
 If Row=""Then 
 	ExcelApp.Quit
	Set ExcelApp=Nothing
	Set ExcelWoorkbook=Nothing
	Set ExcelSheet=Nothing
	WScript.ConnectObject session,     "off"
	WScript.ConnectObject application, "off"
	WScript.Quit
 End If
'*********************

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/ntrip"
session.findById("wnd[0]").sendVKey 0

Do
Call Main
Row=Row+1
Loop Until ExcelSheet.Cells(Row,1).Value = ""

ExcelApp.Quit
	Set ExcelApp=Nothing
	Set ExcelWoorkbook=Nothing
	Set ExcelSheet=Nothing
	WScript.ConnectObject session,     "off"
	WScript.ConnectObject application, "off"
	WScript.Quit
	
Sub Main
session.findById("wnd[0]/tbar[1]/btn[24]").press
session.findById("wnd[1]/usr/ctxtPTP40-PERNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[1]/usr/ctxtPTP40-PERNR").caretPosition = 4
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[33]").press
On Error Resume Next
 strBob=Session.findbyid("wnd[2]").text
On Error Goto 0
	If strBob="Error" Then
		Session.findbyid("wnd[2]").close
		ExcelSheet.Cells(Row,2).Value="Not available at this time"
		strBob="nada"
		Exit Sub
	End If

ExcelSheet.Cells(Row,2).Value=Session.findById("wnd[1]/usr/tabsTABSTRIP1/tabpTRAVELER/ssubSUB1:SAPLHRTRV_UTIL:0300/txtADDR3_VAL-NAME_TEXT").text
session.findById("wnd[1]/tbar[0]/btn[0]").press
End Sub
