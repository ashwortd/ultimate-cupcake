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
Dim row
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible=True
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")'edit the sheet name for different list.
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nzmk05"
session.findById("wnd[0]").sendVKey 0
row=InputBox("Which row would you like to start on?")
Do
Call Main
row=row+1
Loop Until ExcelSheet.Cells(Row,1).Value=""
Call Cleanup

Sub Main
session.findById("wnd[0]/usr/ctxtP_LIFNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/ctxtP_EKORG").text = ExcelSheet.Cells(Row,2).Value
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/chkGW_COCKPIT-BLOCKPMX").selected = true
session.findById("wnd[0]/usr/chkGW_COCKPIT-BLOCKPMX").setFocus
session.findById("wnd[0]/tbar[0]/btn[11]").press
ExcelSheet.Cells(Row,3).Value = session.findById("wnd[0]/sbar").Text
End Sub

Sub Cleanup
	MsgBox("Script Complete")
	ExcelWorkbook.Close(True)
	ExcelApp.Quit
	Set ExcelApp=Nothing
	Set ExcelWorkbook=Nothing
	Set ExcelSheet=Nothing
    WScript.ConnectObject session,     "off"
    WScript.ConnectObject application, "off"	
	WScript.Quit
End sub