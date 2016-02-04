Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim messtxt,z,Row,iRow,strCurrentTab,file

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
Function sbarStatus()
	sbarStatus = Session.findbyid("wnd[0]/sbar").text
End Function

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
ExcelApp.Visible=True
Row=InputBox("Row to start at")
session.findById("wnd[0]").maximize 
session.StartTransaction("VD52")
Do
Call Main
Row=Row+1
Loop Until ExcelSheet.Cells(Row,1).Value=""

Call EndScript

Sub Main()
	session.findById("wnd[0]/usr/ctxtKUNNR").text = ExcelSheet.Cells(Row,1).Value
	Session.findById("wnd[0]/usr/ctxtVKORG").text = ExcelSheet.Cells(Row,2).Value
	Session.findById("wnd[0]/usr/ctxtVTWEG").text = ExcelSheet.Cells(Row,3).Value
	Session.findById("wnd[0]/usr/ctxtMATNR_R-LOW").text = ExcelSheet.Cells(Row,4).Value
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	Session.findbyid("wnd[0]/tbar[1]/btn[2]").press
	session.findById("wnd[0]/usr/txtMV10A-ANTLF").text = ExcelSheet.Cells(Row,6).Value
	session.findById("wnd[0]/tbar[0]/btn[11]").press
	ExcelSheet.Cells(Row,7).Value = sbarStatus()
End Sub

Sub EndScript()
ExcelWorkbook.Close=True
ExcelApp.Quit
MsgBox("Changes Completed")
WScript.Quit 
End Sub
