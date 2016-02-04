Dim ExcelApp,ExcelWorkbook,ExcelSheet,file
Dim Row


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
If file="" Then 
	WScript.Quit
End If
'****************
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
Set ExcelSheet = ExcelWorkbook.Worksheets("500B-ABC")
Row=InputBox("Row to start at")
ExcelApp.Visible=True
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0

Do While ExcelSheet.Cells(Row,1).Value <>""
Call Main
Row=Row+1
Loop

Sub Main()
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[19]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(9).selected = true
'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,9]").setFocus
'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,9]").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,2).Value
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = ExcelSheet.Cells(Row,3).Value
'session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").setFocus
'session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2481/ctxtMARC-MAABC").text = ExcelSheet.Cells(Row,5).Value
'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2481/ctxtMARC-MAABC").setFocus
'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2481/ctxtMARC-MAABC").caretPosition = 1
session.findbyid("wnd[0]/tbar[0]/btn[11]").press
'session.findById("wnd[0]/tbar[0]/btn[0]").press
'session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
ExcelSheet.Cells(Row,11).Value = session.findById("wnd[0]/sbar").Text
End Sub