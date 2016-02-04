Dim file,ExcelApp,ExcelWorkbook,ExcelSheet,Row
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

If Not IsObject(application) Then
   On Error Resume next
   Set SapGuiAuto  = GetObject("SAPGUI")
   	If Err.Number<>0 Then
   		MsgBox("You are not connected to PMx, please connect and try again")
   		On Error Goto 0
   		WScript.Quit
    End If
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

Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible=true
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
Set ExcelSheet = ExcelWorkbook.Worksheets(1)
Row=InputBox("Starting Row?")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0

Do
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
'session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 15
session.findById("wnd[0]").sendVKey 0
session.findbyID("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "500B"
'session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").caretPosition = 4
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text="G001"
session.findById("wnd[1]/usr/ctxtRMMG1-BWTAR").text=""
'session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text=""
'session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text=""
session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text=""
session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text=""
session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = "5013"
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "01"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findbyid("wnd[0]/tbar[1]/btn[8]").press
session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP05").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2156/ctxtMVKE-MVGR1").text = ""
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2156/ctxtMVKE-MVGR1").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2156/ctxtMVKE-MVGR1").caretPosition = 0
session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP12").select
session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2481/ctxtMARC-EKGRP").text="elv"
'session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
ExcelSheet.Cells(Row,8).Value = session.findById("wnd[0]/sbar").Text
Row=Row+1
Stat1=ExcelSheet.Cells(Row,1).Value
Loop While Stat1 <>""

	MsgBox("The end has come")
	ExcelWorkbook.Save
	Set ExcelApp=Nothing
	Set ExcelWorkbook=Nothing
	Set ExcelSheet=Nothing
	WScript.ConnectObject session,     "off"
	WScript.ConnectObject application, "off"
	WScript.Quit
	