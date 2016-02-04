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
If file ="" Then 
WScript.Quit
End If
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
'*Open Excel data file and set worksheet    
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
ExcelApp.Visible=True
Set ExcelSheet = ExcelWorkbook.Worksheets("NCR_ENTRY")
Dim Row,StatusBar,SLBlank,i
Row=InputBox("Which row would you like to start with?")
	If Row = False Then
		Call Endscript
	End If
Do	
Call Main
Row=Row+1
Loop While ExcelSheet.Cells(Row,1).Value <>""

Call Endscript

Sub Endscript
MsgBox("Script completed")
		ExcelWorkbook.Close(True)
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"
		WScript.Quit
End Sub

Sub Main
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nqm01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0101/txtGW_VIQMEL-REFNUM").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0101/ctxtGW_VIQMEL-QMGRP").text = "Z01"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0101/txtGW_VIQMEL-QMCOD").text = ExcelSheet.Cells(Row,3).Value
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0101/ctxtGW_VIQMEL-QMNAM").text = ExcelSheet.Cells(Row,5).Value
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0101/ctxtG_REPENT_VAL").text = "RU-4072"
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0101/ctxtGW_VIQMEL-ZZLOCAL_PROJ").text = ExcelSheet.Cells(Row,4).Value
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0101/ctxtG_VERIFIER").text = ExcelSheet.Cells(Row,7).Value
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0101/ctxtG_VERIFIER").setFocus
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0101/ctxtG_VERIFIER").caretPosition = 6
session.findById("wnd[0]/tbar[0]/btn[11]").press
StatusBar=session.findById("wnd[0]/sbar").Text
	If StatusBar <>"" Then
		Call ErrorSub
		StatusBar=""
		Exit Sub
	End If
session.findById("wnd[0]/tbar[0]/btn[3]").press
End Sub

Sub ErrorSub
ExcelSheet.Cells(Row,8).Value= StatusBar
End Sub




