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
'************Ask for data file
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
'****************
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim Row,elementID,elementLeft,elementFinal,intComplete
Dim strWndTtl,strWndStat2,strGoNogo
Row=InputBox("Starting Row?","Starting Point")

Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible=True
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
Set ExcelSheet = ExcelWorkbook.Worksheets(1)

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
session.findById("wnd[0]").sendVKey 0
Do
strWndTtl="None"
intComplete=2
strGoNogo="go"
Call main
Call TabFind
Row=Row+1
Loop While ExcelSheet.cells(Row,1)<>""

ExcelWorkbook.Close(True)
Set ExcelApp=Nothing
Set ExcelWorkbook=Nothing
Set ExcelSheet=Nothing
ExcelApp.Quit
MsgBox("Process Complete")
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit

Sub main
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.cells(Row,1).value
session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").text = ExcelSheet.cells(Row,1).value
session.findById("wnd[0]").sendVKey 0
If session.findbyid("wnd[0]/sbar").text="Material type ALSTOM Materials and industry Plant engin./construction copied from master record" then
	session.findById("wnd[0]").sendVKey 0
End If
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findbyid("wnd[1]/usr/ctxtRMMG1-BWTAR").text=""
session.findbyid("wnd[1]/usr/ctxtRMMG1-VTWEG").text=""
session.findbyid("wnd[1]/usr/ctxtRMMG1-VKORG").text=""
session.findbyid("wnd[1]/usr/ctxtRMMG1_REF-BWTAR").text=""
session.findbyid("wnd[1]/usr/ctxtRMMG1_REF-VTWEG").text=""
session.findbyid("wnd[1]/usr/ctxtRMMG1_REF-VKORG").text=""
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.cells(Row,6).value
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = ExcelSheet.cells(Row,7).value
session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = ExcelSheet.cells(Row,8).value
session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = ExcelSheet.cells(Row,9).value
session.findById("wnd[1]/usr/ctxtRMMG1_REF-WERKS").text = ExcelSheet.cells(Row,2).value
session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGORT").text = ExcelSheet.cells(Row,3).value
session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGNUM").text = ExcelSheet.cells(Row,4).value
session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGTYP").text = ExcelSheet.cells(Row,5).value
session.findById("wnd[1]/tbar[0]/btn[0]").press
'session.findById("wnd[0]").sendVKey 0
On Error Resume Next
strWndStat2=session.findbyid("wnd[2]").text
On Error Goto 0
If strWndStat2="Error" Then
	session.findById("wnd[0]").sendVKey 0
	ExcelSheet.Cells(Row,20).Value ="Material has already been extended"
	strGoNogo ="nogo"
	session.findById("wnd[1]").close
	Exit Sub
End If
	
End Sub

Sub TabFind
If strGoNogo="nogo" Then
	Exit Sub
End If
do
elementID = session.ActiveWindow.GuiFocus.ID
elementLeft = Left(elementID, 50)
elementFinal = Right(elementLeft, 8)
'MsgBox(elementFinal)
Select Case elementFinal

Case "tabpSP05"
session.findById("wnd[0]").sendVKey 0

Case "tabpSP06"
session.findById("wnd[0]").sendVKey 0

Case "tabpSP07"
session.findById("wnd[0]").sendVKey 0

Case "tabpSP08"
session.findById("wnd[0]/tbar[0]/btn[0]").press

Case "tabpSP09"
session.findById("wnd[0]/tbar[0]/btn[0]").press

Case "tabpSP10"
session.findById("wnd[0]/tbar[0]/btn[0]").press

Case "tabpSP11"
session.findById("wnd[0]/tbar[0]/btn[0]").press

Case "tabpSP12"
session.findById("wnd[0]").sendVKey 0

Case "tabpSP13"
session.findById("wnd[0]/tbar[0]/btn[0]").press

Case "tabpSP14"
session.findById("wnd[0]/tbar[0]/btn[0]").press
If session.findbyid("wnd[0]/sbar").text="Check the consumption periods" Then
	session.findById("wnd[0]/tbar[0]/btn[0]").press
End If

Case "tabpSP15"
session.findById("wnd[0]/tbar[0]/btn[0]").press

Case "tabpSP16"
If ExcelSheet.cells(Row,10).value<>"" then
 session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2523/chkMARC-AUTRU").selected = True
Else
 session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2523/chkMARC-AUTRU").selected = False
End If
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2524/ctxtMPOP-PRMOD").text = ExcelSheet.cells(Row,12).value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2524/ctxtMARC-PERKZ").text = ExcelSheet.cells(Row,13).value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2525/txtMPOP-PERAN").text = ExcelSheet.cells(Row,14).value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2523/ctxtMPOP-KZINI").text = ExcelSheet.cells(Row,15).value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2523/txtMPOP-SIGGR").text = ExcelSheet.cells(Row,16).value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2523/ctxtMPOP-MODAV").text = ExcelSheet.cells(Row,17).value
session.findById("wnd[0]").sendVKey 0
'session.findById("wnd[0]").sendVKey 0

Case "tabpSP17"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP17/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2601/ctxtMARC-FEVOR").text = ExcelSheet.cells(Row,18).value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP17/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2601/ctxtMARC-SFCPF").text = ExcelSheet.cells(Row,19).value
session.findById("wnd[0]").sendVKey 0

Case "tabpSP18"
session.findById("wnd[0]").sendVKey 0

Case "tabpSP19"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
'session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'ExcelSheet.Cells(Row,20).Value = session.findById("wnd[0]/sbar").Text

Case "tabpSP20"
session.findById("wnd[0]/tbar[0]/btn[0]").press

Case "tabpSP21"
session.findById("wnd[0]/tbar[0]/btn[0]").press

Case "tabpSP22"
session.findById("wnd[0]/tbar[0]/btn[0]").press

Case "tabpSP23"
session.findById("wnd[0]").sendVKey 0
'session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'ExcelSheet.Cells(Row,20).Value = session.findById("wnd[0]/sbar").Text

Case "tabpSP24"
session.findById("wnd[0]/tbar[0]/btn[0]").press

Case "tabpSP25"
session.findById("wnd[0]/tbar[0]/btn[0]").press

Case "tabpSP26"
session.findById("wnd[0]/tbar[0]/btn[0]").press

Case "tabpSP27"
session.findById("wnd[0]/tbar[0]/btn[0]").press
End Select
On Error Resume Next
strWndTtl=session.findbyid("wnd[1]").text
On Error Goto 0
	If strWndTtl="Last data screen reached" Then
		session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
		ExcelSheet.Cells(Row,20).Value = session.findById("wnd[0]/sbar").Text
		intComplete=1
	End If

Loop While intComplete <>1
End Sub
