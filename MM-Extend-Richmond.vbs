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
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim messtxt,z,Row,iRow,strCurrentTab
Dim mrp3status,qmStatus,WndTTL

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

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
ExcelApp.Visible=True
Row=InputBox("Row to start at")
If TypeName(Row)="Empty" Then
	MsgBox("Goodbye!")
	ExcelWorkbook.Close(True)
	ExcelApp.Quit
	Set ExcelApp=Nothing
	Set ExcelWorkbook=Nothing
	Set ExcelSheet=Nothing
	WScript.ConnectObject session,     "off"
	WScript.ConnectObject application, "off"
	WScript.Quit
End If

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0
Do While ExcelSheet.Cells(Row,1).Value <>""
	Call GetStarted
	Call Tab4
	Call Tab5
	Call Tab7
	Call Tab9
	Call Tab12
	Call Tab13
	Call Tab14
	Call TabSave
	Row=Row+1
Loop

MsgBox("The end has come")
ExcelWorkbook.Close(True)
ExcelApp.Quit
Set ExcelApp=Nothing
Set ExcelWorkbook=Nothing
Set ExcelSheet=Nothing
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit

Sub GetStarted
Session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
Session.findById("wnd[0]").sendVKey 0
'Session.findById("wnd[0]").sendVKey 0
If session.findById("wnd[0]/sbar").Text<>"" Then
	ExcelSheet.Cells(Row,56)=session.findById("wnd[0]/sbar").Text
	Exit Sub
End If
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,2).Value
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = ExcelSheet.Cells(Row,3).Value
session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = ExcelSheet.Cells(Row,4).Value
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = ExcelSheet.Cells(Row,5).Value
'Session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text=""
'Session.findbyid("wnd[1]/usr/ctxtRMMG1-LGTYP").text=""
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").setFocus
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").caretPosition = 2
session.findById("wnd[1]").sendVKey 0

End Sub

Sub Tab4
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04").select
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").text = ExcelSheet.Cells(Row,6).Value
'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2184/tblSAPLMGD1TC_STEUERN/ctxtMG03STEUER-TAXKM[4,0]").text = ExcelSheet.Cells(Row,7).Value
End Sub

Sub Tab5
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05").select
'if strCurrentTab="tabpSP05" Then
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-MTPOS").text = ExcelSheet.Cells(Row,7).Value
'End If
End Sub



Sub Tab7
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP07").select
'If strCurrentTab="tabpSP07"Then
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP07/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLMGD1:2205/ctxtMARC-STAWN").text = ExcelSheet.Cells(Row,8).Value
'End If
End Sub


Sub Tab9
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09").select
'If strCurrentTab="tabpSP09"Then
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").text = ExcelSheet.Cells(Row,9).Value
'End If
End Sub


Sub Tab12
'If strCurrentTab="tabpSP12" Then
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").text = ExcelSheet.Cells(Row,10).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").text = ExcelSheet.Cells(Row,11).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").text = ExcelSheet.Cells(Row,12).Value
Session.findById("wnd[0]").sendVKey 0
'End If
End Sub

Sub Tab13
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13").select
'If strCurrentTab="tabpSP13" Then
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").text = ExcelSheet.Cells(Row,13).Value
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-PLIFZ").text = ExcelSheet.Cells(Row,14).Value
Smk=Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/ctxtMARC-FHORI").text
If Smk="" Then
	Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/ctxtMARC-FHORI").text ="001"
	ExcelSheet.Cells(Row,17).Value="Schedule margin key was blank set to 001"
End If
End Sub

Sub Tab14
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14").select
'If strCurrentTab="tabpSP14"Then
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/ctxtMARC-MTVFP").text=ExcelSheet.Cells(Row,15).Value
'End If
End Sub

Sub TabSave
	Session.findById("wnd[0]/tbar[0]/btn[11]").press
	ExcelSheet.Cells(Row,16).Value = session.findById("wnd[0]/sbar").Text
End Sub
