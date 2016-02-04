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

Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet,strPARVW,Row,strStartPos,strStat
Dim strTest

Set shell = CreateObject( "WScript.Shell" )
defaultLocalDir = shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Desktop"
Set shell = Nothing

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
ExcelApp.Visible=True
Set ExcelWorkbook = ExcelApp.Workbooks.Open(file)
Set ExcelSheet = ExcelWorkbook.Worksheets(1)
Row=InputBox("Row to start at")
 If Row=""Then 
 	ExcelApp.Quit
	Set ExcelApp=Nothing
	Set ExcelWoorkbook=Nothing
	Set ExcelSheet=Nothing
	WScript.ConnectObject session,     "off"
	WScript.ConnectObject application, "off"
	WScript.Quit
 End if
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nxd03"
session.findById("wnd[0]").sendVKey 0
Do Until ExcelSheet.Cells(Row,1).Value = ""
session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[1]/usr/ctxtRF02D-BUKRS").text = "5000"
session.findById("wnd[1]/usr/ctxtRF02D-VKORG").text = "5013"
session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = "99"
session.findById("wnd[1]/usr/ctxtRF02D-SPART").text = "00"
session.findById("wnd[1]/tbar[0]/btn[0]").press
strStartPos=Session.findbyid("wnd[0]").guifocus.ID
strStartPos=Right(strStartPos,10)
If strStartPos="TITLE_MEDI" Then
	Session.findbyid("wnd[0]/tbar[1]/btn[27]").press
	Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05").select
End if
If strStartPos="KNVV-BZIRK" Then
	Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05").select
End if

vrc=session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN").visiblerowcount
For i = 0 To 14
strPARVW = session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,"&(i)&"]").text 
strTest=Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/txtRF02D-VTXTM[3,"&(i)&"]").text 
 If strPARVW ="Z4" Then
 	ExcelSheet.Cells(Row,2).Value=strTest
 ElseIf strPARVW="Z3" Then
 	ExcelSheet.Cells(Row,3).Value=strTest 
 ElseIf strPARVW="Z9" Then
 	ExcelSheet.Cells(Row,4).Value=strTest
 ElseIf strPARVW="ZA" Then
 	ExcelSheet.Cells(Row,5).Value=strTest
 ElseIf strPARVW="SP" Then
 	ExcelSheet.Cells(Row,6).Value=strTest
 ElseIf strPARVW="SH" Then
 	ExcelSheet.Cells(Row,7).Value=strTest
  End if
Next
session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN").verticalScrollbar.position = vrc
For i = 0 To 14
strPARVW = session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,"&(i)&"]").text 
strTest=Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/txtRF02D-VTXTM[3,"&(i)&"]").text 
 If strPARVW ="Z4" Then
 	ExcelSheet.Cells(Row,2).Value=strTest
 ElseIf strPARVW="Z3" Then
 	ExcelSheet.Cells(Row,3).Value=strTest 
 ElseIf strPARVW="Z9" Then
 	ExcelSheet.Cells(Row,4).Value=strTest
 ElseIf strPARVW="ZA" Then
 	ExcelSheet.Cells(Row,5).Value=strTest
 ElseIf strPARVW="SP" Then
 	ExcelSheet.Cells(Row,6).Value=strTest
 ElseIf strPARVW="SH" Then
 	ExcelSheet.Cells(Row,7).Value=strTest
  End if
Next

session.findById("wnd[0]/tbar[0]/btn[3]").press
'strStat=session.findById("wnd[0]/sbar").Text
'strStat=Left(strStat,7)
'If strStat="Country" Then
'	session.findById("wnd[0]").sendVKey 0
'End if
'ExcelSheet.Cells(Row,15).Value = session.findById("wnd[0]/sbar").Text
Row=Row+1
Loop 
ExcelApp.Quit
Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit
