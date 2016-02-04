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
Dim ExcelApp,ExcelWorkbook,ExcelSheet,Win2Text
Dim strPARVW(8),strKTONR(8),strVTXTM(8)

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
ExcelApp.Visible=true
Set ExcelWorkbook = ExcelApp.Workbooks.Open(file)
Set ExcelSheet = ExcelWorkbook.Worksheets(1)
ExcelApp.Visible=true
Row=InputBox("Row to start at")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nxd02"
session.findById("wnd[0]").sendVKey 0
Do Until ExcelSheet.Cells(Row,1).Value = ""
session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[1]/usr/ctxtRF02D-BUKRS").text = "5000"
session.findById("wnd[1]/usr/ctxtRF02D-VKORG").text = "5013"
session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = "01"
session.findById("wnd[1]/usr/ctxtRF02D-SPART").text = "00"
session.findById("wnd[1]/usr/ctxtRF02D-SPART").setFocus
session.findById("wnd[1]/usr/ctxtRF02D-SPART").caretPosition = 2
session.findById("wnd[1]/tbar[0]/btn[0]").press
On Error Resume Next
Win2Text=session.findbyid("wnd[2]").text
On Error Goto 0
If Win2Text="Error" Then
	Session.findbyid("wnd[2]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = "99"
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	Win2Text="none"
End If
On Error Resume Next
Win2Text=session.findbyid("wnd[2]").text
On Error Goto 0
If Win2Text="Warning" Then
	Session.findbyid("wnd[2]/tbar[0]/btn[0]").press
	Win2Text="none"
End If
Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05").select
session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN").verticalScrollbar.position = 3
session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN").verticalScrollbar.position = 6
strPARVW(0) = Session.findbyid("wnd[0]/usr/subSUBKOPF:SAPMF02D:7003/txtINT_KNA1-NAME1").Text
strPARVW(1) = session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,2]").text 
strPARVW(2) = session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,3]").text
strPARVW(3) = session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,4]").text
strPARVW(4) = session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,5]").text
strPARVW(5) = session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,6]").text
strPARVW(6) = session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,7]").text
strPARVW(7) = session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,8]").text
strPARVW(8) = session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,9]").text
strKTONR(0) = Session.findbyid("wnd[0]/usr/subSUBKOPF:SAPMF02D:7003/txtINT_KNA1-ORT01").text
strKTONR(1) = Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtRF02D-KTONR[2,2]").text
strKTONR(2) = Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtRF02D-KTONR[2,3]").text
strKTONR(3) = Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtRF02D-KTONR[2,4]").text
strKTONR(4) = Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtRF02D-KTONR[2,5]").text
strKTONR(5) = Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtRF02D-KTONR[2,6]").text
strKTONR(6) = Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtRF02D-KTONR[2,7]").text
strKTONR(7) = Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtRF02D-KTONR[2,8]").text
strKTONR(8) = Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtRF02D-KTONR[2,9]").text
strVTXTM(1) = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/txtRF02D-VTXTM[3,2]").text 
strVTXTM(2) = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/txtRF02D-VTXTM[3,3]").text 
strVTXTM(3) = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/txtRF02D-VTXTM[3,4]").text 
strVTXTM(4) = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/txtRF02D-VTXTM[3,5]").text 
strVTXTM(5) = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/txtRF02D-VTXTM[3,6]").text 
strVTXTM(6) = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/txtRF02D-VTXTM[3,7]").text 
strVTXTM(7) = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/txtRF02D-VTXTM[3,8]").text 
strVTXTM(8) = Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/txtRF02D-VTXTM[3,9]").text 
'Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtRF02D-KTONR[2,9]").setFocus
'session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtRF02D-KTONR[2,9]").caretPosition = 5
'session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[3]").press
ExcelSheet.Cells(Row,4).Value=strPARVW(0)
ExcelSheet.Cells(Row,5).Value=strKTONR(0)
ExcelSheet.Cells(Row,6).Value=strPARVW(1)
ExcelSheet.Cells(Row,9).Value=strPARVW(2)
ExcelSheet.Cells(Row,12).Value=strPARVW(3)
ExcelSheet.Cells(Row,15).Value=strPARVW(4)
ExcelSheet.Cells(Row,18).Value=strPARVW(5)
ExcelSheet.Cells(Row,21).Value=strPARVW(6)
ExcelSheet.Cells(Row,24).Value=strPARVW(7)
ExcelSheet.Cells(Row,27).Value=strPARVW(8)
ExcelSheet.Cells(Row,7).Value=strKTONR(1)
ExcelSheet.Cells(Row,10).Value=strKTONR(2)
ExcelSheet.Cells(Row,13).Value=strKTONR(3)
ExcelSheet.Cells(Row,16).Value=strKTONR(4)
ExcelSheet.Cells(Row,19).Value=strKTONR(5)
ExcelSheet.Cells(Row,22).Value=strKTONR(6)
ExcelSheet.Cells(Row,25).Value=strKTONR(7)
ExcelSheet.Cells(Row,28).Value=strKTONR(8)
ExcelSheet.Cells(Row,8).Value=strVTXTM(1)
ExcelSheet.Cells(Row,11).Value=strVTXTM(2)
ExcelSheet.Cells(Row,14).Value=strVTXTM(3)
ExcelSheet.Cells(Row,17).Value=strVTXTM(4)
ExcelSheet.Cells(Row,20).Value=strVTXTM(5)
ExcelSheet.Cells(Row,23).Value=strVTXTM(6)
ExcelSheet.Cells(Row,26).Value=strVTXTM(7)
ExcelSheet.Cells(Row,29).Value=strVTXTM(8)
Row=Row+1
Loop 
ExcelApp.Quit
Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing