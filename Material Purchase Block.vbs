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
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim Row
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
Set ExcelSheet = ExcelWorkbook.Worksheets(1)
ExcelApp.Visible=True

session.findById("wnd[0]").maximize
Row=InputBox("Starting Row?")

Do Until ExcelSheet.Cells(Row,1).Value =""
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(6).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(10).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(11).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(12).selected = true
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,5).Value
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = ExcelSheet.Cells(Row,6).Value
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-MMSTA").text = ExcelSheet.Cells(Row,7).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-PLIFZ").text = ExcelSheet.Cells(Row,2).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").text = ExcelSheet.Cells(Row,3).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:2503/btnMARC_NOTE").press
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,1]").text = "Resourced effective immediately November 2014; to be purchased from"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,2]").text = "Weatherly Castings and NOT Sioux City Foundry"
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
ExcelSheet.Cells(Row,8).Value = session.findById("wnd[0]/sbar").Text
Row=Row+1
Loop
		MsgBox("The end has come")
		ExcelWorkbook.Close=True	
		Set ExcelApp=Nothing
		Set ExcelWorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
  		WScript.ConnectObject application, "off"
		WScript.Quit
