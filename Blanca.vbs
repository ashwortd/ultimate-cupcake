'-------open excel file for win7 or XP
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
'------------End of open excel-dma02
'----variables
Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row,PMxColumn
Set ExcelApp = CreateObject("Excel.Application")
'Next line sets the location of the excel spreadsheet
Set ExcelWorkbook = ExcelApp.Workbooks.Open(file)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
Set ExcelApp.Visible(True)
Row=InputBox("Row to start at")
PMxColumn=1
'------end of variables and excel open
'*********Original Script
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
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "cat2"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtTCATST-VARIANT").text = "3485a"
session.findById("wnd[0]/usr/ctxtCATSFIELDS-PERNR").text = "271800"
session.findById("wnd[0]/usr/ctxtCATSFIELDS-PERNR").caretPosition = 6
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtCATSFIELDS-INPUTDATE").text = "02.02.2015"
session.findById("wnd[0]/usr/ctxtCATSFIELDS-INPUTDATE").setFocus
session.findById("wnd[0]/usr/ctxtCATSFIELDS-INPUTDATE").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[5]").press

ExcelSheet.Cells(Row,13).Value = session.findById("wnd[0]/sbar").Text
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,"&PMxColumn&"]").text = ExcelSheet.Cells(Row,2).Value
'session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,2]").text = "3485-4POM1"
'session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,3]").text = "3485-4POM1"
'session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,4]").text = "3485-4POM1"
'session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,5]").text = "3485-4POM1"
'session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,6]").text = "3485-4POM1"
'session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,7]").text = "3485-4POM1"
'session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,8]").text = "3485-4POM1"
'session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,9]").text = "3485-4POM1"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,"&PMxColumn&"]").text = ExcelSheet.Cells(Row,3).Value
'session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,2]").text = "333485400001"
'session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,3]").text = "333485400003"
'session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,4]").text = "333485400004"
'session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,5]").text = "333485400011"
'session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,6]").text = "333485400012"
'session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,7]").text = "333485400013"
'session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,8]").text = "333485400010"
'session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,9]").text = "333485400008"

session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY1[10,"&PMxColumn&"]").text =ExcelSheet.Cells(Row,13).Value
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY1[10,2]").text = "0"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY1[10,3]").text = ",3"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY2[11,1]").text = ",08"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY2[11,2]").text = "0"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY2[11,3]").text = ",2"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY3[12,1]").text = ",08"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY3[12,2]").text = "0"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY3[12,3]").text = ",1"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY4[13,1]").text = ",08"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY4[13,2]").text = "0"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY4[13,3]").text = "0"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY5[14,1]").text = ",08"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY5[14,2]").text = "0"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY5[14,3]").text = "0"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,1]").setFocus
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,1]").caretPosition = 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "cat2"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtTCATST-VARIANT").text = "3485a"
session.findById("wnd[0]/usr/ctxtTCATST-VARIANT").setFocus
session.findById("wnd[0]/usr/ctxtTCATST-VARIANT").caretPosition = 5
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtCATSFIELDS-PERNR").text = "271875"
session.findById("wnd[0]/usr/ctxtCATSFIELDS-PERNR").caretPosition = 6
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtCATSFIELDS-INPUTDATE").text = "02.02.2015"
session.findById("wnd[0]/usr/ctxtCATSFIELDS-INPUTDATE").setFocus
session.findById("wnd[0]/usr/ctxtCATSFIELDS-INPUTDATE").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[5]").press
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,1]").text = "3485-4POM1"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,2]").text = "3485-4POM1"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,3]").text = "3485-4POM1"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,4]").text = "3485-4POM1"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,5]").text = "3485-4POM1"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,6]").text = "3485-4POM1"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,7]").text = "3485-4POM1"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,8]").text = "3485-4POM1"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,9]").text = "3485-4POM1"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,1]").text = "333485400002"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,2]").text = "333485400001"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,3]").text = "333485400003"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,4]").text = "333485400004"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,5]").text = "333485400011"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,6]").text = "333485400012"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,7]").text = "333485400013"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,8]").text = "333485400010"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RAUFNR[7,9]").text = "333485400008"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY1[10,1]").text = ",08"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY1[10,2]").text = ",5"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY1[10,3]").text = ",25"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY2[11,1]").text = ",08"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY2[11,2]").text = ",5"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY2[11,3]").text = ",25"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY3[12,1]").text = ",08"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY3[12,2]").text = ",5"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY3[12,3]").text = ",25"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY4[13,1]").text = ",08"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY4[13,2]").text = ",5"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY4[13,3]").text = ",25"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY5[14,1]").text = ",08"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY5[14,2]").text = ",5"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/txtCATSD-DAY5[14,3]").text = ",25"
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,1]").setFocus
session.findById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD/ctxtCATSD-RKOSTL[3,1]").caretPosition = 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
'*****end of original script


