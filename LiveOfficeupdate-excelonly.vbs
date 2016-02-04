Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim row,application, SapGuiAuto 

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
Dim objExcel, objWorkbook, objSheet,filelocation,file
'************Ask for data file
file = ChooseFile(defaultLocalDir)
'MsgBox file

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
Function wndStatus()
	wndStatus = Session.findbyid("wnd[1]").text
	If  IsEmpty(wndStatus)Then
		wndStatus = ""
	End If 
End Function

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible=True
Set objWorkbook = objExcel.Workbooks.Open (file)
Set objSheet = objWorkbook.Worksheets("Sheet1")


Row=InputBox("Row Number?")

Session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nva23"
session.findById("wnd[0]").sendVKey 0
Do While objSheet.Cells(Row,7).Value <>""
If objSheet.Cells(Row,7).Value ="Sum:" Then
	Row=Row+1
	Exit Do
End If

session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = objSheet.Cells(Row,7).Value
Session.findById("wnd[0]").sendVKey 0

'Do 
'If wndStatus<>"" Then
'	Session.findbyid("wnd[1]").sendVKey 0
	
'End If
'Loop Until wndStatus = ""
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,0]").setFocus
session.findById("wnd[0]").sendVKey 2
objSheet.Cells(Row,41).Value=session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4451/ctxtVBAP-AWAHR").text
Session.findbyid("wnd[0]/tbar[1]/btn[19]").press
objSheet.Cells(Row,42).Value=session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4451/ctxtVBAP-AWAHR").text
Session.findbyid("wnd[0]/tbar[0]/btn[3]").press

Session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").select

session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectitem "0020","Column1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleclickitem "0020","Column1"
objSheet.Cells(Row,39).Value=(session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text)

session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectitem "0035","Column1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleclickitem "0035","Column1"
objSheet.Cells(Row,40).Value=(session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text)
Session.findbyid("wnd[0]/tbar[0]/btn[3]").press
Session.findbyid("wnd[0]/tbar[0]/btn[3]").press
Row=Row+1
Loop

'****Execl sheet cleanup
objSheet.Cells(1,39).Value="Document Description"
objSheet.cells(1,39).font.bold=True
objSheet.cells(1,39).font.colorindex = 2
objSheet.cells(1,39).interior.colorindex=11
objSheet.Columns(39).columnwidth=30
objSheet.Range("AL:AM").wraptext=true
objSheet.Cells(1,40).Value="Document Text Notes"
objSheet.cells(1,40).font.bold=True
objSheet.cells(1,40).font.colorindex = 2
objSheet.cells(1,40).interior.colorindex=11
objSheet.Columns(40).columnwidth=90
objSheet.columns.autofit


'****Disconnect SAP
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"