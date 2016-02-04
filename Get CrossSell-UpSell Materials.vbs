Dim file,Row,ExcelApp,ExcelWorkbook,ExcelSheet
Dim Material,Mat_Desc,UpSell_Mat,UpSell_Mat_Desc,strUserName
Dim vrc, i, vrc_multiplier
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

Set wshShell = WScript.CreateObject( "WScript.Shell" )
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
'Open Excel file
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
'Set ExcelSheet = ExcelWorkbook.Worksheets("Script Sheet")
ExcelApp.Visible = True
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
'Access VB13 Transaction
Session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nVB43"
session.findById("wnd[0]").sendVKey 0
'Enter Material Determination type
session.findbyid("wnd[0]/usr/ctxtD000-KSCHL").text ="ZCS1"
session.findById("wnd[0]").sendVKey 0
'Set key combination
session.findbyid("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[1,0]").select
session.findbyid("wnd[1]/tbar[0]/btn[0]").press
'run query
session.findbyid("wnd[0]/usr/ctxtF003-LOW").text=""
session.findbyid("wnd[0]/usr/ctxtF003-HIGH").text=""
session.findbyid("wnd[0]/tbar[1]/btn[8]").press
'Get visible row count from table
Row = 2
vrc_multiplier = 1
vrc = session.findbyid("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY").visiblerowcount
'call subroutine to get info
Do
Call get_page
session.findbyid("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY").verticalScrollbar.position=vrc*vrc_multiplier
vrc_multiplier =vrc_multiplier +1
Loop Until Right(Material,1)="_"
Call end_script



Sub get_page
For i = 0 To vrc-1
	Material =session.findbyid("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKOMGD-MATNR[0,"&i&"]").text
	'/app/con[0]/ses[0]/wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKOMGD-MATNR[0,0]
	Mat_Desc =session.findbyid("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/txtRV130-TEXTL[1,"&i&"]").text
	UpSell_Mat =session.findbyid("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKONDD-SMATN[2,"&i&"]").text
	UpSell_Mat_Desc =session.findbyid("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/txt*MAAPV-ARKTX[3,"&i&"]").text
	ExcelSheet.Cells(Row,1).Value = Material
	ExcelSheet.Cells(Row,2).Value = Mat_Desc
	ExcelSheet.Cells(Row,3).Value = UpSell_Mat
	ExcelSheet.Cells(Row,4).Value = UpSell_Mat_Desc
	Row=Row+1
Next
End Sub

Sub end_script
ExcelSheet.SaveAs "C:\Users\"&strUserName&"\Desktop\"&SDNum(0)&"-Cross_Upsell_Parts.xlsx",51
		MsgBox("The end has come")
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"
		WScript.Quit
End sub