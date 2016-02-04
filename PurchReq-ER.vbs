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
Dim ExcelApp,ExcelWorkbook,ExcelSheet,ExcelSheet2
Dim Row,PartnerRow,PartnerName,SalesOrderNum,Row2
Dim file

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
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
Set ExcelSheet2= ExcelWorkbook.Worksheets("Data")
ExcelApp.Visible=True
session.findById("wnd[0]").maximize
PartnerRow=0
Row=2
Row2=1
Do

Call Main
	Do
	Call GetER
	'PartnerRow=0
		'session.findById("wnd[0]/tbar[1]/btn[19]").press
	Loop until session.findById("wnd[0]/sbar").Text ="There are no more items to be displayed"
	
Loop Until ExcelSheet.Cells(Row,1).Value =""

Sub Main
session.findById("wnd[0]/tbar[0]/okcd").text = "/nva03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = ExcelSheet.Cells(Row,1).Value
SalesOrderNum=ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/mbar/menu[2]/menu[2]/menu[10]").select
session.findById("wnd[0]/tbar[1]/btn[19]").press
Row=Row+1
End Sub

Sub GetER
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4353/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/cmbGV_FILTER").key = "PARPE"
PartnerName=session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4353/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,"&PartnerRow&"]").text
ExcelSheet.Cells(Row,13).Value = session.findById("wnd[0]/sbar").Text
'PartnerRow=0
Select Case PartnerName
Case "Employee respons."
	ExcelSheet2.Cells(Row2,1).Value=SalesOrderNum 
	ExcelSheet2.Cells(Row2,2).Value=session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4013/txtVBAP-POSNR").text
	ExcelSheet2.Cells(Row2,4).Value=session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4353/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/txtGVS_TC_DATA-REC-NAME1[3,"&PartnerRow&"]").text
	'session.findById("wnd[0]/tbar[1]/btn[19]").press
	PartnerRow=PartnerRow+1
	'PartnerRow=0
	'Row2=Row2+1
Case "Customer ServiceRep"
	ExcelSheet2.Cells(Row2,1).Value=SalesOrderNum 
	ExcelSheet2.Cells(Row2,2).Value=session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4013/txtVBAP-POSNR").text
	ExcelSheet2.Cells(Row2,5).Value=session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4353/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/txtGVS_TC_DATA-REC-NAME1[3,"&PartnerRow&"]").text
	PartnerRow=PartnerRow+1
Case ""
	session.findById("wnd[0]/tbar[1]/btn[19]").press
	PartnerRow=0
	Row2=Row2+1
	Exit Sub
Case Else
	PartnerRow=PartnerRow+1
	'session.findById("wnd[0]/tbar[1]/btn[19]").press
End Select
End Sub
