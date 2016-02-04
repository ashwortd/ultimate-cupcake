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
'Option Explicit
Dim SapGuiAuto,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath,application
Dim ExcelApp,ExcelWorkbook,ExcelSheet,vrc,vrc2,SOQuantity,v
Dim messtxt,z,Row,SDRow,Itemno,SDNum(4),SAPRow,check1
Dim workingRow,i,strRejectionReason,SDCount,wshShell,strUserName

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
'Set objDialog = CreateObject("UserAccounts.CommonDialog")

'objDialog.Filter = "VBScript Data Files |*.xls;*.xlsx|All Files|*.*"
'objDialog.FilterIndex = 1
'objDialog.InitialDir = "C:\Scripts"
'intResult = objDialog.ShowOpen
 
'If intResult = 0 Then
'    Wscript.Quit
'Else
'    Wscript.Echo objDialog.FileName
'End If
'****************

Set wshShell = WScript.CreateObject( "WScript.Shell" )
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
check1=false
Do While check1=false
SOQuantity=InputBox("How many sales orders are on this invoice?","Quantity of Invoices")
If SOQuantity<4 Then
	check1=True
End If 
loop

For v = 0 To (SOQuantity-1)
SDNum(v)=InputBox("Enter Sales Document Number "&v+1&":","Document Number")

Next 
'For v=0 To (SOQuantity)-1
'MsgBox(SDNum(v))
'Next
SDCount=0
SDRow=0
Row=11
Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible=True
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)'("O:\CustSvc\Temp\Customs Document Script\Customs Document.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
For v=0 To (SOQuantity)-1
Session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nVA03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = SDNum(v)
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/btn").press
session.findById("wnd[1]/usr/cmbAKT_VERSION").key = "Basic setting"
session.findById("wnd[1]/usr/cmbAKT_VERSION").setFocus
session.findById("wnd[1]/tbar[0]/btn[11]").press
vrc=Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").VisibleRowCount
'vrc=Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT").VisibleRowCount
SDRow=0
SAPRow=0
If ExcelSheet.Cells(14,5).Value ="" Then	
	ExcelSheet.Cells(14,5).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBKD-KURSK[84,0]").text
End if
ExcelSheet.Shapes(1).TextFrame.Characters.Text = Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/txtKUAGV-TXTPA").text
ExcelSheet.Cells(1,7).Value = Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-VBELN").text
ExcelSheet.Cells(8,7).Value = Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text

'MsgBox(vrc)
While Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,"&(SDRow)&"]").text <>"__________________"
'While Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]
'While Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,"&(SDRow)&"]").text <>"__________________"
If Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,"&(SDRow)&"]").text="SUM" Then
	SDCount=SDCount+1
	SDRow=SDRow+1
	SAPRow=SAPRow+1
End if
If (SDRow+1) = vrc Then
	Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").verticalScrollbar.position=SDCount
   	SDRow=0
   	SAPRow=-1
End If
If ExcelSheet.Cells(Row,1).Value=ExcelSheet.cells(Row-1,1).value Then
   ExcelSheet.rows(Row).delete
End If

strRejectionReason=Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/cmbVBAP-ABGRU[27,"&(SDRow)&"]").text
'MsgBox(strRejectionReason)
If strRejectionReason <> "Item captured in error" Then
	ExcelSheet.Cells(Row,1).Value = SDNum(v)&"/"&Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,"&(SDRow)&"]").text
	'ExcelSheet.Cells(100,100).Value= Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,"&(SDRow)&"]").text
	ExcelSheet.Cells(Row,3).Value = "'"&Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,"&(SDRow)&"]").text
	ExcelSheet.Cells(Row,2).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,"&(SDRow)&"]").text
	'ExcelSheet.Cells(Row,3).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-KDMAT[6,"&(SDRow)&"]").text
	ExcelSheet.Cells(Row,5).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-NETPR[17,"&(SDRow)&"]").text
	ExcelSheet.Cells(Row,4).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-ARKTX[5,"&(SDRow)&"]").text
	ExcelSheet.Cells(Row,8).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WAERK[21,"&(SDRow)&"]").text
	ExcelSheet.Cells(Row,6).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WAERK[21,"&(SDRow)&"]").text
	ExcelSheet.Cells(Row,7).Value = "=iferror(B"&(Row)&"*E"&(Row)&",0)"'Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-NETWR[21,"&(SDRow)&"]").text
	ExcelSheet.Rows(Row+1).EntireRow.Insert
		
	Row=Row+1
	'SDRow=SDRow+1
	'SAPRow=SAPRow+1
End if
'	MsgBox("vrc="&vrc&" Row="&Row&" SDRow="&SDRow&" SAPRow="&SAPRow)
'End If
'Row=Row+1
SDCount=SDCount+1
SDRow=SDRow+1
SAPRow=SAPRow+1
Wend
Next

ExcelSheet.SaveAs "C:\Users\"&strUserName&"\Desktop\"&SDNum(0)&"-Customs Document.xlsx",51
		MsgBox("The end has come")
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"
		WScript.Quit

