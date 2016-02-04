
'/\  ___\   /\  __ \   /\ "-./  \   /\  == \    /\  ___\   /\  ___\   /\  ___\   /\ \/ /    
'\ \ \____  \ \ \/\ \  \ \ \-./\ \  \ \  _-/    \ \ \__ \  \ \  __\   \ \  __\   \ \  _"-.  
' \ \_____\  \ \_____\  \ \_\ \ \_\  \ \_\       \ \_____\  \ \_____\  \ \_____\  \ \_\ \_\ 
'  \/_____/   \/_____/   \/_/  \/_/   \/_/        \/_____/   \/_____/   \/_____/   \/_/\/_/ 
'Scripted 8/19/2014
'Derek Ashworth
'
'Script uses direct file from Ardmore, it creates a table with the data provided, script also adds 
'three colums (E,L,M). E is used for creating a tracking number that fits into the PMX Field. L is
'used for taking UPS shipments under $30 and making them $30. M is just used as a confirmation that 
'the sales order has been changed.                                                                                          




If Not IsObject(application) Then
   On Error Resume next
   Set SapGuiAuto  = GetObject("SAPGUI")
        If Err.Number<>0 Then
   		MsgBox("You are not connected to PMx, please connect and try again")
   		
   		WScript.Quit
   	  End If
   On Error Goto 0
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

Dim ExcelSheet,ExcelApp,ExcelWorkbook
Dim Row,PMxRow,PMx3,i,MtrLn,j,h,stat1
Dim stat2,Formula(1),intNewRow,Wnd1TTL
Dim MessText,MainTtl

Const xlCellTypeLastCell = 11
Const xlSrcRange=1
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
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
Set ExcelSheet = ExcelWorkbook.Worksheets("ZPC billed")
Row=InputBox("Starting Row?")
If ExcelSheet.cells(Row,12)="" Then
	Call SheetSetup
End If

Sub SheetSetup
Set objrange = ExcelSheet.UsedRange
objrange.SpecialCells(xlCellTypeLastCell).Activate
intNewRow= ExcelApp.ActiveCell.Row
ExcelApp.ActiveSheet.ListObjects.add xlSrcRange, ExcelApp.Range("A1:L"&intNewRow),,XlYes

Set objrange = ExcelApp.Range("E1").EntireColumn
Formula(0)="=Left(C2,5)&""-""&D2"
Formula(1)="=IF([@Carrier]=""UPS (UPSN)"",IF([@[Amount Billed]]<30,30,[@[Amount Billed]]),[@[Amount Billed]])"
objrange.Insert(xlShiftToRight)
With ExcelSheet
	.range("E2").select
	.range("E1").value="Tracking Number"
	.range("E2").formula=Formula(0)
	.range("L1").Value="Amount Invoiced"
	.range("L2").select
	.range("M1").value = "Status"
	.range("L2").formula=Formula(1)
	.range("M1").select
	
	
End With
End Sub

session.findById("wnd[0]/tbar[0]/okcd").text = "/nva02"'<-----Change sales order
session.findById("wnd[0]").sendVKey 0
Do
Call Main
ExcelSheet.Cells(Row,13).Value = session.findById("wnd[0]/sbar").Text
Row=Row+1
Loop Until ExcelSheet.cells(Row,1)=""
ExcelWorkbook.Close(True)
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit

Sub Main
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text =ExcelSheet.Cells(Row,8).Value
session.findById("wnd[0]").sendVKey 0
On Error Resume Next
'session.findById("wnd[1]").sendVKey 0
Wnd1TTL=session.findbyid("wnd[1]").text
MessText=session.findbyid("wnd[1]/usr/txtMESSTXT2").text
If Left(Wnd1TTL,4)="Help" Then
	ExcelSheet.Cells(Row,13).Value = "Order Closed"
	session.findbyid("wnd[1]/tbar[0]/btn[5]").press
	Wnd1TTL="none"
	Exit Sub
 ElseIf Left(MessText,4)="Over" Then
 	ExcelSheet.Cells(Row,14).Value = "Warning - Not Processed"
	session.findbyid("wnd[1]/tbar[0]/btn[5]").press
	MessText="none"
	Exit Sub
End If
session.findById("wnd[1]").sendVKey 0
MainTtl=session.findbyid("wnd[0]").text
MainTtl=Left(MainTtl,21)
MainTtl=Right(MainTtl,2)
If MainTtl="BP" Then
	ExcelSheet.Cells(Row,14).Value = "Warning - BP Order freight not added"
	MainTtl="none"
	session.findbyid("wnd[0]/tbar[0]/btn[3]").press
	Exit Sub
End If

On Error Goto 0
Call FindRow2
'session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POAN").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,"&PMxRow&"]").text = "ship-handling"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,"&PMxRow&"]").text = "1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-ARKTX[5,"&PMxRow&"]").text = ExcelSheet.cells(Row,5).value
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtKOMV-KBETR[8,"&PMxRow&"]").text = ExcelSheet.cells(Row,12).value
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtKOMV-KBETR[8,"&PMxRow&"]").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtKOMV-KBETR[8,"&PMxRow&"]").caretPosition = 16
session.findById("wnd[0]").sendVKey 0
stat1=session.findById("wnd[0]/sbar").Text
On Error Resume Next
stat2=session.findbyId("wnd[1]").text
On Error Goto 0
If stat2 ="Information" Then
	ExcelSheet.cells(Row,14)=session.findbyid("wnd[1]/usr/txtMESSTXT1").text
	session.findbyid("wnd[1]/tbar[0]/btn[0]").press
	stat2="nope"
End if	
If Left(stat1,8)="No goods" then
	session.findById("wnd[0]").sendVKey 0
End if
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,"&PMxRow&"]").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,"&PMxRow&"]").caretPosition = 9
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KOAN").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,1]").text = "pr00"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,1]").text = ExcelSheet.cells(Row,12).value
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,1]").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,1]").caretPosition = 16
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KOAN").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/ctxtKOMV-KSCHL[1,1]").text = "ycmc"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,1]").text = ExcelSheet.cells(Row,10).value
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,1]").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,1]").caretPosition = 16
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05/ssubSUBSCREEN_BODY:SAPLV69A:6201/subSUBSCREEN_PUSHBUTTONS:SAPLV69A:1000/btnBT_KONY").press
session.findById("wnd[1]/usr/lbl[1,5]").setFocus
session.findById("wnd[1]/usr/lbl[1,5]").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[0]").press
WScript.Sleep(2000)
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[1]").close
session.findById("wnd[2]/usr/btnBUTTON_2").press
ExcelSheet.Cells(Row,15).Value = session.findById("wnd[0]/sbar").Text
End Sub

Sub FindRow2
PMx3=True
h=1
i=0
j=session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").visiblerowcount
	Do Until PMx3=False
		MtrLn=session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,"&i&"]").text
		If i=(j-1) Then
			session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").verticalScrollbar.position = j*h
			h=h+1
			i=-1
		End If
		PMxRow=i
		If MtrLn="" Then
			If PMxRow=-1 Then
			PMxRow=0
			End If
			
		PMx3=False
		End If
		i=i+1
	Loop
End Sub
