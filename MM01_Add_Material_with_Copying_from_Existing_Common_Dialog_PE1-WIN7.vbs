'*****************************************************************************************************************************************************
'	Purpose: ADD Material Master records WITH copying from existing plus updating fields from excel file.  PE1 version
'	NOTE:  When adding materials without copying, SAP won't allow the script to update the Moving average price which is 
'   why the copy method was chosen instead		 
'
'	Input: Excel file (Material_Master_ADDS_via_MM01.xlsx)
'	Refer to excel template for columns
'
'	Created on: 2_14-14
'	Created by: ESMarion
'	
'   NOTES:
'	Removed all 'on error goto 0's
'   The Excel template file has a flag in column 75 that needs to indicate whether or not the material has already been extended to a SLOC.  
'   If yes, then this script will skip the row but it will get picked up in the MRP 4 script.   
'   That is, the same template/same data is used in the MRP4 script which will add the MRP4 tab for materials where column 75 = True which skipping the other rows.  
'   It updates columns BX,BY + BZ with status messages and the add date.
'   Lastly, this script skips dist channel 99 but 99 may need to be added/updated via MVKE script
'*****************************************************************************************************************************************************
 
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
Dim ExcelApp,ExcelWorkbook,ExcelSheet,currenttab,ExcelFile
Dim Row, StatusMsg1,StatusMsg2,StatusMsg3,StatusMsg4,CorrectFile

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
'************Ask for data file
'Set objDialog = CreateObject("UserAccounts.CommonDialog")

'objDialog.Filter = "VBScript Data Files |*.xls;*.xlsx;*.xlsm|All Files|*.*"
'objDialog.FilterIndex = 1
'objDialog.InitialDir = "C:\Scripts"
'intResult = objDialog.ShowOpen
 
'If intResult = 0 Then
'    Wscript.Quit
'Else
'    ExcelFile = objDialog.FileName
'End If
'****************


Set ExcelApp = CreateObject("Excel.Application")
'Set ExcelWorkbook = ExcelApp.Workbooks.Open("O:\CustSvc\Parts\Inventory\Inventory Scripts\Excel files for Script execution\Relocation of 50GC to 50GD extensions.xlsx")
'Set ExcelWorkbook = ExcelApp.Workbooks.Open("C:\scripts\MM01_ADD_Material_with_Copying_from_nesc TO 50GD r2.xlsx")
Set ExcelWorkbook = ExcelApp.Workbooks.Open(file)
MsgBox "The Excel File = " & file
Set ExcelSheet = ExcelWorkbook.Worksheets(1)

ExcelApp.Visible=True

CorrectFile = MsgBox ( "Is this the correct excel file?",VBYESNO,"Verify Excel File" )

If CorrectFile = 7 Then WScript.Quit

'User is prompted to enter first row of Excel spreadsheet to be read - usually row 2, Column 1 (A) reverse of excel col/row

Row=InputBox("Enter Starting Row, usually Row 2","Remember to Set Cursor in Input Box First",2)-1

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"  
session.findById("wnd[0]").sendVKey 0

Row = Row + 1

Do Until ExcelSheet.Cells(Row,1).Value = ""  'Do until row x column 1 (A) is null

''MsgBox ExcelSheet.Cells(Row,6).Value & "-" & ExcelSheet.Cells(Row,75).Value 
 
If ExcelSheet.Cells(Row,8).Value = "01" and (ExcelSheet.Cells(Row,75).Value = "False" Or ExcelSheet.Cells(Row,75).Value = "FALSE") THEN 

'Clear Material number value
Session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ""
Session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").text = ""
	
'Populate Material number into text box
Session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = Excelsheet.cells(Row,1).value		'Material number (i.e. MPS-8800345)
session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = Excelsheet.cells(Row,2).value	
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = Excelsheet.cells(Row,3).value	
'If using this script to ADD, then comment out the next 3 lines which provides the 'copy from material' 
'************
Session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").text = Excelsheet.cells(Row,4).value	
session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").setFocus
session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").caretPosition = 7
'****************

'Select which tabs to add'
Session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[20]").press 'Select all tabs
session.findById("wnd[1]/tbar[0]/btn[0]").press

'Fill in the New material info
Session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = Excelsheet.cells(Row,5).value	'"500D" 'copy to plant
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = Excelsheet.cells(Row,6).value	'"0001" 'sloc
session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = Excelsheet.cells(Row,7).value	'"5013" 'org
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = Excelsheet.cells(Row,8).value	'"01"  'dist channel
session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = Excelsheet.cells(Row,9).value	'"U04" ' from MLGN table - needs to be calculated based on new plant 
session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = Excelsheet.cells(Row,10).value '"001"  'storage type

'****************************
'Fill in the copy from info.  If script it being used to simply add, then comment out these rows.
session.findById("wnd[1]/usr/ctxtRMMG1_REF-WERKS").text = Excelsheet.cells(Row,11).value '"500C" 'Copy from plant
session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGORT").text = Excelsheet.cells(Row,12).value '"0001"
session.findById("wnd[1]/usr/ctxtRMMG1_REF-VKORG").text = Excelsheet.cells(Row,13).value '"5013"
session.findById("wnd[1]/usr/ctxtRMMG1_REF-VTWEG").text = Excelsheet.cells(Row,14).value '"01"
session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGNUM").text = Excelsheet.cells(Row,15).value '"U03"
session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGTYP").text = Excelsheet.cells(Row,16).value '"001"
'****************************

On Error Resume next

Session.findById("wnd[1]/tbar[0]/btn[0]").press

IF session.findById("wnd[0]/sbar").Text <> "The material already exists and will be extended" then 
   StatusMsg4 = "Material " & Excelsheet.cells(Row,1).value & " already exists" 
   'MsgBox statusmsg4
   session.findById("wnd[2]/tbar[0]/btn[0]").press 'click ok on error dialog box
   session.findById("wnd[1]/tbar[0]/btn[12]").press 'click x to close the new plant/sloc info dialog
   'On Error Goto 0
Else 'do process 

'On Error Goto 0
'MsgBox "sales Org Select"
'tabpSP04 Sales Org 
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04").select
IF session.findById("wnd[0]/sbar").Text = "Sales: Sales Org. Data 1 not selectable here (data already created)" Then
   StatusMsg4 = "Material skipped: Sales Data 1 already exists: " &  Excelsheet.cells(Row,1).value & ", " & Excelsheet.cells(Row,5).value & "/" & Excelsheet.cells(Row,6).value & "/" & Excelsheet.cells(Row,8).value
   'Write status msg to excel and skip to next record
Else       
   session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").text = Excelsheet.cells(Row,17).value '"50GD"
   session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2184/tblSAPLMGD1TC_STEUERN/ctxtMG03STEUER-TAXKM[4,0]").text = Excelsheet.cells(Row,18).value '"1" 'tax classification
   On Error Resume next
   session.findById("wnd[0]").sendVKey 0
   If session.findById("wnd[0]/sbar").Text = "Material not yet created in supplying plant" Then 
      session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").text = "" ' delivery plant
      StatusMsg1 = "Delivery Plant error"
   End If
   'On Error Goto 0

'tabpSP05 Sales Org 2
'MsgBox "Before  sales Org 2 Select"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KONDM").text = Excelsheet.cells(Row,20).value '"UG"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KTGRM").text = Excelsheet.cells(Row,21).value '"10"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-MTPOS").text = Excelsheet.cells(Row,22).value '"ZVOR"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-PRODH").text = Excelsheet.cells(Row,19).value '"212040103160102"

'tabpSP06 Sales General
'MsgBox "Before Sales General Select"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2161/ctxtMARC-MTVFP").text = Excelsheet.cells(Row,23).value '"04"
' Trans Grp is read only Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2162/ctxtMARA_TRAGR").text = "Z001" 'Excelsheet.cells(Row,24).value '"Z001"
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2162/ctxtMARC-LADGR").text = Excelsheet.cells(Row,25).value '"0002"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").text = Excelsheet.cells(Row,26).value '"5000000019"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-SERNP").text = Excelsheet.cells(Row,27).value '""

Session.findById("wnd[0]").sendVKey 0
'tabpSP07 Foreign Trade Export
session.findById("wnd[0]").sendVKey 0
'tabpSP08 Sales Text 
session.findById("wnd[0]").sendVKey 0

'tabpSP09 Purchasing
'MsgBox "Before Purchasing Select"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").text = Excelsheet.cells(Row,29).value '"ELA"
session.findById("wnd[0]").sendVKey 0
'tabpSP10 Foreign Trade Import
session.findById("wnd[0]").sendVKey 0
'tabpSP11 Purchase Order Text
'session.findById("wnd[0]").sendVKey 0

'tabpSP12 MRP1
'MsgBox "Before MRP1 Select"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").text = Excelsheet.cells(Row,30).value '"PD"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").text = Excelsheet.cells(Row,31).value '"001"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").text = Excelsheet.cells(Row,32).value '"EX"
'session.findById("wnd[0]").sendVKey 0

'tabpSP13 MRP2
'MsgBox "Before MRP2 Select"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").text = Excelsheet.cells(Row,34).value '"F"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-SOBSL").text = Excelsheet.cells(Row,35).value '"40"
'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGPRO").text = "0001" Prod Stor Loc
'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGFSB").text = "" 'Storage Loc for EP
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGPRO").text = excelsheet.cells(Row,40).value '"0001" Prod Stor Loc
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGFSB").text = excelsheet.cells(Row,41).value'  "" 'Storage Loc for EP
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-DZEIT").text = Excelsheet.cells(Row,36).value '"0"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-PLIFZ").text = Excelsheet.cells(Row,37).value '"14"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-WEBAZ").text = Excelsheet.cells(Row,38).value '"1"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/ctxtMARC-FHORI").text = Excelsheet.cells(Row,39).value '"000"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2486/txtMARC-EISBE").text = Excelsheet.cells(Row,42).value ' "0"
'session.findById("wnd[0]").sendVKey 0

'tabpSP14 MRP3 Skip this tab when copying from an existing material since SAP demands entry of the forward consumption period for unknown reasons
'MsgBox "Before MRP3 Select"
'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14").select
'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-STRGR").text = Excelsheet.cells(Row,43).value '"Z1"
'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-VRMOD").text = Excelsheet.cells(Row,44).value '"1"
'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/txtMARC-VINT1").text = Excelsheet.cells(Row,45).value '"30"
'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-MISKZ").text = Excelsheet.cells(Row,46).value '"1"
'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").text = Excelsheet.cells(Row,47).value '"15"
'when manually entering material, sap forces the forward consumption mode, vint2 to be populated so popped it with vint1 which is usually 1
'Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/txtMARC-VINT2").text = "0" 'Excelsheet.cells(Row,45).value '"1"
'session.findById("wnd[0]").sendVKey 0

'tabpSP14 MRP4
'MsgBox "Before MRP4 Select"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2495/ctxtMARC-SBDKZ").text = Excelsheet.cells(Row,48).value '"2"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB6:SAPLMGD1:2498/ctxtMARD-DISKZ").text = Excelsheet.cells(Row,50).value '"2"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB6:SAPLMGD1:2498/ctxtMARD-LSOBS").text = Excelsheet.cells(Row,49).value '"40"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB6:SAPLMGD1:2498/txtMARD-LMINB").text = Excelsheet.cells(Row,51).value '"1"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB6:SAPLMGD1:2498/txtMARD-LBSTF").text = Excelsheet.cells(Row,52).value '"1"
'session.findById("wnd[0]").sendVKey 0

'tabpSP16 Forecasting
'MsgBox "Before Forecasting Select"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16").select
If session.findById("wnd[0]/sbar").text ="The effective-out date is in the past" Then
	session.findById("wnd[0]").sendVKey 0
End If
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2525/txtMPOP-ANZPR").text = Excelsheet.cells(Row,55).value '"12"
'session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2524/ctxtMPOP-PRMOD").text = Excelsheet.cells(Row,54).value '"N" 'mpop table 
'session.findById("wnd[0]").sendVKey 0

'tabpSP17 Workscheduling
'MsgBox "Before Workscheduling Select"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP17").select
'''session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP17/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2601/ctxtMARA-MEINS").text = Excelsheet.cells(Row,53).value 'UOM
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP17/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2601/chkMARC-XCHPF").selected = False 'batch mgmt Row, 58
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP17/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2601/ctxtMARC-FEVOR").text = Excelsheet.cells(Row,56).value '"001"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP17/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2601/ctxtMARC-SFCPF").text = Excelsheet.cells(Row,57).value '"Z00002"
'session.findById("wnd[0]").sendVKey 0

'tabpSP19 plant data stor 1 - (there no tab 18)
'MsgBox "Before Plant data stor 1 Select"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/chkMARC-CCFIX").selected = True 'row, 60
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/txtMARD-LGPBE").text = Excelsheet.cells(Row,61).value '""
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/ctxtMARC-ABCIN").text = Excelsheet.cells(Row,59).value '"D"
session.findById("wnd[0]").sendVKey 0

'tabpSP20 plant data stor 2
session.findById("wnd[0]").sendVKey 0
'tabpSP21 Warehouse Mgmt 1
session.findById("wnd[0]").sendVKey 0
'tabpSP22 Warehouse Mgmt 2
'session.findById("wnd[0]").sendVKey 0

'tabpSP23 Quality Mgmt  
'MsgBox "Before Quality Select"
On Error Resume next
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/chkMARA-QMPUR").selected = true
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-SSQSS").text = "PMX0003"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-QZGTP").text = "USQP"
On Error Resume next
session.findById("wnd[0]").sendVKey 0
IF session.findById("wnd[0]/sbar").Text = "Plants exist in which you have not specified a control key" then 
  'MsgBox "Error - Plants exist in which you have not specified a control key"
  StatusMsg2 = "Quality Mgmt error - Control key set to PMX0000"
  session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/chkMARA-QMPUR").selected = false
  session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-SSQSS").text = "PMX0000"
  session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-QZGTP").text = ""
Else
  'MsgBox "Other Quality tab error"
  'StatusMsg2 = "Quality Mgmt error - May need manual update"
  session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/chkMARA-QMPUR").selected = true
  session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-SSQSS").text = "PMX0003"
  session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-QZGTP").text = "USQP"
End If

'msgbox "Before Accting 1 select"
  
'tabpSP24 Accting 1
'MsgBox "Before Accting 1 Select"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-BKLAS").text = Excelsheet.cells(Row,62).value '"4200"

'MsgBox "Before Setting PC to V"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-VPRSV").text = Excelsheet.cells(Row,63).value '"V or S"

'MsgBox "Before setting MAP - MAP = " & Excelsheet.cells(Row,65).value 
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-VERPR").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-VERPR").caretPosition = 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-VERPR").text = Excelsheet.cells(Row,65).value 
'MsgBox "After setting MAP to Excel value "

'MsgBox "Before setting Std Price to MAP"
'Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-STPRS").text = Excelsheet.cells(Row,65).value 'set std price to MAP

On Error Resume Next 
session.findById("wnd[0]").sendVKey 0
'Trap moving average price error'
IF session.findById("wnd[0]/sbar").Text = "With price control V, enter a moving average price" Then
   'MsgBox "Before resetting after error"
   'Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-STPRS").value = Excelsheet.cells(Row,65).value 'set std Price to MAP
   session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-VERPR").value = Excelsheet.cells(Row,65).value 'Mov Avg Price
   'MsgBox "After resetting after error"
   'On Error Goto 0
   session.findById("wnd[0]").sendVKey 0
    
End if   
'On Error Goto 0

'tabpSP25 Accting 2
'session.findById("wnd[0]").sendVKey 0
'tabpSP26 Costing 1
'MsgBox "Before Costing 1 Select"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/chkMBEW-EKALR").selected = true
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/chkMBEW-HKMAT").selected = true
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/ctxtMBEW-HRKFT").text = Excelsheet.cells(Row,66).value '"B007"
'****Profit ctr is added on sales general tablsession.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/ctxtMARC-PRCTR").text = Excelsheet.cells(Row,26).value'"5000000019"
session.findById("wnd[0]").sendVKey 0

'tabpSP27 Costing 2
'MsgBox "Before costing send Key 0 which will force save of record"
session.findById("wnd[0]").sendVKey 0


'SAVE  
'************ 

''MsgBox "Before Save"  
On Error Resume next

session.findById("wnd[1]/usr/btnSPOP-OPTION1").press 'YES to save
''''''''session.findById("wnd[1]/usr/btnSPOP-OPTION2").press 'NO to save and exit transaction completely
'''''''''''''session.findById("wnd[1]/usr/btnSPOP-OPTION_CAN").press 'Cancel save but stay on screen
       'StatusMsg3 = "Material created for " & Excelsheet.cells(Row,1).value & ", " & Excelsheet.cells(Row,5).value & "/" & Excelsheet.cells(Row,6).value & "/" & Excelsheet.cells(Row,8).value 
       StatusMsg4 = session.findById("wnd[0]/sbar").text ' capture Record save error message
       'MsgBox "Status Message 4: " & StatusMsg4
       If StatusMsg4  = "Choose a valid function" Or StatusMsg4 = "Invalid GUI input data: FOCUS DATA"Then StatusMsg4 = "Error: Material not created"
''MsgBox "After Save"
'***********    

End If 'Sales Org Data 1 already exists
End If 'if material exists 
Else 'ExcelSheet.Cells(Row,75).Value = "TRUE"
   'MsgBox "Material Skipped"
   If ExcelSheet.Cells(Row,75).Value = "True" Or ExcelSheet.Cells(Row,75).Value = "TRUE"  then  
      StatusMsg4 = "Material Skipped: Add via MRP 4 script"
   Else 
      StatusMsg4 = "99 Record skipped - may need updating via MVKE update script" 
   End if
End If 'If NEWMAT3 record for Plant already exists - material should be added via MRP ADD script

    ExcelSheet.Cells(Row,70).Value =  StatusMsg1
    ExcelSheet.Cells(Row,71).Value =  StatusMsg2
    ExcelSheet.Cells(Row,72).Value =  StatusMsg3
    ExcelSheet.Cells(Row,73).Value =  StatusMsg4
    ExcelSheet.Cells(Row,74).Value = Date()
    StatusMsg1 = ""
    StatusMsg2 = ""
    StatusMsg3 = ""
    StatusMsg4 = ""
            

Row = Row + 1	'Move to next row of the Excel data file

On Error Resume next
'session.findById("wnd[0]").resizeWorkingPane 169,20,false
'session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"  
'session.findById("wnd[0]").sendVKey 0

Loop

If ExcelSheet.cells(Row,1).value = "" Then MsgBox "End of File"

ExcelApp.Workbooks.save
ExcelApp.Quit

Set ExcelApp=Nothing
Set ExcelWorkbook=Nothing
Set ExcelSheet=Nothing
wscript.quit
