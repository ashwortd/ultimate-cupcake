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
Dim ExcelApp,ExcelWorkbook,ExcelSheet,vrc,vrc2,SOQuantity,v,EmpResName

row2=InputBox("Starting Row?","Purchase Req Assignments")
Set wshShell = WScript.CreateObject( "WScript.Shell" )
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )

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
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
Set ExcelSheet = ExcelWorkbook.Worksheets(2)
ExcelApp.Visible=True


Do
i=1
Call starthere
Call info1
'Call PartnerSelect
row2=row2+1
Loop Until ExcelSheet.Cells(row2,17).Value =""

Sub starthere
Session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nVA03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = ExcelSheet.Cells(row2,1).Value
session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/btn").press
session.findById("wnd[1]/usr/cmbAKT_VERSION").key = "Basic setting"
session.findById("wnd[1]/usr/cmbAKT_VERSION").setFocus
session.findById("wnd[1]/tbar[0]/btn[11]").press
End Sub

Sub PartnerSelect

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\"&"0"&i).select
j=0
'CSRName=ExcelSheet.Cells(13,2).Value


Do While j<25
	PartnerName=session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,"&(j)&"]").value
 	'MsgBox(PartnerName)
		If PartnerName = "Customer ServiceRep" Then
			Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,"&(j)&"]").setfocus
			Session.findById("wnd[0]").sendVKey 2
			CSRName=Session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0313/txtADDR1_DATA-NAME1").text
			
			'MsgBox(CSRName)
			ExcelSheet.Cells(row2,3).Value=CSRName
				If CSRName="" Then
					ExcelSheet.Cells(row2,3).Value="Not Listed"
				End If
		   	Session.findbyid("wnd[1]/tbar[0]/btn[12]").press
	    
	    ElseIf PartnerName="Employee respons." Then
			Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,"&(j)&"]").setfocus
			Session.findById("wnd[0]").sendVKey 2
			EmpResName=session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0313/txtADDR1_DATA-NAME1").text
	    	ExcelSheet.Cells(row2,4).Value=EmpResName
	    	Session.findbyid("wnd[1]/tbar[0]/btn[12]").press
	    	
	    End If
	
	j=j+1
Loop
End Sub

Sub info1
Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
Do
test4 = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\"&"0"&i).text
boolLoopAgain=False
If test4="Partners" Then
	Call PartnerSelect
	boolLoopAgain=True
End If
i=i+1
Loop While boolLoopAgain = false
End Sub		

session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
session.findById("wnd[0]").sendVKey 0	
ExcelApp.Workbooks.Close(True)
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWorkbook=Nothing
		Set ExcelSheet=Nothing
		
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit