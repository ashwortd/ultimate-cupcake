Dim ExcelApp
Dim ExcelWorkbook
Dim ExcelSheet
Dim file
Dim Row
Dim continueOn
Const updTab="Purchasing"

'****************************************
'Check for Logon status and connect to GUI
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
'******************************************
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
'**********************************
Function sbarStatus()
	sbarStatus = Session.findbyid("wnd[0]/sbar").text
End Function
'**********************************
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
ExcelApp.Visible=true
Row=InputBox("Row to start at")
session.findById("wnd[0]").maximize
'**********************************
Function reset()
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
	session.findById("wnd[0]").sendVKey 0
End Function
'***********************************
Function selectTabs(x)
	session.findById("wnd[1]/tbar[0]/btn[19]").press
	For i = 0 To 16	
		tabText=session.findbyid("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,"&i&"]").text
		If tabText = x Then
			session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(i).selected = true
			Exit For
		End If
		If i=16 Then
		 ExcelSheet.Cells(Row,3).Value = x&" tab not found"
		 continueOn=False
		End If
		 
	Next 
End Function
'************************************
Function RowCheck()
	RowCheck=ExcelSheet.Cells(Row,1).Value
End Function
'************************************
Function critUpdate()
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2313/chkMARC-KZKRI").selected = -1
End Function
'************************************
Function endScript()
	ExcelApp.Save
	Set ExcelApp=Nothing
	Set ExcelWorkbook=Nothing
	Set ExcelSheet=Nothing
	WScript.ConnectObject session,     "off"
    WScript.ConnectObject application, "off"
    WScript.Quit
End Function
'*************************************
	
continueOn=True		
reset
Do While RowCheck<>""
	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
	session.findById("wnd[0]").sendVKey 0
	selectTabs(updTab)
	If continueOn=True then
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,2).Value
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		critUpdate
		session.findById("wnd[0]/tbar[0]/btn[11]").press 
		ExcelSheet.Cells(Row,3).Value = sbarStatus
	Else
		session.findById("wnd[1]").close
	End If
	Row=Row+1
	continueOn=True
Loop
endScript


			
			
	
