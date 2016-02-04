'*******ME0M Source List Change Fixed Vendor
'*Created by Derek Ashworth
'*4/27/2015
'****************************************
Dim FVIndicator,file,PMxRow,Row,testvar1,sbarchk
Dim ExcelApp,ExcelWorkbook,ExcelSheet
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
'*******************************************
'*Open Excel data file and set worksheet

Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible=True
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
PMxRow=0
session.findById("wnd[0]").maximize 
session.findById("wnd[0]/tbar[0]/okcd").text = "/nme0m"
session.findById("wnd[0]").sendVKey 0
Row=InputBox("Starting Row?")
Do
Call Main
Call getinfo
'session.findById("wnd[0]/tbar[0]/btn[11]").press
If sbarchk<>"bad" then
	session.findById("wnd[0]/tbar[0]/btn[3]").press
	session.findById("wnd[0]/tbar[0]/btn[3]").press
End If
sbarchk="good"
PMxRow=0
Row=Row+1
Loop Until ExcelSheet.Cells(Row,1).Value =""
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit

Sub Main
session.findById("wnd[0]/usr/ctxtW_MATNR-LOW").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/ctxtW_WERKS-LOW").text = ExcelSheet.Cells(Row,2).Value
session.findById("wnd[0]/tbar[1]/btn[8]").press
	If session.findbyid("wnd[0]/sbar").text="No selection possible" Then
 		sbarchk="bad"
 		ExcelSheet.cells(Row,3).value="Not on source list"
 		Exit Sub
 	End If
session.findbyId("wnd[0]/tbar[1]/btn[14]").press
End Sub

Sub getinfo
	If sbarcheck="bad" Then
		Exit Sub
	End if
Do
		testvar1=session.findbyid("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,"&PMxRow&"]").text
		If session.findbyid("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,"&PMxRow&"]").text="" Then
			PMxRow=0
			Exit Sub
		End If
		FVIndicator=session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/chkRM06W-FESKZ[8,"&PMxRow&"]").selected
		If FVIndicator="False" Then
			PMxRow=PMxRow+1
	 	Else
	 		ExcelSheet.Cells(Row,3).Value=session.findbyid("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,"&PMxRow&"]").text
	 		ExcelSheet.Cells(Row,4).Value=session.findbyid("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-EKORG[3,"&PMxRow&"]").text
			ExcelSheet.cells(row,5).value="Fixed"
		End if
	Loop Until FVIndicator="True"
End Sub
'session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/chkRM06W-FESKZ[8,0]").selected = 0
'session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/chkRM06W-FESKZ[8,0]").setFocus 
'session.findById("wnd[0]/tbar[0]/btn[11]").press 
'session.findById("wnd[0]/usr/lbl[3,5]").caretPosition = 2
'session.findById("wnd[0]").sendVKey 2
'session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/chkRM06W-FESKZ[8,0]").selected = -1
'session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/chkRM06W-FESKZ[8,0]").setFocus 
'session.findById("wnd[0]/tbar[0]/btn[11]").press 
'session.findById("wnd[0]/tbar[0]/btn[3]").press 
