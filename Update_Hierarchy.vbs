'File: Update_Hierarchy.vbs
'Author: Derek Ashworth
'Edit Date: 05/17/2016
' Column A = Material Number | Column B = Sales Org | Column C = Dist Channel | Column D = New Hierarchy | PMx Response

Option Explicit
Const strTab1 = "Sales: Sales Org. Data 2"
Dim file, excelApp, excelWorkbook, excelWorksheet
Dim Row, SapGuiAuto, application, connection, session
Dim strTabText, fptemp,objWMIService,colItems,objItem
Dim shell, ex

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

Function findPosition
	fptemp = 0
	Do 
	 strTabText = session.findbyid("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,"&fptemp&"]").text
	  If strTabText=strTab1 Then
	    findPosition=fptemp
	  Else
	  	fptemp=fptemp+1
	  End If
	Loop Until findPosition=fptemp
 End Function

Function sbarStatus()
	sbarStatus = Session.findbyid("wnd[0]/sbar").text
End Function

Function newCommand(tcode)
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/n"&tcode
session.findById("wnd[0]").sendVKey 0
End Function

file = ChooseFile("\\winfile02\data\CustSvc\Parts\Pmx Scripting\Script Data")
MsgBox file

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
Row=InputBox("Starting Row?")
Set excelApp = CreateObject("Excel.Application")
Set excelWorkbook = excelApp.Workbooks.Open (file)
Set excelWorksheet = excelWorkbook.Worksheets("Sheet1")
excelApp.Visible=True
newCommand("mm02")
Do Until excelWorksheet.Cells(Row,1).Value =""
	Call UpdateHierarchy
	Row=Row+1
Loop

newCommand("")
excelWorkbook.Save=True
excelApp.Quit
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
Set connection= Nothing
Set session= Nothing

Sub UpdateHierarchy()
	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text =excelWorksheet.Cells(Row,1).Value
	
	session.findById("wnd[0]").sendVKey 0
	
	session.findById("wnd[1]/tbar[0]/btn[19]").press
	session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(findPosition).selected = True
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = excelWorksheet.Cells(Row,2).Value
	session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = excelWorksheet.Cells(Row,3).Value
	Session.findById("wnd[1]").sendVKey 0
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05").select
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-PRODH").text = excelWorksheet.Cells(Row,4).Value
	session.findById("wnd[0]/tbar[0]/btn[11]").press
	excelWorksheet.Cells(Row,5).Value=sbarStatus
End Sub