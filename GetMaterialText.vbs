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
Dim strTabName(14)
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

Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim messtxt,z,Row,i,strSalesText,strPurchText,intPurchTab,intSalesTab,r
Dim t,v

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
ExcelApp.Visible=True
Row=InputBox("Row to start at")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm03"
session.findById("wnd[0]").sendVKey 0

Do
Call GetTabs
Call GetText
ExcelSheet.Cells(Row,2).Value = strSalesText
ExcelSheet.Cells(Row,3).Value = strPurchText
strSalesText="None"
strPurchText="None"
t=0
v=0
Session.findById("wnd[0]/tbar[0]/btn[3]").press
Row=Row+1
Loop Until ExcelSheet.Cells(Row,1).Value=""

MsgBox("The end has come")
ExcelWorkbook.Close(True)
ExcelApp.Quit
Set ExcelApp=Nothing
Set ExcelWorkbook=Nothing
Set ExcelSheet=Nothing
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit
Sub GetTabs
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 15
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[19]").press

For i=0 To 14
 r=i
	strTabName(i)= session.findbyid("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,"&i&"]").text
'		If strTabName(i)="Sales Text" Then
'			session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(r).selected = True
'			v=1
		If strTabName(i)="Purchase Order Text" Then
			session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(r).selected = True
			t=5
		End If
Next 

End Sub

Sub GetText


'Session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(intSalesTab).selected = True

session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,10]").setFocus
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,10]").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[0]").press

'If t+v=6 Then
On Error Resume next
Session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text =""' "5013"
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = ""'"01"
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ""
On Error Goto 0
'ElseIf t+v=5 Then
'	Session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = ""
'	session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = ""
'	session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ""
'ElseIf t+v=1 Then
'	Session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = "5013"
'	session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "01"
'	session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ""
'End if
Session.findById("wnd[1]").sendVKey 0
On Error Resume Next
If Session.findbyid("wnd[2]").text="Warning" Then
	Session.findbyid("wnd[2]").sendVkey 0
End If
strSalesText = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP08/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2121/cntlLONGTEXT_VERTRIEBS/shellcont/shell").text
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11").select
strPurchText = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").text
On Error Goto 0
End sub

