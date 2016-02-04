Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim row,application, SapGuiAuto
Dim file, objExcel,objWorkbook,objSheet,currentTab

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
'***************
'open data file
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
'*************
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible=True
Set objWorkbook = objExcel.Workbooks.Open (file)
Set objSheet = objWorkbook.Worksheets("Sheet1")

Row=InputBox("Row Number?")


session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
session.findById("wnd[0]").sendVKey 0

Do 
Call Main

row=row+1
Loop Until objSheet.Cells(row,1).Value=""
'****Disconnect SAP
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"

Sub Main ()
If objSheet.Cells(row,7).Value<>"" Then
	Exit Sub
End If

Session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = objSheet.Cells(row,1).Value
session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").text = objSheet.Cells(row,1).Value
'session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").setFocus
'session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").caretPosition = 8
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(13).selected = false
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(14).selected = false
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = objSheet.Cells(row,3).Value
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "0001"
session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = "5013"
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "01"
session.findById("wnd[1]/usr/ctxtRMMG1_REF-WERKS").text = "500B"
session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGORT").text = "g001"
session.findById("wnd[1]/usr/ctxtRMMG1_REF-VKORG").text = "5013"
session.findById("wnd[1]/usr/ctxtRMMG1_REF-VTWEG").text = "01"
session.findById("wnd[1]/usr/ctxtRMMG1_REF-VTWEG").setFocus
session.findById("wnd[1]/usr/ctxtRMMG1_REF-VTWEG").caretPosition = 2
session.findById("wnd[1]").sendVKey 0
On Error Resume Next
	objSheet.Cells(row,5).Value=session.findById("wnd[2]/usr/txtMESSTXT1").text
	If objSheet.Cells(row,5).Value="Material already maintained for this" Then
		Session.findById("wnd[2]/tbar[0]/btn[0]").press:
		Session.findById("wnd[1]/tbar[0]/btn[12]").press:
		'messtxt=0
		Exit Sub
	End If
On Error Goto 0
Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP12").select
Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").text=objSheet.Cells(row,5).Value
Session.findById("wnd[0]/tbar[0]/btn[11]").press
currentTab=Session.activewindow.guifocus.ID
currentTab= Left(currentTab,50)
currentTab= Right(currentTab,8)
If currentTab="tabpSP13" then
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGFSB").text = "0001"
	session.findById("wnd[0]").sendVKey 0
	Session.findById("wnd[0]/tbar[0]/btn[11]").press
ElseIf currentTab="tabpSP23" Then
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/chkMARA-QMPUR").selected = true
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-SSQSS").text = "PMX0003"
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-QZGTP").text = "USQP"
	Session.findById("wnd[0]/tbar[0]/btn[11]").press
End If
objSheet.Cells(row,5).Value=session.findById("wnd[0]/sbar").Text
End Sub

