'*******ME01 Purchase Info Record Create
'*Created by Derek Ashworth
'*2/17/2014
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

'*Open Excel data file and set worksheet
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
ExcelApp.Visible=True
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
'*******************************************
'*Maximize window and enter Tcode for Create source list
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nme01"
session.findById("wnd[0]").sendVKey 0
'*******************************************
Dim Row,StatusBar,SLBlank,i
Row=InputBox("Which row would you like to start with?","Create Source List")
	If Row = False Then
		Call Endscript
	End If
Do	
Call Main
Row=Row+1
Loop While ExcelSheet.Cells(Row,1).Value <>""

Call Endscript

Sub Endscript
MsgBox("Script completed")
		ExcelWorkbook.Close(True)
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"
		WScript.Quit
End Sub

Sub Main
session.findById("wnd[0]/usr/ctxtEORD-MATNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/ctxtEORD-WERKS").text = ExcelSheet.Cells(Row,2).Value
session.findById("wnd[0]/tbar[0]/btn[0]").press
StatusBar=session.findById("wnd[0]/sbar").Text
	If StatusBar <>"" Then
		Call ErrorSub
		StatusBar=""
		Exit Sub
	End If
i=0
Do
SLBlank = session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-VDATU[0,"&i&"]").text
'MsgBox SLBlank
i=i+1
Loop While SLBlank <>""
i=i-1
session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-VDATU[0,"&i&"]").text = ExcelSheet.Cells(Row,3).Value
session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-BDATU[1,"&i&"]").text = ExcelSheet.Cells(Row,4).Value
session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,"&i&"]").text = ExcelSheet.Cells(Row,5).Value
session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-EKORG[3,"&i&"]").text = ExcelSheet.Cells(Row,6).Value
If ExcelSheet.Cells(Row,7).Value<>"" then
	session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/chkRM06W-FESKZ[8,"&i&"]").selected = True
End if
session.findById("wnd[0]/tbar[0]/btn[11]").press
StatusBar=session.findById("wnd[0]/sbar").Text
	If StatusBar <>"Source list changed" Then
		ExcelSheet.Cells(Row,8).Value= StatusBar
		session.findbyid("wnd[0]/tbar[0]/btn[15]").press
		session.findById("wnd[0]/tbar[0]/okcd").text = "/nme01"
		session.findById("wnd[0]").sendVKey 0
		StatusBar=""
		Exit Sub
	End If
ExcelSheet.Cells(Row,8).Value = session.findById("wnd[0]/sbar").Text
End Sub

Sub ErrorSub
ExcelSheet.Cells(Row,8).Value= StatusBar
session.findById("wnd[0]/usr/ctxtEORD-MATNR").text = ""
session.findById("wnd[0]/usr/ctxtEORD-WERKS").text = ""
End Sub
