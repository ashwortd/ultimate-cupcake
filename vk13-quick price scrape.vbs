Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim row,application, SapGuiAuto 

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
session.findById("wnd[0]").maximize
Dim objExcel, objWorkbook, objSheet,filelocation,file
'************Ask for data file
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
'****************
Function wndStatus()
	wndStatus = Session.findbyid("wnd[1]").text
	If  IsEmpty(wndStatus)Then
		wndStatus = ""
	End If 
End Function

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible=True
Set objWorkbook = objExcel.Workbooks.Open (file)
Set objSheet = objWorkbook.Worksheets(1)
row =1
x=0
For j= 1 To 11
	session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY").verticalScrollbar.position = x*15
	x=x+1
	For i = 0 To 15
		objSheet.Cells(Row,1).Value=Session.findbyid("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/ctxtKOMG-MATNR[0,"&i&"]").text
		objSheet.Cells(Row,2).Value=Session.findbyid("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtTEXT_DEFAULT-TEXT[2,"&i&"]").text
		objSheet.Cells(Row,3).Value=Session.findbyid("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[4,"&i&"]").text
		row=row+1
	Next 
Next 	