Dim file,row,test,VdtyChek,valLine,valSel
Dim ExcelApp,ExcelWorkbook,ExcelSheet,condLine
Dim condChek
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
Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible=True
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
Set ExcelSheet = ExcelWorkbook.Worksheets(1)
Row=InputBox("Row to start at")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nme12"
session.findById("wnd[0]").sendVKey 0
Do
Call Main
row=row+1
Loop Until ExcelSheet.Cells(Row,1).Value =""

MsgBox("Script Complete")
ExcelWorkbook.Close(True)
ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"
		WScript.Quit
		
Sub Main
valLine=0
condLine=0
session.findById("wnd[0]/usr/ctxtEINA-LIFNR").text = ExcelSheet.Cells(Row,1).Value'"10063063"
session.findById("wnd[0]/usr/ctxtEINA-MATNR").text = ExcelSheet.Cells(Row,6).Value'"05R-210AD"
session.findById("wnd[0]/usr/ctxtEINE-EKORG").text = ExcelSheet.Cells(Row,3).Value'"us31"
session.findById("wnd[0]/usr/ctxtEINE-WERKS").text = ExcelSheet.Cells(Row,2).Value'"500B"
session.findById("wnd[0]/usr/ctxtEINE-WERKS").setFocus
session.findById("wnd[0]/usr/ctxtEINE-WERKS").caretPosition = 4
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
If ExcelSheet.cells(row,11).Value="0.01" Then
 	Call DelRecord 
 Else
 	Call ChgRecord
End if
End Sub

Sub DelRecord
Do
session.findById("wnd[1]/usr/tblSAPLV14ATCTRL_D0102/ctxtVAKE-DATAB[0,"&valLine&"]").setFocus
VdtyChek=session.findById("wnd[1]/usr/tblSAPLV14ATCTRL_D0102/ctxtVAKE-DATAB[0,"&valLine&"]").text
 If VdtyChek<>"__________" Then	
 	valLine=valLine + 1
  Else	
 	valSel=valLine-1
  End If
Loop Until VdtyChek="__________"

session.findById("wnd[1]/usr/tblSAPLV14ATCTRL_D0102/ctxtVAKE-DATAB[0,"&valSel&"]").setFocus
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201").getAbsoluteRow(0).selected = True
Do
session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/ctxtKONP-KSCHL[0,"&condLine&"]").setFocus
condChek=session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/ctxtKONP-KSCHL[0,"&condLine&"]").text
	If condChek<>"PB00" Then
		condLine=condLine+1
	  Else
		session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/ctxtKONP-KSCHL[0,"&condLine&"]").caretPosition = 0
		session.findById("wnd[0]/usr/btnFCODE_DLIN").press
	End If
Loop Until condChek="PB00"

session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
ExcelSheet.Cells(Row,23).Value = session.findById("wnd[0]/sbar").Text
End Sub

Sub ChgRecord
'session.findById("wnd[0]/tbar[1]/btn[8]").press 
session.findById("wnd[1]/tbar[0]/btn[7]").press 
session.findById("wnd[0]/usr/ctxtRV13A-DATBI").text = ExcelSheet.Cells(Row,16).Value'"12/31/2015"
session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/txtKONP-KBETR[2,0]").text =ExcelSheet.Cells(Row,11).Value '"5.44"
session.findById("wnd[0]/usr/ctxtRV13A-DATBI").setFocus 
session.findById("wnd[0]/usr/ctxtRV13A-DATBI").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
ExcelSheet.Cells(Row,23).Value = session.findById("wnd[0]/sbar").Text
End Sub

