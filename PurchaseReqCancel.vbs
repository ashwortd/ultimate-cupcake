If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   On Error Resume Next
   Set connection = application.Children(0)
   If Err.Number <> 0 Then
      MsgBox("You are not properly logged into SAP."& chr(13) &"Please login and try again."& chr(13) & chr(13) &"Script terminating...")
      WScript.Quit
   End If
   On Error Goto 0
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize

Dim ex,wb,ws
Dim Row,strRelStat,strRepeat,delError,a,b
Row =InputBox("Please row to start with in excel sheet","Starting Position")


Call Main
Sub Main()
'************Ask for data file
Set objDialog = CreateObject("UserAccounts.CommonDialog")

objDialog.Filter = "VBScript Data Files |*.xls;*.xlsx;*.xlsm|All Files|*.*"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\Scripts"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
    Wscript.Quit
'Else
'    Wscript.Echo objDialog.FileName
End If
'****************
	Set ex=CreateObject("Excel.Application")
	Set wb=ex.Workbooks.open(objDialog.FileName)
	Set ws= wb.Sheets("Sheet2")
	ex.Visible=True
	
	Do While ws.cells(Row,1).value<>""
	Call ME54N
	Call CancelRelease
	Row=Row+1
	Loop
	
	
	Set ex = Nothing
	Set wb = Nothing
	Set ws = Nothing
	MsgBox("The requested process has been completed." & chr(13) & chr(13) & "Thank you.")					
	WScript.ConnectObject session,     "off"
	WScript.ConnectObject application, "off"
	WScript.Quit
	
End Sub

Sub ME54N
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nME54N"
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	session.findbyid("wnd[0]/tbar[1]/btn[17]").press
	session.findbyid("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-BANFN").text=ws.cells(Row,1).value
	session.findbyid("wnd[1]/tbar[0]/btn[0]").press
	If session.findById("wnd[0]/sbar").Text="Purchase requisition does not exist" Then
		ws.cells(Row,4).value="Purchase Req does not exist"
		Row=Row+1
		Call ME54N
	End if

End Sub

Sub CancelRelease
session.findbyid("wnd[0]").sendvkey 28
session.findbyid("wnd[0]").sendvkey 29
session.findbyid("wnd[0]").sendvkey 27
On Error Resume Next
session.findById("wnd[0]/tbar[1]/btn[9]").press
On Error Goto 0 

strRepeat ="TRUE"
Do While strRepeat ="TRUE"
On Error Resume Next 
strRelStat=session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT12/ssubTABSTRIPCONTROL1SUB:SAPLMERELVI:1101/cntlRELEASE_INFO_ITEM/shellcont/shell").getcellvalue(0,"FUNCTION")
If Err.Number<>0 Then
	delError="FALSE"
	Call TrashError
End If
	
a=session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST").text
If 	strRelStat="@2W\QCancel release@" Then
	session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT12/ssubTABSTRIPCONTROL1SUB:SAPLMERELVI:1101/cntlRELEASE_INFO_ITEM/shellcont/shell").currentCellColumn = "FUNCTION"
	session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT12/ssubTABSTRIPCONTROL1SUB:SAPLMERELVI:1101/cntlRELEASE_INFO_ITEM/shellcont/shell").clickCurrentCell
	ws.cells(Row,3).value = session.findById("wnd[0]/sbar").Text
	session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT002").press
	b=session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST").text
	If a=b Then
		session.findbyid("wnd[0]/tbar[0]/btn[11]").press
		ws.cells(Row,4).value="No Data changed"
		strRepeat="False"
		Exit Sub
	End If
End If
If strRelStat<>"@2W\QCancel release@" Then
	session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT002").press 
	b=session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST").text
	If a=b Then
		session.findbyid("wnd[0]/tbar[0]/btn[11]").press
		ws.cells(Row,4).value="No Changes"
		strRepeat="False"
		Exit Sub
	End if
End If
ws.Cells(Row,4).Value = "Checked" 'session.findById("wnd[0]/sbar").Text
Loop

End Sub

Sub TrashError
	Do While delError = "FALSE"
	session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT002").press
	z=session.findbyid("wnd[0]").getfocus
	MsgBox(z)
	On Error Resume next
	strRelStat=session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT12/ssubTABSTRIPCONTROL1SUB:SAPLMERELVI:1101/cntlRELEASE_INFO_ITEM/shellcont/shell").getcellvalue(0,"FUNCTION")
	delerror="TRUE"
	If Err.Number<>0 Then
		Err.Clear 
		delError="FALSE"
	End If
	
	On Error Goto 0
	Loop
End sub
	
	
	
	