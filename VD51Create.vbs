On Error Resume Next
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
   
   If Err.Number<>0 Then
   MsgBox("You are not properly logged into SAP."& Chr(13) &"Please login and try again."& Chr(13)& Chr(13)&"Script terminating...")
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
Dim ex,wb,ws,row

Call Main
Set ex=WScript.CreateObject("Excel.Application")
ex.Visible=False
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
Set wb = ex.Workbooks.Open (objDialog.FileName)
Set ws = wb.Worksheets(Sheet1)
session.findbyid("wnd[0]/tbar[0]/okcd").text="/nVD51"
session.findbyid("wnd[0]/tbar[0]/btn[0]").press
row=InputBox("Which row would you like to start with?","Starting Point")
session.findById("wnd[0]").maximize
While ws.cells(row,1)value<>""

session.findById("wnd[0]/usr/ctxtMV10A-KUNNR").text = ws.cells(row,1).value
session.findById("wnd[0]/usr/ctxtMV10A-VKORG").text = "5013"
session.findById("wnd[0]/usr/ctxtMV10A-VTWEG").text = "01"
session.findbyid("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tblSAPMV10ATC_CU_MA/ctxtMV10A-MATNR[0,0]").text = "Ship-handling"
session.findById("wnd[0]/usr/tblSAPMV10ATC_CU_MA/ctxtMV10A-MATNR[0,1]").text = "Ciqd"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press

Wend
