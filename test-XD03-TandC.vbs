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
Dim juicy

juicy=session.findbyid("wnd[0]/usr/txtRSTXT-TXBORDER").text
juicy=Right(juicy,16)
juicy=Left(juicy,2)
MsgBox(juicy)


'session.findById("wnd[0]").maximize
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA").verticalScrollbar.position = 30
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA").verticalScrollbar.position = 0
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,23]").caretPosition = 69
'session.findById("wnd[0]").sendVKey 2
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA").columns.elementAt(2).width = 72
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA").verticalScrollbar.position = 1
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA").verticalScrollbar.position = 2
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA").verticalScrollbar.position = 3
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA").verticalScrollbar.position = 4
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA").verticalScrollbar.position = 5
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA").verticalScrollbar.position = 6
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA").verticalScrollbar.position = 7
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA").verticalScrollbar.position = 8
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA").verticalScrollbar.position = 9
