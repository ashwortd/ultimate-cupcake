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
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nsqvi"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "PO_Report"
session.findById("wnd[0]/usr/btnP1").press
session.findById("wnd[0]/usr/ctxtSP$00003-LOW").text = "5013"
session.findById("wnd[0]/usr/ctxtSP$00003-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtSP$00004-LOW").text = ""
session.findById("wnd[0]/usr/ctxtSP$00005-LOW").text = ""
session.findById("wnd[0]/usr/ctxtSP$00006-LOW").text = ""
session.findById("wnd[0]/usr/ctxtSP$00007-LOW").text = ""
session.findById("wnd[0]/usr/rad%ALV").setFocus
session.findById("wnd[0]/usr/rad%ALV").select
session.findById("wnd[0]/tbar[0]/btn[0]").press
   WScript.ConnectObject session,     "off"
   WScript.ConnectObject application, "off"
   WScript.Quit