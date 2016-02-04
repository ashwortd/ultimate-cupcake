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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nvb43"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[1,0]").select
session.findById("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtF003-LOW").text = ""
session.findById("wnd[0]/usr/ctxtF003-LOW").setFocus
session.findById("wnd[0]/usr/ctxtF003-LOW").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY").verticalScrollbar.position = 15
