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
session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/chkRM06W-FESKZ[8,0]").selected = true
session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-VDATU[0,0]").text = "10/1/2014"
session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-BDATU[1,0]").text = "12/31/9999"
session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,0]").text = "99529299"
session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-EKORG[3,0]").text = "US44"
session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/chkRM06W-FESKZ[8,0]").setFocus
session.findById("wnd[0]/tbar[0]/btn[11]").press
