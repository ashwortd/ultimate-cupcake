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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nvk11"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRV13A-KSCHL").text = "zs13"
session.findById("wnd[0]/usr/ctxtRV13A-KSCHL").caretPosition = 4
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[1,0]").select
session.findById("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtKOMG-BUKRS").text = "5000"
session.findById("wnd[0]/usr/ctxtKOMG-VKORG").text = "5013"
session.findById("wnd[0]/usr/ctxtKOMG-VTWEG").text = "01"
session.findById("wnd[0]/usr/ctxtKOMG-SPART").text = "00"
session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/ctxtKOMG-KUNAG[0,0]").text = "10084127"
session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[4,0]").text = "5"
session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[4,0]").setFocus
session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[4,0]").caretPosition = 16
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[26]").press
session.findById("wnd[0]/usr/sub:SAPLV70T:0101/txtLV70T-LTX01[0,37]").text = "Below $200K usage factor"
session.findById("wnd[0]/usr/sub:SAPLV70T:0101/txtLV70T-LTX01[0,37]").caretPosition = 24
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
