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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nme11"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtEINA-LIFNR").text = "10651791"
session.findById("wnd[0]/usr/ctxtEINA-MATNR").text = "ex-550-a"
session.findById("wnd[0]/usr/ctxtEINE-EKORG").text = "us31"
session.findById("wnd[0]/usr/ctxtEINE-WERKS").text = "500b"
session.findById("wnd[0]/tbar[0]/okcd").text = "/nme11"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtEINA-LIFNR").text = "10651791"
session.findById("wnd[0]/usr/ctxtEINA-MATNR").text = "ex-550-a"
session.findById("wnd[0]/usr/ctxtEINE-EKORG").text = "US31"
session.findById("wnd[0]/usr/ctxtEINE-WERKS").text = "500B"
session.findById("wnd[0]/usr/radRM06I-NORMB").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtEINA-LIFNR").text = "10064459"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[7]").press
session.findById("wnd[0]/usr/txtEINE-NORBM").text = "1"
session.findById("wnd[0]/usr/txtEINE-MINBM").text = "1"
session.findById("wnd[0]/usr/txtEINE-NETPR").text = "5.00"
session.findById("wnd[0]/usr/txtEINE-NETPR").setFocus
session.findById("wnd[0]/usr/txtEINE-NETPR").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/txtEINE-NETPR").text = "5,00"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[13]").press
session.findById("wnd[0]/usr/sub:SAPMM06I:0103/chkRM06I-SELKZ[6,0]").selected = true
session.findById("wnd[0]/usr/sub:SAPMM06I:0103/txtRM06I-LTEX1[7,11]").setFocus
session.findById("wnd[0]/usr/sub:SAPMM06I:0103/txtRM06I-LTEX1[7,11]").caretPosition = 40
session.findById("wnd[0]/tbar[0]/btn[11]").press
