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
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/chkMARA-QMPUR").selected = true
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-SSQSS").text = "PMX0003"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-QZGTP").text = "USQP"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-QZGTP").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-QZGTP").caretPosition = 4
session.findById("wnd[0]/tbar[0]/btn[0]").press
