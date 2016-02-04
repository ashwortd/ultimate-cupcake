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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = "gp-769"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").columns.elementAt(-1).width = 2
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,15]").setFocus
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,15]").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "500b"
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/chkMARC-CCFIX").selected = true
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/ctxtMARC-ABCIN").text = "d"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/chkMARC-CCFIX").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-VPRSV").text = "v"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-VPRSV").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-VPRSV").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
