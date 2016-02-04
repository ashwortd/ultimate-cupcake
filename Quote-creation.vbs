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
session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = "zqca"
session.findById("wnd[0]/usr/ctxtVBAK-VKORG").text = "5013"
session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").text = "01"
session.findById("wnd[0]/usr/ctxtVBAK-SPART").text = "00"
session.findById("wnd[0]/usr/ctxtVBAK-VKBUR").text = "5003"
session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").text = "Z27"
session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").setFocus
session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").caretPosition = 3
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text = "v:Test2"
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").text = "10080825"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/ctxtVBAK-BNDDT").text = "7/9/2014"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtVBAP-POSNR[0,0]").text = "10"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,0]").text = "1g-885"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtRV45A-KWMENG[2,0]").text = "1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtRV45A-KWMENG[2,0]").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtRV45A-KWMENG[2,0]").caretPosition = 19
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/lbl[4,10]").setFocus
session.findById("wnd[1]/usr/lbl[4,10]").caretPosition = 6
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/cntlTREE_CONTAINER/shellcont/shell").expandNode "          1"
session.findById("wnd[1]/usr/cntlTREE_CONTAINER/shellcont/shell").topNode = "          1"
session.findById("wnd[1]/usr/cntlTREE_CONTAINER/shellcont/shell").expandNode "          2"
session.findById("wnd[1]/usr/cntlTREE_CONTAINER/shellcont/shell").topNode = "          1"
session.findById("wnd[1]/usr/cntlTREE_CONTAINER/shellcont/shell").expandNode "          8"
session.findById("wnd[1]/usr/cntlTREE_CONTAINER/shellcont/shell").topNode = "          1"
session.findById("wnd[1]/usr/cntlTREE_CONTAINER/shellcont/shell").expandNode "          9"
session.findById("wnd[1]/usr/cntlTREE_CONTAINER/shellcont/shell").topNode = "          1"
session.findById("wnd[1]/usr/cntlTREE_CONTAINER/shellcont/shell").selectedNode = "          9"
session.findById("wnd[1]/usr/cntlTREE_CONTAINER/shellcont/shell").doubleClickNode "          9"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtRV45A-KWMENG[2,5]").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtRV45A-KWMENG[2,5]").caretPosition = 19
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtRV45A-KWMENG[2,5]").showContextMenu
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtRV45A-KWMENG[2,5]").showContextMenu
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,5]").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,5]").caretPosition = 0
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,5]").showContextMenu
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtRV45A-KWMENG[2,0]").text = "100"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").setFocus
session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").caretPosition = 0
