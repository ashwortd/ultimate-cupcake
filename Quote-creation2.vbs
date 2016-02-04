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
session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZOPPORID").text = "1-239146231"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/txtVBAK-KTEXT").text = "Order Description"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_WWBSS").text = "TSe"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_TERANGA").text = "7667"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_PRLINE").text = "b0"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_SCOPE").text = "sp-stde"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_PRLINE2").text = "B04"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_CCTYP").text = "N/A"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_CCTYP").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/ctxtVBAK-ZZ_CCTYP").caretPosition = 3
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/mbar/menu[1]/menu[11]").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/chkVBAK-ZZWFLW_IND").setFocus
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/chkVBAK-ZZWFLW_IND").selected = true
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[1]/usr/tblSAPLZSD_WFLWDETAILS").getAbsoluteRow(0).selected = true
session.findById("wnd[1]/usr/tblSAPLZSD_WFLWDETAILS/txtGW_DETAILS-TEXT[0,0]").setFocus
session.findById("wnd[1]/usr/tblSAPLZSD_WFLWDETAILS/txtGW_DETAILS-TEXT[0,0]").caretPosition = 0
session.findById("wnd[1]/usr/btnPROCEED").press
session.findById("wnd[1]/usr/tblSAPLZSD_WFLWDETAILS").getAbsoluteRow(6).selected = true
session.findById("wnd[1]/usr/tblSAPLZSD_WFLWDETAILS/txtGW_DETAILS-TEXT[0,6]").setFocus
session.findById("wnd[1]/usr/tblSAPLZSD_WFLWDETAILS/txtGW_DETAILS-TEXT[0,6]").caretPosition = 0
session.findById("wnd[1]/usr/btnPROCEED").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/cntlCOMMENTS/shellcont/shell").text = "Contract#" + vbCr + ""
session.findById("wnd[1]/usr/cntlCOMMENTS/shellcont/shell").setSelectionIndexes 9,9
session.findById("wnd[1]/usr/txtG_HOURS").text = "2"
session.findById("wnd[1]/usr/txtG_HOURS").setFocus
session.findById("wnd[1]/usr/txtG_HOURS").caretPosition = 1
session.findById("wnd[1]/usr/btnPROCEED").press
session.findById("wnd[1]/usr/btnBUTTON_2").press
