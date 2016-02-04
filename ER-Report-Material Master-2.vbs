Dim strUserName
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
Set wshShell = WScript.CreateObject( "WScript.Shell" )
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nsqvi"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "1-FIXEDVENDOR"
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 13
session.findById("wnd[0]/usr/btnP1").press
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/ctxtSP$00005-LOW").text = ""
session.findById("wnd[0]/usr/ctxtSP$00005-LOW").setFocus
session.findById("wnd[0]/usr/ctxtSP$00005-LOW").caretPosition = 0
session.findById("wnd[0]").sendVKey 0
'session.findBy'Id("wnd[0]/usr/rad%ALV").setFocus
'session.findBy'Id("wnd[0]/usr/rad%ALV").select
session.findById("wnd[0]/usr/ctxtSP$00001-LOW").text = ""
session.findById("wnd[0]/usr/ctxtSP$00001-LOW").caretPosition = 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findbyId("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\"&strUserName&"\Desktop\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "1-FIXEDVENDOR.txt"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/NSQVI"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "2-PIR-REPORT"
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 12
session.findById("wnd[0]/usr/btnP1").press
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
'session.findById("wnd[0]/usr/rad%ALV").setFocus
'session.findById("wnd[0]/usr/rad%ALV").select
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findbyId("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\"&strUserName&"\Desktop\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "2-PIR-REPORT.txt"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/NSQVI"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "3-MATERIALBASE"
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 14
session.findById("wnd[0]/usr/btnP1").press
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/tbar[0]/btn[2]").press
'session.findById("wnd[0]/usr/rad%ALV").setFocus
'session.findById("wnd[0]/usr/rad%ALV").select
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findbyId("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\"&strUserName&"\Desktop\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "3-500X-MMBASE.txt"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 2
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/btn%_SP$00002_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "50DD"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "50GD"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "50DE"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "50GE"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "50DF"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").text = "50GF"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findbyId("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\"&strUserName&"\Desktop\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "3-50XX-MMBASE.txt"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
session.findById("wnd[0]").sendVKey 0
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"