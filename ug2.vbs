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
session.findById("wnd[0]/tbar[0]/okcd").text = "xd02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = "10071427"
session.findById("wnd[1]/usr/ctxtRF02D-BUKRS").text = "5000"
session.findById("wnd[1]/usr/ctxtRF02D-VKORG").text = "5013"
session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = "01"
session.findById("wnd[1]/usr/ctxtRF02D-SPART").text = "00"
session.findById("wnd[1]/usr/ctxtRF02D-SPART").setFocus
session.findById("wnd[1]/usr/ctxtRF02D-SPART").caretPosition = 2
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/mbar/menu[3]/menu[6]").select
session.findById("wnd[0]/usr/subSUBTAB:SAPMF02D:3502/tblSAPMF02DTCTRL_TEXTE/txtRTEXT-LTEXT[3,10]").setFocus
session.findById("wnd[0]/usr/subSUBTAB:SAPMF02D:3502/tblSAPMF02DTCTRL_TEXTE/txtRTEXT-LTEXT[3,10]").caretPosition = 2
session.findById("wnd[0]").sendVKey 2
On Error Resume next
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,1]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,2]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,3]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,4]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,5]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,6]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,7]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,8]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,9]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,10]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,11]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,12]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,13]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,14]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,15]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,16]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,17]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,18]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,19]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,20]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,21]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,1]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,2]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,3]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,4]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,5]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,6]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,7]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,8]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,9]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,10]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,11]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,12]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,13]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,14]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,15]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,16]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,17]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,18]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,19]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,20]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,21]").text = ""
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,21]").setFocus
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,21]").caretPosition = 0
On Error Goto 0
'session.findById("wnd[0]").sendVKey 0
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,1]").text = "TERMS AND CONDITIONS:     Z_US_ST03938)*********************"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,1]").text = "TERMS AND CONDITIONS:     Z_US_ST03938)*********************"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,1]").caretPosition = 61
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,2]").text = "ALSTOM POWER INC.'S STANDARD TERMS AND CONDITIONS OF SALES"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,3]").text = "(GOODS AND SERVICES) DOMESTIC/REV: 08/12/05 (FORM TC33.DOC) APPLY TO"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,4]").text = " THIS TRANSACTION. TO VIEW THE COMPLETE TERMS AND CONDITIONS DOCUMENT"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,5]").text = "GO TO:  www.service.power.alstom.com"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,5]").setFocus
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,5]").caretPosition = 36
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA").columns.elementAt(2).width = 72
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,7]").text = "TERMS AND CONDITIONS ACCEPTANCE:  (Z_US_ST03490)**********"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,7]").caretPosition = 60
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,8]").text = "PURCHASER AGREES THAT THE RECEIPT OF THE GOODS FROM THE SELLER"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,9]").text = "SIGNIFIES PURCHASER'S ACCEPTANCE OF SELLER'S STANDARD TERMS"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,10]").text = "AND CONDITIONS OF SALE (OR OTHER MUTUALLY AGREED UPON SET OF"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,11]").text = "TERMS AND CONDITIONS WHICH SELLER HAS EXPRESSLY AGREED UPON OR"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,12]").text = "WRITING AS APPLICABLE TO THIS  FACILITY, FREIGHT COLLECT,"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,13]").text = "WHERE THE GOODS WERE ORIGINALLY SHIPPED FROM FOR CREDIT AND/OR TERMS"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,14]").text = "AND CONDITIONS DISCUSSION AFTER RETURNING THE GOODS, CONTACT THE"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,15]").text = "BOILER SERVICE OPERATIONS DIVISION OF ALSTOM POWER INC."
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,15]").setFocus
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,15]").caretPosition = 57
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,14]").text = "PARTS ADMINISTRATION"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,14]").caretPosition = 22
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,14]").text = "200 GREAT POND DRIVE"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,14]").caretPosition = 22
'session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,15]").text = "WINDSOR, CT 06095"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,15]").caretPosition = 17
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,15]").setFocus
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/ctxtRSTXT-TXPARGRAPH[0,15]").caretPosition = 1
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/lbl[1,5]").setFocus
session.findById("wnd[1]/usr/lbl[1,5]").caretPosition = 0
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,15]").text = "(REV 8/23/10)"
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,15]").setFocus
session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,15]").caretPosition = 13
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA").verticalScrollbar.position = 3
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA").verticalScrollbar.position = 0
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,10]").text = "AND CONDITIONS OF SALE (OR OTHER MUTUALLY AGREED UPON SET OF"
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,11]").text = "TERMS AND CONDITIONS WHICH SELLER HAS EXPRESSLY AGREED UPON OR"
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,13]").text = "WHERE THE GOODS WERE ORIGINALLY SHIPPED FROM FOR CREDIT AND/OR TERMS"
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,14]").text = "AND CONDITIONS DISCUSSION AFTER RETURNING THE GOODS, CONTACT THE"
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,15]").text = "BOILER SERVICE OPERATIONS DIVISION OF ALSTOM POWER INC."
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,15]").setFocus
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,15]").caretPosition = 6
'session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA").verticalScrollbar.position = 3
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
