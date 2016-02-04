If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
	On Error Resume next
   Set connection = application.Children(0)
   If Err.Number<>0 Then
   	MsgBox("You are not connected to PMx, please connect and try again")
   	On Error Goto 0
   	WScript.Quit
   End If
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet,vrc,vrc2,SOQuantity,v
Dim messtxt,z,Row,SDRow,Itemno,SDNum,SAPRow,check1
Dim workingRow,i,strRejectionReason,SDCount,PartnerName
Dim j,soldto,ordernum
Dim CSRName,CSREmail,CustEmail
Dim ExcelApp2,ExcelWorkbook2,ExcelSheet2,row2
Dim boolLoopAgain, poNumber
Dim bodyText,recipientNames,filelocation,SubjectText

row2=InputBox("Starting Row?","Order Acknowledgement")
Set wshShell = WScript.CreateObject( "WScript.Shell" )
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )

Do
Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible=true
Set ExcelWorkbook = ExcelApp.Workbooks.Open ("D:\Documents and Settings\dma02\Desktop\OrderAck\OrderAcknowledgement.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
Set ExcelSheet2 = ExcelWorkbook.Worksheets("Sheet3")
SDCount=0
SDRow=0
Row=11

Call starthere
i=i+1
Call info1
row2=row2+1
Call Main
Call SendEMail (recipientNames,SubjectText,bodyText,filelocation,autoSend)
Loop While ExcelSheet2.Cells(row2,1).Value<>""

Sub starthere
Session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nVA03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = ExcelSheet2.Cells(row2,1).Value
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").select
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/btn").press
session.findById("wnd[1]/usr/cmbAKT_VERSION").key = "Basic setting"
session.findById("wnd[1]/usr/cmbAKT_VERSION").setFocus
session.findById("wnd[1]/tbar[0]/btn[11]").press
SDRow=0
SAPRow=0
If ExcelSheet.Cells(14,5).Value ="" Then	
	ExcelSheet.Cells(14,5).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBKD-KURSK[85,0]").text
End If
soldto=Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").text
ordernum=Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-VBELN").text
ExcelSheet.Shapes(1).TextFrame.Characters.Text = Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/txtKUAGV-TXTPA").text
ExcelSheet.Cells(2,2).Value = Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-VBELN").text
ExcelSheet.Cells(8,2).Value = Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text
ExcelSheet.Cells(7,2).Value = Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK").text
poNumber=ExcelSheet.Cells(8,2).Value
ExcelSheet.Cells(40,1).Value = poNumber
Session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
i=1

Call info1
End sub

Sub info1
Do
test4 = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\"&"0"&i).text
boolLoopAgain=false
If test4="Texts" Then
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\"&"0"&i).select
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\0"&i&"/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").topNode = "0025"
	session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\0"&i&"/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectitem "ZCTC","Column1"
	Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\0"&i&"/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "ZCTC","Column1"
	Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\0"&i&"/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleclickitem "ZCTC","Column1"
	ExcelSheet.Shapes(2).TextFrame.Characters.Text=Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\0"&i&"/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text
    boolLoopAgain=true
End If
If test4="Partners" Then
	Call PartnerSelect
	boolLoopAgain=false
End If
i=i+1
Loop While boolLoopAgain = false



ExcelSheet.SaveAs "D:\Documents and Settings\"&strUserName&"\Desktop\OrderAck\XLSX\"&ordernum&"-OrderAck.xlsx",51
		'MsgBox("The end has come")
		ExcelSheet.Printout
		Set WshShell = WScript.CreateObject("WScript.Shell")
		WScript.Sleep 500
		WshShell.SendKeys "D:\Documents and Settings\"&strUserName&"\Desktop\OrderAck\PDF\"&ordernum&"-OrderAck.pdf"
		WshShell.SendKeys "{ENTER}"

		WScript.Sleep(500)
		ExcelSheet2.Cells(row2,2).Value="yes"
		'ExcelApp.Workbooks(ordernum&"-OrderAck.xlsx").Close(True)
'		ExcelApp.Quit
'		Set ExcelApp=Nothing
'		Set ExcelWorkbook=Nothing
'		Set ExcelSheet=Nothing
'		Set ExcelSheet2=Nothing
Call Main
End Sub

Sub PartnerSelect
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\"&"0"&i).select
j=0
CSRName=ExcelSheet.Cells(13,2).Value


Do While j<25
	PartnerName=session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,"&(j)&"]").value
 	'MsgBox(PartnerName)
		If PartnerName = "Customer ServiceRep" Then
			Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,"&(j)&"]").setfocus
			Session.findById("wnd[0]").sendVKey 2
			CSRName=Session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0313/txtADDR1_DATA-NAME1").text
			
			'MsgBox(CSRName)
			ExcelSheet.Cells(13,2).Value=Session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0313/txtADDR1_DATA-NAME1").text
			CSREmail=ExcelSheet.Cells(39,1).Value
				If CSRName="" Then
					ExcelSheet.Cells(13,2).Value="Not Listed"
				End If
			
	    	Session.findbyid("wnd[1]/tbar[0]/btn[12]").press
	    		'If CSREmail="" Then
	    		'	ExcelSheet.Cells(14,2).Value="Not Listed"
	    			
	    		'End If
		ElseIf PartnerName="Sold-to party" Then
			Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,"&(j)&"]").setfocus
			Session.findById("wnd[0]").sendVKey 2
			ExcelSheet.Cells(9,2).Value=session.findById("wnd[1]/usr/subGCS_ADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0313/txtSZA1_D0100-SMTP_ADDR").text
	    	CustEmail=ExcelSheet.Cells(9,2).Value
	    	Session.findbyid("wnd[1]/tbar[0]/btn[12]").press
	    	If CustEmail=""Then
	    		ExcelSheet.Cells(9,2).Value="Not Listed"
	    		CustEmail=";"
	    	End If
	    
	    End If
	
	j=j+1
Loop
End Sub

Sub Main 

	filelocation = "D:\Documents and Settings\"&strUserName&"\Desktop\OrderAck\PDF\"&ordernum&"-OrderAck.pdf"
	Const showWindow = False
	Const autoSend=False
	SubjectText="***ID "&soldto&"***Order "&ordernum&"*** Acknowledgement **** PO "&poNumber
	
	bodyText="Thank you for your order. Please see the attachment for your Purchase Order Number and contact information." &vbCrLf
	ExcelSheet.cells(14,1).calculate
	CSREmail=ExcelSheet.Cells(39,1).Value
	recipientNames = Array(CustEmail,CSREmail,"windsorparts@power.alstom.com")
	WScript.Sleep(1000)
	
End Sub
	
Sub SendEMail (recipientNames,SubjectText,bodyText,attachmentPath,autoSend)
		Const olMailItem = 0
		Dim objOutlook, objOutlookMsg, objOutlookRecip, objOutlookAttach
    	' Create the Outlook session.
    Set objOutlook = CreateObject("Outlook.Application")

    ' Create the message.
    Set objOutlookMsg = objOutlook.CreateItem(olMailItem)
	objOutlookMsg.Display
    With objOutlookMsg
        ' Add the To recipient(s) to the message.
        Dim name
        For Each name In recipientNames
            Set objOutlookRecip = .Recipients.Add(name)
        Next
        
	   ' Set the Subject, Body, and Importance of the message.
        .Subject = subjectText
        .Body = bodyText
       '.Importance = olImportanceHigh  'High importance

       ' Add attachments to the message.
        Set objOutlookAttach = .Attachments.Add(attachmentPath)


       ' Resolve each Recipient's name.
       For Each objOutlookRecip In .Recipients
           objOutlookRecip.Resolve
       Next

       ' Should we display the message before sending?
       If Not autoSend Then
           .Display
       Else
       	MsgBox("nana nana boo boo")
'           .Save
'           .Send
       End If
    End With

End Sub
    	
	
WScript.Sleep(10000)
		Set objOutlookMsg = Nothing
		Set objOutlook = Nothing

session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
session.findById("wnd[0]").sendVKey 0	
ExcelApp.Workbooks(ordernum&"-OrderAck.xlsx").Close(True)
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWorkbook=Nothing
		Set ExcelSheet=Nothing
		Set ExcelSheet2=Nothing
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit