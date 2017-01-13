'***********************************************************
' Originally created by Adam Damke
' Modified by Derek M Ashworth 4/3/2014
' added 500E Plant selective MRP Controller, Procurement Type, Price Control, and Costing

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
Dim part,row,plant,currentTab,success
Dim procType,profitCenter,x,priceControl
Dim ex,wb,ws,MRPController,PriceControl2

session.findById("wnd[0]").maximize


file = ChooseFile(defaultLocalDir)
							    MsgBox file
							    
							    Function ChooseFile (ByVal initialDir)
							        Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
							    
							        Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
							        Dim winVersion
							    
							        ' This collection should contain just the one item
							        For Each objItem in colItems
							            'Caption e.g. Microsoft Windows 7 Professional
							            'Name e.g. Microsoft Windows 7 Professional |C:\windows|...
							            'OSType e.g. 18 / OSArchitecture e.g 64-bit
							            'Version e.g 6.1.7601 / BuildNumber e.g 7601
							            winVersion = CInt(Left(objItem.version, 1))
							        Next
							        Set objWMIService = Nothing
							        Set colItems = Nothing
							    
							        If (winVersion <= 5) Then
							            ' Then we are running XP and can use the original mechanism
							            Set cd = CreateObject("UserAccounts.CommonDialog")
							            cd.InitialDir = initialDir
							            cd.Filter = "VBScript Data Files |*.xls;*.xlsx;*.xlsm|All Files|*.*"
							            ' filter index 4 would show all files by default
							            ' filter index 1 would show zip files by default
							            cd.FilterIndex = 1
							            If cd.ShowOpen = True Then
							                ChooseFile = cd.FileName
							            Else
							                ChooseFile = ""
							            End If
							            Set cd = Nothing    
							    
							        Else
							            ' We are running Windows 7 or later
							            Set shell = CreateObject( "WScript.Shell" )
							            Set ex = shell.Exec( "mshta.exe ""about: <input type=file id=X><script>X.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(X.value);close();resizeTo(0,0);</script>""" )
							            ChooseFile = Replace( ex.StdOut.ReadAll, vbCRLF, "" )
							    
							            Set ex = Nothing
							            Set shell = Nothing
							        End If
							    End Function
												    
Set ex = CreateObject("Excel.Application")
ex.Visible = True
Set wb = ex.Workbooks.Open (file)
Set ws = wb.Worksheets(1)

Call Main



Sub Main

	Call extendParts
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	wb.Close(True)
	ex.Quit 
	Set ex = Nothing
	Set wb = Nothing
	Set ws = Nothing
	MsgBox("The requested parts have been extended." & chr(13) & chr(13) & "Thank you.")				
	WScript.ConnectObject session,     "off"
    WScript.ConnectObject application, "off"
	WScript.Quit
	
End Sub


Sub extendParts
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	row = InputBox("Which row would you like to start on?","Starting Point")
	While ws.Cells(row,1).Value <> ""
		part = Trim(ws.Cells(row,1).Value)
		plant = Trim(ws.Cells(row,2).Value)
		If ws.Cells(row,3).Value = "" Then
			procType = "X"
		Else
			procType = Trim(ws.Cells(row,3).Value)
				procType=Left(procType,1)
		End If
		'MsgBox(procType)
		session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = part
		session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").text = part
		session.findById("wnd[0]/tbar[0]/btn[0]").press
		session.findById("wnd[0]/tbar[0]/btn[0]").press
		session.findById("wnd[1]/tbar[0]/btn[20]").press
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		Select Case plant
			Case "500C"
			session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "500C"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "0001"
			session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = "5013"
			session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "01"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = "U03"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = "001"
			'****************************************************************
			Case "500D"
			session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "500D"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "0001"
			session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = "5013"
			session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "01"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = "U04"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = "001"
			'****************************************************************
			Case "500E"
			session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "500E"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "0001"
			session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = "5013"
			session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "01"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = "U05"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = "001"
			session.findById("wnd[1]/usr/ctxtRMMG1_REF-WERKS").text = "500J"
			session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGORT").text = "0001"
			session.findById("wnd[1]/usr/ctxtRMMG1_REF-VKORG").text = "5013"
			session.findById("wnd[1]/usr/ctxtRMMG1_REF-VTWEG").text = "01"
			
			
			'****************************************************************
			Case "500F"
			session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "500F"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "0001"
			session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = "5013"
			session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "01"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = "U06"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = "001"
			'****************************************************************
			Case "500G"
			session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "500G"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "0001"
			session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = "5013"
			session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "01"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = "U07"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = "001"
			'****************************************************************
			Case "500H"
			session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "500H"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "0001"
			session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = "5013"
			session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "01"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = "U08"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = "001"
			'****************************************************************
			Case "500I"
			session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "500I"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "0001"
			session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = "5013"
			session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "01"
			session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = ""
			session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = ""
			'****************************************************************
		End Select
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		WScript.Sleep(250)
		
		If session.findById("wnd[0]/titl").text = "Create Material (Initial Screen)" Then
			session.findById("wnd[2]/tbar[0]/btn[0]").press
			session.findById("wnd[1]/tbar[0]/btn[12]").press
			ws.Cells(row,4).Value = "Material Already Fully Extended"
		Else
			success = False
			Do While success = False				
				Call whereAmI
			'	On Error Resume Next
				Select Case currentTab
					Case "tabpSP01"
					'	RED FLAG
						session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
						session.findById("wnd[0]/tbar[0]/btn[0]").press
						ws.Cells(row,4).Value = "NEEDS REVIEW!"
						Exit Do
					'****************************************************************
					
					Case "tabpSP02"
					'	RED FLAG
						session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
						session.findById("wnd[0]/tbar[0]/btn[0]").press
						ws.Cells(row,4).Value = "NEEDS REVIEW!"
						Exit Do
					'****************************************************************
					
					Case "tabpSP03"
					'	RED FLAG
						session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
						session.findById("wnd[0]/tbar[0]/btn[0]").press
						ws.Cells(row,4).Value = "NEEDS REVIEW!"
						Exit Do
					'****************************************************************
					
					Case "tabpSP04"
					'	Tax Classification:
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2184/tblSAPLMGD1TC_STEUERN/ctxtMG03STEUER-TAXKM[4,0]").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2184/tblSAPLMGD1TC_STEUERN/ctxtMG03STEUER-TAXKM[4,0]").text = "1"
						End If
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP05"
					'	Acct Assign Group: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KTGRM").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KTGRM").text = "10"
						End If				
					
					'	Material Pricing Group: 
						'If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KONDM").text = "" Then
						'	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26").select
						'	WScript.Sleep(100)
						'	If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/ctxtMBEW-HRKFT").text = "" Then
						'		ws.Cells(row,4).Value = ws.Cells(row,4).Value & " Pricing group & origin group!"
						'	Else
						'		x = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/ctxtMBEW-HRKFT").text
						'		session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05").select
						'		WScript.Sleep(100)
						'		session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KONDM").text = x
						'	End If
						'	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05").select
						'	WScript.Sleep(100)
						'End If
					
					'	Item Category Group: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-MTPOS").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-MTPOS").text = "BANS"
							ws.Cells(row,5).Value = ws.Cells(row,4).Value & " Item category changed from null to BANS!"
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04").select
							WScript.Sleep(100)
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").text = "500B"
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05").select
							WScript.Sleep(100)							
						End If
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP06"
					'	Availability Check: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2161/ctxtMARC-MTVFP").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2161/ctxtMARC-MTVFP").text = "03"
						End If
					
					'	Trans Group:
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2162/ctxtMARA-TRAGR").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2162/ctxtMARA-TRAGR").text = "Z001"
						End If
					
					'	Loading Group:
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2162/ctxtMARC-LADGR").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2162/ctxtMARC-LADGR").text = "0002"
						End If
					
					'	Profit Center:
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").text = "5000000024" Or session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").text = "" Then
						profitCenter = plant
						Select Case profitCenter
							Case "500C"
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").text = "5000000015"
							'***********************************************************
							Case "500D"
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").text = "5000000019"
							'***********************************************************
							Case "500E"
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").text = "5000000017"
							'***********************************************************
							Case "500F"
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").text = "5000000018"
							'***********************************************************
							Case "500G"
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").text = "5000000016"
							'***********************************************************
							Case "500H"
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").text = "5000000014"
							'***********************************************************
							Case "500I"
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").text = "5000000025"
							'***********************************************************
						End Select
						End If
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP07"
					'	Green Ball Checkmark
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP08"
					'	Green Ball Checkmark
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP09"
					'	Purchasing Group: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").text = "ELA"
						End If
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP10"
					'	Green Ball Checkmark
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP11"
					'	Green Ball Checkmark
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP12"
					'	Purchasing Group: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2481/ctxtMARC-EKGRP").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2481/ctxtMARC-EKGRP").text = "ELA"
						End If
					
					'	MRP Type: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").text = "PD"
						End If
					'	MRP Controller: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").text = "" Then
							MRPController=ws.Cells(row,8).Value
								MRPController=Left(MRPController,3)
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").text = MRPController
						End If
					
					'	Lot Size: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").text = "EX"
						End If
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP13"
					'	Procurement Type: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").text = "" Then
						session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").text = procType
						End If
						
					'	Prod. Stor. Location: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGPRO").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGPRO").text = "0001"
						End If
					
					'	Storage loc. for EP: 
						If plant <> "500E" And session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGFSB").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGFSB").text = "0001"
						End If
					
					'	In-house Production: 
'						If procType = "E" Or procType = "X" Then
'							If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-DZEIT").text = "" Then
'								session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-DZEIT").text = "10"
'							End If
'						End If
					
					'	GR Processing Time: 
'						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-WEBAZ").text = "" Then
'							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-WEBAZ").text = "1"
'						End If
					
					'	SchedMargin Key: 
						If procType = "F" And session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/ctxtMARC-FHORI").text = "" Then 
						session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/ctxtMARC-FHORI").text = "000"
						ElseIf procType <> "F" And session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/ctxtMARC-FHORI").text = "" Then
						session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/ctxtMARC-FHORI").text = "001"
						End If
						'	Planned Deliv Time: 
						If procType = "F" And session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-PLIFZ").text = "" Then
							session.findById("wnd[0]/tbar[0]/btn[0]").press
							WScript.Sleep(100)
							session.findById("wnd[0]/tbar[0]/btn[0]").press
						Else
							session.findById("wnd[0]/tbar[0]/btn[0]").press
						End If
					'****************************************************************
					
					Case "tabpSP14"
					'	Strategy Group: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-STRGR").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-STRGR").text = "Z1"
						End If
						
					'	Consumption Mode: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-VRMOD").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-VRMOD").text = "1"
						End If
						
					'	Bwd. Consumption Per: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/txtMARC-VINT1").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/txtMARC-VINT1").text = "30"
						End If
						
					'	Mixed MRP: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-MISKZ").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-MISKZ").text = "1"
						End If
						
					'	*Tot. Repl. 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").text = "22"
						End If
						session.findById("wnd[0]/tbar[0]/btn[0]").press
						WScript.Sleep(100)
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP15"
					'	Green Ball Checkmark
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP16"
					'	*Forecast Model: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2524/ctxtMPOP-PRMOD").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2524/ctxtMPOP-PRMOD").text = "N"
						End If
						
					'	*Forecast Periods: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2525/txtMPOP-ANZPR").text = "" Then
						session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2525/txtMPOP-ANZPR").text = "12"
						End If
						session.findById("wnd[0]/tbar[0]/btn[0]").press
						If session.findbyid("wnd[0]/sbar").text <>"" Then
							session.findById("wnd[0]/tbar[0]/btn[0]").press	
						End if					
					'****************************************************************
					
					Case "tabpSP17"
					'	*Prodn Supervisor: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP17/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2601/ctxtMARC-FEVOR").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP17/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2601/ctxtMARC-FEVOR").text = "001"
						End If
						
					'	*Prod Sched Profile: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP17/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2601/ctxtMARC-SFCPF").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP17/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2601/ctxtMARC-SFCPF").text = "Z00002"
						End If
						ws.Cells(row,7).Value = "prod sch fixed"
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP19"
					'	*CC Phys Inv Ind: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/ctxtMARC-ABCIN").text = "" Then
						session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/ctxtMARC-ABCIN").text = "D"
						
					'	*CC Fixed: 
						session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/chkMARC-CCFIX").selected = True					
						End If
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					
					'****************************************************************
					
					Case "tabpSP20"
					'	Green Ball Checkmark
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP21"
					'	Green Ball Checkmark
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP22"
					'	Green Ball Checkmark
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP23"
					'	Quality Management info
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/chkMARA-QMPUR").selected = False Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/chkMARA-QMPUR").selected = True
						End If
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-SSQSS").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-SSQSS").text = "PMX0003"
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-QZGTP").text = "USQP"
						End If
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP24"
					'	Valuation Class: 
						If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-BKLAS").text = "" Then
							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-BKLAS").text = "4200"
						End If
					
					'	Price Control:
'						PriceControl2=ws.Cells(row,9).Value
'							PriceControl2=Left(PriceControl2,1)
'						If PriceControl2 ="V" Then
'							session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-VPRSV").text = "V"
					'		Moving Price:		
'							If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-VERPR").text = "" And session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-STPRS").text = "" Then
'								session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-VERPR").text = ws.Cells(row,10).Value
'							ElseIf session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-VERPR").text = "" And session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-STPRS").text <> "" Then
'								session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-VERPR").text =ws.Cells(row,10).Value 'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-STPRS").text
'							End If
'						Else
'						session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-VPRSV").text = "S"
					'		Standard Price:
'							If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-STPRS").text = "" And session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-VERPR").text = "" Then
'								session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-STPRS").text = ws.Cells(row,10).Value
'							ElseIf session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-STPRS").text = "" And session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-VERPR").text <> "" Then
'								session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-STPRS").text = ws.Cells(row,10).Value 'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-VERPR").text
'							End If						
'						End If
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP25"
					'	Green Ball Checkmark
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP26"
					'	Origin Group: 
					'	If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/ctxtMBEW-HRKFT").text = "" Then
					'		session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05").select
					'		WScript.Sleep(100)
					'		If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KONDM").text = "" Then
					'			If InStr(ws.Cells(row,4).Value,"Pricing group") = 0 Then
					'			ws.Cells(row,4).Value = ws.Cells(row,4).Value & " Pricing group & origin group!"
					'			End If
					'		Else
					'			x = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KONDM").text
					'			session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26").select
					'			WScript.Sleep(100)
					'			session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/ctxtMBEW-HRKFT").text = x
					'		End If
					'		session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26").select
					'		WScript.Sleep(100)
					'	End If
					
					'	Material Origin: 
						session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/chkMBEW-HKMAT").selected = True
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP27"
					'	Green Ball Checkmark
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP28"
					'	Green Ball Checkmark
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
					
					Case "tabpSP29"
					'	Green Ball Checkmark
						session.findById("wnd[0]/tbar[0]/btn[0]").press
					'****************************************************************
				End Select
			'	If Err.Number <> 0 Then
			'		Err.Clear
			'	End If
			'	On Error Goto 0
				WScript.Sleep(250)
				Call whereAmI
				If currentTab = "-OPTION1" Then
					session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
					success = True
					If ws.Cells(row,4).Value = "" Then
						ws.Cells(row,4).Value = "Extension Completed"
					End If
					WScript.Sleep(250)
				End If
				
			Loop			
		End If
		row = row + 1
	Wend

End Sub
'********************************************

Sub whereAmI

currentTab = session.ActiveWindow.GuiFocus.ID
currentTab = Left(currentTab, 50)
currentTab = Right(currentTab, 8)

End Sub