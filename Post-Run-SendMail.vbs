'==== This script use to send email using CDO object
'==== This is a post-run script
'==== Use this script file with any process file

Dim objCDO
Dim objCDOConfig

Dim SendToAddr			'=== Email address which you want to send
Dim MailSubject			'=== Email subject
Dim MailBody			'=== Email Body/Detail

Dim wshShell
Dim wshUsrEnv

Set wshShell = CreateObject("WScript.Shell")		'=== Get Shell object from WScript object
Set wshUsrEnv = wshShell.Environment("USER")		'=== Get current user environment variables

'=== Set error level to 1 and if no error during program execution then set error level to 0 at the end of program
wshUsrEnv("PRERRORLEVEL") = 1

Set objCDO = CreateObject("CDO.Message")
Set objCDOConfig = CreateObject("CDO.Configuration") 


Set Flds = objCDOConfig.Fields

Flds.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = false
Flds.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1       	'========= Use this to set the authenication method. Values can be one of 0(no authenication),1(Use the basic (clear text) authentication mechanism.),2 (Use the NTLM authentication mechanism.)
Flds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp server" 	'========= Change to your SMTP server
Flds.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "username"  	'========= Change to user name 
Flds.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "password"  	'========= Change to Password for the user name
Flds.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 	 	'========== Use this to set the sending method. Values can be one of 1(Send the message using the local SMTP service pickup directory), 2( Send the message using the network ( SMTP over the network) ) , 3 ( Send the message using the Microsoft Exchange mail submission Uniform Resource Identifier )
Flds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25    	'========= Change this if PORT address from which the mail is sent is not equal to 25
Flds.Update


'====== Email Header =======

SendToAddr = "abc@xyz.com"				'=== Change Email Address , Its compulsory to give
MailSubject = "Process Runner File Run Report"

'====== Email Body =======

MailBody = "Hello there," & vbCRLF
MailBody = MailBody & vbCRLF & "Process Runner finished executing following Process:" & vbCRLF 
MailBody = MailBody & vbCRLF & "Process Name: " & #CURPROC# & vbCRLF	'=== Process Name

If #CURPRTYPE# = "TX" Then				'=== Transaction File Type
	MailBody = MailBody & vbCRLF &  "File Type: Transaction"
ElseIf #CURPRTYPE# = "BA" Then				'=== BAPI/RFM File Type
	MailBody = MailBody & vbCRLF &  "File Type: BAPI/RFM"
ElseIf #CURPRTYPE# = "GS" Then				'=== GUI Script File Type
	MailBody = MailBody & vbCRLF &  "File Type: GUI Script"
ElseIf #CURPRTYPE# = "DE" Then				'=== Data Extractor File Type
	MailBody = MailBody & vbCRLF &  "File Type: Data Extractor"
End If

MailBody = MailBody & vbCRLF

MailBody = MailBody & vbCRLF & "Excel File Name: " & #CURXLFILE# & vbCRLF	'=== Excel File Name
MailBody = MailBody & vbCRLF & "Excel Sheet Name: " & #CURXLSHEET# & vbCRLF	'=== Excel Sheet Name
MailBody = MailBody & vbCRLF & "SAP System: " & #CURSAPSYS# & vbCRLF 	'=== SAP System
MailBody = MailBody & vbCRLF & "SAP User: " & #SAPUSR# & vbCRLF		'=== SAP User

If #ERRORSTAT# = 0 And #ERRORCOUNT# = 0 Then		'=== Check if any error found
	MailBody = MailBody & vbCRLF & "Error Status: No Error Found" & vbCRLF
	MailBody = MailBody & vbCRLF & "No. Of Error: " & #ERRORCOUNT# & vbCRLF	 
Else
	MailBody = MailBody & vbCRLF & "Error Status: Error Found" & vbCRLF
	If #CURPRTYPE# = "TX" Then
		MailBody = MailBody & vbCRLF & "No. Of Error: " & #ERRORCOUNT# & vbCRLF		'=== Send no. of error only for TX
	End If
End If

If #CURPRTYPE# = "DE" Then				'=== Check Data Extractor
	MailBody = MailBody & vbCRLF &  "No. of Record Extracted: " & #NOOFREC# & vbCRLF
Else
	MailBody = MailBody & vbCRLF &  "No. of Calls: " & #NOOFCALL# & vbCRLF
End If


MailBody = MailBody & vbCRLF &  "Date/Time of Run: " & Now() & vbCRLF	'=== Process File
MailBody = MailBody & vbCRLF &  "Process File: " & #CURPRFILE# & vbCRLF	'=== Process File

MailBody = MailBody & vbCRLF & vbCRLF & "This e-mail was auto generated from Process Runner" & vbCRLF

'==============================

Set objCDO.Configuration = objCDOConfig

objCDO.From = "abc@xyz.com"     '========= Change to sender's email id. Its a mendotary entry
objCDO.To = SendToAddr         '========= Email address which you want to send
objCDO.CC = ""		       '==== CC	

objCDO.Subject = MailSubject   '=== Email subject
objCDO.TextBody = MailBody     '=== Email Body/Detail
objCDO.Send		       '=== Send Mail	


wshUsrEnv("PRERRORLEVEL") = 0  '=== Set error level to 0 / Successfully executed program


Set wshUsrEnv = Nothing
Set wshShell = Nothing

Set objCDO = Nothing
Set objCDOConfig = Nothing
