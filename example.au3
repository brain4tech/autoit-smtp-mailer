#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Brain4Tech, Jos (https://www.autoitscript.com/forum/profile/19-jos/)

 Script Function:
	Examplescript for Jos's smtp mailer.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

; Enable less secure apps in case you are using gmail: https://myaccount.google.com/lesssecureapps (source: https://www.autoitscript.com/forum/topic/186822-autoit-gmail-send-mail/?do=findComment&comment=1341666)

; include udf
#include "smtp_mailer_UDF.au3"

; define function parameters
$SmtpServer = ""
$FromName = ""
$FromAddress = ""
$ToAddress = ""
$Subject = ""
$Body = ""
$AttachFiles = ""
$CcAddress = ""
$BccAddress = ""
$Importance = ""
$Username = ""
$Password = ""
$IPPort = 25
$ssl = 0
$tls = 0

; execute functions with error-handling
$rc = _INetSmtpMailCom($SmtpServer, $FromName, $FromAddress, $ToAddress, $Subject, $Body, $AttachFiles, $CcAddress, $BccAddress, $Importance, $Username, $Password, $IPPort, $ssl, $tls)
If @error Then
	MsgBox(0, "Error sending message", "Error code:" & @error & "  Description:" & $rc, 3)
EndIf
