#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Jos (https://www.autoitscript.com/forum/profile/19-jos/)

 Script Function:
	Smtp mailer that supports Html and attachments (from: https://www.autoitscript.com/forum/topic/23860-smtp-mailer-that-supports-html-and-attachments/)

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

#include-once

; ### INCLUDES ###
#include <File.au3>

; ### GLOBALS ###
Global $oMyRet[2]
Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")

; ### MAIN FUNCTION ###
Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Importance = "Normal", $s_Username = "", $s_Password = "", $IPPort = 25, $ssl = 0, $tls = 0)

	#cs ### PARAMETERS EXPLANATOIN ###
	$SmtpServer = "MailServer"              ; address for the smtp-server to use - REQUIRED
	$FromName = "Name"                      ; name from who the email was sent
	$FromAddress = "your@Email.Address.com" ; address from where the mail should come
	$ToAddress = "your@Email.Address.com"   ; destination address of the email - REQUIRED
	$Subject = "Userinfo"                   ; subject from the email - can be anything you want it to be
	$Body = ""                              ; the messagebody from the mail - can be left blank but then you get a blank mail
	$AttachFiles = ""                       ; the file(s) you want to attach seperated with a ; (Semicolon) - leave blank if not needed
	$CcAddress = "CCadress1@test.com"       ; address for cc - leave blank if not needed
	$BccAddress = "BCCadress1@test.com"     ; address for bcc - leave blank if not needed
	$Importance = "Normal"                  ; Send message priority: "High", "Normal", "Low"
	$Username = "******"                    ; username for the account used from where the mail gets sent - REQUIRED
	$Password = "********"                  ; password for the account used from where the mail gets sent - REQUIRED
	$IPPort = 25                            ; port used for sending the mail
	$ssl = 0                                ; enables/disables secure socket layer sending - put to 1 if using httpS
	$tls = 0                                ; enables/disables TLS when required
	#ce

	Local $objEmail = ObjCreate("CDO.Message")

	$objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
	$objEmail.To = $s_ToAddress

	Local $i_Error = 0
	Local $i_Error_desciption = ""

	If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
	If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress

	$objEmail.Subject = $s_Subject

	If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
		$objEmail.HTMLBody = $as_Body
	Else
		$objEmail.Textbody = $as_Body & @CRLF
	EndIf

	If $s_AttachFiles <> "" Then

		Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")

		For $x = 1 To $S_Files2Attach[0]
			$S_Files2Attach[$x] = _PathFull($S_Files2Attach[$x])
			; ConsoleWrite('@@ Debug : $S_Files2Attach[$x] = ' & $S_Files2Attach[$x] & @LF & '>Error code: ' & @error & @LF) ;### Debug Console

			If FileExists($S_Files2Attach[$x]) Then
				ConsoleWrite('+> File attachment added: ' & $S_Files2Attach[$x] & @LF)
				$objEmail.AddAttachment($S_Files2Attach[$x])
			Else
				ConsoleWrite('!> File not found to attach: ' & $S_Files2Attach[$x] & @LF)
				SetError(1)
				Return 0
			EndIf

		Next
	EndIf

	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer

	If Number($IPPort) = 0 Then $IPPort = 25

	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort

	;Authenticated SMTP
	If $s_Username <> "" Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
	EndIf

	; Set security params
	If $ssl Then $objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
	If $tls Then $objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendtls") = True

	;Update settings
	$objEmail.Configuration.Fields.Update

	; Set Email Importance
	Switch $s_Importance
		Case "High"
			$objEmail.Fields.Item("urn:schemas:mailheader:Importance") = "High"
		Case "Normal"
			$objEmail.Fields.Item("urn:schemas:mailheader:Importance") = "Normal"
		Case "Low"
			$objEmail.Fields.Item("urn:schemas:mailheader:Importance") = "Low"
	EndSwitch

	$objEmail.Fields.Update

	; Sent the Message
	$objEmail.Send
	If @error Then
		SetError(2)
		Return $oMyRet[1]
	EndIf

	$objEmail = ""

EndFunc   ;==>_INetSmtpMailCom

; ### COM ERROR HANDLER ###
Func MyErrFunc()
	$HexNumber = Hex($oMyError.number, 8)
	$oMyRet[0] = $HexNumber
	$oMyRet[1] = StringStripWS($oMyError.description, 3)
	ConsoleWrite("### COM Error !  Number: " & $HexNumber & "   ScriptLine: " & $oMyError.scriptline & "   Description:" & $oMyRet[1] & @LF)
	SetError(1) ; something to check for when this function returns
	Return
EndFunc   ;==>MyErrFunc
