<%
Const S_ID_EMAIL = 0
Const S_SERVER_NAME_EMAIL = 1
Const L_SERVER_PORT_EMAIL = 2
Const B_USE_HTML_EMAIL = 3
Const N_BODY_FORMAT_EMAIL = 4
Const N_MAIL_FORMAT_EMAIL = 5
Const S_TO_EMAIL = 6
Const S_CC_EMAIL = 7
Const S_BCC_EMAIL = 8
Const S_FROM_EMAIL = 9
Const S_SUBJECT_EMAIL = 10
Const S_BODY_EMAIL = 11
Const S_ATTACHMENTS_EMAIL = 12
Const O_EMAIL = 13
Const B_COMPONENT_INITIALIZED_EMAIL = 14

Const N_EMAIL_COMPONENT_SIZE = 14

Const B_USE_CDO_2003 = True
Const B_USE_CDO_NTS = False
Const SMTP_SERVER_NAME = "192.168.50.10"
Const SMTP_USER_NAME = "nomina@issste.gob.mx"
Const SMTP_PASSWORD = "jgg11302"
Const SMTP_PORT = 25

Dim aEmailComponent()
ReDim aEmailComponent(N_EMAIL_COMPONENT_SIZE)

Function InitializeEmailComponent(oRequest, aEmailComponent)
'************************************************************
'Purpose: To initialize the empty elements of the e-mail
'         Component using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aEmailComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeEmailComponent"
	Redim Preserve aEmailComponent(N_EMAIL_COMPONENT_SIZE)

	If IsEmpty(aEmailComponent(S_ID_EMAIL)) Then
		If Len(oRequest("MessageID").Item) > 0 Then
			aEmailComponent(S_ID_EMAIL) = oRequest("MessageID").Item
		Else
			aEmailComponent(S_ID_EMAIL) = ""
		End If
	End If

	If IsEmpty(aEmailComponent(S_SERVER_NAME_EMAIL)) Then
		If Len(oRequest("SMTPServerName").Item) > 0 Then
			aEmailComponent(S_SERVER_NAME_EMAIL) = oRequest("SMTPServerName").Item
		Else
			aEmailComponent(S_SERVER_NAME_EMAIL) = SMTP_SERVER_NAME
		End If
	End If

	If IsEmpty(aEmailComponent(L_SERVER_PORT_EMAIL)) Then
		If Len(oRequest("SMTPServerPort").Item) > 0 Then
			aEmailComponent(L_SERVER_PORT_EMAIL) = oRequest("SMTPServerPort").Item
		Else
			aEmailComponent(L_SERVER_PORT_EMAIL) = SMTP_PORT
		End If
	End If

	If IsEmpty(aEmailComponent(B_USE_HTML_EMAIL)) Then
		If Len(oRequest("UseHTML").Item) > 0 Then
			aEmailComponent(B_USE_HTML_EMAIL) = (CInt(oRequest("UseHTML").Item) <> 0)
		Else
			aEmailComponent(B_USE_HTML_EMAIL) = True
		End If
	End If

	If IsEmpty(aEmailComponent(N_BODY_FORMAT_EMAIL)) Then
		If Len(oRequest("BodyFormat").Item) > 0 Then
			aEmailComponent(N_BODY_FORMAT_EMAIL) = CInt(oRequest("BodyFormat").Item)
		Else
			If aEmailComponent(B_USE_HTML_EMAIL) Then
				aEmailComponent(N_BODY_FORMAT_EMAIL) = 0 'CdoBodyFormatHTML
			Else
				aEmailComponent(N_BODY_FORMAT_EMAIL) = 1 'CdoBodyFormatText
			End If
		End If
	End If

	If IsEmpty(aEmailComponent(N_MAIL_FORMAT_EMAIL)) Then
		If Len(oRequest("MailFormat").Item) > 0 Then
			aEmailComponent(N_MAIL_FORMAT_EMAIL) = CInt(oRequest("MailFormat").Item)
		Else
			If aEmailComponent(B_USE_HTML_EMAIL) Then
				aEmailComponent(N_MAIL_FORMAT_EMAIL) = 0 'CdoMailFormatMime
			Else
				aEmailComponent(N_MAIL_FORMAT_EMAIL) = 1 'CdoMailFormatText
			End If
		End If
	End If

	If IsEmpty(aEmailComponent(S_TO_EMAIL)) Then
		If Len(oRequest("To").Item) > 0 Then
			aEmailComponent(S_TO_EMAIL) = oRequest("To").Item
		Else
			aEmailComponent(S_TO_EMAIL) = ""
		End If
	End If
	aEmailComponent(S_TO_EMAIL) = Replace(aEmailComponent(S_TO_EMAIL), ";", ",", 1, -1, vbBinaryCompare)

	If IsEmpty(aEmailComponent(S_CC_EMAIL)) Then
		If Len(oRequest("CC").Item) > 0 Then
			aEmailComponent(S_CC_EMAIL) = oRequest("CC").Item
		Else
			aEmailComponent(S_CC_EMAIL) = ""
		End If
	End If
	aEmailComponent(S_CC_EMAIL) = Replace(aEmailComponent(S_CC_EMAIL), ";", ",", 1, -1, vbBinaryCompare)

	If IsEmpty(aEmailComponent(S_BCC_EMAIL)) Then
		If Len(oRequest("BCC").Item) > 0 Then
			aEmailComponent(S_BCC_EMAIL) = oRequest("BCC").Item
		Else
			aEmailComponent(S_BCC_EMAIL) = ""
		End If
	End If
	aEmailComponent(S_BCC_EMAIL) = Replace(aEmailComponent(S_BCC_EMAIL), ";", ",", 1, -1, vbBinaryCompare)

	If IsEmpty(aEmailComponent(S_FROM_EMAIL)) Then
		If Len(oRequest("From").Item) > 0 Then
			aEmailComponent(S_FROM_EMAIL) = oRequest("From").Item
		Else
			aEmailComponent(S_FROM_EMAIL) = ""
		End If
	End If

	If IsEmpty(aEmailComponent(S_SUBJECT_EMAIL)) Then
		If Len(oRequest("Subject").Item) > 0 Then
			aEmailComponent(S_SUBJECT_EMAIL) = oRequest("Subject").Item
		Else
			aEmailComponent(S_SUBJECT_EMAIL) = ""
		End If
	End If

	If IsEmpty(aEmailComponent(S_BODY_EMAIL)) Then
		If Len(oRequest("Body").Item) > 0 Then
			aEmailComponent(S_BODY_EMAIL) = oRequest("Body").Item
		Else
			aEmailComponent(S_BODY_EMAIL) = ""
		End If
	End If

	If IsEmpty(aEmailComponent(S_ATTACHMENTS_EMAIL)) Then
		If Len(oRequest("Attachments").Item) > 0 Then
			aEmailComponent(S_ATTACHMENTS_EMAIL) = oRequest("Attachments").Item
		Else
			aEmailComponent(S_ATTACHMENTS_EMAIL) = ""
		End If
	End If
	aEmailComponent(S_ATTACHMENTS_EMAIL) = Replace(aEmailComponent(S_ATTACHMENTS_EMAIL), ";", ",", 1, -1, vbBinaryCompare)

	aEmailComponent(B_COMPONENT_INITIALIZED_EMAIL) = True
	InitializeEmailComponent = Err.number
	Err.Clear
End Function

Function SendEmail(oRequest, aEmailComponent, sErrorDescription)
'************************************************************
'Purpose: To send an e-mail
'Inputs:  oRequest, aEmailComponent
'Outputs: aEmailComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SendEmail"
	Dim iIndex
	Dim asAttachments
	Dim sLogEntry
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmailComponent(B_COMPONENT_INITIALIZED_EMAIL)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmailComponent(oRequest, aEmailComponent)
	End If

	If (Not IsObject(aEmailComponent(O_EMAIL))) Or (IsNull(aEmailComponent(O_EMAIL))) Then
		If B_USE_CDO_2003 Then
			Set aEmailComponent(O_EMAIL) = Server.CreateObject("CDO.Message")
			lErrorNumber = Err.number
			If lErrorNumber <> 0 Then
				sErrorDescription = aIntlComponent(AS_DESCRIPTORS_INTL)(1425) 'No se pudo crear una instancia del objeto 'CDO.Message'. El archivo 'CDONTS.dll' no está correctamente registrado en el Servidor Web. Favor de contactar al Administrador.
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR />" & aIntlComponent(AS_DESCRIPTORS_INTL)(1391) & "&nbsp;" & Err.description 'Error del servidor Web:
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 1425, "EmailComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			End If
		ElseIf B_USE_CDO_NTS Then
			Set aEmailComponent(O_EMAIL) = Server.CreateObject("CDONTS.NewMail")
			lErrorNumber = Err.number
			If lErrorNumber <> 0 Then
				sErrorDescription = aIntlComponent(AS_DESCRIPTORS_INTL)(1426) 'No se pudo crear una instancia del objeto 'CDONTS.NewMail'. El archivo 'CDONTS.dll' no se encuentra registrado correctamente en el servidor Web. Favor de contactar al Administrador.
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR />" & aIntlComponent(AS_DESCRIPTORS_INTL)(1391) & "&nbsp;" & Err.description 'Error del servidor Web:
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 1426, "EmailComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			End If
		Else
			Set aEmailComponent(O_EMAIL) = Server.CreateObject("SMTPsvg.Mailer")
			lErrorNumber = Err.number
			If lErrorNumber <> 0 Then
				sErrorDescription = aIntlComponent(AS_DESCRIPTORS_INTL)(1427) 'No se pudo crear una instancia del objeto 'SMTPsvg.Mailer'. El archivo 'SMTPsvg.dll' no se encuentra registrado correctamente en el servidor Web. Favor de contactar al Administrador.
				If Len(Err.description) > 0 Then
					sErrorDescription = sErrorDescription & "<BR />" & aIntlComponent(AS_DESCRIPTORS_INTL)(1391) & "&nbsp;" & Err.description 'Error del servidor Web:
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 1427, "EmailComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			End If
		End If
	End If

	If lErrorNumber = 0 Then
		If Len(aEmailComponent(S_FROM_EMAIL)) = 0 Then
			lErrorNumber = -1
			sErrorDescription = aIntlComponent(AS_DESCRIPTORS_INTL)(1428) 'No se especificó el correo electrónico de la persona que envía el mensaje.
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 1428, "EmailComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		ElseIf (Len(aEmailComponent(S_TO_EMAIL)) = 0) And (Len(aEmailComponent(S_CC_EMAIL)) = 0) And (Len(aEmailComponent(S_BCC_EMAIL)) = 0) Then
'			lErrorNumber = -1
'			sErrorDescription = aIntlComponent(AS_DESCRIPTORS_INTL)(1429) 'No se especificó el correo electrónico de la(s) persona(s) que recibirá(n) el mensaje.
'			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 1429, "EmailComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		Else
			If FileExists(Server.MapPath("Template_Email.htm"), "") Then aEmailComponent(S_BODY_EMAIL) = aEmailComponent(S_BODY_EMAIL) & GetFileContents(Server.MapPath("Template_Email.htm"), "")
			If B_USE_CDO_2003 Then
				If aEmailComponent(B_USE_HTML_EMAIL) Then
					aEmailComponent(O_EMAIL).HTMLBody = aEmailComponent(S_BODY_EMAIL)
				Else
					aEmailComponent(O_EMAIL).TextBody = aEmailComponent(S_BODY_EMAIL)
				End If
				aEmailComponent(O_EMAIL).MimeFormatted = aEmailComponent(B_USE_HTML_EMAIL)
				aEmailComponent(O_EMAIL).To = aEmailComponent(S_TO_EMAIL)
				aEmailComponent(O_EMAIL).CC = aEmailComponent(S_CC_EMAIL)
				aEmailComponent(O_EMAIL).BCC = aEmailComponent(S_BCC_EMAIL)
				aEmailComponent(O_EMAIL).From = aEmailComponent(S_FROM_EMAIL)
				aEmailComponent(O_EMAIL).Subject = aEmailComponent(S_SUBJECT_EMAIL)
				If Len(aEmailComponent(S_ATTACHMENTS_EMAIL)) > 0 Then
					asAttachments = Split(aEmailComponent(S_ATTACHMENTS_EMAIL), ",", -1, vbBinaryCompare)
					For iIndex = 0 To UBound(asAttachments)
						Call aEmailComponent(O_EMAIL).AddAttachment(asAttachments(iIndex), "", "")
					Next
				End If
				aEmailComponent(O_EMAIL).Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing").Value = 2 'cdoSendUsingPort
				aEmailComponent(O_EMAIL).Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver").Value = aEmailComponent(S_SERVER_NAME_EMAIL)
				aEmailComponent(O_EMAIL).Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport").Value = aEmailComponent(L_SERVER_PORT_EMAIL)
				aEmailComponent(O_EMAIL).Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'cdoBasic
				If Len(SMTP_USER_NAME) > 0 Then
					aEmailComponent(O_EMAIL).Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTP_USER_NAME
					aEmailComponent(O_EMAIL).Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTP_PASSWORD
				End If
				aEmailComponent(O_EMAIL).Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
				aEmailComponent(O_EMAIL).Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
				Call aEmailComponent(O_EMAIL).Configuration.Fields.Update()
				Err.number = lErrorNumber
				If lErrorNumber = 0 Then Call aEmailComponent(O_EMAIL).Send()
			ElseIf B_USE_CDO_NTS Then
				aEmailComponent(O_EMAIL).BodyFormat = aEmailComponent(N_BODY_FORMAT_EMAIL)
				aEmailComponent(O_EMAIL).MailFormat = aEmailComponent(N_MAIL_FORMAT_EMAIL)
				aEmailComponent(O_EMAIL).To = aEmailComponent(S_TO_EMAIL)
				aEmailComponent(O_EMAIL).Cc = aEmailComponent(S_CC_EMAIL)
				aEmailComponent(O_EMAIL).Bcc = aEmailComponent(S_BCC_EMAIL)
				aEmailComponent(O_EMAIL).From = aEmailComponent(S_FROM_EMAIL)
				aEmailComponent(O_EMAIL).Subject = aEmailComponent(S_SUBJECT_EMAIL)
				aEmailComponent(O_EMAIL).Body = aEmailComponent(S_BODY_EMAIL)
				If Len(aEmailComponent(S_ATTACHMENTS_EMAIL)) > 0 Then
					asAttachments = Split(aEmailComponent(S_ATTACHMENTS_EMAIL), ",", -1, vbBinaryCompare)
					For iIndex = 0 To UBound(asAttachments)
						aEmailComponent(O_EMAIL).AttachFile(asAttachments(iIndex))
					Next
				End If
				aEmailComponent(O_EMAIL).Send
			Else
				aEmailComponent(O_EMAIL).RemoteHost = aEmailComponent(S_SERVER_NAME_EMAIL)
				Call aEmailComponent(O_EMAIL).AddRecipient(aEmailComponent(S_TO_EMAIL), aEmailComponent(S_TO_EMAIL))
				Call aEmailComponent(O_EMAIL).AddRecipient(aEmailComponent(S_CC_EMAIL), aEmailComponent(S_CC_EMAIL))
				Call aEmailComponent(O_EMAIL).AddRecipient(aEmailComponent(S_BCC_EMAIL), aEmailComponent(S_BCC_EMAIL))
				aEmailComponent(O_EMAIL).FromName = aEmailComponent(S_FROM_EMAIL)
				aEmailComponent(O_EMAIL).FromAddress = aEmailComponent(S_FROM_EMAIL)
				aEmailComponent(O_EMAIL).Subject = aEmailComponent(S_SUBJECT_EMAIL)
				aEmailComponent(O_EMAIL).BodyText = aEmailComponent(S_BODY_EMAIL)
				Call aEmailComponent(O_EMAIL).SendMail()
			End If

			lErrorNumber = Err.number
			sErrorDescription = Err.description
			sLogEntry = sLogEntry & "From: " & aEmailComponent(S_FROM_EMAIL)
			sLogEntry = sLogEntry & "<BR />To: " & aEmailComponent(S_TO_EMAIL)
			sLogEntry = sLogEntry & "<BR />Cc: " & aEmailComponent(S_CC_EMAIL)
			sLogEntry = sLogEntry & "<BR />Bcc: " & aEmailComponent(S_BCC_EMAIL)
			sLogEntry = sLogEntry & "<BR />Subject: " & aEmailComponent(S_SUBJECT_EMAIL)
			sLogEntry = sLogEntry & "<BR />Attachments: " & aEmailComponent(S_ATTACHMENTS_EMAIL)
			sLogEntry = sLogEntry & "<BR />Body length: " & Len(aEmailComponent(S_BODY_EMAIL)) & " characters"
			Call LogErrorInXMLFile(lErrorNumber, sLogEntry, 0, "EmailComponent.asp", S_FUNCTION_NAME, N_EMAIL_LEVEL)

			If lErrorNumber <> 0 Then
				If Len(Err.description) > 0 Then
					sLogEntry = aIntlComponent(AS_DESCRIPTORS_INTL)(1431) & "&nbsp;" & Err.description & "<BR />" & sLogEntry 'Error del servidor SMTP:
				End If
				sLogEntry = aIntlComponent(AS_DESCRIPTORS_INTL)(1430) & "<BR />" & sLogEntry 'No se pudo enviar el mensaje.
				Call LogErrorInXMLFile(lErrorNumber, sLogEntry, 1430, "EmailComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
			End If
			aEmailComponent(O_EMAIL) = Null
		End If
	End If

	SendEmail = lErrorNumber
	Err.Clear
End Function
%>