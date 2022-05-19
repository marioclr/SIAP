<%
Function SendPasswordToUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
'************************************************************
'Purpose: To send an email to the user with his/her password
'Inputs:  oRequest, oADODBConnection, aUserComponent
'Outputs: aUserComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SendPasswordToUser"
	Dim aMessageComponent
	Dim lErrorNumber

	lErrorNumber = CheckExistencyOfUser(oADODBConnection, False, aUserComponent, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not aUserComponent(B_IS_DUPLICATED_USER) Then
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "Usted no tiene una cuenta asignada dentro del sistema. Favor de contactar al administrador del sistema para que le asigne una cuenta y le proporcione sus credenciales de entrada."
		Else
			lErrorNumber = GetUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				If Len(aUserComponent(S_EMAIL_USER)) = 0 Then
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "Su cuenta no tiene registrada su correo electrónico. Favor de contactar al administrador del sistema para que le asigne una cuenta y le proporcione sus credenciales de entrada."
				Else
					Redim aMessageComponent(N_EMAIL_COMPONENT_SIZE)
					aMessageComponent(S_TO_EMAIL) = aUserComponent(S_EMAIL_USER)
					aMessageComponent(S_FROM_EMAIL) = S_ADMIN_EMAIL_ACCOUNT
					aMessageComponent(S_SUBJECT_EMAIL) = "Mensaje enviado por SIAP: Solicitud de contraseña"
					aMessageComponent(S_BODY_EMAIL) = "<FONT FACE=""Verdana"" SIZE=""2"">" & _
														"A petición suya le enviamos su contraseña para que pueda ingresar al Sistema de Administración del Personal.<BR /><BR />" & _
														"<TABLE BGCOLOR=""#000000"" BORDER=""0"" ALIGN=""CENTER"" CELLSPACING=""0"" CELLPADDING=""1""><TR><TD>" & _
															"<TABLE WIDTH=""100%"" BGCOLOR=""#FFFFFF"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & _
																"<TR>" & _
																	"<TD><FONT FACE=""Verdana"" SIZE=""2"">Clave de acceso:&nbsp;</FONT></TD>" & _
																	"<TD BGCOLOR=""#000000""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>" & _
																	"<TD><FONT FACE=""Verdana"" SIZE=""2"">" & _
																		aUserComponent(S_ACCESS_KEY_USER) & _
																	"</FONT></TD>" & _
																"</TR>" & _
																"<TR><TD COLSPAN=""3"" BGCOLOR=""#000000""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>" & _
																"<TR>" & _
																	"<TD><FONT FACE=""Verdana"" SIZE=""2"">Contraseña:&nbsp;</FONT></TD>" & _
																	"<TD BGCOLOR=""#000000""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>" & _
																	"<TD><FONT FACE=""Verdana"" SIZE=""2"">" & _
																		aUserComponent(S_PASSWORD_USER) & _
																	"</FONT></TD>" & _
																"</TR>" & _
															"</TABLE>" & _
														"</TD></TR></TABLE><BR />" & _
														"Le recomendamos que la mantenga en un lugar seguro. De ser posible borre este mensaje de su bandeja de entrada para evitar que otras personas tengan acceso a esta información." & _
													  "</FONT>"
					lErrorNumber = SendEmail(oRequest, aMessageComponent, sErrorDescription)
				End If
			Else
				lErrorNumber = -1
				sErrorDescription = "No se pudo enviar su contraseña a su cuenta de correo electrónico. Favor de contactar al administrador del sistema para que le proporcione sus credenciales de entrada."
			End If
		End If
	Else
		lErrorNumber = -1
		sErrorDescription = "No se pudo enviar su contraseña a su cuenta de correo electrónico. Favor de contactar al administrador del sistema para que le proporcione sus credenciales de entrada."
	End If

	SendPasswordToUser = lErrorNumber
End Function
%>