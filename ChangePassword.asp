<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/LoginComponent.asp" -->
<!-- #include file="Libraries/UserComponent.asp" -->
<%
If Len(oRequest("ChangePwd")) > 0 Then
	aUserComponent(N_ID_USER) = aLoginComponent(N_USER_ID_LOGIN)
	lErrorNumber = GetUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
	If lErrorNumber = 0 Then
		If StrComp(aUserComponent(S_PASSWORD_USER), oRequest("CurrentPassword").Item, vbBinaryCompare) <> 0 Then
			lErrorNumber = -1
			sErrorDescription = "La contraseña proporcionada como la actual no corresponde con la registrada en el sistema."
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ChangePassword.asp", "_root", N_WARNING_LEVEL)
		ElseIf StrComp(aUserComponent(S_OLD_PASSWORD_USER), oRequest("NewPassword").Item, vbBinaryCompare) = 0 Then
			lErrorNumber = -1
			sErrorDescription = "La contraseña proporcionada ya fue utilizada anteriormente. Favor de seleccionar otra contraseña."
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ChangePassword.asp", "_root", N_WARNING_LEVEL)
		Else
			aUserComponent(S_PASSWORD_USER) = oRequest("NewPassword").Item
			lErrorNumber = ModifyUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
		End If
	End If
End If
If Len(oRequest("ChangeEmail")) > 0 Then
	aUserComponent(N_ID_USER) = aLoginComponent(N_USER_ID_LOGIN)
	lErrorNumber = GetUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
	If lErrorNumber = 0 Then
		aUserComponent(S_ADDITIONAL_EMAIL_USER) = oRequest("AdditionalEmail").Item
		aUserComponent(N_ACTIVE_USER) = CInt(oRequest("UserActive").Item)
		aLoginComponent(S_USER_ADDITIONAL_E_MAIL_LOGIN) = aUserComponent(S_ADDITIONAL_EMAIL_USER)
		aLoginComponent(B_ACTIVE_LOGIN) = (aUserComponent(N_ACTIVE_USER) = 1)
		lErrorNumber = ModifyUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
	End If
End If

aHeaderComponent(L_SELECTED_OPTION_HEADER) = TOOLS_TOOLBAR
aHeaderComponent(S_TITLE_NAME_HEADER) = "Cambiar Contraseña"
bWaitMessage = True
Response.Cookies("SoS_SectionID") = 199
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript"><!--
			function CheckPwdFields(oForm) {
				if (oForm) {
					if (oForm.CurrentPassword.value.length == 0) {
						alert('Favor de introducir su contraseña actual.');
						oForm.CurrentPassword.focus();
						return false;
					}
					if (oForm.NewPassword.value.length == 0) {
						alert('Favor de introducir su nueva contraseña.');
						oForm.NewPassword.focus();
						return false;
					}
					if (oForm.NewPassword.value != oForm.UserPwdConfirmation.value) {
						alert('Su nueva contraseña no coincide con la confirmación. Favor de introducirlas de nuevo.');
						oForm.NewPassword.value = '';
						oForm.UserPwdConfirmation.value = '';
						oForm.NewPassword.focus();
						return false;
					}
					if (oForm.CurrentPassword.value == oForm.NewPassword.value) {
						alert('Su nueva contraseña coincide con la actual. Favor de seleccionar otra contraseña e introducirla de nuevo.');
						oForm.NewPassword.value = '';
						oForm.UserPwdConfirmation.value = '';
						oForm.NewPassword.focus();
						return false;
					}
				}
				return true;
			} // End of CheckPwdFields

			function CheckEmailFields(oForm) {
				if (oForm) {
					if ((oForm.UserActive.value == '0') && (oForm.AdditionalEmail.value.length == 0)) {
						alert('Favor de introducir la(s) cuenta(s) de correo electrónico.');
						oForm.AdditionalEmail.focus();
						return false;
					}
				}
				return true;
			} // End of CheckEmailFields
		//--></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<!-- #include file="_Header.asp" -->
		Usted se encuentra aquí: <A HREF="Main.asp">Inicio</A> > <A HREF="Tools.asp">Herramientas</A> > <B>Cambiar contraseña</B><BR /><BR /><BR />
		<%If lErrorNumber <> 0 Then
			Call DisplayErrorMessage("Error al modificar la contraseña", sErrorDescription)
			Response.Write "<BR />"
		End If
		If (Len(oRequest("ChangePwd")) > 0) And (lErrorNumber = 0) Then
			If B_SADE Then
				Call DisplayErrorMessage("Confirmación", "La contraseña fue modificada con éxito<BR />Si usted se encuentra utilizando SADE le recomendamos que cierre su sesión presionando SALIR y vuelva a entrar al sistema.<BR /><BR /><A HREF=""Main.asp""><B>Continuar</B></A>")
			Else
				Call DisplayErrorMessage("Confirmación", "La contraseña fue modificada con éxito<BR /><BR /><A HREF=""Main.asp""><B>Continuar</B></A>")
			End If
			Response.Write "<BR />"
		Else
			If Len(oRequest("Expired").Item) > 0 Then
				Call DisplayErrorMessage("Su contraseña ha expirado", "Para fines de seguridad es necesario que usted cambie su contraseña periódicamente.")
				Response.Write "<BR />"
			End If%>
			<FONT FACE="Arial" SIZE="2"><B>¿Desea cambiar su contraseña?<BR />Indique su contraseña actual e introduzca una nueva contraseña.</B></FONT>
			<FORM NAME="ChangePwdFrm" ID="ChangePwdFrm" ACTION="ChangePassword.asp" METHOD="POST" onSubmit="return CheckPwdFields(this)">
				<FONT FACE="Arial" SIZE="2">Contraseña actual: </FONT><INPUT TYPE="PASSWORD" NAME="CurrentPassword" ID="CurrentPasswordPwd" SIZE="30" MAXLENGTH="30" VALUE="" CLASS="TextFields" /><BR />
				<FONT FACE="Arial" SIZE="2">Contraseña nueva: </FONT><IMG SRC="Images/Transparent.gif" WIDTH="2" HEIGHT="1" /><INPUT TYPE="PASSWORD" NAME="NewPassword" ID="NewPasswordTxt" SIZE="30" MAXLENGTH="30" VALUE="" CLASS="TextFields" /><BR />
				<FONT FACE="Arial" SIZE="2">Confirmación: </FONT><IMG SRC="Images/Transparent.gif" WIDTH="30" HEIGHT="1" /><INPUT TYPE="PASSWORD" NAME="UserPwdConfirmation" ID="UserPwdConfirmationTxt" SIZE="30" MAXLENGTH="30" VALUE="" CLASS="TextFields" /><BR />
				<BR />
				<INPUT TYPE="SUBMIT" NAME="ChangePwd" ID="ChangePwdBtn" VALUE="Cambiar Contraseña" CLASS="Buttons" />
				<IMG SRC="Images/Transparent.gif" WIDTH="113" HEIGHT="1" />
				<INPUT TYPE="BUTTON" NAME="Cancel" ID="CancelBtn" VALUE="Cancelar" CLASS="Buttons" onClick="window.location.href='Main.asp'" />
			</FORM><BR />
			<FONT FACE="Arial" SIZE="2"><B>¿Desea cambiar su e-mail adicional?<BR />Indique la(s) dirección(es) de correo electrónico a las que desea que le lleguen copias de sus alarmas.</B></FONT>
			<FORM NAME="ChangeEmailFrm" ID="ChangeEmailFrm" ACTION="ChangePassword.asp" METHOD="POST" onSubmit="return CheckEmailFields(this)">
				<FONT FACE="Arial" SIZE="2">Correo(s) electrónico(s): </FONT><INPUT TYPE="TEXT" NAME="AdditionalEmail" ID="AdditionalEmailTxt" SIZE="30" MAXLENGTH="100" VALUE="<%Response.Write aLoginComponent(S_USER_ADDITIONAL_E_MAIL_LOGIN)%>" CLASS="TextFields" /><BR />
				<INPUT TYPE="CHECKBOX" NAME="DummyUserActive" ID="DummyUserActiveChk" onClick="SetHiddenValueForCheckBox(!this.checked, this.form.UserActive)"<%
					If Not aLoginComponent(B_ACTIVE_LOGIN) Then
						Response.Write " CHECKED=""1"""
					End If
				%>/>
				<INPUT TYPE="HIDDEN" NAME="UserActive" ID="UserActiveHdn" VALUE="<%If aLoginComponent(B_ACTIVE_LOGIN) Then Response.Write "1" Else Response.Write "0" End If%>" />
				<FONT FACE="Arial" SIZE="2">Enviar copia de alarmas a mis cuenta(s) de correo adicional(es)</FONT>
				<BR /><BR />
				<INPUT TYPE="SUBMIT" NAME="ChangeEmail" ID="ChangeEmailBtn" VALUE="Guardar Cambios" CLASS="Buttons" />
				<IMG SRC="Images/Transparent.gif" WIDTH="113" HEIGHT="1" />
				<INPUT TYPE="BUTTON" NAME="Cancel" ID="CancelBtn" VALUE="Cancelar" CLASS="Buttons" onClick="window.location.href='Main.asp'" />
			</FORM>
		<%End If%>
		<BR /><BR /><BR /><BR />
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>