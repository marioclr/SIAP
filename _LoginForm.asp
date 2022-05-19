<%If Not B_PORTAL Then%>
	<CENTER><FORM NAME="LoginFrm" ID="LoginFrm" ACTION="<%
		Response.Write GetASPFileName("") & "?Session=" & GetSerialNumberForDate("")
	%>760211" METHOD="POST" onSubmit="return CheckLoginFields(this);">
		<TABLE BGCOLOR="#<%Response.Write S_MAIN_COLOR_FOR_GUI%>" BORDER="0" CELLSPACING="0" CELLPADDING="1"><TR><TD>
			<TABLE BGCOLOR="#FFFFFF" WIDTH="100%" BORDER="0" CELLSPACING="2" CELLPADDING="0">
				<TR>
					<TD ROWSPAN="3">&nbsp;&nbsp;<IMG SRC="Images/PicLogin.gif" WIDTH="64" HEIGHT="64" />&nbsp;&nbsp;&nbsp;</TD>
					<TD><FONT FACE="Arial" SIZE="2"><NOBR>Clave de acceso:</NOBR></FONT></TD>
					<TD>&nbsp;&nbsp;</TD>
					<TD WIDTH="1"><INPUT TYPE="TEXT" NAME="AccessKey" ID="AccessKeyTxt" VALUE="<%Response.Write aLoginComponent(S_ACCESS_KEY_LOGIN)%>" SIZE="20" MAXLENGTH="100" CLASS="TextFields" onKeyDown="StartClockToCleanFrom()" /></TD>
					<TD>&nbsp;&nbsp;</TD>
				</TR>
				<TR>
					<TD><FONT FACE="Arial" SIZE="2">Contraseña:</FONT></TD>
					<TD>&nbsp;&nbsp;</TD>
					<TD WIDTH="1"><INPUT TYPE="PASSWORD" NAME="Password" ID="PasswordPwd" VALUE="" SIZE="20" MAXLENGTH="100" CLASS="TextFields" onKeyDown="StartClockToCleanFrom()" /></TD>
					<TD>&nbsp;&nbsp;</TD>
				</TR>
				<TR>
					<TD ALIGN="RIGHT" COLSPAN="3"><BR /><INPUT TYPE="SUBMIT" NAME="Login" ID="LoginBtn" VALUE="Entrar" CLASS="Buttons" /></TD>
					<TD>&nbsp;&nbsp;</TD>
				</TR>
				<TR><TD COLSPAN="5"><IMG SRC="Images/Transparent.gif"  WIDTH="0" HEIGHT="3" /></TD></TR>
				<TR><TD BGCOLOR="#<%Response.Write S_WIDGET_BGCOLOR_FOR_GUI%>" COLSPAN="5"><FONT FACE="Arial" SIZE="2">&nbsp;<B>Nota: </B>Recuerde que la clave de acceso y la contraseña&nbsp;<BR />&nbsp;son sensibles a mayúsculas y minúsculas.<BR /></FONT></TD></TR>
			</TABLE>
		</TD></TR></TABLE><BR />
		<%If Len(oRequest("Redirect").Item) > 0 Then
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Redirect"" ID=""RedirectHdn"" VALUE=""" & oRequest("Redirect").Item & """ />"
		End If
		If B_USE_SMTP Then
			If lErrorNumber = L_ERR_INCORRECT_PASSWORD Then
				Response.Write "<TABLE WIDTH=""1"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR><TD ALIGN=""CENTER""><NOBR><FONT FACE=""Arial"" SIZE=""2"">"
					Response.Write "<IMG SRC=""Images/IcnExclamation.gif"" WIDTH=""32"" HEIGHT=""32"" ALIGN=""LEFT"" HSPACE=""5"" />"
					Response.Write "<FONT SIZE=""1""><BR /></FONT><B>¿Olvidó su contraseña? <A HREF=""SendPassword.asp?UserAccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & """>Solicítela por correo electrónico</A></B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<BR />"
				Response.Write "</FONT></NOBR></TD></TR></TABLE>"
			End If
		End If%>
	</FORM></CENTER>
	<SCRIPT LANGUAGE="JavaScript"><!--
		var iCounter = 0;
		window.focus();
		document.LoginFrm.AccessKey.focus();

		function StartClockToCleanFrom() {
			iCounter = 0;
		} // End of StartClockToCleanFrom

		function CleanFrom() {
			iCounter++;
			if (iCounter >= 60) {
				document.LoginFrm.AccessKey.value = '';
				document.LoginFrm.Password.value = '';
				iCounter = 0;
			}
		} // End of CleanFrom
		window.setInterval('CleanFrom()', 1000);
	//--></SCRIPT>
<%End If%>