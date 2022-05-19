<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/LoginComponentConstants.asp" -->
<!-- #include file="Libraries/SendPasswordLib.asp" -->
<!-- #include file="Libraries/UserComponent.asp" -->
<%
If Not B_USE_SMTP Then
	Response.Redirect "Default.asp"
Else
	lErrorNumber = SendPasswordToUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
End If

aHeaderComponent(L_SELECTED_OPTION_HEADER) = NO_TOOLBAR
aHeaderComponent(S_TITLE_NAME_HEADER) = "Solicitud de Contraseña"
bWaitMessage = False
Response.Cookies("SoS_SectionID") = 208
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<!-- #include file="_Header.asp" -->
		<%If lErrorNumber = 0 Then
			Call DisplayErrorMessage("Confirmación", "Su contraseña ha sido enviada a su cuenta de correo electrónico.")
		Else
			Call DisplayErrorMessage("Error al enviar su contraseña", sErrorDescription)
		End If%>
		<BR /><BR />
		&nbsp;<A HREF="Default.asp"><B>Continuar</B></A>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>