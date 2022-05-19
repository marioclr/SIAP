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
<!-- #include file="Libraries/UserComponent.asp" -->
<%
bWaitMessage = True

aLoginComponent(S_ACCESS_KEY_LOGIN) = "admin"
Call InitializePermissionsForLoginComponent(oRequest, aLoginComponent)

Call InitializeUserComponent(oRequest, aUserComponent)
If Len(aUserComponent(S_ACCESS_KEY_USER)) > 0 Then
	Call CheckExistencyOfUser(oADODBConnection, False, aUserComponent, sErrorDescription)
	aUserComponent(B_IS_DUPLICATED_USER) = False
End If
%>
<HTML>
	<HEAD>
		<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1" />
		<TITLE>Sistema Integral de Administración de Personal del ISSSTE</TITLE>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/CheckFields.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/CommonLibrary.js"></SCRIPT>
		<LINK REL="STYLESHEET" TYPE="text/css" HREF="Styles/SIAP.css" />
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<FONT FACE="Arial" SIZE="2"><B>Cuenta del usuario en el Sistema Integral de Administración del Personal<BR /><BR /></B></FONT>
		<%Call DisplayUserForm(oRequest, oADODBConnection, GetASPFileName(""), aUserComponent, sErrorDescription)%>
	</BODY>
</HTML>