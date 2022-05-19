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
<%
aHeaderComponent(L_SELECTED_OPTION_HEADER) = HOME_TOOLBAR
aHeaderComponent(S_TITLE_NAME_HEADER) = "Access Denied"
bWaitMessage = False
Response.Cookies("SoS_SectionID") = 187
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<!-- #include file="_Header.asp" -->
		Usted se encuentra aquí: <A HREF="Main.asp">Inicio</A><BR /><BR /><BR />
		Usted no cuenta con los permisos necesarios para ver el módulo solicitado.<BR /><BR />
		<FORM><INPUT TYPE="BUTTON" NAME="Back" ID="BackBtn" VALUE="Regresar" onClick="window.history.go(-1)" CLASS="Buttons" /></FORM>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>