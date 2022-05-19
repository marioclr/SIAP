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
<!-- #include file="Libraries/ReportsLib.asp" -->
<%
Dim oItem
aHeaderComponent(L_SELECTED_OPTION_HEADER) = REPORTS_TOOLBAR
aHeaderComponent(S_TITLE_NAME_HEADER) = "Plantilla"
For Each oItem In oRequest("Template")
	sFlags = sFlags & CLng(oItem) & ","
Next
%>
<HTML>
	<HEAD>
		<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1" />
		<META HTTP-EQUIV="Page-Exit" CONTENT="BlendTrans(Duration=0.5)" />
		<TITLE><%Response.Write aHeaderComponent(S_WINDOW_TITLE_HEADER)%></TITLE>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/CommonLibrary.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/Events.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/RollOver.js"></SCRIPT>
		<LINK REL="STYLESHEET" TYPE="text/css" HREF="Styles/SIAP.css" />
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%lErrorNumber = DisplayReportTableTemplate(oRequest, sFlags, sErrorDescription)
		If lErrorNumber <> 0 Then
			Call DisplayErrorMessage("Error en la plantilla del reporte", sErrorDescription)
		End If%>
	</BODY>
</HTML>