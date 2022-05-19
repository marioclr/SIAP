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
<!-- #include file="Libraries/HelpLibrary.asp" -->
<!-- #include file="Libraries/XMLLibrary.asp" -->
<%
aHeaderComponent(S_TITLE_NAME_HEADER) = "Ayuda en Línea"
bWaitMessage = False
Response.Cookies("SoS_SectionID") = 203
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" onLoad="window.focus()">
		<!-- #include file="_HelpHeader.asp" -->
		<TABLE WIDTH="720" BORDER="0" CELLPADDING="0" CELLSPACING="0"><TR>
			<TD WIDTH="199" VALIGN="TOP"><FONT FACE="Verdana" SIZE="1">
				<%lErrorNumber = DisplayHelpMenu(iHelpSection, sErrorDescription)%>
			</FONT></TD>
			<TD BGCOLOR="#<%Response.Write S_MAIN_COLOR_FOR_GUI%>"><IMG SRC="Images/Transparent.gif" WIDTH="1" HEIGHT="1" /></TD>
			<TD><IMG SRC="Images/Transparent.gif" WIDTH="5" HEIGHT="1" /></TD>
			<TD WIDTH="515" VALIGN="TOP"><FONT FACE="Arial" SIZE="2"><%
				If lErrorNumber <> 0 Then
					Call DisplayErrorMessage("Error en la ayuda", sErrorDescription)
					lErrorNumber = 0
					sErrorDescription = ""
				End If
				lErrorNumber = DisplayHelpSection(iHelpSection, sErrorDescription)
				If lErrorNumber <> 0 Then
					Call DisplayErrorMessage("Error en la ayuda", sErrorDescription)
					lErrorNumber = 0
					sErrorDescription = ""
				End If
			%></FONT></TD>
		</TR></TABLE>
		<!-- #include file="_HelpFooter.asp" -->
	</BODY>
</HTML>