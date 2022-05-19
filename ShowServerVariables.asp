<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next

Dim iIndex
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<HTML>
	<HEAD>
		<TITLE>Show Server Variables</TITLE>
	</HEAD>
	<BODY BGCOLOR="#FFFFFF" TOPMARGIN="0" LEFTMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<FONT FACE="Verdana" SIZE="1"><%
			Response.Write "<B>Date</B>: " & Date() & " " & Time() & "<BR />"
			For Each iIndex In Request.ServerVariables 
				Response.Write "<B>" & iIndex & "</B>: " & Request.ServerVariables(iIndex) & "<BR />"
			Next
			Response.Write "<BR /><B>Server.ScriptTimeout</B>: " & Server.ScriptTimeout & "<BR />"
			Response.Write "<B>Session.Timeout</B>: " & Session.Timeout & "<BR />"
			Response.Write "<B>Session ID:</B> &nbsp;&nbsp;&nbsp;" & Session.SessionID & "<BR /><BR />"
			For Each iIndex In Application.Contents 
				Response.Write "<B>" & iIndex & "</B>: " & Application.Contents(iIndex) & "<BR />"
			Next
		%></FONT>
	</BODY>
</HTML>