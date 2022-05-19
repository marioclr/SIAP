<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next

Dim oCookie
Dim oCookieKey
Dim oSessionVariable
%>
<HTML>
	<HEAD>
		<TITLE>Show Cookies</TITLE>
	</HEAD>
	<BODY BGCOLOR="#FFFFFF" TOPMARGIN="0" LEFTMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<FONT FACE="Verdana" SIZE="1"><%
			Response.Write "<B>Date</B>: " & Date() & " " & Time() & "<BR />"
			For Each oCookie In Request.Cookies
				If Request.Cookies(oCookie).HasKeys Then
					For Each oCookieKey In Request.Cookies(oCookie)
						Response.Write "<B>" & oCookie & "(" & oCookieKey & "):</B> "
						Response.Write "&nbsp;&nbsp;&nbsp;" & Request.Cookies(oCookie)(oCookieKey) & "<BR />"
					Next
				Else
					Response.Write "<B>" & oCookie & ":</B> "
					Response.Write "&nbsp;&nbsp;&nbsp;" & Request.Cookies(oCookie) & "<BR />"
				End If
			Next
			Response.Write "<BR /><B>Session ID:</B> &nbsp;&nbsp;&nbsp;" & Session.SessionID & "<BR />"
			Response.Write "<B>Session Timeout:</B> &nbsp;&nbsp;&nbsp;" & Session.Timeout & "<BR />"
			For Each oSessionVariable In Session.Contents
				Response.Write "<B>" & oSessionVariable & ":</B> "
				Response.Write "&nbsp;&nbsp;&nbsp;" & Session.Contents(oSessionVariable) & "<BR />"
			Next
		%></FONT>
	</BODY>
</HTML>