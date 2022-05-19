<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/LogoutLib.asp" -->
<%
Call DoLogout(oRequest, Request.Cookies("SIAP_CurrentAccessKey").Item)
If True Then
	If Len(oRequest) = 0 Then
		Response.Redirect SYSTEM_PATH & "Default.asp"
	Else
		Response.Redirect SYSTEM_PATH & "Default.asp?" & oRequest
	End If
Else
	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "window.close();" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine
End If
%>