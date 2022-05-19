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
Dim iIndex

Call InitializeUserComponent(oRequest, aUserComponent)
If aUserComponent(N_ID_USER) > -1 Then
	lErrorNumber = GetUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
	If lErrorNumber = 0 Then
		lErrorNumber = SendMessageToNewUser(aUserComponent, sErrorDescription)
	End If
ElseIf (Len(oRequest("StartID").Item) > 0) And (Len(oRequest("EndID").Item) > 0) Then
	For iIndex = CLng(oRequest("StartID").Item) To CLng(oRequest("EndID").Item)
		aUserComponent(N_ID_USER) = iIndex
		lErrorNumber = GetUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			lErrorNumber = SendMessageToNewUser(aUserComponent, sErrorDescription)
		End If
	Next
End If
Response.Cookies("SoS_SectionID") = 208
%>
<HTML>
	<HEAD>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" onLoad="window.focus()">
		<%If lErrorNumber <> 0 Then
			Call DisplayErrorMessage("Error al enviar el mensaje", sErrorDescription)
		Else%>
			<SCRIPT LANGUAGE="JavaScript"><!--
				window.close();
			//--></SCRIPT>
		<%End If%>
	</BODY>
</HTML>