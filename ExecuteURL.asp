<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- include file="Libraries/LoginComponent.asp" -->
<%
Dim sURL
Dim iIndex

'If Not aLoginComponent(B_ADMINISTRATOR_LOGIN) And Not aLoginComponent(B_TUTOR_LOGIN) Then
'	Response.Redirect REDIRECT_INVALID_USER_PAGE & "?Permission=0"
'End If

Call LogErrorInXMLFile(0, "El usuario entró en la consola de URL<BR />URL: " & sURL, 000, "_root", "_root", N_MESSAGE_LEVEL)
aHeaderComponent(S_TITLE_NAME_HEADER) = "Enviar URL"
%>
<HTML>
	<HEAD>
		<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1" />
		<TITLE><%Response.Write aHeaderComponent(S_WINDOW_TITLE_HEADER)%></TITLE>
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/CommonLibrary.js"></SCRIPT>
		<SCRIPT LANGUAGE="JavaScript"><!--
			function SendURLAsHiddenToForm(sURL) {
				var oTargetForm = document.ExecuteURLFrm;
				var sTextFields = '';

				if (sURL != '') {
					if (oTargetForm) {
						var aURL = sURL.split('?');
						var aURLElements = aURL[1].split('&');
						var aURLElement;
						var sValue = '';

						for (var i=0; i<aURLElements.length; i++) {
							aURLElement = aURLElements[i].split('=');
							sTextFields += aURLElement[0] + ': <INPUT TYPE="TEXT" NAME="' + aURLElement[0] + '" VALUE="' + aURLElement[1] + '" /><BR />';
						}
						oTargetForm.action = aURL[0];
						oTargetForm.innerHTML = sTextFields + '<INPUT TYPE="SUBMIT" VALUE="Ejecutar URL" CLASS="Buttons" />';
					}
				}
			} // End of SendURLAsHiddenToForm
		//--></SCRIPT>
		<LINK REL="STYLESHEET" TYPE="text/css" HREF="Styles/SIAP.css" />
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" onLoad="window.focus()">
		<!-- BEGIN: CONTENTS -->
		<FONT FACE="Arial" SIZE="2"><BR /><BR />Usted se encuentra aquí:&nbsp;<A HREF="Main.asp">Inicio</A> > <B>Enviar URL</B><BR /><BR /></FONT>
		<FORM NAME="URLFrm" ID="URLFrm">
			&nbsp;<FONT FACE="Arial" SIZE="2">Introduzca el URL:</FONT><BR />
			&nbsp;<INPUT TYPE="TEXT" NAME="URL" ID="URLTxt" SIZE="60" VALUE="<%Response.Write sURL%>" CLASS="TextFields" /><BR /><BR />
			&nbsp;<INPUT TYPE="BUTTON" VALUE="Enviar URL" CLASS="Buttons" onClick="SendURLAsHiddenToForm(this.form.URL.value)" />
		</FORM>
		<FORM NAME="ExecuteURLFrm" ID="ExecuteURLFrm" ACTION="" METHOD="POST">
		</FORM>
		<BR />
		<FONT FACE="Arial" SIZE="2">
			<%If Len(sURL) > 0 Then
				Response.Write "<B>" & CleanStringForHTML(sURL) & "</B><BR />"
			End If%>
		</FONT><BR /><BR /><BR />
		<!-- END: CONTENTS -->
	</BODY>
</HTML>