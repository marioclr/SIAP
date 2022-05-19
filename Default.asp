<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/DefaultLib.asp" -->
<!-- #include file="Libraries/LoginComponent.asp" -->
<%
Call LogCommonErrors(oRequest, sErrorDescription)

aHeaderComponent(L_SELECTED_OPTION_HEADER) = NO_TOOLBAR
aHeaderComponent(S_TITLE_NAME_HEADER) = "Introduzca su clave de acceso"
If Len(sAuxMessage) > 0 Then aHeaderComponent(S_TITLE_NAME_HEADER) = "Boletín"
bWaitMessage = False
Response.Cookies("SoS_SectionID") = 187
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript"><!--
			if (self.parent.name != 'SAPO') {
				var iScreenWidth = (screen.width) ? screen.width : 1024; 
				var iScreenHeight = (screen.height) ? screen.height : 768;
				var iWindowWidth = (iScreenWidth >= 1024) ? 1024 : 800; 
				var iWindowHeight = (iScreenHeight >= 768) ? 740 : 572;
				var iLeftPosition = (iScreenWidth - iWindowWidth) / 2; 
				var iTopPosition = (iScreenHeight - iWindowHeight - 28) / 2;

				window.moveTo(iLeftPosition, iTopPosition);
				window.resizeTo(iWindowWidth, iWindowHeight);
				window.focus();
			}

			function CheckLoginFields(oForm){
			//************************************************************
			//Purpose: To check the access key before trying to log in.
			//Inputs:  oForm
			//************************************************************
				if (oForm.AccessKey.value.length == 0) {
					alert('Favor de introducir la clave de accesso.');
					oForm.AccessKey.focus();
					return false;
				}
				return true;
			} // End of CheckLoginFields
		//--></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<!-- #include file="_Header.asp" -->
		<BR /><BR />
		<%If Len(sErrorDescription) > 0 Then
			Call DisplayErrorMessage("Error al validar las credenciales de entrada", sErrorDescription)
		End If
		If Len(sAuxMessage) = 0 Then
			Response.Write "<BR />"
			%><!-- #include file="_LoginForm.asp" --><%
			Response.Write "<BR /><BR />"
			Call LaunchIntro()
			Response.Write "<BR /><BR />"
			'Call ShowFlashPlayerMessage()
			Response.Write "<BR /><BR />"
		Else
			Response.Write sAuxMessage 
		End If%>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>