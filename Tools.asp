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
aHeaderComponent(L_SELECTED_OPTION_HEADER) = TOOLS_TOOLBAR
aHeaderComponent(S_TITLE_NAME_HEADER) = "Herramientas"
bWaitMessage = False
Response.Cookies("SoS_SectionID") = 207
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<!-- #include file="_Header.asp" -->
		Usted se encuentra aquí: <A HREF="Main.asp">Inicio</A> > <B>Herramientas</B><BR />
		<BR /><BR /><TABLE WIDTH="900" BORDER="0" CELLPADDING="0" CELLSPACING="0">
		<%aMenuComponent(A_ELEMENTS_MENU) = Array(_
			Array("Actualizar descriptores",_
				  "",_
				  "Images/MnLeftArrows.gif", "ReloadDescriptors.asp", (InStr(1, aLoginComponent(S_ACCESS_KEY_LOGIN), "vac", vbBinaryCompare) = 1)),_
			Array("Ejecutar URL",_
				  "",_
				  "Images/MnLeftArrows.gif", "ExecuteURL.asp", (InStr(1, aLoginComponent(S_ACCESS_KEY_LOGIN), "vac", vbBinaryCompare) = 1)),_
			Array("<LINE />", "", "", "", (InStr(1, aLoginComponent(S_ACCESS_KEY_LOGIN), "vac", vbBinaryCompare) = 1)),_
			Array("Bitácora de errores",_
				  "Revise el desempeño de la aplicación a través de la bitácora que muestra a detalle los errores producidos durante el uso de este sistema. En caso de encontrar problemas serios contácte al distribuidor del sistema.",_
				  "Images/MnLeftArrows.gif", "ErrorLog.asp", N_TOOLS_PERMISSIONS),_
			Array("Cambiar contraseña",_
				  "Se recomienda cambiar de contraseña cada cierto tiempo para proteger su cuenta y mantener la seguridad del sistema.",_
				  "Images/MnLeftArrows.gif", "ChangePassword.asp", True),_
			Array("Opciones del sistema",_
				  "Modifique los valores que utiliza el sistema para generar reportes.",_
				  "Images/MnLeftArrows.gif", "Options.asp?Admin=1", N_TOOLS_PERMISSIONS),_
			Array("Preferencias del usuario",_
				  "Cambie su página de inicio, modifique el estilo de las tablas en sus reportes, etc.",_
				  "Images/MnLeftArrows.gif", "Options.asp", True),_
			Array("<LINE />", "", "", "", True),_
			Array("Curso de capacitación",_
				  "Curso interactivo para aprender a utilizar el Sistema de Administración del Personal.",_
				  "Images/MnLeftArrows.gif", "javascript: OpenNewWindow('Course\/index.htm', '', 'Curso', 961, 643, 'yes', 'yes');", False),_
			Array("Manuales de uso",_
				  "Manuales de uso de los distintos módulos.",_
				  "Images/MnLeftArrows.gif", "Docs.asp", True),_
			Array("Soporte técnico",_
				  "¿Necesita ayuda técnica? Registre su caso en nuestro sistema de soporte técnico.",_
				  "Images/MnLeftArrows.gif", "http://200.57.131.24/SoS/AddEvent.asp?SystemID=9&ModuleID=" & Request.Cookies("SoS_SectionID") & "&UserEmail=" & aLoginComponent(S_USER_E_MAIL_LOGIN) & """ TARGET=""SoS_AddEvent", aLoginComponent(B_TECH_SUPPORT_LOGIN))_
		)
		aMenuComponent(B_USE_DIV_MENU) = True
		Call DisplayMenuInThreeSmallColumns(aMenuComponent)%>
		</TABLE>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>