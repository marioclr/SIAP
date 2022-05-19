<!-- BEGIN: HEADER -->
<%
Dim sBoldBeginForHeader
Dim sBoldEndForHeader
Dim iLeftForPopupMenu
iLeftForPopupMenu = 121
If bIsMac Or bIsNetscape Then bWaitMessage = False
If bWaitMessage Then
%>
<DIV ID="WaitDiv" CLASS="ClassPopupItem" STYLE="top: 200px; visibility: visible;">
	<TABLE WIDTH="100%" BORDER="0" CELLPADDING="0" CELLSPACING="0"><TR><TD ALIGN="CENTER">
		<IMG SRC="Images/AniWait.gif" WIDTH="100" HEIGHT="100" ALT="Cargando información..." /><BR /><BR />
		<FONT FACE="Arial" SIZE="2"><B>Cargando información...</B></FONT>
	</TD></TR></TABLE>
</DIV>
<%
Call Response.Flush()
End If%>
<TABLE BORDER="0" WIDTH="100%" HEIGHT="65" CELLSPACING="0" CELLPADDING="0">
	<TR>
		<TD WIDTH="1" ROWSPAN="2"><IMG SRC="Images/TtlLogo.gif" WIDTH="115" HEIGHT="65" /></TD>
		<TD BGCOLOR="#<%Response.Write S_MAIN_COLOR_FOR_GUI%>" WIDTH="*" HEIGHT="33"><IMG SRC="Images/TtlSIAP.gif" WIDTH="494" HEIGHT="33" /></TD>
		<TD WIDTH="161" BACKGROUND="Images/TtlPicture.gif" ROWSPAN="2" ALIGN="RIGHT" VALIGN="TOP"><IMG SRC="Images/Transparent.gif" WIDTH="161" HEIGHT="1" /><BR /><%
			If aLoginComponent(B_VALID_USER_LOGIN) Then
				Response.Write "<A HREF=""Logout.asp""><IMG SRC=""Images/BtnLogout.gif"" WIDTH=""81"" HEIGHT=""33"" ALT=""Salir"" BORDER=""0"" /></A>"
			End If
		%>&nbsp;&nbsp;</TD>
	</TR>
	<TR><TD BGCOLOR="#<%Response.Write S_MAIN2_COLOR_FOR_GUI%>" HEIGHT="32"><FONT FACE="Arial" SIZE="2" COLOR="#FFFFFF"><B><IMG SRC="Images/Transparent.gif" WIDTH="1" HEIGHT="28" />&nbsp;<%
		If aHeaderComponent(L_SELECTED_OPTION_HEADER) <> NO_TOOLBAR Then
			If aLoginComponent(B_VALID_USER_LOGIN) Then
				Response.Write "<A"
					If aHeaderComponent(L_LINKED_OPTION_HEADER) And HOME_TOOLBAR Then
						Response.Write " HREF=""Main.asp"""
					End If
				Response.Write ">"
					Response.Write "<FONT COLOR=""#"
						If aHeaderComponent(L_SELECTED_OPTION_HEADER) And HOME_TOOLBAR Then
							Response.Write S_SELECTED_LINK_FOR_GUI
						Else
							Response.Write S_MENU_LINK_FOR_GUI
						End If
					Response.Write """ CLASS=""SpecialLink"" STYLE=""font-size: 13px;"">INICIO</FONT>"
				Response.Write "</A>&nbsp;&#183;&nbsp;"
				iLeftForPopupMenu = iLeftForPopupMenu + 52
			End If

			If B_ISSSTE Then
				If InStr(1, ",-1,0,1,", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) > 0 Then
					Response.Write "<A"
						If aHeaderComponent(L_LINKED_OPTION_HEADER) And CATALOGS_TOOLBAR Then
							Response.Write " HREF=""Main_ISSSTE.asp?SectionID=1"""
						End If
					Response.Write ">"
						Response.Write "<FONT COLOR=""#"
							If aHeaderComponent(L_SELECTED_OPTION_HEADER) And CATALOGS_TOOLBAR Then
								Response.Write S_SELECTED_LINK_FOR_GUI
							Else
								Response.Write S_MENU_LINK_FOR_GUI
							End If
						Response.Write """ CLASS=""SpecialLink"" STYLE=""font-size: 13px;"">PERSONAL</FONT>"
					Response.Write "</A>&nbsp;&#183;&nbsp;"
					iLeftForPopupMenu = iLeftForPopupMenu + 83
				End If

				If InStr(1, ",-1,0,2,", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) > 0 Then
					Response.Write "<A"
						If aHeaderComponent(L_LINKED_OPTION_HEADER) And PAYMENTS_TOOLBAR Then
							Response.Write " HREF=""Main_ISSSTE.asp?SectionID=2"""
						End If
					Response.Write ">"
						Response.Write "<FONT COLOR=""#"
							If aHeaderComponent(L_SELECTED_OPTION_HEADER) And PAYMENTS_TOOLBAR Then
								Response.Write S_SELECTED_LINK_FOR_GUI
							Else
								Response.Write S_MENU_LINK_FOR_GUI
							End If
						Response.Write """ CLASS=""SpecialLink"" STYLE=""font-size: 13px;"">PRESTACIONES</FONT>"
					Response.Write "</A>&nbsp;&#183;&nbsp;"
					iLeftForPopupMenu = iLeftForPopupMenu + 113
				End If

				If InStr(1, ",-1,0,3,", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) > 0 Then
					Response.Write "<A"
						If aHeaderComponent(L_LINKED_OPTION_HEADER) And HUMAN_RESOURCES_TOOLBAR Then
							Response.Write " HREF=""Main_ISSSTE.asp?SectionID=3"""
						End If
					Response.Write ">"
						Response.Write "<FONT COLOR=""#"
							If aHeaderComponent(L_SELECTED_OPTION_HEADER) And HUMAN_RESOURCES_TOOLBAR Then
								Response.Write S_SELECTED_LINK_FOR_GUI
							Else
								Response.Write S_MENU_LINK_FOR_GUI
							End If
						Response.Write """ CLASS=""SpecialLink"" STYLE=""font-size: 13px;"">DESARROLLO&nbsp;HUMANO</FONT>"
					Response.Write "</A>&nbsp;&#183;&nbsp;"
					iLeftForPopupMenu = iLeftForPopupMenu + 162
				End If

				If InStr(1, ",-1,0,4,", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) > 0 Then
					Response.Write "<A"
						If aHeaderComponent(L_LINKED_OPTION_HEADER) And PAYROLL_TOOLBAR Then
							Response.Write " HREF=""Main_ISSSTE.asp?SectionID=4"""
						End If
					Response.Write ">"
						Response.Write "<FONT COLOR=""#"
							If aHeaderComponent(L_SELECTED_OPTION_HEADER) And PAYROLL_TOOLBAR Then
								Response.Write S_SELECTED_LINK_FOR_GUI
							Else
								Response.Write S_MENU_LINK_FOR_GUI
							End If
						Response.Write """ CLASS=""SpecialLink"" STYLE=""font-size: 13px;"">INFORMÁTICA</FONT>"
					Response.Write "</A>&nbsp;&#183;&nbsp;"
					iLeftForPopupMenu = iLeftForPopupMenu + 102
				End If

				If InStr(1, ",-1,0,5,", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) > 0 Then
					Response.Write "<A"
						If aHeaderComponent(L_LINKED_OPTION_HEADER) And BUDGET_TOOLBAR Then
							Response.Write " HREF=""Main_ISSSTE.asp?SectionID=5"""
						End If
					Response.Write ">"
						Response.Write "<FONT COLOR=""#"
							If aHeaderComponent(L_SELECTED_OPTION_HEADER) And BUDGET_TOOLBAR Then
								Response.Write S_SELECTED_LINK_FOR_GUI
							Else
								Response.Write S_MENU_LINK_FOR_GUI
							End If
						Response.Write """ CLASS=""SpecialLink"" STYLE=""font-size: 13px;"">PRESUPUESTO</FONT>"
					Response.Write "</A>&nbsp;&#183;&nbsp;"
					iLeftForPopupMenu = iLeftForPopupMenu + 109
				End If

				If InStr(1, ",-1,0,6,", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) > 0 Then
					Response.Write "<A"
						If aHeaderComponent(L_LINKED_OPTION_HEADER) And REPORTS_TOOLBAR Then
							Response.Write " HREF=""Main_ISSSTE.asp?SectionID=6"""
						End If
					Response.Write ">"
						Response.Write "<FONT COLOR=""#"
							If aHeaderComponent(L_SELECTED_OPTION_HEADER) And REPORTS_TOOLBAR Then
								Response.Write S_SELECTED_LINK_FOR_GUI
							Else
								Response.Write S_MENU_LINK_FOR_GUI
							End If
						Response.Write """ CLASS=""SpecialLink"" STYLE=""font-size: 13px;"">DEPARTAMENTO&nbsp;TÉCNICO</FONT>"
					Response.Write "</A>&nbsp;&#183;&nbsp;"
					iLeftForPopupMenu = iLeftForPopupMenu + 180
				End If

				If InStr(1, ",-1,0,7,", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) > 0 Then
					Response.Write "<A"
						If aHeaderComponent(L_LINKED_OPTION_HEADER) And LOGOUT_TOOLBAR Then
							Response.Write " HREF=""Main_ISSSTE.asp?SectionID=7"""
						End If
					Response.Write ">"
						Response.Write "<FONT COLOR=""#"
							If aHeaderComponent(L_SELECTED_OPTION_HEADER) And LOGOUT_TOOLBAR Then
								Response.Write S_SELECTED_LINK_FOR_GUI
							Else
								Response.Write S_MENU_LINK_FOR_GUI
							End If
						Response.Write """ CLASS=""SpecialLink"" STYLE=""font-size: 13px;"">DESCONCENTRADOS</FONT>"
					Response.Write "</A>&nbsp;&#183;&nbsp;"
					iLeftForPopupMenu = iLeftForPopupMenu + 146
				End If

				If InStr(1, ",8,", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) > 0 Then
					Response.Write "<A"
						If aHeaderComponent(L_LINKED_OPTION_HEADER) And LOGOUT_TOOLBAR Then
							Response.Write " HREF=""Main_ISSSTE.asp?SectionID=8"""
						End If
					Response.Write ">"
						Response.Write "<FONT COLOR=""#"
							If aHeaderComponent(L_SELECTED_OPTION_HEADER) And LOGOUT_TOOLBAR Then
								Response.Write S_SELECTED_LINK_FOR_GUI
							Else
								Response.Write S_MENU_LINK_FOR_GUI
							End If
						Response.Write """ CLASS=""SpecialLink"" STYLE=""font-size: 13px;"">ATENCIÓN&nbsp;AL&nbsp;PERSONAL</FONT>"
					Response.Write "</A>&nbsp;&#183;&nbsp;"
					iLeftForPopupMenu = iLeftForPopupMenu + 146
				End If
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_BUDGET_PERMISSIONS Then
					Response.Write "<A"
						If aHeaderComponent(L_LINKED_OPTION_HEADER) And BUDGET_TOOLBAR Then
							Response.Write " HREF=""Budget.asp"""
						End If
					Response.Write ">"
						Response.Write "<FONT COLOR=""#"
							If aHeaderComponent(L_SELECTED_OPTION_HEADER) And BUDGET_TOOLBAR Then
								Response.Write S_SELECTED_LINK_FOR_GUI
							Else
								Response.Write S_MENU_LINK_FOR_GUI
							End If
						Response.Write """ CLASS=""SpecialLink"" STYLE=""font-size: 13px;"">PRESUPUESTOS</FONT>"
					Response.Write "</A>&nbsp;&#183;&nbsp;"
					iLeftForPopupMenu = iLeftForPopupMenu + 118
				End If

				If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_AREAS_PERMISSIONS) = N_AREAS_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_POSITIONS_PERMISSIONS) = N_POSITIONS_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_JOBS_PERMISSIONS) = N_JOBS_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_EMPLOYEES_PERMISSIONS) = N_EMPLOYEES_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_SADE_PERMISSIONS) = N_SADE_PERMISSIONS) Then
					Dim aHRMenuComponent()
					Call InitializeMenuComponent(aHRMenuComponent)
					aHRMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Centros de trabajo",_
							  "",_
							  "", "Areas.asp", N_AREAS_PERMISSIONS),_
						Array("Puestos",_
							  "",_
							  "", "Positions.asp", N_POSITIONS_PERMISSIONS),_
						Array("Plazas",_
							  "",_
							  "", "Jobs.asp", N_JOBS_PERMISSIONS),_
						Array("Empleados",_
							  "",_
							  "", "Employees.asp", N_EMPLOYEES_PERMISSIONS),_
						Array("Desarrollo Humano",_
							  "",_
							  "", "SADE.asp", N_SADE_PERMISSIONS)_
					)
					aHRMenuComponent(N_LEFT_FOR_DIV_MENU) = iLeftForPopupMenu
					aHRMenuComponent(N_TOP_FOR_DIV_MENU) = 64
					aHRMenuComponent(N_WIDTH_FOR_DIV_MENU) = 156
					aHRMenuComponent(B_USE_DIV_MENU) = True
					Response.Write "<A"
						If aHeaderComponent(L_LINKED_OPTION_HEADER) And HUMAN_RESOURCES_TOOLBAR Then
							Response.Write " HREF=""HumanResources.asp"""
						End If
						Response.Write " onMouseOver=""ShowPopupItem('" & aHRMenuComponent(S_POPUP_DIV_NAME_MENU) & "'), document." & aHRMenuComponent(S_POPUP_DIV_NAME_MENU) & """ onMouseOut=""HidePopupItem('" & aHRMenuComponent(S_POPUP_DIV_NAME_MENU) & "'), document." & aHRMenuComponent(S_POPUP_DIV_NAME_MENU) & """"
					Response.Write ">"
						Response.Write "<FONT COLOR=""#"
							If aHeaderComponent(L_SELECTED_OPTION_HEADER) And HUMAN_RESOURCES_TOOLBAR Then
								Response.Write S_SELECTED_LINK_FOR_GUI
							Else
								Response.Write S_MENU_LINK_FOR_GUI
							End If
						Response.Write """ CLASS=""SpecialLink"" STYLE=""font-size: 13px;"">PERSONAL</FONT>"
					Response.Write "</A>&nbsp;&#183;&nbsp;"
					iLeftForPopupMenu = iLeftForPopupMenu + 83
				End If

				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_PAYROLL_PERMISSIONS Then
					Response.Write "<A"
						If aHeaderComponent(L_LINKED_OPTION_HEADER) And PAYROLL_TOOLBAR Then
							Response.Write " HREF=""Payroll.asp"""
						End If
					Response.Write ">"
						Response.Write "<FONT COLOR=""#"
							If aHeaderComponent(L_SELECTED_OPTION_HEADER) And PAYROLL_TOOLBAR Then
								Response.Write S_SELECTED_LINK_FOR_GUI
							Else
								Response.Write S_MENU_LINK_FOR_GUI
							End If
						Response.Write """ CLASS=""SpecialLink"" STYLE=""font-size: 13px;"">NÓMINA</FONT>"
					Response.Write "</A>&nbsp;&#183;&nbsp;"
					iLeftForPopupMenu = iLeftForPopupMenu + 64
				End If

				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_PAYMENTS_PERMISSIONS Then
					Response.Write "<A"
						If aHeaderComponent(L_LINKED_OPTION_HEADER) And PAYMENTS_TOOLBAR Then
							Response.Write " HREF=""Payments.asp"""
						End If
					Response.Write ">"
						Response.Write "<FONT COLOR=""#"
							If aHeaderComponent(L_SELECTED_OPTION_HEADER) And PAYMENTS_TOOLBAR Then
								Response.Write S_SELECTED_LINK_FOR_GUI
							Else
								Response.Write S_MENU_LINK_FOR_GUI
							End If
						Response.Write """ CLASS=""SpecialLink"" STYLE=""font-size: 13px;"">CHEQUES</FONT>"
					Response.Write "</A>&nbsp;&#183;&nbsp;"
					iLeftForPopupMenu = iLeftForPopupMenu + 74
				End If

				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REPORTS_PERMISSIONS Then
					Response.Write "<A"
						If aHeaderComponent(L_LINKED_OPTION_HEADER) And REPORTS_TOOLBAR Then
							Response.Write " HREF=""Reports.asp"""
						End If
					Response.Write ">"
						Response.Write "<FONT COLOR=""#"
							If aHeaderComponent(L_SELECTED_OPTION_HEADER) And REPORTS_TOOLBAR Then
								Response.Write S_SELECTED_LINK_FOR_GUI
							Else
								Response.Write S_MENU_LINK_FOR_GUI
							End If
						Response.Write """ CLASS=""SpecialLink"" STYLE=""font-size: 13px;"">REPORTES</FONT>"
					Response.Write "</A>&nbsp;&#183;&nbsp;"
					iLeftForPopupMenu = iLeftForPopupMenu + 82
				End If
			End If

			If aLoginComponent(B_VALID_USER_LOGIN) Then
				Dim aToolsMenuComponent()
				Call InitializeMenuComponent(aToolsMenuComponent)
				aToolsMenuComponent(A_ELEMENTS_MENU) = Array(_
					Array("Bitácora de errores",_
						  "",_
						  "", "ErrorLog.asp", N_TOOLS_PERMISSIONS),_
					Array("Cambiar contraseña",_
						  "",_
						  "", "ChangePassword.asp", True),_
					Array("Cambiar e-mail adicional",_
						  "",_
						  "", "ChangePassword.asp", False),_
					Array("Opciones del sistema",_
						  "",_
						  "", "Options.asp?Admin=1", N_TOOLS_PERMISSIONS),_
					Array("Preferencias del usuario",_
						  "",_
						  "", "Options.asp", True),_
					Array("<LINE />", "", "", "", N_TACO_PERMISSIONS),_
					Array("Tablero de control",_
						  "",_
						  "", "TaCo.asp", N_TACO_PERMISSIONS),_
					Array("<LINE />", "", "", "", True),_
					Array("Curso de capacitación",_
						  "",_
						  "", "javascript: OpenNewWindow('Course\/index.htm', '', 'Curso', 961, 643, 'yes', 'yes');", False),_
					Array("Manuales de uso",_
						  "",_
						  "", "Docs.asp", True),_
					Array("Soporte técnico",_
						  "",_
						  "", "http://200.57.131.24/SoS/AddEvent.asp?SystemID=9&ModuleID=" & Request.Cookies("SoS_SectionID") & "&UserEmail=" & aLoginComponent(S_USER_E_MAIL_LOGIN) & """ TARGET=""SoS_AddEvent", aLoginComponent(B_TECH_SUPPORT_LOGIN))_
				)
				aToolsMenuComponent(N_LEFT_FOR_DIV_MENU) = iLeftForPopupMenu
				aToolsMenuComponent(N_TOP_FOR_DIV_MENU) = 64
				aToolsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 156
				aToolsMenuComponent(B_USE_DIV_MENU) = True
				Response.Write "<A"
					If aHeaderComponent(L_LINKED_OPTION_HEADER) And TOOLS_TOOLBAR Then
						Response.Write " HREF=""Tools.asp"""
					End If
					Response.Write " onMouseOver=""ShowPopupItem('" & aToolsMenuComponent(S_POPUP_DIV_NAME_MENU) & "'), document." & aToolsMenuComponent(S_POPUP_DIV_NAME_MENU) & """ onMouseOut=""HidePopupItem('" & aToolsMenuComponent(S_POPUP_DIV_NAME_MENU) & "'), document." & aToolsMenuComponent(S_POPUP_DIV_NAME_MENU) & """"
				Response.Write ">"
					Response.Write "<FONT COLOR=""#"
						If aHeaderComponent(L_SELECTED_OPTION_HEADER) And TOOLS_TOOLBAR Then
							Response.Write S_SELECTED_LINK_FOR_GUI
						Else
							Response.Write S_MENU_LINK_FOR_GUI
						End If
					Response.Write """ CLASS=""SpecialLink"" STYLE=""font-size: 13px;"">HERRAMIENTAS</FONT>"
				Response.Write "</A>&nbsp;&#183;&nbsp;"
				iLeftForPopupMenu = iLeftForPopupMenu + 51
			End If

			Response.Write "<A HREF=""DocsLibrary.asp"">"
				Response.Write "<FONT COLOR=""#"
					If aHeaderComponent(L_SELECTED_OPTION_HEADER) And DOCS_TOOLBAR Then
						Response.Write S_SELECTED_LINK_FOR_GUI
					Else
						Response.Write S_MENU_LINK_FOR_GUI
					End If
				Response.Write """ CLASS=""SpecialLink"" STYLE=""font-size: 13px;"">NORMATECA</FONT>"
			Response.Write "</A>&nbsp;&#183;&nbsp;"
			iLeftForPopupMenu = iLeftForPopupMenu + 93

			Response.Write "<A HREF=""Docs.asp"">"
				Response.Write "<FONT COLOR=""#"
					If StrComp(GetASPFileName(""), "Docs.asp", vbBinaryCompare) = 0 Then
						Response.Write S_SELECTED_LINK_FOR_GUI
					Else
						Response.Write S_MENU_LINK_FOR_GUI
					End If
				Response.Write """ CLASS=""SpecialLink"" STYLE=""font-size: 13px;"">MANUALES</FONT>"
			Response.Write "</A>"
			iLeftForPopupMenu = iLeftForPopupMenu + 0
		End If
	%></B></FONT></TD></TR>
</TABLE>
<%Response.Write "<TABLE WIDTH=""993"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0""><TR>"
	Response.Write "<TD VALIGN=""TOP"">&nbsp;&nbsp;<FONT FACE=""Century Gothic, Arial"" SIZE="""
		If Len(aHeaderComponent(S_TITLE_NAME_HEADER)) > 50 Then
			Response.Write "4"
			sBoldBeginForHeader = "<B>"
			sBoldEndForHeader = "</B>"
		Else
			Response.Write "5"
			sBoldBeginForHeader = ""
			sBoldEndForHeader = ""
		End If
	Response.Write """ COLOR=""#" & S_MAIN_TITLE_FOR_GUI & """>" & sBoldBeginForHeader & aHeaderComponent(S_TITLE_NAME_HEADER) & sBoldEndForHeader & "</FONT></TD>"
	If Not IsEmpty(aOptionsMenuComponent) Then
		If Not IsEmpty(aOptionsMenuComponent(A_ELEMENTS_MENU)) Then
			Response.Write "<TD ALIGN=""RIGHT"" VALIGN=""TOP"">"
				Call DisplayMenuPopup(True, aOptionsMenuComponent)
			Response.Write "</TD>"
		End If
	End If
Response.Write "</TR></TABLE>"
If Not IsEmpty(aHRMenuComponent) Then
	If Not IsEmpty(aHRMenuComponent(A_ELEMENTS_MENU)) Then
		Call DisplayMenuPopup(False, aHRMenuComponent)
	End If
End If
If Not IsEmpty(aToolsMenuComponent) Then
	If Not IsEmpty(aToolsMenuComponent(A_ELEMENTS_MENU)) Then
		Call DisplayMenuPopup(False, aToolsMenuComponent)
	End If
End If%>
<!-- END: HEADER -->
<TABLE WIDTH="100%" BORDER="0" CELLPADDING="0" CELLSPACING="0"><TR>
	<TD>&nbsp;&nbsp;</TD>
	<TD WIDTH="100%" VALIGN="TOP"><FONT FACE="Arial" SIZE="2"><BR />