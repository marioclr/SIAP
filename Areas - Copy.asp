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
<!-- #include file="Libraries/AreasLib.asp" -->
<!-- #include file="Libraries/AreaComponent.asp" -->
<!-- #include file="Libraries/JobComponent.asp" -->
<!-- #include file="Libraries/ZoneComponent.asp" -->
<%
Dim iSelectedTab
Dim bAction
Dim bError
Dim bCompactStyle
Dim bPaymentCenters

If B_ISSSTE Then
Else
	If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_AREAS_PERMISSIONS) = N_AREAS_PERMISSIONS Then
	Else
		Response.Redirect "AccessDenied.asp?Permission=" & N_AREAS_PERMISSIONS
	End If
End If

bCompactStyle = True
bPaymentCenters = (Len(oRequest("PaymentCenters").Item) > 0)
Call InitializeAreaComponent(oRequest, aAreaComponent)
Call InitializeJobComponent(oRequest, aJobComponent)
Call GetAreasURLValues(oRequest, iSelectedTab, bAction, aAreaComponent(S_QUERY_CONDITION_AREA))
If B_ISSSTE And (iGlobalSectionID = 3) Then iSelectedTab = 1
aAreaComponent(B_SEND_TO_IFRAME_AREA) = True
aJobComponent(B_SEND_TO_IFRAME_JOB) = True

bError = False
If bAction Then
	lErrorNumber = DoAreasAction(oRequest, oADODBConnection, oRequest("Action").Item, sErrorDescription)
	bError = (lErrorNumber <> 0)
End If
If lErrorNumber = 0 Then
	If aAreaComponent(N_ID_AREA) > -1 Then
		lErrorNumber = GetArea(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
'		If (lErrorNumber = 0) And bAction And (Len(oRequest("Add").Item) = 0) Then
'			aAreaComponent(N_ID_AREA) = aAreaComponent(N_PARENT_ID_AREA)
'			lErrorNumber = GetArea(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
'		End If
	End If
	If aJobComponent(N_ID_JOB) > -1 Then
		lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
	End If
End If

Select Case iGlobalSectionID
	Case 1
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
	Case 2
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 3
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 4
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
	Case Else
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
End Select
If bPaymentCenters Then
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Centros de pago"
Else
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Áreas y centros de trabajo"
End If
bWaitMessage = True
Response.Cookies("SoS_SectionID") = 190
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/Export.js"></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%If bPaymentCenters Then
			If Len(oRequest("ReadOnly").Item) = 0 Then
				aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
					Array("Agregar un nuevo centro de pago",_
						  "",_
						  "", "Areas.asp?ParentID=" & aAreaComponent(N_ID_AREA) & "&AreaLevelTypeID=2&Tab=1&New=1", (N_ADD_PERMISSIONS)),_
					Array("<LINE />",_
						  "",_
						  "", "", ((Len(oRequest("New").Item) = 0) And ((aAreaComponent(N_ID_AREA) = -1) Or (Len(oRequest("PositionID").Item) > 0)))),_
					Array("Exportar a Excel",_
						  "",_
						  "", "javascript: OpenNewWindow('Export.asp?Action=Areas&Excel=1&PaymentCenters=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
				)
			Else
				aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
					Array("Exportar a Excel",_
						  "",_
						  "", "javascript: OpenNewWindow('Export.asp?Action=Areas&Excel=1&PaymentCenters=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
				)
			End If
		Else
			If Len(oRequest("ReadOnly").Item) = 0 Then
				aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
					Array("Agregar una nueva área",_
						  "",_
						  "", "Areas.asp?ParentID=-1&AreaLevelTypeID=1&Tab=1&New=1", (N_ADD_PERMISSIONS And (aAreaComponent(N_LEVEL_TYPE_ID_AREA) = -1))),_
					Array("Agregar un nuevo centro de trabajo",_
						  "",_
						  "", "Areas.asp?ParentID=" & aAreaComponent(N_ID_AREA) & "&AreaLevelTypeID=2&Tab=1&New=1", (N_ADD_PERMISSIONS And (aAreaComponent(N_LEVEL_TYPE_ID_AREA) = 1))),_
					Array("<LINE />",_
						  "",_
						  "", "", ((Len(oRequest("New").Item) = 0) And ((aAreaComponent(N_ID_AREA) = -1) Or (Len(oRequest("PositionID").Item) > 0)))),_
					Array("Exportar a Excel",_
						  "",_
						  "", "javascript: OpenNewWindow('Export.asp?Action=Areas&Excel=1&ParentID=" & aAreaComponent(N_ID_AREA) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (aAreaComponent(N_LEVEL_TYPE_ID_AREA) < 2)),_
					Array("Exportar a Excel",_
						  "",_
						  "", "javascript: OpenNewWindow('Export.asp?Action=Jobs&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", ((aAreaComponent(N_LEVEL_TYPE_ID_AREA) = 2) And (CInt(Request.Cookies("SIAP_SectionID")) <> 3))),_
					Array("Exportar información del área a Word",_
						  "",_
						  "", "javascript: OpenNewWindow('Export.asp?Action=Areas&Word=1&AreaID=" & aAreaComponent(N_ID_AREA) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", False),_
					Array("Exportar información del centro a Word",_
						  "",_
						  "", "javascript: OpenNewWindow('Export.asp?Action=Areas&Word=1&AreaID=" & aAreaComponent(N_ID_AREA) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", ((aAreaComponent(N_ID_AREA) > -1) And (aAreaComponent(N_LEVEL_TYPE_ID_AREA) >= 2))),_
					Array("Imprimir",_
						  "",_
						  "", "javascript: SendReportToPrint('ReportDiv', '" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "')", False)_
				)
			Else
				aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
					Array("Exportar a Excel",_
						  "",_
						  "", "javascript: OpenNewWindow('Export.asp?Action=Areas&Excel=1&ParentID=" & aAreaComponent(N_ID_AREA) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (aAreaComponent(N_LEVEL_TYPE_ID_AREA) < 2)),_
					Array("Exportar a Excel",_
						  "",_
						  "", "javascript: OpenNewWindow('Export.asp?Action=Jobs&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", ((aAreaComponent(N_LEVEL_TYPE_ID_AREA) = 2) And (CInt(Request.Cookies("SIAP_SectionID")) <> 3))),_
					Array("Exportar información del área a Word",_
						  "",_
						  "", "javascript: OpenNewWindow('Export.asp?Action=Areas&Word=1&AreaID=" & aAreaComponent(N_ID_AREA) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", False),_
					Array("Exportar información del centro a Word",_
						  "",_
						  "", "javascript: OpenNewWindow('Export.asp?Action=Areas&Word=1&AreaID=" & aAreaComponent(N_ID_AREA) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", ((aAreaComponent(N_ID_AREA) > -1) And (aAreaComponent(N_LEVEL_TYPE_ID_AREA) >= 2))),_
					Array("Imprimir",_
						  "",_
						  "", "javascript: SendReportToPrint('ReportDiv', '" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "')", False)_
				)
			End If
		End If
		aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 733
		aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
		aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 260
		%>
		<!-- #include file="_Header.asp" -->
		<%Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > "
		If B_ISSSTE Then
			Select Case iGlobalSectionID
				Case 1
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Catalogs.asp?CatalogType=1"">Catálogos</A> > "
				Case 2
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Catalogs.asp?CatalogType=2"">Catálogos</A> > "
				Case 3
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=31"">Estructuras ocupacionales</A> > <A HREF=""Catalogs.asp?CatalogType=3"">Catálogos</A> > "
				Case 4
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Catalogs.asp?CatalogType=4"">Catálogos</A> > "
				Case Else
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=31"">Estructuras ocupacionales</A> > <A HREF=""Catalogs.asp?CatalogType=3"">Catálogos</A> > "
			End Select
		Else
			Response.Write "<A HREF=""HumanResources.asp"">Personal</A> > "
		End If
		If bPaymentCenters Then
			Response.Write "<B>Centros de pago</B>"
		ElseIf aAreaComponent(N_ID_AREA) = -1 Then
			If (aAreaComponent(N_ID_AREA) = -1) And (aAreaComponent(N_PARENT_ID_AREA) = -1) Then
				Response.Write "<B>Áreas y centros de trabajo</B>"
			Else
				Dim aTempAreaComponent()
				Redim aTempAreaComponent(N_AREA_COMPONENT_SIZE)
				aTempAreaComponent(N_ID_AREA) = aAreaComponent(N_PARENT_ID_AREA)
				Response.Write "<A HREF=""Areas.asp?ReadOnly=" & oRequest("ReadOnly").Item & """>Áreas y centros de trabajo</A> > "
				Call DisplayAreaPath(oRequest, oADODBConnection, aTempAreaComponent, "")
			End If
		Else
			Response.Write "<A HREF=""Areas.asp?ReadOnly=" & oRequest("ReadOnly").Item & """>Áreas y centros de trabajo</A> > "
			Call DisplayAreaPath(oRequest, oADODBConnection, aAreaComponent, "")
		End If
		Response.Write "<BR /><BR />"
		If Len(oRequest("Search").Item) > 0 Then
			Call DisplayAreasSearchForm(oRequest, oADODBConnection, True, sErrorDescription)
		ElseIf Len(oRequest("DoSearch").Item) > 0 Then
			aAreaComponent(B_SEND_TO_IFRAME_AREA) = False
			lErrorNumber = DisplayAreasTable(oRequest, oADODBConnection, DISPLAY_NOTHING, (Len(oRequest("ReadOnly").Item) = 0), False, False, aAreaComponent, sErrorDescription)
			If lErrorNumber = L_ERR_NO_RECORDS Then
				Call DisplayAreasSearchForm(oRequest, oADODBConnection, True, sErrorDescription)
			End If
		ElseIf (Len(oRequest("Add").Item) > 0) And (lErrorNumber <> 0) Then
			Response.Write "<BR /><BR />"
			Call DisplayErrorMessage("Error en la información del centro de trabajo", sErrorDescription)
			lErrorNumber = 0
			Response.Write "<BR />"
			Select Case oRequest("Action").Item
				Case "Jobs"
					Call DisplayJobForm(oRequest, oADODBConnection, GetASPFileName(""), aJobComponent, sErrorDescription)
				Case Else
					Call DisplayAreaForm(oRequest, oADODBConnection, GetASPFileName(""), aAreaComponent, sErrorDescription)
			End Select
		ElseIf (Len(oRequest("Tab").Item) > 0) Or (bAction And (Len(oRequest("Remove").Item) = 0) And (Len(oRequest("SetActive").Item) = 0)) Then
			If bAction Then
				Response.Write "<BR />"
				If lErrorNumber = 0 Then
					Call DisplayErrorMessage("Confirmación", "La información del centro de trabajo fue guardada con éxito.")
					lErrorNumber = 0
				Else
					Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
					lErrorNumber = 0
				End If
				Response.Write "<BR />"
			End If
			Select Case oRequest("Action").Item
				Case "Jobs"
					If Len(oRequest("ShowInfo").Item) > 0 Then
						lErrorNumber = DisplayJob(oRequest, oADODBConnection, False, aJobComponent, sErrorDescription)
					Else
						lErrorNumber = DisplayJobForm(oRequest, oADODBConnection, GetASPFileName(""), aJobComponent, sErrorDescription)
					End If
				Case Else
					If B_ISSSTE And (iGlobalSectionID = 3) Then
					ElseIf (Not B_ISSSTE) Or (iGlobalSectionID <> 3) Or (aAreaComponent(N_LEVEL_TYPE_ID_AREA) > 1) Then
						Call DisplayAreasTabs(oRequest, bError, sErrorDescription)
					End If
					Response.Write "<BR />"
					Select Case iSelectedTab
						Case 2
							lErrorNumber = DisplayAreaPositionsForm(oRequest, oADODBConnection, GetASPFileName(""), aAreaComponent, sErrorDescription)
						Case 3
							lErrorNumber = DisplayAreaHistoryList(oRequest, oADODBConnection, False, aAreaComponent, sErrorDescription)
						Case 4
							lErrorNumber = DisplayAreaPositionsHistoryList(oRequest, oADODBConnection, False, aAreaComponent, sErrorDescription)
						Case Else
							If Len(oRequest("ShowInfo").Item) > 0 Then
								lErrorNumber = DisplayArea(oRequest, oADODBConnection, False, aAreaComponent, sErrorDescription)
							Else
								lErrorNumber = DisplayAreaForm(oRequest, oADODBConnection, GetASPFileName(""), aAreaComponent, sErrorDescription)
							End If
					End Select
			End Select
			If lErrorNumber <> 0 Then
				Response.Write "<BR /><BR />"
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
				Response.Write "<BR />"
			End If
		Else
			If aAreaComponent(N_ID_AREA) = -1 Then


				Response.Write "<FORM NAME=""ConceptInfoFrm"" ID=""ConceptInfoFrm"" ACTION=""Areas.asp"" METHOD=""POST"">"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartPage"" ID=""StartPageHdn"" VALUE=""1"" />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaFind"" ID=""AreaFindHdn"" VALUE=""1"" />"
					Response.Write "<B>Seleccione los datos para filtrar los registros:&nbsp;&nbsp;&nbsp;</B><BR />"

					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar los registros con clave:&nbsp;"
						Response.Write "<SELECT NAME=""AreaShortName"" ID=""AreaShortNameCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write "<OPTION VALUE=""-2"">Todos</OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "Distinct AreaShortName", "AreaShortName As AreaShortName2", "AreaShortName IS NOT NULL AND AreaShortName <> ' ' AND ParentID >= 0", "AreaShortName", aAreaComponent(S_SHORT_NAME_AREA), "", sErrorDescription)
						Response.Write "</SELECT>"
					Response.Write "<BR />"

					Response.Write "<INPUT TYPE=""SUBMIT"" VALUE=""Ver Reporte"" CLASS=""Buttons""><BR />"
					Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""960"" HEIGHT=""1"" /><BR />"
				Response.Write "</FORM>"



				If bPaymentCenters Then
					aAreaComponent(S_QUERY_CONDITION_AREA) = " And (Areas.PaymentCenterID>-1)"
				Else
					If Len(aAreaComponent(S_SHORT_NAME_AREA)) > 0 Then
						'aAreaComponent(S_QUERY_CONDITION_AREA) = aAreaComponent(S_QUERY_CONDITION_AREA) & " And (Areas.AreaShortName='" & aAreaComponent(S_SHORT_NAME_AREA) & "')"
					Else
						aAreaComponent(S_QUERY_CONDITION_AREA) = " And (Areas.ParentID=-1)"
					End If
				End If
				aAreaComponent(B_SEND_TO_IFRAME_AREA) = False
				lErrorNumber = DisplayAreasTable(oRequest, oADODBConnection, DISPLAY_NOTHING, (Len(oRequest("ReadOnly").Item) = 0) And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)), False, aAreaComponent, sErrorDescription)
				If lErrorNumber <> 0 Then
					Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
					lErrorNumber = 0
					sErrorDescription = ""
				End If
			Else
				If (Not B_ISSSTE) And (aAreaComponent(N_LEVEL_TYPE_ID_AREA) > 1) Then
					Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
						Response.Write "<TD WIDTH=""1"" VALIGN=""TOP"">"
							If bCompactStyle Then
								lErrorNumber = DisplayAreaCompact(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
								If lErrorNumber = 0 Then
									aAreaComponent(S_QUERY_CONDITION_AREA) = " And (Areas.ParentID=" & aAreaComponent(N_ID_AREA) & ")"
									lErrorNumber = DisplayAreasInSmallIcons(oRequest, oADODBConnection, False, aAreaComponent, sErrorDescription)
								End If
							Else
								lErrorNumber = DisplayArea(oRequest, oADODBConnection, False, aAreaComponent, sErrorDescription)
							End If
						Response.Write "</TD>"
						Response.Write "<TD>&nbsp;&nbsp;&nbsp;</TD>"

						Response.Write "<TD WIDTH=""*"" VALIGN=""TOP"">"
							If Not bCompactStyle Then


								Response.Write "<FORM NAME=""ConceptInfoFrm"" ID=""ConceptInfoFrm"" ACTION=""Areas.asp"" METHOD=""POST"">"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartPage"" ID=""StartPageHdn"" VALUE=""1"" />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaFind"" ID=""AreaFindHdn"" VALUE=""1"" />"
									Response.Write "<B>Seleccione los datos para filtrar los registros:&nbsp;&nbsp;&nbsp;</B><BR />"

									Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar los registros con clave:&nbsp;"
										Response.Write "<SELECT NAME=""AreaShortName"" ID=""AreaShortNameCmb"" SIZE=""1"" CLASS=""Lists"">"
											Response.Write "<OPTION VALUE=""-2"">Todos</OPTION>"
											Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "Distinct AreaShortName", "AreaShortName As AreaShortName2", "AreaShortName IS NOT NULL AND AreaShortName <> ' ' AND ParentID >= 0", "AreaShortName", aAreaComponent(S_SHORT_NAME_AREA), "", sErrorDescription)
										Response.Write "</SELECT>"
									Response.Write "<BR />"

									Response.Write "<INPUT TYPE=""SUBMIT"" VALUE=""Ver Reporte"" CLASS=""Buttons""><BR />"
									Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""960"" HEIGHT=""1"" /><BR />"
								Response.Write "</FORM>"

								aAreaComponent(S_QUERY_CONDITION_AREA) = " And (Areas.ParentID=" & aAreaComponent(N_ID_AREA) & ")"
								'If Len(aAreaComponent(S_SHORT_NAME_AREA)) > 0 Then aAreaComponent(S_QUERY_CONDITION_AREA) = aAreaComponent(S_QUERY_CONDITION_AREA) & " And (Areas.AreaShortName='" & aAreaComponent(S_SHORT_NAME_AREA) & "')"

								lErrorNumber = DisplayAreasTable(oRequest, oADODBConnection, DISPLAY_NOTHING, (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)), False, aAreaComponent, sErrorDescription)
								Response.Write "<BR />"
							End If
							If (lErrorNumber <> 0) And (lErrorNumber <> L_ERR_NO_RECORDS) Then
								Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
								lErrorNumber = 0
								sErrorDescription = ""
							End If
							If aAreaComponent(N_LEVEL_TYPE_ID_AREA) = 2 Then
								aJobComponent(S_QUERY_CONDITION_JOB) = " And (Jobs.AreaID=" & aAreaComponent(N_ID_AREA) & ")"
								aJobComponent(N_SHOW_BY_JOB) = N_SHOW_BY_AREA
								Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>PLAZAS</B><BR /><BR /></FONT>"
								Response.Write "<DIV ID=""JobsDiv"" CLASS=""JobsTable"">"
									lErrorNumber = DisplayJobsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)), False, aJobComponent, sErrorDescription)
								Response.Write "</DIV>"
								Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""10"" /><BR />"
								Response.Write "<IMG SRC=""Images/DotBlue.gif"" WIDTH=""750"" HEIGHT=""1"" ALIGN=""ABSMIDDLE"" /><BR />"
								Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""10"" /><BR />"
								Response.Write "<IFRAME SRC=""ShowForms.asp"" NAME=""FormsIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""270""></IFRAME>"
							Else
								Response.Write "<IFRAME SRC=""ShowForms.asp"" NAME=""FormsIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""700""></IFRAME>"
							End If
						Response.Write "</TD>"
					Response.Write "</TR></TABLE>"
				Else


					Response.Write "<FORM NAME=""ConceptInfoFrm"" ID=""ConceptInfoFrm"" ACTION=""Areas.asp"" METHOD=""POST"">"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartPage"" ID=""StartPageHdn"" VALUE=""1"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaFind"" ID=""AreaFindHdn"" VALUE=""1"" />"
						Response.Write "<B>Seleccione los datos para filtrar los registros:&nbsp;&nbsp;&nbsp;</B><BR />"

						Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar los registros con clave:&nbsp;"
							Response.Write "<SELECT NAME=""AreaShortName"" ID=""AreaShortNameCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE=""-2"">Todos</OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "Distinct AreaShortName", "AreaShortName As AreaShortName2", "AreaShortName IS NOT NULL AND AreaShortName <> ' ' AND ParentID >= 0", "AreaShortName", aAreaComponent(S_SHORT_NAME_AREA), "", sErrorDescription)
							Response.Write "</SELECT>"
						Response.Write "<BR />"

						Response.Write "<INPUT TYPE=""SUBMIT"" VALUE=""Ver Reporte"" CLASS=""Buttons""><BR />"
						Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""960"" HEIGHT=""1"" /><BR />"
					Response.Write "</FORM>"


					'aAreaComponent(S_QUERY_CONDITION_AREA) = " And (Areas.ParentID=" & aAreaComponent(N_ID_AREA) & ")"
					aAreaComponent(B_SEND_TO_IFRAME_AREA) = False
					lErrorNumber = DisplayAreasTable(oRequest, oADODBConnection, DISPLAY_NOTHING, (Len(oRequest("ReadOnly").Item) = 0) And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)), False, aAreaComponent, sErrorDescription)
					If lErrorNumber <> 0 Then
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						sErrorDescription = ""
					End If





				End If
			End If
		End If
		If lErrorNumber <> 0 Then
			Response.Write "<BR /><BR />"
			Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			lErrorNumber = 0
			Response.Write "<BR />"
		End If%>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>