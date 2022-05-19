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
<!-- #include file="Libraries/CatalogComponent.asp" -->
<!-- #include file="Libraries/PositionComponent.asp" -->
<!-- #include file="Libraries/CatalogsLib.asp" -->
<!-- #include file="Libraries/PositionsLib.asp" -->
<!-- #include file="Libraries/JobComponent.asp" -->
<%
Dim sNames
Dim bShowForm
Dim bAction
Dim bError
Dim sCondition
Dim sActionErrorDescription
Dim sFilter
Dim bFilter

sFilter = ""

If B_ISSSTE Then
Else
	If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_POSITIONS_PERMISSIONS) = N_POSITIONS_PERMISSIONS Then
	Else
		Response.Redirect "AccessDenied.asp?Permission=" & N_POSITIONS_PERMISSIONS
	End If
End If

aCatalogComponent(S_TABLE_NAME_CATALOG) = "Positions"
Call InitializeCatalogs(oRequest)
Call InitializeValuesForCatalogComponent(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
Call InitializeJobComponent(oRequest, aJobComponent)
Call InitializePositionComponent(oRequest, aPositionComponent)
Call GetPositionsURLValues(oRequest, bShowForm, bAction, aCatalogComponent(S_QUERY_CONDITION_CATALOG))
aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) = CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)))
If Not B_ISSSTE Then aCatalogComponent(S_URL_PARAMETERS_CATALOG) = "Positions.asp?PositionID=<FIELD_0 />"
aJobComponent(B_SEND_TO_IFRAME_JOB) = True
bShowForm = True

If Len(oRequest("ApplyFilter").Item) > 0 Then
    Call GetStartAndEndDatesFromURL("StartForValue", "EndForValue", "Positions.StartDate", False, aPositionComponent(S_FILTER_CONDITION_POSITION))
	If (Len(oRequest("PositionShortNameFilter").Item) > 0) And (CInt(oRequest("PositionShortNameFilter").Item) > 0) Then
		aPositionComponent(S_FILTER_CONDITION_POSITION) = aPositionComponent(S_FILTER_CONDITION_POSITION) & " And (Positions.PositionShortName = '" & UCase(CStr(oRequest("PositionShortNameFilter").Item)) & "')"
	End If
	If (Len(oRequest("GroupGradeLevelIDFilter").Item) > 0) And (CInt(oRequest("GroupGradeLevelIDFilter").Item) > 0) Then
		aPositionComponent(S_FILTER_CONDITION_POSITION) = aPositionComponent(S_FILTER_CONDITION_POSITION) & " And (Positions.GroupGradeLevelID=" & CInt(oRequest("GroupGradeLevelIDFilter").Item) & ")"
	End If
	If (Len(oRequest("EmployeeTypeIDFilter").Item) > 0) And (CInt(oRequest("EmployeeTypeIDFilter").Item) > 0) Then
		aPositionComponent(S_FILTER_CONDITION_POSITION) = aPositionComponent(S_FILTER_CONDITION_POSITION) & " And (Positions.EmployeeTypeID=" & CInt(oRequest("EmployeeTypeIDFilter").Item) & ")"
	End If
	aPositionComponent(S_QUERY_CONDITION_POSITION) = aPositionComponent(S_QUERY_CONDITION_POSITION) & aPositionComponent(S_FILTER_CONDITION_POSITION)
	sFilter = "ApplyFilter=1&PositionShortNameFilter=" & CStr(oRequest("PositionShortNameFilter").Item) & "&GroupGradeLevelIDFilter=" & CStr(oRequest("GroupGradeLevelIDFilter").Item) & "&EmployeeTypeIDFilter=" & CStr(oRequest("EmployeeTypeIDFilter").Item)
End If

bError = False
If bAction Then
	aCatalogComponent(S_QUERY_CONDITION_CATALOG) = ""
	lErrorNumber = DoPositionsAction(oRequest, oADODBConnection, oRequest("Action").Item, sErrorDescription)
	sActionErrorDescription = sErrorDescription
	bError = (lErrorNumber = 0)
	If bError Then
		Response.Redirect "Positions.asp?Success=0"
	Else
		Response.Redirect "Positions.asp?Success=1&ErrorDescription=" & sErrorDescription
	End If
End If
If lErrorNumber = 0 Then
	If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) > -1 Then
		lErrorNumber = GetCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
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
aHeaderComponent(S_TITLE_NAME_HEADER) = "Puestos"
bWaitMessage = True
Response.Cookies("SoS_SectionID") = 191
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
		<SCRIPT LANGUAGE="JavaScript" SRC="JavaScript/Export.js"></SCRIPT>
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" onload="enableFields()">
		<%If Len(oRequest("ReadOnly").Item) = 0 Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Agregar un nuevo puesto",_
					  "",_
					  "", "Positions.asp?PositionID=-1&New=1", N_ADD_PERMISSIONS),_
				Array("<LINE />",_
					  "",_
					  "", "", ((Len(oRequest("New").Item) = 0) And ((aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) = -1) Or (Len(oRequest("PositionID").Item) > 0)))),_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Positions&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) = -1)),_
				Array("Exportar registro a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Positions&Excel=1&PositionID=" & oRequest("PositionID").Item & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')",(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) > -1)),_
				Array("Exportar registro a Word",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Positions&Word=1&PositionID=" & oRequest("PositionID").Item & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')",(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) > -1)),_
				Array("Exportar a Excel los registros mostrados",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Positions&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") &"&StartPage="& oRequest("StartPage") &"&StartForValueDay="& oRequest("StartForValueDay") &"&StartForValueMonth="& oRequest("StartForValueMonth") &"&StartForValueYear="& oRequest("StartForValueYear") &"&EndForValueDay="& oRequest("EndForValueDay") &"&EndForValueMonth="& oRequest("EndForValueMonth") &"&EndForValueYear="& oRequest("EndForValueYear") &"&PositionShortName=" & oRequest("PositionShortName") & "&GroupGradeLevelID=" & oRequest("GroupGradeLevelID") & "&ApplyFilter=++Filtrar++ " & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')",(Len(oRequest("ApplyFilter").Item)>0))_
				)
		Else
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Positions&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) = -1)),_
				Array("Exportar a Word",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Positions&Word=1&PositionID=" & oRequest("PositionID").Item & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) > -1))_
			)
		End If
		aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 793
		aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
		aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 200
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
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=31"">Estructuras ocupacionales</A> > <A HREF=""Catalogs.asp"">Catálogos</A> > "
				Case 4
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Catalogs.asp?CatalogType=4"">Catálogos</A> > "
				Case Else
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=31"">Estructuras ocupacionales</A> > <A HREF=""Catalogs.asp"">Catálogos</A> > "
			End Select
		Else
			Response.Write "<A HREF=""HumanResources.asp"">Recursos Humanos</A> > "
		End If
		If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) = -1 Then
			Response.Write "<B>Puestos</B>"
		Else
			Call GetNameFromTable(oADODBConnection, "Positions", aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG))&","&oRequest("StartDate"), "", "", sNames, sErrorDescription)
			Response.Write "<A HREF=""Positions.asp"">Puestos</A> > <B>" & sNames & "</B>"
		End If
		Response.Write "<BR /><BR />"
		If Len(oRequest("Search").Item) > 0 Then
			Call DisplayPositionsSearchForm(oRequest, oADODBConnection, True, sErrorDescription)
		ElseIf (Len(oRequest("DoSearch").Item) > 0) Or (Len(oRequest("SetActive").Item) > 0) Then
			lErrorNumber = DisplayCatalogsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, (Len(oRequest("ReadOnly").Item) = 0), aCatalogComponent, sErrorDescription)
			If lErrorNumber = L_ERR_NO_RECORDS Then
				Call DisplayPositionsSearchForm(oRequest, oADODBConnection, True, sErrorDescription)
			End If
		ElseIf ((Len(oRequest("Add").Item) > 0) And (lErrorNumber <> 0)) Then
			If (Len(oRequest("Add").Item) > 0) And (lErrorNumber <> 0) Then
				Response.Write "<BR /><BR />"
				sErrorDescription = sActionErrorDescription
				Call DisplayErrorMessage("Error en la información del puesto", sErrorDescription)
				lErrorNumber = 0
				Response.Write "<BR />"
			End If
			Select Case oRequest("Action").Item
				Case "Jobs"
					Call DisplayJobForm(oRequest, oADODBConnection, GetASPFileName(""), aJobComponent, sErrorDescription)
				Case Else
					Call DisplayPositionForm(oRequest, oADODBConnection, GetASPFileName(""), aPositionComponent, sErrorDescription)
			End Select
		ElseIf bShowForm Then
			If bAction Then
				Response.Write "<BR />"
				If lErrorNumber = 0 Then
					If Len(oRequest("Remove").Item) > 0 Then
						Call DisplayErrorMessage("Confirmación", "La información del puesto se ha borrado correctamente.")
					Else
						Call DisplayErrorMessage("Confirmación", "La información del puesto fue guardada con éxito.")
					End If
				Else
					Call DisplayErrorMessage("Mensaje del sistema", sActionErrorDescription)
					lErrorNumber = 0
				End If
				Response.Write "<BR />"
			ElseIf Len(oRequest("Success").Item) > 0 Then
				If CInt(oRequest("Success").Item) = 1 Then
					Call DisplayErrorMessage("Error", CStr(oRequest("ErrorDescription").Item))
				Else
					Call DisplayErrorMessage("Confirmación", "La operación fué realizada de manera exitosa.")
				End If
			End If
			Select Case oRequest("Action").Item
				Case "Jobs"
					If Len(oRequest("ShowInfo").Item) > 0 Then
						lErrorNumber = DisplayJob(oRequest, oADODBConnection, False, aJobComponent, sErrorDescription)
					Else
						lErrorNumber = DisplayJobForm(oRequest, oADODBConnection, GetASPFileName(""), aJobComponent, sErrorDescription)
					End If
				Case Else
					Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
						Response.Write "<TD WIDTH=""600"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
						lErrorNumber = DisplayPisitionsFilters(oRequest, GetASPFileName(""), sErrorDescription)
						Response.Write "<DIV STYLE=""height: 600px; width: 900px; overflow: auto;"">"
						If lErrorNumber = 0 Then
							If Len(oRequest("ApplyFilter").Item) > 0 Then
								Call GetStartAndEndDatesFromURL("StartForValue", "EndForValue", "Positions.StartDate", False, sCondition)
								aPositionComponent(S_QUERY_CONDITION_POSITION) = aPositionComponent(S_QUERY_CONDITION_POSITION) & sCondition
								If CInt(oRequest("PositionShortName").Item) > 0 Then
									aPositionComponent(S_QUERY_CONDITION_POSITION) = aPositionComponent(S_QUERY_CONDITION_POSITION) & " And (Positions.PositionShortName Like '" & S_WILD_CHAR & oRequest("PositionShortName").Item & S_WILD_CHAR & "')"
								End If
								If CInt(oRequest("GroupGradeLevelID").Item) > 0 Then
									aPositionComponent(S_QUERY_CONDITION_POSITION) = aPositionComponent(S_QUERY_CONDITION_POSITION) & " And (Positions.GroupGradeLevelID = " & oRequest("GroupGradeLevelID").Item & ")"
								End If
							End If
							lErrorNumber = DisplayPositionsTable(oRequest, oADODBConnection, False, aPositionComponent, sErrorDescription)
						End If
						Response.Write "</DIV>"
						If lErrorNumber <> 0 Then
							Response.Write "<BR />"
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							sErrorDescription = ""
							bShowForm = True
						End If

						Response.Write "<TD>&nbsp;</TD>"
						Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
						Response.Write "<TD>&nbsp;</TD>"
						Response.Write "<TD WIDTH=""*"" VALIGN=""TOP"">"
							Response.Write "<DIV NAME=""CatalogDiv"" ID=""CatalogDiv"">"
								lErrorNumber = DisplayPositionForm(oRequest, oADODBConnection, GetASPFileName(""), aPositionComponent, sErrorDescription)
							Response.Write "</DIV>"
							If lErrorNumber <> 0 Then
								Response.Write "<BR />"
								Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
								lErrorNumber = 0
								sErrorDescription = ""
							End If
						Response.Write "</TD>"

					Response.Write "</TR></TABLE>"
			End Select
			If lErrorNumber <> 0 Then
				Response.Write "<BR /><BR />"
				Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
				lErrorNumber = 0
				Response.Write "<BR />"
			End If
		Else
			If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) = -1 Then
				Response.Write "<FORM NAME=""ConceptInfoFrm"" ID=""ConceptInfoFrm"" ACTION=""Positions.asp"" METHOD=""POST"">"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartPage"" ID=""StartPageHdn"" VALUE=""1"" />"
					Response.Write "<B>Seleccione los datos para filtrar los registros:&nbsp;&nbsp;&nbsp;</B><BR />"
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar los registros del &nbsp;"
					Response.Write DisplayDateCombos(oRequest("StartForValueYear").Item, oRequest("StartForValueMonth").Item, oRequest("StartForValueDay").Item, "StartForValueYear", "StartForValueMonth", "StartForValueDay", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "&nbsp;al&nbsp;"
					Response.Write DisplayDateCombos(oRequest("EndForValueYear").Item, oRequest("EndForValueMonth").Item, oRequest("EndForValueDay").Item, "EndForValueYear", "EndForValueMonth", "EndForValueDay", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "<BR />"

					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar los registros con clave:&nbsp;"
						Response.Write "<SELECT NAME=""PositionShortName"" ID=""PositionShortNameCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write "<OPTION VALUE=""-2"">Todos</OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "Distinct PositionShortName", "PositionShortName As PositionShortName2", "PositionShortName IS NOT NULL AND PositionShortName <> ' '", "PositionShortName", aPositionComponent(S_SHORT_NAME_POSITION), "", sErrorDescription)
						Response.Write "</SELECT>"
					Response.Write "<BR />"

					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar los registros con Grupo, grado, nivel:&nbsp;"
						Response.Write "<SELECT NAME=""GroupGradeLevelID"" ID=""GroupGradeLevelIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write "<OPTION VALUE=""-2"">Todos</OPTION>"
							Response.Write "<OPTION VALUE=""-1"">Ninguno</OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "GroupGradeLevels", "GroupGradeLevelID", "GroupGradeLevelName", "(GroupGradeLevelID>-1) And (Active=1)", "GroupGradeLevelName", aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11), "", sErrorDescription)
						Response.Write "</SELECT>"
					Response.Write "<BR />"
					Response.Write "<INPUT TYPE=""SUBMIT"" VALUE=""Ver Reporte"" CLASS=""Buttons""><BR />"
					Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""960"" HEIGHT=""1"" /><BR />"
				Response.Write "</FORM>"
				Call GetStartAndEndDatesFromURL("StartForValue", "EndForValue", "Positions.StartDate", False, sCondition)
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & sCondition
				If CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11)) <> -2 Then aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (Positions.GroupGradeLevelID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11) & ")"
				lErrorNumber = DisplayCatalogsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, (Len(oRequest("ReadOnly").Item) = 0) And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)), aCatalogComponent, sErrorDescription)
				If lErrorNumber <> 0 Then
					Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
					lErrorNumber = 0
					sErrorDescription = ""
				End If
			Else
				Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
					Response.Write "<TD WIDTH=""1"" VALIGN=""TOP"">"
						lErrorNumber = DisplayPositionCompact(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
					Response.Write "</TD>"
					Response.Write "<TD>&nbsp;</TD>"
					Response.Write "<TD BGCOLOR=""#" & S_MAIN_COLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
					Response.Write "<TD>&nbsp;&nbsp;&nbsp;</TD>"

					Response.Write "<TD WIDTH=""*"" VALIGN=""TOP"">"
						aJobComponent(S_QUERY_CONDITION_JOB) = " And (Jobs.PositionID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & ")"
						aJobComponent(N_SHOW_BY_JOB) = N_SHOW_BY_POSITION
						Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>PLAZAS</B><BR /><BR /></FONT>"
						Response.Write "<DIV ID=""JobsDiv"" CLASS=""JobsTable"">"
							lErrorNumber = DisplayJobsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)), False, aJobComponent, sErrorDescription)
						Response.Write "</DIV><BR />"
						Response.Write "<IFRAME SRC=""ShowForms.asp"" NAME=""FormsIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""270""></IFRAME>"
					Response.Write "</TD>"
				Response.Write "</TR></TABLE>"
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