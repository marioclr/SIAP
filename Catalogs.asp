<%@LANGUAGE=VBSCRIPT%>
<%
Option Explicit
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Server.ScriptTimeout = 72000
%>
<!-- #include file="Libraries/GlobalVariables.asp" -->
<!-- #include file="Libraries/LoginComponent.asp" -->
<!-- #include file="Libraries/CatalogsLib.asp" -->
<!-- #include file="Libraries/CatalogComponent.asp" -->
<!-- #include file="Libraries/ConceptComponent.asp" -->
<!-- #include file="Libraries/TaxLibraries.asp" -->
<!-- #include file="Libraries/UploadInfoLibrary.asp" -->
<%
Dim sAction
Dim iStep
Dim sError
Dim sFileName
Dim sNames
Dim lSuccess
Dim bDoAction
Dim bShowForm
Dim sCatalogTypes
Dim sRequiredFields

If B_ISSSTE Then
Else
	If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And (N_CATALOGS_PERMISSIONS) Then
	Else
		Response.Redirect "AccessDenied.asp?Permission=" & N_CATALOGS_PERMISSIONS
	End If
End If

sAction = ""
If Not IsEmpty(oRequest("Action")) Then
	sAction = oRequest("Action").Item
End If
Call InitializeCatalogs(oRequest)
Call InitializeConceptComponent(oRequest, aConceptComponent)
Call InitializePayrollResumeForSarComponent(oRequest, aPayrollResumeForSarComponent)
Call InitializeBanamexCensusComponent(oRequest, aBanamexCensusComponent)
bDoAction = ((Len(oRequest("Add").Item) > 0) Or (Len(oRequest("Apply").Item) > 0) Or (Len(oRequest("Modify").Item) > 0) Or (Len(oRequest("SetActive").Item) > 0) Or (Len(oRequest("Remove").Item) > 0) Or (Len(oRequest("Active").Item) > 0) Or (Len(oRequest("Deactive").Item) > 0) Or (Len(oRequest("Unlock").Item) > 0)) Or ((aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) And (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) > -1))
sCatalogTypes = "," & iGlobalSectionID & ","

If Len(oRequest("Success").Item) > 0 Then lSuccess = 1
iStep = 1
If Len(oRequest("Step").Item) > 0 Then iStep = CInt(oRequest("Step").Item)
sFileName = Server.MapPath(UPLOADED_PHYSICAL_PATH & sAction & "_" & aLoginComponent(N_USER_ID_LOGIN) & ".txt")

If Len(oRequest("RawData").Item) > 0 Then
	lErrorNumber = SaveTextToFile(sFileName, oRequest("RawData").Item, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Redirect "Catalogs.asp?Action=" & sAction & "&Step=" & iStep
	End If
End If

If Len(oRequest("ApplyFilter").Item) > 0 Then
	Call GetFilterFromURL(oRequest, sAction, sErrorDescription)
End If
If Len(oRequest("PositionsSpecialJourneysAction").Item) > 0 Then
	If (Len(oRequest("AuthorizationFile").Item) > 0) Then
		lErrorNumber = AddPositionsSpecialJourneysFile(oRequest, oADODBConnection,  oRequest("sQuery").Item, aConceptComponent, sErrorDescription)
		sError = sErrorDescription
		If lErrorNumber = 0 Then
			sError = sError & "El puesto para guardias y suplencias se registró exitosamente<BR />"
		Else
			sError = sError & "Error al registrar el puesto para guardias y suplencias<BR />"
		End If
		Response.Redirect "Catalogs.asp?Action=PositionsSpecialJourneysLKP&Success=1&ErrorDescription=" & sError
	End If
End If
If bDoAction Then
	Call InitializeValuesForCatalogComponent(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
	lErrorNumber = DoAction(sAction, bShowForm, sErrorDescription)
	If lErrorNumber = 0 Then
		Select Case sAction
			Case "ProfessionalRiskMatrix"
				If ((Len(oRequest("Modify").Item) <> 0) And (Len(oRequest("Delete").Item) <> 0)) Then Response.Redirect "Catalogs.asp?Action=ProfessionalRiskMatrix"
			Case "PositionsSpecialJourneysLKP"
				Response.Redirect "Catalogs.asp?Action=PositionsSpecialJourneysLKP&Success=1"
			Case "PayrollsClcs"
				Response.Redirect "Reports.asp?ReportID=1400&ReportStep=2&PayrollID=" & oRequest("PayrollID").Item & "&PayrollCode=" & oRequest("PayrollCode").Item
			Case Else
		End Select
	Else
		Select Case sAction
			Case "PositionsSpecialJourneysLKP"
				Response.Redirect "Catalogs.asp?Action=PositionsSpecialJourneysLKP&Success=0&sErrorDescription=" & sErrorDescription
			Case Else
		End Select
	End If
End If

If B_ISSSTE Then
	If Len(aCatalogComponent(S_NAME_CATALOG)) = 0 Then
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Catálogos"
	Else
		aHeaderComponent(S_TITLE_NAME_HEADER) = aCatalogComponent(S_NAME_CATALOG)
	End If
	Select Case iGlobalSectionID
		Case 1
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
			Response.Cookies("SoS_SectionID") = 1110
		Case 2
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
			Response.Cookies("SoS_SectionID") = 1028
		Case 3
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
			Response.Cookies("SoS_SectionID") = 1311
		Case 4
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
			Response.Cookies("SoS_SectionID") = 1410
		Case 5
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = BUDGET_TOOLBAR
			Response.Cookies("SoS_SectionID") = 1510
	End Select
Else
	aHeaderComponent(S_TITLE_NAME_HEADER) = "Administración de Catálogos"
	aHeaderComponent(L_SELECTED_OPTION_HEADER) = TOOLS_TOOLBAR
End If
bWaitMessage = (Len(sAction) > 0)
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0" onLoad="DocumentOnLoad()">
		<%If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS) = N_ADD_PERMISSIONS) And (Len(oRequest("ReadOnly").Item) = 0) Then
			Select Case sAction
				Case "BanamexCensus"
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Agregar un nuevo registro",_
						  "",_
							"", "Catalogs.asp?Action=BanamexCensus&AddNew=1", True),_
						Array("Exportar a Excel",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (InStr(1, ",PaperworkLists,", "," & sAction & ",", vbBinaryCompare) = 0))_
					)
				Case "ConsarFile"
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Exportar a Excel",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (InStr(1, ",PaperworkLists,", "," & sAction & ",", vbBinaryCompare) = 0))_
					)
				Case "CurrenciesHistoryList"
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Exportar a Excel",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
					)
				Case "EmployeeFields"
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Agregar campo",_
							  "",_
							  "", GetASPFileName("") & "?Action=" & sAction & "&New=1&FormFieldID=-1", True)_
					)
				Case "EmployeesDeleted"
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Exportar a Excel",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (InStr(1, ",PaperworkLists,", "," & sAction & ",", vbBinaryCompare) = 0))_
					)
				Case "Forms"
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Agregar formulario",_
							  "",_
							  "", GetASPFileName("") & "?Action=" & sAction & "&New=1&FormID=-1", True)_
					)
				Case "FormFields"
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Agregar campo",_
							  "",_
							  "", GetASPFileName("") & "?Action=" & sAction & "&New=1&FormID=" & oRequest("FormID").Item & "&FormFieldID=-1", True)_
					)
				Case "PayrollCompare"
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Exportar a Excel",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (InStr(1, ",PaperworkLists,", "," & sAction & ",", vbBinaryCompare) = 0))_
					)
				Case "PayrollResume"
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Exportar a Excel",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (InStr(1, ",PaperworkLists,", "," & sAction & ",", vbBinaryCompare) = 0))_
					)
				Case "PositionsSpecialJourneysLKP"
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Agregar registro",_
							  "",_
							  "", GetASPFileName("") & "?Action=" & sAction & "&New=1&StartPage=" & oRequest("StartPage").Item & "&" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=-1&" & aCatalogComponent(S_URL_CATALOG), True),_
						Array("Exportar a Excel",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (InStr(1, ",PaperworkLists,", "," & sAction & ",", vbBinaryCompare) = 0)),_
						Array("Exportar histórico a Excel",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Excel=1&ShowAll=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (InStr(1, ",PaperworkLists,", "," & sAction & ",", vbBinaryCompare) = 0))_
					)
				Case "Profiles"
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Agregar perfil",_
							  "",_
							  "", GetASPFileName("") & "?Action=" & sAction & "&New=1&ProfileID=-2", True)_
					)
				Case "Users"
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Agregar usuario",_
							  "",_
							  "", GetASPFileName("") & "?Action=" & sAction & "&New=1&UserID=-2", (Not B_PORTAL))_
					)
				Case "Zones"
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Agregar zona",_
							  "",_
							  "", GetASPFileName("") & "?Action=" & sAction & "&New=1&ZoneID=-1&ParentID=" & oRequest("ParentID").Item, True),_
						Array("Exportar a Excel",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
					)
				Case Else
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Agregar registro",_
							  "",_
							  "", GetASPFileName("") & "?Action=" & sAction & "&New=1&StartPage=" & oRequest("StartPage").Item & "&" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=-1&" & aCatalogComponent(S_URL_CATALOG), True),_
						Array("Exportar a Excel",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (InStr(1, ",PaperworkLists,", "," & sAction & ",", vbBinaryCompare) = 0))_
					)
			End Select
		Else
			Select Case sAction
				Case "Zones"
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Exportar a Excel",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
					)
				Case Else
					aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Exportar a Excel",_
							  "",_
							  "", "javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", (InStr(1, ",PaperworkLists,", "," & sAction & ",", vbBinaryCompare) = 0))_
					)
			End Select
		End If
		aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 810
		aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
		aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 183
		%>
		<!-- #include file="_Header.asp" -->
		<!-- BEGIN: PATH -->
		<%Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > "
		If B_ISSSTE Then
			Select Case iGlobalSectionID
				Case 1
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > "
				Case 2
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > "
				Case 3
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo Humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=31"">Estructuras Ocupacionales</A> > "
				Case 4
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > "
				Case 5
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > "
				Case 6
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > "
                Case 8
                    Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=8"">Atención al personal</A> > "
			End Select
			sNames = "Catálogos"
		Else
			Response.Write "<A HREF=""Tools.asp"">Herramientas</A> > "
			sNames = "Catálogos"
		End If
			If Len(sAction) = 0 Then
				Response.Write "<B>" & sNames & "</B>"
			Else
				If InStr(1, ",PaperworkActions,PaperworkAddresses,PaperworkLists,PaperworkOwners,PaperworkSenders,PaperworkTypes,SubjectTypes,", sAction, vbBinaryCompare) > 0 Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=61"">Ventanilla única</A> > "
				ElseIf StrComp(sAction, "ProfessionalRiskMatrix", vbBinaryCompare) = 0 Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=29"">Matriz de Riesgos Profesionales</A> > "
				ElseIf StrComp(sAction, "PayrollCompare", vbBinaryCompare) = 0 Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=491"">Ejercicio bimestral del SAR</A> > <B> Mostrar comparativo de devengos </B>"
				ElseIf StrComp(sAction, "PayrollResume", vbBinaryCompare) = 0 Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=491"">Ejercicio bimestral del SAR</A> > <B> Mostrar resumen de nóminas </B>"
				ElseIf StrComp(sAction, "BanamexCensus", vbBinaryCompare) = 0 Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=491"">Ejercicio bimestral del SAR</A> > <B> Mostrar padrón Banamex </B>"
				ElseIf StrComp(sAction, "ConsarFile", vbBinaryCompare) = 0 Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=491"">Ejercicio bimestral del SAR</A> > <B> Mostrar archivo CONSAR </B>"
				ElseIf StrComp(sAction, "EmployeesDeleted", vbBinaryCompare) = 0 Then
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=491"">Ejercicio bimestral del SAR</A> > <B> Consultar histórico de bajas de empleados </B>"
				Else
					Response.Write "<A HREF=""Catalogs.asp"">" & sNames & "</A> > "
				End If
				sNames = ""
				Select Case sAction
					'Case "CenterSubtypes"
					'	Call GetNameFromTable(oADODBConnection, "CenterTypes", oRequest("CenterTypeID").Item, "", "", sNames, sErrorDescription)
					'	Response.Write "<A HREF=""Catalogs.asp?Action=CenterTypes"">Tipo de centro de trabajo: " & CleanStringForHTML(sNames) & "</A> > <B>" & aCatalogComponent(S_NAME_CATALOG) & "</B><BR /><BR />"
					Case "EmployeeFields"
						Response.Write "<B>Campos para la información de los empleados</B><BR /><BR />"
					Case "EmploymentAllowances"
						Response.Write "<B>Tabla de subsidio mensual y quincenal</B><BR /><BR />"
					Case "Forms"
						Response.Write "<B>Formularios</B><BR /><BR />"
					Case "FormFields"
						Call GetNameFromTable(oADODBConnection, "Forms", oRequest("FormID").Item, "", "", sNames, "")
						Response.Write "<A HREF=""Catalogs.asp?Action=Forms"">Formularios. " & sNames & "</A> > <B>Campos</B><BR /><BR />"
					Case "Profiles"
						Response.Write "<B>Perfiles</B><BR /><BR />"
Case "SubStates"
Call GetNameFromTable(oADODBConnection, "States", oRequest("ParentID").Item, "", "", sNames, "")
Response.Write "<A HREF=""Catalogs.asp?Action=States"">Entidad Federativa: " & sNames & "</A> > <B>" & aCatalogComponent(S_NAME_CATALOG) & "</B><BR /><BR />"
					Case "TaxInvertions"
						Response.Write "<B>Tabla de ISR</B><BR /><BR />"
					Case "TaxLimits"
						Response.Write "<B>Tabla de límites</B><BR /><BR />"
					Case "Users"
						Response.Write "<B>Usuarios del sistema</B><BR /><BR />"
					Case "Zones"
						If Len(oRequest("ParentID").Item) = 0 Then aZoneComponent(N_PARENT_ID_ZONE) = aLoginComponent(L_USER_PERMISSION_ZONE_ID_LOGIN)
						If Len(oRequest("ParentID").Item) = 0 Then
							Response.Write "<B>Entidades federativas</B><BR /><BR />"
						Else
							Response.Write "<A HREF=""Catalogs.asp?Action=Zones"">Entidades federativas</A> > "
								Call DisplayZonePath(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
							Response.Write "<BR /><BR />"
						End If
					Case Else
						Response.Write "<B>" & aCatalogComponent(S_NAME_CATALOG) & "</B><BR /><BR />"
				End Select
			End If
		'<!-- END: PATH -->
		If lErrorNumber <> 0 Then
			Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			lErrorNumber = 0
			sErrorDescription = ""
			bShowForm = (Len(oRequest("Add")) > 0)
		End If
		Response.Write "<BR />"
		If iStep <= 1 Then
			If Len(sAction) = 0 Then
				'<!-- BEGIN: MENU -->
				aMenuComponent(A_ELEMENTS_MENU) = Array(_
					Array("Antigüedades",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Antiquities", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Ausencias",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Absences", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0))),_
					Array("Horarios",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Shifts", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Periodicidad",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Periods", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Tipos de ámbito para las áreas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=ConfineTypes", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Tipos de área",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=AreaTypes", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Dependencias gubernamentales",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=FederalCompanies", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0))),_
					Array("Tipos de crédito",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=CreditTypes", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0))),_
					Array("Tipos de tabulador",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=EmployeeTypes", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Requisitos de documentación para movimiento de personal",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=EmployeesRequirements", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0))),_
					Array("Tipos de movimiento",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Reasons", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0))),_
					Array("Riesgos profesonales",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=RiskLevels", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Clasificación de tipos de movimiento",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=ReasonTypes", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0))),_
					Array("Tipos de niveles para las área",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=AreaLevelTypes", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",-1,", vbBinaryCompare) > 0))),_
					Array("Tipos de ocupación",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=OccupationTypes", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0))),_
					Array("Tipos de puesto",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=PositionTypes", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Turnos",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Journeys", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("<LINE />", "", "", "", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Bancos",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Banks", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Cuentas bancarias del Instituto",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=BankAccounts", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Monedas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Currencies", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Tabla de ISR",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=TaxInvertions", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Tabla de límites",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=TaxLimits", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Tabla de subsidio mensual y quincenal",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=EmploymentAllowances", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Actividad institucional",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=BudgetsActivities1", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",5,", vbBinaryCompare) > 0))),_
					Array("Actividad presupuestaria",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=BudgetsActivities2", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",5,", vbBinaryCompare) > 0))),_
					Array("Ámbito",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=BudgetsConfineTypes", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",5,", vbBinaryCompare) > 0))),_
					Array("Fondo",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=BudgetsFunds", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",5,", vbBinaryCompare) > 0))),_
					Array("Función",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=BudgetsDuties", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",5,", vbBinaryCompare) > 0))),_
					Array("Municipio",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=BudgetsLocations", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",5,", vbBinaryCompare) > 0))),_
					Array("Proceso",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=BudgetsProcesses", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",5,", vbBinaryCompare) > 0))),_
					Array("Programa",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=BudgetsPrograms", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",5,", vbBinaryCompare) > 0))),_
					Array("Programa presupuestario",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=BudgetsProgramDuties", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",5,", vbBinaryCompare) > 0))),_
					Array("Región",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=BudgetsRegions", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",5,", vbBinaryCompare) > 0))),_
					Array("Subfunción activa",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=BudgetsActiveDuties", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",5,", vbBinaryCompare) > 0))),_
					Array("Subfunción específica",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=BudgetsSpecificDuties", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",5,", vbBinaryCompare) > 0))),_
					Array("<LINE />", "", "", "", True),_
					Array("Platilla modificada",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=PositionsAreasLKP", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",5,", vbBinaryCompare) > 0))),_
					Array("<LINE />", "", "", "", True),_
					Array("Tipos de estructuras programáticas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=BudgetTypes2", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",5,", vbBinaryCompare) > 0))),_
					Array("Tipos de partidas presupuestales",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=BudgetTypes", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",5,", vbBinaryCompare) > 0))),_
					Array("<LINE />", "", "", "", True),_
					Array("CLCs",_
						  "",_
						  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1400", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Resumen mensual de nóminas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1403", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("<LINE />", "", "", "", True),_
					Array("Estatus",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Status", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",-1,", vbBinaryCompare) > 0))),_
					Array("Estatus de las áreas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=StatusAreas", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",-1,", vbBinaryCompare) > 0))),_
					Array("Estatus de las partidas presupuestales",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=StatusBudgets", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",-1,", vbBinaryCompare) > 0))),_
					Array("Estatus de los empleados",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=StatusEmployees", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0))),_
					Array("Estatus de los formularios",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=StatusForms", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Estatus de las plazas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=StatusJobs", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0))),_
					Array("Estatus de los niveles",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=StatusLevels", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",-1,", vbBinaryCompare) > 0))),_
					Array("Estatus de los pagos",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=StatusPayments", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Estatus de los puestos",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=StatusPositions", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",-1,", vbBinaryCompare) > 0))),_
					Array("<LINE />", "", "", "", True),_
					Array("Áreas y centros de trabajo",_
						  "",_
						  "Images/MnLeftArrows.gif", "Areas.asp", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Áreas generadoras",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=GeneratingAreas", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Áreas geográficas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=ZoneTypes", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Centros de pago",_
						  "",_
						  "Images/MnLeftArrows.gif", "Areas.asp?PaymentCenters=1", False),_
					Array("Centros de trabajo",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=WorkingCenters", False),_
					Array("Discapacidades",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Handicaps", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0))),_
					Array("Entidades federativas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Zones", (False And (StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Grupos, grados, niveles",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=GroupGradeLevels", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Matriz UNIMED",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=MedicalAreas", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Niveles",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Levels", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Niveles de atención",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=AttentionLevels", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Pagadurías SIPE",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=CashierOffices", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Países",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Countries", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",-1,", vbBinaryCompare) > 0))),_
					Array("Puestos",_
						  "",_
						  "Images/MnLeftArrows.gif", "Positions.asp", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Puestos genéricos",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=GenericPositions", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Entidades federativas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=States", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Puestos para guardias y suplencias",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=PositionsSpecialJourneysLKP", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Puestos por jerarquía",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=PositionsHierarchy", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Ramas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Branches", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Perfil académico",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Requirements", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0))),_
					Array("Servicios",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Services", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Servicios por tipo de centro de trabajo",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=ServicesCenterTypesLKP", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Sindicatos",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Syndicates", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Sociedades y empresas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Companies", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Subramas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=SubBranches", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Subtipos de centro de trabajo",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=CenterSubtypes", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Tipos de reporte UNIMED",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=MedicalAreasTypes", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Tipos de centro de trabajo",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=CenterTypes", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("Zonas económicas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=EconomicZones", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",3,", vbBinaryCompare) > 0))),_
					Array("<LINE />", "", "", "", True),_
					Array("Campos para la información de los empleados",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=EmployeeFields", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",444,", vbBinaryCompare) > 0))),_
					Array("Formularios",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Forms", False),_
					Array("Días de asueto",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Holidays", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0))),_
					Array("<LINE />", "", "", "", ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Perfiles",_
						  "Administre los permisos de los usuarios a través de perfiles.",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Profiles", ((StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_UsuariosDelSistema & ",", vbBinaryCompare) > 0) And ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",44,", vbBinaryCompare) > 0)))),_
					Array("Usuarios del sistema",_
						  "Personas autorizadas para entrar al sistema y acceder a los módulos que lo integran.",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Users", ((StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_UsuariosDelSistema & ",", vbBinaryCompare) > 0) And ((StrComp(sCatalogTypes, ",-1,", vbBinaryCompare) = 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0)))),_

					Array("<TITLE />CONSULTAS", "", "", "", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Horarios",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Shifts&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Tipos de ámbito para las áreas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=ConfineTypes&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Tipos de área",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=AreaTypes&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Tipos de tabulador",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=EmployeeTypes&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Tipos de puesto",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=PositionTypes&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Turnos",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Journeys&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Áreas y centros de trabajo",_
						  "",_
						  "Images/MnLeftArrows.gif", "Areas.asp?ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Áreas generadoras",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=GeneratingAreas&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Áreas geográficas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=ZoneTypes&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Centros de pago",_
						  "",_
						  "Images/MnLeftArrows.gif", "Areas.asp?PaymentCenters=1&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Entidades federativas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Zones&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Grupos, grados, niveles",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=GroupGradeLevels&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Matriz UNIMED",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=MedicalAreas&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Niveles",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Levels&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Niveles de atención",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=AttentionLevels&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Pagadurías SIPE",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=CashierOffices&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Puestos",_
						  "",_
						  "Images/MnLeftArrows.gif", "Positions.asp?ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Puestos genéricos",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=GenericPositions&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Puestos por jerarquía",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=PositionsHierarchy&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Ramas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Branches&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Servicios",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Services&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Servicios por tipo de centro de trabajo",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=ServicesCenterTypesLKP&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Sociedades y empresas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=Companies&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Subramas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=SubBranches&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Subtipos de centro de trabajo",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=CenterSubtypes&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Tipos de reporte UNIMED",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=MedicalAreasTypes&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Tipos de centro de trabajo",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=CenterTypes&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("Zonas económicas",_
						  "",_
						  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=EconomicZones&ReadOnly=1", ((InStr(1, sCatalogTypes, ",1,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",2,", vbBinaryCompare) > 0) Or (InStr(1, sCatalogTypes, ",4,", vbBinaryCompare) > 0))),_
					Array("", "", "", "", False)_
				)
				aMenuComponent(B_USE_DIV_MENU) = True
				Response.Write "<BR /><BR /><TABLE WIDTH=""900"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
					Call DisplayMenuInThreeSmallColumns(aMenuComponent)
				Response.Write "</TABLE>"
				'<!-- END: MENU -->
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "function DocumentOnLoad() {" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			Else
				'<!-- BEGIN: CATALOGS -->
				Select Case sAction
					Case "PositionsSpecialJourneysLKP"
						Response.Write "<IMG SRC=""Images/Crcl1.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: ShowDisplay(document.all['CatalogsDiv']); if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; "">Deseo registrar la información en línea</A><BR /><BR />"
						Response.Write "<DIV NAME=""CatalogsDiv"" ID=""CatalogsDiv"">"
					Case "BanamexCensus", "ConsarFile", "EmployeesDeleted"
						Response.Write "<TABLE WIDTH=""720"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
							Response.Write "<TR>"
								If False Then
								Response.Write "<TD WIDTH=""50%"" COLSPAN=""2"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>Consulta de empleados</B></FONT></TD>"
								End If
								If B_ISSSTE Then
									Response.Write "<TD WIDTH=""50%"" COLSPAN=""2"" VALIGN=""TOP""></TD>"
								Else
									Response.Write "<TD WIDTH=""50%"" COLSPAN=""2"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B><A HREF=""Employees.asp?New=1"">ALTA DE EMPLEADOS</A></B></FONT></TD>"
								End If
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD ROWSPAN=""2"">&nbsp;&nbsp;&nbsp;</TD>"
								Response.Write "<TD WIDTH=""50%"" VALIGN=""TOP"">"
									Call DisplayEmployeesSearchForm(oRequest, oADODBConnection, GetASPFileName(""), False, sErrorDescription)
								Response.Write "</TD>"
								Response.Write "<TD ROWSPAN=""2"">&nbsp;&nbsp;&nbsp;</TD>"
								Response.Write "<TD WIDTH=""50%"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
									If B_ISSSTE And False Then
										Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
											Response.Write "function CheckNewEmployeeFields(oForm) {" & vbNewLine
												Response.Write "if (oForm) {" & vbNewLine
													Response.Write "if (oForm.EmployeeNumber.value == '') {" & vbNewLine
														Response.Write "alert('Favor de introducir el número de empleado.');" & vbNewLine
														Response.Write "oForm.EmployeeNumber.focus();" & vbNewLine
														Response.Write "return false;" & vbNewLine
													Response.Write "}" & vbNewLine
												Response.Write "}" & vbNewLine
												Response.Write "return true;" & vbNewLine
											Response.Write "} // End of CheckNewEmployeeFields" & vbNewLine
										Response.Write "//--></SCRIPT>"

										Response.Write "<FORM NAME=""NewEmployeeFrm"" ID=""NewEmployeeFrm"" ACTION=""Employees.asp"" METHOD=""GET"" onSubmit=""return CheckNewEmployeeFields(this)"">"
											Response.Write "Número del empleado: <INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & oRequest("EmployeeNumber").Item & """ CLASS=""TextFields"" /><BR />"
											Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StatusID"" ID=""StatusIDHdn"" VALUE=""-2"" />"
											Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Continuar"" CLASS=""Buttons"" />"
										Response.Write "</FORM>"
									Else
										'Response.Write "<A HREF=""Employees.asp?New=1"">Registre la información del empleado de nuevo ingreso</A> e indique la plaza que va a ocupar.<BR />"
									End If
								Response.Write "</FONT></TD>"
							Response.Write "</TR>"
						Response.Write "</TABLE>"
					Case Else
				End Select
				Response.Write "<TABLE WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
				Select Case sAction
					Case "PositionsSpecialJourneysLKP"
						Response.Write "<TD WIDTH=""30%"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
							lErrorNumber = DisplayFilters(oRequest, sAction, sErrorDescription)
								Response.Write "<DIV NAME=""EntriesDiv"" ID=""EntriesDiv"" CLASS=""TableScrollDiv"">"
									Response.Write "<FORM NAME=""ReportFrm"" ID=""ReportFrm"" ACTION=""Catalogs.asp"" METHOD=""GET"">"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
										Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Mostrar los registros de: <SELECT NAME=""PositionID"" ID=""PositionID"" CLASS=""Lists"">"
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID", "PositionShortName, PositionName, 'Cia:' As Temp, CompanyID, 'Nivel:' As Temp, LevelID, 'Jornada:' As Temp, WorkingHours", "(CompanyID=1) And (EndDate=30000000) And (Active=1)", "PositionShortName", aConceptComponent(N_POSITION_ID_CONCEPT) & "," & aConceptComponent(N_LEVEL_ID_CONCEPT) & "," & aConceptComponent(D_WORKING_HOURS_CONCEPT), "", sErrorDescription)
										Response.Write "</SELECT><BR />"
										Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Servicio:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT NAME=""ServiceID"" ID=""ServiceID"" CLASS=""Lists"">"
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Services", "ServiceID", "ServiceShortName, ServiceName", "(EndDate=30000000) And (Active=1)", "ServiceShortName", aConceptComponent(N_SERVICE_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
										Response.Write "</SELECT><BR />"
										Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""20"" HEIGHT=""30"" ALIGN=""ABSMIDDLE"" />Tipo de centro de trabajo: <SELECT NAME=""CenterTypeID"" ID=""CenterTypeID"" CLASS=""Lists"">"
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "CenterTypes", "CenterTypeID", "CenterTypeShortName, CenterTypeName", "(EndDate=30000000) And (Active=1)", "CenterTypeShortName", aConceptComponent(N_CENTER_TYPE_ID), "Ninguno;;;-1", sErrorDescription)
										Response.Write "</SELECT><BR />"
										Response.Write "&nbsp;&nbsp;&nbsp;<INPUT TYPE=""SUBMIT"" VALUE=""Consultar registros"" CLASS=""Buttons""><BR />"
										Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""800"" HEIGHT=""1"" /><BR />"
									Response.Write "</FORM>"
									If aConceptComponent(N_POSITION_ID_CONCEPT) <> -1 Then aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (PositionsSpecialJourneysLKP.PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ")"
									If aConceptComponent(N_SERVICE_ID_CONCEPT) <> -1 Then aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (PositionsSpecialJourneysLKP.ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ")"
									If aConceptComponent(N_CENTER_TYPE_ID) <> -1 Then aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (PositionsSpecialJourneysLKP.CenterTypeID=" & aConceptComponent(N_CENTER_TYPE_ID) & ")"
									If lErrorNumber = 0 Then
										aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (PositionsSpecialJourneysLKP.Active=1)"
										lErrorNumber = DisplayTables(sAction, sErrorDescription)
									End If
									If lErrorNumber <> 0 Then
										Response.Write "<BR />"
										Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
										lErrorNumber = 0
										sErrorDescription = ""
										bShowForm = True
									End If
								Response.Write "</DIV>"
					Case Else
						Response.Write "<TD WIDTH=""600"" VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
						lErrorNumber = DisplayFilters(oRequest, sAction, sErrorDescription)
						If lErrorNumber = 0 Then
							Response.Write "<DIV STYLE=""height: 400px; width: 600px; overflow: auto;"">"
								lErrorNumber = DisplayTables(sAction, sErrorDescription)
							Response.Write "</DIV>"
						End If
						If lErrorNumber <> 0 Then
							Response.Write "<BR />"
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							sErrorDescription = ""
							bShowForm = True
						End If
				End Select
				Response.Write "</FONT></TD>"
				If (StrComp(sAction, "PayrollResume", vbBinaryCompare) <> 0) And (StrComp(sAction, "BanamexCensus", vbBinaryCompare) <> 0) And _
					(StrComp(sAction, "ConsarFile", vbBinaryCompare) <> 0) And (StrComp(sAction, "PayrollCompare", vbBinaryCompare) <> 0) And _
					(StrComp(sAction, "EmployeesDeleted", vbBinaryCompare) <> 0) Then 
					If Len(oRequest("ReadOnly").Item) = 0 then
						Response.Write "<TD>&nbsp;</TD>"
						Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
						Response.Write "<TD>&nbsp;</TD>"
						Response.Write "<TD WIDTH=""*"" VALIGN=""TOP"">"
							Response.Write "<DIV NAME=""CatalogDiv"" ID=""CatalogDiv"">"
								lErrorNumber = DisplayForms(sAction, sErrorDescription)
							Response.Write "</DIV>"
							If lErrorNumber <> 0 Then
								Response.Write "<BR />"
								Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
								lErrorNumber = 0
								sErrorDescription = ""
							End If
						Response.Write "</TD>"
					End If
					Response.Write "</TR></TABLE>"
				End If
				Select Case sAction
					Case "PositionsSpecialJourneysLKP"
						Response.Write "</DIV>"
						'Response.Write "<BR />"
					Case Else
				End Select
				'If Len(oRequest("Success").Item) > 0 Then
				'	If CInt(oRequest("Success").Item) = 1 Then
				'		Call DisplayErrorMessage("Confirmación", "La operación fué ejecutada exitosamente." & CStr(oRequest("ErrorDescription").Item))
				'	Else
				'		Call DisplayErrorMessage("Error al realizar la operación", CStr(oRequest("ErrorDescription").Item))
				'	End If
				'End If
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "function DocumentOnLoad() {" & vbNewLine
						bShowForm = (bShowForm Or (Len(oRequest("Import").Item) > 0))
						If bShowForm Then
							Response.Write "ShowDisplay(document.all['CatalogDiv']);" & vbNewLine
							Select Case sAction
								Case "Users"
									Response.Write "ShowDisplay(document.all['UserAccessKeyDiv']);" & vbNewLine
									If Len(oRequest("Import").Item) > 0 Then
										Response.Write "HideDisplay(document.UserFrm.Add); ShowDisplay(document.UserFrm.Modify);" & vbNewLine
									End If
								Case Else
							End Select
						End If
					Response.Write "}" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
				'<!-- END: CATALOGS -->
				Select Case sAction
					Case "PositionsSpecialJourneysLKP"
						sRequiredFields = "Clave del Puesto, Nivel, Jornada, Servicio, Tipo de centro de trabajo, Aplica para guardias (1-Si; 0-No), Aplica para suplencias (1-Si; 0-No), Aplica para rezago q. (1-Si; 0-No), Aplica para PROVAC (1-Si; 0-No), Fecha de inicio de vigencia y Fecha de fin de vigencia (opcional)"
						Response.Write "<IMG SRC=""Images/Crcl2.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['CatalogsDiv'] != null) { HideDisplay(document.all['CatalogsDiv']) }; ShowDisplay(document.all['UploadInfoFormDiv']); if(document.all['UploadValidateInfoFormDiv'] != null) { HideDisplay(document.all['UploadValidateInfoFormDiv']) }; ""><FONT FACE=""Arial"" SIZE=""2"">Deseo subir la información a través de un archivo</FONT></A><BR /><BR />"
						Response.Write "<DIV NAME=""UploadInfoFormDiv"" ID=""UploadInfoFormDiv"" STYLE=""display: none"">"
								Response.Write "<FORM NAME=""UploadInfoFrm"" ID=""UploadInfoFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""StepHdn"" VALUE=""2"" />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Success"" ID=""ActionHdn"" VALUE=""" & lSuccess & """ />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & lEmployeeID & """ />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ErrorDescription"" ID=""EmployeeIDHdn"" VALUE=""" & sError & """ />"
									Call DisplayInstructionsMessage("Paso 1. Introduzca el archivo a utilizar", "<BLOCKQUOTE><OL>" & _
																		"<LI>Abra el documento de Excel con la información que desea subir.</LI>" & _
																		"<LI>Copie únicamente las celdas que contienen la información deseada.</LI>" & _
																		"<LI>Pegue dicha información en la caja de texto.</LI>" & _
																		"<LI><B>O seleccione el archivo de texto que contiene la información a subir.</B></LI>" & _
																	"</OL></BLOCKQUOTE>")
									Response.Write "<BR />"
									Response.Write "<B>Para la carga de este catálogo se requiere: </B>" & sRequiredFields & "."
									Response.Write "<BR />"
									Response.Write "<BR />"
									Response.Write "<TEXTAREA NAME=""RawData"" ID=""RawDataTxtArea"" ROWS=""10"" COLS=""119"" CLASS=""TextFields"" onChange=""bReady = (this.value != '')""></TEXTAREA><BR /><BR />"
									Response.Write "<INPUT TYPE=""SUBMIT"" VALUE=""Continuar"" CLASS=""Buttons"" />"
								Response.Write "</FORM>"
								Response.Write "<IFRAME SRC=""BrowserFileForInfo.asp?Action=" & oRequest("Action").Item & "&UserID=" & aLoginComponent(N_USER_ID_LOGIN) & """ NAME=""UploadInfoIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""100""></IFRAME>"
								Response.Write "<BR />"
						Response.Write "</DIV>"
						Response.Write "<IMG SRC=""Images/Crcl3.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""5"" /><A HREF=""javascript: if(document.all['CatalogsDiv'] != null) { HideDisplay(document.all['CatalogsDiv']) }; if(document.all['UploadInfoFormDiv'] != null) { HideDisplay(document.all['UploadInfoFormDiv']) }; ShowDisplay(document.all['UploadValidateInfoFormDiv']);"">Registros en proceso de aplicación</A><BR /><BR />"
						Response.Write "<DIV NAME=""UploadValidateInfoFormDiv"" ID=""UploadValidateInfoFormDiv"" STYLE=""display: none"">"
							Response.Write "<FORM NAME=""UploadValidateInfoFrm"" ID=""UploadValidateInfoFrm"" METHOD=""POST"" onSubmit=""return CheckPayrollFields(this)"">"
								Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""10"">"
									Response.Write "<TR>"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PositionsSpecialJourneysAction"" ID=""PositionsSpecialJourneysActionHdn"" VALUE=""1"" />"
										'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptValuesAction"" ID=""ConceptValuesActionHdn"" VALUE=""1"" />"
										If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<TD><INPUT TYPE=""SUBMIT"" NAME=""AuthorizationFile"" ID=""ModifyBtn"" VALUE=""Aplicar registros en proceso"" CLASS=""Buttons""/></TD>"
									Response.Write "</TR>"
								Response.Write "</TABLE>"
								aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (PositionsSpecialJourneysLKP.Active<=0)"
								lErrorNumber = DisplayTables(sAction, sErrorDescription)
								If lErrorNumber <> 0 Then
									Response.Write "<BR />"
									Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
									lErrorNumber = 0
									sErrorDescription = ""
								End If
							Response.Write "</FORM>"
						Response.Write "</DIV>"
					Case Else
				End Select
			End If
			If Len(oRequest("Success").Item) > 0 Then
				If CInt(oRequest("Success").Item) = 1 Then
					Call DisplayErrorMessage("Confirmación", "La operación fué ejecutada exitosamente.")
				Else
					Call DisplayErrorMessage("Error al realizar la operación", CStr(oRequest("sErrorDescription").Item))
				End If
			End If
		Else
			Select Case sAction
				Case "BanamexCensus","ConsarFile","EstrQna", "EmployeesDeleted"
				Case Else
					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "function DocumentOnLoad() {" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "//--></SCRIPT>" & vbNewLine
					Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 1. </B>Introduzca el archivo a utilizar.<BR /><BR />"
			End Select

			Select Case sAction
				Case "BanamexCensus","ConsarFile","EstrQna", "EmployeesDeleted"
					lErrorNumber = DisplayTables(sAction, sErrorDescription)
				Case "PositionsSpecialJourneysLKP"
					Select Case iStep
						Case 2
							Call DisplayPositionsSpecialJourneysLKPColumns(sFileName, sErrorDescription)
						Case 3
							Response.Write "<IMG SRC=""Images/IcnCheckBig.gif"" WIDTH=""15"" HEIGHT=""15"" ALIGN=""ABSMIDDLE"">&nbsp;<B>Paso 2. </B>Proceso de la información.<BR /><BR />"
							lErrorNumber = UploadPositionsSpecialJourneysLKPFile(oADODBConnection, sFileName, sErrorDescription)
							Response.Write "<BR />"
							If lErrorNumber = 0 Then
								Call DisplayErrorMessage("Confirmación", "Los puestos para guardias fueron registrados con éxito.")
							Else
								Call DisplayErrorMessage("Error al registrar a los puestos para guardias.", sErrorDescription)
								lErrorNumber = 0
								sErrorDescription = ""
							End If
					End Select
			End Select
		End If
		If lErrorNumber <> 0 Then
			Call DisplayErrorMessage("Error al realizar la operación", sErrorDescription)
			Response.Write "<BR />"
			lErrorNumber = 0
			sErrorDescription = ""
		End If
		%>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>