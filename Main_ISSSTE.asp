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
<!-- #include file="Libraries/CatalogsLib.asp" -->
<!-- #include file="Libraries/EmployeeComponent.asp" -->
<!-- #include file="Libraries/JobComponent.asp" -->
<!-- #include file="Libraries/Main_ISSSTELib.asp" -->
<!-- #include file="Libraries/PayrollComponent.asp" -->
<!-- #include file="Libraries/UploadInfoLibrary.asp" -->
<!-- #include file="Libraries/ReportsLib.asp" -->
<%
Dim iSectionID
Dim iPayrollID
Dim lEmployeeID
Dim sAction
Dim bAction
Dim bSearchForm
Dim bShowForm
Dim bError
Dim sSubSectionID
Dim sUserName
Dim sPayrollName
Dim asLockForPayroll
Dim iStep
Dim lReasonID

iStep = 1
lReasonID = 1
sSubSectionID = ""

If Len(oRequest("ReasonID").Item) > 0 Then lReasonID = CLng(oRequest("ReasonID").Item)
If CLng(oRequest("SubSectionID").Item) > 0 Then
	If CInt(oRequest("SubSectionID").Item) = 1 Then
		sSubSectionID = "&SubSectionID=1"
	ElseIf CInt(oRequest("SubSectionID").Item) = 2 Then
		sSubSectionID = "&SubSectionID=2"
	ElseIf CInt(oRequest("SubSectionID").Item) = 4 Then
		sSubSectionID = "&SubSectionID=4"
	End If
End If

sAction = oRequest("Action").Item
Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)

Select Case oRequest("Action").Item
	Case "DocumentsForLicenses"
		iSectionID = 62
	Case "EmployeesAntiquitiesLKP"
		iSectionID = 261
	Case "EmployeesDocs"
		iSectionID = 267
	Case "EmployeesKardex"
		iSectionID = 352
	Case "EmployeesKardex2"
		iSectionID = 356
	Case "EmployeesKardex3"
		iSectionID = 351
	Case "EmployeesKardex4"
		iSectionID = 282
	Case "EmployeesKardex5"
		iSectionID = 281
	Case "EmployeesSpecialJourneys"
		If Len(oRequest("SpecialJourneyID").Item) = 0 Then
			iSectionID = CInt(oRequest("SectionID").Item)
		Else
			iSectionID = CInt(oRequest("SpecialJourneyID").Item)
		End If
	Case "SADE_NewCourse"
		iSectionID = 369
	Case Else
		iSectionID = CInt(oRequest("SectionID").Item)
End Select
lEmployeeID = 0
If Len(oRequest("EmployeeID").Item) > 0 Then
	If CLng(oRequest("EmployeeID").Item) >= 800000 Then lEmployeeID = 800000
End If

Select Case iSectionID
	Case EMPLOYEES_SERVICE_SHEET
		If Len(oRequest("Add").Item) > 0 Then
			If lErrorNumber = 0 Then
				aEmployeeComponent(N_EMPLOYEE_DOCUMENT_ID) = -1
				lErrorNumber = AddEmployeeDocument(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			End If
			If lErrorNumber = 0 Then
				Response.Redirect "Main_ISSSTE.asp?Action=ServiceSheet&Success=1&ReasonID=" & lReasonID & "&SectionID=" & iSectionID & "&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
			Else
				Response.Redirect "Main_ISSSTE.asp?Action=ServiceSheet&Success=0&ReasonID=" & lReasonID & "&SectionID=" & iSectionID & "&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sErrorDescription
			End If
		ElseIf Len(oRequest("Authorize").Item) > 0 Then
			lErrorNumber = AuthorizeEmployeeDocument(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				Response.Redirect "Main_ISSSTE.asp?Action=ServiceSheet&Success=1&SectionID=" & iSectionID & "&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
			Else
				Response.Redirect "Main_ISSSTE.asp?Action=ServiceSheet&Success=0&SectionID=" & iSectionID & "&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sErrorDescription
			End If
		ElseIf Len(oRequest("GenerateReport").Item) > 0 Then
			aReportsComponent(N_ID_REPORTS) = ISSSTE_1203_REPORTS
			lErrorNumber = BuildReport1203(oRequest, oADODBConnection, sErrorDescription)
			If lErrorNumber = 0 Then
				Response.Redirect "Main_ISSSTE.asp?Action=ServiceSheet&Success=1&SectionID=" & iSectionID & "&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
			Else
				Response.Redirect "Main_ISSSTE.asp?Action=ServiceSheet&Success=0&SectionID=" & iSectionID & "&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sErrorDescription
			End If
		ElseIf Len(oRequest("Remove").Item) > 0 Then
			lErrorNumber = RemoveEmployeeDocument(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				Response.Redirect "Main_ISSSTE.asp?Action=ServiceSheet&Success=1&SectionID=" & iSectionID & "&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
			Else
				Response.Redirect "Main_ISSSTE.asp?Action=ServiceSheet&Success=0&SectionID=" & iSectionID & "&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sErrorDescription
			End If
		ElseIf Len(oRequest("Modify").Item) > 0 Then
			lErrorNumber = ModifyEmployeeDocument(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				Response.Redirect "Main_ISSSTE.asp?Action=ServiceSheet&Success=1&ReasonID=" & lReasonID & "&SectionID=" & iSectionID & "&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
			Else
				Response.Redirect "Main_ISSSTE.asp?Action=ServiceSheet&Success=0&ReasonID=" & lReasonID & "&SectionID=" & iSectionID & "&EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ErrorDescription=" & sErrorDescription
			End If
		End If
End Select

Response.Cookies("SIAP_SubSectionID") = -1
Response.Cookies("SoS_SectionID") = 1000 + iSectionID
Select Case iSectionID
	Case 1
		Response.Cookies("SIAP_SectionID") = 1
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Personal"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
	Case 15
		Response.Cookies("SIAP_SectionID") = 1
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Acumulados anuales"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
	Case 151
		Response.Cookies("SIAP_SectionID") = 1
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de empleados que no desean el ajuste anual del impuesto sobre la renta"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
	Case 156
		Response.Cookies("SIAP_SectionID") = 1
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Aplicación por empleado del ajuste anual del impuesto sobre la renta"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
	Case 16
		Response.Cookies("SIAP_SectionID") = 1
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reclamos de pago por ajustes y deducciones"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
	Case 18
		Response.Cookies("SIAP_SectionID") = 1
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Administración de personal"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
	Case 19
		Response.Cookies("SIAP_SectionID") = 1
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reportes"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
	Case 2
		Response.Cookies("SIAP_SectionID") = 2
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Prestaciones"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 20
		Response.Cookies("SIAP_SectionID") = 2
		aHeaderComponent(S_TITLE_NAME_HEADER) = "SI. Seguro de separación y AE. Seguro adicional de separación individualizado"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 21
		Response.Cookies("SIAP_SectionID") = 2
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Terceros institucionales"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 211
		Response.Cookies("SIAP_SectionID") = 2
		Response.Cookies("SIAP_SubSectionID") = 211
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Carga de discos de terceros"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 22
		Response.Cookies("SIAP_SectionID") = 2
		Response.Cookies("SIAP_SubSectionID") = 22
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Prestaciones e incidencias"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 23
		Response.Cookies("SIAP_SectionID") = 2
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Pensión alimenticia"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 24
		Response.Cookies("SIAP_SectionID") = 2
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reportes"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 25
		Response.Cookies("SIAP_SectionID") = 2
		Response.Cookies("SIAP_SubSectionID") = 25
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Certificaciones y archivo"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 261
		Response.Cookies("SIAP_SectionID") = 2
		Response.Cookies("SIAP_SubSectionID") = 261
		aCatalogComponent(S_TABLE_NAME_CATALOG) = "EmployeesAntiquitiesLKP"

		lErrorNumber = DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription)
		bError = (lErrorNumber <> 0)

		aHeaderComponent(S_TITLE_NAME_HEADER) = "Antigüedad federal"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 262
		Response.Cookies("SIAP_SectionID") = 2
		Response.Cookies("SIAP_SubSectionID") = 262
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Actualización de antigüedades"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 267
		'Response.Cookies("SIAP_SectionID") = 2
		aCatalogComponent(S_TABLE_NAME_CATALOG) = "EmployeesDocs"

		lErrorNumber = DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription)
		bError = (lErrorNumber <> 0)

		aHeaderComponent(S_TITLE_NAME_HEADER) = "Entregas de hojas únicas de servicio"
		If CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Else
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
		End If
	Case 27
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Pensión alimenticia"
		If CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
			Response.Cookies("SIAP_SectionID") = 2
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		Else
			Response.Cookies("SIAP_SectionID") = 1
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
		End If
	Case 28
		Response.Cookies("SIAP_SectionID") = 2
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de bolsa de trabajo y escalafón"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 281
		aCatalogComponent(S_TABLE_NAME_CATALOG) = "EmployeesKardex5"

		lErrorNumber = DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription)
		bError = (lErrorNumber <> 0)
		If ((Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Or (Len(oRequest("RecordID").Item) > 0)) And (Len(oRequest("Add").Item) = 0) Then
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Búsqueda de información de la bolsa de trabajo"
		Else
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de información de la bolsa de trabajo"
		End If

		If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
		Else
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		End If
	Case 282
		aCatalogComponent(S_TABLE_NAME_CATALOG) = "EmployeesKardex4"

		lErrorNumber = DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription)
		bError = (lErrorNumber <> 0)
		If ((Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Or (Len(oRequest("RecordID").Item) > 0)) And (Len(oRequest("Add").Item) = 0) Then
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Búsqueda de información de escalafón"
		Else
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de información de escalafón"
		End If

		If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
		Else
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		End If
	Case 29
		Response.Cookies("SIAP_SectionID") = 2
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Matriz de Riesgos Profesionales"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 291
		Response.Cookies("SIAP_SectionID") = 2
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Acumulados anuales"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
	Case 3
		Response.Cookies("SIAP_SectionID") = 3
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Desarrollo humano"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 31
		Response.Cookies("SIAP_SectionID") = 3
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Estructuras ocupacionales"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 32
		Response.Cookies("SIAP_SectionID") = 3
		Response.Cookies("SIAP_SubSectionID") = 32
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Consulta de tabuladores"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 33
		Response.Cookies("SIAP_SectionID") = 3
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de tabuladores"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 34
		Response.Cookies("SIAP_SectionID") = 3
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reportes"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 35
		Response.Cookies("SIAP_SectionID") = 3
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Selección de personal"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 351
		Response.Cookies("SIAP_SectionID") = 3
		aCatalogComponent(S_TABLE_NAME_CATALOG) = "EmployeesKardex3"

		lErrorNumber = DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription)
		bError = (lErrorNumber <> 0)

		If ((Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Or (Len(oRequest("RecordID").Item) > 0)) And (Len(oRequest("Add").Item) = 0) Then
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Búsqueda de información"
		Else
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de información"
		End If
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 352
		Response.Cookies("SIAP_SectionID") = 3
		aCatalogComponent(S_TABLE_NAME_CATALOG) = "EmployeesKardex"

		lErrorNumber = DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription)
		bError = (lErrorNumber <> 0)

		aHeaderComponent(S_TITLE_NAME_HEADER) = "Validación del proceso de selección de personal"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 353
		Response.Cookies("SIAP_SectionID") = 3
		aCatalogComponent(S_TABLE_NAME_CATALOG) = "EmployeesKardex3"

		lErrorNumber = DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription)
		bError = (lErrorNumber <> 0)

		aHeaderComponent(S_TITLE_NAME_HEADER) = "Búsqueda de información del proceso de selección"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 354
		Response.Cookies("SIAP_SectionID") = 3
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Búsqueda de información de la bolsa de trabajo"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 356
		Response.Cookies("SIAP_SectionID") = 3
		aCatalogComponent(S_TABLE_NAME_CATALOG) = "EmployeesKardex2"

		lErrorNumber = DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription)
		bError = (lErrorNumber <> 0)

		aHeaderComponent(S_TITLE_NAME_HEADER) = "Búsqueda de información de escalafón"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 36
		Response.Cookies("SIAP_SectionID") = 3
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Desarrollo humano"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 369
		Response.Cookies("SIAP_SectionID") = 3
		aCatalogComponent(S_TABLE_NAME_CATALOG) = "SADE_NewCourse"

		lErrorNumber = DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription)
		bError = (lErrorNumber <> 0)

		aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de detección de necesidades"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 37
		Response.Cookies("SIAP_SectionID") = 3
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Planeación de recursos humanos"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 371
		Response.Cookies("SIAP_SectionID") = 3
		aCatalogComponent(S_TABLE_NAME_CATALOG) = "Documents"

		lErrorNumber = DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription)
		bError = (lErrorNumber <> 0)

		aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de procedimientos"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 38
		Response.Cookies("SIAP_SectionID") = 3
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Búsqueda de centros de trabajo y centros de pago"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
	Case 4
		Response.Cookies("SIAP_SectionID") = 4
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Informática"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
	Case 42
		Response.Cookies("SIAP_SectionID") = 4
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Empleados"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
	Case 421
		Response.Cookies("SIAP_SectionID") = 4
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Incidencias"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
	Case 423
		If (StrComp(oRequest("FromSectionID").Item, "7", vbBinaryCompare) = 0) Or (CInt(Request.Cookies("SIAP_SectionID")) = 7) Then
			Response.Cookies("SIAP_SectionID") = 7
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
		Else
			Response.Cookies("SIAP_SectionID") = 4
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
		End If
		aCatalogComponent(S_TABLE_NAME_CATALOG) = "EmployeesSpecialJourneys"

		lErrorNumber = DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription)
		If lEmployeeID = 0 Then lEmployeeID = CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1))
		If CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(16)) = -1 Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(16) = 423
		bError = (lErrorNumber <> 0)
		If ((Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Or (Len(oRequest("RecordID").Item) > 0)) And (Len(oRequest("Add").Item) = 0) Then
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Búsqueda de información"
		ElseIf (Len(oRequest("New").Item) > 0) Or bAction Then
			If lEmployeeID < 800000 Then
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de información para internos"
			Else
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de información para externos"
			End If
		Else
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Guardias"
		End If
	Case 424
		If (StrComp(oRequest("FromSectionID").Item, "7", vbBinaryCompare) = 0) Or (CInt(Request.Cookies("SIAP_SectionID")) = 7) Then
			Response.Cookies("SIAP_SectionID") = 7
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
		Else
			Response.Cookies("SIAP_SectionID") = 4
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
		End If
		aCatalogComponent(S_TABLE_NAME_CATALOG) = "EmployeesSpecialJourneys"

		lErrorNumber = DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription)
		If lEmployeeID = 0 Then lEmployeeID = CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1))
		If CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(16)) = -1 Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(16) = 424
		bError = (lErrorNumber <> 0)
		If ((Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Or (Len(oRequest("RecordID").Item) > 0)) And (Len(oRequest("Add").Item) = 0) Then
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Búsqueda de información"
		ElseIf (Len(oRequest("New").Item) > 0) Or bAction Then
			If lEmployeeID < 800000 Then
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de información para internos"
			Else
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de información para externos"
			End If
		Else
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Suplencias"
		End If
	Case 425
		If (StrComp(oRequest("FromSectionID").Item, "7", vbBinaryCompare) = 0) Or (CInt(Request.Cookies("SIAP_SectionID")) = 7) Then
			Response.Cookies("SIAP_SectionID") = 7
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
		Else
			Response.Cookies("SIAP_SectionID") = 4
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
		End If
		aCatalogComponent(S_TABLE_NAME_CATALOG) = "EmployeesSpecialJourneys"

		lErrorNumber = DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription)
		If lEmployeeID = 0 Then lEmployeeID = CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1))
		If CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(16)) = -1 Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(16) = 425
		bError = (lErrorNumber <> 0)
		If ((Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Or (Len(oRequest("RecordID").Item) > 0)) And (Len(oRequest("Add").Item) = 0) Then
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Búsqueda de información"
		ElseIf (Len(oRequest("New").Item) > 0) Or bAction Then
			If lEmployeeID < 800000 Then
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de información para internos"
			Else
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de información para externos"
			End If
		Else
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Rezago quirúrgico"
		End If
	Case 426
		If (StrComp(oRequest("FromSectionID").Item, "7", vbBinaryCompare) = 0) Or (CInt(Request.Cookies("SIAP_SectionID")) = 7) Then
			Response.Cookies("SIAP_SectionID") = 7
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
		Else
			Response.Cookies("SIAP_SectionID") = 4
			aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
		End If
		aCatalogComponent(S_TABLE_NAME_CATALOG) = "EmployeesSpecialJourneys"

		lErrorNumber = DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription)
		If lEmployeeID = 0 Then lEmployeeID = CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1))
		If CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(16)) = -1 Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(16) = 426
		bError = (lErrorNumber <> 0)
		If ((Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Or (Len(oRequest("RecordID").Item) > 0)) And (Len(oRequest("Add").Item) = 0) Then
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Búsqueda de información"
		ElseIf (Len(oRequest("New").Item) > 0) Or bAction Then
			If lEmployeeID < 800000 Then
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de información para internos"
			Else
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Registro de información para externos"
			End If
		Else
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Programa de vacunación"
		End If
	Case 427
		Response.Cookies("SIAP_SectionID") = 4
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Entrada del archivo de FOVISSSTE"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
	Case 429
		Response.Cookies("SIAP_SectionID") = 4
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reclamos de pago por ajustes y deducciones"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
	Case 47
		Response.Cookies("SIAP_SectionID") = 4
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Cheques y depósitos"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
	Case 48
		Response.Cookies("SIAP_SectionID") = 4
		Select Case CInt(oRequest("Action").Item)
			Case 1
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Apertura y cierre para movimientos de personal y prestaciones"
			Case 2
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Apertura y cierre para incidencias"
			Case 3
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Apertura y cierre para el padrón de madres"
			Case 4
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Apertura y cierre para el registro de cuentas bancarias"
			Case 5
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Apertura y cierre para guardias, suplencias, rezago quirúrgico y programa de vacunación"
			Case 6
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Apertura y cierre para FONAC"
			Case 7
				aHeaderComponent(S_TITLE_NAME_HEADER) = ""
			Case 8
				aHeaderComponent(S_TITLE_NAME_HEADER) = ""
			Case 9
				aHeaderComponent(S_TITLE_NAME_HEADER) = ""
			Case 10
				aHeaderComponent(S_TITLE_NAME_HEADER) = ""
			Case 11
				aHeaderComponent(S_TITLE_NAME_HEADER) = ""
			Case 12
				aHeaderComponent(S_TITLE_NAME_HEADER) = ""
			Case Else
				aHeaderComponent(S_TITLE_NAME_HEADER) = "Apertura y cierre de registros"
		End Select
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
		Call InitializePayrollComponent(oRequest, aPayrollComponent)
		If Len(oRequest("ModifyStatus").Item) > 0 Then
			lErrorNumber = GetPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				Select Case CInt(oRequest("Action").Item)
					Case 1
						aPayrollComponent(N_IS_ACTIVE_1_PAYROLL) = CInt(oRequest("StatusID").Item)
					Case 2
						aPayrollComponent(N_IS_ACTIVE_2_PAYROLL) = CInt(oRequest("StatusID").Item)
					Case 3
						aPayrollComponent(N_IS_ACTIVE_3_PAYROLL) = CInt(oRequest("StatusID").Item)
					Case 4
						aPayrollComponent(N_IS_ACTIVE_4_PAYROLL) = CInt(oRequest("StatusID").Item)
					Case 5
						aPayrollComponent(N_IS_ACTIVE_5_PAYROLL) = CInt(oRequest("StatusID").Item)
					Case 6
						aPayrollComponent(N_IS_ACTIVE_6_PAYROLL) = CInt(oRequest("StatusID").Item)
					Case 7
						aPayrollComponent(N_IS_ACTIVE_7_PAYROLL) = CInt(oRequest("StatusID").Item)
					Case 8
						aPayrollComponent(N_IS_ACTIVE_8_PAYROLL) = CInt(oRequest("StatusID").Item)
					Case 9
						aPayrollComponent(N_IS_ACTIVE_9_PAYROLL) = CInt(oRequest("StatusID").Item)
					Case 10
						aPayrollComponent(N_IS_ACTIVE_10_PAYROLL) = CInt(oRequest("StatusID").Item)
					Case 11
						aPayrollComponent(N_IS_ACTIVE_11_PAYROLL) = CInt(oRequest("StatusID").Item)
					Case 12
						aPayrollComponent(N_IS_ACTIVE_12_PAYROLL) = CInt(oRequest("StatusID").Item)
				End Select
				lErrorNumber = ModifyPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
			End If
			bError = (lErrorNumber <> 0)
		End If
	Case 49
		Response.Cookies("SIAP_SectionID") = 4
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reportes"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
	Case 491
		Response.Cookies("SIAP_SectionID") = 4
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Informática"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
	Case 5
		Response.Cookies("SIAP_SectionID") = 5
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Presupuesto"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = BUDGET_TOOLBAR
	Case 53
		Response.Cookies("SIAP_SectionID") = 5
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Costeo de plazas"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = BUDGET_TOOLBAR
	Case 56
		Response.Cookies("SIAP_SectionID") = 5
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reportes sobre el costeo de plazas"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = BUDGET_TOOLBAR
	Case 58
		Response.Cookies("SIAP_SectionID") = 5
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reportes"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = BUDGET_TOOLBAR
	Case 6
		Response.Cookies("SIAP_SectionID") = 6
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Departamento Técnico"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = REPORTS_TOOLBAR
	Case 61
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Ventanilla única"
		Select Case CInt(Request.Cookies("SIAP_SectionID"))
			Case 1
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = CATALOGS_TOOLBAR
			Case 2
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
			Case 3
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = HUMAN_RESOURCES_TOOLBAR
			Case 4
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYROLL_TOOLBAR
			Case 5
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = BUDGET_TOOLBAR
			Case 6
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = REPORTS_TOOLBAR
			Case Else
				aHeaderComponent(L_SELECTED_OPTION_HEADER) = PAYMENTS_TOOLBAR
		End Select
	Case 62
		Response.Cookies("SIAP_SectionID") = 6
		aCatalogComponent(S_TABLE_NAME_CATALOG) = "DocumentsForLicenses"

		lErrorNumber = DoCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, bSearchForm, bShowForm, sErrorDescription)
		bError = (lErrorNumber <> 0)

		If Len(oRequest("Search").Item) > 0 Then
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Búsqueda de licencias sindicales"
		Else
			aHeaderComponent(S_TITLE_NAME_HEADER) = "Emisión de licencias por comisión sindical"
		End If
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = REPORTS_TOOLBAR
	Case 64
		Response.Cookies("SIAP_SectionID") = 6
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reportes"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = REPORTS_TOOLBAR
	Case 7
		Response.Cookies("SIAP_SectionID") = 7
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Desconcentrados"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
	Case 71
		Response.Cookies("SIAP_SectionID") = 7
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Personal"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
	Case 712
		Response.Cookies("SIAP_SectionID") = 7
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Administración de personal"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
	Case 713
		Response.Cookies("SIAP_SectionID") = 7
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reportes"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
	Case 72
		Response.Cookies("SIAP_SectionID") = 7
		Response.Cookies("SIAP_SubSectionID") = 72
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Prestaciones"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
	Case 721
		Response.Cookies("SIAP_SectionID") = 7
		Response.Cookies("SIAP_SubSectionID") = 721
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Prestaciones e incidencias"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
	Case 73
		Response.Cookies("SIAP_SectionID") = 7
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Informática"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
	Case 731
		Response.Cookies("SIAP_SectionID") = 7
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Empleados"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
	Case 732
		Response.Cookies("SIAP_SectionID") = 7
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Cheques y depósitos"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
	Case 733
		Response.Cookies("SIAP_SectionID") = 7
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Reportes"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
	Case 74
		Response.Cookies("SIAP_SectionID") = 7
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Presupuesto"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
	Case 75
		Response.Cookies("SIAP_SectionID") = 7
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Tablero de Control"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
    Case 8
		Response.Cookies("SIAP_SectionID") = 8
		aHeaderComponent(S_TITLE_NAME_HEADER) = "Atención al personal"
		aHeaderComponent(L_SELECTED_OPTION_HEADER) = LOGOUT_TOOLBAR
End Select
bWaitMessage = True
%>
<HTML>
	<HEAD>
		<!-- #include file="_JavaScript.asp" -->
	</HEAD>
	<BODY TEXT="#000000" LINK="#000000" ALINK="#000000" VLINK="#000000" BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		<%If (InStr(1, ",261,", "," & iSectionID & ",", vbBinaryCompare) > 0) And (Len(oRequest("EmployeeID").Item) > 0) Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Agregar un nuevo registro",_
					  "",_
					  "", "Main_ISSSTE.asp?SectionID=" & iSectionID & "&EmployeeID=" & oRequest("EmployeeID").Item & "&New=1", True)_
			)
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 803
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 190
		ElseIf (InStr(1, ",262,", "," & iSectionID & ",", vbBinaryCompare) > 0) And (Len(oRequest("EmployeeID").Item) > 0) Then
			If ((aLoginComponent(N_USER_PERMISSIONS4_LOGIN) And N_07_PERMISSIONS4) = N_07_PERMISSIONS4) And (InStr(1, ",-1,2,", "," & aLoginComponent(N_PROFILE_ID_LOGIN) & ",", vbBinaryCompare) > 0) Then
				aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
					Array("Agregar un registro al historial",_
						  "",_
						  "", "Main_ISSSTE.asp?SectionID=262&EmployeeID=" & oRequest("EmployeeID").Item & "&EmployeeDate=0", True)_
				)
				aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 793
				aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
				aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 200
			End If
		ElseIf (iSectionID = 267) And (Len(oRequest("DoSearch").Item) > 0) Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Excel=1&Action=EmployeesDocs&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
			)
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 843
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 150
		ElseIf (iSectionID = 281) And (Len(oRequest("DoSearch").Item) > 0) Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Excel=1&Action=EmployeesKardex5&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
			)
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 843
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 150
		ElseIf (iSectionID = 282) And (Len(oRequest("DoSearch").Item) > 0) Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Excel=1&Action=EmployeesKardex4&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
			)
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 843
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 150
		ElseIf (iSectionID = 353) And (Len(oRequest("DoSearch").Item) > 0) Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Excel=1&Action=EmployeesKardex3&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
			)
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 843
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 150
		ElseIf (iSectionID = 352) And (Len(oRequest("DoSearch").Item) > 0) Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Excel=1&Action=EmployeesKardex&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
			)
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 843
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 150
		ElseIf (iSectionID = 356) And (Len(oRequest("DoSearch").Item) > 0) Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Excel=1&Action=EmployeesKardex2&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
			)
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 843
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 150
		ElseIf (iSectionID = 371) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS) = N_ADD_PERMISSIONS) Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Agregar un nuevo procedimiento",_
					  "",_
					  "", "Main_ISSSTE.asp?SectionID=371&New=1", True)_
			)
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 763
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 230
		ElseIf (iSectionID = 369) And (Len(oRequest("DoSearch").Item) > 0) Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=Reports&ReportID=1369&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
			)
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 843
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 150
		ElseIf (InStr(1, ",423,424,425,426,", "," & iSectionID & ",", vbBinaryCompare) > 0) And (Len(oRequest("DoSearch").Item) > 0) Then
			aOptionsMenuComponent(A_ELEMENTS_MENU) = Array(_
				Array("Exportar a Excel",_
					  "",_
					  "", "javascript: OpenNewWindow('Export.asp?Action=EmployeesSpecialJourneys&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "&" & RemoveEmptyParametersFromURLString(oRequest) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')", True)_
			)
			aOptionsMenuComponent(N_LEFT_FOR_DIV_MENU) = 843
			aOptionsMenuComponent(N_TOP_FOR_DIV_MENU) = 82
			aOptionsMenuComponent(N_WIDTH_FOR_DIV_MENU) = 150
		End If%>
		<!-- #include file="_Header.asp" -->
		<%Response.Write "Usted se encuentra aquí: <A HREF=""Main.asp"">Inicio</A> > "
			Select Case iSectionID
				Case 1
					Response.Write "<B>Personal</B>"
				Case 15
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <B>Acumulados anuales</B>"
				Case 151
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=15"">Acumulados anuales</A> > <B>Registro de empleados que no desean el ajuste anual del impuesto sobre la renta</B>"
				Case 156
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=15"">Acumulados anuales</A> > <B>Aplicación por empleado del ajuste anual del impuesto sobre la renta</B>"
				Case 16
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <B>Reclamos de pago por ajustes y deducciones</B>"
				Case 18
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <B>Administración de personal</B>"
				Case 19
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <B>Reportes</B>"
				Case 2
					Response.Write "<B>Prestaciones</B>"
				Case 20
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <B>SI. Seguro de separación y AE. Seguro adicional de separación individualizado</B>"
				Case 21
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <B>Terceros institucionales</B>"
				Case 22
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <B>Prestaciones e incidencias</B>"
				Case 23
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <B>Pensión alimenticia</B>"
				Case 24
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <B>Reportes</B>"
				Case 25
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <B>Certificaciones y archivo</B>"
				Case 26
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <B>Antigüedades</B>"
				Case 261
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=25"">Certificaciones y archivo</A> > <B>Antigüedad federal</B>"
				Case 262
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=25"">Certificaciones y archivo</A> > <B>Actualización de antigüedades</B>"
				Case 267
					If CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
						If Len(oRequest("New").Item) > 0 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=25"">Certificaciones y archivo</A> > <A HREF=""Main_ISSSTE.asp?SectionID=267"">Entregas de hojas únicas de servicio</A> > <B>Registro de información</B>"
						ElseIf ((Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Or (Len(oRequest("RecordID").Item) > 0)) And (Len(oRequest("Add").Item) = 0) Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=25"">Certificaciones y archivo</A> > <A HREF=""Main_ISSSTE.asp?SectionID=267"">Entregas de hojas únicas de servicio</A> > <B>Búsqueda de información</B>"
						Else
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=25"">Certificaciones y archivo</A> > <B>Entregas de hojas únicas de servicio</B>"
						End If
					Else
						If Len(oRequest("New").Item) > 0 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=72"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=267"">Entregas de hojas únicas de servicio</A> > <B>Registro de información</B>"
						ElseIf ((Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Or (Len(oRequest("RecordID").Item) > 0)) And (Len(oRequest("Add").Item) = 0) Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=72"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=267"">Entregas de hojas únicas de servicio</A> > <B>Búsqueda de información</B>"
						Else
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=72"">Prestaciones</A> > <B>Entregas de hojas únicas de servicio</B>"
						End If
					End If
				Case 211
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=21"">Terceros institucionales</A> > <B>Carga de discos de terceros</B>"
				Case 212
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=21"">Terceros institucionales</A> > <B>Generación de archivos para enteros</B>"
				Case 213
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=21"">Terceros institucionales</A> > <B>Reportes</B>"
				Case 241
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=21"">Terceros institucionales</A> > <B>Generación de archivos para enteros</B>"
				Case 27
					If CInt(Request.Cookies("SIAP_SectionID")) = 2 Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <B>Acreedores de los empleados</B>"
					Else
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <B>Acreedores de los empleados</B>"
					End If
				Case 28
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <B>Registro de bolsa de trabajo y escalafón</B>"
				Case 281
					If Len(oRequest("New").Item) > 0 Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=28"">Registro de bolsa de trabajo y escalafón</A> > <B>Registro de información de la bolsa de trabajo</B>"
					ElseIf ((Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Or (Len(oRequest("RecordID").Item) > 0)) And (Len(oRequest("Add").Item) = 0) Then
						If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=35"">Selección de personal</A> > <B>Búsqueda de información de la bolsa de trabajo</B>"
						Else
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=28"">Registro de bolsa de trabajo y escalafón</A> > <B>Búsqueda de información de la bolsa de trabajo</B>"
						End If
					Else
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=28"">Registro de bolsa de trabajo y escalafón</A> > <B>Registro de información de la bolsa de trabajo</B>"
					End If
				Case 282
					If Len(oRequest("New").Item) > 0 Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=28"">Registro de bolsa de trabajo y escalafón</A> > <B>Registro de información de escalafón</B>"
					ElseIf ((Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Or (Len(oRequest("RecordID").Item) > 0)) And (Len(oRequest("Add").Item) = 0) Then
						If CInt(Request.Cookies("SIAP_SectionID")) = 3 Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=35"">Selección de personal</A> > <B>Búsqueda de información de escalafón</B>"
						Else
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=28"">Registro de bolsa de trabajo y escalafón</A> > <B>Búsqueda de información de escalafón</B>"
						End If
					Else
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=28"">Registro de bolsa de trabajo y escalafón</A> > <B>Registro de información de escalafón</B>"
					End If
				Case 29
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <B>Matriz de Riesgos Profesionales </B>"
				Case 291
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <B>Acumulados anuales </B>"
				Case 3
					Response.Write "<B>Desarrollo humano</B>"
				Case 31
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <B>Estructuras ocupacionales</B>"
				Case 32
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=31"">Estructuras ocupacionales</A> > <B>Consulta de tabuladores</B>"
				Case 33
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=31"">Estructuras ocupacionales</A> > <B>Registro de tabuladores</B>"
				Case 34
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <B>Reportes</B>"
				Case 35
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <B>Selección de personal</B>"
				Case 351
					If Len(oRequest("New").Item) > 0 Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=35"">Selección de personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=351"">Registro de información</A> > <B>Registro de información</B>"
					ElseIf (Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Or (Len(oRequest("RecordID").Item) > 0) Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=35"">Selección de personal</A> > <A HREF=""Main_ISSSTE.asp?SectionID=351"">Registro de información</A> > <A HREF=""Main_ISSSTE.asp?SectionID=351&Search=1""><B>Búsqueda de información</B></A>"
					Else
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=35"">Selección de personal</A> > <B>Registro de información</B>"
					End If
				Case 352
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=35"">Selección de personal</A> > <B>Validación del proceso de selección de personal</B>"
				Case 353
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=35"">Selección de personal</A> > <B>Búsqueda de información del proceso de selección</B>"
				Case 354
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=35"">Selección de personal</A> > <B>Búsqueda de información de la bolsa de tabajo</B>"
				Case 356
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=35"">Selección de personal</A> > <B>Búsqueda de información de escalafón</B>"
				Case 36
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <B>Desarrollo humano</B>"
				Case 369
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=36"">Desarrollo humano</A> > "
					If (Len(oRequest("New").Item) > 0) Or bAction Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=369"">Registro de detección de necesidades</A> > <B>Registro de información</B>"
					ElseIf (Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=369"">Registro de detección de necesidades</A> > <B>Búsqueda de información</B>"
					Else
						Response.Write "<B>Registro de detección de necesidades</B>"
					End If
				Case 37
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <B>Planeación de recursos humanos</B>"
				Case 371
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <A HREF=""Main_ISSSTE.asp?SectionID=37"">Planeación de recursos humanos</A> > <B>Registro de procedimientos</B>"
				Case 38
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <B>Búsqueda de centros de trabajo y centros de pago</B>"
				Case 4
					Response.Write "<B>Informática</B>"
				Case 42
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Empleados</B>"
				Case 421
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <B>Incidencias</B>"
				Case 423
					If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > "
					Else
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > "
					End If
					If (Len(oRequest("New").Item) > 0) Or (Len(oRequest("Change").Item) > 0) Or bAction Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=423"">Guardias</A> > "
						If lEmployeeID < 800000 Then
							Response.Write "<B>Registro de información para internos</B>"
						Else
							Response.Write "<B>Registro de información para externos</B>"
						End If
					ElseIf (Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=423"">Guardias</A> > <B>Búsqueda de información</B>"
					Else
						Response.Write "<B>Guardias</B>"
					End If
				Case 424
					If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > "
					Else
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > "
					End If
					If (Len(oRequest("New").Item) > 0) Or (Len(oRequest("Change").Item) > 0) Or bAction Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=424"">Suplencias</A> > "
						If lEmployeeID < 800000 Then
							Response.Write "<B>Registro de información para internos</B>"
						Else
							Response.Write "<B>Registro de información para externos</B>"
						End If
					ElseIf (Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=424"">Suplencias</A> > <B>Búsqueda de información</B>"
					Else
						Response.Write "<B>Suplencias</B>"
					End If
				Case 425
					If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > "
					Else
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > "
					End If
					If (Len(oRequest("New").Item) > 0) Or (Len(oRequest("Change").Item) > 0) Or bAction Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=425"">Rezago quirúrgico</A> > "
						If lEmployeeID < 800000 Then
							Response.Write "<B>Registro de información para internos</B>"
						Else
							Response.Write "<B>Registro de información para externos</B>"
						End If
					ElseIf (Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=425"">Rezago quirúrgico</A> > <B>Búsqueda de información</B>"
					Else
						Response.Write "<B>Rezago quirúrgico</B>"
					End If
				Case 426
					If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=731"">Empleados</A> > "
					Else
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > "
					End If
					If (Len(oRequest("New").Item) > 0) Or (Len(oRequest("Change").Item) > 0) Or bAction Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=426"">Programa de vacunación</A> > "
						If lEmployeeID < 800000 Then
							Response.Write "<B>Registro de información para internos</B>"
						Else
							Response.Write "<B>Registro de información para externos</B>"
						End If
					ElseIf (Len(oRequest("Search").Item) > 0) Or (Len(oRequest("DoSearch").Item) > 0) Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=426"">Programa de vacunación</A> > <B>Búsqueda de información</B>"
					Else
						Response.Write "<B>Programa de vacunación</B>"
					End If
				Case 427
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <B>Entrada del archivo de FOVISSSTE</B>"
				Case 429
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=42"">Empleados</A> > <B>Reclamos de pago por ajustes y deducciones</B>"
				Case 47
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Cheques y depósitos</B>"
				Case 48
					Select Case CInt(oRequest("Action").Item)
						Case 1
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=48"">Apertura y cierre de registros</A> > <B>Apertura y cierre para movimientos de personal y prestaciones</B>"
						Case 2
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=48"">Apertura y cierre de registros</A> > <B>Apertura y cierre para incidencias</B>"
						Case 3
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=48"">Apertura y cierre de registros</A> > <B>Apertura y cierre para el padrón de madres</B>"
						Case 4
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=48"">Apertura y cierre de registros</A> > <B>Apertura y cierre para el registro de cuentas bancarias</B>"
						Case 5
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=48"">Apertura y cierre de registros</A> > <B>Apertura y cierre para guardias, suplencias, rezago quirúrgico y programa de vacunación</B>"
						Case 6
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <A HREF=""Main_ISSSTE.asp?SectionID=48"">Apertura y cierre de registros</A> > <B>Apertura y cierre para FONAC</B>"
						Case Else
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Apertura y cierre de registros</B>"
					End Select
				Case 49
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Reportes</B>"
				Case 491
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Ejercicio bimestral del SAR</B>"
				Case 5
					Response.Write "<B>Presupuesto</B>"
				Case 53
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <B>Costeo de plazas</B>"
				Case 56
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <B>Reportes sobre el costeo de plazas</B>"
				Case 58
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <B>Reportes</B>"
				Case 6
					Response.Write "<B>Departamento técnico</B>"
				Case 61
					Select Case CInt(Request.Cookies("SIAP_SectionID"))
						Case 1
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=1"">Personal</A> > <B>Ventanilla única</B>"
						Case 2
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <B>Ventanilla única</B>"
						Case 3
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=3"">Desarrollo humano</A> > <B>Ventanilla única</B>"
						Case 4
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=4"">Informática</A> > <B>Ventanilla única</B>"
						Case 5
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=5"">Presupuesto</A> > <B>Ventanilla única</B>"
						Case 6
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > <B>Ventanilla única</B>"
                        Case 8
                            Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=8"">Atención al personal</A> > <B>Trámites al personal</B>"
						Case Else
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <B>Ventanilla única</B>"
					End Select
				Case 62
					If (Len(oRequest("Search").Item) = 1) Then
						Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > <A HREF=""Main_ISSSTE.asp?SectionID=62"">Emisión de licencias por comisión sindical</A> > <B>Búsqueda de licencias sindicales</B>"
					Else
						If (Len(oRequest("DoSearch").Item) <> 0) Then
							Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > <A HREF=""Main_ISSSTE.asp?SectionID=62"">Emisión de licencias por comisión sindical</A> > <B>Búsqueda de licencias sindicales</B>"
						Else
							If (Len(oRequest("Change").Item) <> 0) Then
								Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > <A HREF=""Main_ISSSTE.asp?SectionID=62"">Emisión de licencias por comisión sindical</A> > <A HREF=""Main_ISSSTE.asp?SectionID=62&Search=1"">Búsqueda de licencias sindicales</A> > <B>Modificar</B>"
							Else
								If (Len(oRequest("Delete").Item) <> 0) Then
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > <A HREF=""Main_ISSSTE.asp?SectionID=62"">Emisión de licencias por comisión sindical</A> > <A HREF=""Main_ISSSTE.asp?SectionID=62&Search=1"">Búsqueda de licencias sindicales</A> > <B>Eliminar</B>"
								Else
									Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > <B>Emisión de licencias por comisión sindical</B>"
								End If
							End If
						End If
					End If
				Case 64
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=6"">Departamento técnico</A> > <B>Reportes</B>"
				Case EMPLOYEES_SERVICE_SHEET
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=2"">Prestaciones</A> > <A HREF=""Main_ISSSTE.asp?SectionID=25"">Certificaciones y archivos</A> > <B>Hoja única de servicio</B>"
				Case 7
					Response.Write "<B>Desconcentrados</B>"
				Case 71
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <B>Personal</B>"
				Case 712
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > <B>Administración de personal</B>"
				Case 713
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=71"">Personal</A> > <B>Reportes</B>"
				Case 72
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <B>Prestaciones</B>"
				Case 721
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=72"">Prestaciones</A> > <B>Prestaciones e incidencias</B>"
				Case 723
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=72"">Prestaciones</A> > <B>Reportes</B>"
				Case 73
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <B>Informática</B>"
				Case 731
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <B>Empleados</B>"
				Case 732
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <B>Cheques y depósitos</B>"
				Case 733
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <A HREF=""Main_ISSSTE.asp?SectionID=73"">Informática</A> > <B>Reportes</B>"
				Case 74
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <B>Presupuesto</B>"
				Case 75
					Response.Write "<A HREF=""Main_ISSSTE.asp?SectionID=7"">Desconcentrados</A> > <B>Tablero de Control</B>"
				Case 8
					Response.Write "<B>Atención al personal</B>"
			End Select
		Response.Write "<BR />"
		Select Case iSectionID
			Case EMPLOYEES_SERVICE_SHEET
				Response.Write "<BR /><BR />"
			Case Else
				Response.Write "<BR /><BR /><TABLE WIDTH=""720"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		End Select
			Select Case iSectionID
				Case 1 'Personal
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Administración de plazas",_
							  "Busque las plazas que desea administrar.",_
							  "Images/MnJobs.gif", "Jobs.asp", ((aLoginComponent(N_USER_PERMISSIONS4_LOGIN) And N_01_PERMISSIONS4) = N_01_PERMISSIONS4) And False),_
						Array("Asignación de número de empleado",_
							  "Asigne un número al nuevo empleado antes de darlo de alta.",_
							  "Images/MnSection12.gif", "UploadInfo.asp?Action=EmployeesAssignNumber&ReasonID=0", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_AsignacionDeNumeroDeEmpleado & ",", vbBinaryCompare) > 0),_
						Array("Consulta de personal",_
							  "Consulte la información de los empleados, plaza, conceptos de pago, historia.",_
							  "Images/MnEmployees.gif", "Employees.asp", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ConsultaDePersonal & ",", vbBinaryCompare) > 0),_
						Array("Administración de personal",_
							  "Realice movimientos de nuevo ingreso, reingreso, bajas, etc. al personal del Instituto.",_
							  "Images/MnSection34.gif", "Main_ISSSTE.asp?SectionID=18", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_AdministracionDePersonal & ",", vbBinaryCompare) > 0),_
						Array("Aguinaldos",_
							  "Obtenga el listado con la información con que cuenta el sistema de nómina para generar el pago de aguinaldo",_
							  "Images/MnSection16.gif", "Reports.asp?ReportID=1101", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_Aguinaldos & ",", vbBinaryCompare) > 0),_
						Array("Acumulados anuales",_
							  "Revise el estado de los acumulados por año, genere la constancia de percepciones y deducciones anuales, realice el ajuste anual del impuesto sobre la renta y el recálculo anual de impuestos.",_
							  "Images/MnSection15.gif", "Main_ISSSTE.asp?SectionID=15", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_AcumuladosAnuales & ",", vbBinaryCompare) > 0),_
						Array("Reclamos de pago por ajustes y deducciones",_
							  "Registre los reclamos de pago por ajustes y deducciones por empleado que hayan sido evaluados y autorizados.",_
							  "Images/MnSection14.gif", "Main_ISSSTE.asp?SectionID=16", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_ReclamosDePagoPorAjustesYDeducciones & ",", vbBinaryCompare) > 0),_
						Array("Reportes",_
							  "Reportes del área de personal",_
							  "Images/MnReports.gif", "Main_ISSSTE.asp?SectionID=19", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_Reportes & ",", vbBinaryCompare) > 0),_
						Array("Catálogos",_
							  "Altas, bajas y cambios de registros concernientes a los registros del sistema.",_
							  "Images/MnHumanResources.gif", "Catalogs.asp", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_Catalogos & ",", vbBinaryCompare) > 0),_
						Array("Tablero de control de procesos",_
							  "Administre el estatus de los procesos registrados en el sistema.",_
							  "Images/MnSection63.gif", "TaCo.asp", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_TableroDeControl & ",", vbBinaryCompare) > 0),_
						Array("Ventanilla única",_
							  "Control de los trámite que ingresan a la Subdirección de Personal.",_
							  "Images/MnSection61.gif", "Main_ISSSTE.asp?SectionID=61", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_VentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("Acreedores de los empleados",_
							  "Registre y valide los acreedores de los empleados.",_
							  "Images/MnSection24.gif", "Main_ISSSTE.asp?SectionID=27", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_AcreedoresDeLosEmpleados & ",", vbBinaryCompare) > 0)_
					)
				Case 15 'Personal > Acumulados anuales
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Acumulados anuales",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1151", True),_
						Array("Registro de empleados que no desean el ajuste anual del impuesto sobre la renta",_
							  "",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=151", True),_
						Array("Ajuste anual del impuesto sobre la renta",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1153", True),_
						Array("Aplicación del ajuste anual del impuesto sobre la renta",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1155", True),_
						Array("Aplicación por empleado del ajuste anual del impuesto sobre la renta",_
							  "",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=156", False),_
						Array("Recálculo anual de impuestos",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1154", True),_
						Array("Declaración informativa múltiple (DIM)",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1157", True),_
						Array("Constancia de percepciones y deducciones anuales",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1152", True)_
					)
				Case 151 'Personal > Acumulados anuales > Registro de empleados que no desean el ajuste anual del impuesto sobre la renta
					If Len(oRequest("ActivateTax").Item) > 0 Then
						lErrorNumber = ActivateEmployeeTaxAdjustment(oRequest, oADODBConnection, sErrorDescription)
						If lErrorNumber = 0 Then
							Call DisplayErrorMessage("Confirmación", "El ajuste anual del impuesto sobre la renta se aplicó correctamente para el empleado.")
							Response.Write "<BR />"
						Else
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							Response.Write "<BR />"
							lErrorNumber = 0
							sErrorDescription = ""
						End If
					End If
					Call DisplayEmployeeTaxActivationForm(oRequest, oADODBConnection, sErrorDescription)
				Case 156 'Personal > Acumulados anuales > Aplicación por empleado del ajuste anual del impuesto sobre la renta
					If Len(oRequest("ApplyTax").Item) > 0 Then
						lErrorNumber = ApplyEmployeeTaxAdjustment(oRequest, oADODBConnection, sErrorDescription)
						If lErrorNumber = 0 Then
							Call DisplayErrorMessage("Confirmación", "El estatus de la aplicación del ajuste anual del impuesto sobre la renta se actualizó correctamente para el empleado.")
							Response.Write "<BR />"
						Else
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							Response.Write "<BR />"
							lErrorNumber = 0
							sErrorDescription = ""
						End If
					End If
					Call DisplayEmployeeTaxAdjustmentForm(oRequest, oADODBConnection, sErrorDescription)
				Case 16, 429 'Personal > Reclamos de pago por ajustes y deducciones | Informática > Empleados > Reclamos de pago por ajustes y deducciones
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("C9. Devoluciones no gravables",_
							  "Registre los reclamos de pago por ajustes y deducciones por empleado que hayan sido evaluados y autorizados.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=-89", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_C9_DevolucionesNoGravables & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_C9_DevolucionesNoGravables & ",", vbBinaryCompare) > 0),_
						Array("71. Deducción por cobro de sueldos indebidos",_
							  "Registre los reclamos de pago por ajustes y deducciones por empleado que hayan sido evaluados y autorizados.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=-79", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_71_DeduccionPorCobroDeSueldosIndebidos & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_71_DeduccionPorCobroDeSueldosIndebidos & ",", vbBinaryCompare) > 0),_
						Array("72. Otras deducciones",_
							  "Registre los reclamos de pago por ajustes y deducciones por empleado que hayan sido evaluados y autorizados.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=-80", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_72_OtrasDeducciones & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_72_OtrasDeducciones & ",", vbBinaryCompare) > 0),_
						Array("Registro de reclamos",_
							  "Registre los reclamos de pago por empleado que hayan sido evaluados y autorizados.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=-58", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_RegistroDeReclamos & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_RegistroDeReclamos & ",", vbBinaryCompare) > 0),_
						Array("Revisión de nóminas",_
							  "Indique a qué empleados se les realizará una revisión en sus nóminas.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=PayrollRevision", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_RevisionDeNominas & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_RevisionDeNominasReclamos & ",", vbBinaryCompare) > 0),_
						Array("Baja de registros vigentes",_
							  "Cancele el concepto C9, concepto 71 o concepto 72 de los empleados.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&SubSectionID=2&ReasonID=" & CANCEL_EMPLOYEES_C04, (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaDeRegistrosVigentesReclamos & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_BajaDeRegistrosVigentesReclamos & ",", vbBinaryCompare) > 0)_
					)
				Case 18, 712 'Personal > Administración de personal | Desconcentrados > Personal > Administración de personal
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("101 Nuevo ingreso",_
							  "Registre a un nuevo empleado.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=12", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_101NuevoIngreso & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("Alta de honorarios",_
							  "Registre a un nuevo empleado bajo el régimen de honorarios.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=14", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_AltaDeHonorarios & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("103 Alta por interinato",_
							  "Registre la plaza del empleado que cubrirá un interinato.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=13", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_103_AltaPorInterinato & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("106 Alta por reingreso",_
							  "Registre la plaza que ocupará el empleado en su reingreso.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=17", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_106_AltaPorReingreso & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("106 Alta para ocupar puesto de confianza",_
							  "Registre la plaza que ocupará el empleado de licencia sin sueldo dentro del Instituto.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=68", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_106_AltaParaOcuparPuestoDeConfianza & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("107 Alta por reinstalación",_
							  "Registre la plaza que ocupará el empleado en su reinstalación.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=18", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_107_AltaPorReinstalacion & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("130 Alta por reanudación de labores",_
							  "Registre la reanudación de la licencias otorgada al empleados.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=28", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_130_AltaPorReanudacionDeLabores & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("109 Retorno a la vida laboral",_
							  "Registre el retorno del empleado.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=82", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_109_RetornoALaVidaLaboral & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("210 Cambio de plaza misma adscripción",_
							  "Registre el cambio de plaza para un empleado.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=21", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_210_CambioDePlazaMismaAdscripcion & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("212 Cambio de adscripción con plaza / 219 Cambio de servicio / 219 Cambio de turno",_
							  "Registre el movimiento en la plaza para el cambio de adscripción, servicio y/o turno.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=51", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_212_CambioDeAdscripcionConPlaza & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("213 Cambio de adscripción sin plaza",_
							  "Registre la plaza que ocupará el empleado para su cambio de adscripción.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=50", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_213_CambioDeAdscripcionSinPlaza & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("218. Permuta de plazas",_
							  "Registre el cambio de plazas interlazadas.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=26", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_220_PermutaDePlazas & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("220 Cambio de datos del empleado",_
							  "Registre los cambios de datos del empleado.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=57", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_CambioDeDatosDelEmpleado & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("000 Reasignación de número de empleado",_
							  "Reasigne un número de empleado que no ha ingresado al Instituto.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=58", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_000_ReasignacionDeNumeroDeEmpleado & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("222 Inclusión de Riesgos profesionales",_
							  "Agregue o modifique a los empleados que se les otorga una cantidad adicional por laborar en áreas infectocontagiosas o radioactivas.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=53", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_220_InclusionDeRiesgosProfesionales & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("221. Turno opcional (Concepto 07)",_
							  "Agregue o modifique los empleados de base que tienen turno opcional.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=-64", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_220_TurnoOpcional_Concepto_07 & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("221 Percepción adicional (Concepto 08)",_
							  "Agregue o modifique los empleados de confianza que tienen percepción adicional.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=-75", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_220_PercepcionAdicional_Concepto_08 & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("Cambio de Honorarios (Concepto 11)",_
							  "Modifique el importe de la percepción por Honorarios.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=-105", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_CambioDeHonorarios & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("Baja de registros vigentes",_
							  "Cancele el concepto 04, concepto 07 o concepto 08 de los empleados.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&SubSectionID=1&ReasonID=" & CANCEL_EMPLOYEES_C04, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaDeRegistrosVigentes & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("441 Licencia con goce de sueldo por Comisión sindical",_
							  "Registre la licencia con sueldo por comisión sindical.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=29", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_441_LicenciaCGSPorComisionSindical & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("446 Licencia con goce de sueldo por trámite de pensión",_
							  "Registre la licencia con sueldo por trámite de pensión.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=30", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_446_LicenciaCGSPorTramiteDePension & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("448 Licencia con goce de sueldo por contraer matrimonio",_
							  "Registre la licencia con goce de sueldo por contraer matrimonio.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=31", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_448_LicenciaCGSPorContraerMatrimonio & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("449 Licencia con goce de sueldo por fallecimiento de familiar en primer grado",_
							  "Registre la licencia con goce de sueldo por fallecimiento de familiar en primer grado.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=32", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_449_LicenciaCGSPorFallecimientoDeFamiliarEnPrimerGrado & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("451 Licencia con goce de sueldo por otorgamiento de beca",_
							  "Registre la licencia con goce de sueldo por otorgamiento de beca.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=33", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_451_LicenciaCGSPorOtorgamientoDeBeca & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("452 Licencia con goce de sueldo por práctica de servicio social",_
							  "Registre la licencia con goce de sueldo por práctica de servicio social.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=34", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_452_LicenciaCGSPorPracticaDeServicioSocial & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("592 Licencia sin goce de sueldo por asuntos particulares",_
							  "Registre la licencia sin goce de sueldo por asuntos particulares.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=43", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_592_LicenciaSGSPorAsuntosParticulares & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("593 Licencia sin goce de sueldo por comisión sindical",_
							  "Registre la licencia sin goce de sueldo por asuntos particulares.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=44", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_593_LicenciaSGSPorComisionSindical & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("594 Licencia sin goce de sueldo por otorgamiento de beca",_
							  "Registre la licencia con goce de sueldo por otorgamiento de beca.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=45", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_594_LicenciaSGSPorOtorgamientoDeBeca & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("595 Licencia sin goce de sueldo por ocupar cargo de elección popular o puesto de confianza fuera del instituto",_
							  "Registre la licencia sin goce de sueldo por ocupar cargo de elección popular o puesto de confianza fuera del Instituto.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=46", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_595_LicenciaSGSPorOcuparCargoDeEleccionPopular & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("596 Licencia sin goce de sueldo por ocupar puesto de confianza dentro del Instituto",_
							  "Registre la licencia sin goce de sueldo por ocupar puesto de confianza dentro del Instituto.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=47", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_596_LicenciaSGSPorOcuparPuestoDeConfianzaDentroDelInstituto & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("597 Licencia sin goce de sueldo por práctica de servicio social",_
							  "Registre la licencia sin goce de sueldo por práctica de servicio social.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=48", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_597_LicenciaSGSPorPracticaDeServicioSocial & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("468 Prórroga con goce de sueldo por comision",_
							  "Registre la prórroga con goce de sueldo por comisión.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=35", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_468_ProrrogaDeLicenciaConSueldoPorComisionsindical & ",", vbBinaryCompare) > 0),_
							  
						Array("469 Prórroga de licencia con goce de sueldo por otorgamiento de beca",_
							  "Registre la prórroga de licencia con goce de sueldo por otorgamiento de beca.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=36", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_469_ProrrogaDeLicenciaCGSPorOtorgamientoDeBeca & ",", vbBinaryCompare) > 0),_
						Array("570 Prórroga de licencia sin goce de sueldo por comisión sindical",_
							  "Registre la licencia sin goce de sueldo por comisión sindical.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=37", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_570_ProrrogaDeLicenciaSGSPorComisionSindical & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("571 Prórroga de licencia sin goce de sueldo por otorgamiento de beca",_
							  "Registre la licencia sin goce de sueldo por otorgamiento de beca.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=38", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_571_ProrrogaDeLicenciaSGSPorOtorgamientoDeBeca & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("572 Prórroga de licencia sin goce de sueldo por ocupar cargo de elección popular o puesto de confianza fuera del Instituto",_
							  "Registre la prórroga de licencia sin goce de sueldo por ocupar cargo de elección popular o puesto de confianza fuera del Instituto.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=39", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_572_ProrrogaDeLicenciaSGSPorOcuparCargoDeEleccionPopular & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("573 Prórroga de licencia sin goce de sueldo por ocupar puesto de confianza dentro del Instituto",_
							  "Registre la prórroga de licencia sin goce de sueldo por ocupar puesto de confianza dentro del Instituto.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=40", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_573_ProrrogaDeLicenciaSGSPorOcuparPuestoDeConfianzaDentroDelInstituto & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("574 Prórroga de licencia sin goce de sueldo por asuntos particulares",_
							  "Registre la prórroga de licencia sin goce de sueldo por asuntos particulares.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=41", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_574_ProrrogaDeLicenciaSGSPorAsuntosParticulares & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("310 Baja de personal de honorarios",_
							  "Registre la baja de personal de honorarios).",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=66", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_310_BajaDePersonalDeHonorarios & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("340 Baja por renuncia",_
							  "Registre baja por renuncia).",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=1", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_340_BajaPorRenuncia & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("341 Baja por defunción",_
							  "Registre baja por defunción).",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=2", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_341_BajaPorDefuncion & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("342 Baja por cese",_
							  "Registre baja por cese.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=3", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_342_BajaPorCese & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("343 Baja por incapacidad total y permanente",_
							  "Registre baja por incapacidad total y permanente).",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=4", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_343_BajaPorIncapacidadTotalYPermanente & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("344 Baja por pensión",_
							  "Registre baja por pensión).",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=5", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_344_BajaPorPension & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("345 Baja por jubilación",_
							  "Registre baja por jubilación).",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=6", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_345_BajaPorJubilacion & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("350 Baja por interinato",_
							  "Registre baja por interinato).",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=10", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_346_BajaPorInterinato & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("348 Baja por término al puesto de confianza",_
							  "Registre la baja al puesto de confianza.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=8", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_348_BajaPorTerminoAlPuestoDeConfianza & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("349 Baja por sanción administrativa",_
							  "Registre baja por sanción administrativa.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=62", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_349_BajaPorSancionAdministrativa & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("350 Baja por sanción",_
							  "Registre baja por sanción.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=63", False And StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_350_BajaPorSancion & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("Baja por término de nombramiento",_
							  "Registre baja por término de nombramiento.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=7", False And StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorTerminoDeNombramiento & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("Baja por término de provisionalidad",_
							  "Registre la baja por término de provisionalidad.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=78", False And StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorTerminoDeProvisionalidad & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("346 Término de convenio",_
							  "Registre la baja por término por convenio.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=79", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_TerminoDeConvenio & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("347 Baja por no tomar posesión del puesto",_
							  "Registre la baja por no tomar posesión del puesto.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=80", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorNoTomarPosesionDelPuesto & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("348 Baja por incurrir en falta administrativa antes de adquirir la inamovilidad",_
							  "Registre la baja por incurrir en falta administrativa antes de adquirir la inamovilidad.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=81", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorIncurrirEnFaltaAdministrativa & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("380 Baja por transición de la terminación de la relación laboral con el Instituto",_
							  "Registre la baja de la terminación de la relación laboral con el Instituto.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=107", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_380_BajaPorTransicionDLaTerminacionDlaRelacionLaboralConInstituto & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("370 Baja por Renuncia al Instituto con estatus de licencia sin goce de sueldo",_
							  "Registre Baja por Renuncia al Instituto con estatus de licencia sin goce de sueldo.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=101", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_370_BajaPorRenunciaEnLSGS & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("371 Baja por Defuncion con estatus de licencia sin goce de sueldo",_
							  "Registre Baja por Defuncion con estatus de licencia sin goce de sueldo.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=102", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_371_BajaPorDefuncionEnLSGS & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("372 Baja por Cese con estatus de licencia sin goce de sueldo",_
							  "Registre Baja por Cese con estatus de licencia sin goce de sueldo.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=103", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_372_BajaPorCeseEnLSGS & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("373 Baja por Incapacidad total y permanente o invalidez, con estatus de licencia sin goce de sueldo",_
							  "Registre Baja por Incapacidad total y permanente o invalidez, con estatus de licencia sin goce de sueldo.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=104", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_373_BajaPorIncapacidadTotalYPermanenteEnLSGS & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("374 Baja por Pensión en estatus de licencia sin goce de sueldo",_
							  "Registre Baja por Pensión en estatus de licencia sin goce de sueldo.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=105", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_374_BajaPorPensionEnLSGS & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("375 Baja por Jubilacion en estatus de licencia sin goce de sueldo",_
							  "Registre Baja por Jubilacion en estatus de licencia sin goce de sueldo.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=106", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_375_BajaPorJubilacionEnLSGS & ",", vbBinaryCompare) > 0 Or aLoginComponent(N_PROFILE_ID_LOGIN) = 7),_
						Array("",_
							  "",_
							  "", "", False)_
					)
				Case 19, 713 'Personal > Reportes | Desconcentrados > Personal > Reportes
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Reportes guardados",_
							  "Reportes generados por otros usuarios a partir de plantillas y filtros.",_
							  "Images/MnLeftArrows.gif", "SavedReport.asp?ReportType=1", (CInt(Request.Cookies("SIAP_SectionID")) = 1)),_
						Array("<TITLE />REPORTES SOBRE EMPLEADOS",_
							  "",_
							  "", "", True),_
						Array("Pagos cancelados",_
							  "Obtenga el listado de los pagos cancelados por empleado",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1114" & sSubSectionID, True),_
						Array("Conteo de empleados",_
							  "Obtenga un conteo de los empleados que se encuentran registrados en el sistema, agrupando los resultados por diferentes conceptos.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=702" & sSubSectionID, True),_
						Array("Hoja única de servicio",_
							  "Histórico del empleado a lo largo de los años de servicio en el Instituto.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1112" & sSubSectionID, False),_
						Array("Información de los empleados",_
							  "Obtenga un listado de los empleados registrados en el sistema incluyendo la información que usted necesite consultar.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=705" & sSubSectionID, True),_
						Array("Nómina de personal correspondiente a conceptos",_
							  "Obtenga un archivo zip con el listado de los empleados que cobraron el concepto seleccionado en la quincena seleccionada.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1107" & sSubSectionID, True),_
						Array("Plantilla de nómina",_
							  "Listado de la plantilla de nómina",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1311" & sSubSectionID, True),_
						Array("Reporte de honorarios",_
							  "Listado del personal de honorarios",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1106" & sSubSectionID, True),_
						Array("Reporte de personal con conceptos",_
							  "Obtenga un archivo zip con el listado de los empleados que tienen registrado el concepto seleccionado.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1108" & sSubSectionID, True),_
						Array("Reporte de validación del pago de aguinaldo",_
							  "Obtenga el listado con la información con que cuenta el sistema de nómina para generar el pago de aguinaldo.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1113" & sSubSectionID, True),_
						Array("<TITLE />REPORTES SOBRE MOVIMIENTOS",_
							  "",_
							  "", "", True),_
						Array("Impresión del formato FM1",_
							  "Obtenga un archivo zip con el formato FM1 de los movimientos en trámite.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1109" & sSubSectionID, True),_
						Array("Impresión del formato de baja honorarios",_
							  "Obtenga un archivo zip con el formato de honorarios de los movimientos en trámite.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1115" & sSubSectionID, True),_
						Array("Impresión del formato de honorarios",_
							  "Obtenga un archivo zip con el formato de honorarios de los movimientos en trámite.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1110" & sSubSectionID, True),_
						Array("Reporte de movimientos en trámite",_
							  "Listado de los movimientos que tuvieron los empleados en el período especificado o por empleado una vez que fue procesada la nómina.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1103" & sSubSectionID, True),_
						Array("Reporte de reclamos de pago por ajustes y deducciones",_
							  "Listado de los reclamos de pago por ajustes y deducciones que se han registrado a los empleado",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1102" & sSubSectionID, True),_
						Array("Reporte de registro de movimientos",_
							  "Listado de los movimientos registrados para aplicación en la quincena seleccionada.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1105" & sSubSectionID, True),_
						Array("Reporte de movimientos por usuario",_
							  "Listado de los movimientos que realizó cada usuario",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1104" & sSubSectionID, True),_  
						Array("<TITLE />REPORTES SOBRE PLAZAS",_
							  "",_
							  "", "", True),_
						Array("Histórico de plazas",_
							  "Histórico de las plazas incluyendo el número de los empleados que las han ocupado.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1111" & sSubSectionID, True),_
                        Array("<TITLE />REPORTES",_
							  "",_
							  "", "", True),_
                        Array("Ejercicios Fiscales",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1120" & sSubSectionID, True),_
						Array("",_
							  "",_
							  "", "", False)_                        
					)
				Case 2 'Prestaciones
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Consulta de personal",_
							  "Consulte la información de los empleados, plaza, conceptos de pago, historia.",_
							  "Images/MnEmployees.gif", "Employees.asp", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ConsultaDePersonal & ",", vbBinaryCompare) > 0),_
						Array("SI. Seguro de separación y AE. Seguro adicional de separación individualizado",_
							  "Administre el seguro de separación y seguro adicional para personal de mando medio.",_
							  "Images/MnSection17.gif", "Main_ISSSTE.asp?SectionID=20", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_SI_SeguroDeSeparacionYAE_SeguroAdicionalDeSeparacion & ",", vbBinaryCompare) > 0),_
						Array("Certificaciones y archivo",_
							  "Administre las antigüedades, hojas de servicio, constancias para los empleados.",_
							  "Images/MnSection22.gif", "Main_ISSSTE.asp?SectionID=25", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_CertificacionesYArchivo & ",", vbBinaryCompare) > 0),_
						Array("Terceros institucionales",_
							  "Administre información de terceros y genere archivos.",_
							  "Images/MnSection21.gif", "Main_ISSSTE.asp?SectionID=21", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_TercerosInstitucionales & ",", vbBinaryCompare) > 0),_
						Array("Prestaciones e incidencias",_
							  "Registre las prestaciones e incidencias a los empleados.",_
							  "Images/MnSearch.gif", "Main_ISSSTE.asp?SectionID=22", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_PrestacionesEIncidencias & ",", vbBinaryCompare) > 0),_
						Array("Baja de prestaciones vigentes",_
							  "Cancele las prestaciones e incidencias a los empleados.",_
							  "Images/MnSection18.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & CANCEL_EMPLOYEES_CONCEPTS, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_BajaDePrestacionesVigentes & ",", vbBinaryCompare) > 0),_
						Array("Antigüedades",_
							  "Administre las antigüedades y constacias relacionadas.",_
							  "Images/MnSection23.gif", "Main_ISSSTE.asp?SectionID=26", False And StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorIncurrirEnFaltaAdministrativa & ",", vbBinaryCompare) > 0),_
						Array("Pensión alimenticia",_
							  "Registre y valide las pensiones alimenticias de los empleados.",_
							  "Images/MnSection24.gif", "Main_ISSSTE.asp?SectionID=23", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_PensionAlimenticia & ",", vbBinaryCompare) > 0),_
						Array("Acreedores de los empleados",_
							  "Registre y valide los acreedores de los empleados.",_
							  "Images/MnSection24.gif", "Main_ISSSTE.asp?SectionID=27", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AcreedoresDeLosEmpleados & ",", vbBinaryCompare) > 0),_
						Array("Fondo de ahorro capitalizable (FONAC)",_
							  "Administre los procesos y liquidaciones del FONAC.",_
							  "Images/MnSection25.gif", "xxx.asp", False And StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorIncurrirEnFaltaAdministrativa & ",", vbBinaryCompare) > 0),_
						Array("Sistema de ahorro para el retiro",_
							  "Administración del SAR.",_
							  "Images/MnSection26.gif", "xxx.asp", False And StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorIncurrirEnFaltaAdministrativa & ",", vbBinaryCompare) > 0),_
						Array("Reportes",_
							  "Reportes del área de prestaciones",_
							  "Images/MnReports.gif", "Main_ISSSTE.asp?SectionID=24", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_Reportes & ",", vbBinaryCompare) > 0),_
						Array("Catálogos",_
							  "Altas, bajas y cambios de registros concernientes a los registros del sistema.",_
							  "Images/MnHumanResources.gif", "Catalogs.asp", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_Catalogos & ",", vbBinaryCompare) > 0),_
						Array("Ventanilla única",_
							  "Control de los trámite que ingresan a la Subdirección de Personal.",_
							  "Images/MnSection61.gif", "Main_ISSSTE.asp?SectionID=61", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_VentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("Registro de bolsa de trabajo y escalafón",_
							  "La información registrada en esta sección será consultada a través del módulo de Desarrollo Humano.",_
							  "Images/MnSection34.gif", "Main_ISSSTE.asp?SectionID=28", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeBolsaDeTrabajoEscalafon & ",", vbBinaryCompare) > 0),_
						Array("Matriz de riesgos profesionales",_
							  "Administre la matriz de Puestos y Servicios sujetos al pago de Riesgos Profesionales.",_
							  "Images/MnPayments.gif", "Main_ISSSTE.asp?SectionID=29", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_MatrizDeRiesgosProfesionales & ",", vbBinaryCompare) > 0),_
						Array("Acumulados anuales",_
							  "Revise el estado de los acumulados por año, genere la constancia de percepciones y deducciones anuales, realice el ajuste anual del impuesto sobre la renta y el recálculo anual de impuestos.",_
							  "Images/MnSection15.gif", "Main_ISSSTE.asp?SectionID=291", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AcumuladosAnuales & ",", vbBinaryCompare) > 0)_
					)
				Case 20 'Prestaciones > Seguro de separación
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("SI. Seguro de separación individualizado",_
							  "Administre el seguro de separación para personal de mando medio.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=-61", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_SI_SeguroDeSeparacionIndividualizado & ",", vbBinaryCompare) > 0),_
						Array("AE. Seguro adicional de separación individualizado",_
							  "Administre el seguro adicional para personal de mando medio.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=-62", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AE_SeguroAdicionalDeSeparacionIndividualizado & ",", vbBinaryCompare) > 0),_
						Array("Baja de registros vigentes",_
							  "Cancele el seguros de separación y seguro adicional de los empleados.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID="& CANCEL_EMPLOYEES_SSI, (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0) Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_BajaDeRegistrosVigentesSIyAE & ",", vbBinaryCompare) > 0)_
					)
				Case 21 'Prestaciones > Terceros institucionales
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Carga de discos de terceros",_
							  "Registrar los importes a los empleados mediante el archivo que los terceros entregan",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=211", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_CargaDeDiscosDeTerceros & ",", vbBinaryCompare) > 0),_
						Array("Aplicación de registros cargados por cada archivo",_
							  "Realice una validación general seleccionando cada archivo cargado y aplique los movimientos",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=ThirdUploadMovements&ReasonID="&EMPLOYEES_THIRD_PROCESS, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AplicacionDeRegistrosCargadosPorCadaArchivo & ",", vbBinaryCompare) > 0),_
						Array("Registro en línea de terceros institucionales",_
							  "Registrar los importes de terceros a los empleados que tuvieron error en el archivo",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID="&EMPLOYEES_THIRD_CONCEPT, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroEnLineaDeTercerosInstitucionales & ",", vbBinaryCompare) > 0),_
						Array("Reporte de registros cargados desde archivo de terceros",_
							  "Obtenga un archivo zip con el listado de los registros agregados por medio del archivo de un tercero.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1221", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ReporteDeCargaDesdeArchivosDeTerceros & ",", vbBinaryCompare) > 0),_
						Array("Reporte de registros rechazados desde archivos de terceros",_
							  "Obtenga un archivo zip con el listado de los registros que fueron rechazados del archivo de un tercero.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1222", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ReporteDeCargaDesdeArchivosDeTerceros & ",", vbBinaryCompare) > 0),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Generación de archivo Repcsi",_
							  "Generar los archivos que serán enviados a los terceros",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1225", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_GeneracionDeArchivoRepcsi & ",", vbBinaryCompare) > 0),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Registro de calificación de empleados",_
							  "Registre a la calificación a los empleados que la requieran.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_GRADE, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeCalificacionDeEmpleados & ",", vbBinaryCompare) > 0),_
						Array("Reporte de calificación de empleados",_
							  "Revise las calificaciones otorgadas a los empleados para el pago del concepto 28. Estímulo a la productividad, calidad y eficacia para personal médico y de enfermería",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1211" & sSubSectionID, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ReporteDeCalificacionDeEmpleados & ",", vbBinaryCompare) > 0)_
					)
				Case 211 'Prestaciones > Terceros institucionales > Carga de discos de terceros
					Dim iIndex
					Dim sMenuData
					Dim aMenu
					Dim oRecordset

					sErrorDescription = "No se pudo obtener la información de los registros."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From CreditTypes Where (IsOther=1) And (Active=1) And (CreditTypeID>0) Order By CreditTypeName", "Main_ISSSTE.asp", "_root", 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							iIndex = 2
							Do While Not oRecordset.EOF
								iIndex = iIndex + 1
								oRecordset.MoveNext
								If Err.number <> 0 Then Exit Do
							Loop
							ReDim aMenu(iIndex)
							iIndex = 7
							oRecordset.MoveFirst
							sMenuData = "ISSSTE. Préstamos" & LIST_SEPARATOR & "Registre los conceptos <B>60</B> y <B>85</B> incluidos en el archivo Issste por prestamos personales" & LIST_SEPARATOR & "Images/MnLeftArrows.gif" & LIST_SEPARATOR & "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_ISSSTE_CONCEPT & LIST_SEPARATOR & "-1"
							aMenu(0) = Split(sMenuData, LIST_SEPARATOR, -1, vbBinaryCompare)
							sMenuData = "<LINE />" & LIST_SEPARATOR & LIST_SEPARATOR & LIST_SEPARATOR & LIST_SEPARATOR & "-1"
							aMenu(1) = Split(sMenuData, LIST_SEPARATOR, -1, vbBinaryCompare)
							sMenuData = "FOVISSSTE, Crédito hipotecario" & LIST_SEPARATOR & "Registre el concepto <B>62</B> reportados dentro del archivo FOVISSSTE" & LIST_SEPARATOR & "Images/MnLeftArrows.gif" & LIST_SEPARATOR & "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_FOVISSSTE_CONCEPT_62 & LIST_SEPARATOR & "-1"
							aMenu(2) = Split(sMenuData, LIST_SEPARATOR, -1, vbBinaryCompare)
							sMenuData = "FOVISSSTE, Crédito hipotecario" & LIST_SEPARATOR & "Registre los conceptos <B>86</B> reportados dentro del archivo FOVISSSTE" & LIST_SEPARATOR & "Images/MnLeftArrows.gif" & LIST_SEPARATOR & "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_FOVISSSTE_CONCEPT_86 & LIST_SEPARATOR & "-1"
							aMenu(3) = Split(sMenuData, LIST_SEPARATOR, -1, vbBinaryCompare)
							sMenuData = "FOVISSSTE, Crédito hipotecario" & LIST_SEPARATOR & "Registre los conceptos <B>56</B> reportados dentro del archivo FOVISSSTE" & LIST_SEPARATOR & "Images/MnLeftArrows.gif" & LIST_SEPARATOR & "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_FOVISSSTE_CONCEPT_56 & LIST_SEPARATOR & "-1"
							aMenu(4) = Split(sMenuData, LIST_SEPARATOR, -1, vbBinaryCompare)
							sMenuData = "FOVISSSTE, Crédito hipotecario" & LIST_SEPARATOR & "Registre los conceptos <B>NF</B> reportados dentro del archivo FOVISSSTE" & LIST_SEPARATOR & "Images/MnLeftArrows.gif" & LIST_SEPARATOR & "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_FOVISSSTE_CONCEPT_NF & LIST_SEPARATOR & "-1"
							aMenu(5) = Split(sMenuData, LIST_SEPARATOR, -1, vbBinaryCompare)
							sMenuData = "<LINE />" & LIST_SEPARATOR & LIST_SEPARATOR & LIST_SEPARATOR & LIST_SEPARATOR & "-1"
							aMenu(6) = Split(sMenuData, LIST_SEPARATOR, -1, vbBinaryCompare)
							If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS) = N_ADD_PERMISSIONS) Then
								Do While Not oRecordset.EOF
									sMenuData = CleanStringForHTML(CStr(oRecordset.Fields("CreditTypeShortName").Value)) & ". " & CleanStringForHTML(CStr(oRecordset.Fields("CreditTypeName").Value)) & LIST_SEPARATOR
									sMenuData = sMenuData & LIST_SEPARATOR & "Images/MnLeftArrows.gif" & LIST_SEPARATOR & _
													  "UploadInfo.asp?Action=Third&ThirdConcept=" & CStr(oRecordset.Fields("CreditTypeShortName").Value)
									sMenuData = sMenuData & LIST_SEPARATOR & "-1"
									aMenu(iIndex) = Split(sMenuData, LIST_SEPARATOR, -1, vbBinaryCompare)
									iIndex = iIndex + 1
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
							End If
						End If
						oRecordset.Close
						aMenuComponent(A_ELEMENTS_MENU) = aMenu
					End If
					If False Then
						aMenuComponent(A_ELEMENTS_MENU) = Array(_
							Array("Seguros Metlife",_
								  "Registre los conceptos <B>63</B> y <B>64</B> proporcionados por Metlife",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_AHISA_CONCEPT, True),_
							Array("AHORRA. Crédito ahorra Ya",_
								  "Cargue el concepto <B>88</B> para los empleados, reportados en el archivo Ahorra por compras de electrodomesticos",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_AHORRA_CONCEPT, True),_
							Array("AXXA. Seguros",_
								  "Cargue el concepto <B>75</B> para los empleados, reportados en el archivo Asemex",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_ASEMEX_CONCEPT, True),_
							Array("CARDINAL. Seguro de vida interamericana",_
								  "Cargar los registros de seguros proporcionados en el archivo Cardinal, concepto <B>87</B>",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_CARDINAL_CONCEPT, True),_
							Array("ETESA. Compra de electrodomesticos",_
								  "Registre el concepto <B>ET</B> para los empleados reportados en el archivo Etesa por compras de electrodomesticos",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_ETESA_CONCEPT, True),_
							Array("ISSSTE. Préstamos",_
								  "Registre los conceptos <B>60</B> y <B>85</B> incluidos en el archivo Issste por prestamos personales",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_ISSSTE_CONCEPT, True),_
							Array("NASER. Servicios funerarios",_
								  "Cargar los registros del concepto <B>81</B> por servicios funerarios proporcionados en el archivo Naser",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_NASER_CONCEPT, True),_
							Array("Previsora plenitud. Servicios funerarios",_
								  "Cargar los registros del concepto <B>PP</B> de servicios funerarios proporcionados en el archivo Naser",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept=Previsora", True),_
							Array("SERFUN. Servicios funerarios",_
								  "Cargar los registros de servicios funerarios, con clave <B>MT</B> proporcionados en el archivo del tercero",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_SERFUN_CONCEPT, True),_
							Array("OPTICA. Tratamientos ópticos.",_
								  "Cargar los registros del concepto <B>D3</B> de tratamientos ópticos de los empleados proporcionados en el archivo Optica",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_OPTICA_CONCEPT, True),_
							Array("FOVISSSTE, Crédito hipotecario",_
								  "Registre los conceptos <B>55, 56, 62</B> y <B>86</B> reportados dentro del archivo FOVISSSTE",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_FOVISSSTE_CONCEPT, True),_
							Array("65. Baja de seguro institucional",_
								  "Registre el concepto por baja de seguro institucional.",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&SubSectionID=211&ReasonID=" & EMPLOYEES_SAFEDOWN, True),_
							Array("67. Cuota Deportivo",_
								  "Registre a los empleados inscritos al deportivo.",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_SPORTS, True),_
							Array("83. Seguro Auto Grupo Nacional Provincial",_
								  "Cargar los registros del seguro de Auto Grupo Nacional Provincial con clave <B>83</B>",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_GNP_CONCEPT, True),_
							Array("87. Seguro de Vida Interamericana",_
								  "",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=87", False),_
							Array("D5. Seguro de Responsabilidad Civil para Médicos y Enfermeras",_
								  "",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=D5", False),_
							Array("CS. Sindicato Independiente",_
								  "",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=CS", False),_
							Array("MT. Martínez Servicios Funerarios a Futuro",_
								  "",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=MT", False),_
							Array("SR. Seguro de Responsabilidad Civil para Personal de Mando",_
								  "",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=SR", False),_
							Array("FR. Reestructuración Financiera",_
								  "",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=FR", False)_
						)
					End If
				Case 212 'Prestaciones > Terceros institucionales > Generación de archivos para enteros
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("56 y 60. AHISA, seguro Metlife",_
							  "Cargar los registros de seguros proporcionados por Metlife",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_AHISA_CONCEPT, True),_
						Array("88. AHORRA. Crédito ahorra Ya",_
							  "Cargar los registros reportados por AHORRA por compras de electrodomesticos",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_AHORRA_CONCEPT, True),_
						Array("75. ASEMEX. Seguros",_
							  "Cargar los registros de seguros proporcionados por Asemex",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_ASEMEX_CONCEPT, True),_
						Array("CARDINAL. Compañías de seguros",_
							  "Cargar los registros de seguros proporcionados por Cardinal",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_CARDINAL_CONCEPT, True),_
						Array("ET. ETESA. Compra de electrodomesticos",_
							  "Cargar los registros reportados en Etesa por compras de electrodomesticos",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_ETESA_CONCEPT, True),_
						Array("60 y 85. ISSSTE. Préstamos",_
							  "Cargar los registros por prestamos ",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_ISSSTE_CONCEPT, True),_
						Array("81. NASER. Servicios funerarios",_
							  "Cargar los registros de servicios funerarios proporcionados en Naser",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_NASER_CONCEPT, True),_
						Array("PP. Previsora plenitud. Servicios funerarios",_
							  "Cargar los registros de servicios funerarios proporcionados en Naser",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept=Plenitud", True),_
						Array("SERFUN. Servicios funerarios",_
							  "Cargar los registros de servicios funerarios proporcionados en Serfun",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_SERFUN_CONCEPT, True),_
						Array("D3. OPTICA. Tratamientos ópticos.",_
							  "Cargar los registros de tratamientos ópticos ",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_OPTICA_CONCEPT, True),_
						Array("55, 56, 62 y 86. FOVISSSTE, Crédito hipotecario",_
							  "Cargar los registros de créditos proporcionados por el FOVISSSTE",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Third&ThirdConcept="&THIRD_FOVISSSTE_CONCEPT, True),_
				        Array("65. Seguro Institucional",_
							  "",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=65", True),_
						Array("67. Cuota Deportivo",_
							  "",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=67", True),_
						Array("83. Seguro Auto Grupo Nacional Provincial",_
							  "",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=83", True),_
						Array("87. Seguro de Vida Interamericana",_
							  "",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=87", True),_
						Array("D5. Seguro de Responsabilidad Civil para Médicos y Enfermeras",_
							  "",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=D5", True),_
						Array("CS. Sindicato Independiente",_
							  "",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=CS", True),_
						Array("MT. Martínez Servicios Funerarios a Futuro",_
							  "",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=MT", True),_
						Array("SR. Seguro de Responsabilidad Civil para Personal de Mando",_
							  "",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=SR", True),_
						Array("FR. Reestructuración Financiera",_
							  "",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=FR", True)_
					)
				Case 213 'Prestaciones > Terceros institucionales > Reportes
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Reportes guardados",_
							  "Reportes generados por otros usuarios a partir de plantillas y filtros.",_
							  "Images/MnLeftArrows.gif", "SavedReport.asp?ReportType=1", False),_
						Array("Reporte de registros cargados desde archivo de terceros",_
							  "Obtenga un archivo zip con el listado de los registros agregados en por medio de un archivo de un tercero.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1221", True)_
					)
				Case 22, 721 'Prestaciones > Prestaciones e incidencias | 'Desconcentrados > Prestaciones > Prestaciones e incidencias
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Registro de reclamos",_
							  "Registre los reclamos de pago por empleado que hayan sido evaluados y autorizados.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=-58", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeReclamos & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("Revisión de nóminas",_
							  "Indique a qué empleados se les realizará una revisión en sus nóminas.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=PayrollRevision", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RevisionDeNominas & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("Registro de conceptos de empleado",_
							  "Seleccione algún concepto en especial para registrarlo al empleado.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesConcepts", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeConceptosDeEmpleado & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("05. Compensaciones por antigüedad",_
							  "Modifique la compensación por antigüedad que tiene actualmente el empleado.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_ANTIQUITIES, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_05_CompensacionesPorAntiguedad & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("09. Remuneración por horas extraordinarias",_
							  "Registre el tiempo extra laborado por los empleados.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_EXTRAHOURS, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_09_RemuneracionPorHorasExtraordinarias & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("14. Primas dominicales",_
							  "Registre los domingos que los empleados laboraron.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_SUNDAYS, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_14_PrimasDominicales & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("19. Becas para los hijos de los trabajadores",_
							  "Agregue o modifique las becas que los trabajadores recibirán para los estudios de sus hijos.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_CHILDREN_SCHOOLARSHIPS, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_19_BecasParaLosHijosDeLosTrabajadores & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("20. Ayuda de anteojos",_
							  "Registre a los empleados que se les ha otorgado esta prestación para la adquisición de anteojos o lentes de contacto, por prescripción del médico oftalmólogo del ISSSTE.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_GLASSES, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_20_AyudaDeAnteojos & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("28. Estímulo a la productividad, calidad y eficacia para personal médico y de enfermería",_
							  "Registre al personal médico, de odontología y enfermería que se hicieron acreedores al reconocimiento de acuerdo a su evaluación.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_ANUAL_AWARD, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_28_EstimuloALaProductividad & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("41. Premio antigüedad 25 y 30 años",_
							  "Registre al personal con antigüedad de 25 y 30 años.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_ANTIQUITY_25_AND_30_YEARS, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_41_Premio_antiguedad_25_y_30_anios & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("42. Ayuda por muerte de familiar en primer grado",_
							  "Registre al personal al que se le otorgará esta prestación a causa de la muerte de familiar en primer grado.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_FAMILY_DEATH, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_42_AyudaPorMuerteDeFamiliarEnPrimerGrado & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("43. Ayuda impresión de tesis",_
							  "Registre al personal que obtuvo su título.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_PROFESSIONAL_DEGREE, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_43_AyudaImpresionDeTesis & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("49. Premio trabajador del mes",_
							  "Registre a los empleados de base que se hacen acreedores a esta percepción de acuerdo a la evaluación de desempeño que realizan sus compañeros de trabajo.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_MONTHAWARD, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_49_PremioTrabajadorDelMes & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("22. Premio 10 de Mayo",_
							  "Registre a los empleados de base que se hacen acreedores a esta percepción.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_MOTHERAWARD, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_22_Premio10DeMayo & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("AD. Apoyo al deporte",_
							  "Registre a los empleados de base que se hacen acreedores a esta percepción por cumplir con la activación física.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_SPORTS_HELP, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_67_ApoyoDeportivo & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("67. Cuota deportivo",_
							  "Registre a los empleados inscritos al deportivo.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_SPORTS, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_67_CuotaDeportivo & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("C2. Jornada nocturna adicional por día festivo (acumulada)",_
							  "Registre a los empleados que laboran los turnos 21, 22 y 23 hayan iniciado o concluido su jornada en días de descanso obligatorio.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_NIGHTSHIFTS, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_C2_JornadaNocturnaAdicionalPorDiaFestivo & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("C3. Premios, estímulos y recompensas (recompensa del sistema de evaluación del desempeño)",_
							  "Registre al personal acreedor a esta recompensa, de acuerdo a sus aportaciones y evaluación.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_CONCEPT_C3, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_C3_PremiosEstimulosYRecompensas & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("16. Devoluciones por deducciones indebidas",_
							  "Registre el concepto para devoluciones por deducciones indebidas.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_CONCEPT_16, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_16_DevolucionesPorDeduccionesIndebidas & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("71. Deducción por cobro de sueldos indebidos",_
							  "Registre el concepto para devoluciones no excentas y reingreso de sueldo por cobro indebido.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_NON_EXCENT, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_71_DevolucionesNoExcentas & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("72. Otras deducciones",_
							  "Registre el concepto para devoluciones excentas de los empleados que las reportan.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_EXCENT, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_72_OtrasDeducciones & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("61. Indicador comisión de auxilio",_
							  "Registre el concepto por Indicador comisión de auxilio.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_HELP_COMISSION, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_61_IndicadorComisionDeAuxilio & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("65. Baja de seguro institucional",_
							  "Registre el concepto por baja de seguro institucional.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_SAFEDOWN, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_65_BajaDeSeguroColectivo & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("77. Fondo de ahorro capitalizable FONAC",_
							  "Registre el concepto por aportación al fondo de ahorro capitalizable.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_FONAC_CONCEPT, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_77_FondoDeAhorroCapitalizableFonac & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("76. Ajuste FONAC",_
							  "Registre el concepto de ajuste al fondo de ahorro capitalizable FONAC.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_FONAC_ADJUSTMENT, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_76_AjusteFonac & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("7S. Ahorro solidario",_
							  "Registre el concepto de ahorro solidario.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_CONCEPT_7S, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_7s_AhorroSolidario & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7))_
					)
				Case 23, 723 'Prestaciones > Pensión alimenticia | 'Desconcentrados > Prestaciones > Pensión alimenticia
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Padrón de pensionistas",_
							  "Registre en el padrón de pensionistas a los beneficiarios de la prestación.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_ADD_BENEFICIARIES, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDePensionAlimenticia & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("Catálogo de tipos de pensión alimenticia",_
							  "Registre los tipos de pensión alimenticia que son asignadas a los beneficiarios de la prestación.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=AlimonyTypes&ReasonID=" & ALIMONY_TYPES, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_CatalogoDeTiposDePensionAlimenticia & ",", vbBinaryCompare) > 0),_
						Array("Adeudo pensión alimenticia",_
							  "Administre el concepto de adeudo de pensión alimenticia para los empleados.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_BENEFICIARIES_DEBIT, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdeudoPensionAlimenticia & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Reporte de beneficiarios de pensiones alimenticias por empleado",_
							  "Obtenga un archivo zip con el listado de los beneficiarios de pensiones alimenticias por empleado.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1223", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ReporteDeBeneficiariosDePensionesAlimenticiasPorEmpleado & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("Reporte de empleados con pensiones alimenticias",_
							  "Obtenga un archivo zip con el listado de los empleados que tienen registradas pensiones alimenticias.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1224", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ReporteDeEmpleadosConPensionesAlimenticias & ",", vbBinaryCompare) > 0 Or (CInt(Request.Cookies("SIAP_SectionID")) = 7))_
					)
				Case 24 'Prestaciones > Reportes | Desconcentrados > Personal > Reportes
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Conteo de empleados",_
							  "Obtenga un conteo de los empleados que se encuentran registrados en el sistema, agrupando los resultados por diferentes conceptos.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=702" & sSubSectionID, True),_
						Array("Funcionarios y operativos por concepto de pago y empresa",_
							  "Obtenga un archivo zip con el listado de los empleados con conceptos.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1200", True),_
						Array("Reporte de personal con conceptos",_
							  "Obtenga un archivo zip con el listado de los empleados con conceptos y tipo de empleado",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1201", True),_
						Array("Reporte de personal con créditos",_
							  "Obtenga un archivo zip con el listado de los empleados con créditos indicando el número de cuota.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1202", True),_
						Array("Reporte de revisión de nóminas",_
							  "Obtenga un listado de los empleados registrados para revisión de nóminas anteriores.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1209", True),_
						Array("Información de los empleados",_
							  "Obtenga un listado de los empleados registrados en el sistema incluyendo la información que usted necesite consultar.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=705", True),_
						Array("Antigüedad para un empleado",_
							  "Seleccione a un empleado y obtenga su antigüedad junto con sus incidencias.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1702", True),_
						Array("Reporte de antigüedades",_
							  "Obtenga las antigüedades de los empleados, agrupadas por entidad federativa y centro de trabajo.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1703", True),_
						Array("Validación de nómina 1o de Octubre",_
							  "Listado de empleados acreedores del premio por antigüedad y al premio moneda de oro.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1704", True),_
						Array("Conceptos registrados a los empleados",_
							  "Obtenga un archivo zip con el listado de los conceptos registrados para los empleados",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1108", True),_
						Array("Horas extras y primas dominicales registrados a los empleados",_
							  "Obtenga un archivo zip con el listado de horas extras y primas dominicales registrados para los empleados",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1210", True),_
						Array("Reclamos de pago registrados a los empleados",_
							  "Obtenga un archivo zip con el listado de los reclamos de pago registrados para los empleados",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1102", True),_
						Array("Reporte de empleados con derecho al concepto 41",_
							  "Obtenga un archivo zip con el listado de los empleados que cumplen 25 y 30 años en la quincena indicada",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1119", True),_
						Array("",_
							  "",_
							  "", "", False)_
					)
				Case 25 'Prestaciones > Certificaciones y archivo
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Registro de reclamos",_
							  "Registre los reclamos de pago por empleado que hayan sido evaluados y autorizados.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=-58", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RegistroDeReclamosCyA & ",", vbBinaryCompare) > 0),_
						Array("D2. Exceso de incapacidades y licencias médicas",_
							  "Registre los montos por exceso de incapacidades y licencias médicas.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_LICENSES, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_D2_ExcesoDeIncapacidadesYLicenciasMedicas & ",", vbBinaryCompare) > 0),_
						Array("Revisión de nóminas",_
							  "Indique a qué empleados se les realizará una revisión en sus nóminas.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=PayrollRevision", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_RevisionDeNominasCyA & ",", vbBinaryCompare) > 0),_
						Array("72. Otras deducciones",_
							  "Registre el concepto para devoluciones excentas de los empleados que las reportan.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_EXCENT, False),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Actualización de antigüedades",_
							  "Seleccione un empleado y modifique su histórico.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=262", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ActualizacionDeAntiguedades & ",", vbBinaryCompare) > 0),_
						Array("Antigüedad federal",_
							  "Registre la antigüedad federal de los empleados.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=261", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AntiguedadFederal & ",", vbBinaryCompare) > 0),_
						Array("Antigüedad para un empleado",_
							  "Seleccione a un empleado y obtenga su antigüedad junto con sus incidencias.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1204", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AntiguedadParaUnEmpleado & ",", vbBinaryCompare) > 0),_
						Array("Reporte de antigüedades",_
							  "Obtenga las antigüedades de los empleados, agrupadas por entidad federativa y centro de trabajo.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1205", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ReporteDeAntiguedades & ",", vbBinaryCompare) > 0),_
						Array("Validación de nómina 1o de Octubre",_
							  "Listado de empleados acreedores del premio por antigüedad y al premio moneda de oro.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1206", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ValidacionDeNomina1oDeOctubre & ",", vbBinaryCompare) > 0),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Hoja única de servicio",_
							  "Histórico del empleado a lo largo de los años de servicio en el Instituto.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1203", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_HojaUnicaDeServicio & ",", vbBinaryCompare) > 0),_
						Array("(Opción) Hoja única de servicio",_
							  "Histórico del empleado a lo largo de los años de servicio en el Instituto.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=" & EMPLOYEES_SERVICE_SHEET, StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_HojaUnicaDeServicio & ",", vbBinaryCompare) > 0),_
						Array("Entregas de hojas únicas de servicio",_
							  "Control de las entregas a los empleados de sus hojas únicas de servicio.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=267", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_EntregasDeHojasUnicasDeServicio & ",", vbBinaryCompare) > 0),_
						Array("Constancia de servicio activo",_
							  "Indique el número de empleado y obtenga su constancia de servicio activo.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1207", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ConstanciaDeServicioActivo & ",", vbBinaryCompare) > 0),_
						Array("Constancia de descuento",_
							  "Indique el número de empleado y el crédito y obtenga su constancia de descuento.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1208", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_ConstanciaDeDescuento & ",", vbBinaryCompare) > 0)_
					)
				Case 26 'Prestaciones > Antigüedades
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Antigüedad federal",_
							  "Registre la antigüedad federal de los empleados.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=261", True),_
						Array("Antigüedad para un empleado",_
							  "Seleccione a un empleado y obtenga su antigüedad junto con sus incidencias.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1204", True),_
						Array("Reporte de antigüedades",_
							  "Obtenga las antigüedades de los empleados, agrupadas por entidad federativa y centro de trabajo.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1205", True),_
						Array("Validación de nómina 1o de Octubre",_
							  "Listado de empleados acreedores del premio por antigüedad y al premio moneda de oro.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1206", True)_
					)
				Case 27 'Prestaciones > Acreedores de los empleados
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Padrón de acreedores",_
							  "Registre en el padrón de acreedores a los beneficiarios del adeudo.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_CREDITORS, True),_
						Array("Catálogo de tipos de descuentos para acreedores",_
							  "Registre los tipos de descuentos que son asignadas a los acreedores registrados.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=AlimonyTypes&ReasonID=" & CREDITORS_TYPES, True)_
					)
				Case 261 'Prestaciones > Antigüedades > Antigüedad federal
					If Len(oRequest("EmployeeID").Item) = 0 Then bSearchForm = True
					If (lErrorNumber = 0) And bAction Then
						Call DisplayErrorMessage("Confirmación", "La información fue guardada con éxito.<BR /><FORM><INPUT TYPE=""BUTTON"" VALUE=""Registro de información"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=261&EmployeeID=" & oRequest("EmployeeID").Item & "&DoSearch=1'"" /><IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" /><INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=261'"" /></FORM>")
						Response.Write "<BR />"
					End If
					If (lErrorNumber <> 0) And (bAction) Then
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						Response.Write "<BR />"
					End If

					If bSearchForm Then
						lErrorNumber = Display261SearchForm(oRequest, oADODBConnection, sErrorDescription)
					ElseIf Len(oRequest("DoSearch").Item) > 0 Then
						Response.Write "<IFRAME SRC=""SearchRecord.asp?RecordID=" & oRequest("EmployeeID").Item & "&Action=EmployeesInfo"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""500"" HEIGHT=""170""></IFRAME><BR />"
						If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
							aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeeID In (Select EmployeeID From Employees, Jobs Where (Employees.JobID=Jobs.JobID) And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Jobs.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))))"
						End If
						lErrorNumber = DisplayCatalogsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, aCatalogComponent, sErrorDescription)
						If lErrorNumber = L_ERR_NO_RECORDS Then
							Call DisplayErrorMessage("Búsqueda vacía", "No existen registros que cumplan con los criterios de la búsqueda")
							Response.Write "<BR />"
							lErrorNumber = Display261SearchForm(oRequest, oADODBConnection, sErrorDescription)
						End If
					ElseIf bShowForm Then
						Response.Write "<IFRAME SRC=""SearchRecord.asp?RecordID=" & oRequest("EmployeeID").Item & "&Action=EmployeesInfo"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""500"" HEIGHT=""170""></IFRAME><BR />"
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeeID=" & oRequest("EmployeeID").Item & ") And (AntiquityDate=" & oRequest("AntiquityDate").Item & ")"
						lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
						If lErrorNumber = -2 Then
							lErrorNumber = 0
							sErrorDescription = ""
							aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) = -1
							lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "document.CatalogFrm.EmployeeID.value = '" & oRequest("EmployeeID").Item & "';" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine
						End If
					End If
				Case 262 'Prestaciones > Antigüedades > Actualización de antigüedades
					Dim lAntiquityError
					Dim sAntiquityErrorDescription
					lAntiquityError = 0
					Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
					aEmployeeComponent(S_URL_EMPLOYEE) = "SectionID=262"
					If Len(oRequest("Add").Item) > 0 Then
						lErrorNumber = UpdateEmployeeHistoryListRecord(oRequest, oADODBConnection, 1, aEmployeeComponent, sErrorDescription)
						lAntiquityError = lErrorNumber
						sAntiquityErrorDescription = sErrorDescription
					ElseIf Len(oRequest("Modify").Item) > 0 Then
						lErrorNumber = UpdateEmployeeHistoryListRecord(oRequest, oADODBConnection, 2, aEmployeeComponent, sErrorDescription)
						lAntiquityError = lErrorNumber
						sAntiquityErrorDescription = sErrorDescription
					ElseIf Len(oRequest("Remove").Item) > 0 Then
						lErrorNumber = UpdateEmployeeHistoryListRecord(oRequest, oADODBConnection, 0, aEmployeeComponent, sErrorDescription)
						lAntiquityError = lErrorNumber
						sAntiquityErrorDescription = sErrorDescription
					End If
					If (Len(oRequest("EmployeeID").Item) = 0) And (Len(oRequest("AnotherEmployeeID").Item) = 0) Then
						lErrorNumber = Display262SearchForm(oRequest, oADODBConnection, sErrorDescription)
					Else
						If Len(oRequest("AnotherEmployeeID").Item) > 0 Then aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRequest("AnotherEmployeeID").Item)
						lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						If lErrorNumber = 0 Then
							Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
								Response.Write "<TD VALIGN=""TOP""><DIV NAME=""ReportDiv"" ID=""ReportDiv"" STYLE=""height: 450px; width:600px; overflow: auto;"">"
									lErrorNumber = DisplayEmployeeHistoryList(oRequest, oADODBConnection, False, True, aEmployeeComponent, sErrorDescription)
								Response.Write "</DIV></BR>"
								If aEmployeeComponent(N_HISTORY_LIST_RECORTS) <> -1 Then
									Call DisplayInstructionsMessage("Número de registros", "El empleado cuenta con:&nbsp;" & aEmployeeComponent(N_HISTORY_LIST_RECORTS) & " registros en su historial.")
								End If
								Response.Write "</TD>"
								Response.Write "<TD>&nbsp;</TD>"
								Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
								Response.Write "<TD>&nbsp;</TD>"
								Response.Write "<TD WIDTH=""*"" VALIGN=""TOP"">"
									If lErrorNumber = 0 Then
										lErrorNumber = ShowEmployeeHistoryListForm(oRequest, oADODBConnection, GetASPFileName(""), aEmployeeComponent, sErrorDescription)
									End If
									If (Len(oRequest("Add").Item) > 0) Or (Len(oRequest("Modify").Item) > 0) Or (Len(oRequest("Remove").Item) > 0) Then
										If lAntiquityError <> 0 Then
											Call DisplayErrorMessage("Mensaje del sistema", sAntiquityErrorDescription)
										Else
											Call DisplayInstructionsMessage("Mensaje del sistema", "La operación fue realizada exitosamente.")
										End If
										Response.Write "<BR />"
									End If
								Response.Write "</TD>"
							Response.Write "</TR></TABLE>"
						Else
							Response.Write "<BR />"
							Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
								Response.Write "<TR>"
									Response.Write "<TD VALIGN=""TOP"" WIDTH=""10"">&nbsp;</TD>"
											Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><A HREF=""UploadInfo.asp?Action=EmployeeHistoryList&SectionID=262""><IMG SRC=""Images/MnLeftArrows.gif"" WIDTH=""32"" HEIGHT=""32"" ALT=""Incidencias"" BORDER=""0"" /></A><BR /></TD>"
											Response.Write "<TD VALIGN=""TOP"" WIDTH=""290""><FONT FACE=""Arial"" SIZE=""2""><B><A HREF=""Main_ISSSTE.asp?Action=EmployeeHistoryList&SectionID=262"" CLASS=""SpecialLink"">Otro empleado</A></B><BR /></FONT>"
											Response.Write "<DIV CLASS=""MenuOverflow""><FONT FACE=""Arial"" SIZE=""2"">Actualice la antigüedad a un empleado diferente.</FONT></DIV></TD>"
								Response.Write "</TR>"
							Response.Write "</TABLE>"
							Response.Write "<BR /><BR />"
						End If
					End If
				Case 267 'Prestaciones > Antigüedades > Entregas de hojas únicas de servicio
					If (Not bSearchForm) And (Not bShowForm) And (Not bAction) And (Len(oRequest("DoSearch").Item) = 0) Then
						aMenuComponent(A_ELEMENTS_MENU) = Array(_
							Array("Registro de información",_
								  "Alta de entregas de hojas únicas de servicio",_
								  "Images/MnAreas.gif", "Main_ISSSTE.asp?SectionID=267&New=1", True),_
							Array("Búsqueda de información",_
								  "Listado de empleados que han recibido su hoja única de servicio",_
								  "Images/MnJobs.gif", "Main_ISSSTE.asp?SectionID=267&Search=1", True)_
						)
					Else
						If (lErrorNumber = 0) And bAction Then
							Call DisplayErrorMessage("Confirmación", "La operación fue realizada con éxito.<BR /><FORM><INPUT TYPE=""BUTTON"" VALUE=""Registro de información"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=267&New=1'"" /><IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" /><INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=267'"" /></FORM>")
							Response.Write "<BR />"
						ElseIf (lErrorNumber <> 0) And bAction Then
							Call DisplayErrorMessage("Confirmación", "Error al realizar la operación.<BR /><FORM id=form1 name=form1><INPUT TYPE=""BUTTON"" VALUE=""Registro de información"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=267&New=1'"" / id=1 name=1><IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" /><INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=267'"" / id=1 name=1></FORM>")
							Response.Write "<BR />"
						End If
						If lErrorNumber <> 0 Then
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							Response.Write "<BR />"
						End If

						If bSearchForm Then
							lErrorNumber = Display267SearchForm(oRequest, oADODBConnection, sErrorDescription)
						ElseIf Len(oRequest("DoSearch").Item) > 0 Then
							lErrorNumber = DisplayCatalogsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, aCatalogComponent, sErrorDescription)
							If lErrorNumber = L_ERR_NO_RECORDS Then
								Call DisplayErrorMessage("Búsqueda vacía", "No existen registros que cumplan con los criterios de la búsqueda")
								Response.Write "<BR />"
								lErrorNumber = Display267SearchForm(oRequest, oADODBConnection, sErrorDescription)
							End If
						ElseIf bShowForm Then
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine
							lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
						End If
					End If
				Case 28 'Prestaciones > Registro de bolsa de trabajo y escalafón
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Registro de información de la bolsa de trabajo",_
							  "",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=281&New=1", True),_
						Array("Búsqueda de información de la bolsa de trabajo",_
							  "",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=281&Search=1", True),_
						Array("Impresión de información de la bolsa de trabajo",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1354", True),_
						Array("Registro de información de escalafón",_
							  "",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=282&New=1", True),_
						Array("Búsqueda de información de escalafón",_
							  "",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=282&Search=1", True),_
						Array("Impresión de información de escalafón",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1356", True)_
					)
				Case 281 'Prestaciones > Registro de bolsa de trabajo y escalafón > Registro de información de la bolsa de trabajo
					If (Not bSearchForm) And (Not bShowForm) And (Not bAction) And (Len(oRequest("DoSearch").Item) = 0) Then
						aMenuComponent(A_ELEMENTS_MENU) = Array(_
							Array("Registro de información",_
								  "Alta de la información",_
								  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=282&New=1", True),_
							Array("Búsqueda de información",_
								  "Listado de la información registrada",_
								  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=282&Search=1", True)_
						)
					Else
						If (lErrorNumber = 0) And bAction Then
							Call DisplayErrorMessage("Confirmación", "La información fue guardada con éxito.<BR /><FORM><INPUT TYPE=""BUTTON"" VALUE=""Registro de información"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=282&New=1'"" /><IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" /><INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=28'"" /></FORM>")
							Response.Write "<BR />"
						End If
						If lErrorNumber <> 0 Then
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							Response.Write "<BR />"
						End If

						If bSearchForm Then
							lErrorNumber = Display281SearchForm(oRequest, oADODBConnection, sErrorDescription)
						ElseIf Len(oRequest("DoSearch").Item) > 0 Then
							lErrorNumber = DisplayCatalogsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, (CInt(Request.Cookies("SIAP_SectionID")) = 2), aCatalogComponent, sErrorDescription)
							If lErrorNumber = L_ERR_NO_RECORDS Then
								Call DisplayErrorMessage("Búsqueda vacía", "No existen registros que cumplan con los criterios de la búsqueda")
								Response.Write "<BR />"
								lErrorNumber = Display281SearchForm(oRequest, oADODBConnection, sErrorDescription)
							End If
						ElseIf bShowForm Then
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "function ShowKardexFields(sKardex5TypeID) {" & vbNewLine
									Response.Write "oForm = document.CatalogFrm;" & vbNewLine

									Response.Write "if (oForm) {" & vbNewLine
										Response.Write "HideDisplay(document.all['CatalogFrm_StartDateDiv']);" & vbNewLine
										Response.Write "HideDisplay(document.all['CatalogFrm_SchoolarshipIDDiv']);" & vbNewLine
										Response.Write "HideDisplay(document.all['CatalogFrm_RelationshipDiv']);" & vbNewLine
										Response.Write "HideDisplay(document.all['CatalogFrm_ServiceYearsDiv']);" & vbNewLine
										Response.Write "HideDisplay(document.all['CatalogFrm_KardexYearsDiv']);" & vbNewLine
										Response.Write "switch (sKardex5TypeID) {" & vbNewLine
											Response.Write "case '0': // Tipo cronológico" & vbNewLine
												Response.Write "ShowDisplay(document.all['CatalogFrm_StartDateDiv']);" & vbNewLine
												Response.Write "break;" & vbNewLine
											Response.Write "case '1': // Tipo puntuación" & vbNewLine
												Response.Write "ShowDisplay(document.all['CatalogFrm_SchoolarshipIDDiv']);" & vbNewLine
												Response.Write "ShowDisplay(document.all['CatalogFrm_RelationshipDiv']);" & vbNewLine
												Response.Write "ShowDisplay(document.all['CatalogFrm_ServiceYearsDiv']);" & vbNewLine
												Response.Write "ShowDisplay(document.all['CatalogFrm_KardexYearsDiv']);" & vbNewLine
												Response.Write "break;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "} // End of ShowKardexFields" & vbNewLine

								Response.Write "function CheckEmployeeForm() {" & vbNewLine
									Response.Write "oForm = document.EmployeeFrm;" & vbNewLine
									Response.Write "if (oForm) {" & vbNewLine
'										Response.Write "if (oForm.EmployeeID.value == '') {" & vbNewLine
'											Response.Write "alert('Favor de validar el número de empleado.');" & vbNewLine
'											Response.Write "return false;" & vbNewLine
'										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "} // End of CheckEmployeeForm" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine

							lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)

							Response.Write "<FORM NAME=""EmployeeFrm"" ID=""EmployeeFrm"" onSubmit=""return false"">"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" VALUE="""" />"
							Response.Write "</FORM>"
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "ShowKardexFields(document.CatalogFrm.Kardex5TypeID.value);" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine
						End If
					End If
				Case 282 'Prestaciones > Registro de bolsa de trabajo y escalafón > Registro de información de escalafón
					If (Not bSearchForm) And (Not bShowForm) And (Not bAction) And (Len(oRequest("DoSearch").Item) = 0) Then
						aMenuComponent(A_ELEMENTS_MENU) = Array(_
							Array("Registro de información",_
								  "Alta de la información",_
								  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=281&New=1", True),_
							Array("Búsqueda de información",_
								  "Listado de la información registrada",_
								  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=281&Search=1", True)_
						)
					Else
						If (lErrorNumber = 0) And bAction Then
							Call DisplayErrorMessage("Confirmación", "La información fue guardada con éxito.<BR /><FORM><INPUT TYPE=""BUTTON"" VALUE=""Registro de información"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=281&New=1'"" /><IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" /><INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=28'"" /></FORM>")
							Response.Write "<BR />"
						End If
						If lErrorNumber <> 0 Then
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							Response.Write "<BR />"
						End If

						If bSearchForm Then
							lErrorNumber = Display282SearchForm(oRequest, oADODBConnection, sErrorDescription)
						ElseIf Len(oRequest("DoSearch").Item) > 0 Then
							lErrorNumber = DisplayCatalogsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, (CInt(Request.Cookies("SIAP_SectionID")) = 2), aCatalogComponent, sErrorDescription)
							If lErrorNumber = L_ERR_NO_RECORDS Then
								Call DisplayErrorMessage("Búsqueda vacía", "No existen registros que cumplan con los criterios de la búsqueda")
								Response.Write "<BR />"
								lErrorNumber = Display282SearchForm(oRequest, oADODBConnection, sErrorDescription)
							End If
						ElseIf bShowForm Then
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "function ShowKardexFields(sKardexChangeTypeID) {" & vbNewLine
									Response.Write "oForm = document.CatalogFrm;" & vbNewLine

									Response.Write "if (oForm) {" & vbNewLine
										Response.Write "HideDisplay(document.all['CatalogFrm_PositionIDDiv']);" & vbNewLine
										Response.Write "HideDisplay(document.all['CatalogFrm_ServiceIDDiv']);" & vbNewLine
										Response.Write "HideDisplay(document.all['CatalogFrm_BranchIDDiv']);" & vbNewLine
										Response.Write "switch (sKardexChangeTypeID) {" & vbNewLine
											Response.Write "case '0': // Cambio de puesto" & vbNewLine
												Response.Write "ShowDisplay(document.all['CatalogFrm_PositionIDDiv']);" & vbNewLine
												Response.Write "ShowDisplay(document.all['CatalogFrm_ServiceIDDiv']);" & vbNewLine
												Response.Write "break;" & vbNewLine
											Response.Write "case '1': // Cambio de rama" & vbNewLine
												Response.Write "ShowDisplay(document.all['CatalogFrm_PositionIDDiv']);" & vbNewLine
												Response.Write "ShowDisplay(document.all['CatalogFrm_ServiceIDDiv']);" & vbNewLine
												Response.Write "ShowDisplay(document.all['CatalogFrm_BranchIDDiv']);" & vbNewLine
												Response.Write "break;" & vbNewLine
											Response.Write "case '2': // Cambio de residencia estado de procedencia a otro estado" & vbNewLine
												Response.Write "break;" & vbNewLine
											Response.Write "case '3': // Cambio de residencia de otro estado al estado de procedencia" & vbNewLine
												Response.Write "break;" & vbNewLine
											Response.Write "case '4': // Cambio de turno" & vbNewLine
												Response.Write "break;" & vbNewLine
											Response.Write "case '5': // Cambio de adscripción" & vbNewLine
												Response.Write "break;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "} // End of ShowKardexFields" & vbNewLine

								Response.Write "function CheckEmployeeForm() {" & vbNewLine
									Response.Write "oForm = document.EmployeeFrm;" & vbNewLine
									Response.Write "if (oForm) {" & vbNewLine
										Response.Write "if (oForm.EmployeeID.value == '') {" & vbNewLine
											Response.Write "alert('Favor de validar el número de empleado.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine

										Response.Write "if (oForm.JobID.value == '') {" & vbNewLine
											Response.Write "alert('Favor de validar el número de plaza.');" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine

									Response.Write "if (document.CatalogFrm.PositionID.value == '-1')" & vbNewLine
										Response.Write "document.CatalogFrm.PositionID.value = oForm.PositionID.value;" & vbNewLine

'									Response.Write "if (parseInt(document.CatalogFrm.StartDateYear.value + document.CatalogFrm.StartDateMonth.value + document.CatalogFrm.StartDateDay.value) < " & Left(GetSerialNumberForDate(""), Len("00000000")) & ") {" & vbNewLine
'										Response.Write "alert('La fecha de registro no puede ser anterior al día de hoy.');" & vbNewLine
'										Response.Write "document.CatalogFrm.StartDateDay.focus();" & vbNewLine
'										Response.Write "return false;" & vbNewLine
'									Response.Write "}" & vbNewLine
									Response.Write "return true;" & vbNewLine
								Response.Write "} // End of CheckEmployeeForm" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine
							Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
									lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
									Response.Write "<FORM NAME=""EmployeeFrm"" ID=""EmployeeFrm"" onSubmit=""return false"">"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" VALUE="""" />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobID"" ID=""JobIDTxt"" VALUE="""" />"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PositionID"" ID=""PositionIDTxt"" VALUE="""" />"
									Response.Write "</FORM>"
								Response.Write "</TD>"
								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
									Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""240""></IFRAME>"
									Response.Write "<BR />"
									Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchJobNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""240""></IFRAME>"
								Response.Write "</TD>"
							Response.Write "</TR></TABLE>"
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "ShowKardexFields(document.CatalogFrm.KardexChangeTypeID.value);" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine
						End If
					End If
				Case 29 'Matriz de riesgos profesionales
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Administración de Matriz de Riesgos",_
							  "Administración de la matriz de puestos sujetos al pago de Riesgos Profesionales.",_
							  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=ProfessionalRiskMatrix", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministracionDeMatrizDeRiesgos & ",", vbBinaryCompare) > 0),_
						Array("Carga de matriz de riesgos profesionales",_
							  "Utilice un archivo para registrar la Matriz de Riesgos Pprofesionales vigente",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=ProfessionalRisk", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_CargaDeMatrizDeRiesgosProfesionales & ",", vbBinaryCompare) > 0)_
					)
				Case 291
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Acumulados anuales",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1151", True),_
						Array("Registro de empleados que no desean el ajuste anual del impuesto sobre la renta",_
							  "",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=151", False),_
						Array("Ajuste anual del impuesto sobre la renta",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1153", False),_
						Array("Aplicación del ajuste anual del impuesto sobre la renta",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1155", False),_
						Array("Aplicación por empleado del ajuste anual del impuesto sobre la renta",_
							  "",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=156", False),_
						Array("Recálculo anual de impuestos",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1154", False),_
						Array("Declaración informativa múltiple (DIM)",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1157", False),_
						Array("Constancia de percepciones y deducciones anuales",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1152", False)_
					)
				Case 3 'Desarrollo humano
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Administración de plazas",_
							  "Busque las plazas que desea administrar.",_
							  "Images/MnJobs.gif", "Jobs.asp", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_AdministracionDePlazas & ",", vbBinaryCompare) > 0),_
						Array("Estructuras ocupacionales",_
							  "Administración de plantillas de personal, puestos, plazas, centros de trabajo y tabuladores.",_
							  "Images/MnSection27.gif", "Main_ISSSTE.asp?SectionID=31", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_EstructurasOcupacionales & ",", vbBinaryCompare) > 0),_
						Array("Reportes",_
							  "Reportes del área de desarrollo humano",_
							  "Images/MnReports.gif", "Main_ISSSTE.asp?SectionID=34", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_Reportes & ",", vbBinaryCompare) > 0),_
						Array("Selección de personal",_
							  "Seleccione a las personas que ocuparán las plazas vacantes de la bolsa de trabajo del Instituto.",_
							  "Images/MnSection34.gif", "Main_ISSSTE.asp?SectionID=35", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_SeleccionDePersonal & ",", vbBinaryCompare) > 0),_
						Array("Desarrollo humano",_
							  "Administre los cursos que tomarán los empledos del Instituto y obtenga reportes.",_
							  "Images/MnSection35.gif", "Main_ISSSTE.asp?SectionID=36", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_DesarrolloHumano & ",", vbBinaryCompare) > 0),_
						Array("Planeación de recursos humanos",_
							  "Administre los procedimientos, programas institucionales y especiales, metas institucionales, registros únicos de servidores públicos y plantilla de personal vacante.",_
							  "Images/MnSection36.gif", "Main_ISSSTE.asp?SectionID=37", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_PlaneacionDeRecursosHumanos & ",", vbBinaryCompare) > 0),_
						Array("Búsqueda de centros de trabajo y centros de pago",_
							  "A través de la clave, obtenga un listado de los centros de trabajo y de los centros de pago.",_
							  "Images/MnSearch.gif", "Main_ISSSTE.asp?SectionID=38", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_BusquedaDeCentrosDeTrabajoYCentrosDePago & ",", vbBinaryCompare) > 0),_
						Array("Ventanilla única",_
							  "Control de los trámite que ingresan a la Subdirección de Personal.",_
							  "Images/MnSection61.gif", "Main_ISSSTE.asp?SectionID=61", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_VentanillaUnica & ",", vbBinaryCompare) > 0)_
					)
				Case 31 'Desarrollo humano > Estructuras ocupacionales
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Catálogos",_
							  "Altas, bajas y cambios de registros concernientes a los registros del sistema.",_
							  "Images/MnHumanResources.gif", "Catalogs.asp", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_Catalogos & ",", vbBinaryCompare) > 0),_
						Array("Consulta de tabuladores",_
							  "Consulte la información por cada tipo de tabulador.",_
							  "Images/MnSearch.gif", "Main_ISSSTE.asp?SectionID=32", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_ConsultaDeTabuladores & ",", vbBinaryCompare) > 0),_
						Array("Registro de tabuladores",_
							  "Cargue el archivo que contiene la información por cada tipo de tabulador.",_
							  "Images/MnSearch.gif", "Main_ISSSTE.asp?SectionID=33", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_RegistroDeTabuladores & ",", vbBinaryCompare) > 0),_
						Array("Carga UNIMED",_
							  "Cargue el archivo que contiene la información UNIMED.",_
							  "Images/MnSection31.gif", "UploadInfo.asp?Action=MedicalAreas", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_CargaUNIMED & ",", vbBinaryCompare) > 0),_
						Array("Tabuladores",_
							  "Administrar los valores de los conceptos de pago y por tipos de tabulador.",_
							  "Images/MnBudget.gif", "Payroll.asp?Action=EmployeeTypes", False)_
					)
				Case 32, 33 'Desarrollo humano > Estructuras ocupacionales > Tabuladores
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Alta responsabilidad",_
							  "Cargue el archivo que contiene el tabulador de alta responsabilidad.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=ConceptsValues&EmployeeTypeID=3", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorIncurrirEnFaltaAdministrativa & ",", vbBinaryCompare) > 0),_
						Array("Becario",_
							  "Cargue el archivo que contiene el tabulador de becarios",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=ConceptsValues&EmployeeTypeID=6", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorIncurrirEnFaltaAdministrativa & ",", vbBinaryCompare) > 0),_
						Array("Enlace",_
							  "Cargue el archivo que contiene el tabulador de enlace.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=ConceptsValues&EmployeeTypeID=4", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorIncurrirEnFaltaAdministrativa & ",", vbBinaryCompare) > 0),_
						Array("Funcionario",_
							  "Cargue el archivo que contiene el tabulador de funcionarios.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=ConceptsValues&EmployeeTypeID=1", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorIncurrirEnFaltaAdministrativa & ",", vbBinaryCompare) > 0),_
						Array("Médica, paramédica y grupos afines",_
							  "Cargue el archivo que contiene el tabulador Médica, paramédica y grupos afines.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=ConceptsValues&EmployeeTypeID=0", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorIncurrirEnFaltaAdministrativa & ",", vbBinaryCompare) > 0),_
						Array("Operativo",_
							  "Cargue el archivo que contiene el tabulador de operativos.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=ConceptsValues&EmployeeTypeID=2", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorIncurrirEnFaltaAdministrativa & ",", vbBinaryCompare) > 0),_
						Array("Residente",_
							  "Cargue el archivo que contiene el tabulador de residentes.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=ConceptsValues&EmployeeTypeID=5", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_01_BajaPorIncurrirEnFaltaAdministrativa & ",", vbBinaryCompare) > 0)_
					)
				Case 34 'Desarrollo humano > Reportes
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Reportes guardados",_
							  "Reportes generados por otros usuarios a partir de plantillas y filtros.",_
							  "Images/MnLeftArrows.gif", "SavedReport.asp?ReportType=1", True),_
						Array("Catálogo de puestos y tabuladores de puestos",_
							  "Obtenga los tabuladores de médica, paramédica y grupos afines, funcionarios, operativos, alta responsabilidad, enlace y residentes",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1335", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_Reportes & ",", vbBinaryCompare) > 0),_
						Array("Centros de pago",_
							  "Obtenga el listado de los centros de pago.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1337", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_Reportes & ",", vbBinaryCompare) > 0),_
						Array("Centros de trabajo",_
							  "Obtenga el listado de los centros de trabajo.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1336", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_Reportes & ",", vbBinaryCompare) > 0),_
						Array("Generación del archivo para registro de servidores públicos",_
							  "Obtenga el listado de registros únicos de empleados por empleado, entidad federativa y nómina.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1371", False And StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_Reportes & ",", vbBinaryCompare) > 0),_
						Array("Generación del reporte de unidades medicas UNIMED",_
							  "Obtenga los listados de los reportes UNIMED por ubicación, tipo de reporte y quincena.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1334", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_CargaUNIMED & ",", vbBinaryCompare) > 0),_
						Array("Plantilla de personal",_
							  "Generación de plantilla de personal por tipo de centro de trabajo",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1311", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_Reportes & ",", vbBinaryCompare) > 0),_
						Array("Conteo de empleados",_
							  "Obtenga un conteo de los empleados que se encuentran registrados en el sistema, agrupando los resultados por diferentes conceptos.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=702", True),_
						Array("Información de los empleados",_
							  "Obtenga un listado de los empleados registrados en el sistema incluyendo la información que usted necesite consultar.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=705", True),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("RUSP. Información básica",_
							  "",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1372", True),_
						Array("RUSP. Bajas",_
							  "",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1373", True),_
						Array("RUSP. Datos personales",_
							  "",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1374", True)_
					)
				Case 35 'Desarrollo humano > Selección de personal
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Registro de información",_
							  "",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=351", True),_
						Array("Validación del proceso de selección de personal",_
							  "",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=352&Search=1", True),_
						Array("Búsqueda de información del proceso de selección",_
							  "",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=353&Search=1", True),_
						Array("Búsqueda de información de la bolsa de trabajo",_
							  "",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=281&Search=1", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_BolsaDeTrabajo & ",", vbBinaryCompare) > 0),_
						Array("Búsqueda de información de escalafón",_
							  "",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=282&Search=1", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_Escalafon & ",", vbBinaryCompare) > 0),_
						Array("",_
							  "",_
							  "", "", False)_
					)
				Case EMPLOYEES_SERVICE_SHEET
					If iStep <= 1 Then
						Dim sAltDescription
						Dim sDescription
						Select Case sAction
							Case "EmployeesMovements"
								Select Case lReasonID
									Case EMPLOYEES_BANK_ACCOUNTS
										sAltDescription = "Cuenta bancaria"
										sDescription = "Registre una cuenta bancaria a un empleado diferente."
									Case EMPLOYEES_ADD_BENEFICIARIES
										sAltDescription = "Beneficiario de pensión alimenticia"
										sDescription = "Registre un(a) beneficiario(a) de pensión alimenticia a un empleado diferente."
									Case EMPLOYEES_CREDITORS
										sAltDescription = "Acreedores"
										sDescription = "Registre un(a) acreedor(a) de adeudos a un empleado diferente."
									Case -96,-75,-64,1,2,3,4,5,6,8,10,12,13,14,17,18,21,26,28,29,30,31,32,33,34,37,38,39,40,41,43,44,45,46,47,48,50,51,53,57,58,62,63,66,68
										sAltDescription = "Movimientos de personal"
										sDescription = "Registre el movimiento a un empleado diferente."
									Case Else
										sAltDescription = "Prestación"
										sDescription = "Registre la prestación a un empleado diferente."
								End Select
							Case "Absences"
								sAltDescription = "Incidencias"
								sDescription = "Registre incidencias a un empleado diferente."
							Case Else
								sAltDescription = "Hoja única de servicios"
								sDescription = "Registre la solicitud de una Hoja única de servicios a un empleado diferente."
						End Select
						If Len(oRequest("EmployeeNumber").Item) > 0 Then
							aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = CLng(aEmployeeComponent(S_NUMBER_EMPLOYEE))
							lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							If lErrorNumber = 0 Then
								'If VerifyRequerimentsForEmployeesConcepts(oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription) Then
								If True Then
									Call DisplayUploadForm("ServiceSheet", -1, EMPLOYEES_SERVICE_SHEET)
								Else
									lErrorNumber = -1
									Call DisplayAnotherEmployeeForm(oRequest, oADODBConnection, "Main_ISSSTE.asp", "ServiceSheet", 10, EMPLOYEES_SERVICE_SHEET, sAltDescription, sDescription, sErrorDescription)
								End If
							Else
								lErrorNumber = -1
								Call DisplayAnotherEmployeeForm(oRequest, oADODBConnection, "Main_ISSSTE.asp", "ServiceSheet", 10, EMPLOYEES_SERVICE_SHEET, sAltDescription, sDescription, sErrorDescription)
							End If
						Else
							Call DisplayUploadForm("ServiceSheet", -1, EMPLOYEES_SERVICE_SHEET)
						End If
						If Len(oRequest("Success").Item) > 0 Then
							If CInt(oRequest("Success").Item) = 1 Then
								Select Case sAction
									Case "Absences"
										'Call DisplayErrorMessage("Confirmación", "La operación con la incidencia " & sAbsenceShortName & " fué ejecutada exitosamente.")
									Case Else
										Call DisplayErrorMessage("Confirmación", "La operación fué ejecutada exitosamente." & CStr(oRequest("ErrorDescription").Item))
								End Select
							Else
								Select Case sAction
									Case "Absences"				
										'Call DisplayErrorMessage("Error al realizar la operación con la incidencia " & sAbsenceShortName, CStr(oRequest("ErrorDescription").Item))
									Case Else
										Call DisplayErrorMessage("Error al realizar la operación", CStr(oRequest("ErrorDescription").Item))
								End Select
							End If
						End If
					End If
				Case 351 'Desarrollo humano > Selección de personal > Registro de información
					If (Not bSearchForm) And (Not bShowForm) And (Not bAction) And (Len(oRequest("DoSearch").Item) = 0) Then
						aMenuComponent(A_ELEMENTS_MENU) = Array(_
							Array("Registro de información",_
								  "Alta de la información",_
								  "Images/MnAreas.gif", "Main_ISSSTE.asp?SectionID=351&New=1", True),_
							Array("Búsqueda de información",_
								  "Listado de la información registrada",_
								  "Images/MnJobs.gif", "Main_ISSSTE.asp?SectionID=351&Search=1", True)_
						)
					Else
						If (lErrorNumber = 0) And bAction Then
							Call DisplayErrorMessage("Confirmación", "La información fue guardada con éxito.<BR /><FORM><INPUT TYPE=""BUTTON"" VALUE=""Registro de información"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=351&New=1'"" /><IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" /><INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=351'"" /></FORM>")
							Response.Write "<BR />"
						End If
						If lErrorNumber <> 0 Then
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							Response.Write "<BR />"
						End If

						If bSearchForm Then
							lErrorNumber = Display351SearchForm(oRequest, oADODBConnection, sErrorDescription)
						ElseIf Len(oRequest("DoSearch").Item) > 0 Then
							lErrorNumber = DisplayCatalogsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, aCatalogComponent, sErrorDescription)
							If lErrorNumber = L_ERR_NO_RECORDS Then
								Call DisplayErrorMessage("Búsqueda vacía", "No existen registros que cumplan con los criterios de la búsqueda")
								Response.Write "<BR />"
								lErrorNumber = Display351SearchForm(oRequest, oADODBConnection, sErrorDescription)
							End If
						ElseIf bShowForm Then
							Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchPositionsIFrame"" FRAMEBORDER=""0"" WIDTH=""0"" HEIGHT=""0""></IFRAME>"
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "var bDocReady = false;" & vbNewLine
								Response.Write "var KardexOrigins = new Array("
									Call GenerateJavaScriptArrayFromQuery(oADODBConnection, "KardexOrigins", "KardexOriginID", "KardexOriginName, KardexTypeID", "(Active=1)", "KardexOriginID", sErrorDescription)
								Response.Write "['-2', '', '']);" & vbNewLine

								Response.Write "var KardexRequirements = new Array("
									Call GenerateJavaScriptArrayFromQuery(oADODBConnection, "KardexRequirements", "KardexRequirementID", "KardexRequirementName, RequirementsTypeID, IsOptional", "(Active=1)", "KardexRequirementID", sErrorDescription)
								Response.Write "['-2', '', '', '']);" & vbNewLine

								Response.Write "var RequirementsTypes = new Array("
									Call GenerateJavaScriptArrayFromQuery(oADODBConnection, "RequirementsTypes", "RequirementsTypeID", "RequirementsTypeName, KardexTypeID", "(Active=1)", "RequirementsTypeID", sErrorDescription)
								Response.Write "['-2', '', '']);" & vbNewLine

								Response.Write "function CheckKardexRequirements(oForm) {" & vbNewLine
									If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)) = -1 Then
										If Month(Date()) = 1 Then
											Response.Write "if (parseInt(oForm.StartDateYear.value + oForm.StartDateMonth.value + oForm.StartDateDay.value) < " & (Year(Date()) - 1) & "1201) {" & vbNewLine
										Else
											Response.Write "if (parseInt(oForm.StartDateYear.value + oForm.StartDateMonth.value + oForm.StartDateDay.value) < " & Year(Date()) & Right(("00" & (Month(Date()) - 1)), Len("00")) & "01) {" & vbNewLine
										End If
											Response.Write "alert('No se pueden registrar trámites anteriores al mes pasado.');" & vbNewLine
											Response.Write "oForm.StartDateDay.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (parseInt(oForm.StartDateYear.value + oForm.StartDateMonth.value + oForm.StartDateDay.value) > " & Left(GetSerialNumberForDate(""), Len("00000000")) & ") {" & vbNewLine
											Response.Write "alert('La fecha de registro no puede ser posterior al día de hoy.');" & vbNewLine
											Response.Write "oForm.StartDateDay.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									End If

									Response.Write "if (parseInt(oForm.DocumentsDateYear.value + oForm.DocumentsDateMonth.value + oForm.DocumentsDateDay.value) > 0) {" & vbNewLine
										Response.Write "if (parseInt(oForm.StartDateYear.value + oForm.StartDateMonth.value + oForm.StartDateDay.value) > parseInt(oForm.DocumentsDateYear.value + oForm.DocumentsDateMonth.value + oForm.DocumentsDateDay.value)) {" & vbNewLine
											Response.Write "alert('La fecha de recepción de documentos no puede ser anterior a la fecha de inicio del trámite.');" & vbNewLine
											Response.Write "oForm.DocumentsDateDay.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine

									Response.Write "if (oForm.KnowledgeStatusID.value != '') {" & vbNewLine
										Response.Write "if ((parseInt(oForm.KnowledgeDateYear.value) == 0) || (parseInt('1' + oForm.KnowledgeDateMonth.value) == 100) || (parseInt('1' + oForm.KnowledgeDateDay.value) == 100)) {" & vbNewLine
											Response.Write "alert('Favor de introducir la fecha de evaluación de conocimientos.');" & vbNewLine
											Response.Write "oForm.KnowledgeDateDay.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine

									Response.Write "if (parseInt(oForm.KnowledgeDateYear.value + oForm.KnowledgeDateMonth.value + oForm.KnowledgeDateDay.value) > 0) {" & vbNewLine
										Response.Write "if (parseInt(oForm.DocumentsDateYear.value + oForm.DocumentsDateMonth.value + oForm.DocumentsDateDay.value) > parseInt(oForm.KnowledgeDateYear.value + oForm.KnowledgeDateMonth.value + oForm.KnowledgeDateDay.value)) {" & vbNewLine
											Response.Write "alert('La fecha de evaluación de conocimientos no puede ser anterior a la fecha de recepción de documentos.');" & vbNewLine
											Response.Write "oForm.KnowledgeDateDay.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine

									Response.Write "if (oForm.PsychologicStatusID.value != '') {" & vbNewLine
										Response.Write "if ((parseInt(oForm.PsychologicDateYear.value) == 0) || (parseInt('1' + oForm.PsychologicDateMonth.value) == 100) || (parseInt('1' + oForm.PsychologicDateDay.value) == 100)) {" & vbNewLine
											Response.Write "alert('Favor de introducir la fecha de la evaluación psicológica.');" & vbNewLine
											Response.Write "oForm.PsychologicDateDay.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine

									Response.Write "if ((parseInt(oForm.PsychologicDateYear.value) > 0) || (parseInt('1' + oForm.PsychologicDateMonth.value) > 100) || (parseInt('1' + oForm.PsychologicDateDay.value) > 100)) {" & vbNewLine
										Response.Write "if (parseInt(oForm.KnowledgeDateYear.value + oForm.KnowledgeDateMonth.value + oForm.KnowledgeDateDay.value) > parseInt(oForm.PsychologicDateYear.value + oForm.PsychologicDateMonth.value + oForm.PsychologicDateDay.value)) {" & vbNewLine
											Response.Write "alert('La fecha de la evaluación psicológica no puede ser anterior a la fecha de evaluación de conocimientos.');" & vbNewLine
											Response.Write "oForm.PsychologicDateDay.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine

									Response.Write "switch (oForm.KardexTypeID.value) {" & vbNewLine
										Response.Write "case '1':" & vbNewLine 'Nuevo ingreso. Base
											Response.Write "if (oForm.KnowledgeStatusID.value != '0') {" & vbNewLine
												Response.Write "if ((parseInt(oForm.PsychologicDateYear.value) == 0) && (parseInt(oForm.PsychologicDateMonth.value) == 0) && (parseInt(oForm.PsychologicDateDay.value) == 0)) {" & vbNewLine
													Response.Write "if (parseInt(oForm.Registration1DateYear.value + oForm.Registration1DateMonth.value + oForm.Registration1DateDay.value) > 0) {" & vbNewLine
														Response.Write "alert('Favor de indicar la fecha de la evaluación psicológica.');" & vbNewLine
														Response.Write "oForm.PsychologicDateDay.focus();" & vbNewLine
														Response.Write "return false;" & vbNewLine
													Response.Write "}" & vbNewLine
												Response.Write "} else {" & vbNewLine
													Response.Write "if ((oForm.PsychologicStatusID.value == '2') && (parseInt(oForm.Registration1DateYear.value + oForm.Registration1DateMonth.value + oForm.Registration1DateDay.value) > 0)) {" & vbNewLine
														Response.Write "alert('La fecha de registro en bolsa de trabajo no puede ser registrada mientras la evaluación psicológica esté programada.');" & vbNewLine
														Response.Write "oForm.Registration1DateDay.focus();" & vbNewLine
														Response.Write "return false;" & vbNewLine
													Response.Write "}" & vbNewLine

													Response.Write "if (parseInt(oForm.Registration1DateYear.value + oForm.Registration1DateMonth.value + oForm.Registration1DateDay.value) > 0) {" & vbNewLine
														Response.Write "if (parseInt(oForm.PsychologicDateYear.value + oForm.PsychologicDateMonth.value + oForm.PsychologicDateDay.value) > parseInt(oForm.Registration1DateYear.value + oForm.Registration1DateMonth.value + oForm.Registration1DateDay.value)) {" & vbNewLine
															Response.Write "alert('La fecha de registro en bolsa de trabajo no puede ser anterior a la fecha de la evaluación psicológica.');" & vbNewLine
															Response.Write "oForm.Registration1DateDay.focus();" & vbNewLine
															Response.Write "return false;" & vbNewLine
														Response.Write "}" & vbNewLine
													Response.Write "}" & vbNewLine
												Response.Write "}" & vbNewLine
											Response.Write "} else {" & vbNewLine
												Response.Write "if (parseInt(oForm.Registration1DateYear.value + oForm.Registration1DateMonth.value + oForm.Registration1DateDay.value) > 0) {" & vbNewLine
													Response.Write "if (parseInt(oForm.KnowledgeDateYear.value + oForm.KnowledgeDateMonth.value + oForm.KnowledgeDateDay.value) > parseInt(oForm.Registration1DateYear.value + oForm.Registration1DateMonth.value + oForm.Registration1DateDay.value)) {" & vbNewLine
														Response.Write "alert('La fecha de registro en bolsa de trabajo no puede ser anterior a la fecha de evaluación de conocimientos.');" & vbNewLine
														Response.Write "oForm.Registration1DateDay.focus();" & vbNewLine
														Response.Write "return false;" & vbNewLine
													Response.Write "}" & vbNewLine
												Response.Write "}" & vbNewLine
											Response.Write "}" & vbNewLine
											Response.Write "break;" & vbNewLine
										Response.Write "case '2':" & vbNewLine 'Nuevo ingreso. Confianza
											Response.Write "if (oForm.KnowledgeStatusID.value != '0') {" & vbNewLine
												Response.Write "if ((parseInt(oForm.PsychologicDateYear.value) == 0) || (parseInt('1' + oForm.PsychologicDateMonth.value) == 100) || (parseInt('1' + oForm.PsychologicDateDay.value) == 100)) {" & vbNewLine
													Response.Write "if (parseInt(oForm.Registration3DateYear.value + oForm.Registration3DateMonth.value + oForm.Registration3DateDay.value) > 0) {" & vbNewLine
														Response.Write "alert('Favor de indicar la fecha de la evaluación psicológica.');" & vbNewLine
														Response.Write "oForm.PsychologicDateDay.focus();" & vbNewLine
														Response.Write "return false;" & vbNewLine
													Response.Write "}" & vbNewLine
												Response.Write "} else {" & vbNewLine
													Response.Write "if ((oForm.PsychologicStatusID.value == '2') && (parseInt(oForm.Registration3DateYear.value + oForm.Registration3DateMonth.value + oForm.Registration3DateDay.value) > 0)) {" & vbNewLine
														Response.Write "alert('La fecha de envío al área de recursos humanos no puede ser registrada mientras la evaluación psicológica esté programada.');" & vbNewLine
														Response.Write "oForm.Registration3DateDay.focus();" & vbNewLine
														Response.Write "return false;" & vbNewLine
													Response.Write "}" & vbNewLine

													Response.Write "if (parseInt(oForm.Registration3DateYear.value + oForm.Registration3DateMonth.value + oForm.Registration3DateDay.value) > 0) {" & vbNewLine
														Response.Write "if (parseInt(oForm.PsychologicDateYear.value + oForm.PsychologicDateMonth.value + oForm.PsychologicDateDay.value) > parseInt(oForm.Registration3DateYear.value + oForm.Registration3DateMonth.value + oForm.Registration3DateDay.value)) {" & vbNewLine
															Response.Write "alert('La fecha de envío al área de recursos humanos no puede ser anterior a la fecha de la evaluación psicológica.');" & vbNewLine
															Response.Write "oForm.Registration3DateDay.focus();" & vbNewLine
															Response.Write "return false;" & vbNewLine
														Response.Write "}" & vbNewLine
													Response.Write "}" & vbNewLine
												Response.Write "}" & vbNewLine
											Response.Write "} else {" & vbNewLine
												Response.Write "if (parseInt(oForm.Registration3DateYear.value + oForm.Registration3DateMonth.value + oForm.Registration3DateDay.value) > 0) {" & vbNewLine
													Response.Write "if (parseInt(oForm.KnowledgeDateYear.value + oForm.KnowledgeDateMonth.value + oForm.KnowledgeDateDay.value) > parseInt(oForm.Registration3DateYear.value + oForm.Registration3DateMonth.value + oForm.Registration3DateDay.value)) {" & vbNewLine
														Response.Write "alert('La fecha de envío al área de recursos humanos no puede ser anterior a la fecha de evaluación de conocimientos.');" & vbNewLine
														Response.Write "oForm.Registration3DateDay.focus();" & vbNewLine
														Response.Write "return false;" & vbNewLine
													Response.Write "}" & vbNewLine
												Response.Write "}" & vbNewLine
											Response.Write "}" & vbNewLine
											Response.Write "break;" & vbNewLine
										Response.Write "case '3':" & vbNewLine 'Trabajador. Base
											Response.Write "if (parseInt(oForm.Registration2DateYear.value + oForm.Registration2DateMonth.value + oForm.Registration2DateDay.value) > 0) {" & vbNewLine
												Response.Write "if (parseInt(oForm.KnowledgeDateYear.value + oForm.KnowledgeDateMonth.value + oForm.KnowledgeDateDay.value) > parseInt(oForm.Registration2DateYear.value + oForm.Registration2DateMonth.value + oForm.Registration2DateDay.value)) {" & vbNewLine
													Response.Write "alert('La fecha de registro en bolsa de trabajo no puede ser anterior a la fecha de evaluación de conocimientos.');" & vbNewLine
													Response.Write "oForm.Registration2DateDay.focus();" & vbNewLine
													Response.Write "return false;" & vbNewLine
												Response.Write "}" & vbNewLine
											Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine

									Response.Write "return true;" & vbNewLine
								Response.Write "} // End of CheckKardexRequirements" & vbNewLine

								Response.Write "function ShowKardexRequirements(oForm, sRequirementsTypeID) {" & vbNewLine
									Response.Write "var oRegExp = null;" & vbNewLine
									Response.Write "var i=0;" & vbNewLine

									Response.Write "UnselectAllItemsFromList(oForm.Requirements);" & vbNewLine

									Response.Write "oRegExp = eval('/,' + oForm.KardexTypeID.value + ',/gi');" & vbNewLine
									Response.Write "RemoveAllItemsFromList(null, oForm.KardexOriginID);" & vbNewLine
									Response.Write "for (i=0; i<KardexOrigins.length; i++)" & vbNewLine
										Response.Write "if (KardexOrigins[i][2].search(oRegExp) > -1) {" & vbNewLine
											Response.Write "AddItemToList(KardexOrigins[i][1], KardexOrigins[i][0], null, oForm.KardexOriginID);" & vbNewLine
										Response.Write "}" & vbNewLine

									Response.Write "oForm.Requirements.size = oForm.Requirements.options.length;" & vbNewLine
									Response.Write "UpdateDocumentsDate(oForm);" & vbNewLine

									Response.Write "RemoveAllItemsFromList(null, oForm.RequirementsTypeID);" & vbNewLine
									Response.Write "for (i=0; i<RequirementsTypes.length; i++)" & vbNewLine
										Response.Write "if (RequirementsTypes[i][2] == oForm.KardexTypeID.value) {" & vbNewLine
											Response.Write "AddItemToList(RequirementsTypes[i][1], RequirementsTypes[i][0], null, oForm.RequirementsTypeID);" & vbNewLine
										Response.Write "}" & vbNewLine

									Response.Write "SendURLValuesToForm('RequirementsTypeID=' + sRequirementsTypeID, oForm);" & vbNewLine

									Response.Write "oRegExp = eval('/,' + oForm.RequirementsTypeID.value + ',/gi');" & vbNewLine
									Response.Write "RemoveAllItemsFromList(null, oForm.Requirements);" & vbNewLine
									Response.Write "for (i=0; i<KardexRequirements.length; i++)" & vbNewLine
										Response.Write "if (KardexRequirements[i][2].search(oRegExp) > -1) {" & vbNewLine
											Response.Write "AddItemToList(KardexRequirements[i][1], KardexRequirements[i][0], null, oForm.Requirements);" & vbNewLine
										Response.Write "}" & vbNewLine

									Response.Write "oForm.Requirements.size = oForm.Requirements.options.length;" & vbNewLine
									Response.Write "UpdateDocumentsDate(oForm);" & vbNewLine
									Response.Write "if (document.CatalogFrm.KardexTypeID.value == '2')" & vbNewLine
										Response.Write "SearchRecord('2', 'PositionsByType', 'SearchPositionsIFrame', 'CatalogFrm.PositionID');" & vbNewLine
									Response.Write "else" & vbNewLine
										Response.Write "SearchRecord('1', 'PositionsByType', 'SearchPositionsIFrame', 'CatalogFrm.PositionID');" & vbNewLine
								Response.Write "} // End of ShowKardexRequirements" & vbNewLine

								Response.Write "function UpdateDocumentsDate(oForm) {" & vbNewLine
									Response.Write "var bReady = true;" & vbNewLine

									Response.Write "if (bDocReady) {" & vbNewLine
										Response.Write "oForm.DocumentsDateYear.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.DocumentsDateMonth.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.DocumentsDateDay.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.KnowledgeDateYear.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.KnowledgeDateMonth.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.KnowledgeDateDay.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.KnowledgeStatusID.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.PsychologicDateYear.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.PsychologicDateMonth.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.PsychologicDateDay.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.PsychologicStatusID.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.Registration1DateYear.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.Registration1DateMonth.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.Registration1DateDay.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.Registration2DateYear.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.Registration2DateMonth.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.Registration2DateDay.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.Registration3DateYear.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.Registration3DateMonth.options[0].selected = true;" & vbNewLine
										Response.Write "oForm.Registration3DateDay.options[0].selected = true;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "HideDisplay(document.all['CatalogFrm_DocumentsDateDiv']);" & vbNewLine
									Response.Write "HideDisplay(document.all['TempDiv']);" & vbNewLine
									Response.Write "HideDisplay(document.all['CatalogFrm_KnowledgeDateDiv']);" & vbNewLine
									Response.Write "HideDisplay(document.all['CatalogFrm_KnowledgeStatusIDDiv']);" & vbNewLine
									Response.Write "HideDisplay(document.all['CatalogFrm_PsychologicDateDiv']);" & vbNewLine
									Response.Write "HideDisplay(document.all['CatalogFrm_PsychologicStatusIDDiv']);" & vbNewLine
									Response.Write "HideDisplay(document.all['CatalogFrm_Registration1DateDiv']);" & vbNewLine
									Response.Write "HideDisplay(document.all['CatalogFrm_Registration2DateDiv']);" & vbNewLine
									Response.Write "HideDisplay(document.all['CatalogFrm_Registration3DateDiv']);" & vbNewLine

									Response.Write "if (oForm.Requirements.options.length > 0) {" & vbNewLine
										Response.Write "for (var i=0; i<oForm.Requirements.options.length; i++) {" & vbNewLine
											Response.Write "if (oForm.RequirementsTypeID.value != '7') {" & vbNewLine
												Response.Write "if (! oForm.Requirements.options[i].selected)" & vbNewLine
													Response.Write "for (var j=0; j<KardexRequirements.length; j++)" & vbNewLine
														Response.Write "if ((KardexRequirements[j][0] == oForm.Requirements.options[i].value) && (KardexRequirements[j][3] == '0')) {" & vbNewLine
															Response.Write "if (bDocReady) {" & vbNewLine
																Response.Write "oForm.DocumentsDateYear.options[0].selected = true;" & vbNewLine
																Response.Write "oForm.DocumentsDateMonth.options[0].selected = true;" & vbNewLine
																Response.Write "oForm.DocumentsDateDay.options[0].selected = true;" & vbNewLine
															Response.Write "}" & vbNewLine
															Response.Write "bReady = false;" & vbNewLine
															Response.Write "break;" & vbNewLine
														Response.Write "}" & vbNewLine
											Response.Write "} else {" & vbNewLine
												Response.Write "var bReady = false;" & vbNewLine
												Response.Write "if (oForm.Requirements.options[i].selected) {" & vbNewLine
													Response.Write "bReady = true;" & vbNewLine
													Response.Write "break;" & vbNewLine
												Response.Write "}" & vbNewLine
											Response.Write "}" & vbNewLine
										Response.Write "}" & vbNewLine

										Response.Write "if (bReady) {" & vbNewLine
											Response.Write "SetDateCombos('', '', '', oForm.DocumentsDateYear, oForm.DocumentsDateMonth, oForm.DocumentsDateDay)" & vbNewLine
											Response.Write "ShowDisplay(document.all['CatalogFrm_DocumentsDateDiv']);" & vbNewLine
											Response.Write "ShowDisplay(document.all['TempDiv']);" & vbNewLine
											Response.Write "ShowDisplay(document.all['CatalogFrm_KnowledgeDateDiv']);" & vbNewLine
											Response.Write "ShowDisplay(document.all['CatalogFrm_KnowledgeStatusIDDiv']);" & vbNewLine
											Response.Write "UpdatePsychologicDate(oForm);" & vbNewLine
											If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11)) = 0 Then
												Response.Write "if (oForm.RequirementsTypeID.value == '7')" & vbNewLine
													Response.Write "alert('Para un Trabajador (CR) de Base, Profesional (P), es necesario al menos un documento.\nPara continuar guarde los cambios presionando el botón ""Modificar"".\nEsto desactivará los campos de arriba y sólo le permitirá modificar la información\nde la sección ""Procesos de selección de personal""');" & vbNewLine
												Response.Write "else" & vbNewLine
													Response.Write "alert('Todos los requisitos documentales que son obligatorios han sido seleccionados.\nPara continuar guarde los cambios presionando el botón ""Modificar"".\nEsto desactivará los campos de arriba y sólo le permitirá modificar la información\nde la sección ""Procesos de selección de personal""');" & vbNewLine
											End If
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "} // End of UpdateDocumentsDate" & vbNewLine

								Response.Write "function UpdateKnowledgeStatus(oForm) {" & vbNewLine
									Response.Write "if ((! oForm.KnowledgeDateYear.options[0].selected) && (! oForm.KnowledgeDateMonth.options[0].selected) && (! oForm.KnowledgeDateDay.options[0].selected) && (oForm.KnowledgeStatusID.options[0].selected)) {" & vbNewLine
										Response.Write "oForm.KnowledgeStatusID.options[2].selected = true;" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "} // End of UpdateKnowledgeStatus" & vbNewLine

								Response.Write "function UpdatePsychologicDate(oForm) {" & vbNewLine
									Response.Write "HideDisplay(document.all['CatalogFrm_PsychologicDateDiv']);" & vbNewLine
									Response.Write "HideDisplay(document.all['CatalogFrm_PsychologicStatusIDDiv']);" & vbNewLine
									Response.Write "HideDisplay(document.all['CatalogFrm_Registration1DateDiv']);" & vbNewLine
									Response.Write "HideDisplay(document.all['CatalogFrm_Registration2DateDiv']);" & vbNewLine
									Response.Write "HideDisplay(document.all['CatalogFrm_Registration3DateDiv']);" & vbNewLine

									Response.Write "switch (oForm.KnowledgeStatusID.value) {" & vbNewLine
										Response.Write "case '1':" & vbNewLine 'Aprobado
											Response.Write "if (oForm.KardexTypeID.value != '3') {" & vbNewLine
												Response.Write "ShowDisplay(document.all['CatalogFrm_PsychologicDateDiv']);" & vbNewLine
												Response.Write "ShowDisplay(document.all['CatalogFrm_PsychologicStatusIDDiv']);" & vbNewLine
											Response.Write "} else {" & vbNewLine
												Response.Write "oForm.PsychologicStatusID.options[0].selected = true;" & vbNewLine
											Response.Write "}" & vbNewLine
											Response.Write "switch (oForm.KardexTypeID.value) {" & vbNewLine
												Response.Write "case '1':" & vbNewLine 'Nuevo ingreso (NI): Base
													Response.Write "ShowDisplay(document.all['CatalogFrm_Registration1DateDiv']);" & vbNewLine
													Response.Write "if (bDocReady) {" & vbNewLine
														Response.Write "oForm.Registration2DateYear.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration2DateMonth.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration2DateDay.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration3DateYear.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration3DateMonth.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration3DateDay.options[0].selected = true;" & vbNewLine
													Response.Write "}" & vbNewLine
													Response.Write "break;" & vbNewLine
												Response.Write "case '2':" & vbNewLine 'Nuevo ingreso (NI): Confianza
													Response.Write "ShowDisplay(document.all['CatalogFrm_Registration3DateDiv']);" & vbNewLine
													Response.Write "if (bDocReady) {" & vbNewLine
														Response.Write "oForm.Registration1DateYear.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration1DateMonth.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration1DateDay.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration2DateYear.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration2DateMonth.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration2DateDay.options[0].selected = true;" & vbNewLine
													Response.Write "}" & vbNewLine
													Response.Write "break;" & vbNewLine
												Response.Write "case '3':" & vbNewLine 'Trabajador (CR): Base
													Response.Write "ShowDisplay(document.all['CatalogFrm_Registration2DateDiv']);" & vbNewLine
													Response.Write "if (bDocReady) {" & vbNewLine
														Response.Write "oForm.Registration1DateYear.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration1DateMonth.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration1DateDay.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration3DateYear.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration3DateMonth.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration3DateDay.options[0].selected = true;" & vbNewLine
													Response.Write "}" & vbNewLine
													Response.Write "break;" & vbNewLine
											Response.Write "}" & vbNewLine
											Response.Write "break;" & vbNewLine
										Response.Write "case '0':" & vbNewLine 'Reprobado
											Response.Write "if (bDocReady) {" & vbNewLine
												Response.Write "oForm.PsychologicStatusID.options[0].selected = true;" & vbNewLine
												Response.Write "oForm.PsychologicDateYear.options[0].selected = true;" & vbNewLine
												Response.Write "oForm.PsychologicDateMonth.options[0].selected = true;" & vbNewLine
												Response.Write "oForm.PsychologicDateDay.options[0].selected = true;" & vbNewLine
											Response.Write "}" & vbNewLine
											Response.Write "switch (oForm.KardexTypeID.value) {" & vbNewLine
												Response.Write "case '1':" & vbNewLine 'Nuevo ingreso (NI): Base
													Response.Write "ShowDisplay(document.all['CatalogFrm_Registration1DateDiv']);" & vbNewLine
													Response.Write "if (bDocReady) {" & vbNewLine
														Response.Write "oForm.Registration2DateYear.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration2DateMonth.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration2DateDay.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration3DateYear.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration3DateMonth.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration3DateDay.options[0].selected = true;" & vbNewLine
													Response.Write "}" & vbNewLine
													Response.Write "break;" & vbNewLine
												Response.Write "case '2':" & vbNewLine 'Nuevo ingreso (NI): Confianza
													Response.Write "ShowDisplay(document.all['CatalogFrm_Registration3DateDiv']);" & vbNewLine
													Response.Write "if (bDocReady) {" & vbNewLine
														Response.Write "oForm.Registration1DateYear.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration1DateMonth.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration1DateDay.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration2DateYear.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration2DateMonth.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration2DateDay.options[0].selected = true;" & vbNewLine
													Response.Write "}" & vbNewLine
													Response.Write "break;" & vbNewLine
												Response.Write "case '3':" & vbNewLine 'Trabajador (CR): Base
													Response.Write "ShowDisplay(document.all['CatalogFrm_Registration2DateDiv']);" & vbNewLine
													Response.Write "if (bDocReady) {" & vbNewLine
														Response.Write "oForm.Registration1DateYear.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration1DateMonth.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration1DateDay.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration3DateYear.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration3DateMonth.options[0].selected = true;" & vbNewLine
														Response.Write "oForm.Registration3DateDay.options[0].selected = true;" & vbNewLine
													Response.Write "}" & vbNewLine
													Response.Write "break;" & vbNewLine
											Response.Write "}" & vbNewLine
											Response.Write "break;" & vbNewLine
										Response.Write "default:" & vbNewLine
											Response.Write "if (bDocReady) {" & vbNewLine
												Response.Write "oForm.PsychologicStatusID.options[0].selected = true;" & vbNewLine
												Response.Write "oForm.PsychologicDateYear.options[0].selected = true;" & vbNewLine
												Response.Write "oForm.PsychologicDateMonth.options[0].selected = true;" & vbNewLine
												Response.Write "oForm.PsychologicDateDay.options[0].selected = true;" & vbNewLine
												Response.Write "oForm.Registration1DateYear.options[0].selected = true;" & vbNewLine
												Response.Write "oForm.Registration1DateMonth.options[0].selected = true;" & vbNewLine
												Response.Write "oForm.Registration1DateDay.options[0].selected = true;" & vbNewLine
												Response.Write "oForm.Registration2DateYear.options[0].selected = true;" & vbNewLine
												Response.Write "oForm.Registration2DateMonth.options[0].selected = true;" & vbNewLine
												Response.Write "oForm.Registration2DateDay.options[0].selected = true;" & vbNewLine
												Response.Write "oForm.Registration3DateYear.options[0].selected = true;" & vbNewLine
												Response.Write "oForm.Registration3DateMonth.options[0].selected = true;" & vbNewLine
												Response.Write "oForm.Registration3DateDay.options[0].selected = true;" & vbNewLine
											Response.Write "}" & vbNewLine
											Response.Write "break;" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "} // End of UpdatePsychologicDate" & vbNewLine

								Response.Write "function UpdatePsychologicStatus(oForm) {" & vbNewLine
									Response.Write "if ((! oForm.PsychologicDateYear.options[0].selected) && (! oForm.PsychologicDateMonth.options[0].selected) && (! oForm.PsychologicDateDay.options[0].selected) && (oForm.PsychologicStatusID.options[0].selected)) {" & vbNewLine
										Response.Write "oForm.PsychologicStatusID.options[2].selected = true;" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "} // End of UpdatePsychologicStatus" & vbNewLine

								Response.Write "function UpdateRegistrationDate(oForm, iRegistrationID) {" & vbNewLine
'									Response.Write "HideDisplay(document.all['CatalogFrm_Registration1DateDiv']);" & vbNewLine
'									Response.Write "HideDisplay(document.all['CatalogFrm_Registration3DateDiv']);" & vbNewLine
'
'									Response.Write "if (oForm.PsychologicStatusID.value == '1') {" & vbNewLine
'										Response.Write "if (oForm.RequirementsTypeID.value != '8') {" & vbNewLine
'											Response.Write "ShowDisplay(document.all['CatalogFrm_Registration1DateDiv']);" & vbNewLine
'											Response.Write "ShowDisplay(document.all['CatalogFrm_Registration3DateDiv']);" & vbNewLine
'										Response.Write "}" & vbNewLine
'										Response.Write "ShowDisplay(document.all['CatalogFrm_Registration2DateDiv']);" & vbNewLine
'									Response.Write "}" & vbNewLine
'
'									Response.Write "switch (iRegistrationID) {" & vbNewLine
'										Response.Write "case 1:" & vbNewLine
'											Response.Write "if ((oForm.Registration1DateDay.value != '0') || (oForm.Registration1DateMonth.value != '0') || (oForm.Registration1DateYear.value != '0')) {" & vbNewLine
'												Response.Write "HideDisplay(document.all['CatalogFrm_Registration2DateDiv']);" & vbNewLine
'												Response.Write "HideDisplay(document.all['CatalogFrm_Registration3DateDiv']);" & vbNewLine
'											Response.Write "}" & vbNewLine
'											Response.Write "break;" & vbNewLine
'										Response.Write "case 2:" & vbNewLine
'											Response.Write "if ((oForm.Registration2DateDay.value != '0') || (oForm.Registration2DateMonth.value != '0') || (oForm.Registration2DateYear.value != '0')) {" & vbNewLine
'												Response.Write "HideDisplay(document.all['CatalogFrm_Registration1DateDiv']);" & vbNewLine
'												Response.Write "HideDisplay(document.all['CatalogFrm_Registration3DateDiv']);" & vbNewLine
'											Response.Write "}" & vbNewLine
'											Response.Write "break;" & vbNewLine
'										Response.Write "case 3:" & vbNewLine
'											Response.Write "if ((oForm.Registration3DateDay.value != '0') || (oForm.Registration3DateMonth.value != '0') || (oForm.Registration3DateYear.value != '0')) {" & vbNewLine
'												Response.Write "HideDisplay(document.all['CatalogFrm_Registration1DateDiv']);" & vbNewLine
'												Response.Write "HideDisplay(document.all['CatalogFrm_Registration2DateDiv']);" & vbNewLine
'											Response.Write "}" & vbNewLine
'											Response.Write "break;" & vbNewLine
'									Response.Write "}" & vbNewLine
								Response.Write "} // End of UpdateRegistrationDate" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine

							If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)) > -1 Then lErrorNumber = GetCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
							If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11)) > 0 Then
								aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) = "1,2,3,4,5,6,7,8,9,10,11"
							ElseIf CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)) > -1 Then
								aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) = "2"
							End If

							If (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12)) > 0) And (InStr(1, "0,1", aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13), vbBinaryCompare) > 0) Then
								aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) = aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",12"
								aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG)(13) = 1
								aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(13) = "StatusKnowledges;,;StatusID;,;StatusName;,;(StatusID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13) & ");,;StatusName;,;;,;Ninguno;;;-1"
							End If
							If (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14)) > 0) And (InStr(1, "0,1", aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(15), vbBinaryCompare) > 0) Then
								aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) = aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",14"
								aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG)(15) = 1
								aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(15) = "StatusPsychologics;,;StatusID;,;StatusName;,;(StatusID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(15) & ");,;StatusName;,;;,;Ninguno;;;-1"
							End If
							If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(16)) > 0 Then aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) = aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",16"
							If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(17)) > 0 Then aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) = aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",17"
							If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(18)) > 0 Then aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) = aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",18"
							If (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(16)) > 0) Or (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(17)) > 0) Or (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(18)) > 0) Then
								Call DisplayInstructionsMessage("Proceso concluido", "Se ha concluido el proceso y a partir de este momento el registro solo será de consulta")
							End If
							
							lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
							If (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11)) = 0) And (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)) > -1) Then
								Call DisplayErrorMessage("Estatus", "<B>Pendiente de entregar sus requerimientos documentales.</B>")
							End If

							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								If (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(16)) > 0) Or (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(17)) > 0) Or (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(18)) > 0) Then
									Response.Write "HideDisplay(document.CatalogFrm.Modify);" & vbNewLine
								End If
								If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)) = -1 Then
									Response.Write "document.CatalogFrm.StartDateMonth.options[" & Month(Date()) - 1 & "].selected = true;" & vbNewLine
									Response.Write "document.CatalogFrm.StartDateDay.options[" & Day(Date()) - 1 & "].selected = true;" & vbNewLine
								End If
								Response.Write "ShowKardexRequirements(document.CatalogFrm, '" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(9) & "');" & vbNewLine
								Response.Write "SendURLValuesToForm('Requirements=" & Replace(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10), " ", "") & "', document.CatalogFrm);" & vbNewLine
								Response.Write "UpdateDocumentsDate(document.CatalogFrm);" & vbNewLine
								aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11) = Right(("00000000" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11)), Len("00000000"))
								Response.Write "SendURLValuesToForm('DocumentsDateYear=" & Left(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11), Len("0000")) & "&DocumentsDateMonth=" & Mid(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11), Len("00000"), Len("00")) & "&DocumentsDateDay=" & Right(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11), Len("00")) & "', document.CatalogFrm);" & vbNewLine
								Response.Write "UpdatePsychologicDate(document.CatalogFrm);" & vbNewLine
								If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)) > -1 Then
									Response.Write "SendURLValuesToForm('"
										If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2)) > 0 Then Response.Write "StartDateYear=" & Left(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2), Len("0000")) & "&StartDateMonth=" & Mid(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2), Len("00000"), Len("00")) & "&StartDateDay=" & Right(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2), Len("00"))
										If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11)) > 0 Then Response.Write "DocumentsDateYear=" & Left(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11), Len("0000")) & "&DocumentsDateMonth=" & Mid(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11), Len("00000"), Len("00")) & "&DocumentsDateDay=" & Right(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11), Len("00"))
										If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12)) > 0 Then Response.Write "KnowledgeDateYear=" & Left(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12), Len("0000")) & "&KnowledgeDateMonth=" & Mid(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12), Len("00000"), Len("00")) & "&KnowledgeDateDay=" & Right(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12), Len("00"))
										If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14)) > 0 Then Response.Write "PsychologicDateYear=" & Left(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14), Len("0000")) & "&PsychologicDateMonth=" & Mid(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14), Len("00000"), Len("00")) & "&PsychologicDateDay=" & Right(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2), Len("00"))
										If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(16)) > 0 Then Response.Write "Registration1DateYear=" & Left(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(16), Len("0000")) & "&Registration1DateMonth=" & Mid(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(16), Len("00000"), Len("00")) & "&Registration1DateDay=" & Right(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2), Len("00"))
										If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(17)) > 0 Then Response.Write "Registration2DateYear=" & Left(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(17), Len("0000")) & "&Registration2DateMonth=" & Mid(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(17), Len("00000"), Len("00")) & "&Registration2DateDay=" & Right(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2), Len("00"))
										If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(18)) > 0 Then Response.Write "Registration3DateYear=" & Left(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(18), Len("0000")) & "&Registration3DateMonth=" & Mid(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(18), Len("00000"), Len("00")) & "&Registration3DateDay=" & Right(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2), Len("00"))
									Response.Write "', document.CatalogFrm);" & vbNewLine
									Response.Write "window.setTimeout(""SendURLValuesToForm('a=1"
										If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2)) > 0 Then Response.Write "&StartDateDay=" & Right(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2), Len("00"))
										If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11)) > 0 Then Response.Write "&DocumentsDateDay=" & Right(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11), Len("00"))
										If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12)) > 0 Then Response.Write "&KnowledgeDateDay=" & Right(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12), Len("00"))
										If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14)) > 0 Then Response.Write "&PsychologicDateDay=" & Right(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14), Len("00"))
										If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(16)) > 0 Then Response.Write "&Registration1DateDay=" & Right(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(16), Len("00"))
										If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(17)) > 0 Then Response.Write "&Registration2DateDay=" & Right(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(17), Len("00"))
										If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(18)) > 0 Then Response.Write "&Registration3DateDay=" & Right(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(18), Len("00"))
									Response.Write "', document.CatalogFrm)"", 500);" & vbNewLine
								End If
								Response.Write "document.CatalogFrm.ModifyDate.value='" & Left(GetSerialNumberForDate(""), Len("00000000")) & "';" & vbNewLine
								Response.Write "bDocReady = true;" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine

							'If (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2)) > 0) And (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2)) < (CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) - 600)) And (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11)) = 0) Then
							If (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2)) > 0) And (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2)) < (CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) - 600)) Then
								Call DisplayErrorMessage("Trámite expirado", "Ya han pasado más de seis meses desde que se inició este trámite.")
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									Response.Write "HideDisplay(document.CatalogFrm.Modify);" & vbNewLine
								Response.Write "//--></SCRIPT>" & vbNewLine
							End If
						End If
					End If
				Case 352 'Desarrollo humano > Selección de personal > Validación del proceso de selección de personal
					If (Not bSearchForm) And (Not bShowForm) And (Not bAction) And (Len(oRequest("DoSearch").Item) = 0) Then
						aMenuComponent(A_ELEMENTS_MENU) = Array(_
							Array("Registro de información",_
								  "Alta de la información del proceso de selección",_
								  "Images/MnAreas.gif", "Main_ISSSTE.asp?SectionID=352&New=1", True),_
							Array("Búsqueda de información",_
								  "Listado de las personas registradas en el proceso de selección",_
								  "Images/MnJobs.gif", "Main_ISSSTE.asp?SectionID=352&Search=1", True)_
						)
					Else
						If (lErrorNumber = 0) And bAction Then
							Call DisplayErrorMessage("Confirmación", "La información fue guardada con éxito.<BR /><FORM><INPUT TYPE=""BUTTON"" VALUE=""Registro de información"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=352&New=1'"" /><IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" /><INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=352'"" /></FORM>")
							Response.Write "<BR />"
						End If
						If lErrorNumber <> 0 Then
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							Response.Write "<BR />"
						End If

						If bSearchForm Then
							lErrorNumber = Display352SearchForm(oRequest, oADODBConnection, sErrorDescription)
						ElseIf Len(oRequest("DoSearch").Item) > 0 Then
							'lErrorNumber = DisplayCatalogsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, aCatalogComponent, sErrorDescription)
							lErrorNumber = Display352SearchResults(oRequest, oADODBConnection, False, sErrorDescription)
							If lErrorNumber = L_ERR_NO_RECORDS Then
								Call DisplayErrorMessage("Búsqueda vacía", "No existen registros que cumplan con los criterios de la búsqueda")
								Response.Write "<BR />"
								lErrorNumber = Display352SearchForm(oRequest, oADODBConnection, sErrorDescription)
							End If
						ElseIf bShowForm Then
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "function ShowRequirements(sPositionTypeID) {" & vbNewLine
									Response.Write "ShowDisplay(document.all['CatalogFrm_Requirement7Div']);" & vbNewLine
									Response.Write "ShowDisplay(document.all['CatalogFrm_Requirement8Div']);" & vbNewLine
									Response.Write "ShowDisplay(document.all['CatalogFrm_Requirement9Div']);" & vbNewLine
									Response.Write "ShowDisplay(document.all['CatalogFrm_Requirement10Div']);" & vbNewLine
									Response.Write "ShowDisplay(document.all['CatalogFrm_Requirement11Div']);" & vbNewLine
									Response.Write "switch(sPositionTypeID) {" & vbNewLine
										Response.Write "case '1':" & vbNewLine 'Especialista
											Response.Write "HideDisplay(document.all['CatalogFrm_Requirement7Div']);" & vbNewLine
											Response.Write "HideDisplay(document.all['CatalogFrm_Requirement8Div']);" & vbNewLine
											Response.Write "HideDisplay(document.all['CatalogFrm_Requirement9Div']);" & vbNewLine
											Response.Write "HideDisplay(document.all['CatalogFrm_Requirement10Div']);" & vbNewLine
											Response.Write "break;" & vbNewLine
										Response.Write "case '2':" & vbNewLine 'Profesional
											Response.Write "HideDisplay(document.all['CatalogFrm_Requirement7Div']);" & vbNewLine
											Response.Write "HideDisplay(document.all['CatalogFrm_Requirement8Div']);" & vbNewLine
											Response.Write "HideDisplay(document.all['CatalogFrm_Requirement11Div']);" & vbNewLine
											Response.Write "break;" & vbNewLine
										Response.Write "case '3':" & vbNewLine 'Técnico
											Response.Write "HideDisplay(document.all['CatalogFrm_Requirement9Div']);" & vbNewLine
											Response.Write "HideDisplay(document.all['CatalogFrm_Requirement10Div']);" & vbNewLine
											Response.Write "HideDisplay(document.all['CatalogFrm_Requirement11Div']);" & vbNewLine
											Response.Write "break;" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "} // End of ShowRequirements" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine

							lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)

							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "ShowRequirements(document.CatalogFrm.PositionTypeID.value);" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine
						End If
					End If
				Case 353 'Desarrollo humano > Selección de personal > Búsqueda de información del proceso de selección
					If lErrorNumber <> 0 Then
						Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
						lErrorNumber = 0
						Response.Write "<BR />"
					End If

					If bSearchForm Then
						lErrorNumber = Display353SearchForm(oRequest, oADODBConnection, sErrorDescription)
					ElseIf Len(oRequest("DoSearch").Item) > 0 Then
						lErrorNumber = Display353SearchResults(oRequest, oADODBConnection, False, sErrorDescription)
						If lErrorNumber = L_ERR_NO_RECORDS Then
							Call DisplayErrorMessage("Búsqueda vacía", "No existen registros que cumplan con los criterios de la búsqueda")
							Response.Write "<BR />"
							lErrorNumber = Display353SearchForm(oRequest, oADODBConnection, sErrorDescription)
						End If
					End If
				Case 354 'Desarrollo humano > Selección de personal > Búsqueda de información de la bolsa de trabajo
					Call DisplayErrorMessage("Mensaje del sistema", "No existe información registrada en el sistema por parte del área de prestaciones.")
				Case 356 'Desarrollo humano > Selección de personal > Búsqueda de información de escalafón
					If True Then
						Call DisplayErrorMessage("Mensaje del sistema", "No existe información registrada en el sistema por parte del área de prestaciones.")
					Else
					If (Not bSearchForm) And (Not bShowForm) And (Not bAction) And (Len(oRequest("DoSearch").Item) = 0) Then
						aMenuComponent(A_ELEMENTS_MENU) = Array(_
							Array("Registro de información",_
								  "Alta de la información del registros de escalafón",_
								  "Images/MnAreas.gif", "Main_ISSSTE.asp?SectionID=356&New=1", True),_
							Array("Búsqueda de información",_
								  "Listado de los empleados en el registro de escalafón",_
								  "Images/MnJobs.gif", "Main_ISSSTE.asp?SectionID=356&Search=1", True)_
						)
					Else
						If (lErrorNumber = 0) And bAction Then
							Call DisplayErrorMessage("Confirmación", "La información fue guardada con éxito.<BR /><FORM><INPUT TYPE=""BUTTON"" VALUE=""Registro de información"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=356&New=1'"" /><IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" /><INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=356'"" /></FORM>")
							Response.Write "<BR />"
						End If
						If lErrorNumber <> 0 Then
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							Response.Write "<BR />"
						End If

						If bSearchForm Then
							lErrorNumber = Display356SearchForm(oRequest, oADODBConnection, sErrorDescription)
						ElseIf Len(oRequest("DoSearch").Item) > 0 Then
							lErrorNumber = Display356SearchResults(oRequest, oADODBConnection, True, sErrorDescription)
							If lErrorNumber = L_ERR_NO_RECORDS Then
								Call DisplayErrorMessage("Búsqueda vacía", "No existen registros que cumplan con los criterios de la búsqueda")
								Response.Write "<BR />"
								lErrorNumber = Display356SearchForm(oRequest, oADODBConnection, sErrorDescription)
							End If
						ElseIf bShowForm Then
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "function CheckEmployeeNumber(oForm) {" & vbNewLine
									Response.Write "if (oForm)" & vbNewLine
										Response.Write "if ((oForm.EmployeeID.value == '') || (oForm.EmployeeID.value == '-1')) {" & vbNewLine
											Response.Write "alert('Introduzca el número de empleado y verifique su existencia en el sistema');" & vbNewLine
											Response.Write "document.EmployeeFrm.EmployeeNumber.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "return true;" & vbNewLine
								Response.Write "} // End of CheckEmployeeNumber" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine

							Response.Write "<FORM NAME=""EmployeeFrm"" ID=""EmployeeFrm"" onSubmit=""return false"">"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">No. del empleado:&nbsp;</FONT>"
							    Response.Write "<INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""
									If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) > -1 Then Response.Write aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG))
							    Response.Write """ CLASS=""TextFields"" onChange=""document.CatalogFrm.EmployeeID.value='-1';"" />"
							    Response.Write "<A HREF=""javascript: document.CatalogFrm.EmployeeID.value='-1'; SearchRecord(document.EmployeeFrm.EmployeeNumber.value, 'EmployeeNumber', 'SearchEmployeeNumberIFrame', 'CatalogFrm.EmployeeID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>&nbsp;"
							    Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""22""></IFRAME>"
							Response.Write "</FORM>"
							lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
						End If
					End If
					End If
				Case 36 'Desarrollo humano > Desarrollo humano
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Registro de detección de necesidades",_
							  "",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=369", True),_
						Array("Registro de lista de cursos",_
							  "",_
							  "Images/MnLeftArrows.gif", "SADE.asp?SectionID=361", True),_
						Array("Registro de diplomados",_
							  "",_
							  "Images/MnLeftArrows.gif", "SADE.asp?SectionID=361&Diploma=1", True),_
						Array("Cursos del programa anual de capacitación",_
							  "",_
							  "Images/MnLeftArrows.gif", "SADE.asp?SectionID=362", True),_
						Array("Registro de personal para capacitación",_
							  "",_
							  "Images/MnLeftArrows.gif", "SADE.asp?SectionID=363", True),_
						Array("Seguimiento al programa autorizado de capacitación",_
							  "",_
							  "Images/MnLeftArrows.gif", "SADE.asp?SectionID=365", True),_
						Array("SADE",_
							  "Sistema de Administración de la Educación",_
							  "Images/MnLeftArrows.gif", "javascript: OpenNewWindow('/SADE/Default.asp?SessionID=" & GenerateRandomNumbersSecuence(100) & "&FromSAPo=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&Password=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&Login=1', null, 'SADE', 1012, 717, 'yes', 'yes')", True),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Reportes",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1364", True),_
						Array("Reporte de curriculum por empleado",_
							  "",_
							  "Images/MnLeftArrows.gif", "SADE.asp?SectionID=367", True),_
						Array("Generación de diplomas de reconocimiento",_
							  "",_
							  "Images/MnLeftArrows.gif", "SADE.asp?SectionID=366", True)_
					)
				Case 369 'Desarrollo humano > Desarrollo humano > Registro de detección de necesidades
					If (Not bSearchForm) And (Not bShowForm) And (Not bAction) And (Len(oRequest("DoSearch").Item) = 0) Then
						aMenuComponent(A_ELEMENTS_MENU) = Array(_
							Array("Registro de información",_
								  "Alta de la información del registros de escalafón",_
								  "Images/MnAreas.gif", "Main_ISSSTE.asp?SectionID=369&New=1", True),_
							Array("Búsqueda de información",_
								  "Listado de los empleados en el registro de escalafón",_
								  "Images/MnJobs.gif", "Main_ISSSTE.asp?SectionID=369&Search=1", True)_
						)
					Else
						If (lErrorNumber = 0) And bAction Then
							Call DisplayErrorMessage("Confirmación", "La información fue guardada con éxito.<BR /><FORM><INPUT TYPE=""BUTTON"" VALUE=""Registro de información"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=369&New=1'"" /><IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" /><INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=369'"" /></FORM>")
							Response.Write "<BR />"
						End If
						If lErrorNumber <> 0 Then
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							Response.Write "<BR />"
						End If

						If bSearchForm Then
							lErrorNumber = Display356SearchForm(oRequest, oADODBConnection, sErrorDescription)
						ElseIf Len(oRequest("DoSearch").Item) > 0 Then
							lErrorNumber = Display369SearchResults(oRequest, oADODBConnection, False, sErrorDescription)
							If lErrorNumber = L_ERR_NO_RECORDS Then
								Call DisplayErrorMessage("Búsqueda vacía", "No existen registros que cumplan con los criterios de la búsqueda")
								Response.Write "<BR />"
								lErrorNumber = Display356SearchForm(oRequest, oADODBConnection, sErrorDescription)
							End If
						ElseIf bShowForm Then
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "function ShowNewCourseFields(sSchoolarshipID) {" & vbNewLine
									Response.Write "oForm = document.CatalogFrm;" & vbNewLine
									Response.Write "if (oForm) {" & vbNewLine
										Response.Write "RemoveItemByValueFromList('2', null, oForm.StatusID);" & vbNewLine
										Response.Write "if ((sSchoolarshipID == '6') || (sSchoolarshipID == '7')) {" & vbNewLine
											Response.Write "AddItemToList('Pasante', '2', null, oForm.StatusID);" & vbNewLine
											Response.Write "ShowDisplay(document.all['CatalogFrm_BSNameDiv']);" & vbNewLine
											Response.Write "ShowDisplay(document.all['CatalogFrm_UniversityNameDiv']);" & vbNewLine
											Response.Write "ShowDisplay(document.all['CatalogFrm_BSYearsDiv']);" & vbNewLine
											Response.Write "ShowDisplay(document.all['SemestersText1Div']);" & vbNewLine
											Response.Write "HideDisplay(document.all['SemestersText2Div']);" & vbNewLine
										Response.Write "} else {" & vbNewLine
											Response.Write "HideDisplay(document.all['CatalogFrm_BSNameDiv']);" & vbNewLine
											Response.Write "HideDisplay(document.all['CatalogFrm_UniversityNameDiv']);" & vbNewLine
											Response.Write "HideDisplay(document.all['CatalogFrm_BSYearsDiv']);" & vbNewLine
											Response.Write "HideDisplay(document.all['SemestersText1Div']);" & vbNewLine
											Response.Write "ShowDisplay(document.all['SemestersText2Div']);" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "} // End of ShowNewCourseFields" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine

							Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
								Response.Write "<TD VALIGN=""TOP"">"
									Response.Write "<FORM NAME=""EmployeeFrm"" ID=""EmployeeFrm"" onSubmit=""return false"">"
										Response.Write "<FONT FACE=""Arial"" SIZE=""2"">No. del empleado:&nbsp;</FONT>"
									    Response.Write "<INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""
											If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) > -1 Then Response.Write aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG))
									    Response.Write """ CLASS=""TextFields"" onChange=""document.CatalogFrm.EmployeeID.value='-1';"" />"
										Response.Write "<A HREF=""javascript: document.CatalogFrm.EmployeeID.value='-1'; SearchRecord(document.EmployeeFrm.EmployeeNumber.value, 'EmployeesInfo', 'EmployeeInfoIFrame', 'CatalogFrm.EmployeeID');""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>"
									Response.Write "</FORM>"
									lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
								Response.Write "</TD>"
								Response.Write "<TD VALIGN=""TOP"">"
									Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""EmployeeInfoIFrame"" FRAMEBORDER=""0"" WIDTH=""250"" HEIGHT=""300""></IFRAME>"
								Response.Write "</TD>"
							Response.Write "</TR></TABLE>"
								
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "ShowNewCourseFields('" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2) & "');" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine
						End If
					End If
				Case 37 'Desarrollo humano > Planeación de recursos humanos
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Registro de procedimientos",_
							  "Suba y actualice los documentos digitales con los procedimientos.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=371", True),_
						Array("Alta y modificación de procesos",_
							  "Registre y administre los procesos del tablero de control.",_
							  "Images/MnLeftArrows.gif", "TaCo.asp?Action=Projects", True),_
						Array("Seguimiento de procesos",_
							  "Revise y actualice el avance de los procesos registrados en el tablero de control.",_
							  "Images/MnLeftArrows.gif", "Projects.asp", True)_
					)
				Case 371 'Desarrollo humano > Planeación de recursos humanos > Registro de procedimientos
					If bShowForm Then
						Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							Response.Write "function IsFileReady(oForm) {" & vbNewLine
								Response.Write "if (oForm.FilePath.value == '') {" & vbNewLine
									Response.Write "alert('No se ha guardado el archivo del procedimiento');" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "} else {" & vbNewLine
									Response.Write "oForm.FileType.value = oForm.FilePath.value.substr(oForm.FilePath.value.search(/\./gi) + 1);" & vbNewLine
								Response.Write "}" & vbNewLine

								Response.Write "return true;" & vbNewLine
							Response.Write "} // End of IsFileReady" & vbNewLine
						Response.Write "//--></SCRIPT>" & vbNewLine
						
						If Len(oRequest("Delete").Item) = 0 Then Response.Write "<IFRAME SRC=""BrowserFileForInfo.asp?Action=Documents&NoOriginalFile=1&UserID=" & aLoginComponent(N_USER_ID_LOGIN) & """ NAME=""UploadInfoIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""60""></IFRAME><BR />"
						Response.Write "<B>Registro de procedimientos</B><BR />"
						lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
					Else
						lErrorNumber = Display371SearchResults(oRequest, oADODBConnection, False, sErrorDescription)
						If lErrorNumber <> 0 Then
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							Response.Write "<BR />"
						End If
					End If
				Case 38 'Desarrollo humano > Búsqueda de centros de trabajo y centros de pago
					If Len(oRequest("DoSearch").Item) = 0 Then
						lErrorNumber = Display38SearchForm(oRequest, oADODBConnection, sErrorDescription)
					Else
						lErrorNumber = Display38SearchResults(oRequest, oADODBConnection, sErrorDescription)
					End If
				Case 4 'Informática
					asLockForPayroll = CStr(Application.Contents("SIAP_CalculatePayroll")) & LIST_SEPARATOR & LIST_SEPARATOR & LIST_SEPARATOR
					asLockForPayroll = Split(asLockForPayroll, LIST_SEPARATOR)
					Call GetNameFromTable(oADODBConnection, "Users", asLockForPayroll(0), "", "", sUserName, sErrorDescription)
					Call GetNameFromTable(oADODBConnection, "Payrolls", asLockForPayroll(1), "", "", sPayrollName, sErrorDescription)
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Conceptos de pago",_
							  "Administrar los valores de los conceptos de pago y por tipos de tabulador.",_
							  "Images/MnBudget.gif", "Payroll.asp?Action=Concepts", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_ConceptosDePago & ",", vbBinaryCompare) > 0),_
						Array("Empleados",_
							  "Administre la información de los empleados.",_
							  "Images/MnEmployees.gif", "Main_ISSSTE.asp?SectionID=42", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_Empleados & ",", vbBinaryCompare) > 0),_
						Array("Crear una nueva nómina",_
							  "Se creará el registro de la nómina para que los administradores responsables puedan registrar movimientos de personal e incidencias.",_
							  "Images/MnPayroll.gif", "Payroll.asp?Action=AddPayroll", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_CrearUnaNuevaNomina & ",", vbBinaryCompare) > 0),_
						Array("Modificar nómina",_
							  "Seleccione la nómina para modificar su nombre o su fecha.",_
							  "Images/MnSection14.gif", "Payroll.asp?Action=UpdatePayroll", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_CrearUnaNuevaNomina & ",", vbBinaryCompare) > 0),_
						Array("Prenómina",_
							  "Prepare todos los conceptos de pago para la revisión previa al cierre de la nómina. Esto le permitirá hacer cualquier corrección necesaria sobre las nóminas abiertas.",_
							  "Images/MnReportList.gif", "Payroll.asp?Action=ModifyPayroll", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_Prenomina & ",", vbBinaryCompare) > 0) And Len(asLockForPayroll(2)) = 0),_
						Array("Prenómina",_
							  "El usuario <B>" & CleanStringForHTML(sUserName) & "</B> lanzó el proceso de cálculo de la nómina <B>" & CleanStringForHTML(sPayrollName) & "</B> el <B>" & DisplayDateAndTimeFromSerialNumber(Left(asLockForPayroll(2), Len("YYYYMMDD")), Right(asLockForPayroll(2), Len("HHMMSS"))) & "</B>. No se podrá realizar otro cálculo hasta que no termine el que está en proceso.",_
							  "Images/MnReportListDis.gif", "", (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_Prenomina & ",", vbBinaryCompare) > 0) And Len(asLockForPayroll(2)) > 0),_
						Array("Cerrar nómina",_
							  "Si ya no se agregarán más conceptos de pago para los empleados, es necesario cerrar las nóminas para que se puedan pagar.",_
							  "Images/MnSection48.gif", "Payroll.asp?Action=ClosePayroll", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_CerrarNomina & ",", vbBinaryCompare) > 0),_
						Array("Nóminas especiales",_
							  "Agregue las nóminas de cancelación, extraordinaria y de pagos retroactivos",_
							  "Images/MnPayments.gif", "Payroll.asp?Action=AddSpecialPayroll", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_NominasEspeciales & ",", vbBinaryCompare) > 0),_
						Array("Cheques y depósitos",_
							  "Administración de cheques, asignación de folios, reexpedición de cheques.",_
							  "Images/MnSection47.gif", "Main_ISSSTE.asp?SectionID=47", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_Cheques & ",", vbBinaryCompare) > 0),_
						Array("Apertura y cierre de registros",_
							  "Active o desactive los módulos de registro de incidencias, padrón de madres, cuentas bancarias, etc.",_
							  "Images/MnProjects00.gif", "Main_ISSSTE.asp?SectionID=48", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_AperturaYCierreDeRegistros & ",", vbBinaryCompare) > 0),_
						Array("Reportes",_
							  "Reportes del área de informática.",_
							  "Images/MnReports.gif", "Main_ISSSTE.asp?SectionID=49", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_Reportes & ",", vbBinaryCompare) > 0),_
						Array("Catálogos",_
							  "Altas, bajas y cambios de registros concernientes a los registros del sistema.",_
							  "Images/MnHumanResources.gif", "Catalogs.asp", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_Catalogos & ",", vbBinaryCompare) > 0),_
						Array("Ventanilla única",_
							  "Control de los trámite que ingresan a la Subdirección de Personal.",_
							  "Images/MnSection61.gif", "Main_ISSSTE.asp?SectionID=61", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_VentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("Tablero de control de procesos",_
							  "Administre el estatus de los procesos registrados en el sistema.",_
							  "Images/MnSection63.gif", "TaCo.asp", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_TableroDeControlDeProcesos & ",", vbBinaryCompare) > 0),_
						Array("Ejercicio bimestral del SAR",_
							  "Ejercicio bimestral del SAR.",_
							  "Images/MnSection27.gif", "Main_ISSSTE.asp?SectionID=491", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_EjercicioBimestralDelSAR & ",", vbBinaryCompare) > 0)_
					)
				Case 42, 731 'Informática > Empleados |  Desconcentrados > Informática > Empleados
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Consulta de personal",_
							  "Consulte la información de los empleados, plaza, conceptos de pago, historia.",_
							  "Images/MnLeftArrows.gif", "Employees.asp", (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("Incidencias",_
							  "Registre las incidencias a los empleados.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Absences&SubSectionID=422", True),_
						Array("Aplicación de incidencias",_
							  "Aplique las incidencias en proceso de los empleados.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=ApplyAbsences&ReasonID=0", (Request.Cookies("SIAP_SectionID")) <> 7),_
						Array("Reclamos de pago por ajustes y deducciones",_
							  "Registre los reclamos de pago por ajustes y deducciones por empleado que hayan sido evaluados y autorizados.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=429", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_ReclamosDePagoPorAjustesYDeducciones & ",", vbBinaryCompare) > 0),_
						Array("41. Premio antigüedad 25 y 30 años",_
							  "Registre al personal con antigüedad de 25 y 30 años.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_ANTIQUITY_25_AND_30_YEARS, True),_							  
						Array("22. Premio 10 de Mayo",_
							  "Registre a los empleados de base que se hacen acreedores a esta percepción.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_MOTHERAWARD, True),_
						Array("Registro de cuentas bancarias",_
							  "Registre los numeros de cuentas bancarias para los empleados.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_BANK_ACCOUNTS,, True),_
						Array("Entrada del archivo de FOVISSSTE",_
							  "Suba la información de los empleados del FOVISSSTE con sus aportaciones del FONAC",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=FONAC", (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("Revisión de nóminas",_
							  "Indique a qué empleados se les realizará una revisión en sus nóminas.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=PayrollRevision", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_04_RevisionDeNominas & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_RevisionDeNominas & ",", vbBinaryCompare) > 0),_
						Array("<TITLE />Guardias y Suplencias",_
							  "",_
							  "", "", True),_
						Array("Guardias",_
							  "Registre, modifique y elimine los registros de guardias.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=423&FromSectionID=" & oRequest("FromSectionID").Item, True),_
						Array("Suplencias",_
							  "Registre, modifique y elimine los registros de suplencias.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=424&FromSectionID=" & oRequest("FromSectionID").Item, True),_
						Array("Rezago quirúrgico",_
							  "Registre, modifique y elimine los registros de rezago quirúrgico.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=425&FromSectionID=" & oRequest("FromSectionID").Item, True),_
						Array("Programa de vacunación",_
							  "Registre, modifique y elimine los registros del programa de vacunación.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=426&FromSectionID=" & oRequest("FromSectionID").Item, True),_
						Array("Registro de Personal externo",_
							  "Registre, modifique y elimine los registros del personal externo.",_
							  "Images/MnLeftArrows.gif", "SpecialJourney.asp?SubSectionID=427&External=1", True),_
						Array("Registro de beneficiario(a)s de pensión",_
							  "Registre, modifique y elimine los registros de beneficiario(a)s de pensión.",_
							  "Images/MnLeftArrows.gif", "SpecialJourney.asp?SubSectionID=428&Beneficiaries=1", True),_
						Array("<TITLE />Reportes",_
							  "",_
							  "", "", True),_
						Array("Reporte de personal interno",_
							  "Totales para guardias, suplencias y/o PROVAC fitlrados por quincena y agrupados por delegaciones, para personal interno.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2420" & sSubSectionID, True),_
						Array("Reporte del concepto 15 (guardias), 31 (suplencias) y C5 (guardias PROVAC) de personal interno",_
							  "Listado de empleados y sus registros agrupados por periodo y totalizados por montos y horas, para personal interno.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2421" & sSubSectionID, True),_
						Array("Reporte de captura de personal externo",_
							  "Reporte en orden de captura de movimientos de guardias, suplencias y PROVAC correspondientes a una quincena, para personal externo.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2422" & sSubSectionID, True),_
						Array("Reporte de validación de personal externo",_
							  "Reporte de validación de movimientos de guardias, suplencias y PROVAC correspondientes a una quincena, para personal externo.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2423" & sSubSectionID, True),_
						Array("Volantes",_
							  "Generación de volantes de pago para guardias, suplencias y PROVAC correspondientes a una quincena, para personal externo.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=2426" & sSubSectionID, True),_
						Array("Listado de firmas",_
							  "Listado de firmas de guardias, suplencias y PROVAC correspondientes a una quincena, para personal externo.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2427" & sSubSectionID, True),_
						Array("Reporte de totales",_
							  "Reporte de totales de guardias, suplencias y PROVAC correspondientes a una quincena.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2428" & sSubSectionID, True),_
						Array("Reporte concentrado por quincena",_
							  "Reporte concentrado por quincena del archivo histórico de totales.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2429" & sSubSectionID, True),_
						Array("Reporte estadístico de causas",_
							  "Reporte estadístico de causas de guardias, suplencias y PROVAC.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2430" & sSubSectionID, True),_
						Array("Reporte de cuentas bancarias",_
							  "Reporte histórico de cuentas bancarias de los empleados.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=2431" & sSubSectionID, True),_
						Array("Listado de actualización de cuentas bancarias",_
							  "Reporte comparativo de las cuentas bancarias de los empleados.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=2432" & sSubSectionID, True),_
						Array("Registro de revisiones salariales",_
							  "",_
							  "Images/MnLeftArrows.gif", "xxx.asp", False)_
					)
				Case 421 'Informática > Empleados > Incidencias
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Incidencias diarias",_
							  "Capture incidencias que se generan en la jornada: Retardos, inasistencias, etc.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Absences&ReasonID=0", True),_
						Array("Incidencias por periodo",_
							  "Registre las incidencias identificadas para un periodo especifíco: Licencias, vacaciones, incapacidades, etc.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=Absences&ReasonID=1", True)_
					)
                Case 423 'Informática > Empleados > Guardias
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Registro de información para internos",_
								"Agregue los registros del personal interno.",_
								"Images/MnAreas.gif", "SpecialJourney.asp?New=1&SubSectionID=423&SpecialJourneyID=423&SpecialJourneyType=1&FromSectionID=" & oRequest("FromSectionID").Item, True),_
						Array("Registro de información para externos",_
								"Agregue los registros del personal externo.",_
								"Images/MnEmployees.gif", "SpecialJourney.asp?SubSectionID=423&SpecialJourneyID=423&SpecialJourneyType=2&New=1&FromSectionID=" & oRequest("FromSectionID").Item, True),_
						Array("Búsqueda de información",_
								"A partir de un filtro, obtenga los registros tanto del personal interno como externo.",_
								"Images/MnJobs.gif", "SpecialJourney.asp?SubSectionID=423&Search=1&FromSectionID=" & oRequest("FromSectionID").Item, True),_
						Array("<TITLE />Reportes",_
							    "",_
							    "", "", True),_
						Array("Reporte de personal interno",_
							    "Totales para guardias, suplencias y/o PROVAC fitlrados por quincena y agrupados por delegaciones, para personal interno.",_
							    "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2420" & sSubSectionID, True),_
						Array("Reporte del concepto 15 (guardias), 31 (suplencias) y C5 (guardias PROVAC) de personal interno",_
							    "Listado de empleados y sus registros agrupados por periodo y totalizados por montos y horas, para personal interno.",_
							    "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2421" & sSubSectionID, True),_
						Array("Reporte de captura de personal externo",_
							    "Reporte en orden de captura de movimientos de guardias, suplencias y PROVAC correspondientes a una quincena, para personal externo.",_
							    "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2422" & sSubSectionID, True)_
                    )
				Case 424, 425, 426 'Informática > Empleados > Guardias | Suplencias | Rezago quirúrgico | Programa de vacunación
					If (Not bSearchForm) And (Not bShowForm) And (Not bAction) And (Len(oRequest("DoSearch").Item) = 0) Then
						aMenuComponent(A_ELEMENTS_MENU) = Array(_
							Array("Registro de información para internos",_
								  "Agregue los registros del personal interno.",_
								  "Images/MnAreas.gif", "SpecialJourney.asp?New=1&SubSectionID=424&SpecialJourneyType=1&FromSectionID=" & oRequest("FromSectionID").Item, True),_
							Array("Registro de información para externos",_
								  "Agregue los registros del personal externo.",_
								  "Images/MnEmployees.gif", "SpecialJourney.asp?SubSectionID=424&SpecialJourneyType=2&New=1&FromSectionID=" & oRequest("FromSectionID").Item, True),_
							Array("Búsqueda de información",_
								  "A partir de un filtro, obtenga los registros tanto del personal interno como externo.",_
								  "Images/MnJobs.gif", "SpecialJourney.asp?SubSectionID=423&Search=1&FromSectionID=" & oRequest("FromSectionID").Item, True),_
						    Array("<TITLE />Reportes",_
							      "",_
							      "", "", True),_
						    Array("Reporte de personal interno",_
							      "Totales para guardias, suplencias y/o PROVAC fitlrados por quincena y agrupados por delegaciones, para personal interno.",_
							      "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2420" & sSubSectionID, True),_
						    Array("Reporte del concepto 15 (guardias), 31 (suplencias) y C5 (guardias PROVAC) de personal interno",_
							      "Listado de empleados y sus registros agrupados por periodo y totalizados por montos y horas, para personal interno.",_
							      "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2421" & sSubSectionID, True),_
						    Array("Reporte de captura de personal externo",_
							      "Reporte en orden de captura de movimientos de guardias, suplencias y PROVAC correspondientes a una quincena, para personal externo.",_
							      "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2422" & sSubSectionID, True),_
						    Array("Reporte de validación de personal externo",_
							      "Reporte de validación de movimientos de guardias, suplencias y PROVAC correspondientes a una quincena, para personal externo.",_
							      "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2423" & sSubSectionID, True),_
						    Array("Volantes",_
							      "Generación de volantes de pago para guardias, suplencias y PROVAC correspondientes a una quincena, para personal externo.",_
							      "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=2426" & sSubSectionID, True),_
						    Array("Listado de firmas",_
							      "Listado de firmas de guardias, suplencias y PROVAC correspondientes a una quincena, para personal externo.",_
							      "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2427" & sSubSectionID, True),_
						    Array("Reporte de totales",_
							      "Reporte de totales de guardias, suplencias y PROVAC correspondientes a una quincena.",_
							      "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2428" & sSubSectionID, True),_
						    Array("Reporte concentrado por quincena",_
							      "Reporte concentrado por quincena del archivo histórico de totales.",_
							      "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2429" & sSubSectionID, True),_
						    Array("Reporte estadístico de causas",_
							      "Reporte estadístico de causas de guardias, suplencias y PROVAC.",_
							      "Images/MnLeftArrows.gif", "Reports.asp?ReportID=2430" & sSubSectionID, True),_
						    Array("Reporte de cuentas bancarias",_
							      "Reporte histórico de cuentas bancarias de los empleados.",_
							      "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=2431" & sSubSectionID, True),_
						    Array("Listado de actualización de cuentas bancarias",_
							      "Reporte comparativo de las cuentas bancarias de los empleados.",_
							      "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=2432" & sSubSectionID, True),_
						    Array("Registro de revisiones salariales",_
							      "",_
							      "Images/MnLeftArrows.gif", "xxx.asp", False)_
						)
					Else
						If (lErrorNumber = 0) And bAction Then
							If Len(oRequest("Remove").Item) > 0 Then
								sErrorDescription = "La información fue eliminada con éxito."
							Else
								sErrorDescription = "La información fue guardada con éxito."
							End If
							Call DisplayErrorMessage("Confirmación", sErrorDescription & "<BR /><FORM><INPUT TYPE=""BUTTON"" VALUE=""Registro de información"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=" & iSectionID & "&New=1&EmployeeID=" & oRequest("EmployeeID").Item & "'"" /><IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" /><INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=" & iSectionID & "'"" /></FORM>")
							Response.Write "<BR />"
							sErrorDescription = ""
						End If
						If lErrorNumber <> 0 Then
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							Response.Write "<BR />"
						End If

						If bSearchForm Then
							lErrorNumber = Display423SearchForm(oRequest, oADODBConnection, iSectionID, sErrorDescription)
						ElseIf Len(oRequest("DoSearch").Item) > 0 Then
							If iSectionID = 425 Then 'RQ
								If StrComp(oRequest("Internal").Item, "1", vbBinaryCompare) = 0 Then
									aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("2,3,4,5,6,10,17,18,19,28,31", ",")
								Else
									aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("3,4,5,6,10,17,18,19,28,31", ",")
								End If
							End If
							If Len(aCatalogComponent(S_QUERY_CONDITION_CATALOG)) = 0 Then aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (AppliedDate In (Select PayrollID From Payrolls Where IsActive_5<>0))"
							lErrorNumber = DisplayCatalogsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, aCatalogComponent, sErrorDescription)
							aCatalogComponent(S_QUERY_CONDITION_CATALOG) = ""
							If lErrorNumber = L_ERR_NO_RECORDS Then
								Call DisplayErrorMessage("Búsqueda vacía", "No existen registros que cumplan con los criterios de la búsqueda")
								Response.Write "<BR />"
								lErrorNumber = Display423SearchForm(oRequest, oADODBConnection, sErrorDescription)
							End If
						ElseIf bShowForm Then
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "var bBlock = true;" & vbNewLine

								Response.Write "function DoBlock() {" & vbNewLine
									Response.Write "if (bBlock) {" & vbNewLine
										Response.Write "document.CatalogFrm.EmployeeNumber.focus();" & vbNewLine
									Response.Write "}" & vbNewLine
								Response.Write "} // End of DoBlock" & vbNewLine

								Response.Write "function ValitadeCatalogFields(oForm) {" & vbNewLine
									Response.Write "var lEmployeeID = parseInt(oForm.EmployeeID.value);" & vbNewLine

									Response.Write "if (isNaN(lEmployeeID))" & vbNewLine
										Response.Write "lEmployeeID = -1;" & vbNewLine

									Response.Write "if (lEmployeeID < 800000) {" & vbNewLine
										Response.Write "if ((oForm.CheckEmployeeID.value == '') || (oForm.CheckEmployeeID.value == '-1')) {" & vbNewLine
											Response.Write "alert('Favor de validar el número del empleado.');" & vbNewLine
											Response.Write "oForm.EmployeeNumber.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "} else {" & vbNewLine
										Response.Write "if ((oForm.CheckEmployeeID.value == '') || (oForm.CheckEmployeeID.value == '-1')) {" & vbNewLine
											Response.Write "alert('Favor de validar el RFC del empleado externo.');" & vbNewLine
											Response.Write "oForm.RFC.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (oForm.RFC.value.length < 13) {" & vbNewLine
											Response.Write "alert('El RFC debe ser de 13 posiciones.');" & vbNewLine
											Response.Write "oForm.RFC.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (oForm.CURP.value.length < 18) {" & vbNewLine
											Response.Write "alert('El CURP debe ser de 18 posiciones.');" & vbNewLine
											Response.Write "oForm.CURP.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									Response.Write "}" & vbNewLine

									Select Case iSectionID
										Case 423 'Guardias
											If lEmployeeID < 800000 Then
												aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,5,11,11,11,11,11,11,11,6,11,11,11,11,11,11,5,1,1,11,11,6,2,6,11,6,5,2,11,11,6,11,11,11,11,11"
												aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ onChange=""document.CatalogFrm.CheckEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeNumber.value, 'EmployeesGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&AreaID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10) & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckEmployeeID"" ID=""CheckEmployeeIDHdn"" VALUE="""" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onChange=""document.CatalogFrm.CheckOriginalEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.OriginalEmployeeID.value, 'EmployeesGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&Original=1&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado a suplir"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckOriginalEmployeeID"" ID=""CheckOriginalEmployeeIDHdn"" VALUE="""" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" ÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeID.value, 'RecordsForGyS&TheRecordID=' + document.CatalogFrm.RecordID.value + '&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&OriginalEmployeeID=' + document.CatalogFrm.OriginalEmployeeID.value + '&RiskLevelID=' + document.CatalogFrm.RiskLevelID.value + '&MovementID=' + document.CatalogFrm.MovementID.value + '&StartDate=' + BuildDateString(document.CatalogFrm.StartDateYear.value, document.CatalogFrm.StartDateMonth.value, document.CatalogFrm.StartDateDay.value) + '&EndDate=' + BuildDateString(document.CatalogFrm.EndDateYear.value, document.CatalogFrm.EndDateMonth.value, document.CatalogFrm.EndDateDay.value) + '&WorkingHours=' + document.CatalogFrm.WorkingHours.value + '&WorkedHours=' + document.CatalogFrm.WorkedHours.value + '&PayrollDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeJourneysIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar las horas reportadas por quincena para el empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""ReportedHours"" ID=""ReportedHoursHdn"" VALUE="""" ÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" ÞÞÞÞÞÞÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
											Else
												aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,11,5,5,5,5,5,11,6,6,6,6,6,11,11,11,5,1,1,11,11,6,2,6,11,6,5,2,11,11,6,11,11,11,11,11"
												aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ onChange=""document.CatalogFrm.EmployeeID.value=this.value; document.CatalogFrm.ReportedHours.value='';"" /><INPUT TYPE=""HIDDEN"" NAME=""CheckEmployeeID"" ID=""CheckEmployeeIDHdn"" VALUE="""" ÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""document.CatalogFrm.CheckEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.RFC.value, 'ExternalGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&CURP=' + document.CatalogFrm.CURP.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar RFC del empleado externo"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckEmployeeID"" ID=""CheckEmployeeIDHdn"" VALUE="""" ÞÞÞÞÞÞ onChange=""document.CatalogFrm.CheckOriginalEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.OriginalEmployeeID.value, 'EmployeesGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&Original=1&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado a suplir"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckOriginalEmployeeID"" ID=""CheckOriginalEmployeeIDHdn"" VALUE="""" ÞÞÞ onChange=""SearchRecord('P', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=-1&ServiceID=-1&LevelID=-1&WorkingHours=-1&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞ onChange=""SearchRecord('A', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&ServiceID=-1&LevelID=-1&WorkingHours=-1&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞ onChange=""SearchRecord('S', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&ServiceID=' + document.CatalogFrm.ServiceID.value + '&LevelID=-1&WorkingHours=-1&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞ onChange=""SearchRecord('L', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AeraID=' + document.CatalogFrm.AeraID.value + '&ServiceID=' + document.CatalogFrm.ServiceID.value + '&LevelID=' + document.CatalogFrm.LevelID.value + '&WorkingHours=-1&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞ onChange=""SearchRecord('W', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&ServiceID=' + document.CatalogFrm.ServiceID.value + '&LevelID=' + document.CatalogFrm.LevelID.value + '&WorkingHours=' + document.CatalogFrm.WorkingHours.value + '&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" ÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeID.value, 'RecordsForGyS&TheRecordID=' + document.CatalogFrm.RecordID.value + '&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&OriginalEmployeeID=' + document.CatalogFrm.OriginalEmployeeID.value + '&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&RiskLevelID=' + document.CatalogFrm.RiskLevelID.value + '&MovementID=' + document.CatalogFrm.MovementID.value + '&StartDate=' + BuildDateString(document.CatalogFrm.StartDateYear.value, document.CatalogFrm.StartDateMonth.value, document.CatalogFrm.StartDateDay.value) + '&EndDate=' + BuildDateString(document.CatalogFrm.EndDateYear.value, document.CatalogFrm.EndDateMonth.value, document.CatalogFrm.EndDateDay.value) + '&WorkingHours=' + document.CatalogFrm.WorkingHours.value + '&WorkedHours=' + document.CatalogFrm.WorkedHours.value + '&PayrollDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeJourneysIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar las horas reportadas por quincena para el empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""ReportedHours"" ID=""ReportedHoursHdn"" VALUE="""" ÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" ÞÞÞÞÞÞÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
												aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) = 800000
												aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2) = 800000
											End If
											aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
											aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
											Response.Write "oForm.OriginalEmployeeID.value = '-1';" & vbNewLine
										Case 424 'Suplencias
											If lEmployeeID < 800000 Then
												aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,5,5,5,5,5,5,5,11,11,11,11,11,11,11,11,5,1,1,11,11,6,2,6,11,6,5,2,11,11,6,11,11,11,11,11"
												aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ onChange=""document.CatalogFrm.CheckEmployeeID.value=''; document.CatalogFrm.CheckOriginalEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeNumber.value, 'EmployeesGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&AreaID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10) & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckEmployeeID"" ID=""CheckEmployeeIDHdn"" VALUE="""" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onChange=""document.CatalogFrm.CheckOriginalEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: if (parseInt(document.CatalogFrm.EmployeeID.value) == parseInt(document.CatalogFrm.OriginalEmployeeID.value)) {alert('El empleado suplido no puede ser el mismo que el empleado suplente.'); document.CatalogFrm.OriginalEmployeeID.focus();} else {SearchRecord(document.CatalogFrm.OriginalEmployeeID.value, 'EmployeesGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&Original=1&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm');}""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado a suplir"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckOriginalEmployeeID"" ID=""CheckOriginalEmployeeIDHdn"" VALUE="""" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" ÞÞÞ onFocus=""DoBlock();"" onChange=""document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeID.value, 'RecordsForGyS&TheRecordID=' + document.CatalogFrm.RecordID.value + '&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&OriginalEmployeeID=' + document.CatalogFrm.OriginalEmployeeID.value + '&RiskLevelID=' + document.CatalogFrm.RiskLevelID.value + '&MovementID=' + document.CatalogFrm.MovementID.value + '&StartDate=' + BuildDateString(document.CatalogFrm.StartDateYear.value, document.CatalogFrm.StartDateMonth.value, document.CatalogFrm.StartDateDay.value) + '&EndDate=' + BuildDateString(document.CatalogFrm.EndDateYear.value, document.CatalogFrm.EndDateMonth.value, document.CatalogFrm.EndDateDay.value) + '&WorkingHours=' + document.CatalogFrm.WorkingHours.value + '&WorkedHours=' + document.CatalogFrm.WorkedHours.value + '&PayrollDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeJourneysIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar las horas reportadas por quincena para el empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""ReportedHours"" ID=""ReportedHoursHdn"" VALUE="""" ÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" ÞÞÞÞÞÞÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
											Else
												aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,11,5,5,5,5,5,5,11,6,11,11,11,11,11,11,5,1,1,11,11,6,2,6,11,6,5,2,11,11,6,11,11,11,11,11"
												aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(10) = Replace(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(10), "Seleccione un puesto", "Valide el número de empleado a suplir")
												aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(11) = Replace(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(11), "Seleccione un puesto", "Valide el número de empleado a suplir")
												aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(12) = Replace(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(12), "Seleccione un puesto", "Valide el número de empleado a suplir")
												aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(13) = Replace(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(13), "Seleccione un puesto", "Valide el número de empleado a suplir")
												aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ onChange=""document.CatalogFrm.EmployeeID.value=this.value; document.CatalogFrm.ReportedHours.value='';"" /><INPUT TYPE=""HIDDEN"" NAME=""CheckEmployeeID"" ID=""CheckEmployeeIDHdn"" VALUE="""" ÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""document.CatalogFrm.CheckEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.RFC.value, 'ExternalGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&CURP=' + document.CatalogFrm.CURP.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar RFC del empleado externo"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckEmployeeID"" ID=""CheckEmployeeIDHdn"" VALUE="""" ÞÞÞÞÞÞ onChange=""document.CatalogFrm.CheckOriginalEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: if (parseInt(document.CatalogFrm.EmployeeID.value) == parseInt(document.CatalogFrm.OriginalEmployeeID.value)) {alert('El empleado suplido no puede ser el mismo que el empleado suplente.'); document.CatalogFrm.OriginalEmployeeID.focus();} else {SearchRecord(document.CatalogFrm.OriginalEmployeeID.value, 'EmployeesGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&Original=1&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm');}""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado a suplir"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckOriginalEmployeeID"" ID=""CheckOriginalEmployeeIDHdn"" VALUE="""" ÞÞÞÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" ÞÞÞ onFocus=""DoBlock();"" onChange=""document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeID.value, 'RecordsForGyS&TheRecordID=' + document.CatalogFrm.RecordID.value + '&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&OriginalEmployeeID=' + document.CatalogFrm.OriginalEmployeeID.value + '&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&RiskLevelID=' + document.CatalogFrm.RiskLevelID.value + '&MovementID=' + document.CatalogFrm.MovementID.value + '&StartDate=' + BuildDateString(document.CatalogFrm.StartDateYear.value, document.CatalogFrm.StartDateMonth.value, document.CatalogFrm.StartDateDay.value) + '&EndDate=' + BuildDateString(document.CatalogFrm.EndDateYear.value, document.CatalogFrm.EndDateMonth.value, document.CatalogFrm.EndDateDay.value) + '&WorkingHours=' + document.CatalogFrm.WorkingHours.value + '&WorkedHours=' + document.CatalogFrm.WorkedHours.value + '&PayrollDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeJourneysIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar las horas reportadas por quincena para el empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""ReportedHours"" ID=""ReportedHoursHdn"" VALUE="""" ÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
												aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) = 800000
												aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2) = 800000
											End If
											aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(8) = ""
											aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
											aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
											aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("2,3,4,5,6,8,10,17,18,19,22,23,24,28,31", ",")

											Response.Write "if (oForm.OriginalEmployeeID.value == '') {" & vbNewLine
												Response.Write "alert('Favor de especificar y validar el número del empleado a suplir.');" & vbNewLine
												Response.Write "oForm.OriginalEmployeeID.focus();" & vbNewLine
												Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine

											Response.Write "if ((oForm.CheckOriginalEmployeeID.value == '') || (oForm.CheckOriginalEmployeeID.value == '-1')) {" & vbNewLine
												Response.Write "alert('Favor de validar el número del empleado a suplir.');" & vbNewLine
												Response.Write "oForm.OriginalEmployeeID.focus();" & vbNewLine
												Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
										Case 425 'RQ
											If lEmployeeID < 800000 Then
												aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,5,11,11,11,11,11,11,11,6,11,11,11,11,11,11,11,1,1,11,11,11,11,11,11,11,5,2,11,11,6,11,11,11,11,11"
												aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ onChange=""document.CatalogFrm.CheckEmployeeID.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeNumber.value, 'EmployeesGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&AreaID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10) & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckEmployeeID"" ID=""CheckEmployeeIDHdn"" VALUE="""" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞ onChange=""document.CatalogFrm.CheckOriginalEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.OriginalEmployeeID.value, 'EmployeesGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&Original=1&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado a suplir"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckOriginalEmployeeID"" ID=""CheckOriginalEmployeeIDHdn"" VALUE="""" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeID.value, 'RecordsForGyS&TheRecordID=' + document.CatalogFrm.RecordID.value + '&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&OriginalEmployeeID=' + document.CatalogFrm.OriginalEmployeeID.value + '&RiskLevelID=' + document.CatalogFrm.RiskLevelID.value + '&MovementID=' + document.CatalogFrm.MovementID.value + '&StartDate=' + BuildDateString(document.CatalogFrm.StartDateYear.value, document.CatalogFrm.StartDateMonth.value, document.CatalogFrm.StartDateDay.value) + '&EndDate=' + BuildDateString(document.CatalogFrm.EndDateYear.value, document.CatalogFrm.EndDateMonth.value, document.CatalogFrm.EndDateDay.value) + '&WorkingHours=' + document.CatalogFrm.WorkingHours.value + '&WorkedHours=' + document.CatalogFrm.WorkedHours.value + '&PayrollDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeJourneysIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar las horas reportadas por quincena para el empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""ReportedHours"" ID=""ReportedHoursHdn"" VALUE="""" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
											Else
												aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,11,5,5,5,5,5,11,11,6,11,11,11,11,11,11,11,1,1,11,11,11,11,11,11,11,5,2,11,11,6,11,11,11,11,11"
												aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ onChange=""document.CatalogFrm.EmployeeID.value=this.value;"" /><INPUT TYPE=""HIDDEN"" NAME=""CheckEmployeeID"" ID=""CheckEmployeeIDHdn"" VALUE="""" ÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""document.CatalogFrm.CheckEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.RFC.value, 'ExternalGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&CURP=' + document.CatalogFrm.CURP.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar RFC del empleado externo"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckEmployeeID"" ID=""CheckEmployeeIDHdn"" VALUE="""" ÞÞÞÞÞÞ onChange=""document.CatalogFrm.CheckOriginalEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.OriginalEmployeeID.value, 'EmployeesGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&Original=1&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado a suplir"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckOriginalEmployeeID"" ID=""CheckOriginalEmployeeIDHdn"" VALUE="""" ÞÞÞ onChange=""SearchRecord('P', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=-1&ServiceID=-1&LevelID=-1&WorkingHours=-1&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞ onChange=""SearchRecord('A', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&ServiceID=-1&LevelID=-1&WorkingHours=-1&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞ onChange=""SearchRecord('S', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&ServiceID=' + document.CatalogFrm.ServiceID.value + '&LevelID=-1&WorkingHours=-1&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞ onChange=""SearchRecord('L', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AeraID=' + document.CatalogFrm.AeraID.value + '&ServiceID=' + document.CatalogFrm.ServiceID.value + '&LevelID=' + document.CatalogFrm.LevelID.value + '&WorkingHours=-1&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞ onChange=""SearchRecord('W', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&ServiceID=' + document.CatalogFrm.ServiceID.value + '&LevelID=' + document.CatalogFrm.LevelID.value + '&WorkingHours=' + document.CatalogFrm.WorkingHours.value + '&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeID.value, 'RecordsForGyS&TheRecordID=' + document.CatalogFrm.RecordID.value + '&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&OriginalEmployeeID=' + document.CatalogFrm.OriginalEmployeeID.value + '&RiskLevelID=' + document.CatalogFrm.RiskLevelID.value + '&MovementID=' + document.CatalogFrm.MovementID.value + '&StartDate=' + BuildDateString(document.CatalogFrm.StartDateYear.value, document.CatalogFrm.StartDateMonth.value, document.CatalogFrm.StartDateDay.value) + '&EndDate=' + BuildDateString(document.CatalogFrm.EndDateYear.value, document.CatalogFrm.EndDateMonth.value, document.CatalogFrm.EndDateDay.value) + '&WorkingHours=' + document.CatalogFrm.WorkingHours.value + '&WorkedHours=' + document.CatalogFrm.WorkedHours.value + '&PayrollDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeJourneysIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar las horas reportadas por quincena para el empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""ReportedHours"" ID=""ReportedHoursHdn"" VALUE=""""ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
												aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(10) = "Areas;,;AreaID;,;AreaCode, AreaName;,;(CenterTypeID In (Select Distinct CenterTypeID From PositionsSpecialJourneysLKP)) And (ParentID>-1);,;AreaCode;,;;,;Ninguno;;;-1"
												aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) = 800000
												aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2) = 800000
											End If
											aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
											aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
											aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG)(17) = 0
											Response.Write "oForm.DocumentNumber.value = ' ';" & vbNewLine
											Response.Write "oForm.OriginalEmployeeID.value = '-1';" & vbNewLine
										Case 426 'PROVAC
											Response.Write "oForm.OriginalEmployeeID.value = '-1';" & vbNewLine
											If lEmployeeID < 800000 Then
												aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,5,11,11,11,11,11,11,11,6,11,11,11,11,11,11,5,1,1,11,11,6,2,6,11,6,5,2,11,11,6,11,11,11,11,11"
												aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ onChange=""document.CatalogFrm.CheckEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeNumber.value, 'EmployeesGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&AreaID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10) & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckEmployeeID"" ID=""CheckEmployeeIDHdn"" VALUE="""" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onChange=""document.CatalogFrm.CheckOriginalEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.OriginalEmployeeID.value, 'EmployeesGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&Original=1&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado a suplir"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckOriginalEmployeeID"" ID=""CheckOriginalEmployeeIDHdn"" VALUE="""" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" ÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeID.value, 'RecordsForGyS&TheRecordID=' + document.CatalogFrm.RecordID.value + '&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&OriginalEmployeeID=' + document.CatalogFrm.OriginalEmployeeID.value + '&RiskLevelID=' + document.CatalogFrm.RiskLevelID.value + '&MovementID=' + document.CatalogFrm.MovementID.value + '&StartDate=' + BuildDateString(document.CatalogFrm.StartDateYear.value, document.CatalogFrm.StartDateMonth.value, document.CatalogFrm.StartDateDay.value) + '&EndDate=' + BuildDateString(document.CatalogFrm.EndDateYear.value, document.CatalogFrm.EndDateMonth.value, document.CatalogFrm.EndDateDay.value) + '&WorkingHours=' + document.CatalogFrm.WorkingHours.value + '&WorkedHours=' + document.CatalogFrm.WorkedHours.value + '&PayrollDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeJourneysIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar las horas reportadas por quincena para el empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""ReportedHours"" ID=""ReportedHoursHdn"" VALUE="""" ÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" ÞÞÞÞÞÞÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
											Else
												aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,11,5,5,5,5,5,11,6,6,6,6,6,11,11,11,5,1,1,11,11,6,2,6,11,6,5,2,11,11,6,11,11,11,11,11"
												aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) = 800000
												aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2) = 800000
												aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ onChange=""document.CatalogFrm.EmployeeID.value=this.value; document.CatalogFrm.ReportedHours.value='';"" /><INPUT TYPE=""HIDDEN"" NAME=""CheckEmployeeID"" ID=""CheckEmployeeIDHdn"" VALUE="""" ÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""document.CatalogFrm.CheckEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.RFC.value, 'ExternalGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&CURP=' + document.CatalogFrm.CURP.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar RFC del empleado externo"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckEmployeeID"" ID=""CheckEmployeeIDHdn"" VALUE="""" ÞÞÞÞÞÞ onChange=""document.CatalogFrm.CheckOriginalEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.OriginalEmployeeID.value, 'EmployeesGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&Original=1&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado a suplir"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckOriginalEmployeeID"" ID=""CheckOriginalEmployeeIDHdn"" VALUE="""" ÞÞÞ onChange=""SearchRecord('P', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=-1&ServiceID=-1&LevelID=-1&WorkingHours=-1&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞ onChange=""SearchRecord('A', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&ServiceID=-1&LevelID=-1&WorkingHours=-1&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞ onChange=""SearchRecord('S', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&ServiceID=' + document.CatalogFrm.ServiceID.value + '&LevelID=-1&WorkingHours=-1&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞ onChange=""SearchRecord('L', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AeraID=' + document.CatalogFrm.AeraID.value + '&ServiceID=' + document.CatalogFrm.ServiceID.value + '&LevelID=' + document.CatalogFrm.LevelID.value + '&WorkingHours=-1&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞ onChange=""SearchRecord('W', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&ServiceID=' + document.CatalogFrm.ServiceID.value + '&LevelID=' + document.CatalogFrm.LevelID.value + '&WorkingHours=' + document.CatalogFrm.WorkingHours.value + '&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" ÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeID.value, 'RecordsForGyS&TheRecordID=' + document.CatalogFrm.RecordID.value + '&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&OriginalEmployeeID=' + document.CatalogFrm.OriginalEmployeeID.value + '&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&RiskLevelID=' + document.CatalogFrm.RiskLevelID.value + '&MovementID=' + document.CatalogFrm.MovementID.value + '&StartDate=' + BuildDateString(document.CatalogFrm.StartDateYear.value, document.CatalogFrm.StartDateMonth.value, document.CatalogFrm.StartDateDay.value) + '&EndDate=' + BuildDateString(document.CatalogFrm.EndDateYear.value, document.CatalogFrm.EndDateMonth.value, document.CatalogFrm.EndDateDay.value) + '&WorkingHours=' + document.CatalogFrm.WorkingHours.value + '&WorkedHours=' + document.CatalogFrm.WorkedHours.value + '&PayrollDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeJourneysIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar las horas reportadas por quincena para el empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""ReportedHours"" ID=""ReportedHoursHdn"" VALUE="""" ÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" ÞÞÞÞÞÞÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
											End If
											aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
											aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
									End Select

									Response.Write "if ((oForm.TempStartDate.value != BuildDateString(document.CatalogFrm.StartDateYear.value, document.CatalogFrm.StartDateMonth.value, document.CatalogFrm.StartDateDay.value)) || (oForm.TempEndDate.value != BuildDateString(document.CatalogFrm.EndDateYear.value, document.CatalogFrm.EndDateMonth.value, document.CatalogFrm.EndDateDay.value))) {" & vbNewLine
										Response.Write "alert('Favor de validar que el empleado no tenga registros en las fechas indicadas.');" & vbNewLine
										Response.Write "oForm.StartDateDay.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									If iSectionID <> 425 Then
										Response.Write "if (oForm.ReportedHours.value == '') {" & vbNewLine
											Response.Write "alert('Favor de validar los días/horas registradas para el empleado.');" & vbNewLine
											Response.Write "oForm.WorkedHours.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (oForm.ReportedHours.value == '-2') {" & vbNewLine
											Response.Write "alert('Ya existen otros registros en las fechas indicadas.');" & vbNewLine
											Response.Write "oForm.WorkedHours.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "} else {" & vbNewLine
											Response.Write "if (oForm.ReportedHours.value != '1') {" & vbNewLine
												Response.Write "alert('Las horas registradas para el empleado en la quincena exceden el número de horas establecidas como máximo.');" & vbNewLine
												Response.Write "oForm.WorkedHours.focus();" & vbNewLine
												Response.Write "return false;" & vbNewLine
											Response.Write "}" & vbNewLine
										Response.Write "}" & vbNewLine
									Else
										Response.Write "if (oForm.ReportedHours.value == '') {" & vbNewLine
											Response.Write "alert('Favor de validar que el empleado no tenga registros en las fechas indicadas.');" & vbNewLine
											Response.Write "oForm.ConceptAmount.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (oForm.ReportedHours.value == '-2') {" & vbNewLine
											Response.Write "alert('Ya existen otros registros en las fechas indicadas.');" & vbNewLine
											Response.Write "oForm.WorkedHours.focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									End If

									Response.Write "return true;" & vbNewLine
								Response.Write "} // End of ValitadeCatalogFields" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine

							Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
								Response.Write "<TD VALIGN=""TOP"">"
									lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
								Response.Write "</TD>"
								Response.Write "<TD>&nbsp;&nbsp;&nbsp;</TD>"
								Response.Write "<TD VALIGN=""TOP"">"
									Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""250""></IFRAME><BR />"
									Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeJourneysIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""150""></IFRAME>"
								Response.Write "</TD>"
							Response.Write "</TR></TABLE>"
								
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1)) = -1 Then
									Response.Write "document.CatalogFrm.EmployeeID.value = '';" & vbNewLine
								Else
									Response.Write "document.CatalogFrm.CheckEmployeeID.value = '" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & "';" & vbNewLine
									Response.Write "document.CatalogFrm.OriginalEmployeeID.value = '" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(8) & "';" & vbNewLine
									Response.Write "document.CatalogFrm.ReportedHours.value = '1';" & vbNewLine
									Response.Write "SearchRecord(document.CatalogFrm.EmployeeNumber.value, 'EmployeesGyS&RecordType=" & iSectionID - 422 & "&AreaID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10) & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm');" & vbNewLine
								End If
								If lEmployeeID < 800000 Then
								Else
									If iSectionID <> 424 Then Response.Write "SearchRecord('1', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&ServiceID=' + document.CatalogFrm.ServiceID.value + '&LevelID=' + document.CatalogFrm.LevelID.value + '&WorkingHours=' + document.CatalogFrm.WorkingHours.value + '&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm');" & vbNewLine
									Response.Write "document.CatalogFrm.RFC.focus();" & vbNewLine
								End If
							Response.Write "//--></SCRIPT>" & vbNewLine
						End If
					End If
				Case 427 'Informática > Empleados > Entrada del archivo de FOVISSSTE
					
				Case 47, 732 'Informática > Cheques | Desconcentrados > Informática > Cheques
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Asignación de folios",_
							  "",_
							  "Images/MnLeftArrows.gif", "Payments.asp?Action=PaymentsRecords", True),_
						Array("Eliminar folios no impresos",_
							  "",_
							  "Images/MnLeftArrows.gif", "Payments.asp?Action=RemovePaymentsRecords", True),_
						Array("Impresión de pagos",_
							  "",_
							  "Images/MnLeftArrows.gif", "Payments.asp?Action=PrintPayments", True),_
						Array("Reposiciones",_
							  "",_
							  "Images/MnLeftArrows.gif", "Payments.asp?Action=Replacement", True),_
						Array("Cancelación de pagos",_
							  "",_
							  "Images/MnLeftArrows.gif", "Payments.asp?Action=CancelPayments", True),_
						Array("Bloqueo de depósitos",_
							  "",_
							  "Images/MnLeftArrows.gif", "Payments.asp?Action=BlockPayments", True),_
						Array("Reexpedición",_
							  "",_
							  "Images/MnLeftArrows.gif", "Payments.asp?Action=Reexpedition", True),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Reporte de cifras de cancelaciones",_
							  "<BR />",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=4703", True),_
						Array("Listado de firmas de cancelaciones",_
							  "<BR />",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=4701", True),_
						Array("Concentrado de conceptos de cancelaciones",_
							  "<BR />",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=4702", True),_
						Array("Recibo por pago de honorarios",_
							  "",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1476", True),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Recibo de distribución y recepción de cheques",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1471", True),_
						Array("Archivo de depósitos bancarios",_
							  "",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1474", True),_
						Array("Archivo de liberación de cheques",_
							  "",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1475", True),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Pagos cancelados",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1472", True),_
						Array("Bloqueos aplicados",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1473", True),_
						Array("",_
							  "",_
							  "", "", False)_
					)
				Case 48 'Informática > Apertura y cierre de registros
					If Len(oRequest("Action").Item) = 0 Then
						aMenuComponent(A_ELEMENTS_MENU) = Array(_
							Array("Apertura y cierre para movimientos de personal",_
								  "",_
								  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=48&Action=1", True),_
							Array("Apertura y cierre para incidencias",_
								  "",_
								  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=48&Action=2", True),_
							Array("Apertura y cierre para el padrón de madres",_
								  "",_
								  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=48&Action=3", True),_
							Array("Apertura y cierre para el registro de cuentas bancarias",_
								  "",_
								  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=48&Action=4", True),_
							Array("Apertura y cierre para guardias, suplencias, rezago quirúrgico y programa de vacunación",_
								  "",_
								  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=48&Action=5", True),_
							Array("Apertura y cierre para FONAC",_
								  "",_
								  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=48&Action=6", True),_
							Array("Apertura y cierre para prestaciones",_
								  "",_
								  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=48&Action=7", True),_
							Array("",_
								  "",_
								  "", "", False)_
						)
					Else
						Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
								lErrorNumber = DisplayPayrollsStatusTable(oRequest, oADODBConnection, CInt(oRequest("Action").Item), aPayrollComponent, sErrorDescription)
							Response.Write "</TD>"
							Response.Write "<TD>&nbsp;</TD>"
							Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
							Response.Write "<TD>&nbsp;</TD>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
								Call DisplayErrorMessage("Instrucciones", "Seleccione la quincena sobre la cual desea habilitar o deshabilitar el registro <B>" & Replace(aHeaderComponent(S_TITLE_NAME_HEADER), "Apertura y cierre ", "") & "</B>.")
								Response.Write "<BR />"
								Response.Write "<FORM ACTION=""Main_ISSSTE.asp"" METHOD=""GET"">"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""48"" />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""Action"" VALUE=""" & oRequest("Action").Item & """ />"

									Response.Write "Quincena:&nbsp;"
									Response.Write "<SELECT NAME=""PayrollID"" ID=""PayrollIDCmb"" SIZE=""1"" CLASS=""Lists"">"
										If CInt(oRequest("Action").Item) = 5 Then
											Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsActive_" & oRequest("Action").Item & "<>0) And (IsClosed<>1) And (PayrollTypeID=5)", "PayrollID Desc", oRequest("PayrollID").Item, "Ninguna;;;-1", sErrorDescription)
										Else
											Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsActive_" & oRequest("Action").Item & "<>0) And (IsClosed<>1) And (PayrollTypeID<>0)", "PayrollID Desc", oRequest("PayrollID").Item, "Ninguna;;;-1", sErrorDescription)
										End If
									Response.Write "</SELECT><BR /><BR />"

									Response.Write "Acción:&nbsp;"
									Response.Write "<SELECT NAME=""StatusID"" ID=""StatusIDCmb"" SIZE=""1"" CLASS=""Lists"">"
										Response.Write "<OPTION VALUE=""2"">Deshabilitar registro</OPTION>"
										Response.Write "<OPTION VALUE=""1"">Habilitar registro</OPTION>"
									Response.Write "</SELECT>"

									Response.Write "<BR /><BR />"
									Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ModifyStatus"" ID=""ModifyStatusBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
									Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
									Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=48';"" />"
								Response.Write "</FORM>"
							Response.Write "</TD>"
						Response.Write "</TR></TABLE>"
					End If
				Case 49, 733 'Informática > Reportes | Desconcentrados > Informática > Reportes
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Reportes guardados",_
							  "Reportes generados por otros usuarios a partir de plantillas y filtros.",_
							  "Images/MnLeftArrows.gif", "SavedReport.asp?ReportType=1", (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("Tabuladores de pago",_
							  "Obtenga el listado de los pagos cancelados por empleado",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1435" & sSubSectionID, True),_
						Array("<TITLE />REPORTES SOBRE NÓMINAS",_
							  "",_
							  "", "", True),_
						Array("Reporte de cifras",_
							  "Montos pagados para la quincena especificada por área.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1490" & sSubSectionID, True),_
						Array("Remesa para cubrir la nómina",_
							  "Resumen por delegación del monto total que se autorizó por cuenta.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1007" & sSubSectionID, True),_
						Array("Concentrado de conceptos de la nómina ordinaria",_
							  "Concentrado del total de percepciones y deducciones clasificados por conceptos de nómina ordinaria.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1006" & sSubSectionID, True),_
						Array("Listado de firmas",_
							  "Para cada empleado, obtenga el listado de los montos de sus percepciones y deducciones.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1003" & sSubSectionID, True),_
						Array("Revisión de nóminas",_
							  "Para cada empleado, obtenga el listado de los montos de sus percepciones y deducciones por cada nómina revisada.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1338" & sSubSectionID, True),_
						Array("Ejercicio Bimestral del SAR",_
							  "Ejercicio Bimestral del SAR.",_
							  "Images/MnLeftArrowsZIP.gif", "Main_ISSSTE.asp?SectionID=491" & sSubSectionID, True),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Archivo para carga del SPEP",_
							  "Genere el archivo en texto para el SPEP de los conceptos pagados en la nómina.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1401" & sSubSectionID, True),_
						Array("Archivo para carga del SPEP por centro de trabajo",_
							  "Genere el archivo en texto para el SPEP de los conceptos pagados en la nómina, totalizados por centro de trabajo.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1402" & sSubSectionID, True),_
						Array("Archivo SICAD",_
							  "Generación del archivo SICAD de emisión en formato corto o largo.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1339" & sSubSectionID, True),_
						Array("Archivo SICAD de cancelaciones",_
							  "Generación del archivo SICAD de los empleados con cancelaciones.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1340" & sSubSectionID, True),_
						Array("Registro de movimientos",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=0000" & sSubSectionID, False),_
						Array("Prenóminas de movimientos",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=0000" & sSubSectionID, False),_
						Array("Auditorías de nóminas",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=0000" & sSubSectionID, False),_
						Array("Empleados por tipo de empleado y por empresa",_
							  "Listado de empleados agrupados por empresa y por su tipo (médicos, funcionarios, operativos, etc).",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1008" & sSubSectionID, True),_
						Array("Hoja de cifras",_
							  "Montos pagados para la quincena especificada por concepto de pagos.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1001" & sSubSectionID, True),_
						Array("Hoja informativa de servicios de informática",_
							  "Hojas informativas de nómina por quincena y centro de trabajo.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1005" & sSubSectionID, True),_
						Array("Memoria de cálculo para el entero de cuotas sindicales",_
							  "Para las cuotas sindicales del SNTISSSTE y FSTSE y del SITISSSTE.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1494" & sSubSectionID, (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("Impuesto sobre la renta",_
							  "Listado de ISR por CLC",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1477" & sSubSectionID, (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("Cálculo del impuesto sobre nómina",_
							  "Listado de ISN por CLC",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1478" & sSubSectionID, (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("Pensión de Ramas médica, paramédica, de grupos afines y operativa, de enlace y de alto nivel de responsabilidad",_
							  "Resumen por banco de los totales pagados a través de cheque, de depósito en cuenta de débito, foráneo y local.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1011" & sSubSectionID, True),_
						Array("Ramas médica, paramédica, de grupos afines y operativa, de enlace y de alto nivel de responsabilidad",_
							  "Resumen por banco de los totales pagados a través de cheque, de depósito en cuenta de débito, foráneo y local.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1010" & sSubSectionID, True),_
						Array("Resumen mensual de nóminas",_
							  "Listado de CLC para el pago de nóminas.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1009" & sSubSectionID, True),_
						Array("Resumen por conceptos de nómina",_
							  "Total de pagos realizados por cada concepto de pago, en las dos nóminas del mes seleccionado.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1004" & sSubSectionID, True),_
						Array("Incidencias de los empleados",_
							  "Listado de las incidencias registradas a los empleados.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1026" & sSubSectionID, True),_
						Array("Incidencias con horas extras y primas dominicales",_
							  "Listado de incidencias registradas y que en el mismo día tengan registros de horas extras o primas dominicales.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1030" & sSubSectionID, True),_
						Array("Concentrado de incidencias",_
							  "Total de incidencias registradas a los empleados.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1029" & sSubSectionID, True),_
						Array("Prestaciones de los empleados",_
							  "Listado de las prestaciones registradas a los empleados.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1108" & sSubSectionID, True),_
						Array("Reporte de estímulos",_
							  "Listado de los estímulos pagados a los empleados.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1433" & sSubSectionID, True),_
						Array("Reporte de incidencias",_
							  "Listado de las incidencias descontadas a los empleados.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1434" & sSubSectionID, True),_
						Array("Listado de cheques",_
							  "Listado para impresión en formatos de cheques y depósitos para empleados.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1027" & sSubSectionID, False),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Altas de empleados por unidad administrativa",_
							  "Auditoría de nómina. Listado de los empleados de nuevo ingreso por unidad administrativa.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1015" & sSubSectionID, True),_
						Array("Bajas de empleados por unidad administrataiva",_
							  "Auditoría de nómina. Listado de los empleados dados de baja por unidad administrativa.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1017" & sSubSectionID, True),_
						Array("Bajas por unidad administrativa",_
							  "Auditoría de nómina. Número de empleados dados de baja por delegación.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1016" & sSubSectionID, True),_
						Array("Cambios de puesto por unidad administrativa",_
							  "Auditoría de nómina. Listado de los empleados que cambiaron de puesto comparado con la nómina anterior.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1019" & sSubSectionID, True),_
						Array("Diferencias de empleados por unidad administrativa",_
							  "Auditoría de nómina. Comparativo del total de registros por delegación, en las dos nóminas del mes seleccionado.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1013" & sSubSectionID, True),_
						Array("Diferencias de sueldo por unidad administrativa",_
							  "Auditoría de nómina. Listado de los empleados que tienen diferencia de salario base comparado con la nómina anterior.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1018" & sSubSectionID, True),_
						Array("Diferencias totales por concepto",_
							  "Auditoría de nómina. Comparativo de nóminas por cada concepto de pago, en las dos nóminas del mes seleccionado.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1012" & sSubSectionID, True),_
						Array("Funcionarios con líquido mayor a la suma de sueldo base más compensación",_
							  "Auditoría de nómina. Comparativo del sueldo base mas compensación con el importe líquido por delegación.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1020" & sSubSectionID, True),_
						Array("Resumen de altas por unidad administrativa",_
							  "Auditoría de nómina. Resumen de altas por unidad administrativa.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1014" & sSubSectionID, True),_
						Array("Revisión de diferencias",_
							  "Para cada emplados, compare los montos de sus percepciones y deducciones entre una nómina y la anterior para detectar diferencias sustanciales.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1002" & sSubSectionID, False),_
						Array("Totales por nómina",_
							  "Auditoría de nómina. Comparativo del total de registros, devengos, retenido y líquido, en las dos nóminas del mes seleccionado.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1021" & sSubSectionID, True),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Constancia de percepciones y deducciones",_
							  "Obtenga la constancia de percepciones y deducciones de un empleado para una quincena.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1499" & sSubSectionID, True),_
						Array("Empleados con líquidos mayores",_
							  "Auditoría de nómina. Listado de empleados con líquidos mayores que cumplan las condiciones proporcionadas por el usuario.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1022" & sSubSectionID, True),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Conteo de empleados",_
							  "Obtenga un conteo de los empleados que se encuentran registrados en el sistema, agrupando los resultados por diferentes conceptos.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=702" & sSubSectionID, True),_
						Array("Funcionarios y operativos por concepto de pago y empresa",_
							  "Listado de empleados agrupados por concepto de pago y por empresa, clasificados en funcionarios y operativos.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1024" & sSubSectionID, True),_
						Array("Reporte de movimientos",_
							  "Listado de empleados que tuvieron movimentos de asignación de titularidad o de licencias en la nómina seleccionada.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1023" & sSubSectionID, True),_
						Array("Información de los empleados",_
							  "Obtenga un listado de los empleados registrados en el sistema incluyendo la información que usted necesite consultar.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=705" & sSubSectionID, True),_
						Array("<TITLE />REPORTES DE DICIEMBRE",_
							  "",_
							  "", "", (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("Fajillas por estado",_
							  "Listado de las fajillas de los vales decembrinos",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1404" & sSubSectionID, (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("<TITLE />FONAC",_
							  "",_
							  "", "", (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("Archivo para carga del SPEP del FONAC",_
							  "Genere el archivo en texto para el SPEP del FONAC.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1412" & sSubSectionID, (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("Archivo para contabilidad",_
							  "Genere el archivo en texto para el área de contabilidad.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1413" & sSubSectionID, (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("Cifras iniciales",_
							  "Cifras iniciales, global por concepto.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1411" & sSubSectionID, (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("Cifras para el pago de nómina",_
							  "Cifras globales por empresa para el pago de nómina FONAC.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1417" & sSubSectionID, (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("Detalle de las aportaciones por quincena",_
							  "Registros de los empleados y sus aportaciones en un archivo de texto.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1414" & sSubSectionID, (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("Reporte para el Fiduciario",_
							  "Genere el archivo en texto para el Fiduciario",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1415" & sSubSectionID, False),_
						Array("Respaldo de los empleados cotizantes por quincena",_
							  "Registros de los empleados inscritos en el FONAC, listados en un archivo de texto.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1416" & sSubSectionID, (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("<TITLE />Guardias y suplencias",_
							  "",_
							  "", "", True),_
						Array("Reporte de personal interno",_
							  "Totales para guardias, suplencias y/o PROVAC fitlrados por quincena y agrupados por delegaciones, para personal interno.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1420" & sSubSectionID, True),_
						Array("Reporte del concepto 15 (guardias), 31 (suplencias) y C5 (guardias PROVAC) de personal interno",_
							  "Listado de empleados y sus registros agrupados por periodo y totalizados por montos y horas, para personal interno.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1421" & sSubSectionID, True),_
						Array("Reporte de captura de personal externo",_
							  "Reporte en orden de captura de movimientos de guardias, suplencias y PROVAC correspondientes a una quincena, para personal externo.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1422" & sSubSectionID, True),_
						Array("Reporte de validación de personal externo",_
							  "Reporte de validación de movimientos de guardias, suplencias y PROVAC correspondientes a una quincena, para personal externo.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1423" & sSubSectionID, True),_
						Array("Volantes",_
							  "Generación de volantes de pago para guardias, suplencias y PROVAC correspondientes a una quincena, para personal externo.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1426" & sSubSectionID, True),_
						Array("Listado de firmas",_
							  "Listado de firmas de guardias, suplencias y PROVAC correspondientes a una quincena, para personal externo.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1427" & sSubSectionID, True),_
						Array("Reporte de totales",_
							  "Reporte de totales de guardias, suplencias y PROVAC correspondientes a una quincena.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1428" & sSubSectionID, True),_
						Array("Reporte concentrado por quincena",_
							  "Reporte concentrado por quincena del archivo histórico de totales.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1429" & sSubSectionID, True),_
						Array("Reporte estadístico de causas",_
							  "Reporte estadístico de causas de guardias, suplencias y PROVAC.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1430" & sSubSectionID, True),_
						Array("<TITLE />Programa de prevención del rezago en la atención médica",_
							  "",_
							  "", "", True),_
						Array("Consolidado de personal interno",_
							  "Registros e importes en la quincena seleccionada agrupados por unidad administrativa.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1424" & sSubSectionID, True),_
						Array("Consolidado de los prestadores de servicio",_
							  "Registros e importes en la quincena seleccionada agrupados por unidad administrativa.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1424&External=1" & sSubSectionID, True),_
						Array("Validación de captura de personal interno",_
							  "Montos pagados en la quincena seleccionada agrupados por unidad administrativa para personal interno.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1425" & sSubSectionID, True),_
						Array("Validación de captura de prestadores de servicios",_
							  "Montos pagados en la quincena seleccionada agrupados por unidad administrativa para prestadores de servicios.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1425&External=1" & sSubSectionID, True),_
						Array("<TITLE />Terceros",_
							  "",_
							  "", "", (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("Archivo para carga del SPEP por concepto",_
							  "Genere el archivo en texto para el SPEP de terceros.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1493" & sSubSectionID, (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("Reportes de terceros",_
							  "Seleccione la empresa y el concepto para obtener el reporte de terceros que necesita.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1492" & sSubSectionID, (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("Salida de archivo de terceros",_
							  "Seleccione la empresa y el concepto para generar el reporte de salida de terceros.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1491" & sSubSectionID, (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("<TITLE />Registros de auditoria",_
							  "",_
							  "", "", (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("Movimientos a los registros",_
							  "Listado de operaciones de conceptos e incidencias de los empleados.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1494" & sSubSectionID, (CInt(Request.Cookies("SIAP_SectionID")) = 4)),_
						Array("",_
							  "",_
							  "", "", False)_
					)
				Case 491 'Ejercicio bimestral del SAR
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Iniciar bimestre",_
							  "Iniciar el periodo de análisis.",_
							  "Images/MnLeftArrows.gif", "PayrollResumeForSar.asp?Action=StartPeriod" & sSubSectionID, True),_
						Array("Cargar resumen de nóminas",_
							  "Cargar el archivo del resumen de nóminas para el ejercicio bimestral del SAR.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=ProcessForSar&Load=PayrollSummary" & sSubSectionID, True),_
						Array("Cargar Padrón SAR",_
							  "Cargar la base correspondiente para los movimientos del padrón SAR.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=ProcessForSar&Load=BanamexCensus" & sSubSectionID, True),_
						Array("Cargar archivo de línea de captura",_
							  "Cargar el archivo de la línea de captura.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=ProcessForSar&Load=ConsarFile" & sSubSectionID, True),_
						Array("Actualización masiva de padrón SAR",_
							  "Cargar el padrón enviado por la CONSAR conteniendo la información actualizada de los empleados.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=ProcessForSar&Load=SarCensus" & sSubSectionID, True),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Mostrar resumen de nóminas",_
							  "Muestra el resumen de nóminas para el ejercicio bimestral.",_
							  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=PayrollResume" & sSubSectionID, True),_
						Array("Consultar padrón SAR",_
							  "Muestra el padrón Banamex actualizado.",_
							  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=BanamexCensus" & sSubSectionID, True),_
						Array("Mostrar comparativo de devengos",_
							  "Muestra el comparativo entre los historiales y el resumen de nóminas.",_
							  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=PayrollCompare" & sSubSectionID, True),_
						Array("Mostrar archivo de línea de captura",_
							  "Muestra el contenido de la línea de captura.",_
							  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=ConsarFile" & sSubSectionID, True),_
						Array("Consultar histórico de bajas de empleados",_
							  "Muestra el historial de las bajas registradas.",_
							  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=EmployeesDeleted" & sSubSectionID, True),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Generar comparativo de nóminas",_
							  "Compara la información del resumen de nóminas, las CLC y los archivos de nóminas.",_
							  "Images/MnLeftArrows.gif", "PayrollResumeForSar.asp?Action=payrollCompare" & sSubSectionID, True),_
						Array("Generar ejercicio bimestral",_
							  "Procesa la información de las nóminas del bimestre por fecha de imputación generando susbotales por bimestre y se realiza la contabilización de días.",_
							  "Images/MnLeftArrows.gif", "PayrollResumeForSar.asp?Action=EstrQna" & sSubSectionID, True),_
						Array("Generar proceso de altas, bajas y cambios",_
							  "Determina las altas y bajas de empleados así como los cambios en su información comparando el padrón contra las nómonas del bimestre.",_
							  "Images/MnLeftArrows.gif", "PayrollResumeForSar.asp?Action=employeesMovements" & sSubSectionID, True),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Generar reporte de cifras del bimestre",_
							  "Genere el reporte de cifras bimestral.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1028" & sSubSectionID, True),_
						Array("Generar reportes de movimientos del bimestre",_
							  "Genere los reportes de altas, cambios y bajas en el bimestre.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1031" & sSubSectionID, True),_
						Array("Generar reportes de dispersión por unidad administrativa y empresa",_
							  "Genere los reportes de dispersión por unidad administrativa y por empresa.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1032" & sSubSectionID, True),_
						Array("Generar reportes de aportaciones",_
							  "Genere los reportes de aportaciones voluntarias de patronales y de trabajadores, Cesantía y vejez y al FOVISSSTE.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1033" & sSubSectionID, True),_
						Array("Generar control y distribución de comprobantes de abono en cuenta de trabajadores",_
							  "Genere el control y distribución de comprobantes de abono en cuenta de trabajadores.",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1034" & sSubSectionID, True),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Borrar resumen de nóminas",_
							  "Borra el resumén de nóminas del último bimestre abierto.",_
							  "Images/MnLeftArrows.gif", "PayrollResumeForSar.asp?Action=deleteResume" & sSubSectionID, True),_
						Array("Cerrar bimestre",_
							  "Cierre el bimestre actual y genere los reportes finales.",_
							  "Images/MnLeftArrowsZIP.gif", "PayrollResumeForSar.asp?Action=ClosePeriod" & sSubSectionID, True)_
					)
				Case 5 'Presupuesto
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Estructuras programáticas",_
							  "Registre y modifique la estructura programática así como sus montos correspondientes.",_
							  "Images/MnSection52.gif", "Budget.asp?Section=Program", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_EstructurasProgramaticas & ",", vbBinaryCompare) > 0),_
						Array("Clasificador por objeto del gasto",_
							  "Catálogo jerárquico de partidas, subpartidas y tipos de pago.",_
							  "Images/MnBudget.gif", "Budget.asp?Section=Budget", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_ClasificadorPorObjetoDelGasto & ",", vbBinaryCompare) > 0),_
						Array("Costeo de plazas",_
							  "Costeo de plazas a partir de las condiciones indicadas en el armado del reporte.",_
							  "Images/MnAreas.gif", "Main_ISSSTE.asp?SectionID=53", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_CosteoDePlazas & ",", vbBinaryCompare) > 0),_
						Array("Reportes sobre el costeo de plazas",_
							  "A partir del costeo de plazas, arme los diferentes reportes utilizando los montos configurados.",_
							  "Images/MnSection55.gif", "Main_ISSSTE.asp?SectionID=56", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_ReportesSobreElCosteoDePlazas & ",", vbBinaryCompare) > 0),_
						Array("Registro de un costeo como original",_
							  "Con la información generada en los reportes del costeo de plazas registre su costeo como el presupuesto original del siguiente periodo.",_
							  "Images/MnSection58.gif", "Reports.asp?ReportID=1571", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_RegistroDeUnCosteoComoOriginal & ",", vbBinaryCompare) > 0),_
						Array("Administración del presupuesto",_
							  "Realice ajustes al presupuesto original.",_
							  "Images/MnSection510.gif", "Budget.asp?Section=Money", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_AdministracionDelPresupuesto & ",", vbBinaryCompare) > 0),_
						Array("Consulta de presupuesto",_
							  "Obtenga un reporte de los montos presupuestados por centros de trabajo, puestos, fondos, etc.",_
							  "Images/MnSection512.gif", "Reports.asp?ReportID=1504", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_ConsultaDePresupuesto & ",", vbBinaryCompare) > 0),_
						Array("Reportes",_
							  "Consulte el personal ocupado, las prestaciones a favor de los servidores públicos del ISSSTE y los trabajadores cotizantes al régimen del ISSSTE",_
							  "Images/MnAreas.gif", "Main_ISSSTE.asp?SectionID=58", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_Reportes & ",", vbBinaryCompare) > 0),_
						Array("Catálogos",_
							  "Altas, bajas y cambios de registros concernientes a los registros del sistema.",_
							  "Images/MnHumanResources.gif", "Catalogs.asp", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_Catalogos & ",", vbBinaryCompare) > 0),_
						Array("Ventanilla única",_
							  "Control de los trámite que ingresan a la Subdirección de Personal.",_
							  "Images/MnSection61.gif", "Main_ISSSTE.asp?SectionID=61", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_05_VentanillaUnica & ",", vbBinaryCompare) > 0)_
					)
				Case 53 'Presupuesto > Costeo de plazas
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Costeos guardados",_
							  "",_
							  "Images/MnLeftArrows.gif", "SavedReport.asp?AllReports=1&ReportToShow=1503", True),_
						Array("Generar un nuevo costeo",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1503", True)_
					)
				Case 56 'Presupuesto > Reportes sobre el costeo de plazas
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Proyecto de presupuesto",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1561", True),_
						Array("Archivo de carga para el SPEP",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1562", True),_
						Array("Formato único de movimientos presupuestales",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1563", True)_
					)
				Case 58 'Presupuesto > Reportes
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Reportes guardados",_
							  "Reportes generados por otros usuarios a partir de plantillas y filtros.",_
							  "Images/MnLeftArrows.gif", "SavedReport.asp?ReportType=1", True),_
						Array("Personal ocupado por rama de actividad",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1581", True),_
						Array("Personal ocupado y pago de sueldos y salarios en la administración pública federal",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1582", True),_
						Array("Prestaciones a favor de los servidores públicos del ISSSTE",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1583", True),_
						Array("Trabajadores cotizantes al régimen del ISSSTE",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1584", True)_
					)
				Case 6 'Departamento técnico
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Emisión de licencias por comisión sindical",_
							  "Registro y generación de oficios por comisión sindical",_
							  "Images/MnSection62.gif", "Main_ISSSTE.asp?SectionID=62", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_EmisionDeLicenciasPorComisionSindical & ",", vbBinaryCompare) > 0),_
						Array("Reportes",_
							  "Reportes para el departamento técnico",_
							  "Images/MnReports.gif", "Main_ISSSTE.asp?SectionID=64", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_Reportes & ",", vbBinaryCompare) > 0),_
						Array("Tablero de control de procesos",_
							  "Administre el estatus de los procesos registrados en el sistema.",_
							  "Images/MnSection63.gif", "TaCo.asp", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_TableroDeDontrolDeProcesos & ",", vbBinaryCompare) > 0),_
						Array("Ventanilla única",_
							  "Control de los trámite que ingresan a la Subdirección de Personal.",_
							  "Images/MnSection61.gif", "Main_ISSSTE.asp?SectionID=61", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_VentanillaUnica & ",", vbBinaryCompare) > 0)_
					)
				Case 61 'Departamento técnico > Ventanilla única
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Registro de documentos",_
							  "",_
							  "Images/MnLeftArrows.gif", "EmployeeSupport.asp?New=1", True),_
						Array("Búsqueda de documentos",_
							  "",_
							  "Images/MnLeftArrows.gif", "EmployeeSupport.asp?Search=1", True),_
						Array("Descargo de documentos",_
							  "",_
							  "Images/MnLeftArrows.gif", "EmployeeSupport.asp?Close=1", True),_
						Array("Asignación de documentos",_
							  "",_
							  "Images/MnLeftArrows.gif", "EmployeeSupport.asp?Assign=1", True),_
						Array("<LINE />",_
							  "",_
							  "", "", (CInt(Request.Cookies("SIAP_SectionID")) = 2)),_
						Array("Catálogo de acciones para turnado",_
							  "",_
							  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=PaperworkActions", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("Catálogo de procedencias",_
							  "",_
							  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=PaperworkSenders", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("Catálogo de remitentes y destinatarios para guías",_
							  "",_
							  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=PaperworkAddresses", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("Catálogo de tipos de asunto",_
							  "",_
							  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=SubjectTypes", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("Catálogo de tipos de trámite",_
							  "",_
							  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=PaperworkTypes", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("Catálogo de responsables",_
							  "",_
							  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=PaperworkOwners", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("Permisos de los usuarios para ver responsables",_
							  "",_
							  "Images/MnLeftArrows.gif", "EmployeeSupport.asp?Owners=1", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("<LINE />",_
							  "",_
							  "", "", True),_
						Array("Generación de volantes",_
							  "",_
							  "Images/MnLeftArrows.gif", "EmployeeSupport.asp?ForReport=1", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("Impresión de guías",_
							  "",_
							  "Images/MnLeftArrows.gif", "EmployeeSupport.asp?ForReport=1&ForGuides=1", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("Creación de listas",_
							  "",_
							  "Images/MnLeftArrows.gif", "Catalogs.asp?Action=PaperworkLists", True),_
						Array("Concentrado de control de correspondencia",_
							  "",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1607", True),_
						Array("Reporte de estatus de documentos",_
							  "",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1603", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("Reportes de asuntos defasados/resueltos",_
							  "",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1604", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("Reporte de Documentos",_
							  "",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1608", True),_
						Array("Listas de Entrega",_
							  "",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1609", True),_
						Array("Reporte ESP",_
							  "",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1619", True),_
						Array("<LINE />",_
							  "",_
							  "", "", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("Número de asuntos recibidos por destinatario",_
							  "",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1610", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("Número de asuntos recibidos por rangos de fecha",_
							  "",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1611", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("Número de asuntos recibidos por destinatario y rangos de fecha",_
							  "",_
							  "Images/MnLeftArrowsZIP.gif", "Reports.asp?ReportID=1612", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0),_
						Array("Asuntos pendientes de descargo",_
							  "",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1613", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_02_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_06_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_AdministrarVentanillaUnica & ",", vbBinaryCompare) > 0)_
					)
				Case 62 'Departamento técnico > Emisión de licencias por comisión sindical
					If (Not bSearchForm) And (Not bShowForm) And (Not bAction) And (Len(oRequest("DoSearch").Item) = 0) Then
						aMenuComponent(A_ELEMENTS_MENU) = Array(_
							Array("Registro de empleados para generar oficio de licencia",_
								  "Registre a los empleados para la generación de oficios de licencia sindical.",_
								  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesMovements&ReasonID=" & EMPLOYEES_DOCUMENTS_FOR_LICENSES, True),_
							Array("Búsqueda de licencias sindicales",_
								  "Listado de los empleados con licencia sindical",_
								  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=62&Search=1", True),_
							Array("Generación de oficios",_
								  "Generación de los oficios en formato word para su posterior impresión",_
								  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1605", True)_
						)
					Else
						If (lErrorNumber = 0) And bAction Then
							Call DisplayErrorMessage("Confirmación", "La información fue guardada con éxito.<BR /><FORM><INPUT TYPE=""BUTTON"" VALUE=""Registro de información"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=62&New=1'"" /><IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" /><INPUT TYPE=""BUTTON"" VALUE=""Regresar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=62'"" /></FORM>")
							Response.Write "<BR />"
						End If
						If lErrorNumber <> 0 Then
							Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
							lErrorNumber = 0
							Response.Write "<BR />"
						End If

						If bSearchForm Then
							lErrorNumber = Display62SearchForm(oRequest, oADODBConnection, sErrorDescription)
						ElseIf Len(oRequest("DoSearch").Item) > 0 Then
							lErrorNumber = DisplayCatalogsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, aCatalogComponent, sErrorDescription)
							If lErrorNumber = L_ERR_NO_RECORDS Then
								Call DisplayErrorMessage("Búsqueda vacía", "No existen registros que cumplan con los criterios de la búsqueda")
								Response.Write "<BR />"
								lErrorNumber = Display62SearchForm(oRequest, oADODBConnection, sErrorDescription)
							End If
						ElseIf bShowForm Then
							lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
						End If
					End If
				Case 64 'Departamento técnico > Reportes
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Reportes guardados",_
							  "Reportes generados por otros usuarios a partir de plantillas y filtros.",_
							  "Images/MnLeftArrows.gif", "SavedReport.asp?ReportType=1", True),_
						Array("Reporte de los empleados con licencia sindical",_
							  "Obtenga el listado de los empleados que fueron registrados para generar los oficios para el otorgamiento o cancelación de licencia sindical.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1606", True)_
					)
				Case 7 'Desconcentrados
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Personal",_
							  "Administre la información de los empleados que laboran en su área.",_
							  "Images/MnHumanResources.gif", "Main_ISSSTE.asp?SectionID=71", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_Personal & ",", vbBinaryCompare) > 0),_
						Array("Prestaciones",_
							  "Registre las prestaciones e incidencias a los empleados que laboran en su área.",_
							  "Images/MnSection2.gif", "Main_ISSSTE.asp?SectionID=72", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_Prestaciones & ",", vbBinaryCompare) > 0),_
						Array("Informática",_
							  "Ejecute los reportes sobre las nóminas de los empleados que laboran en su área.",_
							  "Images/MnSection4.gif", "Main_ISSSTE.asp?SectionID=73", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_Informatica & ",", vbBinaryCompare) > 0),_
						Array("Presupuesto",_
							  "Consulte el presupuesto para la unidad responsable.",_
							  "Images/MnBudget.gif", "Main_ISSSTE.asp?SectionID=74", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_Presupuesto & ",", vbBinaryCompare) > 0),_
						Array("Tablero de Control",_
							  "Tablero de control para desconcentrados",_
							  "Images/MnSection63.gif", "Projects.asp", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_TableroDeControl & ",", vbBinaryCompare) > 0),_
						Array("Reportes guardados",_
							  "Reportes generados por otros usuarios a partir de plantillas y filtros.",_
							  "Images/MnReports.gif", "SavedReport.asp?ReportType=1", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_ReportesGuardados & ",", vbBinaryCompare) > 0)_
					)
				Case 71 'Desconcentrados > Personal
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Asignación de número temporal de empleado",_
							  "Asigne un número al nuevo empleado antes de darlo de alta, este número se reasignará cuando el movimiento sea aplicado por la Subdirección de Personal.",_
							  "Images/MnLeftArrows.gif", "UploadInfo.asp?Action=EmployeesAssignTemporalNumber&ReasonID=67", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_AsignacionDeNumeroTemporalDeEmpleado & ",", vbBinaryCompare) > 0),_
						Array("Administración de personal",_
							  "Realice movimientos de nuevo ingreso, reingreso, bajas, etc. al personal del Instituto.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=712", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_AdministracionDePersonal & ",", vbBinaryCompare) > 0),_
						Array("Consulta de personal",_
							  "Consulte la información de los empleados, plaza, conceptos de pago, historia.",_
							  "Images/MnLeftArrows.gif", "Employees.asp", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_ConsultaDePersonal & ",", vbBinaryCompare) > 0),_
						Array("Consulta de plazas",_
							  "Consulte la información de las plazas, su historial y los ocupantes que ha tenido.",_
							  "Images/MnLeftArrows.gif", "Jobs.asp", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_ConsultaDePlazas & ",", vbBinaryCompare) > 0),_
						Array("Reportes",_
							  "Reportes del área de desconcentrados sobre personal.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=713&SubSectionID=1", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_ReportesPersonal & ",", vbBinaryCompare) > 0)_
					)
				Case 72 'Desconcentrados > Prestaciones
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Prestaciones e incidencias",_
							  "Registro de las incidencias y las prestaciones de los empleados",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=721", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_PrestacionesEIncidencias & ",", vbBinaryCompare) > 0),_
						Array("Pensión alimenticia",_
							  "Registre y valide las pensiones alimenticias de los empleados.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=723", (CInt(Request.Cookies("SIAP_SectionID")) = 7)),_
						Array("Entregas de hojas únicas de servicio",_
							  "Control de las entregas a los empleados de sus hojas únicas de servicio.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=267", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_EntregasDeHojasUnicasDeServicio & ",", vbBinaryCompare) > 0),_
						Array("Reportes",_
							  "Reportes del área de desconcentrados sobre prestaciones.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=723&SubSectionID=2", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_ReportesPrestaciones & ",", vbBinaryCompare) > 0)_
					)
				Case 73 'Desconcentrados > Informática
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Empleados",_
							  "Administre la información de los empleados.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=731&FromSectionID=7", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_Empleados & ",", vbBinaryCompare) > 0),_
						Array("Cheques y depósitos",_
							  "Administración de cheques, asignación de folios, reexpedición de cheques.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=732&FromSectionID=7", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_ChequesYDepositos & ",", vbBinaryCompare) > 0),_
						Array("Reportes",_
							  "Reportes sobre las nóminas.",_
							  "Images/MnLeftArrows.gif", "Main_ISSSTE.asp?SectionID=733&FromSectionID=7&SubSectionID=4", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_ReportesInformatica & ",", vbBinaryCompare) > 0)_
					)
				Case 74 'Desconcentrados > Presupuesto
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Consulta de presupuesto",_
							  "Obtenga un reporte de los montos presupuestados por centros de trabajo, puestos, fondos, etc.",_
							  "Images/MnLeftArrows.gif", "Reports.asp?ReportID=1701", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_07_Presupuesto & ",", vbBinaryCompare) > 0)_
					)
                Case 8 'Nuevo perfil
					aMenuComponent(A_ELEMENTS_MENU) = Array(_
						Array("Trámites al personal",_
							  "Control de los trámite que ingresan a la Subdirección de Personal.",_
							  "Images/userMngmt.png", "Main_ISSSTE.asp?SectionID=61", StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_08_VentanillaUnica & ",", vbBinaryCompare) > 0)_
					)
			End Select
			aMenuComponent(B_USE_DIV_MENU) = True
			If iSectionID < 10 Then
				Call DisplayMenuInTwoColumns(aMenuComponent)
			ElseIf InStr(1, ",31,", ("," & iSectionID & ","), vbBinaryCompare) > 0 Then
				Call DisplayMenuInTwoColumns(aMenuComponent)
			ElseIf InStr(1, ",267,351,352,356,369,423,424,425,426,", ("," & iSectionID & ","), vbBinaryCompare) > 0 Then
				Call DisplayMenuInTwoColumns(aMenuComponent)
			ElseIf InStr(1, ",371,", ("," & iSectionID & ","), vbBinaryCompare) > 0 Then
			Else
				Call DisplayMenuInThreeSmallColumns(aMenuComponent)
			End If
		Select Case iSectionID
			Case EMPLOYEES_SERVICE_SHEET
			Case Else
				Response.Write "</TABLE>"
		End Select
		If lErrorNumber <> 0 Then
			Call DisplayErrorMessage("Mensaje del sistema", sErrorDescription)
			Response.Write "<BR />"
		End If%>
		<!-- #include file="_Footer.asp" -->
	</BODY>
</HTML>