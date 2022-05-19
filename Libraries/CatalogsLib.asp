<!-- #include file="BanamexCensusComponent.asp" -->
<!-- #include file="CalendarComponent.asp" -->
<!-- #include file="EmployeeFieldComponent.asp" -->
<!-- #include file="EmployeesLib.asp" -->
<!-- #include file="EmployeeAddComponent.asp" -->
<!-- #include file="FormComponent.asp" -->
<!-- #include file="FormFieldComponent.asp" -->
<!-- #include file="PayrollResumeForSarComponent.asp" -->
<!-- #include file="ProfileComponent.asp" -->
<!-- #include file="ProfessionalRiskComponent.asp" -->
<!-- #include file="TaCoTaskComponent.asp" -->
<!-- #include file="UserComponent.asp" -->
<!-- #include file="ZoneComponent.asp" -->
<%
Function InitializeCatalogs(oRequest)
'************************************************************
'Purpose: To initialize each component that is a catalog
'Inputs:  oRequest
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeCatalogs"
	Dim lParentID
	Dim sNames

	Call InitializeCalendarComponent(oRequest, aCalendarComponent)
	Call InitializeFormComponent(oRequest, aFormComponent)
	Call InitializeFormFieldComponent(oRequest, aFormFieldComponent)
	Call InitializeProfileComponent(oRequest, aProfileComponent)
	Call InitializeTaskComponent(oRequest, aTaskComponent)
	Call InitializeUserComponent(oRequest, aUserComponent)
	Call InitializeZoneComponent(oRequest, aZoneComponent)
	Call InitializeCatalogComponent(oRequest, aCatalogComponent)
	Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
Case "SubStates"
	aCatalogComponent(S_NAME_CATALOG) = "Entidades federativas 222"
	aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Padre,Clave,Nombre abreviado,Nombre,Banco para pago de cheques,Activo"
	aCatalogComponent(S_ORDER_CATALOG) = "StateName"
	aCatalogComponent(N_NAME_CATALOG) = 3
	aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
	aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "StateID,ParentID,StateCode,StateShortName,StateName,BankID,Active"
	aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
	aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1,1"
	aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
	aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,5,5,5,6,0"
	aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
	aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,5,10,100,0,0"
	aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
	aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0"
	aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
	aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0,-1,0"
	aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
	aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,,,,0"
	aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
	aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1," & oRequest("ParentID").Item & ",,,,3,1"
	aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
	aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1," & oRequest("ParentID").Item & ",,,,3,1"
	aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
	aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞBanks;,;BankID;,;BankName;,;;,;BankName;,;;,;Ninguno;;;-1ÞÞÞ"
	aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
	aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
	aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
	aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,4", ",")
aCatalogComponent(S_URL_CATALOG) = "ParentID=" & oRequest("ParentID").Item
aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (ParentID=" & oRequest("ParentID").Item & ")"
aCatalogComponent(S_CHECK_EXISTENCY_CONDITION_CATALOG) = " And (ParentID=" & oRequest("ParentID").Item & ")"
		Case "Absences"
			aCatalogComponent(S_NAME_CATALOG) = "Ausencias"
			aCatalogComponent(S_ORDER_CATALOG) = "AbsenceName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Tipo de incidencia,¿Es deducción?,¿Está justificada?,Activo,Aplica a,Estatus del empleado,Justificación,Conceptos para nómina,Aplica para periodo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "AbsenceID,AbsenceShortName,AbsenceName,AbsenceTypeID2,IsDeduction,IsJustified,Active,AppliesToID,AppliesToEmployesStatusID,JustificationID,ConceptsIDs,IsForPeriod"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1,1,1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,6,0,0,0,8,8,6,8,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,255,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,00,,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,,-1,1,1,-1,-1,-1,-1,0"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,,-1,1,1,-1,-1,-1,-1,0"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞAbsenceTypes2;,;AbsenceTypeID;,;AbsenceTypeName;,;(Active=1);,;AbsenceTypeName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞConcepts;,;ConceptID;,;ConceptShortName, ConceptName;,;(EndDate=30000000);,;ConceptShortName, ConceptName;,;;,;Ninguno;;;-1ÞÞÞStatusEmployees;,;StatusID;,;StatusName;,;(StatusID>-1);,;StatusName;,;;,;Ninguno;;;-1ÞÞÞJustifications;,;JustificationID;,;JustificationShortName, JustificationName;,;(Active=1);,;JustificationShortName, JustificationName;,;;,;Ninguno;;;-1ÞÞÞConcepts;,;ConceptID;,;ConceptShortName, ConceptName;,;(EndDate=30000000);,;ConceptShortName, ConceptName;,;;,;Ninguno;;;-1ÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
		Case "Antiquities"
			aCatalogComponent(S_NAME_CATALOG) = "Antigüedades"
			aCatalogComponent(S_ORDER_CATALOG) = "StartYears, EndYears"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,De,Hasta"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "AntiquityID,AntiquityName,StartYears,EndYears"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,2,2"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,255,4,4"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,15,15"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,100,100"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,0,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,0,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3", ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
		Case "AreaLevelTypes"
			aCatalogComponent(S_NAME_CATALOG) = "Tipos de niveles para las áreas"
			aCatalogComponent(S_ORDER_CATALOG) = "AreaLevelTypeID"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "AreaLevelTypeID,AreaLevelTypeName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
		Case "AreaTypes"
			aCatalogComponent(S_NAME_CATALOG) = "Tipos de área"
			aCatalogComponent(S_ORDER_CATALOG) = "AreaTypeShortName, AreaTypes.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "AreaTypeID,AreaTypeShortName,AreaTypeName,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,2,255,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 3
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 4
		Case "AttentionLevels"
			aCatalogComponent(S_NAME_CATALOG) = "Niveles de atención"
			aCatalogComponent(S_ORDER_CATALOG) = "AttentionLevelShortName, AttentionLevels.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "AttentionLevelID,AttentionLevelShortName,AttentionLevelName,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,2,255,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 3
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 4
		Case "BanamexCensus"
			aCatalogComponent(S_TABLE_NAME_CATALOG) = "DM_PADRON_BANAMEX"
			aCatalogComponent(S_NAME_CATALOG) = "Padrón Banamex"
			aCatalogComponent(S_ORDER_CATALOG) = "EmployeeID Asc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "Versión,Núm. Empleado,RFC,CURP,Núm. seguro social,Apellido paterno,Apellido materno,Nombre,CT,Fecha de nacimiento,Estado de nacimiento,Genero,Fecha de ingreso,Fecha_Cot,Salario,FOVISSSTE,Periodo,Estatus,Abre-cierra,Estado civil,Domicilio,Colonia,Ciudad,C.P.,Estado,Nombramiento,Cuenta Afore,Clave ICEFA,Núm. Control interno,Motivo baja,Salario base V,Pago integrado,Días laborados,Días inhabiles,Días ausencia,Contribución del empleado,Monto de contribución,Fecha de inicio,Fecha fin,Comentarios"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "u_version,EmployeeID,RFC,CURP,SocialSecurityNumber,EmployeeLastName,EmployeeLastName2,EmployeeName,CT,BirthDate,BirthState,GenderShortName,JoinDate,CotDate,Salary,Fovi,Period,Status,ChangeFlag,MaritalStatusID,Address,Colony,City,ZipZone,State,Nombram,Afore,ICEFA,ICNumber,mot_baja,Salary_v,FullPay,WorkingDays,InabilityDays,AbsenceDays,EmployeeContributions,EmployeeContributionsAmount,StartDate,EndDate,Comments"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,4,5,5,5,5,5,5,4,4,4,5,4,4,4,4,4,4,4,4,5,5,5,5,5,5,4,4,4,4,4,4,4,4,4,4,4,4,4,5"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,255,255,255,255,255,255,0,0,0,255,0,0,0,0,0,0,0,0,255,255,255,255,255,255,0,0,0,0,0,0,0,0,0,0,0,0,0,255"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 3
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 4
		Case "Banks"
			aCatalogComponent(S_NAME_CATALOG) = "Bancos"
			aCatalogComponent(S_ORDER_CATALOG) = "BankName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Clave,Nombre,País,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "BankID,BankShortName,BankName,CountryID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,6,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,255,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,100000000,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,0,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,0,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞCountries;,;CountryID;,;CountryName;,;(Active=1);,;CountryName;,;;,;Ninguno;;;-1ÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("2", ",")
		Case "BankAccounts"
			aCatalogComponent(S_NAME_CATALOG) = "Cuentas bancarias del Instituto"
			aCatalogComponent(S_ORDER_CATALOG) = "StateName, AccountNumber"
			aCatalogComponent(N_NAME_CATALOG) = 3
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,ID,Banco,No. Cuenta,Entidad federativa,Fecha de inicio,Fecha de término,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "AccountID,EmployeeID,BankID,AccountNumber,StateID,StartDate,EndDate,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1,0,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,6,5,6,1,1,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,0,100,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0," & N_START_YEAR & "," & N_START_YEAR & ",0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,,," & Year(Date) + 10 & "," & Year(Date) + 10 & ","
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,-1,-1,,9,0,30000000,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,-1,-1,,9,0,30000000,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞBanks;,;BankID;,;BankName;,;(Active=1);,;BankName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞStates;,;StateID;,;StateName;,;(Active=1);,;StateName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("4,2,3", ",")
		Case "Branches"
			aCatalogComponent(S_NAME_CATALOG) = "Ramas"
			aCatalogComponent(S_ORDER_CATALOG) = "BranchShortName, Branches.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "BranchID,BranchShortName,BranchName,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,2,255,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 3
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 4
		Case "BudgetsActiveDuties"
			aCatalogComponent(S_NAME_CATALOG) = "Subfunción activa"
			aCatalogComponent(S_ORDER_CATALOG) = "ActiveDutyShortName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Clave,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "ActiveDutyID,ActiveDutyShortName,ActiveDutyName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,10,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
		Case "BudgetsActivities1"
			aCatalogComponent(S_NAME_CATALOG) = "Actividad institucional"
			aCatalogComponent(S_ORDER_CATALOG) = "ActivityShortName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Clave,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "ActivityID,ActivityShortName,ActivityName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,10,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
		Case "BudgetsActivities2"
			aCatalogComponent(S_NAME_CATALOG) = "Actividad presupuestaria"
			aCatalogComponent(S_ORDER_CATALOG) = "ActivityShortName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Clave,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "ActivityID,ActivityShortName,ActivityName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,10,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
		Case "BudgetsConfineTypes"
			aCatalogComponent(S_NAME_CATALOG) = "Ámbito"
			aCatalogComponent(S_ORDER_CATALOG) = "ConfineTypeShortName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Clave,Nombre"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "ConfineTypeID,ConfineTypeShortName,ConfineTypeName"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,10,255"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
			aCatalogComponent(N_ACTIVE_CATALOG) = -2
		Case "BudgetsDuties"
			aCatalogComponent(S_NAME_CATALOG) = "Función"
			aCatalogComponent(S_ORDER_CATALOG) = "DutyShortName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Clave,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "DutyID,DutyShortName,DutyName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,10,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
		Case "BudgetsFunds"
			aCatalogComponent(S_NAME_CATALOG) = "Fondo"
			aCatalogComponent(S_ORDER_CATALOG) = "FundShortName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Clave,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "FundID,FundShortName,FundName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,10,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
		Case "BudgetsLocations"
			aCatalogComponent(S_NAME_CATALOG) = "Municipio"
			aCatalogComponent(S_ORDER_CATALOG) = "LocationShortName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Clave,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "LocationID,LocationShortName,LocationName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,10,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
		Case "BudgetsProcesses"
			aCatalogComponent(S_NAME_CATALOG) = "Proceso"
			aCatalogComponent(S_ORDER_CATALOG) = "ProcessShortName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Clave,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "ProcessID,ProcessShortName,ProcessName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,10,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
		Case "BudgetsProgramDuties"
			aCatalogComponent(S_NAME_CATALOG) = "Programa presupuestario"
			aCatalogComponent(S_ORDER_CATALOG) = "ProgramDutyShortName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Clave,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "ProgramDutyID,ProgramDutyShortName,ProgramDutyName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,10,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
		Case "BudgetsPrograms"
			aCatalogComponent(S_NAME_CATALOG) = "Programa"
			aCatalogComponent(S_ORDER_CATALOG) = "ProgramShortName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Clave,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "ProgramID,ProgramShortName,ProgramName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,10,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
		Case "BudgetsRegions"
			aCatalogComponent(S_NAME_CATALOG) = "Región"
			aCatalogComponent(S_ORDER_CATALOG) = "RegionShortName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Clave,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "RegionID,RegionShortName,RegionName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,10,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
		Case "BudgetsSpecificDuties"
			aCatalogComponent(S_NAME_CATALOG) = "Subfunción específica"
			aCatalogComponent(S_ORDER_CATALOG) = "SpecificDutyShortName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Clave,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "SpecificDutyID,SpecificDutyShortName,SpecificDutyName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,10,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
		Case "BudgetTypes"
			aCatalogComponent(S_NAME_CATALOG) = "Partidas presupuestales"
			aCatalogComponent(S_ORDER_CATALOG) = "BudgetTypeID"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "BudgetTypeID,BudgetTypeName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,50,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
		Case "BudgetTypes2"
			aCatalogComponent(S_NAME_CATALOG) = "Estructuras programáticas"
			aCatalogComponent(S_ORDER_CATALOG) = "BudgetTypeID"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "BudgetTypeID,BudgetTypeName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,50,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
		Case "CashierOffices"
			aCatalogComponent(S_NAME_CATALOG) = "Pagadurías SIPE"
			aCatalogComponent(S_ORDER_CATALOG) = "CashierOfficeShortName, CashierOffices.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "CashierOfficeID,CashierOfficeShortName,CashierOfficeName,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,255,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 3
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 4
		Case "CenterSubtypes"
			aCatalogComponent(S_NAME_CATALOG) = "Subtipos de centro de trabajo"
			aCatalogComponent(S_ORDER_CATALOG) = "CenterSubtypeShortName, CenterSubtypes.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 2
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Tipo de centro de trabajo,Código,Nombre,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "CenterSubtypeID,CenterTypeID,CenterSubtypeShortName,CenterSubtypeName,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,6,5,5,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,5,255,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,15,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,100000000,,," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1," & oRequest("CenterTypeID").Item & ",,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1," & oRequest("CenterTypeID").Item & ",,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞCenterTypes;,;CenterTypeID;,;CenterTypeShortName, CenterTypeName;,;(EndDate=30000000);,;CenterTypeShortName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4,5", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 4
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 5
			'aCatalogComponent(S_URL_CATALOG) = "CenterTypeID=" & oRequest("CenterTypeID").Item
			'aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (CenterTypeID=" & oRequest("CenterTypeID").Item & ")"
			'aCatalogComponent(S_CHECK_EXISTENCY_CONDITION_CATALOG) = " And (CenterTypeID=" & oRequest("CenterTypeID").Item & ")"
		Case "CenterTypes"
			aCatalogComponent(S_NAME_CATALOG) = "Tipos de centro de trabajo"
			aCatalogComponent(S_ORDER_CATALOG) = "CenterTypeShortName, CenterTypes.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "CenterTypeID,CenterTypeShortName,CenterTypeName,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,255,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 3
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 4
			'aCatalogComponent(S_URL_PARAMETERS_CATALOG) = "Catalogs.asp?Action=CenterSubtypes&CenterTypeID=<FIELD_0 />"
			''aCatalogComponent(S_URL_CATALOG) = "Action=CenterSubtypes&CenterTypeID=<FIELD_0 />"
		Case "Companies"
			aCatalogComponent(S_NAME_CATALOG) = "Sociedades y empresas"
			aCatalogComponent(S_ORDER_CATALOG) = "CompanyShortName, Companies.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 2
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Padre,Código,Nombre,Tipo,RFC,Dirección,Ciudad,Código postal,Estado,País,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "CompanyID,ParentID,CompanyShortName,CompanyName,CompanyTypeID,CompanyRFC,CompanyAddress,CompanyCity,CompanyZip,StateID,CountryID,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,0,0,0,0,1,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			If Not B_ISSSTE Then
				aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,5,5,6,5,5,5,5,6,6,1,1,11,11,0"
			Else
				aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,5,5,6,11,11,11,11,11,11,1,1,11,11,0"
			End If
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,5,255,0,14,2000,100,10,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0,0,0,0,0,0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,0,0,0,0,0,0,0,0,100000000,100000000," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			If Len(oRequest("ParentID").Item) > 0 Then
				lParentID = CLng(oRequest("ParentID").Item)
			Else
				lParentID = -1
			End If
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1," & lParentID & ",,,-1,,,,,9,0," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,-1,,,-1,,,,,9,0," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞCompanyTypes;,;CompanyTypeID;,;CompanyTypeName;,;;,;CompanyTypeName;,;(EndDate=30000000) And (Active=1);,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞStates;,;StateID;,;StateName;,;(Active=1);,;StateName;,;;,;Ninguno;;;-1ÞÞÞCountries;,;CountryID;,;CountryName;,;;,;CountryName;,;(Active=1);,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("2,3,4,11,12", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 11
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 12
			aCatalogComponent(S_URL_PARAMETERS_CATALOG) = "Catalogs.asp?ParentID=<FIELD_0 />&Action=Companies&ReadOnly=" & oRequest("ReadOnly").Item
			aCatalogComponent(S_URL_CATALOG) = "ParentID=" & lParentID & "&ReadOnly=" & oRequest("ReadOnly").Item
			aCatalogComponent(S_QUERY_CONDITION_CATALOG) = "(ParentID=" & lParentID & ")"
			aCatalogComponent(S_CHECK_EXISTENCY_CONDITION_CATALOG) = "(ParentID=" & lParentID & ")"
		Case "CompanyTypes"
			aCatalogComponent(S_NAME_CATALOG) = "Tipos de compañía"
			aCatalogComponent(S_ORDER_CATALOG) = "CompanyTypeID"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "CompanyTypeID,CompanyTypeName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,50,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
		Case "ConfineTypes"
			aCatalogComponent(S_NAME_CATALOG) = "Tipos de ámbito para las áreas"
			aCatalogComponent(S_ORDER_CATALOG) = "ConfineTypeShortName, ConfineTypes.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "ConfineTypeID,ConfineTypeShortName,ConfineTypeName,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,2,255,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0" & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 3
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 4
		Case "Countries"
			aCatalogComponent(S_NAME_CATALOG) = "Países"
			aCatalogComponent(S_ORDER_CATALOG) = "CountryID"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "CountryID,CountryName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,100,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
		Case "CreditTypes"
			aCatalogComponent(S_NAME_CATALOG) = "Tipos de crédito"
			aCatalogComponent(S_ORDER_CATALOG) = "CreditTypeShortName,CreditTypeName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "CreditTypeID,CreditTypeShortName,CreditTypeName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
		Case "Currencies"
			aCatalogComponent(S_NAME_CATALOG) = "Monedas"
			aCatalogComponent(S_ORDER_CATALOG) = "CurrencyName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Símbolo,Valor,Llave,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "CurrencyID,CurrencyName,CurrencySymbol,CurrencyValue,CurrencyKey,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,2,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,50,10,10,3,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,0,,1"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,$,1,MXN,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,$,1,MXN,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3", ",")
			aCatalogComponent(S_URL_PARAMETERS_CATALOG) = "Catalogs.asp?Action=CurrenciesHistoryList&CurrencyID=<FIELD_0 />&MonthForFilter=" & Right(("0" & Month(Date())), Len("00")) & "&YearForFilter=" & Year(Date())
		Case "CurrenciesHistoryList"
			aCatalogComponent(S_NAME_CATALOG) = "Historial de valores"
			aCatalogComponent(S_ORDER_CATALOG) = "CurrencyDate"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Fecha,Valor"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "CurrencyID,CurrencyDate,CurrencyValue"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,1,2"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,20"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,2000,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000," & Year(Date()) + 1 & ",0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,20100105,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,20100105,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
			aCatalogComponent(S_URL_CATALOG) = "CurrencyDate=<FIELD_1 />"
			If Len(oRequest("CurrencyDate").Item) > 0 Then
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (CurrencyDate=" & oRequest("CurrencyDate").Item & ")"
			ElseIf Len(oRequest("CurrencyDateYear").Item) > 0 Then
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (CurrencyDate=" & oRequest("CurrencyDateYear").Item & oRequest("CurrencyDateMonth").Item & oRequest("CurrencyDateDay").Item & ")"
			Else
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (CurrencyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ")"
			End If
			aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = False
		Case "DocsLibrary", "Documents"
			aCatalogComponent(S_NAME_CATALOG) = "Procedimientos"
			aCatalogComponent(S_ORDER_CATALOG) = "DocumentName, StartDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Ruta,Tipo,Fecha de inicio,Fecha de término,Fecha de modificación,Descripción,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "DocumentID,DocumentName,FilePath,FileType,StartDate,EndDate,ModificationDate,Description,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,0,1,0,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,11,11,1,1,11,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,255,0,0,0,0,0,2000,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0," & Year(Date()) - 10 & "," & Year(Date()) - 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,," & Year(Date()) + 1 & "," & Year(Date()) + 10 & ",,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,4,5", ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(S_ADDITIONAL_FORM_SCRIPT_CATALOG) = "return IsFileReady(oForm);"
		Case "DocumentsForLicenses"
			aCatalogComponent(S_NAME_CATALOG) = "Empleados con licencia sindical"
			aCatalogComponent(S_ORDER_CATALOG) = "DocumentForLicenseNumber"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,No. oficio,No. cancelación,Plantilla,No. solicitud,No. empleado,Tipo de licencia,Fecha del documento,Fecha inicio licencia,Fecha término licencia,Fecha de cancelación,Usuario"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "DocumentForLicenseID,DocumentForLicenseNumber,DocumentForCancelLicenseNumber,DocumentTemplate,RequestNumber,EmployeeID,LicenseSyndicateTypeID,DocumentLicenseDate,LicenseStartDate,LicenseEndDate,LicenseCancelDate,UserID"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,0,1,1,1,1,1,1,1,0,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,5,5,5,6,1,1,1,1,11"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,25,25,50,25,6,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,,,,,,,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,,,,,,,,,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG)  = "-1,,,,,,,,,,0,"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,,,,,,,,0,"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞLicenseSyndicateTypes;,;LicenseSyndicateTypeID;,;LicenseSyndicateTypeName;,;(LicenseSyndicateTypeID>0);,;LicenseSyndicateTypeName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,4,5", ",")
		Case "EconomicZones"
			aCatalogComponent(S_NAME_CATALOG) = "Zonas económicas"
			aCatalogComponent(S_ORDER_CATALOG) = "EconomicZoneCode, EconomicZoneName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "EconomicZoneID,EconomicZoneCode,EconomicZoneName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
		Case "EmployeesAntiquitiesLKP"
			aCatalogComponent(S_NAME_CATALOG) = "Antigüedad federal"
			aCatalogComponent(S_ORDER_CATALOG) = "AntiquityDate"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "No. del empleado,Fecha de inicio,Fecha de término,Institución,Puesto,Años,Meses,Días,Retroactivo,Fecha de modificación,Usuario"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "EmployeeID,AntiquityDate,EndDate,FederalCompanyID,PositionID,AntiquityYears,AntiquityMonths,AntiquityDays,ForRetro,ModifyDate,UserID"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1,1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,1,1,6,11,4,4,4,11,11,11"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "6,0,0,0,0,2,2,2,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,5,15,15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0," & N_FORM_START_YEAR & "," & N_FORM_START_YEAR & ",0,,0,0,0,,,"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "999999," & Year(Date()) & "," & Year(Date()) & ",0,,99,12,30,,,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,0,0,-1,-1,0,0,0,0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN)
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,0,0,-1,-1,0,0,0,0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN)
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞFederalCompanies;,;FederalCompanyID;,;FederalCompanyName;,;;,;FederalCompanyName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("0,1,2,3,5,6,7", ",")
			aCatalogComponent(S_FIELDS_TO_SUM_CATALOG) = "5,6,7"
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			If Len(oRequest("PrevAntiquityDate").Item) > 0 Then
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeeID=" & oRequest("EmployeeID").Item & ") And (AntiquityDate=" & oRequest("PrevAntiquityDate").Item & ")"
			ElseIf Len(oRequest("AntiquityDate").Item) > 0 Then
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeeID=" & oRequest("EmployeeID").Item & ") And (AntiquityDate=" & oRequest("AntiquityDate").Item & ")"
			ElseIf Len(oRequest("AntiquityDateYear").Item) > 0 Then
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeeID=" & oRequest("EmployeeID").Item & ") And (AntiquityDate=" & oRequest("AntiquityDateYear").Item & oRequest("AntiquityDateMonth").Item & oRequest("AntiquityDateDay").Item & ")"
			ElseIf Len(oRequest("EmployeeID").Item) > 0 Then
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (EmployeeID=" & oRequest("EmployeeID").Item & ")"
			End If
			If Len(oRequest("AntiquityDate").Item) > 0 Then
				aCatalogComponent(S_CHECK_EXISTENCY_CONDITION_CATALOG) = " And (EmployeeID=" & oRequest("EmployeeID").Item & ") And (AntiquityDate=" & oRequest("AntiquityDate").Item & ")"
			Else
				aCatalogComponent(S_CHECK_EXISTENCY_CONDITION_CATALOG) = " And (EmployeeID=" & oRequest("EmployeeID").Item & ") And (AntiquityDate=" & oRequest("AntiquityDateYear").Item & oRequest("AntiquityDateMonth").Item & oRequest("AntiquityDateDay").Item & ")"
			End If
			aCatalogComponent(S_CANCEL_BUTTON_ACTION_CATALOG) = "window.location.href='Main_ISSSTE.asp?SectionID=261&EmployeeID=" & oRequest("EmployeeID").Item & "&DoSearch=1'"
			aCatalogComponent(N_ID_CATALOG) = -1
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "function CheckDatesForAntiquity() {" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "if ((parseInt(document.CatalogFrm.EndDateDay.value) * parseInt(document.CatalogFrm.EndDateMonth.value) * parseInt(document.CatalogFrm.EndDateYear.value)) > 0) {" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "if ((parseInt(document.CatalogFrm.AntiquityDateDay.value) + parseInt(document.CatalogFrm.AntiquityDateMonth.value)*100 + parseInt(document.CatalogFrm.AntiquityDateYear.value) * 10000) > (parseInt(document.CatalogFrm.EndDateDay.value) + parseInt(document.CatalogFrm.EndDateMonth.value)*100 + parseInt(document.CatalogFrm.EndDateYear.value) * 10000)) {" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "alert('Favor de verificar la vigencia del movimiento');" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "document.CatalogFrm.AntiquityDateDay.focus();" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "return false;" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "} else {" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "if (parseInt(document.CatalogFrm.EndDateYear.value + document.CatalogFrm.EndDateMonth.value + document.CatalogFrm.EndDateDay.value) > " & Left(GetSerialNumberForDate(""), Len("00000000")) & ") {" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "alert('La fecha de fin de la antigüedad no puede ser posterior al día de hoy.');" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "alert('Se regresara false');" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "document.CatalogFrm.EndDateDay.focus();" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "return false;" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "} else {" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "return true;" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "}" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "}" & vbNewLine
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "} else {" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "alert('Favor de verificar la vigencia del movimiento');" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "document.CatalogFrm.AntiquityDateDay.focus();" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "return false;" & vbNewLine
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "}" & vbNewLine
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "} // End of CheckDatesForAntiquity" & vbNewLine
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "//--></SCRIPT>" & vbNewLine
			aCatalogComponent(S_ADDITIONAL_FORM_SCRIPT_CATALOG) = "return CheckDatesForAntiquity(oForm);"
		Case "EmployeesDocs"
			aCatalogComponent(S_NAME_CATALOG) = "Entregas de hojas únicas de servicio"
			aCatalogComponent(S_ORDER_CATALOG) = "EmployeeID, DocumentDate, DocumentHour"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,No. empleado,Fecha de entrega,Hora de entrega,Fecha de recepción,Hora de recepción,Fecha de trámite,Hora de trámite,No. de oficio,Sello digital,Clave documento"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "RecordID,EmployeeID,DocumentDate,DocumentHour,Document2Date,Document2Hour,Document3Date,Document3Hour,DocumentNumber, Sign1, Sign2"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1,1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,4,1,3,1,3,11,11,5,11,11"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,6,0,0,0,0,0,0,100,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,15,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,1," & Year(Date()) & ",0," & Year(Date()) & ",0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,999999," & Year(Date()) & ",23," & Year(Date()) & ",23," & Year(Date()) & ",23,0,176,32"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,,,,,,,,"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,,,,,,,,"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4,5,8", ",")
			aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = False
		Case "EmployeesKardex"
			aCatalogComponent(S_NAME_CATALOG) = "Validación del proceso de selección de personal"
			aCatalogComponent(S_ORDER_CATALOG) = "EmployeeLastName, EmployeeLastName2, EmployeeName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Apellido Paterno,Apellido Materno,Acta de nacimiento,Currículum vitae,2 Fotos,2 cartas de recomendación,Certificado médico,Constancia de estudios,Tipo de puesto,Constancia de la especialidad,Diplomado o documento que lo avale como técnico,Título,Cédula profesional,Constancia de especialidad (para médicos especialistas),Hizo examen de conocimientos,Hizo examen psicométrico,Constancia de envío a registro"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "EmployeeID,EmployeeName,EmployeeLastName,EmployeeLastName2,Requirement1,Requirement2,Requirement3,Requirement4,Requirement5,Requirement6,PositionTypeID,Requirement7,Requirement8,Requirement9,Requirement10,Requirement11,Requirement12,Requirement13,Requirement14"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,5,0,0,0,0,0,0,6,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,100,100,100,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞPositionTypes2;,;PositionTypeID;,;PositionTypeName;,;;,;PositionTypeName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""ShowRequirements(this.value);""ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("2,3,1,10", ",")
			aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = False
		Case "EmployeesKardex2"
			aCatalogComponent(S_NAME_CATALOG) = "Consulta de los registros de escalafón"
			aCatalogComponent(S_ORDER_CATALOG) = "EmployeeID"
			aCatalogComponent(N_NAME_CATALOG) = 0
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Verificación documental,Validación de sus conocimientos"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "EmployeeID,Requirement1,Requirement2"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,0,0"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "11,0,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,0,0"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,0,0"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(S_ADDITIONAL_FORM_SCRIPT_CATALOG) = "return CheckEmployeeNumber(oForm);"
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("0,1,2", ",")
			aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = False
		Case "EmployeesKardex3"
			aCatalogComponent(S_NAME_CATALOG) = "Registro de información"
			aCatalogComponent(S_ORDER_CATALOG) = "RecordID"
			aCatalogComponent(N_NAME_CATALOG) = 0
			If Len(oRequest("DoSearch").Item) > 0 Then
				aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Procedimiento para,Fecha de inicio de trámite,Propuesto por,Nombre,Apellido paterno,Apellido materno,Puesto,Unidad administrativa,Requisitos documentales por grupo,<B>Requerimientos documentales por entregar</B>,<B>Procesos de selección de personal por fecha</B><BR /></FONT></TD></TR><TR><TD><FONT FACE=""Arial"" SIZE=""2"">Recepción de documentos,Evaluación de conocimientos,&nbsp;&nbsp;&nbsp;Estatus,Evaluación psicológica,&nbsp;&nbsp;&nbsp;Estatus,Envío a registro en bolsa de trabajo,Envío a registro en escalafón,Envío al área de recursos humanos,Fecha de modificación,Modificó"
			Else
				aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Procedimiento para,Fecha de inicio de trámite,Propuesto por,<B>Datos del aspirante o trabajador</B></FONT></TD><TD>&nbsp;</TD></TR><TR><TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Nombre,Apellido paterno,Apellido materno,Puesto,Unidad administrativa,Requisitos documentales por grupo,<B>Requerimientos documentales por entregar</B><IMG SRC=""Images/Transparent.gif"" WIDTH=""60"" HEIGHT=""1"" />Utilizando la tecla <B>CTRL</B> seleccione los requerimientos que el aspirante vaya entregando.,<B>Procesos de selección de personal por fecha</B><BR /></FONT></TD></TR><TR NAME=""TempDiv"" ID=""TempDiv""><TD><FONT FACE=""Arial"" SIZE=""2"">Recepción de documentos,Evaluación de conocimientos,&nbsp;&nbsp;&nbsp;Estatus,Evaluación psicológica,&nbsp;&nbsp;&nbsp;Estatus,Envío a registro en bolsa de trabajo,Envío a registro en escalafón,Envío al área de recursos humanos,Fecha de modificación,Modificó"
			End If
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "RecordID,KardexTypeID,StartDate,KardexOriginID,PersonName,PersonLastName,PersonLastName2,PositionID,AreaID,RequirementsTypeID,Requirements,DocumentsDate,KnowledgeDate,KnowledgeStatusID,PsychologicDate,PsychologicStatusID,Registration1Date,Registration2Date,Registration3Date,ModifyDate,UserID"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1,0,1,1,1,0,0,0,0,0,0,0,0,0,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,6,1,6,5,5,5,6,6,6,8,1,1,6,1,6,1,1,1,11,11"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,0,0,100,100,100,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "0,0,0,0,0,0,0,0,,00,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0," & Year(Date()) & ",0,0,0,0,0,0,0,0," & Year(Date()) & "," & Year(Date()) & ",0," & Year(Date()) & ",0," & Year(Date()) & "," & Year(Date()) & "," & Year(Date()) & ",0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "0,0," & Year(Date()) & ",0,0,0,0,0,0,0,0," & Year(Date()) & "," & Year(Date()) + 1 & ",0," & Year(Date()) + 1 & ",0," & Year(Date()) & "," & Year(Date()) & "," & Year(Date()) & ",0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,-1," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",-1,,,,-1," & aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) & ",-1,,0,0,-1,0,-1,0,0,0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN)
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,-1," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",-1,,,,-1," & aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) & ",-1,,0,0,-1,0,-1,0,0,0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN)
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞKardexTypes;,;KardexTypeID;,;KardexTypeName;,;(KardexTypeID>-1) And (Active=1);,;KardexTypeID;,;;,;Ninguno;;;-1ÞÞÞÞÞÞKardexOrigins;,;KardexOriginID;,;KardexOriginName;,;(Active=1);,;KardexOriginName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞPositions;,;PositionID;,;PositionShortName, PositionName;,;(Positions.PositionID>0) And (EndDate=30000000) And (Active=1);,;PositionShortName;,;;,;Ninguno;;;-1ÞÞÞAreas;,;AreaID;,;AreaCode, AreaName;,;(AreaID>-1) And (ParentID=-1) And (EndDate=30000000) And (Active=1);,;AreaCode, AreaName;,;;,;Ninguno;;;-1ÞÞÞRequirementsTypes;,;RequirementsTypeID;,;RequirementsTypeName;,;(Active=1);,;RequirementsTypeName;,;;,;Ninguno;;;-1ÞÞÞKardexRequirements;,;KardexRequirementID;,;KardexRequirementName;,;(Active=1);,;KardexRequirementName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞStatusKnowledges;,;StatusID;,;StatusName;,;(StatusID>-1) And (Active=1);,;StatusName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞStatusPsychologics;,;StatusID;,;StatusName;,;(StatusID>-1) And (Active=1);,;StatusName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞ onChange=""ShowKardexRequirements(this.form, '-1')""ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""ShowKardexRequirements(this.form, this.value)""ÞÞÞ onChange=""UpdateDocumentsDate(this.form);""ÞÞÞÞÞÞ onChange=""UpdateKnowledgeStatus(this.form);""ÞÞÞ onChange=""UpdatePsychologicDate(this.form);""ÞÞÞ onChange=""UpdatePsychologicStatus(this.form);""ÞÞÞ onChange=""UpdateRegistrationDate(this.form, -1);""ÞÞÞ onChange=""UpdateRegistrationDate(this.form, 1);""ÞÞÞ onChange=""UpdatePsychologicDate(this.form); UpdateRegistrationDate(this.form, 2);""ÞÞÞ onChange=""UpdateRegistrationDate(this.form, 3);""ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(S_ADDITIONAL_FORM_SCRIPT_CATALOG) = "return CheckKardexRequirements(oForm);"
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4,5,6,8", ",")
			aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = False
			aCatalogComponent(S_ORDER_CATALOG) = "EmployeesKardex3.StartDate"
		Case "EmployeesKardex4"
			aCatalogComponent(S_NAME_CATALOG) = "Registro de información de la bolsa de trabajo"
			aCatalogComponent(S_ORDER_CATALOG) = "EmployeesKardex4.StartDate"
			aCatalogComponent(N_NAME_CATALOG) = 0
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Registro de,Número de registro general,Número de registro individual,No. empleado,No. plaza,Puesto solicitado,Turno solicitado,Adscripción solicitada,Servicio solicitado,Rama solicitada,Fecha de registro,Fecha de resolución,Observaciones,Subcomisión mixta de Escalafón"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "RecordID,KardexChangeTypeID,KardexNumber1,KardexNumber2,EmployeeID,JobID,PositionID,JourneyID,AreaID,ServiceID,BranchID,StartDate,EndDate,Comments,DocumentNumber"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1,1,1,1,1,1,1,0,0,0"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,6,5,5,5,5,6,6,6,6,6,1,1,5,5"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,10,10,6,6,0,0,0,0,0,0,0,2000,100"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "0,0,0,0,15,15,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,1,1,0,0,0,0,0,2009," & Year(Date()) & ",,"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "0,0,0,0,999999,999999,0,0,0,0,0," & Year(Date()) & "," & Year(Date()) & ",,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,-1,,,,,-1,-1,-1,-1,-1," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0,,"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,-1,,,,,-1,-1,-1,-1,-1," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0,,"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞKardexChangeTypes;,;KardexChangeTypeID;,;KardexChangeTypeName;,;(KardexChangeTypes.Active=1);,;KardexChangeTypeName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞPositions;,;PositionID;,;PositionShortName, PositionName;,;(Positions.EndDate=30000000) And (Positions.Active=1);,;PositionShortName;,;;,;Ninguno;;;-1ÞÞÞJourneys;,;JourneyID;,;JourneyShortName, JourneyName;,;(Journeys.EndDate=30000000) And (Journeys.Active=1);,;JourneyShortName;,;;,;Ninguno;;;-1ÞÞÞAreas;,;AreaID;,;AreaCode, AreaName;,;(Areas.EndDate=30000000) And (Areas.ParentID<>-1);,;AreaCode;,;;,;Ninguno;;;-1ÞÞÞServices;,;ServiceID;,;ServiceShortName, ServiceName;,;(Services.EndDate=30000000) And (Services.Active=1);,;ServiceShortName;,;;,;Ninguno;;;-1ÞÞÞBranches;,;BranchID;,;BranchShortName, BranchName;,;(Branches.EndDate=30000000) And (Branches.Active=1);,;BranchShortName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞ onChange=""ShowKardexFields(this.value);""ÞÞÞÞÞÞÞÞÞ /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeID.value, 'EmployeesInfo&Full=1', 'SearchEmployeeNumberIFrame', 'EmployeeFrm.EmployeeID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" ÞÞÞ /><A HREF=""javascript: SearchRecord(document.CatalogFrm.JobID.value, 'JobsInfo&SendPosition=1', 'SearchJobNumberIFrame', 'EmployeeFrm.JobID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de plaza"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4,5,6", ",")
			aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = False
			aCatalogComponent(N_ACTIVE_CATALOG) = -2
			aCatalogComponent(S_ADDITIONAL_FORM_SCRIPT_CATALOG) = "return CheckEmployeeForm();"
			aCatalogComponent(S_CANCEL_BUTTON_ACTION_CATALOG) = "window.location.href='Main_ISSSTE.asp?SectionID=28'"
		Case "EmployeesKardex5"
			aCatalogComponent(S_NAME_CATALOG) = "Registro de información de escalafón"
			aCatalogComponent(S_ORDER_CATALOG) = "EmployeesKardex5.StartDate"
			aCatalogComponent(N_NAME_CATALOG) = 0
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Tipo de registro,Referido por,Refrendo,Región,Subcomisión mixta de Bolsa de Trabajo en,Rama,Puesto,Nombre,Apellido paterno,Apellido materno,Fecha de registro,Escolaridad,Parentesco,Tiempo de servicio en el Instituto,Tiempo en el registro de Bolsa de Trabajo,Nominación,Motivo de baja"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "RecordID,Kardex5TypeID,Kardex5OriginID,DocumentNumber,KardexZone,KardexOffice,BranchID,PositionID,EmployeeName,EmployeeLastName,EmployeeLastName2,StartDate,SchoolarshipID,Relationship,ServiceYears,KardexYears,Nomination,Reasons"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1,1,1,1,1,0,1,1,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,6,6,5,5,5,6,6,5,5,5,1,6,5,1,1,5,5"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,0,10,255,255,0,0,100,100,100,0,0,100,0,0,255,255"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "-1,0,0,,,,-1,-1,,,,2009,-1,,2009,2009,,"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "0,0,0,0,0,0,0,0,,,," & Year(Date()) & ",0,0," & Year(Date()) & "," & Year(Date()) & ",0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,0,0,,,,-1,-1,,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",-1,,0,0,,"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,0,0,,,,-1,-1,,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",-1,,0,0,,"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞKardex5Types;,;Kardex5TypeID;,;Kardex5TypeName;,;(Kardex5Types.Active=1);,;Kardex5TypeName;,;;,;Ninguno;;;-1ÞÞÞKardex5Origins;,;Kardex5OriginID;,;Kardex5OriginName;,;(Kardex5Origins.Active=1);,;Kardex5OriginName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞBranches;,;BranchID;,;BranchShortName, BranchName;,;(Branches.EndDate=30000000) And (Branches.Active=1);,;BranchShortName;,;;,;Ninguno;;;-1ÞÞÞPositions;,;PositionID;,;PositionShortName, PositionName;,;(Positions.EndDate=30000000) And (Positions.Active=1);,;PositionShortName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞSchoolarships;,;SchoolarshipID;,;SchoolarshipName;,;(Schoolarships.Active=1);,;SchoolarshipName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞ onChange=""ShowKardexFields(this.value);""ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("11,8,9,10,1,2,3", ",")
			aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = False
			aCatalogComponent(N_ACTIVE_CATALOG) = -2
			'aCatalogComponent(S_ADDITIONAL_FORM_SCRIPT_CATALOG) = "return CheckEmployeeForm();"
			aCatalogComponent(S_CANCEL_BUTTON_ACTION_CATALOG) = "window.location.href='Main_ISSSTE.asp?SectionID=28'"
		Case "EmployeesSpecialJourneys"
			aCatalogComponent(S_NAME_CATALOG) = "Guardias y Suplencias"
			aCatalogComponent(S_ORDER_CATALOG) = "EmployeeID, EmployeesSpecialJourneys.StartDate, EmployeesSpecialJourneys.EndDate"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,ID del empleado,No. del empleado,Nombre,Apellido paterno,Apellido materno,RFC,CURP,No. del empleado a suplir,Puesto,Adscripción,Servicio,Nivel/subnivel,Horas laboradas,Horario,Riesgo profesional,Tipo de registro,Folio de autorización,Fecha desde,Fecha hasta,Hora de entrada,Hora de salida,Turno,Días/horas reportadas,Movimiento,Factor,Motivo,<BR />Comentarios,Percepción,Registrado por,Fecha de registro,Quincena de aplicación,Eliminado,Eliminado por,Fecha de eliminación,Aplicación de la eliminación,Activación"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "RecordID,EmployeeID,EmployeeNumber,EmployeeName,EmployeeLastName,EmployeeLastName2,RFC,CURP,OriginalEmployeeID,PositionID,AreaID,ServiceID,LevelID,WorkingHours,ShiftID,RiskLevelID,SpecialJourneyID,DocumentNumber,StartDate,EndDate,StartHour,EndHour,JourneyID,WorkedHours,MovementID,FactorID,ReasonID,Comments,ConceptAmount,AddUserID,AddDate,AppliedDate,Removed,RemoveUserID,RemovedDate,AppliedRemoveDate,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,1,1,1,1,1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			If lEmployeeID < 800000 Then
				aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,5,11,11,11,11,11,11,11,6,11,11,11,11,11,11,5,1,1,11,11,6,2,6,11,6,5,2,11,11,6,11,11,11,11,11"
			Else
				aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,11,5,5,5,5,5,11,6,6,6,6,6,11,11,11,5,1,1,11,11,6,2,6,11,6,5,2,11,11,6,11,11,11,11,11"
			End If
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,6,100,100,100,13,18,6,0,0,0,0,0,4,0,0,50,0,0,0,0,0,4,0,0,0,2000,20,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,15,0,0,0,0,0,0,0,0,0,0,0,0,15,0,0,0,0,0,0,0,0,5,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0," & Year(Date()) - 1 & "," & Year(Date()) - 1 & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,999999,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0," & Year(Date()) & "," & Year(Date()) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,-1,,,,,,,-1,-1,-1,-1,-1,0,-1,-1,-1,," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0,0,-1,0,-1,-1,-1,,0.00," & aLoginComponent(N_USER_ID_LOGIN) & "," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",-1,0,-1,-1,-1,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,-1,,,,,,,-1,-1,-1,-1,-1,0,-1,-1,-1,," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0,0,-1,0,-1,-1,-1,,0.00," & aLoginComponent(N_USER_ID_LOGIN) & "," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",-1,0,-1,-1,-1,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞPositionsSpecialJourneysLKP, Positions;,;Distinct Positions.PositionID;,;PositionShortName, PositionName, 'Horas laboradas:' As Temp, PositionsSpecialJourneysLKP.WorkingHours;,;(PositionsSpecialJourneysLKP.PositionID=Positions.PositionID) And (Positions.EndDate=30000000) And (Positions.Active=1);,;PositionShortName, PositionsSpecialJourneysLKP.WorkingHours;,;;,;Ninguno;;;-1ÞÞÞAreas;,;AreaID;,;AreaCode, AreaName;,;(ParentID=-1) Or (ParentID=-2);,;AreaCode;,;;,;Seleccione un puesto;;;-1ÞÞÞServices;,;ServiceID;,;ServiceShortName, ServiceName;,;(ServiceID=-1) And (ServiceID=-2);,;ServiceShortName;,;;,;Seleccione un puesto;;;-1ÞÞÞLevels;,;LevelID;,;LevelShortName;,;(LevelID=-1) And (LevelID=-2);,;LevelShortName;,;;,;Seleccione un puesto;;;-1ÞÞÞLevels;,;LevelID;,;LevelShortName;,;(LevelID=-1) And (LevelID=-2);,;LevelShortName;,;;,;Seleccione un puesto;;;-1ÞÞÞShifts;,;ShiftID;,;ShiftShortName, ShiftName;,;(EndDate=30000000) And (Active=1);,;ShiftShortName;,;;,;Ninguno;;;-1ÞÞÞRiskLevels;,;RiskLevelID;,;RiskLevelName;,;(Active=1);,;RiskLevelName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞSpecialJourneys;,;JourneyID;,;JourneyShortName, JourneyName;,;(RecordTypeID In (-1," & oRequest("SectionID").Item & ")) And (Active=1);,;JourneyShortName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞSpecialJourneysMovements;,;MovementID;,;MovementShortName, MovementName;,;(RecordTypeID In (-1," & oRequest("SectionID").Item & ")) And (Active=1);,;MovementShortName;,;;,;Ninguno;;;-1ÞÞÞSpecialJourneysFactors;,;FactorID;,;FactorShortName;,;(RecordTypeID In (-1," & oRequest("SectionID").Item & ")) And (Active=1);,;FactorShortName;,;;,;Ninguno;;;-1ÞÞÞSpecialJourneysReasons;,;ReasonID;,;ReasonShortName, ReasonName;,;(RecordTypeID In (-1," & oRequest("SectionID").Item & ")) And (Active=1);,;ReasonShortName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞPayrolls;,;PayrollID;,;PayrollDate, PayrollName;,;(IsActive_5=1) And (IsClosed<>1);,;PayrollID;,;;,;ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), "ÞÞÞ")
			If lEmployeeID < 800000 Then
				aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ onChange=""document.CatalogFrm.CheckEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeNumber.value, 'EmployeesGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&AreaID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10) & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckEmployeeID"" ID=""CheckEmployeeIDHdn"" VALUE="""" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onChange=""document.CatalogFrm.CheckOriginalEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.OriginalEmployeeID.value, 'EmployeesGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&Original=1&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado a suplir"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckOriginalEmployeeID"" ID=""CheckOriginalEmployeeIDHdn"" VALUE="""" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" ÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeID.value, 'RecordsForGyS&TheRecordID=' + document.CatalogFrm.RecordID.value + '&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&OriginalEmployeeID=' + document.CatalogFrm.OriginalEmployeeID.value + '&RiskLevelID=' + document.CatalogFrm.RiskLevelID.value + '&MovementID=' + document.CatalogFrm.MovementID.value + '&StartDate=' + BuildDateString(document.CatalogFrm.StartDateYear.value, document.CatalogFrm.StartDateMonth.value, document.CatalogFrm.StartDateDay.value) + '&EndDate=' + BuildDateString(document.CatalogFrm.EndDateYear.value, document.CatalogFrm.EndDateMonth.value, document.CatalogFrm.EndDateDay.value) + '&WorkingHours=' + document.CatalogFrm.WorkingHours.value + '&WorkedHours=' + document.CatalogFrm.WorkedHours.value + '&PayrollDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeJourneysIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar las horas reportadas por quincena para el empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""ReportedHours"" ID=""ReportedHoursHdn"" VALUE="""" ÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" ÞÞÞÞÞÞÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			Else
				aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ onChange=""document.CatalogFrm.EmployeeID.value=this.value; document.CatalogFrm.ReportedHours.value='';"" /><INPUT TYPE=""HIDDEN"" NAME=""CheckEmployeeID"" ID=""CheckEmployeeIDHdn"" VALUE="""" ÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""document.CatalogFrm.CheckEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.RFC.value, 'ExternalGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&CURP=' + document.CatalogFrm.CURP.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar RFC del empleado externo"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckEmployeeID"" ID=""CheckEmployeeIDHdn"" VALUE="""" ÞÞÞÞÞÞ onChange=""document.CatalogFrm.CheckOriginalEmployeeID.value=''; document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.OriginalEmployeeID.value, 'EmployeesGyS&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&Original=1&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado a suplir"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""CheckOriginalEmployeeID"" ID=""CheckOriginalEmployeeIDHdn"" VALUE="""" ÞÞÞ onChange=""SearchRecord('P', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=-1&ServiceID=-1&LevelID=-1&WorkingHours=-1&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞ onChange=""SearchRecord('A', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&ServiceID=-1&LevelID=-1&WorkingHours=-1&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞ onChange=""SearchRecord('S', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&ServiceID=' + document.CatalogFrm.ServiceID.value + '&LevelID=-1&WorkingHours=-1&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞ onChange=""SearchRecord('L', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AeraID=' + document.CatalogFrm.AeraID.value + '&ServiceID=' + document.CatalogFrm.ServiceID.value + '&LevelID=' + document.CatalogFrm.LevelID.value + '&WorkingHours=-1&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞ onChange=""SearchRecord('W', 'PositionsGyS&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&ServiceID=' + document.CatalogFrm.ServiceID.value + '&LevelID=' + document.CatalogFrm.LevelID.value + '&WorkingHours=' + document.CatalogFrm.WorkingHours.value + '&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&RecordDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'CatalogFrm')"" ÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" ÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeID.value, 'RecordsForGyS&TheRecordID=' + document.CatalogFrm.RecordID.value + '&RecordType=" & CInt(oRequest("SectionID").Item) - 422 & "&OriginalEmployeeID=' + document.CatalogFrm.OriginalEmployeeID.value + '&PositionID=' + document.CatalogFrm.PositionID.value + '&AreaID=' + document.CatalogFrm.AreaID.value + '&RiskLevelID=' + document.CatalogFrm.RiskLevelID.value + '&MovementID=' + document.CatalogFrm.MovementID.value + '&StartDate=' + BuildDateString(document.CatalogFrm.StartDateYear.value, document.CatalogFrm.StartDateMonth.value, document.CatalogFrm.StartDateDay.value) + '&EndDate=' + BuildDateString(document.CatalogFrm.EndDateYear.value, document.CatalogFrm.EndDateMonth.value, document.CatalogFrm.EndDateDay.value) + '&WorkingHours=' + document.CatalogFrm.WorkingHours.value + '&WorkedHours=' + document.CatalogFrm.WorkedHours.value + '&PayrollDate=' + document.CatalogFrm.AppliedDate.value, 'SearchEmployeeJourneysIFrame', 'CatalogFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar las horas reportadas por quincena para el empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><INPUT TYPE=""HIDDEN"" NAME=""ReportedHours"" ID=""ReportedHoursHdn"" VALUE="""" ÞÞÞ onChange=""document.CatalogFrm.ReportedHours.value='';"" ÞÞÞÞÞÞÞÞÞÞÞÞ onFocus=""DoBlock();"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			End If
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
			aCatalogComponent(S_ADDITIONAL_FORM_SCRIPT_CATALOG) = "return ValitadeCatalogFields(window.document.CatalogFrm);"
			If StrComp(oRequest("Internal").Item, "1", vbBinaryCompare) = 0 Then
				aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("2,3,4,5,6,10,17,18,19,22,23,24,28,31", ",")
			Else
				aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("3,4,5,6,10,17,18,19,22,23,24,28,31", ",")
			End If
			aCatalogComponent(S_FIELDS_TO_SUM_CATALOG) = "28"
			aCatalogComponent(S_CANCEL_BUTTON_ACTION_CATALOG) = "window.location.href='Main_ISSSTE.asp?SectionID=" & iSectionID & "';"
			aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = False
			aCatalogComponent(S_ADD_LINES_BEFORE_FIELDS_CATALOG) = ",17,"
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = "<INPUT TYPE=""HIDDEN"" NAME=""TempStartDate"" ID=""TempStartDateHdn"" VALUE="""" /><INPUT TYPE=""HIDDEN"" NAME=""TempEndDate"" ID=""TempEndDateHdn"" VALUE="""" />"
		Case "EmployeesRequirements"
			aCatalogComponent(S_NAME_CATALOG) = "Requisitos de documentación para movimiento de personal"
			aCatalogComponent(S_ORDER_CATALOG) = "ReasonShortName, EmployeeRequirementName"
			aCatalogComponent(N_NAME_CATALOG) = -1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Tipos de movimientos,Requisito"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "EmployeeRequirementID,ReasonID,EmployeeRequirementName"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,6,5"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,255"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,0,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,1,"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,1,"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞReasons;,;ReasonID;,;ReasonShortName, ReasonName;,;;,;ReasonShortName;,;;,;Ninguna;;;-1ÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
			aCatalogComponent(N_ACTIVE_CATALOG) = -2
		Case "EmployeesRevisions"
		Case "EmployeeTypes"
			aCatalogComponent(S_NAME_CATALOG) = "Tipos de tabulador"
			aCatalogComponent(S_ORDER_CATALOG) = "EmployeeTypeShortName, EmployeeTypes.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Es operativo,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "EmployeeTypeID,EmployeeTypeShortName,EmployeeTypeName,TypeID,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,0,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,11,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,255,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,0," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,4,5", ",")
			aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (EmployeeTypeID Not In (8,9,10))"
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 4
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 5
		Case "FederalCompanies"
			aCatalogComponent(S_NAME_CATALOG) = "Dependencias gubernamentales"
			aCatalogComponent(S_ORDER_CATALOG) = "FederalCompanyName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "FederalCompanyID,FederalCompanyName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
		Case "GeneratingAreas"
			aCatalogComponent(S_NAME_CATALOG) = "Áreas generadoras"
			aCatalogComponent(S_ORDER_CATALOG) = "GeneratingAreaShortName, GeneratingAreas.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "GeneratingAreaID,GeneratingAreaShortName,GeneratingAreaName,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,255,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 3
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 4
		Case "GenericPositions"
			aCatalogComponent(S_NAME_CATALOG) = "Puestos genéricos"
			aCatalogComponent(S_ORDER_CATALOG) = "GenericPositions.GenericPositionID"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "GenericPositionID,GenericPositionName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("0,1", ",")
		Case "GroupGradeLevels"
			aCatalogComponent(S_NAME_CATALOG) = "Grupos, grados, niveles"
			aCatalogComponent(S_ORDER_CATALOG) = "GroupGradeLevelShortName, GroupGradeLevels.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "GroupGradeLevelID,GroupGradeLevelShortName,GroupGradeLevelName,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,100,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 3
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 4
		Case "Handicaps"
			aCatalogComponent(S_NAME_CATALOG) = "Discapacidades"
			aCatalogComponent(S_ORDER_CATALOG) = "HandicapName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "HandicapID,HandicapName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,100,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
		Case "Holidays"
			aCatalogComponent(S_NAME_CATALOG) = "Días de asueto"
		Case "JobTypes"
			aCatalogComponent(S_NAME_CATALOG) = "Tipos de plaza"
			aCatalogComponent(S_ORDER_CATALOG) = "JobTypeShortName, JobTypeName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "JobTypeID,JobTypeShortName,JobTypeName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
		Case "Journeys"
			aCatalogComponent(S_NAME_CATALOG) = "Turnos"
			aCatalogComponent(S_ORDER_CATALOG) = "JourneyShortName, Journeys.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Inicio,Fin,Inicio,Fin,Tipo de jornada,Factor para cálculos de nómina,Jornada adicional,Factor de turno de suplencias,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "JourneyID,JourneyShortName,JourneyName,StartHour1,EndHour1,StartHour2,EndHour2,JourneyTypeID,JourneyFactor,SpecialJourneyFactor1,SpecialJourneyFactor2,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,0,0,1,1,1,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,11,11,11,11,6,2,2,2,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,100,0,0,0,0,1,4,4,4,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,15,15,15,15,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0,0,0,1,0,0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,23,23,23,23,4,24,24,24," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,0,0,0,0,1,1,1,5," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,0,0,0,0,1,1,1,5," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ<OPTION VALUE=""1"">1</OPTION><OPTION VALUE=""2"">2</OPTION><OPTION VALUE=""3"">3</OPTION><OPTION VALUE=""4"">4</OPTION>ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,7,8,11,12", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 11
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 12
		Case "Levels"
			aCatalogComponent(S_NAME_CATALOG) = "Niveles"
			aCatalogComponent(S_ORDER_CATALOG) = "LevelShortName, Levels.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 2
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nivel,Nombre,Descripción,Estatus,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "LevelID,LevelShortName,LevelName,LevelDescription,StatusID,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,0,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,5,6,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,100,255,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,,100000000," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,,-1," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,,-1," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞStatusLevels;,;StatusID;,;StatusName;,;(Active=1);,;StatusName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4,5,6", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 5
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 6
		Case "MaritalStatus"
			aCatalogComponent(S_NAME_CATALOG) = "Estado civil"
			aCatalogComponent(S_ORDER_CATALOG) = "MaritalStatusID"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "MaritalStatusID,MaritalStatusName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,100,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
		Case "MedicalAreas"
			aCatalogComponent(S_NAME_CATALOG) = "Matriz UNIMED"
			aCatalogComponent(S_ORDER_CATALOG) = "MedicalAreasID"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "Renglón,Empresa,Tipo de reporte UNIMED,Puesto,Servicio,Anexo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "MedicalAreasID,CompanyID,MedicalAreasTypeID,PositionID,ServiceID,ColumnNumber"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,6,6,6,6,5"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,0,0,0,2"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,0,0,0,0,,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,9,9,9,9,"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,9,9,9,9,"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞCompanies;,;CompanyID;,;CompanyShortName, CompanyName;,;(CompanyID>0) And (EndDate=30000000) And (Active=1);,;CompanyShortName;,;;,;Ninguno;;;-1ÞÞÞMedicalAreasTypes;,;MedicalAreasTypeID;,;MedicalAreasTypeName;,;(MedicalAreasTypeID>0) And (Active=1);,;MedicalAreasTypeName;,;;,;Ninguno;;;-1ÞÞÞPositions;,;PositionID;,;PositionShortName, PositionName;,;(Positions.PositionID>0) And (EndDate=30000000) And (Active=1);,;PositionShortName;,;;,;Ninguno;;;-1ÞÞÞServices;,;ServiceID;,;ServiceShortName, ServiceName;,;(Services.ServiceID>0) And (EndDate=30000000) And (Active=1);,;ServiceShortName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("0,1,2,3,4,5", ",")
		Case "MedicalAreasTypes"
			aCatalogComponent(S_NAME_CATALOG) = "Tipos de reporte UNIMED"
			aCatalogComponent(S_ORDER_CATALOG) = "MedicalAreasTypeName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "MedicalAreasTypeID,MedicalAreasTypeName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
		Case "OccupationTypes"
			aCatalogComponent(S_NAME_CATALOG) = "Tipos de ocupación"
			aCatalogComponent(S_ORDER_CATALOG) = "OccupationTypeShortName, OccupationTypeID"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "OccupationTypeID,OccupationTypeShortName,OccupationTypeName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
		Case "PaperworkActions"
			aCatalogComponent(S_NAME_CATALOG) = "Acciones para turnado"
			aCatalogComponent(S_ORDER_CATALOG) = "PaperworkActionShortName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Clave,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "PaperworkActionID,PaperworkActionShortName,PaperworkActionName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,100,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
		Case "PaperworkAddresses"
			aCatalogComponent(S_NAME_CATALOG) = "Remitentes y destinatarios para guías"
			aCatalogComponent(S_ORDER_CATALOG) = "StateName, AddressLevel"
			aCatalogComponent(N_NAME_CATALOG) = 2
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nivel,Nombre,Puesto,Calle y número,Colonia,Ciudad,Estado,Código postal,Teléfono"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "AddressID,AddressLevel,OwnerName,PositionName,OwnerAddress,OwnerAddress2,OwnerCity,StateID,OwnerZipCode,OwnerPhone"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1,1,1,1,0"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,4,5,5,5,5,5,6,5,5"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,1,255,255,255,100,100,0,5,100"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,15,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,1,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,3,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,1,,,,,,9,,"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,1,,,,,,9,,"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("2,3,4,5,6,7,8", ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞStates;,;StateID;,;StateName;,;;,;StateName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(N_ACTIVE_CATALOG) = -2
		Case "PaperworkLists"
			aCatalogComponent(S_NAME_CATALOG) = "Creación de listas"
			aCatalogComponent(S_ORDER_CATALOG) = "ListID Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Número de lista,Procedencia,Dirigido a,Trámites"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "ListID,ListNumber,SenderName,RecipientName,PaperworkIDs"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,4,5,5,11"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,6,255,255,2000"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,,,"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,100000,,,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,,"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,,"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			lErrorNumber = GetConsecutiveID(oADODBConnection, 1062, aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1), sErrorDescription)
			lErrorNumber = GetConsecutiveID(oADODBConnection, 1062, aCatalogComponent(AS_DEFAULT_VALUES_CATALOG)(1), sErrorDescription)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,4", ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<TD VALIGN=""TOP"">"
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<FONT FACE=""Arial"" SIZE=""2"">No. de folio:&nbsp;</FONT>"
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<INPUT TYPE=""TEXT"" NAME=""PaperworkNumberTemp"" ID=""PaperworkNumberTempTxt"" SIZE=""10"" MAXLENGTH=""10"" CLASS=""TextFields"" /><BR />"
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</TD>"
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<TD VALIGN=""TOP"">"
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "&nbsp;<A HREF=""javascript: AddPaperworkToList()""><IMG SRC=""Images/BtnCrclAdd.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Turnar"" BORDER=""0"" /></A>&nbsp;"
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<BR /><BR />"
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "&nbsp;<A HREF=""javascript: RemovePaperworkToList()""><IMG SRC=""Images/BtnCrclDelete.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Remover"" BORDER=""0"" /></A>&nbsp;"
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</TD>"
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<TD VALIGN=""TOP"">"
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<SELECT NAME=""PaperworkNumbers"" ID=""PaperworkNumbersLst"" SIZE=""3"" MULTIPLE=""1"" CLASS=""Lists""></SELECT>"
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</TD>"
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "</TR></TABLE><BR />"
			aCatalogComponent(N_ACTIVE_CATALOG) = -2
		Case "PaperworkOwners"
			aCatalogComponent(S_NAME_CATALOG) = "Responsables"
			aCatalogComponent(S_ORDER_CATALOG) = "LEVELID, OWNERID"
			aCatalogComponent(N_NAME_CATALOG) = 0
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Pertenece a,Nivel,Descripción,Empleado,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "OWNERID,PARENTID,LEVELID,OWNERNAME,EMPLOYEEID,ACTIVE"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "6,4,1,255,6,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,15,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,1,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "1000000,0,0,0,999999,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("0,1,2,3,4", ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞ onChange=""document.CatalogFrm.OwnerIDTemp.value = '';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.PARENTID.value, 'OwnerParentID&LevelID=' + document.CatalogFrm.LEVELID.value, 'SearchOwnerParentIDIFrame', 'CatalogFrm.OwnerIDTemp')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar la jefatura"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><BR /><IFRAME SRC=""SearchRecord.asp"" NAME=""SearchOwnerParentIDIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""26""></IFRAME><INPUT TYPE=""HIDDEN"" NAME=""OwnerIDTemp"" ID=""OwnerIDTempTxt"" ÞÞÞÞÞÞÞÞÞ onChange=""document.CatalogFrm.EmployeeIDTemp.value = '';"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EMPLOYEEID.value, 'EmployeeNumber', 'SearchEmployeeNumberIFrame', 'CatalogFrm.EmployeeIDTemp')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><BR /><IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""26""></IFRAME><INPUT TYPE=""HIDDEN"" NAME=""EmployeeIDTemp"" ID=""EmployeeIDTempTxt"" "
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(B_SHOW_ID_FIELD_CATALOG) = True
			'aCatalogComponent(N_ACTIVE_CATALOG) = -2
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
                If Len(oRequest("Change").Item) = 0 Then
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "document.CatalogFrm.OwnerIDTemp.value = document.CatalogFrm.PARENTID.value;" & vbNewLine
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "document.CatalogFrm.EmployeeIDTemp.value = document.CatalogFrm.EMPLOYEEID.value;" & vbNewLine
                End If
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "function CheckOwnersValidation() {" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "if (document.CatalogFrm.OwnerIDTemp.value == '') {" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "alert('Favor de validar la existencia de la jefatura a la que pertenece este registro.');" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "document.CatalogFrm.OWNERID.focus();" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "return false;" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "}" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "if (document.CatalogFrm.EmployeeIDTemp.value == '') {" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "alert('Favor de validar la existencia del empleado encargadro de este registro.');" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "document.CatalogFrm.EMPLOYEEID.focus();" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "return false;" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "}" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "return true;" & vbNewLine
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "} // End of CheckOwnersValidation" & vbNewLine
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "//--></SCRIPT>" & vbNewLine
			aCatalogComponent(S_ADDITIONAL_FORM_SCRIPT_CATALOG) = "return CheckOwnersValidation();"
		Case "PaperworkSenders"
			aCatalogComponent(S_NAME_CATALOG) = "Procedencias"
			aCatalogComponent(S_ORDER_CATALOG) = "SenderName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Puesto,Empleado,Área"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "SenderID,SenderName,PositionName,EmployeeName,AreaID"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,5,6"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,255,255,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,,,,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,,-1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,,-1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("0,1,2,3,4", ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞAreas;,;AreaID;,;AreaCode, AreaName;,;(AreaID>-1) And (ParentID=-1) And (EndDate=30000000) And (Active=1);,;AreaPath;,;;,;Ninguna;;;-1"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ />&nbsp;<SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""document.CatalogFrm.PositionName.value=this.value""><OPTION VALUE=""""> - Catálogo de puestos - </OPTION>" & GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionName", "PositionShortName, PositionName", "(PositionID>-1) And (EndDate=30000000) And (Active=1)", "PositionShortName", -1, "Ninguno;;;-1", sErrorDescription) & "</SELECT><INPUT TYPE=""HIDDEN"" VALUE=""""  ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = "<FONT FACE=""Arial"" SIZE=""2"">Consultar número de empleado: </FONT><INPUT TYPE=""TEXT"" NAME=""EmployeeIDTemp"" ID=""EmployeeIDTempTxt"" SIZE=""6"" CLASS=""TextFields"" /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeIDTemp.value, 'EmployeesInfo', 'SearchEmployeeNumberIFrame', 'CatalogFrm.EmployeeIDTemp')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><BR /><IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""240""></IFRAME>"
			aCatalogComponent(N_ACTIVE_CATALOG) = -2
			aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = False
		Case "PaperworkTypes"
			aCatalogComponent(S_NAME_CATALOG) = "Tipos de trámite"
			aCatalogComponent(S_ORDER_CATALOG) = "PaperworkTypeName, PaperworkTypeID"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "PaperworkTypeID,PaperworkTypeName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,100,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
		Case "PaymentCenters"
			aCatalogComponent(S_NAME_CATALOG) = "Centros de pago"
			aCatalogComponent(S_ORDER_CATALOG) = "PaymentCenterShortName, PaymentCenters.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Dirección,Estado,Descripción,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "PaymentCenterID,PaymentCenterShortName,PaymentCenterName,Address,Description,ZoneID,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,0,0,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,5,5,6,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,255,255,255,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,,,0," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,,,9," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞZones;,;ZoneID;,;ZoneCode, ZoneName;,;(ZoneID>-1) And (ParentID=-1) And (EndDate=30000000) And (Active=1);,;ZoneCode;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,6,7", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 6
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 7
		Case "PaymentsMessages"
			aCatalogComponent(S_NAME_CATALOG) = "Asignación de folios"
			aCatalogComponent(S_ORDER_CATALOG) = "RecordID"
			aCatalogComponent(N_NAME_CATALOG) = 0
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID;;;Nómina;;;Empleado;;;Empresa;;;Unidad administrativa;;;Entidad;;;&nbsp;<INPUT TYPE=""RADIO"" NAME=""ZoneTypeID"" ID=""ZoneTypeIDRd"" onClick=""UnselectAllItemsFromList(document.CatalogFrm.ZoneIDs); SelectItemByValue('9', false, document.CatalogFrm.ZoneIDs)"" />Local&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""ZoneTypeID"" ID=""ZoneTypeIDRd"" onClick=""SelectAllItemsFromList(document.CatalogFrm.ZoneIDs); UnSelectItemByValue('9', false, document.CatalogFrm.ZoneIDs)"" />Foráneos<BR /></FONT></TD></TR><TR><TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de tabulador;;;Puesto;;;Banco;;;Tipo de pago;;;Especial;;;Mensaje"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ";;;")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "RecordID,PayrollID,EmployeeID,CompanyID,AreaIDs,ZoneIDs,EmployeeTypeID,PositionID,BankID,ConceptID,bSpecial,Comments"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) = 0 Then
				aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,0,0,0,0,0,0,0,1,0,1"
			Else
				aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,0,0,1,0,0,0,0,1,0,1"
			End If
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) = 0 Then
				aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,4,6,8,8,6,6,6,4,11,5"
			Else
				aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,4,6,6,8,6,6,6,4,11,5"
			End If
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,6,0,0,0,0,0,0,2,0,2000"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,15,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,0,999999,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1," & oRequest("PayrollID").Item & ",,-1,-1,-1,-1,-1,-1,-1,0,"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1," & oRequest("PayrollID").Item & ",,-1,-1,-1,-1,-1,-1,-1,0,"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) = 0 Then
				aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞCompanies;,;CompanyID;,;CompanyShortName, CompanyName;,;(ParentID>-1) And (EndDate=30000000) And (Active=1);,;CompanyShortName;,;;,;Ninguna;;;-1ÞÞÞAreas;,;AreaID;,;AreaCode, AreaName;,;(AreaID>-1) And (ParentID=-1) And (EndDate=30000000) And (Active=1);,;AreaCode;,;;,;Ninguna;;;-1ÞÞÞStates;,;StateID;,;StateName;,;(StateID>-1) And (Active=1);,;StateName;,;;,;Ninguna;;;-1ÞÞÞEmployeeTypes;,;EmployeeTypeID;,;EmployeeTypeShortName, EmployeeTypeName;,;(EmployeeTypeID>-1) And (EmployeeTypeID<=6) And (EndDate=30000000) And (Active=1);,;EmployeeTypeShortName;,;;,;Ninguno;;;-1ÞÞÞPositions;,;PositionID;,;PositionShortName, PositionName;,;(PositionID>-1) And (EndDate=30000000) And (Active=1);,;PositionShortName;,;;,;Ninguno;;;-1ÞÞÞBanks;,;BankID;,;BankName;,;(BankID>-1) And (Active=1);,;BankName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞ"
			Else
				aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞCompanies;,;CompanyID;,;CompanyShortName, CompanyName;,;(ParentID>-1) And (EndDate=30000000) And (Active=1);,;CompanyShortName;,;;,;Ninguna;;;-1ÞÞÞAreas;,;AreaID;,;AreaCode, AreaName;,;(AreaID>-1) And (AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) And (ParentID=-1) And (EndDate=30000000) And (Active=1);,;AreaCode;,;;,;Ninguna;;;-1ÞÞÞStates;,;StateID;,;StateName;,;(StateID>-1) And (Active=1);,;StateName;,;;,;Ninguna;;;-1ÞÞÞEmployeeTypes;,;EmployeeTypeID;,;EmployeeTypeShortName, EmployeeTypeName;,;(EmployeeTypeID>-1) And (EmployeeTypeID<=6) And (EndDate=30000000) And (Active=1);,;EmployeeTypeShortName;,;;,;Ninguno;;;-1ÞÞÞPositions;,;PositionID;,;PositionShortName, PositionName;,;(PositionID>-1) And (EndDate=30000000) And (Active=1);,;PositionShortName;,;;,;Ninguno;;;-1ÞÞÞBanks;,;BankID;,;BankName;,;(BankID>-1) And (Active=1);,;BankName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞ"
			End If
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ /><A HREF=""javascript: SearchRecord(document.CatalogFrm.EmployeeID.value, 'EmployeeNumber', 'SearchAccountsCatalogsIFrame', 'CatalogFrm.EmployeeID');""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><IFRAME SRC=""SearchRecord.asp"" NAME=""SearchAccountsCatalogsIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""20""></IFRAME><INPUT TYPE=""HIDDEN"" NAME=""Step"" ID=""Step"" VALUE=""2"" ÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""SearchRecord(this.value, 'PositionsForEmployeeType', 'SearchPositionsIFrame', 'CatalogFrm.PositionID');"" ÞÞÞÞÞÞÞÞÞ STYLE=""width: 0px"" /><SELECT NAME="""" ID=""Cmb"" SIZE=""1"" onChange=""document.CatalogFrm.ConceptID.value=this.value;""><OPTION VALUE=""-1"">Todos</OPTION><OPTION VALUE=""0"">Cheque</OPTION><OPTION VALUE=""1"">Depósito</OPTION><OPTION VALUE=""2"">Pensión alimenticia</OPTION><OPTION VALUE=""3"">Honorarios</OPTION><OPTION VALUE=""4"">Acreedores</OPTION></SELECT><INPUT TYPE=""HIDDEN"" ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4,5,6,7,8,11", ",")
			aCatalogComponent(S_ADDITIONAL_FORM_SCRIPT_CATALOG) = "if (oForm.EmployeeID.value == '') {oForm.EmployeeID.value='0';}"
			aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = False
			aCatalogComponent(S_CANCEL_BUTTON_ACTION_CATALOG) = "window.location.href='Payments.asp?Action=PrintPayments&PayrollID=" & lPaymentID & "&Step=1';"
			aCatalogComponent(S_EXTRA_BUTTON_CATALOG) = "<INPUT TYPE=""BUTTON"" VALUE=""Continuar con la Impresión"" CLASS=""Buttons"" onClick=""window.location.href='Payments.asp?Action=PrintPayments&PayrollID=" & lPayrollID & "&Step=3';"" />" & vbNewLine
		Case "PaymentsRecords"
			aCatalogComponent(S_NAME_CATALOG) = "Asignación de folios"
			aCatalogComponent(S_ORDER_CATALOG) = "PayrollDate, CompanyName, EmployeeTypeName, BankName, AccountNumber"
			aCatalogComponent(N_NAME_CATALOG) = 0
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID;;;Nómina;;;Empresa;;;Centros de pago;;;Entidad (Centros de pago);;;&nbsp;<INPUT TYPE=""RADIO"" NAME=""ZoneTypeID"" ID=""ZoneTypeIDRd"" onClick=""UnselectAllItemsFromList(document.CatalogFrm.ZoneIDs); SelectItemByValue('9', false, document.CatalogFrm.ZoneIDs)"" />Local&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""ZoneTypeID"" ID=""ZoneTypeIDRd"" onClick=""SelectAllItemsFromList(document.CatalogFrm.ZoneIDs); UnSelectItemByValue('9', false, document.CatalogFrm.ZoneIDs)"" />Foráneos<BR /></FONT></TD></TR><TR><TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de tabulador del empleado;;;&nbsp;<INPUT TYPE=""RADIO"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDRd"" VALUE=""1"" onClick=""UnselectAllItemsFromList(document.CatalogFrm.EmployeeTypeIDs); SelectItemByValue('1', false, document.CatalogFrm.EmployeeTypeIDs)"" />Funcionarios&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=""RADIO"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDRd"" VALUE=""-1"" onClick=""SelectAllItemsFromList(document.CatalogFrm.EmployeeTypeIDs); UnSelectItemByValue('1', false, document.CatalogFrm.EmployeeTypeIDs)"" />Operativos<BR /></FONT></TD></TR><TR><TD><FONT FACE=""Arial"" SIZE=""2"">Banco;;;Cuenta bancaria;;;Numeración inicial para reposición;;;Numeración final para reposición;;;Inicio de la numeración;;;Término de la numeración;;;Primer ID;;; Último ID;;;Pagos de;;;Impreso;;;Fecha de generación;;;Usuario;;;Empleado"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ";;;")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "RecordID,PayrollID,CompanyIDs,AreaIDs,ZoneIDs,EmployeeTypeIDs,BankID,AccountID,ReexpeditionNumber,EndNumber,FirstNumber,LastNumber,FirstPaymentID,LastPaymentID,ConceptID,bPrinted,ModifyDate,UserID,EmployeeID"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,6,8,8,8,8,6,11,11,11,4,11,11,11,4,11,11,11,5"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,0,0,0,0,0,0,10,10,10,10,-1,-1,0,0,0,0,6"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0,5,0,5,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,1,1,1,1,0,0,1,1,1,1,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,-1,,0,-1,,-1,-1,-1,-1,1,1,-1,-1,0,0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ","
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,-1,,,-1,,-1,-1,-1,-1,1,1,-1,-1,0,0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ","
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) = 0 Then
				aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞPayrolls;,;PayrollID;,;PayrollDate, PayrollName;,;(PayrollTypeID<>0) And (IsClosed=1);,;PayrollID Desc;,;;,;Ninguna;;;-1ÞÞÞCompanies;,;CompanyID;,;CompanyShortName, CompanyName;,;(ParentID>-1) And (EndDate=30000000) And (Active=1);,;CompanyShortName;,;;,;Ninguna;;;-1ÞÞÞAreas;,;AreaID;,;AreaCode, AreaName;,;(AreaID>-1) And (ParentID>-1) And (EndDate=30000000) And (Active=1);,;AreaCode;,;0;,;Ninguna;;;-1ÞÞÞStates;,;StateID;,;StateName;,;(StateID>-1) And (Active=1);,;StateName;,;;,;Ninguna;;;-1ÞÞÞEmployeeTypes;,;EmployeeTypeID;,;EmployeeTypeShortName, EmployeeTypeName;,;((EmployeeTypeID>=0) And (EmployeeTypeID<=7) Or (EmployeeTypeID IN (12))) And (EndDate=30000000) And (Active=1);,;EmployeeTypeShortName;,;;,;Ninguno;;;-1ÞÞÞBanks;,;BankID;,;BankName;,;(BankID>0) And (Active=1);,;BankName;,;;,;Ninguno;;;-1ÞÞÞBankAccounts, Banks;,;AccountID;,;BankName, AccountNumber;,;(BankAccounts.BankID=Banks.BankID) And (AccountID>-1) And (Banks.BankID>-1) And (BankAccounts.EmployeeID<0) And (BankAccounts.Active=1) And (Banks.Active=1);,;BankName, AccountNumber;,;;,;Ninguna;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			Else
				aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞPayrolls;,;PayrollID;,;PayrollDate, PayrollName;,;(PayrollTypeID<>0) And (IsClosed=1);,;PayrollID Desc;,;;,;Ninguna;;;-1ÞÞÞCompanies;,;CompanyID;,;CompanyShortName, CompanyName;,;(ParentID>-1) And (EndDate=30000000) And (Active=1);,;CompanyShortName;,;;,;Ninguna;;;-1ÞÞÞAreas;,;AreaID;,;AreaCode, AreaName;,;(AreaID>-1) And (AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) And (ParentID>-1) And (EndDate=30000000) And (Active=1);,;AreaCode;,;0;,;Ninguna;;;-1ÞÞÞStates;,;StateID;,;StateName;,;(StateID>-1) And (Active=1);,;StateName;,;;,;Ninguna;;;-1ÞÞÞEmployeeTypes;,;EmployeeTypeID;,;EmployeeTypeShortName, EmployeeTypeName;,;((EmployeeTypeID>=0) And (EmployeeTypeID<=7) Or (EmployeeTypeID IN (12))) And (EndDate=30000000) And (Active=1);,;EmployeeTypeShortName;,;;,;Ninguno;;;-1ÞÞÞBanks;,;BankID;,;BankName;,;(BankID>0) And (Active=1);,;BankName;,;;,;Ninguno;;;-1ÞÞÞBankAccounts, Banks;,;AccountID;,;BankName, AccountNumber;,;(BankAccounts.BankID=Banks.BankID) And (AccountID>-1) And (Banks.BankID>-1) And (BankAccounts.EmployeeID<0) And (BankAccounts.Active=1) And (Banks.Active=1);,;BankName, AccountNumber;,;;,;Ninguna;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			End If
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""ShowPaymentsRecordsFields(this.value);"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ STYLE=""width: 0px"" /><SELECT NAME="""" ID=""Cmb"" SIZE=""1"" CLASS=""Lists"" onChange=""document.CatalogFrm.ConceptID.value=this.value;""><OPTION VALUE=""0"">Cheque</OPTION><OPTION VALUE=""1"">Depósito</OPTION><OPTION VALUE=""2"">Pensión alimenticia</OPTION><OPTION VALUE=""3"">Honorarios</OPTION><OPTION VALUE=""4"">Acreedores</OPTION></SELECT><INPUT TYPE=""HIDDEN"" ÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4,5,6,7,10,11,18,8,9", ",")
			aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = False
			aCatalogComponent(S_CANCEL_BUTTON_ACTION_CATALOG) = "window.location.href='Main_ISSSTE.asp?SectionID=47';"
		Case "PaymentsRecords2"
			aCatalogComponent(S_NAME_CATALOG) = "Asignación de folios"
			aCatalogComponent(S_ORDER_CATALOG) = "PayrollDate, CompanyName, EmployeeTypeName, BankName, AccountNumber"
			aCatalogComponent(N_NAME_CATALOG) = 0
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nómina,No. del empleado,Banco,Cuenta,No. del cheque a reexpedir,No. del nuevo pago,Reexpedición en,Impreso,Nueva fecha de pago,ID,Fecha de generación,Usuario"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "RecordID,PayrollID,EmployeeID,BankID,AccountID,CheckNumber,ReplacementNumber,ConceptID,bPrinted,PaymentDate,PaymentID,ModifyDate,UserID"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,0,1,1,1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,6,4,6,11,5,5,4,11,1,11,11,11"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,6,0,0,100,100,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,1,0,0,0,0,0,0," & N_PAYROLL_START_YEAR & ",0,,00"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,0,999999,0,0,0,0,0,0," & Year(Date()) & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,-1,,-1,-1,,,0,0," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",-1," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN)
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,-1,,-1,-1,,,0,0," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",-1," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN)
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞPayrolls;,;PayrollID;,;PayrollDate, PayrollName;,;(PayrollTypeID>0) And (IsClosed=1);,;PayrollID Desc;,;;,;Ninguna;;;-1ÞÞÞÞÞÞBanks;,;BankID;,;BankName;,;(Active=1);,;BankName;,;;,;Ninguno;;;-1ÞÞÞBankAccounts, Banks;,;AccountID;,;BankName, AccountNumber;,;(BankAccounts.BankID=Banks.BankID) And (AccountID>-1) And (Banks.BankID>-1) And (BankAccounts.EmployeeID<0) And (BankAccounts.Active=1) And (Banks.Active=1);,;BankName, AccountNumber;,;;,;Ninguna;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ /><A HREF=""javascript: SearchForPayment();""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado y obtener el número de cheque"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><IFRAME SRC=""SearchRecord.asp"" NAME=""SearchAccountsCatalogsIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""20""></IFRAME><INPUT TYPE=""HIDDEN"" NAME=""CheckEmployee"" ID=""CheckEmployee"" VALUE=""" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(3) & """ ÞÞÞ onChange=""ShowPaymentsRecordsFields(this.value);"" ÞÞÞÞÞÞ onFocus=""this.form.ReplacementNumber.focus()"" ÞÞÞÞÞÞSTYLE=""width: 0px"" /><SELECT NAME="""" ID=""Cmb"" SIZE=""1"" CLASS=""Lists"" onChange=""document.CatalogFrm.ConceptID.value=this.value;""><OPTION VALUE=""0"">Cheque</OPTION><OPTION VALUE=""1"">Depósito</OPTION></SELECT><INPUT TYPE=""HIDDEN"" ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4,5", ",")
			aCatalogComponent(S_ADDITIONAL_FORM_SCRIPT_CATALOG) = "return CheckEmployeeValidation();"
			aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = False
			aCatalogComponent(S_CANCEL_BUTTON_ACTION_CATALOG) = "window.location.href='Main_ISSSTE.asp?SectionID=47';"
		Case "PayrollResume"
			aCatalogComponent(S_TABLE_NAME_CATALOG) = "DM_HIST_NOMSAR"
			aCatalogComponent(S_NAME_CATALOG) = "Resumen de nóminas"
			aCatalogComponent(S_ORDER_CATALOG) = "PaymentDate , CompanyID, BankID Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "Sociedad,Empresa,Periodo,CLC,Banco,Fecha de pago,Tipo de empleado,Ingresos,Deducciones,Líquido,Cpt 01,Cpt 04,Cpt 05,Cpt 06,Cpt 07,Cpt 08,Cpt 11,Cpt 44,Cpt b2,Cpt 7s,Comentarios"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "SocietyID,CompanyID,periodID,CLC,BankID,PaymentDate,EmployeeType,Income,Deductions,NetIncome,Cpt_01,Cpt_04,Cpt_05,Cpt_06,Cpt_07,Cpt_08,Cpt_11,Cpt_44,Cpt_b2,Cpt_7s,Comments"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,4,4,5,4,1,4,4,4,4,4,4,4,4,4,4,4,4,4,4,5"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,0,100,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,255"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "1000000000,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 3
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 4
		Case "Periods"
			aCatalogComponent(S_NAME_CATALOG) = "Periodicidad"
			aCatalogComponent(S_ORDER_CATALOG) = "PeriodID"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Días,Fecha,Especial,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "PeriodID,PeriodName,PeriodDays,PeriodDate,bSpecial,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,4,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,100,6,255,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,1,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,100000000,,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1,,0,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1,,0,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
		Case "Positions"
			aCatalogComponent(S_NAME_CATALOG) = "Puesto"
			aCatalogComponent(S_ORDER_CATALOG) = "PositionShortName, PositionName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Denominación,Nombre largo,Descripción,Fecha de inicio,Fecha de término,Tipo de tabulador,Tipo de puesto,Compañía,Clasificación,Grupo-grado-nivel,Integración,Nivel,Rama,Subrama,Jerarquía,Puesto genérico,Horas laboradas,¿Es estratégico?,¿Está nominado?,Estatus,Bloqueado,Activo,Zona Económica, Descripción"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "PositionID,PositionShortName,PositionName,PositionLongName,PositionDescription,StartDate,EndDate,EmployeeTypeID,PositionTypeID,CompanyID,ClassificationID,GroupGradeLevelID,IntegrationID,LevelID,BranchID,SubBranchID,HierarchyID,GenericPositionID,WorkingHours,Strategic,Nomination,StatusID,Depreciated,Active, EconomicZoneID,Comments"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,0,1,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,5,11,1,1,6,6,6,4,6,4,6,6,6,4,6,2,11,11,11,0,11"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "4,10,255,2000,2000,0,0,0,0,0,2,0,2,3,0,0,3,0,3,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0,0,0,5,0,5,0,0,0,0,0,15,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0," & N_FORM_START_YEAR & "," & N_FORM_START_YEAR & ",0,0,0,-1,-1,-1,0,-1,-1,0,-1,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,,," & Year(Date()) & "," & Year(Date()) & ",0,0,0,0,0,0,0,0,0,0,0,24,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,,,0,0,-1,-1,-1,0,-1,0,0,-1,-1,0,-1,0,1,1,-1,0,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,,,0,0,-1,-1,-1,-1,-1,-1,0,-1,-1,0,-1,0,1,1,-1,0,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞEmployeeTypes;,;EmployeeTypeID;,;EmployeeTypeShortName, EmployeeTypeName;,;(EndDate=30000000) And (Active=1);,;EmployeeTypeShortName;,;;,;Ninguno;;;-1ÞÞÞPositionTypes;,;PositionTypeID;,;PositionTypeShortName, PositionTypeName;,;(EndDate=30000000) And (Active=1);,;PositionTypeShortName;,;;,;Ninguno;;;-1ÞÞÞCompanies;,;CompanyID;,;CompanyShortName, CompanyName;,;(ParentID=0) And (EndDate=30000000) And (Active=1);,;CompanyShortName;,;;,;Ninguna;;;-1ÞÞÞÞÞÞGroupGradeLevels;,;GroupGradeLevelID;,;GroupGradeLevelShortName, GroupGradeLevelName;,;(EndDate=30000000) And (Active=1);,;GroupGradeLevelShortName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞLevels;,;LevelID;,;LevelName;,;(Active=1);,;LevelName;,;;,;Ninguno;;;-1ÞÞÞBranches;,;BranchID;,;BranchShortName, BranchName;,;(Active=1);,;BranchShortName;,;;,;Ninguna;;;-1ÞÞÞSubBranches;,;SubBranchID;,;SubBranchShortName, SubBranchName;,;(Active=1);,;SubBranchShortName;,;;,;Ninguna;;;-1ÞÞÞÞÞÞGenericPositions;,;GenericPositionID;,;GenericPositionName;,;(Active=1);,;GenericPositionName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞStatusPositions;,;StatusID;,;StatusName;,;(Active=1);,;StatusName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,7,8,9,10,11,12,13,14,15,16,17,18,5,6", ",")
			'aCatalogComponent(S_ADDITIONAL_FORM_SCRIPT_CATALOG) = "return GetCatalogsIDs();"
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 5
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 6
		Case "PositionsAreasLKP"
			aCatalogComponent(S_NAME_CATALOG) = "Plantilla original"
			aCatalogComponent(S_ORDER_CATALOG) = "AreaCode, PositionShortName, PositionsAreasLKP.StartDate"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "Puesto,Área,Fecha de inicio,Fecha de término,Plazas originales,Plazas modificadas"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "PositionID,AreaID,StartDate,EndDate,OriginalJobsInArea,ModifiedJobsInArea"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "6,6,1,1,4,4"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,0,0,5,5"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0," & Year(Date()) & "," & Year(Date()) & ",0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "0,0," & Year(Date()) + 1 & "," & Year(Date()) + 1 & ",0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,-1," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0,0,0"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,-1," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0,0"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "Positions;,;PositionID;,;PositionShortName, PositionName;,;(Positions.PositionID>0) And (EndDate=30000000) And (Active=1);,;PositionShortName;,;;,;Ninguno;;;-1ÞÞÞAreas;,;AreaID;,;AreaCode, AreaName;,;(Areas.EndDate=30000000) And (Areas.ParentID<>-1);,;AreaCode;,;;,;Ninguno;;;-1ÞÞÞServices;,;ServiceID;,;ServiceShortName, ServiceName;,;(Services.EndDate=30000000) And (Services.Active=1);,;ServiceShortName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
			aCatalogComponent(N_ACTIVE_CATALOG) = -2
			aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = False
			aCatalogComponent(B_SHOW_ID_FIELD_CATALOG) = True
			aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (Positions.EndDate=30000000) And (Areas.EndDate=30000000)"
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 2
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 3
		Case "PositionsHierarchy"
			aCatalogComponent(B_SHOW_ID_FIELD_CATALOG) = True
			aCatalogComponent(S_NAME_CATALOG) = "Jerarquía de puestos"
			aCatalogComponent(S_ORDER_CATALOG) = "JobID, AreaCode, PositionShortName, GGNShortName"
			aCatalogComponent(N_NAME_CATALOG) = 0
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "Plaza,Centro de trabajo,Puesto,Grupo-grado-nivel,Plaza del puesto superior,Centro de trabajo del puesto superior,Puesto del puesto superior,Grupo-grado-nivel del puesto superior,Fecha de inicio,Fecha de término,Fecha de modificación,Usuario"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "JobID,AreaCode,PositionShortName,GGNShortName,ParentJobID,ParentAreaCode,ParentPositionShortName,ParentGGNShortName,StartDate,EndDate,ModifyDate,UserID"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1,1,1,1,0,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,5,4,5,5,5,1,1,11,11"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "6,5,7,3,6,5,7,3,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,15,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,1,0,0,0,2009,2009,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "999999,0,0,0,999999,0,0,0," & Year(Date()) + 1 & "," & Year(Date()) + 1 & ",0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,,,,,,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN)
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,,,,,,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN)
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("0,1,2,3,4,5,6,7,8,9", ",")
			aCatalogComponent(N_ACTIVE_CATALOG) = -2
'			aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = False
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 8
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 9
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = "<INPUT TYPE=""HIDDEN"" NAME=""AreaCodeOld"" ID=""AreaCodeOldHdn"" VALUE="""" />"
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<FONT FACE=""Arial"" SIZE=""2""><INPUT TYPE=""CHECKBOX"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDChk"" VALUE=""2"" onClick=""ShowParentFields(this.checked);"" /> Operativos<BR /><BR /></FONT>" & vbNewLine
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "var sJobID = '';" & vbNewLine
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "var sPositionShortName = '';" & vbNewLine
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "var sGGNShortName = '';" & vbNewLine

				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "function ShowParentFields (bEmployeeType) {" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "var oForm = document.CatalogFrm;" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "if (oForm) {" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "if (bEmployeeType) {" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "HideDisplay(document.all['CatalogFrm_JobIDDiv']);" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "HideDisplay(document.all['CatalogFrm_PositionShortNameDiv']);" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "HideDisplay(document.all['CatalogFrm_GGNShortNameDiv']);" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "oForm.JobID.value = '0';" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "oForm.PositionShortName.value = '0';" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "oForm.GGNShortName.value = '0';" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "} else {" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "ShowDisplay(document.all['CatalogFrm_JobIDDiv']);" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "ShowDisplay(document.all['CatalogFrm_PositionShortNameDiv']);" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "ShowDisplay(document.all['CatalogFrm_GGNShortNameDiv']);" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "oForm.JobID.value = sJobID;" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "oForm.PositionShortName.value = sPositionShortName;" & vbNewLine
							aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "oForm.GGNShortName.value = sGGNShortName;" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "}" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "}" & vbNewLine
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "} // End of ShowParentFields" & vbNewLine
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "//--></SCRIPT>" & vbNewLine
			aCatalogComponent(S_URL_CATALOG) = "AreaCode=<FIELD_1 />&ParentJobID=<FIELD_4 />&ReadOnly=" & oRequest("ReadOnly").Item
		Case "PositionsSpecialJourneysLKP"
			aCatalogComponent(S_NAME_CATALOG) = "Puestos para guardias y suplecias"
			aCatalogComponent(S_ORDER_CATALOG) = "PositionShortName, ServiceShortName, CenterTypeShortName, PositionsSpecialJourneysLKP.StartDate"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Fecha de inicio,Fecha final,Puesto,Nivel,Jornada,Servicio,Tipo de C. de trabajo,Guardias,Suplencias,R. Quirúrgico,PROVAC"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "RecordID,StartDate,EndDate,PositionID,LevelID,WorkingHours,ServiceID,CenterTypeID,IsActive1,IsActive2,IsActive3,IsActive4"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,0,1,1,1,1,1,0,0,0,0"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,1,1,6,11,11,6,6,0,0,0,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,2000,2000,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000," & Year(Date()) + 1 & "," & Year(Date()) + 1 & ",-1,-1,-1,-1,-1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,0,0,-1,-1-1,-1,-1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,0,0,-1,-1,-1,-1,-1,1,1,1,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞPositions;,;PositionID;,;PositionShortName, PositionName;,;(EndDate=30000000) And (Active=1);,;PositionShortName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞServices;,;ServiceID;,;ServiceShortName, ServiceName;,;(EndDate=30000000) And (Active=1);,;ServiceShortName;,;;,;Ninguno;;;-1ÞÞÞCenterTypes;,;CenterTypeID;,;CenterTypeShortName, CenterTypeName;,;(EndDate=30000000) And (Active=1);,;CenterTypeShortName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), "ÞÞÞ")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), "ÞÞÞ")
			'If StrComp(GetASPFileName(""), "Export.asp", vbBinaryCompare) = 0 Then
				aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4,5,6,7,8,9,10,11", ",")
			'Else
			'	aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4,5,6,7", ",")
			'End If
			'If Len(oRequest("ShowAll").Item) = 0 Then aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (PositionsSpecialJourneysLKP.EndDate=30000000)"
			aCatalogComponent(N_ACTIVE_CATALOG) = -2
			aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = False
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "function CheckPositionsSpecialJourneysLKPValidation() {" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "if ((document.CatalogFrm.PositionID.value == '') || (document.CatalogFrm.PositionID.value == '-1')) {" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "alert('Favor de indicar el puesto para registro de guargias y suplencias.');" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "document.CatalogFrm.PositionID.focus();" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "return false;" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "}" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "if ((document.CatalogFrm.ServiceID.value == '') || (document.CatalogFrm.ServiceID.value == '-1')) {" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "alert('Favor de indicar el servicio del puesto para guardias y suplencias.');" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "document.CatalogFrm.ServiceID.focus();" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "return false;" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "}" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "if ((document.CatalogFrm.CenterTypeID.value == '') || (document.CatalogFrm.CenterTypeID.value == '-1')) {" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "alert('Favor de indicar el Tipo de centro de trabajo para guardias y suplencias.');" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "document.CatalogFrm.CenterTypeID.focus();" & vbNewLine
						aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "return false;" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "}" & vbNewLine
					aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "return true;" & vbNewLine
				aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "} // End of CheckPositionsSpecialJourneysLKPValidation" & vbNewLine
			aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & "//--></SCRIPT>" & vbNewLine
			aCatalogComponent(S_ADDITIONAL_FORM_SCRIPT_CATALOG) = "return CheckPositionsSpecialJourneysLKPValidation();"
		Case "PositionTypes"
			aCatalogComponent(S_NAME_CATALOG) = "Tipos de puesto"
			aCatalogComponent(S_ORDER_CATALOG) = "PositionTypeShortName, PositionTypes.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "PositionTypeID,PositionTypeShortName,PositionTypeName,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,255,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 3
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 4
		Case "ProfessionalRiskMatrix"
			aCatalogComponent(S_NAME_CATALOG) = "Administración de Matriz de Riesgos"
			aCatalogComponent(S_ORDER_CATALOG) = "BranchID, PaymentCenterID, PositionID, ServiceID"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "Rama o Grupo,Centro de Trabajo,Puesto,Servicio,Nivel de Riesgo,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "BranchID,PaymentCenterID,PositionID,ServiceID,RiskLevel,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "5,4,4,4,4,11,11,11,4,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4,5,6", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 3
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 4
		Case "Projects", "TACO_Projects"
			If InStr(1, aCatalogComponent(S_TABLE_NAME_CATALOG), TACO_PREFIX, vbBinaryCompare) = 0 Then aCatalogComponent(S_TABLE_NAME_CATALOG) = TACO_PREFIX & aCatalogComponent(S_TABLE_NAME_CATALOG)
			aCatalogComponent(S_NAME_CATALOG) = "Proyectos"
			aCatalogComponent(S_ORDER_CATALOG) = "ProjectName"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Número,Sección,Descripción,Objetivo,Sección que podrá trabajar sobre el proceso,Modo sencillo,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "ProjectID,ProjectName,ProjectNumber,ProjectOwner,ProjectDescription,ProjectObjective,ProjectFile,EasyMode,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,5,5,5,8,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,255,30,100,4000,4000,7,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			If Request.Cookies("SIAP_SectionID") > -1 Then Call GetNameFromTable(oADODBConnection, "UserProfiles", iGlobalSectionID, "", "", sNames, sErrorDescription)
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,," & sNames & ",,," & iGlobalSectionID & ",1,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,," & sNames & ",," & iGlobalSectionID & ",1,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞUserProfiles;,;ProfileID;,;ProfileName;,;(ProfileID>0);,;ProfileName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(S_URL_PARAMETERS_CATALOG) = "TaCo.asp?ProjectID=<FIELD_0 />&Action=Tasks"
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("2,1,3", ",")
		Case "Reasons"
			aCatalogComponent(S_NAME_CATALOG) = "Tipos de movimiento"
			aCatalogComponent(S_ORDER_CATALOG) = "ReasonShortName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Clave,Descripción,Clasificación,Estatus de la plaza para efectuar el movimiento,Estatus debe estar la plaza a ocupar,Estatus de la plaza después del movimiento,Estatus del empleado para realizar el movimiento,Estatus del empleado para validación o autorización,Estatus empleado después de movimiento,¿Activar al empleado después del movimiento?,Requerimiento para movimiento,Secciones a mostrar en el formulario del empleado"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "ReasonID,ReasonShortName,ReasonName,ReasonTypeID,StatusJob1,StatusJob2,StatusJob3,StatusEmployeesIDs,StatusEmployeesIDs1,StatusEmployeeID,ActiveEmployeeID,ReasonRequirementIDs,Sections"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1,1,1,1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,6,8,8,6,8,8,6,0,8,8"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,255,3,100,100,4,255,255,4,4,255,255"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,,,,,,,,,,"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,,,,,,,,,,"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,,,,,,,,,,"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞReasonTypes;,;ReasonTypeID;,;ReasonTypeName;,;;,;ReasonTypeName;,;;,;Ninguna;;;-1ÞÞÞStatusJobs;,;StatusID;,;StatusName;,;;,;StatusName;,;;,;Ninguno;;;-1ÞÞÞStatusJobs;,;StatusID;,;StatusName;,;;,;StatusName;,;;,;Ninguno;;;-1ÞÞÞStatusJobs;,;StatusID;,;StatusName;,;;,;StatusName;,;;,;Ninguno;;;-1ÞÞÞStatusEmployees;,;StatusID;,;StatusName;,;;,;StatusName;,;;,;Ninguno;;;-1ÞÞÞStatusEmployees;,;StatusID;,;StatusName;,;;,;StatusName;,;;,;Ninguno;;;-1ÞÞÞStatusEmployees;,;StatusID;,;StatusName;,;;,;StatusName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞEmployeesRequirements;,;EmployeeRequirementID;,;EmployeeRequirementName;,;;,;EmployeeRequirementName;,;;,;Ninguna;;;-1ÞÞÞSections;,;SectionID;,;SectionName;,;;,;SectionName;,;;,;Ninguna;;;-1"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3", ",")
			aCatalogComponent(N_ACTIVE_CATALOG) = -2
		Case "ReasonTypes"
			aCatalogComponent(S_NAME_CATALOG) = "Clasificación de tipos de movimiento"
			aCatalogComponent(S_ORDER_CATALOG) = "ReasonTypeName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "ReasonTypeID,ReasonTypeName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,255,1"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,15"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,9"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
		Case "Requirements"
			aCatalogComponent(S_NAME_CATALOG) = "Perfiles académicos"
			aCatalogComponent(S_ORDER_CATALOG) = "RequirementID"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "RequirementID,RequirementName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,100,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
		Case "RiskLevels"
			aCatalogComponent(S_NAME_CATALOG) = "Riesgos profesionales"
			aCatalogComponent(S_ORDER_CATALOG) = "RiskLevelID"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "RiskLevelID,RiskLevelName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,100,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
		Case "SADE_NewCourse"
			aCatalogComponent(S_NAME_CATALOG) = "Registro de detección de necesidades"
			aCatalogComponent(S_ORDER_CATALOG) = "EmployeeID, RecordID"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Empleado,Escolaridad,Nombre de la institución,Nombre de la carrera,<SPAN ID=""SemestersText1Div"">Semestres cursados</SPAN><SPAN ID=""SemestersText2Div"">Grados cursados</SPAN>,Duración de la carrera en años,Titulado,Describa brevemente las tres principales funciones que desempeña en su área de trabajo,<BR /><B>Anote los cursos que usted considere necesarios para desempeñar adecuadamente sus funciones laborales.</B><BR />1. Capacitación vinculada a servicios de salud,2. Capacitación en apoyo a los procesos jurídicos financieros y técnico administrativo,3. Capacitación en tecnología de información,4. Capacitación pedagógica,5. Capacitación sobre asuntos técnico-operativo,6. Capacitación para la superación personal,7. Otros cursos,Al recibir estos cursos ¿cuáles son las deficiencias o problemas que usted resolverá en el desempeño de sus funciones?"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "RecordID,EmployeeID,SchoolarshipID,UniversityName,BSName,Semesters,BSYears,StatusID,EmployeeActivities,CourseNames_1,CourseNames_2,CourseNames_3,CourseNames_4,CourseNames_5,CourseNames_6,CourseNames_7,CoursesResults"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,0,0,1,1,1,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,6,5,5,4,4,6,5,8,8,8,8,8,8,8,5"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,0,0,255,255,2,2,0,10,10,10,10,10,10,10,10,2000"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,15,15,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0,1,1,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,0,0,0,0,20,10,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,-1,-1,,,1,1,1,,,,,,,,,"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,-1,-1,,,1,1,1,,,,,,,,,"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞSchoolarships;,;SchoolarshipID;,;SchoolarshipName;,;(SchoolarshipID>-1) And (Active=1);,;SchoolarshipID;,;;,;Ninguna;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞStatusBachelors;,;StatusID;,;StatusName;,;(StatusID>-1) And (Active=1);,;StatusName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞSADE_Perfiles;,;ID_Perfil;,;Nombre_Perfil;,;(ID_Padre=1);,;Nombre_Perfil;,;;,;ÞÞÞSADE_Perfiles;,;ID_Perfil;,;Nombre_Perfil;,;(ID_Padre=2);,;Nombre_Perfil;,;;,;ÞÞÞSADE_Perfiles;,;ID_Perfil;,;Nombre_Perfil;,;(ID_Padre=3);,;Nombre_Perfil;,;;,;ÞÞÞSADE_Perfiles;,;ID_Perfil;,;Nombre_Perfil;,;(ID_Padre=4);,;Nombre_Perfil;,;;,;ÞÞÞSADE_Perfiles;,;ID_Perfil;,;Nombre_Perfil;,;(ID_Padre=5);,;Nombre_Perfil;,;;,;ÞÞÞSADE_Perfiles;,;ID_Perfil;,;Nombre_Perfil;,;(ID_Padre=6);,;Nombre_Perfil;,;;,;ÞÞÞSADE_Perfiles;,;ID_Perfil;,;Nombre_Perfil;,;(ID_Padre=7);,;Nombre_Perfil;,;;,;ÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ onChange=""ShowNewCourseFields(this.value);""ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"" ÞÞÞ onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"" ÞÞÞ onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"" ÞÞÞ onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"" ÞÞÞ onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"" ÞÞÞ onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"" ÞÞÞ onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"" ÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3", ",")
			aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = False
		Case "Schoolarships"
			aCatalogComponent(S_NAME_CATALOG) = "Escolaridad"
			aCatalogComponent(S_ORDER_CATALOG) = "SchoolarshipName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "SchoolarshipID,SchoolarshipName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,255,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
		Case "Services"
			aCatalogComponent(S_NAME_CATALOG) = "Servicios"
			aCatalogComponent(S_ORDER_CATALOG) = "ServiceShortName, Services.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Código padre,Tipo de centro de trabajo,Fondo,Programa presupuestario,Función,Subfunción activa,Subfunción específica,Programa,Actividad institucional,Actividad presupuestaria,Proceso,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "ServiceID,ServiceShortName,ServiceName,ParentShortName,CenterTypeID,FundID,ProgramDutyID,DutyID,ActiveDutyID,SpecificDutyID,ProgramID,ActivityID1,ActivityID2,ProcessID,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,5,6,6,6,6,6,6,6,6,6,6,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,255,5,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0,0,0,0,0,0,0,0,0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",30000000," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",30000000," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞCenterTypes;,;CenterTypeID;,;CenterTypeShortName, CenterTypeName;,;(EndDate=30000000) And (Active=1);,;CenterTypeShortName, CenterTypeName;,;;,;Ninguno;;;-1ÞÞÞBudgetsFunds;,;FundID;,;FundShortName, FundName;,;(Active=1);,;FundShortName, FundName;,;;,;Ninguna;;;-1ÞÞÞBudgetsProgramDuties;,;ProgramDutyID;,;ProgramDutyShortName, ProgramDutyName;,;(Active=1);,;ProgramDutyShortName, ProgramDutyName;,;;,;Ninguna;;;-1ÞÞÞBudgetsDuties;,;DutyID;,;DutyShortName, DutyName;,;(Active=1);,;DutyShortName, DutyName;,;;,;Ninguna;;;-1ÞÞÞBudgetsActiveDuties;,;ActiveDutyID;,;ActiveDutyShortName, ActiveDutyName;,;(Active=1);,;ActiveDutyShortName, ActiveDutyName;,;;,;Ninguna;;;-1ÞÞÞBudgetsSpecificDuties;,;SpecificDutyID;,;SpecificDutyShortName, SpecificDutyName;,;(Active=1);,;SpecificDutyShortName, SpecificDutyName;,;;,;Ninguna;;;-1ÞÞÞBudgetsPrograms;,;ProgramID;,;ProgramShortName, ProgramName;,;(Active=1);,;ProgramShortName, ProgramName;,;;,;Ninguno;;;-1ÞÞÞBudgetsActivities1;,;ActivityID;,;ActivityShortName As ActivityShortName1, BudgetsActivities1.ActivityName As ActivityName1;,;(BudgetsActivities1.Active=1);,;ActivityShortName, ActivityName;,;;,;Ninguna;;;-1ÞÞÞBudgetsActivities2;,;ActivityID;,;ActivityShortName As ActivityShortName2, BudgetsActivities2.ActivityName As ActivityName2;,;(BudgetsActivities2.Active=1);,;BudgetsActivities2.ActivityShortName, BudgetsActivities2.ActivityName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞBudgetsProcesses;,;ProcessID;,;ProcessShortName, ProcessName;,;(Active=1);,;ProcessShortName, ProcessName;,;;,;Ninguna;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4,5,6,7,8,9,10,11,12,13,14,15", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 14
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 15
		Case "ServicesCenterTypesLKP"
			aCatalogComponent(S_NAME_CATALOG) = "Servicios por tipo de centro de trabajo"
			aCatalogComponent(S_ORDER_CATALOG) = "RecordShortName, ServiceShortName, CenterTypeShortName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Servicios,Tipo de centro de trabajo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "RecordID,RecordShortName,ServiceID,CenterTypeID"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,6,6"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,-1,-1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,-1,-1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞServices;,;ServiceID;,;ServiceShortName, ServiceName;,;(EndDate=30000000) And (Active=1);,;ServiceShortName, ServiceName;,;;,;Ninguno;;;-1ÞÞÞCenterTypes;,;CenterTypeID;,;CenterTypeShortName, CenterTypeName;,;(EndDate=30000000) And (Active=1);,;CenterTypeShortName, CenterTypeName;,;;,;Ninguno;;;-1"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3", ",")
			aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (Services.EndDate=30000000) And (CenterTypes.EndDate=30000000)"
		Case "Shifts"
			aCatalogComponent(S_NAME_CATALOG) = "Horarios"
			aCatalogComponent(S_ORDER_CATALOG) = "ShiftShortName, Shifts.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Horario,Entrada 1,Salida 1,Entrada 2,Salida 2,Horas laboradas,Tipo de jornada,Turno,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "ShiftID,ShiftShortName,ShiftName,StartHour1,EndHour1,StartHour2,EndHour2,WorkingHours,JourneyTypeID,JourneyID,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,0,0,1,1,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,3,3,3,3,2,6,6,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,100,0,0,0,0,4,1,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,15,15,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0,0,0,0,1,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,24,24,24,24,24,4,-1," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,0,0,0,0,,1,-1" & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,0,0,0,0,,1,-1" & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ<OPTION VALUE=""1"">1</OPTION><OPTION VALUE=""2"">2</OPTION><OPTION VALUE=""3"">3</OPTION><OPTION VALUE=""4"">4</OPTION>ÞÞÞJourneys;,;JourneyID;,;JourneyShortName, JourneyName;,;(EndDate=30000000) And (Active=1);,;JourneyShortName, JourneyName;,;;,;Ninguno;;;-1ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4,5,6,7,8,9,10,11", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 10
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 11
		Case "States"
			aCatalogComponent(S_NAME_CATALOG) = "Entidades federativas"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Clave,Nombre abreviado,Nombre,Banco para pago de cheques,Activo"
			aCatalogComponent(S_ORDER_CATALOG) = "StateName"
			aCatalogComponent(N_NAME_CATALOG) = 3
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "StateID,StateCode,StateShortName,StateName,BankID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,5,6,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,10,100,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,-1,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,,3,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,,3,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞBanks;,;BankID;,;BankName;,;;,;BankName;,;;,;Ninguno;;;-1ÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,3", ",")
aCatalogComponent(S_URL_PARAMETERS_CATALOG) = "Catalogs.asp?Action=SubStates&ParentID=<FIELD_0 />"
		Case "Status", "StatusAreas", "StatusBudgets", "StatusForms", "StatusLevels", "StatusPaperworks", "StatusPositions"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Activo"
			Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
				Case "StatusAreas"
					aCatalogComponent(S_NAME_CATALOG) = "Estatus de las áreas"
				Case "StatusBudgets"
					aCatalogComponent(S_NAME_CATALOG) = "Estatus de las partidas presupuestales"
				Case "StatusForms"
					aCatalogComponent(S_NAME_CATALOG) = "Estatus de los formularios"
				Case "StatusJobs"
					aCatalogComponent(S_NAME_CATALOG) = "Estatus de las plazas"
				Case "StatusLevels"
					aCatalogComponent(S_NAME_CATALOG) = "Estatus de los niveles"
				Case "StatusPaperworks"
					aCatalogComponent(S_NAME_CATALOG) = "Estatus de los trámites"
					aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Con este estatus el trámite estará abierto"
				Case "StatusPayments"
					aCatalogComponent(S_NAME_CATALOG) = "Estatus de los pagos"
				Case "StatusPositions"
					aCatalogComponent(S_NAME_CATALOG) = "Estatus de los puestos"
				Case Else
					aCatalogComponent(S_NAME_CATALOG) = "Estatus"
			End Select
			aCatalogComponent(S_ORDER_CATALOG) = "StatusName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "StatusID,StatusName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,100,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
		Case "StatusJobs", "StatusPayments"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Clave,Nombre,Activo"
			Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
				Case "StatusJobs"
					aCatalogComponent(S_NAME_CATALOG) = "Estatus de las plazas"
				Case "StatusPayments"
					aCatalogComponent(S_NAME_CATALOG) = "Estatus de los pagos"
			End Select
			aCatalogComponent(S_ORDER_CATALOG) = "StatusShortName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "StatusID,StatusShortName,StatusName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,100,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
		Case "StatusEmployees"
			aCatalogComponent(S_NAME_CATALOG) = "Estatus de los empleados"
			aCatalogComponent(S_ORDER_CATALOG) = "StatusName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Tipo de movimiento,Clasificación del tipo de movimiento,Con este estatus el empleado estará activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "StatusID,StatusName,ReasonID,StatusReasonID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,6,6,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,255,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,-1,-1,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,-1,-1,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = "ÞÞÞÞÞÞReasons;,;ReasonID;,;ReasonShortName, ReasonName;,;;,;ReasonShortName;,;;,;Ninguno;;;-1ÞÞÞReasonTypes;,;ReasonTypeID;,;ReasonTypeName;,;(Active=1);,;ReasonTypeName;,;;,;Ninguno;;;-1ÞÞÞ"
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2", ",")
			aCatalogComponent(N_ACTIVE_CATALOG) = -2
		Case "SubBranches"
			aCatalogComponent(S_NAME_CATALOG) = "Subramas"
			aCatalogComponent(S_ORDER_CATALOG) = "SubBranchShortName, SubBranches.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "SubBranchID,SubBranchShortName,SubBranchName,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,255,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 3
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 4
		Case "SubjectTypes"
			aCatalogComponent(S_NAME_CATALOG) = "Tipos de asunto"
			aCatalogComponent(S_ORDER_CATALOG) = "SubjectTypeID, SubjectTypeName"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Nombre,Días para atención,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "SubjectTypeID,SubjectTypeName,DaysForAttention,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,4,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,100,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,15,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,365,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,0,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,0,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("0,1,2", ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
		Case "Syndicates"
			aCatalogComponent(S_NAME_CATALOG) = "Sindicatos"
			aCatalogComponent(S_ORDER_CATALOG) = "SyndicateShortName, Syndicates.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "SyndicateID,SyndicateShortName,SyndicateName,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,100,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 3
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 4
		Case "WorkingCenters"
			aCatalogComponent(S_NAME_CATALOG) = "Centros de trabajo"
			aCatalogComponent(S_ORDER_CATALOG) = "WorkingCenterShortName, WorkingCenters.EndDate Desc"
			aCatalogComponent(N_NAME_CATALOG) = 1
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Código,Nombre,Fecha de inicio,Fecha de término,Fecha de modificación,Modificó,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "WorkingCenterID,WorkingCenterShortName,WorkingCenterName,StartDate,EndDate,ModifyDate,UserID,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1,0,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,5,5,1,1,11,11,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,5,255,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0,0,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0," & N_START_YEAR & "," & N_START_YEAR & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,,," & Year(Date()) & "," & Year(Date()) + 10 & ",0,0,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,,," & Left(GetSerialNumberForDate(""), Len("00000000")) & ",0," & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & aLoginComponent(N_USER_ID_LOGIN) & ",1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1,2,3,4", ",")
			aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = 3
			aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = 4
		Case "ZoneTypes"
			aCatalogComponent(S_NAME_CATALOG) = "Áreas geográficas"
			aCatalogComponent(S_ORDER_CATALOG) = "ZoneTypeID"
			aCatalogComponent(N_NAME_CATALOG) = 2
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = "ID,Zona económica,Nombre,Activo"
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = "ZoneTypeID,ZoneTypeID2,ZoneTypeName,Active"
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = "1,1,1,1"
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = "4,11,5,0"
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = "0,1,1,0"
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = "15,0,0,0"
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = "0,0,0,0"
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = "100000000,0,,0"
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1,4,,1"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1,4,,1"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("2", ",")
			aCatalogComponent(AS_SCRIPT_CATALOG) = "ÞÞÞÞÞÞ />&nbsp;&nbsp;&nbsp;<SELECT NAME=""ZoneTypeID2Temp"" ID=""ZoneTypeID2Temp"" CLASS=""Lists"" onChange=""document.CatalogFrm.ZoneTypeID2.value = parseInt(this.value) + 2;"">" & GenerateListOptionsFromQuery(oADODBConnection, "EconomicZones", "EconomicZoneID", "EconomicZoneName", "(EconomicZoneID>0) And (Active=1)", "EconomicZoneID", "", "Ninguna;;;-1", sErrorDescription) & "</SELECT><INPUT TYPE=""HIDDEN"" ÞÞÞ"
			aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)
	End Select
	If aCatalogComponent(N_ACTIVE_CATALOG) = -1 Then aCatalogComponent(N_ACTIVE_CATALOG) = UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
	If Len(oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Item) > 0 Then
		aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) = oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Item
	End If

	InitializeCatalogs = Err.number
	Err.Clear
End Function

Function GetFilterFromURL(oRequest, sAction, sErrorDescription)
'************************************************************
'Purpose: To get the Filter condition from the URL.
'Inputs:  oRequest, sAction
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetFilterFromURL"
	Dim lErrorNumber

	Select Case sAction
		Case "EmployeeFields"
		Case "Forms"
		Case "FormFields"
		Case "Profiles"
		Case "Users"
		Case "Zones"
		Case Else
	End Select

	GetFilterFromURL = lErrorNumber
	Err.Clear
End Function

Function DoAction(sAction, bShowForm, sErrorDescription)
'************************************************************
'Purpose: To add, change or delete the information of the
'         specified catalog.
'Inputs:  sAction
'Outputs: sAction, bShowForm, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DoAction"
	Dim sCondition
	Dim oRecordset
	Dim lErrorNumber

	Select Case sAction
		Case "BanamexCensus"
			If (Len(oRequest("ModifyCensus").Item) > 0) Or (Len(oRequest("Add").Item) > 0) Or _
				(Len(oRequest("RemoveCensus").Item) > 0) Then
				aBanamexCensusComponent(N_EMPLOYEE_ID) = oRequest("EmployeeID").Item
				If (Len(oRequest("ModifyCensus").Item) > 0) Or (Len(oRequest("Add").Item) > 0) Then
					aBanamexCensusComponent(N_U_VERSION) = oRequest("u_version").Item
					aBanamexCensusComponent(N_EMPLOYEE_ID) = oRequest("EmployeeID").Item
					aBanamexCensusComponent(S_RFC) = oRequest("RFC").Item
					aBanamexCensusComponent(S_CURP) = oRequest("CURP").Item
					aBanamexCensusComponent(S_SOCIAL_SECURITY_NUMBER) = oRequest("SocialSecurityNumber").Item
					aBanamexCensusComponent(S_EMPLOYEE_LAST_NAME) = oRequest("EmployeeLastName").Item
					aBanamexCensusComponent(S_EMPLOYEE_LAST_NAME_2) = oRequest("EmployeeLastName2").Item
					aBanamexCensusComponent(S_EMPLOYEE_NAME) = oRequest("EmployeeName").Item
					aBanamexCensusComponent(S_CT) = oRequest("CT").Item
					aBanamexCensusComponent(N_BIRTH_DATE) =  oRequest("BirthDateYear").Item & oRequest("BirthDateMonth").Item & oRequest("BirthDateDay").Item
					aBanamexCensusComponent(N_BIRTH_STATE_ID) = oRequest("BirthState").Item
					aBanamexCensusComponent(S_GENDER_SHORT_NAME) = oRequest("GenderShortName").Item
					aBanamexCensusComponent(N_JOIN_DATE) = oRequest("JoinDateYear").Item & oRequest("JoinDateMonth").Item & oRequest("JoinDateDay").Item
					aBanamexCensusComponent(N_COT_DATE) = oRequest("CotDateYear").Item & oRequest("CotDateMonth").Item & oRequest("CotDateDay").Item
					aBanamexCensusComponent(N_SALARY) = oRequest("Salary").Item
					aBanamexCensusComponent(N_FOVY) = oRequest("Fovi").Item
					aBanamexCensusComponent(N_PERIOD_ID) = oRequest("PeriodID").Item
					aBanamexCensusComponent(N_STATUS_ID) = 2
					aBanamexCensusComponent(N_CHANGE_FLAG) = 1
					aBanamexCensusComponent(N_MARITAL_STATUS_ID) = oRequest("MaritalStatus").Item
					aBanamexCensusComponent(S_ADDRESS) = oRequest("Address").Item
					aBanamexCensusComponent(S_COLONY) = oRequest("Colony").Item
					aBanamexCensusComponent(S_CITY) = oRequest("City").Item
					aBanamexCensusComponent(S_STATE) = oRequest("State").Item
					aBanamexCensusComponent(N_NOMBRAM) = oRequest("Nombram").Item
					aBanamexCensusComponent(N_AFORE) = oRequest("Afore").Item
					aBanamexCensusComponent(N_ICEFA) = oRequest("ICEFA").Item
					aBanamexCensusComponent(N_IC_NUMBER) = oRequest("ICNumber").Item
					aBanamexCensusComponent(S_MOT_BAJA) = oRequest("mot_baja").Item
					aBanamexCensusComponent(N_SALARY_V) = oRequest("Salary_v").Item
					aBanamexCensusComponent(N_FULL_PAY) = oRequest("FullPay").Item
					aBanamexCensusComponent(N_WORKING_DAYS) = oRequest("WorkingDays").Item
					aBanamexCensusComponent(N_INABILITY_DAYS) = oRequest("InabilityDays").Item
					aBanamexCensusComponent(N_ABSENCE_DAYS) = oRequest("AbsenceDays").Item
					aBanamexCensusComponent(N_EMPLOYEE_CONTRIBUTIONS) = oRequest("EmployeeContributions").Item
					aBanamexCensusComponent(N_EMPLOYEE_CONTRIBUTIONS_AMOUNT) = oRequest("EmployeeContributionsAmount").Item
					aBanamexCensusComponent(N_START_DATE_FOR_CENSUS) = oRequest("StartDateYear").Item & oRequest("StartDateMonth").Item & oRequest("StartDateDay").Item
					aBanamexCensusComponent(N_END_DATE_FOR_CENSUS) = oRequest("EndDateYear").Item & oRequest("EndDateMonth").Item & oRequest("EndDateDay").Item
					aBanamexCensusComponent(B_CHECK_FOR_DUPLICATED_BANAMEX_CENSUS) = False
					aBanamexCensusComponent(B_COMPONENT_INITIALIZED_BANAMEX_CENSUS) = True
					If Len(oRequest("ModifyCensus").Item) > 0 Then
						lErrorNumber = ModifyBanamexCensusRecord(oRequest, oADODBConnection, aBanamexCensusComponent, sErrorDescription)
					ElseIf Len(oRequest("Add").Item) > 0 Then
						lErrorNumber = AddBanamexCensusRecord(oRequest, oADODBConnection, aBanamexCensusComponent, sErrorDescription)
					End If
				ElseIf Len(oRequest("RemoveCensus").Item) > 0 Then
					lErrorNumber = MarkRecordForDeleting(oRequest, oADODBConnection, aBanamexCensusComponent, sErrorDescription)
				End If
			End If
		Case "CurrenciesHistoryList"
			If Len(oRequest("Modify").Item) > 0 Then
				sErrorDescription = "No se pudo modificar la información del registro."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update CurrenciesHistoryList Set CurrencyValue=" & oRequest("CurrencyValue").Item & " Where (CurrencyID=" & oRequest("CurrencyID").Item & ") And (CurrencyDate>=" & oRequest("CurrencyDateYear").Item & oRequest("CurrencyDateMonth").Item & oRequest("CurrencyDateDay").Item & ")", "CatalogsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudo modificar la información del registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Currencies Set CurrencyValue=" & oRequest("CurrencyValue").Item & " Where (CurrencyID=" & oRequest("CurrencyID").Item & ")", "CatalogsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
		Case "EmployeeFields"
			If Len(oRequest("Add").Item) > 0 Then
				lErrorNumber = AddEmployeeField(oRequest, oADODBConnection, aEmployeeFieldComponent, sErrorDescription)
			ElseIf Len(oRequest("Modify").Item) > 0 Then
				lErrorNumber = ModifyEmployeeField(oRequest, oADODBConnection, aEmployeeFieldComponent, sErrorDescription)
			ElseIf Len(oRequest("Remove").Item) > 0 Then
				lErrorNumber = RemoveEmployeeField(oRequest, oADODBConnection, aEmployeeFieldComponent, sErrorDescription)
			End If
			Redim aEmployeeFieldComponent(N_EMPLOYEE_FIELD_COMPONENT_SIZE)
			aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD) = -1
			aEmployeeFieldComponent(N_IS_OPTIONAL_EMPLOYEE_FIELD) = 0
			aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD) = 5
			aEmployeeFieldComponent(N_SIZE_EMPLOYEE_FIELD) = 1
			aEmployeeFieldComponent(N_LIMIT_TYPE_EMPLOYEE_FIELD) = 0
			aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD) = 0
			aEmployeeFieldComponent(N_MAXIMUM_EMPLOYEE_FIELD) = 0
			aEmployeeFieldComponent(S_DSN_FOR_SOURCE_EMPLOYEE_FIELD) = SIAP_DATABASE_PATH
			aEmployeeFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_EMPLOYEE_FIELD) = iConnectionType
			aEmployeeFieldComponent(B_COMPONENT_INITIALIZED_EMPLOYEE_FIELD) = True
		Case "EmploymentAllowances"
			If Len(oRequest("Modify").Item) > 0 Then
				lErrorNumber = ModifyEmploymentAllowances(oRequest, oADODBConnection, sErrorDescription)
			End If
		Case "FormFields"
			If Len(oRequest("Add").Item) > 0 Then
				lErrorNumber = AddFormField(oRequest, oADODBConnection, aFormFieldComponent, sErrorDescription)
			ElseIf Len(oRequest("Import").Item) > 0 Then
				lErrorNumber = ImportFormField(oRequest, oADODBConnection, aFormFieldComponent, sErrorDescription)
			ElseIf Len(oRequest("Modify").Item) > 0 Then
				lErrorNumber = ModifyFormField(oRequest, oADODBConnection, aFormFieldComponent, sErrorDescription)
			ElseIf Len(oRequest("Remove").Item) > 0 Then
				lErrorNumber = RemoveFormField(oRequest, oADODBConnection, aFormFieldComponent, sErrorDescription)
			End If
			If Len(oRequest("Import").Item) = 0 Then
				Redim aFormFieldComponent(N_FORM_FIELD_COMPONENT_SIZE)
				aFormFieldComponent(N_ID_FORM_FIELD) = CLng(oRequest("FormID").Item)
				aFormFieldComponent(N_FIELD_ID_FORM_FIELD) = -1
				aFormFieldComponent(S_DSN_FORM_FIELD) = SIAP_DATABASE_PATH
				aFormFieldComponent(N_CONNECTION_TYPE_FORM_FIELD) = iConnectionType
				aFormFieldComponent(S_DATABASE_NAME_FORM_FIELD) = SIAP_DATABASE_NAME
				aFormFieldComponent(S_TABLE_NAME_FORM_FIELD) = "FormFields"
				aFormFieldComponent(N_IS_OPTIONAL_FORM_FIELD) = 0
				aFormFieldComponent(N_TYPE_ID_FORM_FIELD) = 5
				aFormFieldComponent(N_SIZE_FORM_FIELD) = 1
				aFormFieldComponent(N_LIMIT_TYPE_FORM_FIELD) = 0
				aFormFieldComponent(N_MINIMUM_FORM_FIELD) = 0
				aFormFieldComponent(N_MAXIMUM_FORM_FIELD) = 0
				aFormFieldComponent(S_DSN_FOR_SOURCE_FORM_FIELD) = aFormFieldComponent(S_DSN_FORM_FIELD)
				aFormFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_FORM_FIELD) = iConnectionType
				aFormFieldComponent(B_COMPONENT_INITIALIZED_FORM_FIELD) = True
			End If
		Case "Forms"
			If Len(oRequest("Add").Item) > 0 Then
				lErrorNumber = AddForm(oRequest, oADODBConnection, aFormComponent, sErrorDescription)
			ElseIf Len(oRequest("Import").Item) > 0 Then
				lErrorNumber = ImportForm(oRequest, oADODBConnection, aFormComponent, sErrorDescription)
			ElseIf Len(oRequest("Modify").Item) > 0 Then
				lErrorNumber = ModifyForm(oRequest, oADODBConnection, aFormComponent, sErrorDescription)
			ElseIf Len(oRequest("Remove").Item) > 0 Then
				lErrorNumber = RemoveForm(oRequest, oADODBConnection, aFormComponent, sErrorDescription)
			ElseIf Len(oRequest("SetActive").Item) > 0 Then
				lErrorNumber = SetActiveForForm(oRequest, oADODBConnection, aFormComponent, sErrorDescription)
				Redim aFormComponent(N_FORM_COMPONENT_SIZE)
				aFormComponent(N_ID_FORM) = -1
				bShowForm = False
			End If
			If Len(oRequest("Import").Item) = 0 Then
				Redim aFormComponent(N_FORM_COMPONENT_SIZE)
				aFormComponent(N_ID_FORM) = -1
			End If
		Case "PayrollResume"
			If Len(oRequest("ModifyResume").Item) > 0 Or Len(oRequest("Remove").Item) > 0 Then
				aPayrollResumeForSarComponent(N_SOCIETY_ID) = oRequest("SocietyID").Item
				aPayrollResumeForSarComponent(N_COMPANY_ID) = oRequest("CompanyID").Item
				aPayrollResumeForSarComponent(S_CLC) = oRequest("CLC").Item
				aPayrollResumeForSarComponent(N_BANK_ID) = GetBankIDFromShortName(oRequest)
				aPayrollResumeForSarComponent(S_BANK_SHORT_NAME) = oRequest("BankID").Item
				aPayrollResumeForSarComponent(N_PAYMENT_DATE) = oRequest("PaymentDate").Item
				aPayrollResumeForSarComponent(N_EMPLOYEE_TYPE_ID) = oRequest("EmployeeTypeID").Item
				aPayrollResumeForSarComponent(S_EMPLOYEE_TYPE_NAME) = oRequest("EmployeeType").Item
				aPayrollResumeForSarComponent(B_COMPONENT_INITIALIZED_PAYROLL_RESUME_FOR_SAR_COMPONENT) = True
				aPayrollResumeForSarComponent(B_CHECK_FOR_DUPLICATED_PAYROLL_RESUME_FOR_SAR) = False
				If Len(oRequest("ModifyResume").Item) > 0 Then
					aPayrollResumeForSarComponent(N_INCOME) = oRequest("Income").Item
					aPayrollResumeForSarComponent(N_DEDUCTIONS) = oRequest("Deductions").Item
					aPayrollResumeForSarComponent(N_NET_INCOME) = oRequest("NetIncome").Item
					aPayrollResumeForSarComponent(N_CPT_01) = oRequest("Cpt_01").Item
					aPayrollResumeForSarComponent(N_CPT_04) = oRequest("Cpt_04").Item
					aPayrollResumeForSarComponent(N_CPT_05) = oRequest("Cpt_05").Item
					aPayrollResumeForSarComponent(N_CPT_06) = oRequest("Cpt_06").Item
					aPayrollResumeForSarComponent(N_CPT_07) = oRequest("Cpt_07").Item
					aPayrollResumeForSarComponent(N_CPT_08) = oRequest("Cpt_08").Item
					aPayrollResumeForSarComponent(N_CPT_11) = oRequest("Cpt_11").Item
					aPayrollResumeForSarComponent(N_CPT_44) = oRequest("Cpt_44").Item
					aPayrollResumeForSarComponent(N_CPT_B2) = oRequest("Cpt_B2").Item
					aPayrollResumeForSarComponent(N_CPT_7S) = oRequest("Cpt_7S").Item
					lErrorNumber = ModifyPayrollResumeForSarRecord(oRequest, oADODBConnection, aPayrollResumeForSarComponent, sErrorDescription)
				End If
				If Len(oRequest("Remove").Item) > 0 Then
					lErrorNumber = RemovePayrollResumeForSarRecord(oRequest, oADODBConnection, aPayrollResumeForSarComponent, sErrorDescription)
				End If
			End If
		Case "PayrollsClcs"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PayrollsClcs Set PayrollCLC='', PayrollCode='', FilterParameters='' Where PayrollCLC = '" & oRequest("PayrollClc").Item & "' ", "CatalogsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Case "Profiles"
			If Len(oRequest("Add").Item) > 0 Then
				lErrorNumber = AddProfile(oRequest, oADODBConnection, aProfileComponent, sErrorDescription)
			ElseIf Len(oRequest("Modify").Item) > 0 Then
				lErrorNumber = ModifyProfile(oRequest, oADODBConnection, aProfileComponent, sErrorDescription)
			ElseIf Len(oRequest("Remove").Item) > 0 Then
				lErrorNumber = RemoveProfile(oRequest, oADODBConnection, aProfileComponent, sErrorDescription)
			End If
		Case "ProfessionalRiskMatrix"
			aProfessionalRiskComponent(N_BRANCH_ID) = oRequest("BranchID").Item
			aProfessionalRiskComponent(N_CENTER_TYPE_ID_PROFESSIONAL_RISK) = oRequest("CenterTypeID").Item
			aProfessionalRiskComponent(N_POSITION_ID) = oRequest("PositionID").Item
			aProfessionalRiskComponent(N_SERVICE_ID) = oRequest("ServiceID").Item
			aProfessionalRiskComponent(B_IS_DUPLICATED_PROFESSIONAL_RISK) = False
			If Len(aProfessionalRiskComponent(N_BRANCH_ID)) = 0 Then
				aProfessionalRiskComponent(B_COMPONENT_INITIALIZED_PROFESSIONAL_RISK) = False
			Else
				aProfessionalRiskComponent(B_COMPONENT_INITIALIZED_PROFESSIONAL_RISK) = True
			End If
			If aProfessionalRiskComponent(B_COMPONENT_INITIALIZED_PROFESSIONAL_RISK) Then
				lErrorNumber = GetProfessionalRisk(oRequest, oADODBConnection, aProfessionalRiskComponent, sErrorDescription)
			End If
			If StrComp(oRequest("Modify").Item, "Modificar", vbBinaryCompare) = 0 Then
				lErrorNumber = ModifyProfessionalRisk(oRequest, oADODBConnection, aProfessionalRiskComponent, sErrorDescription)
			ElseIf Len(Orequest("Remove").Item) > 0 Then
				lErrorNumber = RemoveProfessionalRisk(oRequest, oADODBConnection, aProfessionalRiskComponent, sErrorDescription)
			End If
		Case "Projects", "TACO_Projects"
			If Len(oRequest("Add").Item) > 0 Then
				lErrorNumber = AddCatalog(oRequest, oSIAPTACOADODBConnection, aCatalogComponent, sErrorDescription)
				Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
					Case "111"
				End Select
				If lErrorNumber = 0 Then
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_DEFAULT_VALUES_CATALOG)
				End If
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_DEFAULT_VALUES_CATALOG)
			ElseIf Len(oRequest("Modify").Item) > 0 Then
				sCondition = aCatalogComponent(S_QUERY_CONDITION_CATALOG)
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = ""
				lErrorNumber = ModifyCatalog(oRequest, oSIAPTACOADODBConnection, aCatalogComponent, sErrorDescription)
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = sCondition
			ElseIf Len(oRequest("Remove").Item) > 0 Then
				sCondition = aCatalogComponent(S_QUERY_CONDITION_CATALOG)
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = ""
				lErrorNumber = RemoveCatalog(oRequest, oSIAPTACOADODBConnection, aCatalogComponent, sErrorDescription)
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = sCondition
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) = -1
			ElseIf Len(oRequest("SetActive").Item) > 0 Then
				lErrorNumber = SetActiveForCatalog(oRequest, oSIAPTACOADODBConnection, aCatalogComponent, sErrorDescription)
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_DEFAULT_VALUES_CATALOG)
				bShowForm = False
			End If
		Case "Tasks"
			If Len(oRequest("Add").Item) > 0 Then
				lErrorNumber = AddTask(oRequest, oSIAPTACOADODBConnection, aTaskComponent, sErrorDescription)
			ElseIf Len(oRequest("Import").Item) > 0 Then
				lErrorNumber = ImportTask(oRequest, oSIAPTACOADODBConnection, aTaskComponent, sErrorDescription)
			ElseIf Len(oRequest("Modify").Item) > 0 Then
				lErrorNumber = ModifyTask(oRequest, oSIAPTACOADODBConnection, aTaskComponent, sErrorDescription)
			ElseIf Len(oRequest("Remove").Item) > 0 Then
				lErrorNumber = RemoveTask(oRequest, oSIAPTACOADODBConnection, aTaskComponent, sErrorDescription)
			End If
			If Len(oRequest("Import").Item) = 0 Then
				Redim aTaskComponent(N_TASK_COMPONENT_SIZE)
				aTaskComponent(N_ID_TASK) = -1
				aTaskComponent(N_PROJECT_ID_TASK) = oRequest("ProjectID").Item
				aTaskComponent(S_PATH_TASK) = oRequest("TaskPath").Item
			End If
		Case "TaxInvertions"
			If Len(oRequest("Modify").Item) > 0 Then
				lErrorNumber = ModifyTaxInvertions(oRequest, oADODBConnection, sErrorDescription)
			End If
		Case "TaxLimits"
			If Len(oRequest("Modify").Item) > 0 Then
				lErrorNumber = ModifyTaxLimits(oRequest, oADODBConnection, sErrorDescription)
			End If
		Case "Users"
			If Len(oRequest("Add").Item) > 0 Then
				lErrorNumber = AddUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
				If lErrorNumber <> 0 Then aUserComponent(N_ID_USER) = -2
			ElseIf Len(oRequest("Import").Item) > 0 Then
				lErrorNumber = ImportUser(oRequest, oADODBConnection, oSADEADODBConnection, aUserComponent, sErrorDescription)
			ElseIf Len(oRequest("Modify").Item) > 0 Then
				lErrorNumber = ModifyUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
			ElseIf Len(oRequest("Remove").Item) > 0 Then
				lErrorNumber = RemoveUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
				If lErrorNumber = 0 Then aUserComponent(N_ID_USER) = -2
			ElseIf Len(oRequest("SetActive").Item) > 0 Then
				lErrorNumber = SetActiveForUser(oRequest, oADODBConnection, aUserComponent, sErrorDescription)
				Redim aUserComponent(N_USER_COMPONENT_SIZE)
				aUserComponent(N_ID_USER) = -2
				bShowForm = False
			ElseIf Len(oRequest("Unlock").Item) > 0 Then
				sErrorDescription = "No se pudo modificar la información del registro."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Users Set SecurityLock=0 Where (UserAccessKey='" & Replace(oRequest.Item("UserToUnlock"), "'", "") & "')", "CatalogsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		Case "Zones"
			If Len(oRequest("Add").Item) > 0 Then
				lErrorNumber = AddZone(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
			ElseIf Len(oRequest("Modify").Item) > 0 Then
				lErrorNumber = ModifyZone(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
			ElseIf Len(oRequest("Remove").Item) > 0 Then
				lErrorNumber = RemoveZone(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
			ElseIf Len(oRequest("SetActive").Item) > 0 Then
				lErrorNumber = SetActiveForZone(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
				Redim aZoneComponent(N_ZONE_COMPONENT_SIZE)
				aZoneComponent(N_ID_ZONE) = -1
				aZoneComponent(N_PARENT_ID_ZONE) = oRequest("ParentID").Item
				bShowForm = False
			End If
		Case Else
			Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
				Case "BankAccounts"
					If StrComp(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(6), "00000", vbBinaryCompare) = 0 Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(6) = 30000000
				Case "PositionsSpecialJourneysLKP"
					If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2)) = 0 Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2) = 30000000
			End Select
			If Len(oRequest("Add").Item) > 0 Then
				Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
					Case "EmployeesAntiquitiesLKP"
						lErrorNumber = VerifyExistencyOfRecordsInPeriodDates(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
						If aCatalogComponent(B_IS_DUPLICATED_CATALOG) Then
							lErrorNumber = L_ERR_DUPLICATED_RECORD
							sErrorDescription = "Ya existe un registro en el período indicado."
							Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogsLib.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
						End If
						If lErrorNumber = 0 Then
							If Not VerifyPeriodDatesForAntiquities(CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1)), CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2)), CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(5)), CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(6)), CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(7)), sErrorDescription) Then
								lErrorNumber = -1
							End If
						End If
				End Select
				If lErrorNumber = 0 Then
					Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
						Case "PositionsSpecialJourneysLKP"
							If aConceptComponent(N_POSITION_ID_CONCEPT) = -1 Then
								lErrorNumber = -1
								sErrorDescription = "No se especificó el identificador del puesto para registrarlo para guardias y suplencias."
								Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
							Else
								If aConceptComponent(N_END_DATE_CONCEPT) = 0 Then aConceptComponent(N_END_DATE_CONCEPT) = 30000000
								lErrorNumber = GetPositionDataForSpecialJourneysLKP(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
								If lErrorNumber = 0 Then
									lErrorNumber = AddPositionsSpecialJourneysLKP(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
								End If
							End If
						Case "EmployeesSpecialJourneys"
							If Len(oRequest("Add").Item) > 0 Then
								lErrorNumber = AddEmployeesSpecialJourney(oRequest, oADODBConnection, sErrorDescription)
							End If
						Case Else
							lErrorNumber = AddCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
					End Select
				End If
				If lErrorNumber = 0 Then
					Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
						Case "PaperworkLists"
							lErrorNumber = UpdateConsecutiveID(oADODBConnection, 1062, aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1), sErrorDescription)
							lErrorNumber = GetConsecutiveID(oADODBConnection, 1062, aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1), sErrorDescription)
							lErrorNumber = GetConsecutiveID(oADODBConnection, 1062, aCatalogComponent(AS_DEFAULT_VALUES_CATALOG)(1), sErrorDescription)
					End Select
				End If
				If lErrorNumber = 0 Then
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_DEFAULT_VALUES_CATALOG)
				End If
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_DEFAULT_VALUES_CATALOG)
			ElseIf Len(oRequest("Apply").Item) > 0 Then
				Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
					Case "PositionsSpecialJourneysLKP"
						lErrorNumber = SetActiveForPositionsSpecialJourneys(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
				End Select
			ElseIf Len(oRequest("Modify").Item) > 0 Then
				sCondition = aCatalogComponent(S_QUERY_CONDITION_CATALOG)
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = ""
				Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
					Case "PositionsHierarchy"
						If Len(oRequest("ParentJobID").Item) > 0 Then aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (AreaCode='" & oRequest("AreaCodeOld").Item & "') And (ParentJobID=" & oRequest("ParentJobID").Item & ")"
				End Select
				Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
					Case "ProfessionalRiskMatrix"
					Case "PositionsSpecialJourneysLKP"
						lErrorNumber = ModifyPositionsSpecialJourneysLKP(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
					Case Else
						lErrorNumber = ModifyCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
				End Select
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = sCondition
			ElseIf Len(oRequest("Remove").Item) > 0 Then
				sCondition = aCatalogComponent(S_QUERY_CONDITION_CATALOG)
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = ""
				Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
					Case "PositionsHierarchy"
						If Len(oRequest("ParentJobID").Item) > 0 Then aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (AreaCode='" & oRequest("AreaCodeOld").Item & "') And (ParentJobID=" & oRequest("ParentJobID").Item & ")"
				End Select
				lErrorNumber = RemoveCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
				Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
					Case "Status"
						'sErrorDescription = "No se pudo eliminar la información del registro."
						'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Substatus Where (" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & ")", "CatalogsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End Select
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = sCondition
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) = -1
			ElseIf Len(oRequest("SetActive").Item) > 0 Then
				lErrorNumber = SetActiveForCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_DEFAULT_VALUES_CATALOG)
				bShowForm = False
			End If
	End Select

	DoAction = lErrorNumber
	Err.Clear
End Function

Function DisplayFilters(oRequest, sAction, sErrorDescription)
'************************************************************
'Purpose: To display the filter of the specified catalog.
'Inputs:  sAction
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFilters"
	Dim bHasFilter
	Dim iIndex
	Dim lErrorNumber

	bHasFilter = False
	Response.Write "<FORM NAME=""FilterFrm"" ID=""FilterFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""GET"" /><FONT FACE=""Arial"" SIZE=""2"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
		Select Case sAction
			Case "EmployeeFields"
			Case "FormFields"
			Case "Forms"
			Case "Profiles"
			Case "Users"
			Case "Zones"
			Case Else
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Mostrar registros que contengan:&nbsp;</B>"
				Response.Write "<INPUT TYPE=""TEXT"" NAME=""FilterName"" ID=""FilterNameTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & oRequest("FilterName").Item & """ CLASS=""TextFields"" />"
				bHasFilter = True
		End Select
		If bHasFilter Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ApplyFilter"" ID=""ApplyFilterBtn"" VALUE=""  Filtrar  "" CLASS=""Buttons"" />"
	Response.Write "</FONT></FORM>"

	DisplayFilters = lErrorNumber
	Err.Clear
End Function

Function DisplayTables(sAction, sErrorDescription)
'************************************************************
'Purpose: To display the table of the specified catalog.
'Inputs:  sAction
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayTables"
	Dim sNames
	Dim lErrorNumber
	Dim sCaseOptions
	Dim iIndex

	Select Case sAction
		Case "BanamexCensus"
			If (Len(oRequest("Modify").Item) = 0) And (Len(oRequest("Delete").Item) = 0) And (Len(oRequest("AddNew").Item) = 0) Then
				If iStep > 1 Then lErrorNumber = DisplayBanamexCensusList(oRequest, oADODBConnection, sErrorDescription)
			Else
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"" WIDTH=""100%""><TR>"
				Response.Write "<TD VALIGN=""TOP""><DIV NAME=""ReportDiv"" ID=""ReportDiv"" STYLE=""height: 350px; width:500px; overflow: auto;"">"
						lErrorNumber = DisplayBanamexCensusList(oRequest, oADODBConnection, sErrorDescription)
					Response.Write "</DIV></TD>"
					Response.Write "<TD>&nbsp;</TD>"
					Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
					Response.Write "<TD>&nbsp;</TD>"
					Response.Write "<TD WIDTH=""*"" VALIGN=""TOP"">"
						lErrorNumber = DisplayBanamexCensusForm(oRequest, oADODBConnection, sAction, aBanamexCensusComponent, sErrorDescription)
					Response.Write "</TD></TR>"
				Response.Write "</TABLE>"
			End If
		Case "ConsarFile"
			If iStep > 1 Then lErrorNumber = DisplayConsarFile(oRequest, oADODBConnection, sErrorDescription)
		Case "Currencies"
			If Len(oRequest("ApplyFilter").Item) > 0 Then
				aCatalogComponent(S_URL_CATALOG) = aCatalogComponent(S_URL_CATALOG) & "FilterName=" & oRequest("FilterName").Item & "&ApplyFilter=1"
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (CurrencyName Like '" & S_WILD_CHAR & oRequest("FilterName").Item & S_WILD_CHAR & "')"
			End If
			lErrorNumber = DisplayCatalogsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, False, aCatalogComponent, sErrorDescription)
		Case "CurrenciesHistoryList"
			lErrorNumber = DisplayCurrencyHistoryTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, aCatalogComponent, sErrorDescription)
		Case "EmployeeFields"
			lErrorNumber = DisplayEmployeeFieldsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, aEmployeeFieldComponent, sErrorDescription)
		Case "EmployeesDeleted"
			If iStep > 1 Then lErrorNumber = DisplayDeletedHistoryList(oRequest, oADODBConnection, sErrorDescription)
		Case "EmploymentAllowances"
			lErrorNumber = DisplayEmploymentAllowancesTable(oRequest, oADODBConnection, False, sErrorDescription)
		Case "FormFields"
			lErrorNumber = DisplayFormFieldsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, aFormFieldComponent, sErrorDescription)
		Case "Forms"
			If (Len(oRequest("Import").Item) > 0) And (lErrorNumber = 0) Then
				Call DisplayErrorMessage("Confirmación", "La información del formulario fue importada con éxito.")
				Response.Write "<BR />"
			End If
			lErrorNumber = DisplayFormsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, aFormComponent, sErrorDescription)
		Case "Holidays"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "function AddHolidayDescription(sValue) {" & vbNewLine
					Response.Write "if (!VerifyHolidayIsMarked(sValue)) {" & vbNewLine
						Response.Write "oAnchor = document.getElementById(sValue);" & vbNewLine
						Response.Write "oText = document.getElementById('HolidayName').value;" & vbNewLine
						Response.Write "if (oText.length != 0) {" & vbNewLine
							'Response.Write "alert('id =' + sValue);" & vbNewLine
							Response.Write "if (oText.length == 0) oText = 'vacia' + ' otro';"
							Response.Write "document.getElementById(sValue).href += '&HolidayDescription=' + oText;" & vbNewLine
							Response.Write "return true;" & vbNewLine
						Response.Write "}"  & vbNewLine
						Response.Write "else {"  & vbNewLine
							Response.Write "alert('Ingrese la descripción para registrar el día de asueto');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}"  & vbNewLine
					Response.Write "}"  & vbNewLine
				Response.Write "} // End of AddHolidayDescription" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Descripcion:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""HolidayName"" ID=""HolidayNameTxt"" SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"
			Response.Write "<BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
				If (Len(oRequest("Year").Item) > 0) And (Len(oRequest("Month").Item) > 0) And (Len(oRequest("Day").Item) > 0) Then
					lErrorNumber = GetNameFromTable(oADODBConnection, "Holiday", oRequest("Year").Item & Right(("0" & oRequest("Month").Item), Len("00")) & Right(("0" & oRequest("Day").Item), Len("00")), "", ",", sNames, sErrorDescription)
					sErrorDescription = "No se pudo actualizar el listado de días de asueto."
					If Len(sNames) > 0 Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Holidays Where (Holiday=" & oRequest("Year").Item & Right(("0" & oRequest("Month").Item), Len("00")) & Right(("0" & oRequest("Day").Item), Len("00")) & ")", "CatalogsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Holidays (Holiday, HolidayDescription) Values (" & oRequest("Year").Item & Right(("0" & oRequest("Month").Item), Len("00")) & Right(("0" & oRequest("Day").Item), Len("00")) & ", '" & CStr(oRequest("HolidayDescription").Item) & "')", "CatalogsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
				Call InitializeCalendarComponent(oRequest, aCalendarComponent)
				aCalendarComponent(N_MONTH_CALENDAR) = 0
				aCalendarComponent(N_DAY_CALENDAR) = 0
				lErrorNumber = GetNameFromTable(oADODBConnection, "Holidays", aCalendarComponent(N_YEAR_CALENDAR), "", ",", aCalendarComponent(S_MARKED_DAYS_CALENDAR), sErrorDescription)
				aCalendarComponent(S_TARGET_PAGE_CALENDAR) = "Catalogs.asp?Action=Holidays"
				Response.Write "<TD VALIGN=""TOP""><A HREF=""Catalogs.asp?Action=Holidays&Year=" & aCalendarComponent(N_YEAR_CALENDAR) - 1 & """><IMG SRC=""Images/ArrLeftBlack.gif"" WIDTH=""7"" HEIGHT=""13"" BORDER=""0"" ALT=""" & aCalendarComponent(N_YEAR_CALENDAR) - 1 & """ /></A>&nbsp;</TD>"
				Response.Write "<TD VALIGN=""TOP"">"
					lErrorNumber = DisplayYear(oRequest, aCalendarComponent, sErrorDescription)
					If Len(sErrorDescription) > 0 Then
						lErrorNumber = DisplayErrorMessage("Error", sErrorDescription)
					End If
				Response.Write "</TD>"
				sURLForCalendar = RemoveParameterFromURLString(ReplaceValueInURLString(oRequest, "Year", aCalendarComponent(N_YEAR_CALENDAR) + 1), "Holiday")
				Response.Write "<TD VALIGN=""TOP""><A HREF=""Catalogs.asp?Action=Holidays&Year=" & aCalendarComponent(N_YEAR_CALENDAR) + 1 & """><IMG SRC=""Images/ArrRightBlack.gif"" WIDTH=""7"" HEIGHT=""13"" BORDER=""0"" ALT=""" & aCalendarComponent(N_YEAR_CALENDAR) + 1 & """ /></A></TD>"
			Response.Write "</TR></TABLE>"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "function VerifyHolidayIsMarked(sValue) {" & vbNewLine
					sCaseOptions = Split(aCalendarComponent(S_MARKED_DAYS_CALENDAR), "," , -1, vbBinaryCompare)
					Response.Write "switch (sValue) {" & vbNewLine
						Response.Write "case '-1':" & vbNewLine
							Response.Write "break;" & vbNewLine
						For iIndex = 0 To UBound(sCaseOptions)
							Response.Write "case '" & CLng(sCaseOptions(iIndex)) & "':" & vbNewLine
						Next
							Response.Write "return true;" & vbNewLine
						Response.Write "default:" & vbNewLine
							Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "} // End of VerifyHolidayIsMarked" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		Case "PayrollCompare"
			lErrorNumber = DisplayPayrollCompareList(oRequest, oADODBConnection, sErrorDescription)
		Case "PayrollResume"
			If Len(oRequest("Modify").Item) = 0 And Len(oRequest("Delete").Item) = 0 Then
				lErrorNumber = DisplayPayrollResumeForSarList(oRequest, oADODBConnection, sErrorDescription)
			Else
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR>"
				Response.Write "<TD VALIGN=""TOP""><DIV NAME=""ReportDiv"" ID=""ReportDiv"" STYLE=""height: 450px; width:500px; overflow: auto;"">"
						lErrorNumber = DisplayPayrollResumeForSarList(oRequest, oADODBConnection, sErrorDescription)
					Response.Write "</DIV></TD>"
					Response.Write "<TD>&nbsp;</TD>"
					Response.Write "<TD BGCOLOR=""" & S_MAIN_COLOR_FOR_GUI & """ WIDTH=""1"" ><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
					Response.Write "<TD>&nbsp;</TD>"
					Response.Write "<TD WIDTH=""*"" VALIGN=""TOP"">"
						lErrorNumber = DisplayPayrollResumeForSarForm(oRequest, oADODBConnection, sAction, aPayrollResumeForSarComponent, sErrorDescription)
					Response.Write "</TD></TR>"
				Response.Write "</TABLE>"
			End If
		Case "ProfessionalRiskMatrix"
			lErrorNumber = DisplayProfessionalRiskMatrix(oRequest, oADODBConnection, sErrorDescription)
		Case "Profiles"
			lErrorNumber = DisplayProfilesTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, aProfileComponent, sErrorDescription)
		Case "Projects", "TACO_Projects"
			lErrorNumber = DisplayCatalogsTable(oRequest, oSIAPTACOADODBConnection, DISPLAY_NOTHING, True, aCatalogComponent, sErrorDescription)
		Case "Tasks"
			Response.Write "<DIV NAME=""TasksTableDiv"" ID=""TasksTableDiv"">"
				lErrorNumber = DisplayTasksTable(oRequest, oSIAPTACOADODBConnection, DISPLAY_NOTHING, True, aTaskComponent, sErrorDescription)
			Response.Write "</DIV>"
		Case "TaxInvertions"
			lErrorNumber = DisplayTaxInvertionsTable(oRequest, oADODBConnection, False, sErrorDescription)
		Case "TaxLimits"
			lErrorNumber = DisplayTaxLimitsTable(oRequest, oADODBConnection, False, sErrorDescription)
		Case "Users"
			If (Len(oRequest("Import").Item) > 0) And (lErrorNumber = 0) Then
				Call DisplayErrorMessage("Confirmación", "La información del usuario fue importada con éxito. Es necesario que revise la información de este usuario, en especial sus permisos dentro del sistema, el correo electrónico de su jefe inmediato y la zona a la que pertenece.")
				Response.Write "<BR />"
			End If
			lErrorNumber = DisplayUsersTable(oRequest, oADODBConnection, DISPLAY_NOTHING, True, aUserComponent, sErrorDescription)
		Case "Zones"
			If Len(oRequest("ParentID").Item) > 0 Then
				aZoneComponent(S_QUERY_CONDITION_ZONE) = " And (ParentID=" & oRequest("ParentID").Item & ")"
			Else
				aZoneComponent(S_QUERY_CONDITION_ZONE) = " And (ParentID=-1)"
			End If
			If Len(oRequest("ApplyFilter").Item) > 0 Then
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (ZoneName Like '" & S_WILD_CHAR & oRequest("FilterName").Item & S_WILD_CHAR & "')"
			End If
			lErrorNumber = DisplayZonesTable(oRequest, oADODBConnection, DISPLAY_NOTHING, (Len(oRequest("ReadOnly").Item) = 0), False, aZoneComponent, sErrorDescription)
		Case Else
			If Len(oRequest("ApplyFilter").Item) > 0 Then
				aCatalogComponent(S_URL_CATALOG) = aCatalogComponent(S_URL_CATALOG) & "FilterName=" & oRequest("FilterName").Item & "&ApplyFilter=1"
				Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
					Case "Areas"
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And ((AreaShortName Like '" & S_WILD_CHAR & oRequest("FilterName").Item & S_WILD_CHAR & "') Or (AreaName Like '" & S_WILD_CHAR & oRequest("FilterName").Item & S_WILD_CHAR & "') Or (AreaCode Like '" & S_WILD_CHAR & oRequest("FilterName").Item & S_WILD_CHAR & "'))"
					Case "Budgets"
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And ((BudgetShortName Like '" & S_WILD_CHAR & oRequest("FilterName").Item & S_WILD_CHAR & "') Or (BudgetName Like '" & S_WILD_CHAR & oRequest("FilterName").Item & S_WILD_CHAR & "'))"
					Case "Banks"
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (BankName Like '" & S_WILD_CHAR & oRequest("FilterName").Item & S_WILD_CHAR & "')"
					Case "PaperworkOwners"
						'aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) &  " And (PAPERWORKOWNERS.OWNERID = " & oRequest("FilterName").Item &")"
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) &  " And ((PAPERWORKOWNERS.OWNERID = " & oRequest("FilterName").Item & ") OR (PAPERWORKOWNERS.PARENTID = " & oRequest("FilterName").Item & "))"
					Case "PaperworkSenders"
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) &  " And (PAPERWORKSENDERS.SENDERID = " & oRequest("FilterName").Item & ")"
					Case "GeneratingAreas"
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (GeneratingAreaShortName Like '" & S_WILD_CHAR & oRequest("FilterName").Item & S_WILD_CHAR & "')"
					Case "Journeys"
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (JourneyShortName Like '" & S_WILD_CHAR & oRequest("FilterName").Item & S_WILD_CHAR & "')"
					Case "Positions"
						Call GetStartAndEndDatesFromURL("StartForValue", "EndForValue", "Positions.StartDate", False, sCondition)
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & sCondition
						If (oRequest("GroupGradeLevelID").Item) < 0 Then
							aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) &  " And (POSITIONS.POSITIONSHORTNAME Like '" & S_WILD_CHAR & oRequest("PositionShortName").Item & S_WILD_CHAR & "')"
						ElseIf (Len(oRequest("PositionName").Item) = 0) Then
							aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) &  " And (Positions.GroupGradeLevelID = " & oRequest("GroupGradeLevelID").Item &")"
						Else
							aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) &  " And (POSITIONS.POSITIONSHORTNAME Like '" & S_WILD_CHAR & oRequest("PositionShortName").Item & S_WILD_CHAR & "') And (Positions.GroupGradeLevelID = " & oRequest("GroupGradeLevelID").Item &")"
						End If
					Case "Shifts"
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (ShiftShortName Like '" & S_WILD_CHAR & oRequest("FilterName").Item & S_WILD_CHAR & "')"
					Case "States"
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (StateName Like '" & S_WILD_CHAR & oRequest("FilterName").Item & S_WILD_CHAR & "')"
				End Select
			End If
			lErrorNumber = DisplayCatalogsTable(oRequest, oADODBConnection, DISPLAY_NOTHING, (Len(oRequest("ReadOnly").Item) = 0), aCatalogComponent, sErrorDescription)
	End Select

	DisplayTables = lErrorNumber
	Err.Clear
End Function

Function DisplayForms(sAction, sErrorDescription)
'************************************************************
'Purpose: To display the HTML form of the specified catalog.
'Inputs:  sAction
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayForms"
	Dim lErrorNumber
	Dim iIndex
	Dim asNames
	Dim sHolidayDescription

	Select Case sAction
		Case "CurrenciesHistoryList"
			If Len(oRequest("CurrencyDate").Item) > 0 Then
				aCatalogComponent(S_URL_PARAMETERS_CATALOG) = ""
				lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
			End If
		Case "EmployeeFields"
			lErrorNumber = DisplayEmployeeFieldForm(oRequest, oADODBConnection, GetASPFileName(""), aEmployeeFieldComponent, sErrorDescription)
		Case "EmploymentAllowances"
		Case "FormFields"
			lErrorNumber = DisplayFormFieldForm(oRequest, oADODBConnection, GetASPFileName(""), aFormFieldComponent, sErrorDescription)
		Case "Forms"
			lErrorNumber = DisplayFormForm(oRequest, oADODBConnection, GetASPFileName(""), aFormComponent, sErrorDescription)
		Case "Holidays"
			lErrorNumber = GetNameFromTable(oADODBConnection, "Holidays", aCalendarComponent(N_YEAR_CALENDAR), "", ",", aCalendarComponent(S_MARKED_DAYS_CALENDAR), sErrorDescription)
			aCalendarComponent(S_MARKED_DAYS_CALENDAR) = Split(aCalendarComponent(S_MARKED_DAYS_CALENDAR), ",")
			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Días de asueto en " & aCalendarComponent(N_YEAR_CALENDAR) & ":</B><BR />"
				For iIndex = 0 To UBound(aCalendarComponent(S_MARKED_DAYS_CALENDAR))
					Call GetNameFromTable(oADODBConnection, "HolidayDescription", CLng(aCalendarComponent(S_MARKED_DAYS_CALENDAR)(iIndex)), "", "", sHolidayDescription, "")
					Response.Write "&nbsp;&nbsp;&nbsp;" & DisplayDateFromSerialNumber(aCalendarComponent(S_MARKED_DAYS_CALENDAR)(iIndex), -1, -1, -1) & " - " & sHolidayDescription & "<BR />"
				Next
			Response.Write "</FONT>"
		Case "PaperworkLists"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "function AddPaperworkToList() {" & vbNewLine
					Response.Write "var bCorrect = true;" & vbNewLine
					Response.Write "var oForm = document.CatalogFrm;" & vbNewLine
					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "if (oForm.PaperworkNumberTemp.value == '') {" & vbNewLine
							Response.Write "alert('Favor de especificar el número de folio del documento a cerrar');" & vbNewLine
							Response.Write "oForm.PaperworkNumberTemp.focus();" & vbNewLine
							Response.Write "bCorrect = false;" & vbNewLine
						Response.Write "} else {" & vbNewLine
							Response.Write "if (! CheckIntegerValue(oForm.PaperworkNumberTemp, 'el folio del documento a cerrar', N_MINIMUM_ONLY_FLAG, N_OPEN_FLAG, 0, 0))" & vbNewLine
								Response.Write "bCorrect = false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (bCorrect) {" & vbNewLine
							Response.Write "UnselectAllItemsFromList(oForm.PaperworkNumbers);" & vbNewLine
							Response.Write "SelectListItemByValue(oForm.PaperworkNumberTemp.value, true, oForm.PaperworkNumbers);" & vbNewLine
							Response.Write "RemoveSelectedItemsFromList(null, oForm.PaperworkNumbers);" & vbNewLine
							Response.Write "AddItemToList(oForm.PaperworkNumberTemp.value, oForm.PaperworkNumberTemp.value, null, oForm.PaperworkNumbers);" & vbNewLine
							Response.Write "ResizePaperworksToList();" & vbNewLine
							Response.Write "oForm.PaperworkNumberTemp.value = '';" & vbNewLine
							Response.Write "oForm.PaperworkNumberTemp.focus();" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "oForm.PaperworkIDs.value = '';" & vbNewLine
						Response.Write "for (var i=0; i<oForm.PaperworkNumbers.options.length; i++)" & vbNewLine
							Response.Write "oForm.PaperworkIDs.value += oForm.PaperworkNumbers.options[i].text + ',';" & vbNewLine
						Response.Write "if (oForm.PaperworkIDs.value != '')" & vbNewLine
							Response.Write "oForm.PaperworkIDs.value = oForm.PaperworkIDs.value.substr(0, oForm.PaperworkIDs.value.length-1);" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "} // End of AddPaperworkToList" & vbNewLine

				Response.Write "function RemovePaperworkToList() {" & vbNewLine
					Response.Write "var oForm = document.CatalogFrm;" & vbNewLine
					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "RemoveSelectedItemsFromList(null, oForm.PaperworkNumbers);" & vbNewLine
						Response.Write "ResizePaperworksToList();" & vbNewLine
						Response.Write "oForm.PaperworkNumberTemp.focus();" & vbNewLine

						Response.Write "oForm.PaperworkIDs.value = '';" & vbNewLine
						Response.Write "for (var i=0; i<oForm.PaperworkNumbers.options.length; i++)" & vbNewLine
							Response.Write "oForm.PaperworkIDs.value += oForm.PaperworkNumbers.options[i].text + ',';" & vbNewLine
						Response.Write "if (oForm.PaperworkIDs.value != '')" & vbNewLine
							Response.Write "oForm.PaperworkIDs.value = oForm.PaperworkIDs.value.substr(0, oForm.PaperworkIDs.value.length-1);" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "} // End of RemovePaperworkToList" & vbNewLine

				Response.Write "function ResizePaperworksToList() {" & vbNewLine
					Response.Write "var oForm = document.CatalogFrm;" & vbNewLine
					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "if (oForm.PaperworkNumbers.options.length > 3) {" & vbNewLine
							Response.Write "oForm.PaperworkNumbers.size = oForm.PaperworkNumbers.options.length;" & vbNewLine
						Response.Write "} else {" & vbNewLine
							Response.Write "oForm.PaperworkNumbers.size = 3;" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "} // End of ResizePaperworksToList" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine

			aCatalogComponent(S_URL_PARAMETERS_CATALOG) = ""
			lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)

			If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)) > -1 Then
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><A HREF=""javascript: OpenNewWindow('Export.asp?Excel=1&AccessKey=vac2&SIAP_SectionID=2&Action=PaperworkList&ListID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')""><IMG SRC=""Images/IcnPrint.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Imprimir"" BORDER=""0"" HSPACE=""5"" /><B>Imprimir</B></A><BR /></FONT>"

				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "var oForm = document.CatalogFrm;" & vbNewLine
					asNames = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(4), ",")
					For iIndex = 0 To UBound(asNames)
						Response.Write "AddItemToList('" & asNames(iIndex) & "', '" & asNames(iIndex) & "', null, oForm.PaperworkNumbers);" & vbNewLine
					Next
					Response.Write "ResizePaperworksToList();" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If
		Case "PositionsHierarchy"
			aCatalogComponent(S_URL_PARAMETERS_CATALOG) = ""
			If Len(oRequest("ParentJobID").Item) > 0 Then aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " And (AreaCode='" & oRequest("AreaCode").Item & "') And (ParentJobID=" & oRequest("ParentJobID").Item & ")"
			lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)) = -1 Then
					Response.Write "document.CatalogFrm.JobID.value='';" & vbNewLine
				Else
					Response.Write "document.CatalogFrm.AreaCodeOld.value = '" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & "';" & vbNewLine
					Response.Write "sJobID = '" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & "';" & vbNewLine
					Response.Write "sPositionShortName = '" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2) & "';" & vbNewLine
					Response.Write "sGGNShortName = '" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(3) & "';" & vbNewLine
					If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)) = 0 Then
						Response.Write "document.CatalogFrm.EmployeeTypeID.checked = true;" & vbNewLine
						Response.Write "ShowParentFields(true);" & vbNewLine
					End If
				End If
			Response.Write "//--></SCRIPT>" & vbNewLine
		Case "Profiles"
			lErrorNumber = DisplayProfileForm(oRequest, oADODBConnection, GetASPFileName(""), aProfileComponent, sErrorDescription)
		Case "ProfessionalRiskMatrix"
			lErrorNumber = DisplayProfessoionalRiskForm(oRequest, oADODBConnection, sAction, aProfessionalRiskComponent, sErrorDescription)
		Case "Tasks"
			lErrorNumber = DisplayTaskForm(oRequest, oSIAPTACOADODBConnection, GetASPFileName(""), aTaskComponent, sErrorDescription)
		Case "ServicesCenterTypesLKP"
			aCatalogComponent(S_QUERY_CONDITION_CATALOG) = ""
			aCatalogComponent(S_URL_PARAMETERS_CATALOG) = ""
			lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
		Case "TaxInvertions"
		Case "TaxLimits"
		Case "Projects", "TACO_Projects"
			aCatalogComponent(S_URL_PARAMETERS_CATALOG) = ""
			lErrorNumber = DisplayCatalogForm(oRequest, oSIAPTACOADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
		Case "Users"
			lErrorNumber = DisplayUserForm(oRequest, oADODBConnection, GetASPFileName(""), aUserComponent, sErrorDescription)
		Case "Zones"
			lErrorNumber = DisplayZoneForm(oRequest, oADODBConnection, GetASPFileName(""), aZoneComponent, sErrorDescription)
		Case "ZoneTypes"
			lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
			If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)) > -1 Then
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "SelectItemByValue('" & (CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1)) - 2) & "', false, document.CatalogFrm.ZoneTypeID2Temp);" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If
		Case Else
			aCatalogComponent(S_URL_PARAMETERS_CATALOG) = ""
			lErrorNumber = DisplayCatalogForm(oRequest, oADODBConnection, GetASPFileName(""), aCatalogComponent, sErrorDescription)
	End Select

	DisplayForms = lErrorNumber
	Err.Clear
End Function

Function DisplayCurrencyHistoryTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the currencies from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aCurrencyComponent
'Outputs: aCurrencyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayCurrencyHistoryTable"
	Dim iYear
	Dim iMonth
	Dim iIndex
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sBoldBegin
	Dim sBoldEnd
	Dim bForExport
	Dim lErrorNumber

	bForExport = (StrComp(GetASPFileName(""), "Export.asp", vbbinaryCompare) = 0)
	If bForExport Then bUseLinks = False
	If lErrorNumber = 0 Then
		iMonth = Month(Date())
		iYear = Year(Date())
		If Len(oRequest("CurrencyDateYear").Item) > 0 Then
			iMonth = CInt(oRequest("CurrencyDateMonth").Item)
			iYear = CInt(oRequest("CurrencyDateYear").Item)
		Else
			If Len(oRequest("Month").Item) > 0 Then iMonth = CInt(oRequest("Month").Item)
			If Len(oRequest("Year").Item) > 0 Then iYear = CInt(oRequest("Year").Item)
		End If
		sErrorDescription = "No se pudo obtener la información de los registros."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From CurrenciesHistoryList Where (CurrencyID=" & oRequest("CurrencyID").Item & ") And (CurrencyDate>" & iYear & Right(("0" & iMonth), Len("00")) & "00) And (CurrencyDate<" & iYear & Right(("0" & iMonth), Len("00")) & "99) Order By CurrencyDate", "CatalogsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If
	If lErrorNumber = 0 Then
		If Not bForExport Then
			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Historial para el mes de <SELECT NAME=""Month"" ID=""MonthCmb"" SIZE=""1"" CLASS=""Lists"">" & vbNewLine
				For iIndex = 1 To 12
					Response.Write "<OPTION VALUE=""" & Right(("0" & iIndex), Len("00")) & """"
						If iMonth = iIndex Then Response.Write " SELECTED=""1"""
					Response.Write ">" & asMonthNames_es(iIndex) & "</OPTION>" & vbNewLine
				Next
			Response.Write "</SELECT> de "
			Response.Write "<SELECT NAME=""Year"" ID=""YearCmb"" SIZE=""1"" CLASS=""Lists"">" & vbNewLine
				For iIndex = 2000 To Year(Date()) + 1
					Response.Write "<OPTION VALUE=""" & iIndex & """"
						If iYear = iIndex Then Response.Write " SELECTED=""1"""
					Response.Write ">" & iIndex & "</OPTION>" & vbNewLine
				Next
			Response.Write "</SELECT>.</FONT>"
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""40"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Ver Historial"" CLASS=""BUTTONS"" onClick=""window.location.href='" & GetASPFileName("") & "?" & RemoveParameterFromURLString(RemoveParameterFromURLString(oRequest, "Month"), "Year") & "&Month=' + window.document.all['MonthCmb'].value + '&Year=' + window.document.all['YearCmb'].value;"" /><BR /><BR />" & vbNewLine
		End If
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""390"" BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks Then
					asColumnsTitles = Split("&nbsp;,Fecha,Valor,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,210,80,80", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("&nbsp;,Fecha,Valor", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,290,80", ",", -1, vbBinaryCompare)
				End If
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split(",,RIGHT,CENTER", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("CurrencyDate").Value), oRequest("CurrencyDate").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""CurrencyDate"" ID=""CurrencyDateRd"" VALUE=""" & CStr(oRecordset.Fields("CurrencyDate").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""CurrencyDate"" ID=""CurrencyDateChk"" VALUE=""" & CStr(oRecordset.Fields("CurrencyDate").Value) & """ />"
						Case Else
							sRowContents = sRowContents & "&nbsp;"
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & DisplayDateFromSerialNumber(CStr(oRecordset.Fields("CurrencyDate").Value), -1, -1, -1) & sBoldEnd & "</A>"
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & FormatNumber(CStr(oRecordset.Fields("CurrencyValue").Value), 4, True, False, True) & sBoldEnd
					If bUseLinks Then
						sRowContents = sRowContents & TABLE_SEPARATOR
						sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "&CurrencyID=" & CStr(oRecordset.Fields("CurrencyID").Value) & "&CurrencyDate=" & CStr(oRecordset.Fields("CurrencyDate").Value) & "&Year=" & oRequest("Year").Item &  "&Month=" & oRequest("Month").Item &  "&Change=1"">"
							sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
						sRowContents = sRowContents & "</A>"
						'sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;<A HREF=""http://www.oanda.com/convert/classic?script=../convert/classic&LANGUAGE=en&lang=en&value=1&date=" & Mid(CStr(oRecordset.Fields("CurrencyDate").Value), 5, 2) & "/" & Mid(CStr(oRecordset.Fields("CurrencyDate").Value), 7, 2) & "/" & Mid(CStr(oRecordset.Fields("CurrencyDate").Value), 3, 2) & "&date_fmt=us" & "&exch=" & aCurrencyComponent(S_KEY_CURRENCY) & "&expr=MXN&margin_fixed=0"" TARGET=""oanda""><IMG SRC=""Images/BtnCurrency.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Obtener tipo de cambio"" BORDER=""0"" /></A>"
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen monedas registradas en la base de datos."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayCurrencyHistoryTable = lErrorNumber
	Err.Clear
End Function


Function VerifyExistencyOfRecordsInPeriodDates(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific record exists in the database
'Inputs:  oADODBConnection, aCatalogComponent
'Outputs: aCatalogComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistencyOfRecordsInPeriodDates"
	Dim sQuery
	Dim sCondition
	Dim iIndex
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCatalogComponent(B_COMPONENT_INITIALIZED_CATALOG)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCatalogComponent(oRequest, aCatalogComponent)
	End If

	Select Case CStr(oRequest("Action").Item)
		Case "EmployeesAntiquitiesLKP"
			If VerifyPeriodDatesForConcepts(CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1)), CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2)), sErrorDescription) Then
				If (Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)) = 0) And (Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1)) = 0) And (Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2)) = 0) Then
					lErrorNumber = -1
					sErrorDescription = "No se especificó el número de empleado, la fecha de inicio, ni la fecha de fin para revisar su existencia en la base de datos."
					Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogsLib.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
				Else
					sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
					sQuery = "Select * from EmployeesAntiquitiesLKP" & _
							 " Where (EmployeeID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & ")" & _
							 " And (((AntiquityDate >=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (AntiquityDate <=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2) & "))" & _
							 " Or ((EndDate >=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EndDate <=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2) & "))" & _
							 " Or ((EndDate >=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (AntiquityDate <=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2) & ")))"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "CatalogsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						aCatalogComponent(B_IS_DUPLICATED_CATALOG) = (Not oRecordset.EOF)
					End If
				End If
			Else
				lErrorNumber = -1
			End If
		Case Else
	End Select

	oRecordset.Close
	Set oRecordset = Nothing

	VerifyExistencyOfRecordsInPeriodDates = lErrorNumber
	Err.Clear
End Function

Function VerifyPeriodDatesForAntiquities(lStartDate, lEndDate, lYear, lMonth, lDay, sErrorDescription)
'************************************************************
'Purpose: To check if a specific record is concisten
'Inputs:  lStartDate, lEndDate
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyPeriodDatesForAntiquities"
	Dim lDiferenceYear
	Dim lDiferenceMonth
	Dim lDiferenceDay

	Call GetAntiquityFromSerialDates(lStartDate, lEndDate, lDiferenceYear, lDiferenceMonth, lDiferenceDay)

	lDiferenceYear = Right("0" & lDiferenceYear, Len("00"))
	lDiferenceMonth = Right("0" & lDiferenceMonth, Len("00"))
	lDiferenceDay = Right("0" & lDiferenceDay, Len("00"))
	lYear = Right("0" & lYear, Len("00"))
	lMonth = Right("0" & lMonth, Len("00"))
	lDay = Right("0" & lDay, Len("00"))

	If (CInt(lDiferenceYear & lDiferenceMonth & lDiferenceDay) > CInt(lYear & lMonth & lDay)) Then
		VerifyPeriodDatesForAntiquities = True
	Else
		sErrorDescription = "Los días, meses y años indicados no pueden exederse respecto al periodo de las fechas especificadas."
		VerifyPeriodDatesForAntiquities = False
	End If

	Err.Clear
End Function
%>