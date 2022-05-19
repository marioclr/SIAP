<%
Const S_TABLE_NAME_CATALOG = 0
Const N_ID_CATALOG = 1
Const S_NAME_CATALOG = 2
Const N_NAME_CATALOG = 3
Const AS_FIELDS_TEXTS_CATALOG = 4
Const AS_FIELDS_NAMES_CATALOG = 5
Const AS_FIELDS_REQUIRED_CATALOG = 6
Const AS_FIELDS_TYPES_CATALOG = 7
Const AS_FIELDS_SIZES_CATALOG = 8
Const AS_FIELDS_LIMITS_CATALOG = 9
Const AS_FIELDS_MINIMUMS_CATALOG = 10
Const AS_FIELDS_MAXIMUMS_CATALOG = 11
Const AS_FIELDS_VALUES_CATALOG = 12
Const AS_DEFAULT_VALUES_CATALOG = 13
Const AS_CATALOG_PARAMETERS_CATALOG = 14
Const AS_SCRIPT_CATALOG = 15
Const N_ACTIVE_CATALOG = 16
Const AS_FIELDS_TO_SHOW_CATALOG = 17
Const S_FIELDS_TO_BLOCK_CATALOG = 18
Const S_FIELDS_TO_SUM_CATALOG = 19
Const S_ADD_LINES_BEFORE_FIELDS_CATALOG = 20
Const S_URL_CATALOG = 21
Const S_URL_PARAMETERS_CATALOG = 22
Const S_ADDITIONAL_FORM_HTML_CATALOG = 23
Const S_ADDITIONAL_FORM_SCRIPT_CATALOG = 24
Const N_START_FIELD_FOR_HISTORY_LIST_CATALOG = 25
Const N_END_FIELD_FOR_HISTORY_LIST_CATALOG = 26
Const S_IDS_NOT_UPDATABLE_CATALOG = 27
Const B_MODIFY_CATALOG = 28
Const B_DELETE_CATALOG = 29
Const B_ACTIVE_CATALOG = 30
Const S_SHOW_LINKS_FOR_IDS_CATALOG = 31
Const B_SHOW_ID_FIELD_CATALOG = 32
Const B_FORCE_SHOW_MODIFY_BUTTON_CATALOG = 33
Const B_FORCE_SHOW_DELETE_BUTTON_CATALOG = 34
Const S_CANCEL_BUTTON_ACTION_CATALOG = 35
Const S_EXTRA_BUTTON_CATALOG = 36
Const S_FULL_QUERY_CATALOG = 37
Const S_QUERY_CONDITION_CATALOG = 38
Const S_CHECK_EXISTENCY_CONDITION_CATALOG = 39
Const S_ORDER_CATALOG = 40
Const B_CHECK_FOR_DUPLICATED_CATALOG = 41
Const B_IS_DUPLICATED_CATALOG = 42
Const S_FORM_NAME_CATALOG = 43
Const B_COMPONENT_INITIALIZED_CATALOG = 44

Const N_CATALOG_COMPONENT_SIZE = 44

Dim aCatalogComponent()
Redim aCatalogComponent(N_CATALOG_COMPONENT_SIZE)

Function InitializeCatalogComponent(oRequest, aCatalogComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Catalog Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aCatalogComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeCatalogComponent"
	Redim Preserve aCatalogComponent(N_CATALOG_COMPONENT_SIZE)
	Dim iIndex

	If IsEmpty(aCatalogComponent(S_TABLE_NAME_CATALOG)) Then
		If Len(oRequest("TableName").Item) > 0 Then
			aCatalogComponent(S_TABLE_NAME_CATALOG) = oRequest("TableName").Item
		ElseIf Len(oRequest("Action").Item) > 0 Then
			aCatalogComponent(S_TABLE_NAME_CATALOG) = oRequest("Action").Item
		ElseIf Len(oRequest("RegistryType").Item) > 0 Then
			aCatalogComponent(S_TABLE_NAME_CATALOG) = oRequest("RegistryType").Item
		Else
			aCatalogComponent(S_TABLE_NAME_CATALOG) = ""
		End If
	End If

	If IsEmpty(aCatalogComponent(N_ID_CATALOG)) Then
		If Len(oRequest("FieldID").Item) > 0 Then
			aCatalogComponent(N_ID_CATALOG) = CInt(oRequest("FieldID").Item)
		Else
			aCatalogComponent(N_ID_CATALOG) = 0
		End If
	End If

	If IsEmpty(aCatalogComponent(N_NAME_CATALOG)) Then
		If Len(oRequest("NameFieldID").Item) > 0 Then
			aCatalogComponent(N_NAME_CATALOG) = CInt(oRequest("NameFieldID").Item)
		Else
			aCatalogComponent(N_NAME_CATALOG) = aCatalogComponent(N_ID_CATALOG) + 1
		End If
	End If

	If IsEmpty(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)) Then
		If Len(oRequest("FieldsTexts").Item) > 0 Then
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = oRequest("FieldsTexts").Item
		Else
			aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = ""
		End If
	End If
	If Not IsArray(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)) Then aCatalogComponent(AS_FIELDS_TEXTS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG), ",")

	If IsEmpty(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)) Then
		If Len(oRequest("FieldsNames").Item) > 0 Then
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = oRequest("FieldsNames").Item
		Else
			aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = ""
		End If
	End If
	If Not IsArray(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)) Then aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")

	If IsEmpty(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG)) Then
		If Len(oRequest("FieldsRequired").Item) > 0 Then
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = oRequest("FieldsRequired").Item
		Else
			aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = ""
		End If
	End If
	If Not IsArray(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG)) Then aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG) = Split(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG), ",")

	If IsEmpty(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)) Then
		If Len(oRequest("FieldsTypes").Item) > 0 Then
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = oRequest("FieldsTypes").Item
		Else
			aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = ""
		End If
	End If
	If Not IsArray(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)) Then aCatalogComponent(AS_FIELDS_TYPES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_TYPES_CATALOG), ",")

	If IsEmpty(aCatalogComponent(AS_FIELDS_SIZES_CATALOG)) Then
		If Len(oRequest("FieldsSizes").Item) > 0 Then
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = oRequest("FieldsSizes").Item
		Else
			aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = ""
		End If
	End If
	If Not IsArray(aCatalogComponent(AS_FIELDS_SIZES_CATALOG)) Then aCatalogComponent(AS_FIELDS_SIZES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_SIZES_CATALOG), ",")

	If IsEmpty(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG)) Then
		If Len(oRequest("FieldsLimits").Item) > 0 Then
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = oRequest("FieldsLimits").Item
		Else
			aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = ""
		End If
	End If
	If Not IsArray(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG)) Then aCatalogComponent(AS_FIELDS_LIMITS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG), ",")

	If IsEmpty(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)) Then
		If Len(oRequest("FieldsMinimums").Item) > 0 Then
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = oRequest("FieldsMinimums").Item
		Else
			aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = ""
		End If
	End If
	If Not IsArray(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)) Then aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG), ",")

	If IsEmpty(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)) Then
		If Len(oRequest("FieldsMaximums").Item) > 0 Then
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = oRequest("FieldsMaximums").Item
		Else
			aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = ""
		End If
	End If
	If Not IsArray(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)) Then aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG) = Split(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG), ",")

	If IsEmpty(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)) Then
		If Len(oRequest("FieldsValues").Item) > 0 Then
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = oRequest("FieldsValues").Item
		Else
			For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
				If CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex)) = N_BOOLEAN Then
					If Len(oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex)).Item) = 0 Then
						aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG) & "0" & CATALOG_SEPARATOR
					Else
						aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG) & oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex)).Item & CATALOG_SEPARATOR
					End If
				Else
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG) & oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex)).Item & CATALOG_SEPARATOR
				End If
			Next
			If Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)) > 0 Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Left(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), (Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)) - Len(CATALOG_SEPARATOR)))
		End If
	End If
	If Not IsArray(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)) Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), CATALOG_SEPARATOR)
	If aCatalogComponent(N_ID_CATALOG) > -1 Then
		If Len(oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Item) > 0 Then
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) = oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Item
		End If
	End If

	If IsEmpty(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG)) Then
		If Len(oRequest("DefaultValues").Item) > 0 Then
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = oRequest("FieldsMaximums").Item
		Else
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = ""
		End If
	End If
	If Not IsArray(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG)) Then aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")

	If IsEmpty(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)) Then
		If Len(oRequest("CatalogParameters").Item) > 0 Then
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = oRequest("CatalogParameters").Item
		Else
			aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = ""
		End If
	End If
	If Not IsArray(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)) Then aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)

	If IsEmpty(aCatalogComponent(AS_SCRIPT_CATALOG)) Then
		If Len(oRequest("CatalogScripts").Item) > 0 Then
			aCatalogComponent(AS_SCRIPT_CATALOG) = oRequest("CatalogScripts").Item
		Else
			aCatalogComponent(AS_SCRIPT_CATALOG) = ""
		End If
	End If
	If Not IsArray(aCatalogComponent(AS_SCRIPT_CATALOG)) Then aCatalogComponent(AS_SCRIPT_CATALOG) = Split(aCatalogComponent(AS_SCRIPT_CATALOG), CATALOG_SEPARATOR)

	If IsEmpty(aCatalogComponent(N_ACTIVE_CATALOG)) Then
		If Len(oRequest("ActiveFieldID").Item) > 0 Then
			aCatalogComponent(N_ACTIVE_CATALOG) = CInt(oRequest("ActiveFieldID").Item)
		ElseIf IsArray(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)) Then
			aCatalogComponent(N_ACTIVE_CATALOG) = UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
		Else
			aCatalogComponent(N_ACTIVE_CATALOG) = -1
		End If
	End If

	aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG) = Split("1", ",")
	aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) = ""
	aCatalogComponent(S_FIELDS_TO_SUM_CATALOG) = ""
	aCatalogComponent(S_ADD_LINES_BEFORE_FIELDS_CATALOG) = ""
	aCatalogComponent(S_URL_CATALOG) = ""
	aCatalogComponent(S_URL_PARAMETERS_CATALOG) = ""
	aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) = ""
	aCatalogComponent(S_ADDITIONAL_FORM_SCRIPT_CATALOG) = ""
	aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = -1
	aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = -1
	aCatalogComponent(S_IDS_NOT_UPDATABLE_CATALOG) = ""
	aCatalogComponent(B_MODIFY_CATALOG) = ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS)
	aCatalogComponent(B_DELETE_CATALOG) = B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)
	aCatalogComponent(B_ACTIVE_CATALOG) = aCatalogComponent(B_MODIFY_CATALOG)
	aCatalogComponent(S_SHOW_LINKS_FOR_IDS_CATALOG) = ""
	aCatalogComponent(B_SHOW_ID_FIELD_CATALOG) = False
	aCatalogComponent(B_FORCE_SHOW_MODIFY_BUTTON_CATALOG) = False
	aCatalogComponent(B_FORCE_SHOW_DELETE_BUTTON_CATALOG) = False
	aCatalogComponent(S_CANCEL_BUTTON_ACTION_CATALOG) = "window.location.href='" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "'"
	aCatalogComponent(S_EXTRA_BUTTON_CATALOG) = ""
	aCatalogComponent(S_FULL_QUERY_CATALOG) = ""
	aCatalogComponent(S_QUERY_CONDITION_CATALOG) = ""
	aCatalogComponent(S_CHECK_EXISTENCY_CONDITION_CATALOG) = ""
	aCatalogComponent(S_ORDER_CATALOG) = ""
	aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) = True
	aCatalogComponent(B_IS_DUPLICATED_CATALOG) = False

	aCatalogComponent(S_FORM_NAME_CATALOG) = "CatalogFrm"
	aCatalogComponent(B_COMPONENT_INITIALIZED_CATALOG) = True
	InitializeCatalogComponent = Err.number
	Err.Clear
End Function

Function InitializeValuesForCatalogComponent(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To initialize the Values for the Catalog Component
'         using the URL parameters
'Inputs:  oRequest
'Outputs: aCatalogComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeValuesForCatalogComponent"
	Dim iIndex

	aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = ""
	For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
		If (IsEmpty(oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex)).Item)) And (IsEmpty(oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Year").Item)) And (IsEmpty(oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Hour").Item)) Then
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG) & aCatalogComponent(AS_DEFAULT_VALUES_CATALOG)(iIndex) & CATALOG_SEPARATOR
		ElseIf CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex)) = N_BOOLEAN Then
			If Len(oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex)).Item) = 0 Then
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG) & "0" & CATALOG_SEPARATOR
			Else
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG) & oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex)).Item & CATALOG_SEPARATOR
			End If
		ElseIf CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex)) = N_DATE Then
			If Len(oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex)).Item) = 0 Then
				If Len(oRequest(Replace(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), "Date", "Year")).Item) = 0 Then
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG) & oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Year").Item & Right(("0" & oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Month").Item), Len("00")) & Right(("0" & oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Day").Item), Len("00")) & CATALOG_SEPARATOR
				Else
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG) & oRequest(Replace(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), "Date", "Year")).Item & Right(("0" & oRequest(Replace(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), "Date", "Month")).Item), Len("00")) & Right(("0" & oRequest(Replace(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), "Date", "Day")).Item), Len("00")) & CATALOG_SEPARATOR
				End If
			Else
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG) & oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex)).Item & CATALOG_SEPARATOR
			End If
		ElseIf CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex)) = N_HOUR Then
			If Len(oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex)).Item) = 0 Then
				If Len(oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Minute").Item) > 0 Then
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG) & oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Hour").Item & Right(("0" & oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Minute").Item), Len("00")) & CATALOG_SEPARATOR
				ElseIf Len(oRequest(Replace(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), "Hour", "Minute")).Item) = 0 Then
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG) & oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Hour").Item & Right(("0" & oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Minute").Item), Len("00")) & CATALOG_SEPARATOR
				Else
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG) & oRequest(Replace(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), "Date", "Hour")).Item & Right(("0" & oRequest(Replace(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), "Date", "Minute")).Item), Len("00")) & CATALOG_SEPARATOR
				End If
			Else
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG) & oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex)).Item & CATALOG_SEPARATOR
			End If
		ElseIf Len(oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex)).Item) > 0 Then
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG) & oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex)).Item & CATALOG_SEPARATOR
		Else
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG) & aCatalogComponent(AS_DEFAULT_VALUES_CATALOG)(iIndex) & CATALOG_SEPARATOR
		End If
	Next
	If Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)) > 0 Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Left(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), (Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)) - Len(CATALOG_SEPARATOR)))
	If Not IsArray(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)) Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), CATALOG_SEPARATOR)

	InitializeValuesForCatalogComponent = Err.number
	Err.Clear
End Function

Function AddCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new record into the given catalog
'Inputs:  oRequest, oADODBConnection
'Outputs: aCatalogComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddCatalog"
	Dim sNames
	Dim sValues
	Dim iIndex
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sCompaniesNames
	Dim sGroupGradeLevelNames
	Dim sEmployeeTypesNames

	bComponentInitialized = aCatalogComponent(B_COMPONENT_INITIALIZED_CATALOG)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCatalogComponent(oRequest, aCatalogComponent)
	End If

	If Len(aCatalogComponent(S_TABLE_NAME_CATALOG)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se pudo agregar el registro pues no se especificó el nombre de la tabla."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aCatalogComponent(N_ID_CATALOG) > -1 Then
			If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) < 0 Then
				sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
				lErrorNumber = GetNewIDFromTable(oADODBConnection, aCatalogComponent(S_TABLE_NAME_CATALOG), aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)), "", 1, aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)), sErrorDescription)
			End If
		End If
		If aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) Then
			Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
				Case "Positions"
					lErrorNumber = CheckExistencyOfPositionInCatalog(oADODBConnection, aCatalogComponent, sErrorDescription)
				Case Else
					lErrorNumber = CheckExistencyOfRecordInCatalog(oADODBConnection, aCatalogComponent, sErrorDescription)
			End Select
			If aCatalogComponent(B_IS_DUPLICATED_CATALOG) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
					Case "Positions"
						Call GetNameFromTable(oADODBConnection, "Companies", aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(9), "", "", sCompaniesNames, sErrorDescription)
						Call GetNameFromTable(oADODBConnection, "GroupGradeLevels", aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11), "", "", sGroupGradeLevelNames, sErrorDescription)
						Call GetNameFromTable(oADODBConnection, "EmployeeTypes", aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(17), "", "", sEmployeeTypesNames, sErrorDescription)
						sErrorDescription = "Ya existe un puesto registrado con la clave '" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & "' para la compañía '" & sCompaniesNames & "' clasificación '" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10) & "' grupo-grado-nivel '" & sGroupGradeLevelNames & "' integración '" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12) & "' nivel '" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13) & "' tabulador '" & sEmployeeTypesNames & "' jornada '" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(18) & "' en el periódo indicado."
					Case Else
						If aCatalogComponent(N_ID_CATALOG) > -1 Then
							sErrorDescription = "Ya existe un registro con " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(aCatalogComponent(N_NAME_CATALOG)) & " '" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_NAME_CATALOG)) & "' registrado en el catálogo."
						Else
							sErrorDescription = "Ya existe un registro con estas características dado de alta en el sistema."
						End If
				End Select
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			End If
		End If

		Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
			Case "Paperworks"
				If lErrorNumber = 0 Then
					If Not VerifyOwnersRelationship(oRequest, oADODBConnection, 1, sErrorDescription) Then
						lErrorNumber = -1
					End If
				End If
		End Select

		If lErrorNumber = 0 Then
			If Not CheckCatalogInformationConsistency(aCatalogComponent, sErrorDescription) Then
				lErrorNumber = -1
			Else
				For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
					If CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex)) <> N_FILE Then
						sNames = sNames & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & ", "
						sValues = sValues & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) & "', '"
					End If
				Next
				If Len(sNames) > 0 Then sNames = Left(sNames, (Len(sNames) - Len(", ")))
				If Len(sValues) > 0 Then sValues = Left(sValues, (Len(sValues) - Len("', '")))
				sErrorDescription = "No se pudo guardar la información del nuevo registro."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " (" & sNames & ") Values ('" & sValues & "')", "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
	End If

	AddCatalog = lErrorNumber
	Err.Clear
End Function

Function GetCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To get the record information from the given catalog
'Inputs:  oRequest, oADODBConnection
'Outputs: aCatalogComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetCatalog"
	Dim oRecordset
	Dim iIndex
	Dim sTemp
	Dim lErrorNumber
	Dim sFieldsToSelect
	Dim bComponentInitialized

	bComponentInitialized = aCatalogComponent(B_COMPONENT_INITIALIZED_CATALOG)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCatalogComponent(oRequest, aCatalogComponent)
	End If

	For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
		sFieldsToSelect = sFieldsToSelect & CStr(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex)) & ", "
	Next
	If Len(sFieldsToSelect) > 0 Then sFieldsToSelect  = Left(sFieldsToSelect , (Len(sFieldsToSelect ) - Len(", ")))

	If Len(aCatalogComponent(S_TABLE_NAME_CATALOG)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se pudo obtener la información el registro pues no se especificó el nombre de la tabla."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	ElseIf aCatalogComponent(N_ID_CATALOG) = -1 Then
		aCatalogComponent(S_QUERY_CONDITION_CATALOG) = Trim(aCatalogComponent(S_QUERY_CONDITION_CATALOG))
		If Len(aCatalogComponent(S_QUERY_CONDITION_CATALOG)) > 0 Then
			If InStr(1, aCatalogComponent(S_QUERY_CONDITION_CATALOG), "And ", vbTextCompare) = 1 Then
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = Right(aCatalogComponent(S_QUERY_CONDITION_CATALOG), (Len(aCatalogComponent(S_QUERY_CONDITION_CATALOG)) - Len("And ")))
			End If
			aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
		End If
		If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) And (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) Then aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & ")"
		If Len(aCatalogComponent(S_QUERY_CONDITION_CATALOG)) > 0 Then
			sErrorDescription = "No se pudo obtener la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & sFieldsToSelect & " From " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Where " & aCatalogComponent(S_QUERY_CONDITION_CATALOG), "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If oRecordset.EOF Then
					lErrorNumber = L_ERR_NO_RECORDS
					sErrorDescription = "El registro especificado no se encuentra en el sistema."
					Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
				Else
					aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = ""
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = ""
					For iIndex = 0 To oRecordset.Fields.Count - 1
						aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = aCatalogComponent(AS_FIELDS_NAMES_CATALOG) & oRecordset.Fields(iIndex).Name & ","
						sTemp = ""
						sTemp = CStr(oRecordset.Fields(iIndex).Value)
						If Err.number <> 0 Then Err.Clear
						aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG) & sTemp & CATALOG_SEPARATOR
					Next
					If Len(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)) > 0 Then aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Left(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), (Len(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)) - Len(",")))
					If Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)) > 0 Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Left(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), (Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)) - Len(CATALOG_SEPARATOR)))
					aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), CATALOG_SEPARATOR)
				End If
				oRecordset.Close
			End If
		End If
	ElseIf CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG))) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		aCatalogComponent(S_QUERY_CONDITION_CATALOG) = Trim(aCatalogComponent(S_QUERY_CONDITION_CATALOG))
		If Len(aCatalogComponent(S_QUERY_CONDITION_CATALOG)) > 0 Then
			If InStr(1, aCatalogComponent(S_QUERY_CONDITION_CATALOG), "And ", vbTextCompare) <> 1 Then
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = "And " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
			End If
			aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
		End If
		If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) And (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) Then aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & ")"
		sErrorDescription = "No se pudo obtener la información del registro."
		If aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) = N_TEXT Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & sFieldsToSelect & " From " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Where (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "='" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "') " & aCatalogComponent(S_QUERY_CONDITION_CATALOG), "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & sFieldsToSelect & " From " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Where (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & ") " & aCatalogComponent(S_QUERY_CONDITION_CATALOG), "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		End If
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = ""
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = ""
				For iIndex = 0 To oRecordset.Fields.Count - 1
					aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = aCatalogComponent(AS_FIELDS_NAMES_CATALOG) & oRecordset.Fields(iIndex).Name & ","
					sTemp = ""
					sTemp = CStr(oRecordset.Fields(iIndex).Value)
					If Err.number <> 0 Then Err.Clear
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG) & sTemp & CATALOG_SEPARATOR
				Next
				If Len(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)) > 0 Then aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Left(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), (Len(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)) - Len(",")))
				If Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)) > 0 Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Left(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), (Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)) - Len(CATALOG_SEPARATOR)))
				aCatalogComponent(AS_FIELDS_NAMES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_NAMES_CATALOG), ",")
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), CATALOG_SEPARATOR)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetCatalog = lErrorNumber
	Err.Clear
End Function

Function GetCatalogs(oRequest, oADODBConnection, aCatalogComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the records from
'		  the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aCatalogComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetCatalogs"
	Dim sFields
	Dim sTables
	Dim sJoinCondition
	Dim asTemp
	Dim sTemp
	Dim iIndex
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim bShowQueryAsHidden

	bComponentInitialized = aCatalogComponent(B_COMPONENT_INITIALIZED_CATALOG)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCatalogComponent(oRequest, aCatalogComponent)
	End If

	If Len(aCatalogComponent(S_TABLE_NAME_CATALOG)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se pudo obtener la información el registro pues no se especificó el nombre de la tabla."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	ElseIf Len(aCatalogComponent(S_FULL_QUERY_CATALOG)) > 0 Then
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, aCatalogComponent(S_FULL_QUERY_CATALOG), "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Else
		aCatalogComponent(S_QUERY_CONDITION_CATALOG) = Trim(aCatalogComponent(S_QUERY_CONDITION_CATALOG))
		If Len(aCatalogComponent(S_QUERY_CONDITION_CATALOG)) > 0 Then
			If InStr(1, aCatalogComponent(S_QUERY_CONDITION_CATALOG), "And (PositionsSpecialJourneysLKP.Active<=0)", vbTextCompare) > 0 Then
				bShowQueryAsHidden = True
			End If
			If InStr(1, aCatalogComponent(S_QUERY_CONDITION_CATALOG), "And ", vbTextCompare) <> 1 Then
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = "And " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
			End If
			aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
		End If
		sFields = ""
		sTables = ""
		sJoinCondition = ""
		For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_TYPES_CATALOG))
			If InStr(1, ",6,7,", "," & aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex) & ",") > 0 Then
				If InStr(1, aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex), "<OPTION", vbBinaryCompare) = 0 Then
					If IsArray(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)) Then
						asTemp = aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)
					Else
						asTemp = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex), SECOND_LIST_SEPARATOR)
					End If
					sTemp = asTemp(0)
					If InStr(1, asTemp(0), " As ", vbBinaryCompare) > 0 Then
						sTemp = Split(asTemp(0), " ")
						sFields = sFields & ", " & sTemp(UBound(sTemp)) & "." & asTemp(2)
						sJoinCondition = sJoinCondition & " And (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "=" & sTemp(UBound(sTemp)) & "." & asTemp(1) & ")"
					Else
						sFields = sFields & ", " & asTemp(0) & "." & asTemp(2)
						sJoinCondition = sJoinCondition & " And (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "=" & asTemp(0) & "." & asTemp(1) & ")"
					End If
					sTables = sTables & ", " & asTemp(0)
				End If
			End If
		Next
		sErrorDescription = "No se pudieron obtener los registros del catálogo."
		If aCatalogComponent(N_ID_CATALOG) > -1 Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & aCatalogComponent(S_TABLE_NAME_CATALOG) & ".*" & sFields & " From " & aCatalogComponent(S_TABLE_NAME_CATALOG) & sTables & " Where (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & " > -1) " & sJoinCondition & aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " Order By " & aCatalogComponent(S_ORDER_CATALOG), "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If bShowQueryAsHidden Then
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & "Select " & aCatalogComponent(S_TABLE_NAME_CATALOG) & ".*" & sFields & " From " & aCatalogComponent(S_TABLE_NAME_CATALOG) & sTables & " Where (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & " > -1) " & sJoinCondition & aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " Order By " & aCatalogComponent(S_ORDER_CATALOG) & """ />"
			End If
		Else
			sJoinCondition = Trim(sJoinCondition & aCatalogComponent(S_QUERY_CONDITION_CATALOG))
			If InStr(1, sJoinCondition, "And ", vbBinaryCompare) = 1 Then sJoinCondition = Right(sJoinCondition, (Len(sJoinCondition) - Len("And ")))
			If Len(sJoinCondition) > 0 Then sJoinCondition = " Where " & sJoinCondition
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & aCatalogComponent(S_TABLE_NAME_CATALOG) & ".*" & sFields & " From " & aCatalogComponent(S_TABLE_NAME_CATALOG) & sTables & sJoinCondition & " Order By " & aCatalogComponent(S_ORDER_CATALOG), "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If bShowQueryAsHidden Then
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & "Select " & aCatalogComponent(S_TABLE_NAME_CATALOG) & ".*" & sFields & " From " & aCatalogComponent(S_TABLE_NAME_CATALOG) & sTables & sJoinCondition & " Order By " & aCatalogComponent(S_ORDER_CATALOG) & """ />"
			End If
		End If
	End If

	GetCatalogs = lErrorNumber
	Err.Clear
End Function

Function ModifyCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing record in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aCatalogComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyCatalog"
	Dim sQuery
	Dim sNames
	Dim sValues
	Dim iIndex
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCatalogComponent(B_COMPONENT_INITIALIZED_CATALOG)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCatalogComponent(oRequest, aCatalogComponent)
	End If

	If Len(aCatalogComponent(S_TABLE_NAME_CATALOG)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se pudo modificar la información el registro pues no se especificó el nombre de la tabla."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	ElseIf aCatalogComponent(N_ID_CATALOG) = -1 Then
		If lErrorNumber = 0 Then
			aCatalogComponent(S_QUERY_CONDITION_CATALOG) = Trim(aCatalogComponent(S_QUERY_CONDITION_CATALOG))
			If Len(aCatalogComponent(S_QUERY_CONDITION_CATALOG)) > 0 Then
				If InStr(1, aCatalogComponent(S_QUERY_CONDITION_CATALOG), "And ", vbTextCompare) = 1 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = Right(aCatalogComponent(S_QUERY_CONDITION_CATALOG), (Len(aCatalogComponent(S_QUERY_CONDITION_CATALOG)) - Len("And ")))
				End If
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
			End If

			If Not CheckCatalogInformationConsistency(aCatalogComponent, sErrorDescription) Then
				lErrorNumber = -1
			Else
				If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = -1) Or (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = -1) Then
					sQuery = "Update " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Set "
					For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
						If CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex)) <> N_FILE Then
							sQuery = sQuery & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "='" & Replace(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex), "'", "´") & "', "
						End If
					Next
					sQuery = Left(sQuery, (Len(sQuery) - Len(", ")))
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = Trim(aCatalogComponent(S_QUERY_CONDITION_CATALOG))
					sQuery = sQuery & " Where " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
					sErrorDescription = "No se pudo modificar la información del registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				Else
					If StrComp(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)), oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & "Old").Item, vbBinaryCompare) = 0 Then
						sQuery = "Update " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Set "
						For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
							If CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex)) <> N_FILE Then
								sQuery = sQuery & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "='" & Replace(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex), "'", "´") & "', "
							End If
						Next
						sQuery = Left(sQuery, (Len(sQuery) - Len(", ")))
						sQuery = sQuery & " Where (" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & "=30000000) " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
						sErrorDescription = "No se pudo modificar la información del registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Else
						sQuery = "Update " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Set " & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & "=" & AddDaysToSerialDate(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)), -1) & " "
						If aCatalogComponent(N_ACTIVE_CATALOG) > -1 Then sQuery = sQuery & ", " & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ACTIVE_CATALOG)) & "=0 "
						sQuery = sQuery & " Where (" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & "=30000000) " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
						sErrorDescription = "No se pudo modificar la información del registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

						sNames = ""
						sValues = ""
						For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
							If CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex)) <> N_FILE Then
								If iIndex = N_END_FIELD_FOR_HISTORY_LIST_CATALOG Then
									sNames = sNames & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & ", "
									sValues = sValues & "30000000', '"
								Else
									sNames = sNames & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & ", "
									sValues = sValues & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) & "', '"
								End If
							End If
						Next
						If Len(sNames) > 0 Then sNames = Left(sNames, (Len(sNames) - Len(", ")))
						If Len(sValues) > 0 Then sValues = Left(sValues, (Len(sValues) - Len("', '")))
						sErrorDescription = "No se pudo guardar la información del nuevo registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " (" & sNames & ") Values ('" & sValues & "')", "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
			End If
		End If
	ElseIf CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG))) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para modificar su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aCatalogComponent(B_CHECK_FOR_DUPLICATED_CATALOG) Then
			lErrorNumber = CheckExistencyOfRecordInCatalog(oADODBConnection, aCatalogComponent, sErrorDescription)
			If aCatalogComponent(B_IS_DUPLICATED_CATALOG) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				If aCatalogComponent(N_ID_CATALOG) > -1 Then
					sErrorDescription = "Ya existe un registro con " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(aCatalogComponent(N_NAME_CATALOG)) & " '" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_NAME_CATALOG)) & "' registrado en el catálogo."
				Else
					sErrorDescription = "Ya existe un registro con estas características dado de alta en el sistema."
				End If
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			End If
		End If

		Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
			Case "Paperworks"
				If Not VerifyOwnersRelationship(oRequest, oADODBConnection, 1, sErrorDescription) Then
					lErrorNumber = -1
				End If
		End Select

		If lErrorNumber = 0 Then
			aCatalogComponent(S_QUERY_CONDITION_CATALOG) = Trim(aCatalogComponent(S_QUERY_CONDITION_CATALOG))
			If Len(aCatalogComponent(S_QUERY_CONDITION_CATALOG)) > 0 Then
				If InStr(1, aCatalogComponent(S_QUERY_CONDITION_CATALOG), "And ", vbTextCompare) <> 1 Then
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = "And " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
				End If
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
			End If

			If Not CheckCatalogInformationConsistency(aCatalogComponent, sErrorDescription) Then
				lErrorNumber = -1
			Else
				If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = -1) Or (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = -1) Then
					sQuery = "Update " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Set "
					For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
						If CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex)) <> N_FILE Then
							sQuery = sQuery & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "='" & Replace(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex), "'", "´") & "', "
						End If
					Next
					sQuery = Left(sQuery, (Len(sQuery) - Len(", ")))
					aCatalogComponent(S_QUERY_CONDITION_CATALOG) = Trim(aCatalogComponent(S_QUERY_CONDITION_CATALOG))
					sQuery = sQuery & " Where (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & ") " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
					sErrorDescription = "No se pudo modificar la información del registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				Else
					If StrComp(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)), oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & "Old").Item, vbBinaryCompare) = 0 Then
						sQuery = "Update " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Set "
						For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
							If CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex)) <> N_FILE Then
								sQuery = sQuery & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "='" & Replace(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex), "'", "´") & "', "
							End If
						Next
						sQuery = Left(sQuery, (Len(sQuery) - Len(", ")))
						aCatalogComponent(S_QUERY_CONDITION_CATALOG) = Trim(aCatalogComponent(S_QUERY_CONDITION_CATALOG))
						sQuery = sQuery & " Where (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & ") And (" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & "=" & oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & "Old").Item & ") " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
						sErrorDescription = "No se pudo modificar la información del registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Else
						sErrorDescription = "No se pudo modificar la información del registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Where (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & ") And (" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & ">=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & ") And (" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & "<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & ") " & aCatalogComponent(S_QUERY_CONDITION_CATALOG), "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						sQuery = "Update " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Set " & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & "=" & AddDaysToSerialDate(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)), -1) & " "
						'If aCatalogComponent(N_ACTIVE_CATALOG) > -1 Then sQuery = sQuery & ", " & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ACTIVE_CATALOG)) & "=0 "
						sQuery = sQuery & " Where (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & ") And (" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & ">=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & ") And (" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & "<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & ") " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
						sErrorDescription = "No se pudo modificar la información del registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

						sQuery = "Update " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Set " & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & "=" & AddDaysToSerialDate(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)), 1) & " "
						'If aCatalogComponent(N_ACTIVE_CATALOG) > -1 Then sQuery = sQuery & ", " & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ACTIVE_CATALOG)) & "=0 "
						sQuery = sQuery & " Where (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & ") And (" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & "<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & ") And (" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & ">=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & ") " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
						sErrorDescription = "No se pudo modificar la información del registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

						sNames = ""
						sValues = ""
						For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
							If CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex)) <> N_FILE Then
								If iIndex = N_END_FIELD_FOR_HISTORY_LIST_CATALOG Then
									sNames = sNames & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & ", "
									sValues = sValues & "30000000', '"
								Else
									sNames = sNames & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & ", "
									sValues = sValues & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) & "', '"
								End If
							End If
						Next
						If Len(sNames) > 0 Then sNames = Left(sNames, (Len(sNames) - Len(", ")))
						If Len(sValues) > 0 Then sValues = Left(sValues, (Len(sValues) - Len("', '")))
						sErrorDescription = "No se pudo guardar la información del nuevo registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " (" & sNames & ") Values ('" & sValues & "')", "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
			End If
		End If
	End If

	ModifyCatalog = lErrorNumber
	Err.Clear
End Function

Function SetActiveForCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given catalog
'Inputs:  oRequest, oADODBConnection
'Outputs: aCatalogComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForCatalog"
	Dim sQuery
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCatalogComponent(B_COMPONENT_INITIALIZED_CATALOG)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCatalogComponent(oRequest, aCatalogComponent)
	End If

	If Len(aCatalogComponent(S_TABLE_NAME_CATALOG)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se pudo modificar la información el registro pues no se especificó el nombre de la tabla."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	ElseIf aCatalogComponent(N_ID_CATALOG) = -1 Then
		sCondition = Trim(aCatalogComponent(S_QUERY_CONDITION_CATALOG))
		If Len(sCondition) > 0 Then
			If InStr(1, sCondition, "And ", vbBinaryCompare) = 1 Then sCondition = Right(sCondition, (Len(sCondition) - Len("And ")))
		End If
		If Len(oRequest("SetActive").Item) > 0 Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ACTIVE_CATALOG)) = CInt(oRequest("SetActive").Item)
		If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = -1) Or (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = -1) Then
			sQuery = "Update " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Set " & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ACTIVE_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ACTIVE_CATALOG)) & " Where " & sCondition
		Else
			sQuery = "Update " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Set " & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ACTIVE_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ACTIVE_CATALOG)) & " Where (" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & ")" & sCondition
		End If
		sErrorDescription = "No se pudo modificar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	ElseIf CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG))) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para modificar su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sCondition = Trim(aCatalogComponent(S_QUERY_CONDITION_CATALOG))
		If Len(sCondition) > 0 Then
			If InStr(1, sCondition, "And ", vbBinaryCompare) = 0 Then sCondition = " And " & sCondition
		End If
		If Len(oRequest("SetActive").Item) > 0 Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ACTIVE_CATALOG)) = CInt(oRequest("SetActive").Item)
		If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = -1) Or (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = -1) Then
			sQuery = "Update " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Set " & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ACTIVE_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ACTIVE_CATALOG)) & " Where (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & ")" & sCondition
		Else
			sQuery = "Update " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Set " & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ACTIVE_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ACTIVE_CATALOG)) & " Where (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & ") And (" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & ")" & sCondition
		End If
		sErrorDescription = "No se pudo modificar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	SetActiveForZone = lErrorNumber
	Err.Clear
End Function

Function RemoveCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a record from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aCatalogComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveCatalog"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCatalogComponent(B_COMPONENT_INITIALIZED_CATALOG)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCatalogComponent(oRequest, aCatalogComponent)
	End If

	If Len(aCatalogComponent(S_TABLE_NAME_CATALOG)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se pudo eliminar la información el registro pues no se especificó el nombre de la tabla."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	ElseIf aCatalogComponent(N_ID_CATALOG) = -1 Then
		aCatalogComponent(S_QUERY_CONDITION_CATALOG) = Trim(aCatalogComponent(S_QUERY_CONDITION_CATALOG))
		If Len(aCatalogComponent(S_QUERY_CONDITION_CATALOG)) > 0 Then
			If InStr(1, aCatalogComponent(S_QUERY_CONDITION_CATALOG), "And ", vbTextCompare) = 1 Then
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = Right(aCatalogComponent(S_QUERY_CONDITION_CATALOG), (Len(aCatalogComponent(S_QUERY_CONDITION_CATALOG)) - Len("And ")))
			End If
			aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
		End If
		If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) And (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) Then
			aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & ")"
		End If
		If Len(aCatalogComponent(S_QUERY_CONDITION_CATALOG)) > 0 Then aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " Where " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)

		sErrorDescription = "No se pudo eliminar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & aCatalogComponent(S_TABLE_NAME_CATALOG) & aCatalogComponent(S_QUERY_CONDITION_CATALOG), "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	ElseIf CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG))) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para eliminar su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		aCatalogComponent(S_QUERY_CONDITION_CATALOG) = Trim(aCatalogComponent(S_QUERY_CONDITION_CATALOG))
		If Len(aCatalogComponent(S_QUERY_CONDITION_CATALOG)) > 0 Then
			If InStr(1, aCatalogComponent(S_QUERY_CONDITION_CATALOG), "And ", vbTextCompare) <> 1 Then
				aCatalogComponent(S_QUERY_CONDITION_CATALOG) = "And " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
			End If
			aCatalogComponent(S_QUERY_CONDITION_CATALOG) = " " & aCatalogComponent(S_QUERY_CONDITION_CATALOG)
		End If
		If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) And (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) Then
			aCatalogComponent(S_QUERY_CONDITION_CATALOG) = aCatalogComponent(S_QUERY_CONDITION_CATALOG) & " And (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & ")"
		End If

		sErrorDescription = "No se pudo eliminar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Where (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & ") " & aCatalogComponent(S_QUERY_CONDITION_CATALOG), "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	RemoveCatalog = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfPositionInCatalog(oADODBConnection, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific position exists in the database
'Inputs:  aCatalogComponent
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfPositionInCatalog"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó la clave del puesto para para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
		'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Positions Where (CompanyID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(9) & ") And (ClassificationID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10) & ") And (GroupGradeLevelID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11) & ") And (IntegrationID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12) & ") And (PositionShortName ='" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & "')  And (LevelID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Positions Where (CompanyID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(9) & ") And (ClassificationID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10) & ") And (GroupGradeLevelID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11) & ") And (IntegrationID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12) & ") And (PositionShortName ='" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & "')  And (LevelID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13) & ") And (WorkingHours=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(18) & ") And (EmployeeTypeID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(7) & ") And (((StartDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(5) & ") And (EndDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(6) & ")) Or ((EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(5) & ") And (EndDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(6) & ")) Or ((EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(5) & ") And (StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(6) & ")))", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			aCatalogComponent(B_IS_DUPLICATED_CATALOG) = (Not oRecordset.EOF)
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfPositionInCatalog = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfRecordInCatalog(oADODBConnection, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific record exists in the database
'Inputs:  oADODBConnection, aCatalogComponent
'Outputs: aCatalogComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfRecordInCatalog"
	Dim sCondition
	Dim iIndex
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aCatalogComponent(B_COMPONENT_INITIALIZED_CATALOG)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeCatalogComponent(oRequest, aCatalogComponent)
	End If

	If aCatalogComponent(N_ID_CATALOG) > -1 Then
		If Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_NAME_CATALOG))) = 0 Then
			lErrorNumber = -1
			sErrorDescription = "No se especificó el nombre del registro para revisar su existencia en la base de datos."
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
		Else
			sCondition = Trim(aCatalogComponent(S_CHECK_EXISTENCY_CONDITION_CATALOG))
			If Len(sCondition) > 0 Then
				If InStr(1, sCondition, "And ", vbBinaryCompare) = 0 Then sCondition = " And " & sCondition
			End If
			For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
				sCondition = Replace(sCondition, "<FIELD_" & iIndex & " />", aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex))
			Next
			sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
			If CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(aCatalogComponent(N_ID_CATALOG))) = N_TEXT Then
				If aCatalogComponent(N_ID_CATALOG) <> aCatalogComponent(N_NAME_CATALOG) Then sCondition = sCondition & " And (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "<>'" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "')"
				If CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(aCatalogComponent(N_NAME_CATALOG))) = N_TEXT Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Where (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_NAME_CATALOG)) & "='" & Replace(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_NAME_CATALOG)), "'", "´") & "')" & sCondition, "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Where (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_NAME_CATALOG)) & "=" & Replace(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_NAME_CATALOG)), "'", "") & ")" & sCondition, "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				End If
			Else
				If aCatalogComponent(N_ID_CATALOG) <> aCatalogComponent(N_NAME_CATALOG) Then sCondition = sCondition & " And (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "<>" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & ")"
                Select Case CStr(oRequest("Action").Item)
                    Case "PaperworkOwners"
                        If Len(oRequest("Modify").Item) > 0 Then
                            sCondition = sCondition & " And (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "<>" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & ")"
                        End If
                End Select
				If CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(aCatalogComponent(N_NAME_CATALOG))) = N_TEXT Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Where (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_NAME_CATALOG)) & "='" & Replace(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_NAME_CATALOG)), "'", "´") & "')" & sCondition, "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Where (" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_NAME_CATALOG)) & "=" & Replace(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_NAME_CATALOG)), "'", "") & ")" & sCondition, "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				End If
			End If
			If lErrorNumber = 0 Then
				aCatalogComponent(B_IS_DUPLICATED_CATALOG) = (Not oRecordset.EOF)
			End If
		End If
		oRecordset.Close
	ElseIf Len(aCatalogComponent(S_CHECK_EXISTENCY_CONDITION_CATALOG)) > 0 Then
		If Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_NAME_CATALOG))) = 0 Then
			lErrorNumber = -1
			sErrorDescription = "No se especificó el nombre del registro para revisar su existencia en la base de datos."
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
		Else
			sCondition = Trim(aCatalogComponent(S_CHECK_EXISTENCY_CONDITION_CATALOG))
			If Len(sCondition) > 0 Then
				If InStr(1, sCondition, "And ", vbBinaryCompare) = 1 Then sCondition = Replace(sCondition, "And ", "", 1, 1, vbBinaryCompare)
			End If
			For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
				sCondition = Replace(sCondition, "<FIELD_" & iIndex & " />", aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex))
			Next
			sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From " & aCatalogComponent(S_TABLE_NAME_CATALOG) & " Where " & sCondition, "CatalogComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				aCatalogComponent(B_IS_DUPLICATED_CATALOG) = (Not oRecordset.EOF)
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfRecordInCatalog = lErrorNumber
	Err.Clear
End Function

Function CheckCatalogInformationConsistency(aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aCatalogComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckCatalogInformationConsistency"
	Dim iIndex
	Dim bIsCorrect

	bIsCorrect = True

	For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
		Select Case CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex))
			Case N_BOOLEAN
				If Not IsNumeric(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)) Then
					sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & " no es numérico."
					bIsCorrect = False
				Else
					If (CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)) <> 0) And (CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)) <> 1) Then
						sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & " no es booleano."
						bIsCorrect = False
					End If
				End If
			Case N_DATE
				If Not IsNumeric(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)) Then
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) = Left(GetSerialNumberForDate(""), Len("00000000"))
				Else
					If (CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)) < 19000000) And (CInt(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG)(iIndex)) = 1) Then
						aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) = Left(GetSerialNumberForDate(""), Len("00000000"))
					End If
				End If
			Case N_FLOAT, N_INTEGER
				If Not IsNumeric(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)) Then
					sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & " no es numérico."
					bIsCorrect = False
				Else
					If CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex)) = N_FLOAT Then
						aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) = CDbl(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex))
						aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) = CDbl(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex))
						aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) = CDbl(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex))
					Else
						aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) = CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex))
						aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) = CLng(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex))
						aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) = CLng(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex))
					End If
					Select Case CInt(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG)(iIndex))
						Case N_OPEN_MINIMUM
							If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) <= aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) Then
								sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & " debe ser mayor a " & aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) & "."
								bIsCorrect = False
							End If
						Case N_OPEN_MAXIMUM
							If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) >= aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) Then
								sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & " debe ser menor a " & aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) & "."
								bIsCorrect = False
							End If
						Case N_OPEN_MINIMUM_OPEN_MAXIMUM
							If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) <= aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) Then
								sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & " debe ser mayor a " & aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) & "."
								bIsCorrect = False
							End If
							If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) >= aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) Then
								sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & " debe ser menor a " & aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) & "."
								bIsCorrect = False
							End If
						Case N_CLOSED_MINIMUM
							If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) < aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) Then
								sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & " debe ser mayor o igual a " & aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) & "."
								bIsCorrect = False
							End If
						Case N_CLOSED_MINIMUM_OPEN_MAXIMUM
							If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) < aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) Then
								sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & " debe ser mayor o igual a " & aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) & "."
								bIsCorrect = False
							End If
							If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) >= aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) Then
								sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & " debe ser menor a " & aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) & "."
								bIsCorrect = False
							End If
						Case N_CLOSED_MAXIMUM
							If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) > aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) Then
								sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & " debe ser menor o igual a " & aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) & "."
								bIsCorrect = False
							End If
						Case N_OPEN_MINIMUM_CLOSED_MAXIMUM
							If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) <= aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) Then
								sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & " debe ser mayor a " & aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) & "."
								bIsCorrect = False
							End If
							If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) > aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) Then
								sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & " debe ser menor o igual a " & aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) & "."
								bIsCorrect = False
							End If
						Case N_CLOSED_MINIMUM_CLOSED_MAXIMUM
							If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) < aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) Then
								sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & " debe ser mayor o igual a " & aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) & "."
								bIsCorrect = False
							End If
							If aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) > aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) Then
								sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & " debe ser menor o igual a " & aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) & "."
								bIsCorrect = False
							End If
					End Select
				End If
			Case N_HOUR
				If Not IsNumeric(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)) Then
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) = Mid(GetSerialNumberForDate(""), Len("000000000"), Len("0000"))
				Else
					If Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)) > 2359 Then
						aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) = Mid(GetSerialNumberForDate(""), Len("000000000"), Len("0000"))
					End If
				End If
			Case N_TEXT
				If CInt(aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex)) > 0 Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) = Left(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex), aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex))
				If (CInt(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG)(iIndex)) = 1) And (Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)) = 0) Then
					sErrorDescription = sErrorDescription & "<BR />&nbsp;- El valor del campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & " está vacío."
					bIsCorrect = False
				End If
			Case N_FILE
		End Select
	Next
	If aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) > -1 Then
		If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG))) = 0 Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) = 30000000
	End If

	If Len(aCatalogComponent(S_TABLE_NAME_CATALOG)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre de la tabla del catálogo está vacío."
		bIsCorrect = False
	End If

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del registro contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "CatalogComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckCatalogInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayCatalogForm(oRequest, oADODBConnection, sAction, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a record from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aCatalogComponent
'Outputs: aCatalogComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayCatalogForm"
	Dim sTemp
	Dim asTemp
	Dim iStart
	Dim iIndex
	Dim jIndex
	Dim kIndex
	Dim sNames
	Dim lErrorNumber

	If aCatalogComponent(N_ID_CATALOG) > -1 Then
		If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG))) <> -1 Then
			lErrorNumber = GetCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
            Select Case CStr(oRequest("Action").Item)
                Case "PaperworkOwners"
                    aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) = "0"
            End Select
		End If
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function Check" & aCatalogComponent(S_FORM_NAME_CATALOG) & "Fields(oForm) {" & vbNewLine
				If Len(oRequest("Delete").Item) = 0 Or Len(oRequest("PaperworkID").Item) > 0 Then
					Response.Write "if (oForm) {" & vbNewLine
						If Not aCatalogComponent(B_SHOW_ID_FIELD_CATALOG) Then
							iStart = aCatalogComponent(N_ID_CATALOG) + 1
						Else
							iStart = 0
						End If
						For iIndex = iStart To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
							aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex) = CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex))
							aCatalogComponent(AS_FIELDS_LIMITS_CATALOG)(iIndex) = CInt(aCatalogComponent(AS_FIELDS_LIMITS_CATALOG)(iIndex))
							sTemp = aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex)
							If InStr(1, sTemp, "<", vbBinaryCompare) > 0 Then sTemp = Right(sTemp, (Len(sTemp) - InStrRev(sTemp, ">")))
							sTemp = Replace(sTemp, "&nbsp;", "")
							Select Case aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex)
								Case N_FLOAT, N_INTEGER
									If CInt(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG)(iIndex)) = 0 Then Response.Write "if (oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & ".value.length > 0) {" & vbNewLine
										Response.Write "oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & ".value = oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & ".value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
										Response.Write "if (!"
											If aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex) = 4 Then
												Response.Write "CheckIntegerValue"
											Else
												Response.Write "CheckFloatValue"
											End If
											Response.Write "(oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & ", 'el campo \'" & sTemp & "\'', "
												Select Case aCatalogComponent(AS_FIELDS_LIMITS_CATALOG)(iIndex)
													Case N_NONE
														Response.Write "N_NO_RANK_FLAG, N_CLOSED_FLAG"
													Case N_OPEN_MINIMUM
														Response.Write "N_MINIMUM_ONLY_FLAG, N_OPEN_FLAG"
													Case N_OPEN_MAXIMUM
														Response.Write "N_MAXIMUM_ONLY_FLAG, N_OPEN_FLAG"
													Case N_OPEN_MINIMUM_OPEN_MAXIMUM
														Response.Write "N_BOTH_FLAG, N_MINIMUM_OPEN_FLAG"
													Case N_CLOSED_MINIMUM
														Response.Write "N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG"
													Case N_CLOSED_MINIMUM_OPEN_MAXIMUM
														Response.Write "N_BOTH_FLAG, N_MAXIMUM_OPEN_FLAG"
													Case N_CLOSED_MAXIMUM
														Response.Write "N_MAXIMUM_ONLY_FLAG, N_CLOSED_FLAG"
													Case N_OPEN_MINIMUM_CLOSED_MAXIMUM
														Response.Write "N_BOTH_FLAG, N_MINIMUM_OPEN_FLAG"
													Case N_CLOSED_MINIMUM_CLOSED_MAXIMUM
														Response.Write "N_BOTH_FLAG, N_CLOSED_FLAG"
												End Select
											Response.Write ", " & aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) & ", " & aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) & ")"
										Response.Write ")" & vbNewLine
											Response.Write "return false;" & vbNewLine
									If CInt(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG)(iIndex)) = 0 Then Response.Write "}" & vbNewLine
								Case N_TEXT
									If CInt(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG)(iIndex)) = 1 Then
										Response.Write "if (oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & ".value.length == 0) {" & vbNewLine
											Response.Write "alert('Favor de introducir un valor para el campo \'" & sTemp & "\'.');" & vbNewLine
											Response.Write "oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & ".focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									End If
								Case N_LIST, N_HIERARCHY_LIST
									If CInt(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex)) > 0 Then
										Response.Write "if (CountSelectedItems(oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & ") < " & aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) & ") {" & vbNewLine
											Response.Write "alert('Favor de seleccionar al menos " & aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) & " elemento(s) de la lista \'" & sTemp & "\'.');" & vbNewLine
											Response.Write "oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & ".focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									End If
									If CInt(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex)) > 1 Then
										Response.Write "if (CountSelectedItems(oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & ") > " & aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) & ") {" & vbNewLine
											Response.Write "alert('Favor de seleccionar a lo más " & aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) & " elementos de la lista \'" & sTemp & "\'.');" & vbNewLine
											Response.Write "oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & ".focus();" & vbNewLine
											Response.Write "return false;" & vbNewLine
										Response.Write "}" & vbNewLine
									End If
							End Select
						Next

						If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) And (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) Then
							Response.Write "if (parseInt(oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & "Year.value + oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & "Month.value + oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & "Day.value) > 0) {" & vbNewLine
								Response.Write "if (parseInt(oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & "Year.value + oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & "Month.value + oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & "Day.value) > parseInt(oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & "Year.value + oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & "Month.value + oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & "Day.value)) {" & vbNewLine
									Response.Write "alert('El campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & " no puede ser posterior al campo " & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & ".');" & vbNewLine
									Response.Write "oForm." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & "Day.focus();" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine
							Response.Write "}" & vbNewLine
						End If
					Response.Write "}" & vbNewLine
					Response.Write aCatalogComponent(S_ADDITIONAL_FORM_SCRIPT_CATALOG) & vbNewLine
				End If
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of Check" & aCatalogComponent(S_FORM_NAME_CATALOG) & "Fields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""" & aCatalogComponent(S_FORM_NAME_CATALOG) & """ ID=""" & aCatalogComponent(S_FORM_NAME_CATALOG) & """ ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return Check" & aCatalogComponent(S_FORM_NAME_CATALOG) & "Fields(this)"">"
			If Len(aCatalogComponent(S_URL_CATALOG)) > 0 Then
				sTemp = aCatalogComponent(S_URL_CATALOG)
				For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
					sTemp = Replace(sTemp, "<FIELD_" & iIndex & " />", CleanStringForJavaScript(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)))
				Next
				For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
					sTemp = RemoveParameterFromURLString(sTemp, aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex))
				Next
				Call DisplayURLParametersAsHiddenValues(sTemp)
			End If
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & aCatalogComponent(S_TABLE_NAME_CATALOG) & """ />"
			If Not aCatalogComponent(B_SHOW_ID_FIELD_CATALOG) Then
				For iIndex = 0 To aCatalogComponent(N_ID_CATALOG)
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & """ ID=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Hdn"" VALUE=""" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) & """ />"
				Next
				iStart = aCatalogComponent(N_ID_CATALOG) + 1
			Else
				iStart = 0
			End If

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">" & vbNewLine
				For iIndex = iStart To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
					If InStr(1, aCatalogComponent(S_ADD_LINES_BEFORE_FIELDS_CATALOG), ("," & iIndex & ","), vbBinaryCompare) > 0 Then
						Response.Write "<TR><TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2""><BR /></FONT></TD></TR>"
					End If
					aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex) = CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex))
					Response.Write "<TR NAME=""" & aCatalogComponent(S_FORM_NAME_CATALOG) & "_" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Div"" ID=""" & aCatalogComponent(S_FORM_NAME_CATALOG) & "_" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Div"">"
						Select Case aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(iIndex)
							Case N_BOOLEAN
								Response.Write "<TD COLSPAN=""2""><NOBR>"
								Response.Write "<INPUT TYPE=""CHECKBOX"" NAME=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & """ ID=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Chk"" VALUE=""1"""
									If CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)) <> 0 Then Response.Write " CHECKED=""1"""
								Response.Write " "
									Response.Write aCatalogComponent(AS_SCRIPT_CATALOG)(iIndex)
									If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",", "," & iIndex & ",", vbBinaryCompare) > 0 Then Response.Write " STYLE=""display: none"""
								Response.Write " />"
								If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",", "," & iIndex & ",", vbBinaryCompare) > 0 Then Response.Write "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" HSPACE=""2"" />"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"""
									If StrComp(oRequest("Highlight").Item, aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), vbBinaryCompare) = 0 Then Response.Write " COLOR=""#" & S_WARNING_FOR_GUI & """"
								Response.Write ">&nbsp;" & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & "</NOBR></FONT></TD>"
							Case N_DATE
								aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) = CLng(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex))
								aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) = CLng(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex))
								If aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) <= 0 Then aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) = Year(Date()) + aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex)
								If aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) <= 0 Then aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) = Year(Date()) + Abs(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex))
								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"""
									If StrComp(oRequest("Highlight").Item, aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), vbBinaryCompare) = 0 Then Response.Write " COLOR=""#" & S_WARNING_FOR_GUI & """"
								If StrComp(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex), "<EMPTY />", vbBinaryCompare) <> 0 Then
									Response.Write "><NOBR>" & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & ":&nbsp;</NOBR></FONT></TD>"
								Else
									Response.Write ">&nbsp;</FONT></TD>"
								End If
								Response.Write "<TD VALIGN=""TOP"""
									If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",", "," & iIndex & ",", vbBinaryCompare) > 0 Then Response.Write " STYLE=""display: none"""
								Response.Write "><FONT FACE=""Arial"" SIZE=""2""><NOBR>"
									If Len(aCatalogComponent(AS_SCRIPT_CATALOG)(iIndex)) = 0 Then
										Response.Write DisplayDateCombosUsingSerial(CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)), aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex), aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex), True, (CInt(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG)(iIndex)) <> 1))
									Else
										Response.Write Replace(DisplayDateCombosUsingSerial(CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)), aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex), aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex), True, (CInt(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG)(iIndex)) <> 1)), " onChange=""", " onChange=""" & Replace(Replace(aCatalogComponent(AS_SCRIPT_CATALOG)(iIndex), "onChange=""", ""), """", " "))
									End If
								Response.Write "</NOBR></FONT></TD>"
								If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",", "," & iIndex & ",", vbBinaryCompare) > 0 Then
									Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><NOBR>"
										If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)) = 0 Then
											Response.Write "---"
										Else
											Response.Write DisplayDateFromSerialNumber(CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)), -1, -1, -1)
										End If
									Response.Write "</NOBR></FONT></TD>"
								End If
							Case N_FLOAT, N_INTEGER
								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"""
									If StrComp(oRequest("Highlight").Item, aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), vbBinaryCompare) = 0 Then Response.Write " COLOR=""#" & S_WARNING_FOR_GUI & """"
								If StrComp(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex), "<EMPTY />", vbBinaryCompare) <> 0 Then
									Response.Write "><NOBR>" & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & ":&nbsp;</NOBR></FONT></TD>"
								Else
									Response.Write ">&nbsp;</FONT></TD>"
								End If
								Response.Write "<TD VALIGN=""TOP"""
									If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",", "," & iIndex & ",", vbBinaryCompare) > 0 Then Response.Write " STYLE=""display: none"""
								Response.Write "><INPUT TYPE=""TEXT"" NAME=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & """ ID=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Txt"" SIZE=""" & aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex) & """ MAXLENGTH=""" & aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex) & """ VALUE=""" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) & """ CLASS=""TextFields"" "
									Response.Write aCatalogComponent(AS_SCRIPT_CATALOG)(iIndex)
								Response.Write " /></TD>"
								If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",", "," & iIndex & ",", vbBinaryCompare) > 0 Then
									Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) & "</FONT></TD>"
								End If
							Case N_HOUR
								aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) = CLng(aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex))
								aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) = CLng(aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex))
								aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex) = CInt(aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex))
								If aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) <= 0 Then aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex) = 0
								If aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) <= 0 Then aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex) = 23
								If (aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex) <= 0) Or (aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex) > 60) Then aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex) = 1
								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"""
									If StrComp(oRequest("Highlight").Item, aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), vbBinaryCompare) = 0 Then Response.Write " COLOR=""#" & S_WARNING_FOR_GUI & """"
								If StrComp(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex), "<EMPTY />", vbBinaryCompare) <> 0 Then
									Response.Write "><NOBR>" & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & ":&nbsp;</NOBR></FONT></TD>"
								Else
									Response.Write ">&nbsp;</FONT></TD>"
								End If
								Response.Write "<TD VALIGN=""TOP"""
									If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",", "," & iIndex & ",", vbBinaryCompare) > 0 Then Response.Write " STYLE=""display: none"""
								Response.Write "><FONT FACE=""Arial"" SIZE=""2""><NOBR>"
									If Len(aCatalogComponent(AS_SCRIPT_CATALOG)(iIndex)) = 0 Then
										Response.Write DisplayTimeCombosUsingSerial(CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)), aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex), aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex), aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex), (CInt(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG)(iIndex)) <> 1))
									Else
										Response.Write Replace(DisplayTimeCombosUsingSerial(CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)), aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), aCatalogComponent(AS_FIELDS_MINIMUMS_CATALOG)(iIndex), aCatalogComponent(AS_FIELDS_MAXIMUMS_CATALOG)(iIndex),  aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex), (CInt(aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG)(iIndex)) <> 1)), " onChange=""", " onChange=""" & Replace(Replace(aCatalogComponent(AS_SCRIPT_CATALOG)(iIndex), "onChange=""", ""), """", " "))
									End If
								Response.Write "</NOBR></FONT></TD>"
								If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",", "," & iIndex & ",", vbBinaryCompare) > 0 Then
									Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & DisplayTimeFromSerialNumber(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)) & "</NOBR></FONT></TD>"
								End If
							Case N_TEXT
								If CLng(aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex) <= 255) Then
									Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"""
										If StrComp(oRequest("Highlight").Item, aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), vbBinaryCompare) = 0 Then Response.Write " COLOR=""#" & S_WARNING_FOR_GUI & """"
									If StrComp(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex), "<EMPTY />", vbBinaryCompare) <> 0 Then
										Response.Write "><NOBR>" & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & ":&nbsp;</NOBR></FONT></TD>"
									Else
										Response.Write ">&nbsp;</FONT></TD>"
									End If
									Response.Write "<TD VALIGN=""TOP"""
										If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",", "," & iIndex & ",", vbBinaryCompare) > 0 Then Response.Write " STYLE=""display: none"""
									Response.Write "><INPUT TYPE=""TEXT"" NAME=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & """ ID=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Txt"" SIZE="""
										If CLng(aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex) < 30) Then
											Response.Write aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex)
										Else
											Response.Write "30"
										End If
									Response.Write """ MAXLENGTH=""" & aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex) & """ VALUE=""" & CleanStringForHTML(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)) & """ CLASS=""TextFields"" "
										Response.Write aCatalogComponent(AS_SCRIPT_CATALOG)(iIndex)
									Response.Write " /></TD>"
									If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",", "," & iIndex & ",", vbBinaryCompare) > 0 Then
										Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)) & "</FONT></TD>"
									End If
								Else
									Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"""
										If StrComp(oRequest("Highlight").Item, aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), vbBinaryCompare) = 0 Then Response.Write " COLOR=""#" & S_WARNING_FOR_GUI & """"
									If StrComp(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex), "<EMPTY />", vbBinaryCompare) <> 0 Then
										Response.Write ">" & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & ":<BR /></FONT>"
									Else
										Response.Write ">&nbsp;</FONT>"
									End If
										Response.Write "<TEXTAREA NAME=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & """ ID=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "TxtArea"" ROWS=""5"" COLS=""50"" MAXLENGTH=""" & aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex) & """ CLASS=""TextFields"" "
											Response.Write aCatalogComponent(AS_SCRIPT_CATALOG)(iIndex)
											If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",", "," & iIndex & ",", vbBinaryCompare) > 0 Then Response.Write " STYLE=""display: none"""
										Response.Write ">" & CleanStringForHTML(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)) & "</TEXTAREA>"
										If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",", "," & iIndex & ",", vbBinaryCompare) > 0 Then
											Response.Write "<SPAN CLASS=""FakeTextArea""><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)) & "</FONT></SPAN>"
										End If
									Response.Write "</TD>"
								End If
							Case N_CATALOG
								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"""
									If StrComp(oRequest("Highlight").Item, aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), vbBinaryCompare) = 0 Then Response.Write " COLOR=""#" & S_WARNING_FOR_GUI & """"
								If StrComp(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex), "<EMPTY />", vbBinaryCompare) <> 0 Then
									Response.Write "><NOBR>" & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & ":&nbsp;</NOBR></FONT></TD>"
								Else
									Response.Write ">&nbsp;</FONT></TD>"
								End If
								Response.Write "<TD VALIGN=""TOP"""
									If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",", "," & iIndex & ",", vbBinaryCompare) > 0 Then Response.Write " STYLE=""display: none"""
								Response.Write "><SELECT NAME=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & """ ID=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Cmb"" SIZE=""1"" CLASS=""Lists"" "
									Response.Write aCatalogComponent(AS_SCRIPT_CATALOG)(iIndex)
								Response.Write ">"
									If InStr(1, aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex), "<OPTION", vbBinaryCompare) = 0 Then
										aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex), SECOND_LIST_SEPARATOR)
										If aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG)(iIndex) = 0 Then Response.Write "<OPTION VALUE=""""></OPTION>"
										If (StrComp(aCatalogComponent(S_TABLE_NAME_CATALOG), "PositionsSpecialJourneysLKP", vbTextCompare) = 0) And (StrComp(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), "PositionID", vbTextCompare) = 0) Then
											If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG))) = -1 Then
												Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID", "PositionShortName, PositionName, 'Cia:' As Temp, CompanyID, 'Nivel:' As Temp, LevelID, 'Jornada:' As Temp, WorkingHours", "(CompanyID=1) And (EndDate=30000000) And (Active=1)", "PositionShortName", aConceptComponent(N_POSITION_ID_CONCEPT) & "," & aConceptComponent(N_LEVEL_ID_CONCEPT) & "," & aConceptComponent(D_WORKING_HOURS_CONCEPT), "", sErrorDescription)
											Else
												Response.Write GenerateListOptionsFromQuery(oADODBConnection, aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(0), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(1), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(2), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(3), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(4), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(5), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(6), sErrorDescription)
											End If
										Else
											Response.Write GenerateListOptionsFromQuery(oADODBConnection, aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(0), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(1), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(2), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(3), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(4), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(5), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(6), sErrorDescription)
										End If
									Else
										Response.Write aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)
									End If
								Response.Write "</SELECT></TD>"
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									Response.Write "SelectItemByValue('" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) & "', false, document." & aCatalogComponent(S_FORM_NAME_CATALOG) & "." & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & ")" & vbNewLine
								Response.Write "//--></SCRIPT>" & vbNewLine
								If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",", "," & iIndex & ",", vbBinaryCompare) > 0 Then
									If InStr(1, aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex), "<OPTION", vbBinaryCompare) = 0 Then
										Call GetNameFromTable(oADODBConnection, aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(0), aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex), "", "<BR />", sNames, sErrorDescription)
										Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2""><NOBR>" & CleanStringForHTML(sNames) & "</NOBR></FONT>"
									Else
										Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2""><NOBR>"
											asTemp = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex), "</OPTION>")
											For kIndex = 0 To UBound(asTemp)
												If InStr(1, asTemp(kIndex), "VALUE=""" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) & """", vbBinaryCompare) > 0 Then
													asTemp(kIndex) = Right(asTemp(kIndex), (Len(asTemp(kIndex)) - InStr(1, asTemp(kIndex), ">", vbBinaryCompare)))
													Response.Write CleanStringForHTML(asTemp(kIndex))
													Exit For
												End If
											Next
										Response.Write "</NOBR></FONT>"
									End If
								End If
							Case N_LIST
								If InStr(1, aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex), "<OPTION", vbBinaryCompare) = 0 Then
									aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex), SECOND_LIST_SEPARATOR)
								End If
								Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"""
									If StrComp(oRequest("Highlight").Item, aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), vbBinaryCompare) = 0 Then Response.Write " COLOR=""#" & S_WARNING_FOR_GUI & """"
								If StrComp(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex), "<EMPTY />", vbBinaryCompare) <> 0 Then
									If StrComp(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex),"AreaIDs",VbBinaryCompare) = 0 And StrComp(oRequest("Action").Item,"PaymentsRecords",VbBinaryCompare) = 0 Then
										Response.Write "><NOBR>" & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & ":&nbsp;</NOBR></FONT>"
									ElseIf StrComp(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex),"ZoneIDs",VbBinaryCompare) = 0 And StrComp(oRequest("Action").Item,"PaymentsRecords",VbBinaryCompare) = 0 Then
										Response.Write "><NOBR>" & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & ":&nbsp;</NOBR></FONT>"
									Else
										Response.Write "><NOBR>" & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & ":&nbsp;</NOBR></FONT>"
									end If
								Else
									Response.Write ">&nbsp;</FONT>"
								End If
								Response.Write "<BR />"
								If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",", "," & iIndex & ",", vbBinaryCompare) > 0 Then
									If InStr(1, aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex), "<OPTION", vbBinaryCompare) = 0 Then
										Call GetNameFromTable(oADODBConnection, aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(0), aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex), "&nbsp;&nbsp;&nbsp;", "<BR />", sNames, sErrorDescription)
										Response.Write "<FONT FACE=""Arial"" SIZE=""2""><NOBR>" & Replace(CleanStringForHTML(sNames), "&#38;nbsp;&#38;nbsp;&#38;nbsp;", "&nbsp;&nbsp;&nbsp;") & "</NOBR>"
									Else
										Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2""><NOBR>"
											asTemp = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex), "</OPTION>")
											For kIndex = 0 To UBound(asTemp)
												If InStr(1, asTemp(kIndex), "VALUE=""" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) & """", vbBinaryCompare) > 0 Then
													asTemp(kIndex) = Right(asTemp(kIndex), (Len(asTemp(kIndex)) - InStr(1, asTemp(kIndex), ">", vbBinaryCompare)))
													Response.Write CleanStringForHTML(asTemp(kIndex))
													Exit For
												End If
											Next
										Response.Write "</NOBR></FONT>"
									End If
								End If
								Response.Write "&nbsp;&nbsp;&nbsp;<SELECT NAME=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & """ ID=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex)
									If Len(aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex)) > 0 Then
										If CInt(aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex)) > 1 Then
											Response.Write "Lst"" SIZE=""" & aCatalogComponent(AS_FIELDS_SIZES_CATALOG)(iIndex) & """ MULTIPLE=""1"""
										Else
											Response.Write "Lst"" SIZE=""5"" MULTIPLE=""1"""
										End If
									Else
										Response.Write "Lst"" SIZE=""5"" MULTIPLE=""1"""
									End If
								Response.Write " CLASS=""Lists"" "
									Response.Write aCatalogComponent(AS_SCRIPT_CATALOG)(iIndex)
									If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_BLOCK_CATALOG) & ",", "," & iIndex & ",", vbBinaryCompare) > 0 Then Response.Write " STYLE=""display: none"""
								Response.Write ">"
									If (StrComp(aCatalogComponent(S_TABLE_NAME_CATALOG), "PaymentsRecords", vbTextCompare) = 0) Then
										If aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG)(iIndex) = 0 Then Response.Write "<OPTION VALUE="""">NO APLICA</OPTION>"
									Else
										If aCatalogComponent(AS_FIELDS_REQUIRED_CATALOG)(iIndex) = 0 Then Response.Write "<OPTION VALUE=""""></OPTION>"
									End If
									If InStr(1, aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex), "<OPTION", vbBinaryCompare) = 0 Then
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(0), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(1), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(2), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(3), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(4), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(5), aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)(6), sErrorDescription)
									Else
										Response.Write aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(iIndex)
									End If
									If (StrComp(aCatalogComponent(S_TABLE_NAME_CATALOG), "PaymentsRecords", vbTextCompare) = 0) Then
										Response.Write "<OPTION VALUE=""38"">HOSP. REG. PDTE. JUAREZ OAXACA, OAX.</OPTION>"
									End If
								Response.Write "</SELECT></TD>"
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
									Response.Write "SendURLValuesToForm('" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "=" & Replace(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex), ", ", ",") & "', document." & aCatalogComponent(S_FORM_NAME_CATALOG) & ");" & vbNewLine
								Response.Write "//--></SCRIPT>" & vbNewLine
							Case N_FILE
								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"""
									If StrComp(oRequest("Highlight").Item, aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex), vbBinaryCompare) = 0 Then Response.Write " COLOR=""#" & S_WARNING_FOR_GUI & """"
								If StrComp(aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex), "<EMPTY />", vbBinaryCompare) <> 0 Then
									Response.Write "><NOBR>" & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(iIndex) & ":&nbsp;</NOBR></FONT></TD>"
								Else
									Response.Write ">&nbsp;</FONT></TD>"
								End If
								Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Agregar archivo</FONT></TD>"
							Case N_HIDDEN
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & """ ID=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Hdn"" VALUE=""" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) & """ />"
						End Select
					Response.Write "</TR>" & vbNewLine
					If InStr(1, "," & aCatalogComponent(B_SHOW_ID_FIELD_CATALOG) & ",", "," & iIndex & ",", vbBinaryCompare) > 0 Then Response.Write "<TR><TD COLSPAN=""2""><BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""100%"" HEIGHT=""1"" /><BR /><BR /></TD></TR>"
				Next
			Response.Write "</TABLE><BR />" & vbNewLine

			If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) <> -1) And (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) <> -1) Then
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & "Old"" ID=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & "OldHdn"" VALUE=""" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG)) & """ />" & vbNewLine
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & "Old"" ID=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & "OldHdn"" VALUE=""" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG)) & """ />" & vbNewLine
			End If

			Response.Write aCatalogComponent(S_ADDITIONAL_FORM_HTML_CATALOG) & vbNewLine
			Response.Write "<DIV NAME=""ExtraHTMLFor" & aCatalogComponent(S_FORM_NAME_CATALOG) & "Div"" ID=""ExtraHTMLFor" & aCatalogComponent(S_FORM_NAME_CATALOG) & "Div""></DIV>" & vbNewLine

			Response.Write "<SPAN NAME=""Buttons" & aCatalogComponent(S_FORM_NAME_CATALOG) & "Div"" ID=""Buttons" & aCatalogComponent(S_FORM_NAME_CATALOG) & "Div"">" & vbNewLine
				If aCatalogComponent(N_ID_CATALOG) > -1 Then
					If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG))) = -1 Then
						If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />" & vbNewLine
					ElseIf Len(oRequest("Delete").Item) > 0 Then
						If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['Remove" & aCatalogComponent(S_FORM_NAME_CATALOG) & "WngDiv']); " & aCatalogComponent(S_FORM_NAME_CATALOG) & ".Remove.focus()"" />" & vbNewLine
					Else
						If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />" & vbNewLine
					End If
				ElseIf aCatalogComponent(B_FORCE_SHOW_MODIFY_BUTTON_CATALOG) Then
					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />" & vbNewLine
				ElseIf aCatalogComponent(B_FORCE_SHOW_DELETE_BUTTON_CATALOG) Then
					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['Remove" & aCatalogComponent(S_FORM_NAME_CATALOG) & "WngDiv']); " & aCatalogComponent(S_FORM_NAME_CATALOG) & ".Remove.focus()"" />" & vbNewLine
				Else
					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />" & vbNewLine
				End If
				If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS) = N_ADD_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />" & vbNewLine
				End If
				If Len(aCatalogComponent(S_CANCEL_BUTTON_ACTION_CATALOG)) = 0 Then aCatalogComponent(S_CANCEL_BUTTON_ACTION_CATALOG) = "window.location.href='" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "'"
				Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""" & aCatalogComponent(S_CANCEL_BUTTON_ACTION_CATALOG) & """ />" & vbNewLine
				If Len(aCatalogComponent(S_EXTRA_BUTTON_CATALOG)) > 0 Then
					Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />" & vbNewLine
					Response.Write aCatalogComponent(S_EXTRA_BUTTON_CATALOG)
				End If
				Response.Write "<BR /><BR />" & vbNewLine
			Response.Write "</SPAN>" & vbNewLine
			Call DisplayWarningDiv("Remove" & aCatalogComponent(S_FORM_NAME_CATALOG) & "WngDiv", "¿Está seguro que desea borrar el registro de la &nbsp;base de datos?")
		Response.Write "</FORM>" & vbNewLine
		
		If (Len(oRequest("Change").Item) = 0) And (Len(oRequest("Modify").Item) = 0) And (Len(oRequest("Delete").Item) = 0) And (Len(oRequest("Remove").Item) = 0) Then
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				sTemp = RemoveParameterFromURLString(aCatalogComponent(S_URL_PARAMETERS_CATALOG), "Action")
				If InStr(1, sTemp, "?", vbBinaryCompare) > 0 Then sTemp = Right(sTemp, (Len(sTemp) - InStr(1, sTemp, "?", vbBinaryCompare)))
				For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
					sTemp = Replace(sTemp, "<FIELD_" & iIndex & " />", CleanStringForJavaScript(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex)))
				Next
				Response.Write "SendURLValuesToForm('" & sTemp & "', document." & aCatalogComponent(S_FORM_NAME_CATALOG) & ");" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
	End If

	DisplayCatalogForm = lErrorNumber
	Err.Clear
End Function

Function DisplayCatalogAsEmptyURL(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a record using
'		  hidden form fields
'Inputs:  oRequest, oADODBConnection, aCatalogComponent
'Outputs: aCatalogComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Dim sTemp
	Dim iIndex
	Const S_FUNCTION_NAME = "DisplayCatalogAsEmptyURL"

	For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
		Response.Write aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "=" & aCatalogComponent(AS_DEFAULT_VALUES_CATALOG)(iIndex) & "&"
	Next
	If Len(aCatalogComponent(S_URL_PARAMETERS_CATALOG)) > 0 Then
		sTemp = aCatalogComponent(S_URL_PARAMETERS_CATALOG)
		If InStr(1, sTemp, "?", vbBinaryCompare) > 0 Then sTemp = Right(sTemp, (Len(sTemp) - InStr(1, sTemp, "?", vbBinaryCompare)))
		For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
			sTemp = Replace(sTemp, "<FIELD_" & iIndex & " />", aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex))
		Next
		Response.Write sTemp & "&"
	End If

	DisplayCatalogAsEmptyURL = Err.number
	Err.Clear
End Function

Function DisplayCatalogAsHiddenFields(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a record using
'		  hidden form fields
'Inputs:  oRequest, oADODBConnection, aCatalogComponent
'Outputs: aCatalogComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Dim iIndex
	Const S_FUNCTION_NAME = "DisplayCatalogAsHiddenFields"

	If Len(aCatalogComponent(S_URL_CATALOG)) > 0 Then Call DisplayURLParametersAsHiddenValues(aCatalogComponent(S_URL_CATALOG))
	For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & """ ID=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(iIndex) & "Hdn"" VALUE=""" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(iIndex) & """ />"
	Next

	DisplayCatalogAsHiddenFields = Err.number
	Err.Clear
End Function

Function DisplayCatalogsTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the records from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aCatalogComponent
'Outputs: aCatalogComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayCatalogsTable"
	Dim iIndex
	Dim jIndex
	Dim sTemp
	Dim sValue
	Dim bLinkDisplayed
	Dim asTemp
	Dim iRecordCounter
	Dim adTotals
	Dim oRecordset
	Dim sColumnsTitles
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim sCellAlignments
	Dim asCellAlignments
	Dim sBoldBegin
	Dim sBoldEnd
	Dim lErrorNumber
	Dim iStartPage
	Dim bForExport

	bForExport = (StrComp(GetASPFileName(""), "Export.asp", vbbinaryCompare) = 0)
	If bForExport Then bUseLinks = False
	lErrorNumber = GetCatalogs(oRequest, oADODBConnection, aCatalogComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			iStartPage = 1
			If Len(oRequest("StartPage").Item) > 0 Then iStartPage = CInt(oRequest("StartPage").Item)
			If Not bForExport Then Call DisplayIncrementalFetch(oRequest, iStartPage, ROWS_CATALOG, oRecordset)
			If bForExport Then Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>" & aCatalogComponent(S_NAME_CATALOG) & "</B></FONT><BR /><BR />"
			Response.Write "<TABLE WIDTH=""250"" BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				sColumnsTitles = ""
				sCellAlignments = ""
				For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG))
					sColumnsTitles = sColumnsTitles & aCatalogComponent(AS_FIELDS_TEXTS_CATALOG)(CInt(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex))) & ","
					sCellAlignments = sCellAlignments & ","
				Next
				sColumnsTitles = Left(sColumnsTitles, (Len(sColumnsTitles) - Len(",")))
				If bUseLinks And (aCatalogComponent(B_MODIFY_CATALOG) Or aCatalogComponent(B_DELETE_CATALOG) Or aCatalogComponent(B_ACTIVE_CATALOG)) Then
					asColumnsTitles = Split("Acciones," & sColumnsTitles & "", ",", -1, vbBinaryCompare)
					asCellWidths = Split("80,150", ",", -1, vbBinaryCompare)
				ElseIf bForExport Then
					asColumnsTitles = Split(sColumnsTitles, ",", -1, vbBinaryCompare)
					asCellWidths = Split("230", ",", -1, vbBinaryCompare)
					sCellAlignments = Left(sCellAlignments, (Len(sCellAlignments) - Len(",")))
				Else
					asColumnsTitles = Split("&nbsp;," & sColumnsTitles, ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,230", ",", -1, vbBinaryCompare)
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

				asCellAlignments = Split("CENTER," & sCellAlignments, ",", -1, vbBinaryCompare)
				iRecordCounter = 0
				adTotals = Split(aCatalogComponent(S_FIELDS_TO_SUM_CATALOG), ",")
				For iIndex = 0 To UBound(adTotals)
					adTotals(iIndex) = Array(adTotals(iIndex), 0)
					adTotals(iIndex)(0) = CInt(adTotals(iIndex)(0))
					adTotals(iIndex)(1) = 0
				Next
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) = -1) Or (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) = -1) Then
						If aCatalogComponent(N_ID_CATALOG) > -1 Then
							If StrComp(CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Value), oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Item, vbBinaryCompare) = 0 Then
								sBoldBegin = "<B>"
								sBoldEnd = "</B>"
							End If
						End If
					Else
						If aCatalogComponent(N_ID_CATALOG) > -1 Then
							If (StrComp(CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Value), oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Item, vbBinaryCompare) = 0) And (StrComp(CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG))).Value), oRequest(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG))).Item, vbBinaryCompare) = 0) Then
								sBoldBegin = "<B>"
								sBoldEnd = "</B>"
							End If
						End If
					End If
					If aCatalogComponent(N_ACTIVE_CATALOG) > -1 Then
						If CInt(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ACTIVE_CATALOG))).Value) = 0 Then
							sBoldBegin = sBoldBegin & "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
							sBoldEnd = sBoldEnd & "</FONT>"
						End If
					End If
					sRowContents = ""

					If bUseLinks And (aCatalogComponent(B_MODIFY_CATALOG) Or aCatalogComponent(B_DELETE_CATALOG)) Then
						sRowContents = sRowContents & "<NOBR>"
							If (Len(aCatalogComponent(S_SHOW_LINKS_FOR_IDS_CATALOG)) = 0) Or (InStr(1, "," & aCatalogComponent(S_SHOW_LINKS_FOR_IDS_CATALOG) & ",", "," & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Value) & ",", vbBinaryCompare) > 0) Then
								If InStr(1, "," & aCatalogComponent(S_IDS_NOT_UPDATABLE_CATALOG) & ",", "," & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Value) & ",", vbBinaryCompare) = 0 Then
									sTemp = aCatalogComponent(S_URL_CATALOG)
									For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
										sTemp = Replace(sTemp, "<FIELD_" & iIndex & " />", CleanStringForJavaScript(CStr(oRecordset.Fields(iIndex).Value)))
									Next

									If StrComp(aCatalogComponent(S_TABLE_NAME_CATALOG), "PositionsSpecialJourneysLKP", vbTextCompare) = 0 Then
										If CInt(oRecordset.Fields("Active").Value) <= 0 Then
											Select Case CInt(oRecordset.Fields("Active").Value)
												Case 0
													'sRowContents = sRowContents & "<IMG SRC=""Images/IcnExclamationSmall.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Al agregar este registro se detectaron registros empalmados que se ajustaran al aplicar este registro"" BORDER=""0"" />"
													sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
												Case -1
													sRowContents = sRowContents & "<IMG SRC=""Images/IcnExclamationSmall.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Al agregar este registro se detectaron registros posteriores que serán ajustados al aplicar este registro"" BORDER=""0"" />"
													sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
												Case -2
													sRowContents = sRowContents & "<IMG SRC=""Images/IcnInformation.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Al agregar este registro se detectaron registros dentro de los efectos de este que se ajustaran al aplicar este registro"" BORDER=""0"" />"
													sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
												Case -3
													sRowContents = sRowContents & "<IMG SRC=""Images/IcnExclamationSmall.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Al agregar este registro se detectaron registros que cubren todo el periodo de este, los cuales se se ajustaran al aplicar este registro"" BORDER=""0"" />"
													sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
											End Select
											If aCatalogComponent(B_MODIFY_CATALOG) Then
												If aCatalogComponent(N_ID_CATALOG) > -1 Then
													sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "&" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Value)
												Else
													sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG)
												End If
													If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) And (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) Then sRowContents = sRowContents & "&StartDate=" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG))).Value)
												sRowContents = sRowContents & "&Apply=1&StartPage=" & oRequest("StartPage").Item & "&" & sTemp & """>"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
											End If
											If aCatalogComponent(B_DELETE_CATALOG) Then
												If aCatalogComponent(N_ID_CATALOG) > -1 Then
													sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "&" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Value)
												Else
													sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG)
												End If
													If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) And (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) Then sRowContents = sRowContents & "&StartDate=" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG))).Value)
												sRowContents = sRowContents & "&Remove=1&StartPage=" & oRequest("StartPage").Item & "&" & sTemp & """>"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
											End If
										Else
											If aCatalogComponent(B_MODIFY_CATALOG) Then
												If aCatalogComponent(N_ID_CATALOG) > -1 Then
													sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "&" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Value)
												Else
													sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG)
												End If
													If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) And (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) Then sRowContents = sRowContents & "&StartDate=" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG))).Value)
												sRowContents = sRowContents & "&Change=1&StartPage=" & oRequest("StartPage").Item & "&" & sTemp & """>"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
											End If
											If aCatalogComponent(B_DELETE_CATALOG) Then
												If aCatalogComponent(N_ID_CATALOG) > -1 Then
													sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "&" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Value)
												Else
													sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG)
												End If
													If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) And (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) Then sRowContents = sRowContents & "&StartDate=" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG))).Value)
												sRowContents = sRowContents & "&Remove=1&StartPage=" & oRequest("StartPage").Item & "&" & sTemp & """>"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
												sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
											End If
										End If
									Else
										If aCatalogComponent(B_MODIFY_CATALOG) Then
											If aCatalogComponent(N_ID_CATALOG) > -1 Then
												sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "&" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Value)
											Else
												sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG)
											End If
												If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) And (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) Then sRowContents = sRowContents & "&StartDate=" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG))).Value)
											sRowContents = sRowContents & "&Change=1&StartPage=" & oRequest("StartPage").Item & "&" & sTemp & """>"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
										End If

										If B_DELETE And aCatalogComponent(B_DELETE_CATALOG) Then
											If aCatalogComponent(N_ID_CATALOG) > -1 Then
												sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "&" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Value)
											Else
												sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG)
											End If
												If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) And (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) Then sRowContents = sRowContents & "&StartDate=" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG))).Value)
											sRowContents = sRowContents & "&Remove=1&StartPage=" & oRequest("StartPage").Item & "&" & sTemp & """>"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
										End If

										If aCatalogComponent(B_ACTIVE_CATALOG) And (aCatalogComponent(N_ACTIVE_CATALOG) > -1) Then
											If CInt(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ACTIVE_CATALOG))).Value) = 0 Then
												If aCatalogComponent(N_ID_CATALOG) > -1 Then
													sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "&" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Value)
												Else
													sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG)
												End If
													If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) And (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) Then sRowContents = sRowContents & "&StartDate=" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG))).Value)
												sRowContents = sRowContents & "&SetActive=1&StartPage=" & oRequest("StartPage").Item & "&" & sTemp & """><IMG SRC=""Images/BtnActive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Activar registro"" BORDER=""0"" /></A>"
											Else
												If aCatalogComponent(N_ID_CATALOG) > -1 Then
													sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG) & "&" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "=" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Value)
												Else
													sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=" & aCatalogComponent(S_TABLE_NAME_CATALOG)
												End If
													If (aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) And (aCatalogComponent(N_END_FIELD_FOR_HISTORY_LIST_CATALOG) > -1) Then sRowContents = sRowContents & "&StartDate=" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_START_FIELD_FOR_HISTORY_LIST_CATALOG))).Value)
												sRowContents = sRowContents & "&SetActive=0&StartPage=" & oRequest("StartPage").Item & "&" & sTemp & """><IMG SRC=""Images/BtnDeactive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Desactivar registro"" BORDER=""0"" /></A>"
											End If
										End If
									End If
								End If
							End If
						sRowContents = sRowContents & "&nbsp;</NOBR>"
					End If

					If Not bForExport Then
						If aCatalogComponent(N_ID_CATALOG) > -1 Then
							Select Case lIDColumn
								Case DISPLAY_RADIO_BUTTONS
									sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & """ ID=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "Rd"" VALUE=""" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Value) & """ />"
								Case DISPLAY_CHECKBOXES
									sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & """ ID=""" & aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) & "Chk"" VALUE=""" & CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(aCatalogComponent(N_ID_CATALOG))).Value) & """ />"
								Case Else
									sRowContents = sRowContents & "&nbsp;"
							End Select
						Else
							sRowContents = sRowContents & "&nbsp;"
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR
					End If
					bLinkDisplayed = False
					If Not IsArray(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)) Then aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG), CATALOG_SEPARATOR)
					For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG))
						sValue = ""
						sValue = CStr(oRecordset.Fields(aCatalogComponent(AS_FIELDS_NAMES_CATALOG)(CInt(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex)))).Value)
						Err.Clear
						sRowContents = sRowContents & "<A"
							If (Len(aCatalogComponent(S_URL_PARAMETERS_CATALOG)) > 0) And (Not bLinkDisplayed) And (Not bForExport) Then
								sTemp = aCatalogComponent(S_URL_PARAMETERS_CATALOG)
								For jIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_NAMES_CATALOG))
									sTemp = Replace(sTemp, "<FIELD_" & jIndex & " />", CStr(oRecordset.Fields(jIndex).Value))
								Next
								sRowContents = sRowContents & " HREF=""" & sTemp & """"
								If Len(sValue) > 0 Then bLinkDisplayed = True
							End If
						sRowContents = sRowContents & ">" & sBoldBegin
							aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex)) = CInt(aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex)))
							Select Case aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex))
								Case N_BOOLEAN
									sRowContents = sRowContents & DisplayYesNo(sValue, True)
								Case N_DATE
									If (CLng(sValue) > 0) And (CLng(sValue) < 30000000) Then
										sRowContents = sRowContents & DisplayDateFromSerialNumber(sValue, -1, -1, -1)
									Else
										sRowContents = sRowContents & "<CENTER>---</CENTER>"
									End If
								Case N_FLOAT
									sRowContents = sRowContents & FormatNumber(CDbl(sValue), 2, True, False, True)
									If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_SUM_CATALOG) & ",", "," & aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex) & ",", vbBinaryCompare) > 0 Then
										For jIndex = 0 to UBound(adTotals)
											If adTotals(jIndex)(0) = CInt(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex)) Then
												adTotals(jIndex)(1) = adTotals(jIndex)(1) + CDbl(sValue)
												Exit For
											End If
										Next
									End If
								Case N_HOUR
									sTemp = sValue
									sTemp = Right(("0000" & sTemp), Len("0000"))
'									If Len(sTemp) < 6 Then sTemp = Right(("000000" & sTemp & "00"), Len("000000"))
									sRowContents = sRowContents & DisplayTimeFromSerialNumber(sTemp)
								Case N_INTEGER
									If aCatalogComponent(N_ID_CATALOG) = -1 Then
										sRowContents = sRowContents & sValue
									ElseIf aCatalogComponent(N_ID_CATALOG) <> CInt(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex)) Then
										sRowContents = sRowContents & FormatNumber(CLng(sValue), 0, True, False, True)
									Else
										sRowContents = sRowContents & sValue
									End If
									If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_SUM_CATALOG) & ",", "," & aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex) & ",", vbBinaryCompare) > 0 Then
										For jIndex = 0 to UBound(adTotals)
											If adTotals(jIndex)(0) = CInt(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex)) Then
												adTotals(jIndex)(1) = adTotals(jIndex)(1) + CDbl(sValue)
												Exit For
											End If
										Next
									End If
								Case N_CATALOG, N_HIERARCHY_CATALOG
									If IsArray(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(CInt(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex)))) Then
										asTemp = aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(CInt(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex)))
									Else
										asTemp = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(CInt(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex))), SECOND_LIST_SEPARATOR)
									End If
									asTemp(2) = Split(Replace(asTemp(2), " ", ""), ",")
									For jIndex = 0 To UBound(asTemp(2))
										sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields(asTemp(2)(jIndex)).Value))
										If jIndex < UBound(asTemp(2)) Then sRowContents = sRowContents & " "
									Next
								Case N_LIST, N_HIERARCHY_LIST
									If Not IsArray(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex))) Then aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex)) = Split(aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex)), SECOND_LIST_SEPARATOR)
									Call GetNameFromTable(oADODBConnection, aCatalogComponent(AS_CATALOG_PARAMETERS_CATALOG)(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex))(0), sValue, "", vbNewLine, sTemp, sErrorDescription)
									sRowContents = sRowContents & Replace(CleanStringForHTML(sTemp), vbNewLine, "<BR />")
								Case Else
									If bForExport And (InStr(1, sValue, vbNewLine, vbBinaryCompare) = 0) And (InStr(1, sValue, """", vbBinaryCompare) = 0) Then
										sRowContents = sRowContents & "=T(""" & CleanStringForHTML(sValue) & """)"
									Else
										sRowContents = sRowContents & CleanStringForHTML(sValue)
									End If
							End Select
						sRowContents = sRowContents & sBoldEnd & "</A>" & TABLE_SEPARATOR
					Next
					If Len(sRowContents) > 0 Then sRowContents = Left(sRowContents, (Len(sRowContents) - Len(TABLE_SEPARATOR)))

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					iRecordCounter = iRecordCounter + 1
					If (Not bForExport) And (iRecordCounter >= ROWS_CATALOG) Then Exit Do
					If Err.number <> 0 Then Exit Do
				Loop

				If Len(aCatalogComponent(S_FIELDS_TO_SUM_CATALOG)) > 0 Then
If StrComp(oRequest("SectionID").Item, "261", vbBinaryCompare) = 0 Then
	If adTotals(2)(1) >= 30 Then
		adTotals(1)(1) = adTotals(1)(1) + Int(adTotals(2)(1) / 30.4)
		adTotals(2)(1) = adTotals(2)(1) Mod 30.4
	End If
	If adTotals(1)(1) >= 12 Then
		adTotals(0)(1) = adTotals(0)(1) + Int(adTotals(1)(1) / 12)
		adTotals(1)(1) = adTotals(1)(1) Mod 12
	End If
End If
					sRowContents = ""
					For iIndex = 0 To UBound(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG))
						If InStr(1, "," & aCatalogComponent(S_FIELDS_TO_SUM_CATALOG) & ",", "," & aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex) & ",", vbBinaryCompare) > 0 Then
							For jIndex = 0 to UBound(adTotals)
								If adTotals(jIndex)(0) = CInt(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex)) Then
									If aCatalogComponent(AS_FIELDS_TYPES_CATALOG)(aCatalogComponent(AS_FIELDS_TO_SHOW_CATALOG)(iIndex)) = N_FLOAT Then
										sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(jIndex)(1), 2, True, False, True) & "</B>"
									Else
										sRowContents = sRowContents & TABLE_SEPARATOR & "<B>" & FormatNumber(adTotals(jIndex)(1), 0, True, False, True) & "</B>"
									End If
									Exit For
								End If
							Next
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						End If
					Next
					If bUseLinks And (aCatalogComponent(B_MODIFY_CATALOG) Or aCatalogComponent(B_DELETE_CATALOG) Or aCatalogComponent(B_ACTIVE_CATALOG)) Then sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
				End If
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayCatalogsTable = lErrorNumber
	Err.Clear
End Function
%>