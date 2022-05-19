<%
Const N_ID_EMPLOYEE_FIELD = 0
Const S_NAME_EMPLOYEE_FIELD = 1
Const S_TEXT_EMPLOYEE_FIELD = 2
Const N_IS_OPTIONAL_EMPLOYEE_FIELD  = 3
Const N_TYPE_ID_EMPLOYEE_FIELD = 4
Const N_SIZE_EMPLOYEE_FIELD = 5
Const N_LIMIT_TYPE_EMPLOYEE_FIELD = 6
Const N_MINIMUM_EMPLOYEE_FIELD = 7
Const N_MAXIMUM_EMPLOYEE_FIELD = 8
Const S_DSN_FOR_SOURCE_EMPLOYEE_FIELD = 9
Const N_CONNECTION_TYPE_FOR_SOURCE_EMPLOYEE_FIELD = 10
Const S_QUERY_FOR_SOURCE_EMPLOYEE_FIELD = 11
Const S_DEFAULT_VALUE_EMPLOYEE_FIELD = 12
Const S_JAVASCRIPT_CODE_EMPLOYEE_FIELD = 13
Const S_DESCRIPTION_EMPLOYEE_FIELD = 14
Const B_CHECK_FOR_DUPLICATED_EMPLOYEE_FIELD = 15
Const B_IS_DUPLICATED_EMPLOYEE_FIELD = 16
Const B_COMPONENT_INITIALIZED_EMPLOYEE_FIELD = 17

Const N_EMPLOYEE_FIELD_COMPONENT_SIZE = 17

Dim aEmployeeFieldComponent()
Redim aEmployeeFieldComponent(N_EMPLOYEE_FIELD_COMPONENT_SIZE)

Function InitializeEmployeeFieldComponent(oRequest, aEmployeeFieldComponent)
'************************************************************
'Purpose: To initialize the empty elements of the EmployeeField Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aEmployeeFieldComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeEmployeeFieldComponent"
	Redim Preserve aEmployeeFieldComponent(N_EMPLOYEE_FIELD_COMPONENT_SIZE)
	Dim aTemp

	If IsEmpty(aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD)) Then
		If Len(oRequest("FormFieldID").Item) > 0 Then
			aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD) = CLng(oRequest("FormFieldID").Item)
		Else
			aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD) = -1
		End If
	End If

	If IsEmpty(aEmployeeFieldComponent(S_NAME_EMPLOYEE_FIELD)) Then
		If Len(oRequest("FormFieldName").Item) > 0 Then
			aEmployeeFieldComponent(S_NAME_EMPLOYEE_FIELD) = oRequest("FormFieldName").Item
		Else
			aEmployeeFieldComponent(S_NAME_EMPLOYEE_FIELD) = ""
		End If
	End If
	aEmployeeFieldComponent(S_NAME_EMPLOYEE_FIELD) = Left(aEmployeeFieldComponent(S_NAME_EMPLOYEE_FIELD), 255)

	If IsEmpty(aEmployeeFieldComponent(S_TEXT_EMPLOYEE_FIELD)) Then
		If Len(oRequest("FormFieldText").Item) > 0 Then
			aEmployeeFieldComponent(S_TEXT_EMPLOYEE_FIELD) = oRequest("FormFieldText").Item
		Else
			aEmployeeFieldComponent(S_TEXT_EMPLOYEE_FIELD) = ""
		End If
	End If
	aEmployeeFieldComponent(S_TEXT_EMPLOYEE_FIELD) = Left(aEmployeeFieldComponent(S_TEXT_EMPLOYEE_FIELD), 255)

	If IsEmpty(aEmployeeFieldComponent(N_IS_OPTIONAL_EMPLOYEE_FIELD)) Then
		If Len(oRequest("IsOptional").Item) > 0 Then
			aEmployeeFieldComponent(N_IS_OPTIONAL_EMPLOYEE_FIELD) = CInt(oRequest("IsOptional").Item)
		Else
			aEmployeeFieldComponent(N_IS_OPTIONAL_EMPLOYEE_FIELD) = 0
		End If
	End If

	If IsEmpty(aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD)) Then
		If Len(oRequest("FieldTypeID").Item) > 0 Then
			aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD) = CLng(oRequest("FieldTypeID").Item)
		Else
			aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD) = 5
		End If
	End If

	If IsEmpty(aEmployeeFieldComponent(N_SIZE_EMPLOYEE_FIELD)) Then
		If aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD) = 5 Then
			If Len(oRequest("FormFieldSize").Item) > 0 Then
				aEmployeeFieldComponent(N_SIZE_EMPLOYEE_FIELD) = CInt(oRequest("FormFieldSize").Item)
			Else
				aEmployeeFieldComponent(N_SIZE_EMPLOYEE_FIELD) = 1
			End If
		Else
			aEmployeeFieldComponent(N_SIZE_EMPLOYEE_FIELD) = 1
		End If
	End If

	If IsEmpty(aEmployeeFieldComponent(N_LIMIT_TYPE_EMPLOYEE_FIELD)) Then
		If Len(oRequest("LimitTypeID").Item) > 0 Then
			aEmployeeFieldComponent(N_LIMIT_TYPE_EMPLOYEE_FIELD) = CInt(oRequest("LimitTypeID").Item)
		Else
			aEmployeeFieldComponent(N_LIMIT_TYPE_EMPLOYEE_FIELD) = 0
		End If
	End If

	If IsEmpty(aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD)) Then
		If aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD) = 1 Then
			If Len(oRequest("MinimumYearValue").Item) > 0 Then
				aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD) = CDbl(oRequest("MinimumYearValue").Item)
			Else
				aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD) = 0
			End If
		Else
			If aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD) = 5 Then
				If Len(oRequest("MinimumSize").Item) > 0 Then
					aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD) = CInt(oRequest("MinimumSize").Item)
				Else
					aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD) = 0
				End If
			ElseIf Len(oRequest("MinimumValue").Item) > 0 Then
				aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD) = CDbl(oRequest("MinimumValue").Item)
			Else
				aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD) = 0
			End If
		End If
	End If

	If IsEmpty(aEmployeeFieldComponent(N_MAXIMUM_EMPLOYEE_FIELD)) Then
		If aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD) = 1 Then
			If Len(oRequest("MaximumYearValue").Item) > 0 Then
				aEmployeeFieldComponent(N_MAXIMUM_EMPLOYEE_FIELD) = CDbl(oRequest("MaximumYearValue").Item)
			Else
				aEmployeeFieldComponent(N_MAXIMUM_EMPLOYEE_FIELD) = 0
			End If
		Else
			If Len(oRequest("MaximumValue").Item) > 0 Then
				aEmployeeFieldComponent(N_MAXIMUM_EMPLOYEE_FIELD) = CDbl(oRequest("MaximumValue").Item)
			Else
				aEmployeeFieldComponent(N_MAXIMUM_EMPLOYEE_FIELD) = 0
			End If
		End If
	End If

	If IsEmpty(aEmployeeFieldComponent(S_DSN_FOR_SOURCE_EMPLOYEE_FIELD)) Then
		If Len(oRequest("DSNForSource").Item) > 0 Then
			aEmployeeFieldComponent(S_DSN_FOR_SOURCE_EMPLOYEE_FIELD) = oRequest("DSNForSource").Item
		Else
			aEmployeeFieldComponent(S_DSN_FOR_SOURCE_EMPLOYEE_FIELD) = SIAP_DATABASE_PATH
		End If
	End If
	aEmployeeFieldComponent(S_DSN_FOR_SOURCE_EMPLOYEE_FIELD) = Left(aEmployeeFieldComponent(S_DSN_FOR_SOURCE_EMPLOYEE_FIELD), 255)

	If IsEmpty(aEmployeeFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_EMPLOYEE_FIELD)) Then
		If Len(oRequest("ConnectionTypeForSource").Item) > 0 Then
			aEmployeeFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_EMPLOYEE_FIELD) = CInt(oRequest("ConnectionTypeForSource").Item)
		Else
			aEmployeeFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_EMPLOYEE_FIELD) = iConnectionType
		End If
	End If

	If IsEmpty(aEmployeeFieldComponent(S_QUERY_FOR_SOURCE_EMPLOYEE_FIELD)) Then
		If Len(oRequest("QueryForSource").Item) > 0 Then
			aEmployeeFieldComponent(S_QUERY_FOR_SOURCE_EMPLOYEE_FIELD) = oRequest("QueryForSource").Item
		Else
			aEmployeeFieldComponent(S_QUERY_FOR_SOURCE_EMPLOYEE_FIELD) = ""
		End If
	End If
	aEmployeeFieldComponent(S_QUERY_FOR_SOURCE_EMPLOYEE_FIELD) = Left(aEmployeeFieldComponent(S_QUERY_FOR_SOURCE_EMPLOYEE_FIELD), 2000)

	If IsEmpty(aEmployeeFieldComponent(S_DEFAULT_VALUE_EMPLOYEE_FIELD)) Then
		If Len(oRequest("DefaultValue").Item) > 0 Then
			aEmployeeFieldComponent(S_DEFAULT_VALUE_EMPLOYEE_FIELD) = oRequest("DefaultValue").Item
		Else
			aEmployeeFieldComponent(S_DEFAULT_VALUE_EMPLOYEE_FIELD) = ""
		End If
	End If
	aEmployeeFieldComponent(S_DEFAULT_VALUE_EMPLOYEE_FIELD) = Left(aEmployeeFieldComponent(S_DEFAULT_VALUE_EMPLOYEE_FIELD), 255)

	If IsEmpty(aEmployeeFieldComponent(S_JAVASCRIPT_CODE_EMPLOYEE_FIELD)) Then
		If Len(oRequest("JavaScriptCode").Item) > 0 Then
			aEmployeeFieldComponent(S_JAVASCRIPT_CODE_EMPLOYEE_FIELD) = oRequest("JavaScriptCode").Item
		Else
			aEmployeeFieldComponent(S_JAVASCRIPT_CODE_EMPLOYEE_FIELD) = ""
		End If
	End If
	aEmployeeFieldComponent(S_JAVASCRIPT_CODE_EMPLOYEE_FIELD) = Left(aEmployeeFieldComponent(S_JAVASCRIPT_CODE_EMPLOYEE_FIELD), 255)

	If IsEmpty(aEmployeeFieldComponent(S_DESCRIPTION_EMPLOYEE_FIELD)) Then
		If Len(oRequest("FormFieldDescription").Item) > 0 Then
			aEmployeeFieldComponent(S_DESCRIPTION_EMPLOYEE_FIELD) = oRequest("FormFieldDescription").Item
		Else
			aEmployeeFieldComponent(S_DESCRIPTION_EMPLOYEE_FIELD) = ""
		End If
	End If
	aEmployeeFieldComponent(S_DESCRIPTION_EMPLOYEE_FIELD) = Left(aEmployeeFieldComponent(S_DESCRIPTION_EMPLOYEE_FIELD), 2000)

	aEmployeeFieldComponent(B_CHECK_FOR_DUPLICATED_EMPLOYEE_FIELD) = True
	aEmployeeFieldComponent(B_IS_DUPLICATED_EMPLOYEE_FIELD) = False

	aEmployeeFieldComponent(B_COMPONENT_INITIALIZED_EMPLOYEE_FIELD) = True
	InitializeEmployeeFieldComponent = Err.number
	Err.Clear
End Function

Function AddEmployeeField(oRequest, oADODBConnection, aEmployeeFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new form into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeFieldComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeField"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeFieldComponent(B_COMPONENT_INITIALIZED_EMPLOYEE_FIELD)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeFieldComponent(oRequest, aEmployeeFieldComponent)
	End If

	If aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para el nuevo campo."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "EmployeeFields", "FormFieldID", "", 1, aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD), sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		If aEmployeeFieldComponent(B_CHECK_FOR_DUPLICATED_EMPLOYEE_FIELD) Then
			lErrorNumber = CheckExistencyOfEmployeeField(oADODBConnection, aEmployeeFieldComponent, sErrorDescription)
			If aEmployeeFieldComponent(B_IS_DUPLICATED_EMPLOYEE_FIELD) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un campo con el nombre '" & aEmployeeFieldComponent(S_NAME_EMPLOYEE_FIELD) & "'."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeFieldComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			End If
		End If

		If lErrorNumber = 0 Then
			If Not CheckEmployeeFieldInformationConsistency(aEmployeeFieldComponent, sErrorDescription) Then
				lErrorNumber = -1
			Else
				sErrorDescription = "No se pudo guardar la información del nuevo campo."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeeFields (FormFieldID, FormFieldName, FormFieldText, IsOptional, FieldTypeID, FormFieldSize, LimitTypeID, MinimumValue, MaximumValue, DSNForSource, ConnectionTypeForSource, QueryForSource, DefaultValue, JavaScriptCode, FormFieldDescription) Values (" & aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD) & ", '" & Replace(aEmployeeFieldComponent(S_NAME_EMPLOYEE_FIELD), "'", "") & "', '" & Replace(aEmployeeFieldComponent(S_TEXT_EMPLOYEE_FIELD), "'", "´") & "', " & aEmployeeFieldComponent(N_IS_OPTIONAL_EMPLOYEE_FIELD) & ", " & aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD) & ", " & aEmployeeFieldComponent(N_SIZE_EMPLOYEE_FIELD) & ", " & aEmployeeFieldComponent(N_LIMIT_TYPE_EMPLOYEE_FIELD) & ", " & aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD) & ", " & aEmployeeFieldComponent(N_MAXIMUM_EMPLOYEE_FIELD) & ", '" & Replace(aEmployeeFieldComponent(S_DSN_FOR_SOURCE_EMPLOYEE_FIELD), "'", "") & "', " & aEmployeeFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_EMPLOYEE_FIELD) & ", '" & Replace(aEmployeeFieldComponent(S_QUERY_FOR_SOURCE_EMPLOYEE_FIELD), "'", "´") & "', '" & Replace(aEmployeeFieldComponent(S_DEFAULT_VALUE_EMPLOYEE_FIELD), "'", "´") & "', '" & Replace(aEmployeeFieldComponent(S_JAVASCRIPT_CODE_EMPLOYEE_FIELD), "'", "´") & "', '" & Replace(aEmployeeFieldComponent(S_DESCRIPTION_EMPLOYEE_FIELD), "'", "´") & "')", "EmployeeFieldComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
	End If

	AddEmployeeField = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeField(oRequest, oADODBConnection, aEmployeeFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a form from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeFieldComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeField"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeFieldComponent(B_COMPONENT_INITIALIZED_EMPLOYEE_FIELD)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeFieldComponent(oRequest, aEmployeeFieldComponent)
	End If

	If aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del campo para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeFieldComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del campo."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeeFields Where (FormFieldID=" & aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD) & ")", "EmployeeFieldComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El campo especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeFieldComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aEmployeeFieldComponent(S_NAME_EMPLOYEE_FIELD) = CStr(oRecordset.Fields("FormFieldName").Value)
				aEmployeeFieldComponent(S_TEXT_EMPLOYEE_FIELD) = CStr(oRecordset.Fields("FormFieldText").Value)
				aEmployeeFieldComponent(N_IS_OPTIONAL_EMPLOYEE_FIELD) = CInt(oRecordset.Fields("IsOptional").Value)
				aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD) = CLng(oRecordset.Fields("FieldTypeID").Value)
				aEmployeeFieldComponent(N_SIZE_EMPLOYEE_FIELD) = CInt(oRecordset.Fields("FormFieldSize").Value)
				aEmployeeFieldComponent(N_LIMIT_TYPE_EMPLOYEE_FIELD) = CInt(oRecordset.Fields("LimitTypeID").Value)
				aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD) = CDbl(oRecordset.Fields("MinimumValue").Value)
				aEmployeeFieldComponent(N_MAXIMUM_EMPLOYEE_FIELD) = CDbl(oRecordset.Fields("MaximumValue").Value)
				aEmployeeFieldComponent(S_DSN_FOR_SOURCE_EMPLOYEE_FIELD) = CStr(oRecordset.Fields("DSNForSource").Value)
				aEmployeeFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_EMPLOYEE_FIELD) = CLng(oRecordset.Fields("ConnectionTypeForSource").Value)
				aEmployeeFieldComponent(S_QUERY_FOR_SOURCE_EMPLOYEE_FIELD) = Replace(CStr(oRecordset.Fields("QueryForSource").Value), "´", "'")
				aEmployeeFieldComponent(S_DEFAULT_VALUE_EMPLOYEE_FIELD) = Replace(CStr(oRecordset.Fields("DefaultValue").Value), "´", "'")
				aEmployeeFieldComponent(S_JAVASCRIPT_CODE_EMPLOYEE_FIELD) = Replace(CStr(oRecordset.Fields("JavaScriptCode").Value), "´", "'")
				aEmployeeFieldComponent(S_DESCRIPTION_EMPLOYEE_FIELD) = CStr(oRecordset.Fields("FormFieldDescription").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetEmployeeField = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeFields(oRequest, oADODBConnection, aEmployeeFieldComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the forms from
'		  the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeFieldComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeFields"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeFieldComponent(B_COMPONENT_INITIALIZED_EMPLOYEE_FIELD)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeFieldComponent(oRequest, aEmployeeFieldComponent)
	End If

	sErrorDescription = "No se pudo obtener la información de los campos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeFields.*, FieldTypeName From EmployeeFields, FieldTypes Where (EmployeeFields.FieldTypeID=FieldTypes.FieldTypeID) And (FormFieldID > -1) Order By FormFieldName", "EmployeeFieldComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetEmployeeFields = lErrorNumber
	Err.Clear
End Function

Function ModifyEmployeeField(oRequest, oADODBConnection, aEmployeeFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing form in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeFieldComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyEmployeeField"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeFieldComponent(B_COMPONENT_INITIALIZED_EMPLOYEE_FIELD)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeFieldComponent(oRequest, aEmployeeFieldComponent)
	End If

	If aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del campo para modificar su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeFieldComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aEmployeeFieldComponent(B_CHECK_FOR_DUPLICATED_EMPLOYEE_FIELD) Then
			lErrorNumber = CheckExistencyOfEmployeeField(oADODBConnection, aEmployeeFieldComponent, sErrorDescription)
			If aEmployeeFieldComponent(B_IS_DUPLICATED_EMPLOYEE_FIELD) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un campo con el nombre '" & aEmployeeFieldComponent(S_NAME_EMPLOYEE_FIELD) & "'."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeFieldComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			End If
		End If

		If lErrorNumber = 0 Then
			If Not CheckEmployeeFieldInformationConsistency(aEmployeeFieldComponent, sErrorDescription) Then
				lErrorNumber = -1
			Else
				sErrorDescription = "No se pudo modificar la información del campo."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeeFields Set FormFieldName='" & Replace(aEmployeeFieldComponent(S_NAME_EMPLOYEE_FIELD), "'", "") & "', FormFieldText='" & Replace(aEmployeeFieldComponent(S_TEXT_EMPLOYEE_FIELD), "'", "´") & "', IsOptional=" & aEmployeeFieldComponent(N_IS_OPTIONAL_EMPLOYEE_FIELD) & ", FieldTypeID=" & aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD) & ", FormFieldSize=" & aEmployeeFieldComponent(N_SIZE_EMPLOYEE_FIELD) & ", LimitTypeID=" & aEmployeeFieldComponent(N_LIMIT_TYPE_EMPLOYEE_FIELD) & ", MinimumValue=" & aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD) & ", MaximumValue=" & aEmployeeFieldComponent(N_MAXIMUM_EMPLOYEE_FIELD) & ", DSNForSource='" & Replace(aEmployeeFieldComponent(S_DSN_FOR_SOURCE_EMPLOYEE_FIELD), "'", "") & "', ConnectionTypeForSource=" & aEmployeeFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_EMPLOYEE_FIELD) & ", QueryForSource='" & Replace(aEmployeeFieldComponent(S_QUERY_FOR_SOURCE_EMPLOYEE_FIELD), "'", "´") & "', DefaultValue='" & Replace(aEmployeeFieldComponent(S_DEFAULT_VALUE_EMPLOYEE_FIELD), "'", "´") & "', JavaScriptCode='" & Replace(aEmployeeFieldComponent(S_JAVASCRIPT_CODE_EMPLOYEE_FIELD), "'", "´") & "', FormFieldDescription='" & Replace(aEmployeeFieldComponent(S_DESCRIPTION_EMPLOYEE_FIELD), "'", "´") & "' Where (FormFieldID=" & aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD) & ")", "EmployeeFieldComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
	End If

	ModifyEmployeeField = lErrorNumber
	Err.Clear
End Function

Function RemoveEmployeeField(oRequest, oADODBConnection, aEmployeeFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a form from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeFieldComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployeeField"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeFieldComponent(B_COMPONENT_INITIALIZED_EMPLOYEE_FIELD)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeFieldComponent(oRequest, aEmployeeFieldComponent)
	End If

	If aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del campo a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeFieldComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del campo."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeeFields Where (FormFieldID=" & aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD) & ")", "EmployeeFieldComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la información del formulario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesInformation Where (FormFieldID=" & aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD) & ")", "EmployeeFieldComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	RemoveEmployeeField = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfEmployeeField(oADODBConnection, aEmployeeFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific form exists in the database
'Inputs:  oADODBConnection, aEmployeeFieldComponent
'Outputs: aEmployeeFieldComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfEmployeeField"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeFieldComponent(B_COMPONENT_INITIALIZED_EMPLOYEE_FIELD)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeFieldComponent(oRequest, aEmployeeFieldComponent)
	End If

	If (aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD) = -1) Or (Len(aEmployeeFieldComponent(S_NAME_EMPLOYEE_FIELD)) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el nombre del campo para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeFieldComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del campo en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeeFields Where (FormFieldID<>" & aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD) & ") And (FormFieldName='" & Replace(aEmployeeFieldComponent(S_NAME_EMPLOYEE_FIELD), "'", "") & "')", "EmployeeFieldComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			aEmployeeFieldComponent(B_IS_DUPLICATED_EMPLOYEE_FIELD) = (Not oRecordset.EOF)
		End If
	End If
	oRecordset.Close

	Set oRecordset = Nothing
	CheckExistencyOfEmployeeField = lErrorNumber
	Err.Clear
End Function

Function CheckEmployeeFieldInformationConsistency(aEmployeeFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aEmployeeFieldComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckEmployeeFieldInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del campo no es un valor numérico."
		bIsCorrect = False
	End If
	If Len(aEmployeeFieldComponent(S_NAME_EMPLOYEE_FIELD)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del campo está vacío."
		bIsCorrect = False
	End If
	If Len(aEmployeeFieldComponent(S_TEXT_EMPLOYEE_FIELD)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El descriptivo del campo está vacío."
		bIsCorrect = False
	End If
	If Not IsNumeric(aEmployeeFieldComponent(N_IS_OPTIONAL_EMPLOYEE_FIELD)) Then aEmployeeFieldComponent(N_IS_OPTIONAL_EMPLOYEE_FIELD) = 0
	If Not IsNumeric(aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD)) Then aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD) = 5
	If Not IsNumeric(aEmployeeFieldComponent(N_SIZE_EMPLOYEE_FIELD)) Then aEmployeeFieldComponent(N_SIZE_EMPLOYEE_FIELD) = 1
	If Not IsNumeric(aEmployeeFieldComponent(N_LIMIT_TYPE_EMPLOYEE_FIELD)) Then aEmployeeFieldComponent(N_LIMIT_TYPE_EMPLOYEE_FIELD) = 0
	If Not IsNumeric(aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD)) Then aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD) = 0
	If Not IsNumeric(aEmployeeFieldComponent(N_MAXIMUM_EMPLOYEE_FIELD)) Then aEmployeeFieldComponent(N_MAXIMUM_EMPLOYEE_FIELD) = 0
	If Len(aEmployeeFieldComponent(S_DSN_FOR_SOURCE_EMPLOYEE_FIELD)) = 0 Then aEmployeeFieldComponent(S_DSN_FOR_SOURCE_EMPLOYEE_FIELD) = SIAP_DATABASE_PATH
	If Not IsNumeric(aEmployeeFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_EMPLOYEE_FIELD)) Then aEmployeeFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_EMPLOYEE_FIELD) = iConnectionType

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del campo contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeFieldComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckEmployeeFieldInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayEmployeeFieldForm(oRequest, oADODBConnection, sAction, aEmployeeFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a form from the
'		  database using a HTML EmployeeField
'Inputs:  oRequest, oADODBConnection, sAction, aEmployeeFieldComponent
'Outputs: aEmployeeFieldComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFieldForm"
	Dim sNames
	Dim asFields
	Dim iIndex
	Dim lErrorNumber

	If aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD) <> -1 Then
		lErrorNumber = GetEmployeeField(oRequest, oADODBConnection, aEmployeeFieldComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckEmployeeFields(oForm) {" & vbNewLine
				If Len(oRequest("Delete").Item) = 0 Then
					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "if (oForm.FormFieldName.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir el nombre del campo.');" & vbNewLine
							Response.Write "oForm.FormFieldName.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (oForm.FormFieldText.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir el descriptivo del campo.');" & vbNewLine
							Response.Write "oForm.FormFieldName.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (! CheckIntegerValue(oForm.FormFieldSize, 'el tamaño del campo', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 1, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "if (oForm.MinimumSize.value != '') {" & vbNewLine
							Response.Write "if (! CheckIntegerValue(oForm.MinimumSize, 'el mínimo número de caracteres', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 1, 0))" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "if (parseInt(oForm.MinimumSize.value) > parseInt(oForm.FormFieldSize.value)) {" & vbNewLine
								Response.Write "alert('El número mínimo de caracteres no puede ser mayor al tamaño del campo.');" & vbNewLine
								Response.Write "oForm.MinimumSize.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (oForm.FieldTypeID.value != '1') {" & vbNewLine
							Response.Write "if (! CheckFloatValue(oForm.MinimumValue, 'el valor mínimo del campo', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "if (! CheckFloatValue(oForm.MaximumValue, 'el valor máximo del campo', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
								Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
							
						Response.Write "if (oForm.LimitTypeID.value == '1') {" & vbNewLine
							Response.Write "oForm.MinimumValue.value = oForm.MinimumYearValue.value;" & vbNewLine
							Response.Write "oForm.MaximumValue.value = oForm.MaximumYearValue.value;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (oForm.LimitTypeID.value >= '3')" & vbNewLine
							Response.Write "if (parseFloat(oForm.MinimumValue.value) > parseFloat(oForm.MaximumValue.value)) {" & vbNewLine
								Response.Write "alert('El valor mínimo del campo es mayor al valor máximo.');" & vbNewLine
								Response.Write "oForm.MinimumValue.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "if (oForm.DSNForSource.value.length == 0)" & vbNewLine
							Response.Write "oForm.DSNForSource.value = oForm.DSNForField.value;" & vbNewLine
						Response.Write "if (IsDisplayed(document.all['CatalogFieldDiv']))" & vbNewLine
							Response.Write "oForm.QueryForSource.value = oForm.Catalog_0.value + '" & LIST_SEPARATOR & "' + oForm.Catalog_1.value + '" & LIST_SEPARATOR & "' + oForm.Catalog_2.value + '" & LIST_SEPARATOR & "' + oForm.Catalog_3.value + '" & LIST_SEPARATOR & "' + oForm.Catalog_4.value + '" & LIST_SEPARATOR & "' + oForm.Catalog_5.value;" & vbNewLine
					Response.Write "}" & vbNewLine
				End If
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckEmployeeFields" & vbNewLine

			Response.Write "function ShowFieldsForFieldType(iFieldType) {" & vbNewLine
				Response.Write "HideDisplay(document.all['SizeDiv']);" & vbNewLine
				Response.Write "HideDisplay(document.all['MinimumSizeDiv']);" & vbNewLine
				Response.Write "HideDisplay(document.all['LimitTypeDiv']);" & vbNewLine
				Response.Write "HideDisplay(document.all['MinimumDiv']);" & vbNewLine
				Response.Write "HideDisplay(document.all['MaximumDiv']);" & vbNewLine
				Response.Write "HideDisplay(document.all['MinimumYearDiv']);" & vbNewLine
				Response.Write "HideDisplay(document.all['MaximumYearDiv']);" & vbNewLine
				Response.Write "ShowDisplay(document.all['NoCatalogFieldDiv']);" & vbNewLine
				Response.Write "HideDisplay(document.all['CatalogFieldDiv']);" & vbNewLine

				Response.Write "switch (iFieldType) {" & vbNewLine
					Response.Write "case '1':" & vbNewLine 'Fecha
						Response.Write "ShowDisplay(document.all['MinimumYearDiv']);" & vbNewLine
						Response.Write "ShowDisplay(document.all['MaximumYearDiv']);" & vbNewLine
						Response.Write "break;" & vbNewLine
					Response.Write "case '2':" & vbNewLine 'Flotante
						Response.Write "ShowDisplay(document.all['LimitTypeDiv']);" & vbNewLine
						Response.Write "ShowDisplay(document.all['MinimumDiv']);" & vbNewLine
						Response.Write "ShowDisplay(document.all['MaximumDiv']);" & vbNewLine
						Response.Write "break;" & vbNewLine
					Response.Write "case '4':" & vbNewLine 'Numérico
						Response.Write "ShowDisplay(document.all['LimitTypeDiv']);" & vbNewLine
						Response.Write "ShowDisplay(document.all['MinimumDiv']);" & vbNewLine
						Response.Write "ShowDisplay(document.all['MaximumDiv']);" & vbNewLine
						Response.Write "break;" & vbNewLine
					Response.Write "case '5':" & vbNewLine 'Texto
						Response.Write "ShowDisplay(document.all['SizeDiv']);" & vbNewLine
						Response.Write "ShowDisplay(document.all['MinimumSizeDiv']);" & vbNewLine
						Response.Write "break;" & vbNewLine
					Response.Write "case '6':" & vbNewLine 'Catálogo
						Response.Write "HideDisplay(document.all['NoCatalogFieldDiv']);" & vbNewLine
						Response.Write "ShowDisplay(document.all['CatalogFieldDiv']);" & vbNewLine
						Response.Write "HideDisplay(document.all['HierarchyFieldDiv']);" & vbNewLine
						Response.Write "break;" & vbNewLine
					Response.Write "case '7':" & vbNewLine 'Catálogo jerárquico
						Response.Write "HideDisplay(document.all['NoCatalogFieldDiv']);" & vbNewLine
						Response.Write "ShowDisplay(document.all['CatalogFieldDiv']);" & vbNewLine
						Response.Write "ShowDisplay(document.all['HierarchyFieldDiv']);" & vbNewLine
						Response.Write "break;" & vbNewLine
					Response.Write "case '8':" & vbNewLine 'Lista
						Response.Write "HideDisplay(document.all['NoCatalogFieldDiv']);" & vbNewLine
						Response.Write "ShowDisplay(document.all['CatalogFieldDiv']);" & vbNewLine
						Response.Write "HideDisplay(document.all['HierarchyFieldDiv']);" & vbNewLine
						Response.Write "break;" & vbNewLine
					Response.Write "case '9':" & vbNewLine 'Catálogo jerárquico
						Response.Write "HideDisplay(document.all['NoCatalogFieldDiv']);" & vbNewLine
						Response.Write "ShowDisplay(document.all['CatalogFieldDiv']);" & vbNewLine
						Response.Write "ShowDisplay(document.all['HierarchyFieldDiv']);" & vbNewLine
						Response.Write "break;" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of ShowFieldsForFieldType" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""EmployeeFieldFrm"" ID=""EmployeeFieldFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckEmployeeFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""EmployeeFields"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormFieldID"" ID=""FormFieldIDHdn"" VALUE=""" & aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD) & """ />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Nombre:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""FormFieldName"" ID=""FormFieldNameTxt"" SIZE=""35"" MAXLENGTH=""255"" VALUE=""" & aEmployeeFieldComponent(S_NAME_EMPLOYEE_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Descriptivo:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""FormFieldText"" ID=""FormFieldTextTxt"" SIZE=""35"" MAXLENGTH=""255"" VALUE=""" & aEmployeeFieldComponent(S_TEXT_EMPLOYEE_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">¿Es opcional?&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><SELECT NAME=""IsOptional"" ID=""IsOptionalCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE=""0"""
							If aEmployeeFieldComponent(N_IS_OPTIONAL_EMPLOYEE_FIELD) = 0 Then Response.Write " SELECTED=""1"""
						Response.Write ">No</OPTION>"
						Response.Write "<OPTION VALUE=""1"""
							If aEmployeeFieldComponent(N_IS_OPTIONAL_EMPLOYEE_FIELD) = 1 Then Response.Write " SELECTED=""1"""
						Response.Write ">Sí</OPTION>"
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Tipo:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><SELECT NAME=""FieldTypeID"" ID=""FieldTypeIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowFieldsForFieldType(this.value)"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "FieldTypes", "FieldTypeID", "FieldTypeName", "", "FieldTypeName", aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""SizeDiv"" ID=""SizeDiv"" STYLE=""display: none"">"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Tamaño:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""FormFieldSize"" ID=""FormFieldSizeTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & aEmployeeFieldComponent(N_SIZE_EMPLOYEE_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""MinimumSizeDiv"" ID=""MinimumSizeDiv"" STYLE=""display: none"">"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Mínimo número de caracteres:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""MinimumSize"" ID=""MinimumSizeTxt"" SIZE=""3"" MAXLENGTH=""3"" VALUE="""
						If aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD) > 0 Then Response.Write aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD)
					Response.Write """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""LimitTypeDiv"" ID=""LimitTypeDiv"" STYLE=""display: none"">"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Tipo de límite:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><SELECT NAME=""LimitTypeID"" ID=""LimitTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "LimitTypes", "LimitTypeID", "LimitTypeName", "", "LimitTypeID", aEmployeeFieldComponent(N_LIMIT_TYPE_EMPLOYEE_FIELD), "Ninguno;;;0", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""MinimumDiv"" ID=""MinimumDiv"" STYLE=""display: none"">"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Mínimo:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""MinimumValue"" ID=""MinimumValueTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""MaximumDiv"" ID=""MaximumDiv"" STYLE=""display: none"">"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Máximo:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""MaximumValue"" ID=""MaximumValueTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & aEmployeeFieldComponent(N_MAXIMUM_EMPLOYEE_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""MinimumYearDiv"" ID=""MinimumYearDiv"" STYLE=""display: none"">"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Mostrar desde el año:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><SELECT NAME=""MinimumYearValue"" ID=""MinimumYearValueCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE=""0"">Año en curso</OPTION>"
						For iIndex = Year(Date()) To N_FORM_START_YEAR Step -1
							Response.Write "<OPTION VALUE=""" & iIndex & """"
								If iIndex = aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD) Then Response.Write " SELECTED=""1"""
							Response.Write ">" & iIndex & "</OPTION>"
						Next
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""MaximumYearDiv"" ID=""MaximumYearDiv"" STYLE=""display: none"">"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Hasta el año:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><SELECT NAME=""MaximumYearValue"" ID=""MaximumYearValueCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE=""0"">Año en curso</OPTION>"
						For iIndex = 1 To 5
							Response.Write "<OPTION VALUE=""" & iIndex & """"
								If iIndex = Abs(aEmployeeFieldComponent(N_MAXIMUM_EMPLOYEE_FIELD)) Then Response.Write " SELECTED=""1"""
							Response.Write ">Año en curso +" & iIndex & "</OPTION>"
						Next
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">DSN de origen de los datos:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""DSNForSource"" ID=""DSNForSourceTxt"" SIZE=""35"" MAXLENGTH=""255"" VALUE=""" & aEmployeeFieldComponent(S_DSN_FOR_SOURCE_EMPLOYEE_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Tipo de conexión con el origen de los datos:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><SELECT NAME=""ConnectionTypeForSource"" ID=""ConnectionTypeForSourceCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE=""" & SQL_SERVER & """"
							If aEmployeeFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_EMPLOYEE_FIELD) = SQL_SERVER Then Response.Write " SELECTED=""1"""
						Response.Write ">SQL Server</OPTION>"
						Response.Write "<OPTION VALUE=""" & ACCESS & """"
							If aEmployeeFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_EMPLOYEE_FIELD) = ACCESS Then Response.Write " SELECTED=""1"""
						Response.Write ">MS Access File</OPTION>"
						Response.Write "<OPTION VALUE=""" & ACCESS_DSN & """"
							If aEmployeeFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_EMPLOYEE_FIELD) = ACCESS_DSN Then Response.Write " SELECTED=""1"""
						Response.Write ">MS Access (DSN)</OPTION>"
						Response.Write "<OPTION VALUE=""" & ORACLE & """"
							If aEmployeeFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_EMPLOYEE_FIELD) = ORACLE Then Response.Write " SELECTED=""1"""
						Response.Write ">Oracle</OPTION>"
						Response.Write "<OPTION VALUE=""" & MYSQL & """"
							If aEmployeeFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_EMPLOYEE_FIELD) = MYSQL Then Response.Write " SELECTED=""1"""
						Response.Write ">MySQL</OPTION>"
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""NoCatalogFieldDiv"" ID=""NoCatalogFieldDiv"""
					If (aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD) = N_CATALOG) Or (aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD) = N_LIST) Then Response.Write " STYLE=""display: none"""
				Response.Write "><TD COLSPAN=""2"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Query de origen de los datos:</FONT><BR />"
					Response.Write "<SELECT onChange=""document.EmployeeFieldFrm.QueryForSource.value=this.value"">"
						Response.Write "<OPTION VALUE=""""></OPTION>"
						Response.Write "<OPTION VALUE=""Select EmployeeName From Employees Where EmployeeID=<EMPLOYEE_ID />"">Nombre del empleado</OPTION>"
						Response.Write "<OPTION VALUE=""Select EmployeeLastName From Employees Where EmployeeID=<EMPLOYEE_ID />"">Apellido paterno</OPTION>"
						Response.Write "<OPTION VALUE=""Select EmployeeLastName2 From Employees Where EmployeeID=<EMPLOYEE_ID />"">Apellido materno</OPTION>"
						Response.Write "<OPTION VALUE=""Select EmployeeName, EmployeeLastName, EmployeeLastName2 From Employees Where EmployeeID=<EMPLOYEE_ID />"">Nombre y apellidos del usuario</OPTION>"
						Response.Write "<OPTION VALUE=""Select EmployeeNumber From Employees Where EmployeeID=<EMPLOYEE_ID />"">Número del empleado</OPTION>"
						Response.Write "<OPTION VALUE=""Select JobNumber From Jobs Where JobID=<EMPLOYEE_JOB_ID />"">Número de la plaza</OPTION>"
						Response.Write "<OPTION VALUE=""Select PaymentCenterShortName From PaymentCenters Where PaymentCenterID=<EMPLOYEE_PAYMENT_CENTER_ID />"">Centro de pago</OPTION>"
						Response.Write "<OPTION VALUE=""Select SocialSecurityNumber From Employees Where EmployeeID=<EMPLOYEE_ID />"">Número de seguro social</OPTION>"
						Response.Write "<OPTION VALUE=""Select ClassificationID From Employees Where EmployeeID=<EMPLOYEE_ID />"">Clasificación</OPTION>"
						Response.Write "<OPTION VALUE=""Select IntegrationID From Employees Where EmployeeID=<EMPLOYEE_ID />"">Integración</OPTION>"
						Response.Write "<OPTION VALUE=""Select StartHour1 From Employees Where EmployeeID=<EMPLOYEE_ID />"">Hora de entrada 1</OPTION>"
						Response.Write "<OPTION VALUE=""Select EndHour1 From Employees Where EmployeeID=<EMPLOYEE_ID />"">Hora de salida 1</OPTION>"
						Response.Write "<OPTION VALUE=""Select StartHour2 From Employees Where EmployeeID=<EMPLOYEE_ID />"">Hora de entrada 2</OPTION>"
						Response.Write "<OPTION VALUE=""Select EndHour2 From Employees Where EmployeeID=<EMPLOYEE_ID />"">Hora de salida 2</OPTION>"
						Response.Write "<OPTION VALUE=""Select StartHour3 From Employees Where EmployeeID=<EMPLOYEE_ID />"">Hora de entrada turno opcional</OPTION>"
						Response.Write "<OPTION VALUE=""Select EndHour3 From Employees Where EmployeeID=<EMPLOYEE_ID />"">Hora de salida turno opcional</OPTION>"
						Response.Write "<OPTION VALUE=""Select WorkingHours From Employees Where EmployeeID=<EMPLOYEE_ID />"">Horas laboradas</OPTION>"
						Response.Write "<OPTION VALUE=""Select BirthDate From Employees Where EmployeeID=<EMPLOYEE_ID />"">Fecha de nacimiento</OPTION>"
						Response.Write "<OPTION VALUE=""Select StartDate From Employees Where EmployeeID=<EMPLOYEE_ID />"">Fecha de inicio en el ISSSTE</OPTION>"
						Response.Write "<OPTION VALUE=""Select StartDate2 From Employees Where EmployeeID=<EMPLOYEE_ID />"">Fecha de inicio en gobierno</OPTION>"
						Response.Write "<OPTION VALUE=""Select RFC From Employees Where EmployeeID=<EMPLOYEE_ID />"">RFC del empleado</OPTION>"
						Response.Write "<OPTION VALUE=""Select CURP From Employees Where EmployeeID=<EMPLOYEE_ID />"">CURP del empleado</OPTION>"
						Response.Write "<OPTION VALUE=""Select EmployeeEmail From Employees Where EmployeeID=<EMPLOYEE_ID />"">Correo electrónico del empleado</OPTION>"
						Response.Write "<OPTION VALUE=""Select Active From Employees Where EmployeeID=<EMPLOYEE_ID />"">¿El empleado está activo?</OPTION>"
					Response.Write "</SELECT><BR />"
					Response.Write "<TEXTAREA NAME=""QueryForSource"" ID=""QueryForSourceTxtArea"" ROWS=""5"" COLS=""40"" MAXLENGTH=""4000"" CLASS=""TextFields"">" & aEmployeeFieldComponent(S_QUERY_FOR_SOURCE_EMPLOYEE_FIELD) & "</TEXTAREA><BR />"
				Response.Write "</TD></TR>"
				Response.Write "<TR NAME=""CatalogFieldDiv"" ID=""CatalogFieldDiv"""
					If (aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD) <> N_CATALOG) And (aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD) <> N_LIST) Then Response.Write " STYLE=""display: none"""
				Response.Write "><TD COLSPAN=""2"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Query de origen de los datos:</FONT><BR />"
					Response.Write "<SELECT onChange=""if (this.value != '') {var sFieldTemp = this.value.split('" & LIST_SEPARATOR & "'); document.EmployeeFieldFrm.Catalog_0.value=sFieldTemp[0]; document.EmployeeFieldFrm.Catalog_1.value=sFieldTemp[1]; document.EmployeeFieldFrm.Catalog_2.value=sFieldTemp[2]; document.EmployeeFieldFrm.Catalog_3.value=sFieldTemp[3]; document.EmployeeFieldFrm.Catalog_4.value=sFieldTemp[4]; document.EmployeeFieldFrm.Catalog_5.value=sFieldTemp[5];}"">"
						Response.Write "<OPTION VALUE=""""></OPTION>"
						Response.Write "<OPTION VALUE=""Companies;;;CompanyID;;;;;;CompanyName;;;;;;CompanyName;;;<EMPLOYEE_COMPANY_ID />"">Catálogo de compañías</OPTION>"
						Response.Write "<OPTION VALUE=""States;;;StateID;;;;;;StateName;;;;;;StateName"">Catálogo de estados</OPTION>"
						Response.Write "<OPTION VALUE=""StatusEmployees;;;StatusID;;;;;;StatusName;;;;;;StatusName;;;<EMPLOYEE_STATUS_ID />"">Catálogo de estatus de empleados</OPTION>"
						Response.Write "<OPTION VALUE=""Genders;;;GenderID;;;;;;GenderName;;;;;;GenderName;;;<EMPLOYEE_GENDER_ID />"">Catálogo de géneros</OPTION>"
						Response.Write "<OPTION VALUE=""GroupGradeLevels;;;GroupGradeLevelID;;;;;;GroupGradeLevelName;;;;;;GroupGradeLevelName;;;<EMPLOYEE_GROUP_GRADE_LEVEL_ID />"">Catálogo de grupo, grado, nivel</OPTION>"
						Response.Write "<OPTION VALUE=""Journeys;;;JourneyID;;;;;;JourneyName;;;;;;JourneyName;;;<EMPLOYEE_JOURNEY_ID />"">Catálogo de jornadas</OPTION>"
						Response.Write "<OPTION VALUE=""Levels;;;LevelID;;;;;;LevelName;;;;;;LevelName;;;<EMPLOYEE_LEVEL_ID />"">Catálogo de niveles</OPTION>"
						Response.Write "<OPTION VALUE=""Countries;;;CountryID;;;;;;CountryName;;;;;;CountryName;;;<EMPLOYEE_COUNTRY_ID />"">Catálogo de países</OPTION>"
						Response.Write "<OPTION VALUE=""EmployeeTypes;;;EmployeeTypeID;;;;;;EmployeeTypeName;;;;;;EmployeeTypeName;;;<EMPLOYEE_TYPE_ID />"">Catálogo de tipos de empleado</OPTION>"
						Response.Write "<OPTION VALUE=""PositionTypes;;;PositionTypeID;;;;;;PositionTypeName;;;;;;PositionTypeName;;;<EMPLOYEE_POSITION_TYPE_ID />"">Catálogo de tipos de puesto</OPTION>"
						Response.Write "<OPTION VALUE=""Shifts;;;ShiftID;;;;;;ShiftName;;;;;;ShiftName;;;<EMPLOYEE_SHIFT_ID />"">Catálogo de turnos</OPTION>"
					Response.Write "</SELECT><BR /><FONT FACE=""Arial"" SIZE=""2"">"
					If InStr(1, aEmployeeFieldComponent(S_QUERY_FOR_SOURCE_EMPLOYEE_FIELD), LIST_SEPARATOR, vbBinaryCompare) = 0 Then asFields = Split(BuildList("", LIST_SEPARATOR, 6), LIST_SEPARATOR) 
					asFields = Split(aEmployeeFieldComponent(S_QUERY_FOR_SOURCE_EMPLOYEE_FIELD), LIST_SEPARATOR, -1, vbBinaryCompare)
					If UBound(asFields) < 5 Then asFields = Split(JoinLists(aEmployeeFieldComponent(S_QUERY_FOR_SOURCE_EMPLOYEE_FIELD), BuildList("", LIST_SEPARATOR, 6 - UBound(asFields)), LIST_SEPARATOR), LIST_SEPARATOR, -1, vbBinaryCompare)
					Response.Write "Tabla:<IMG SRC=""Images/Transparent.gif"" WIDTH=""50"" HEIGHT=""1"" /><INPUT TYPE=""TEXT"" NAME=""Catalog_0"" ID=""Catalog_0Txt"" SIZE=""20"" MAXLENGTH=""100"" VALUE=""" & asFields(0) & """ CLASS=""TextFields"" /><BR />"
					Response.Write "Campo llave:<IMG SRC=""Images/Transparent.gif"" WIDTH=""11"" HEIGHT=""1"" /><INPUT TYPE=""TEXT"" NAME=""Catalog_1"" ID=""Catalog_1Txt"" SIZE=""20"" MAXLENGTH=""100"" VALUE=""" & asFields(1) & """ CLASS=""TextFields"" /><BR />"
					Response.Write "<SPAN NAME=""HierarchyFieldDiv"" ID=""HierarchyFieldDiv"">Campo padre: <INPUT TYPE=""TEXT"" NAME=""Catalog_2"" ID=""Catalog_2Txt"" SIZE=""20"" MAXLENGTH=""100"" VALUE=""" & asFields(2) & """ CLASS=""TextFields"" /><BR /></SPAN>"
					Response.Write "Campo:<IMG SRC=""Images/Transparent.gif"" WIDTH=""40"" HEIGHT=""1"" /><INPUT TYPE=""TEXT"" NAME=""Catalog_3"" ID=""Catalog_3Txt"" SIZE=""20"" MAXLENGTH=""100"" VALUE=""" & asFields(3) & """ CLASS=""TextFields"" /><BR />"
					Response.Write "Condición:<IMG SRC=""Images/Transparent.gif"" WIDTH=""24"" HEIGHT=""1"" /><INPUT TYPE=""TEXT"" NAME=""Catalog_4"" ID=""Catalog_4Txt"" SIZE=""20"" MAXLENGTH=""100"" VALUE=""" & asFields(4) & """ CLASS=""TextFields"" /><BR />"
					Response.Write "Ordenar por:<IMG SRC=""Images/Transparent.gif"" WIDTH=""13"" HEIGHT=""1"" /><INPUT TYPE=""TEXT"" NAME=""Catalog_5"" ID=""Catalog_5Txt"" SIZE=""20"" MAXLENGTH=""100"" VALUE=""" & asFields(5) & """ CLASS=""TextFields"" /><BR />"
				Response.Write "</FONT></TD></TR>"
				Response.Write "<TR><TD COLSPAN=""2"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Valor predeterminado:&nbsp;</FONT>"
					Response.Write "<INPUT TYPE=""TEXT"" NAME=""DefaultValue"" ID=""DefaultValueTxt"" SIZE=""46"" MAXLENGTH=""255"" VALUE=""" & aEmployeeFieldComponent(S_DEFAULT_VALUE_EMPLOYEE_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR><TD COLSPAN=""2"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">JavaScript:<BR /></FONT>"
					Response.Write "<INPUT TYPE=""TEXT"" NAME=""JavaScriptCode"" ID=""JavaScriptCodeTxt"" SIZE=""59"" MAXLENGTH=""255"" VALUE=""" & aEmployeeFieldComponent(S_JAVASCRIPT_CODE_EMPLOYEE_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR><TD COLSPAN=""2"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Descripción:</FONT><BR />"
					Response.Write "<TEXTAREA NAME=""FormFieldDescription"" ID=""FormFieldDescriptionTxtArea"" ROWS=""5"" COLS=""40"" MAXLENGTH=""4000"" CLASS=""TextFields"">" & aEmployeeFieldComponent(S_DESCRIPTION_EMPLOYEE_FIELD) & "</TEXTAREA><BR />"
				Response.Write "</TD></TR>"
			Response.Write "</TABLE>"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "ShowFieldsForFieldType(document.EmployeeFieldFrm.FieldTypeID.value);" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine

			Response.Write "<BR />"
			If Len(oRequest("Change").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveEmployeeFieldWngDiv']); EmployeeFieldFrm.Remove.focus()"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=EmployeeFields&FormID=" & oRequest("FormID").Item & "'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveEmployeeFieldWngDiv", "¿Está seguro que desea borrar el campo de la &nbsp;base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayEmployeeFieldForm = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeeFieldAsHiddenFields(oRequest, oADODBConnection, aEmployeeFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a form using
'		  hidden form fields
'Inputs:  oRequest, oADODBConnection, aEmployeeFieldComponent
'Outputs: aEmployeeFieldComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFieldAsHiddenFields"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormFieldID"" ID=""FormFieldIDHdn"" VALUE=""" & aEmployeeFieldComponent(N_ID_EMPLOYEE_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormFieldName"" ID=""FormFieldNameHdn"" VALUE=""" & aEmployeeFieldComponent(S_NAME_EMPLOYEE_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormFieldText"" ID=""FormFieldTextHdn"" VALUE=""" & aEmployeeFieldComponent(S_TEXT_EMPLOYEE_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IsOptional"" ID=""IsOptionalHdn"" VALUE=""" & aEmployeeFieldComponent(N_IS_OPTIONAL_EMPLOYEE_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FieldTypeID"" ID=""FieldTypeIDHdn"" VALUE=""" & aEmployeeFieldComponent(N_TYPE_ID_EMPLOYEE_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormFieldSize"" ID=""FormFieldSizeHdn"" VALUE=""" & aEmployeeFieldComponent(N_SIZE_EMPLOYEE_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""LimitTypeID"" ID=""LimitTypeIDHdn"" VALUE=""" & aEmployeeFieldComponent(N_LIMIT_TYPE_EMPLOYEE_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""MinimumValue"" ID=""MinimumValueHdn"" VALUE=""" & aEmployeeFieldComponent(N_MINIMUM_EMPLOYEE_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""MaximumValue"" ID=""MaximumValueHdn"" VALUE=""" & aEmployeeFieldComponent(N_MAXIMUM_EMPLOYEE_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DSNForSource"" ID=""DSNForSourceHdn"" VALUE=""" & aEmployeeFieldComponent(S_DSN_FOR_SOURCE_EMPLOYEE_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConnectionTypeForSource"" ID=""ConnectionTypeForSourceHdn"" VALUE=""" & aEmployeeFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_EMPLOYEE_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""QueryForSource"" ID=""QueryForSourceHdn"" VALUE=""" & aEmployeeFieldComponent(S_QUERY_FOR_SOURCE_EMPLOYEE_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DefaultValue"" ID=""DefaultValueHdn"" VALUE=""" & aEmployeeFieldComponent(S_DEFAULT_VALUE_EMPLOYEE_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JavaScriptCode"" ID=""JavaScriptCodeHdn"" VALUE=""" & aEmployeeFieldComponent(S_JAVASCRIPT_CODE_EMPLOYEE_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormFieldDescription"" ID=""FormFieldDescriptionHdn"" VALUE=""" & aEmployeeFieldComponent(S_DESCRIPTION_EMPLOYEE_FIELD) & """ />"

	DisplayEmployeeFieldAsHiddenFields = Err.number
	Err.Clear
End Function

Function DisplayEmployeeFieldsTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, aEmployeeFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the forms from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aEmployeeFieldComponent
'Outputs: aEmployeeFieldComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeFieldsTable"
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sBoldBegin
	Dim sBoldEnd
	Dim lErrorNumber

	lErrorNumber = GetEmployeeFields(oRequest, oADODBConnection, aEmployeeFieldComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""400"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
					asColumnsTitles = Split("&nbsp;,Campo,Texto,Tipo,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,150,150,100,80", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("&nbsp;,Tabla,Campo,Tipo", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,180,180,120", ",", -1, vbBinaryCompare)
				End If
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				asCellAlignments = Split(",,,,CENTER", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("FormFieldID").Value), oRequest("FormFieldID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""FormFieldID"" ID=""FormFieldIDRd"" VALUE=""" & CStr(oRecordset.Fields("FormFieldID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""FormFieldID"" ID=""FormFieldIDChk"" VALUE=""" & CStr(oRecordset.Fields("FormFieldID").Value) & """ />"
						Case Else
							sRowContents = sRowContents & "&nbsp;"
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("FormFieldName").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("FormFieldText").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("FieldTypeName").Value))
						If CInt(oRecordset.Fields("IsOptional").Value) = 0 Then sRowContents = sRowContents & " *"
					sRowContents = sRowContents & sBoldEnd
					If bUseLinks And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
						sRowContents = sRowContents & TABLE_SEPARATOR
							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=EmployeeFields&FormFieldID=" & CStr(oRecordset.Fields("FormFieldID").Value) & "&Change=1"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>"
							End If

							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS Then
								sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=EmployeeFields&FormFieldID=" & CStr(oRecordset.Fields("FormFieldID").Value) & "&Delete=1"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>"
							End If
						sRowContents = sRowContents & "&nbsp;"
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen campos registrados en la base de datos."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayEmployeeFieldsTable = lErrorNumber
	Err.Clear
End Function
%>