<%
Const N_ID_FORM_FIELD = 0
Const N_FIELD_ID_FORM_FIELD = 1
Const S_DSN_FORM_FIELD = 2
Const N_CONNECTION_TYPE_FORM_FIELD = 3
Const S_DATABASE_NAME_FORM_FIELD = 4
Const S_TABLE_NAME_FORM_FIELD = 5
Const S_NAME_FORM_FIELD = 6
Const S_TEXT_FORM_FIELD = 7
Const N_IS_OPTIONAL_FORM_FIELD  = 8
Const N_TYPE_ID_FORM_FIELD = 9
Const N_SIZE_FORM_FIELD = 10
Const N_LIMIT_TYPE_FORM_FIELD = 11
Const N_MINIMUM_FORM_FIELD = 12
Const N_MAXIMUM_FORM_FIELD = 13
Const S_DSN_FOR_SOURCE_FORM_FIELD = 14
Const N_CONNECTION_TYPE_FOR_SOURCE_FORM_FIELD = 15
Const S_QUERY_FOR_SOURCE_FORM_FIELD = 16
Const S_DEFAULT_VALUE_FORM_FIELD = 17
Const S_JAVASCRIPT_CODE_FORM_FIELD = 18
Const S_DESCRIPTION_FORM_FIELD = 19
Const B_CHECK_FOR_DUPLICATED_FORM_FIELD = 20
Const B_IS_DUPLICATED_FORM_FIELD = 21
Const B_COMPONENT_INITIALIZED_FORM_FIELD = 22

Const N_FORM_FIELD_COMPONENT_SIZE = 22

Dim aFormFieldComponent()
Redim aFormFieldComponent(N_FORM_FIELD_COMPONENT_SIZE)

Function InitializeFormFieldComponent(oRequest, aFormFieldComponent)
'************************************************************
'Purpose: To initialize the empty elements of the FormField Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aFormFieldComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeFormFieldComponent"
	Redim Preserve aFormFieldComponent(N_FORM_FIELD_COMPONENT_SIZE)
	Dim aTemp

	If Len(oRequest("FormFieldIdentificator").Item) > 0 Then
		aTemp = Split(oRequest("FormFieldIdentificator").Item, ",", 2, vbBinaryCompare)
		If IsEmpty(aFormFieldComponent(N_ID_FORM_FIELD)) Then
			aFormFieldComponent(N_ID_FORM_FIELD) = CLng(aTemp(0))
		End If

		If IsEmpty(aFormFieldComponent(N_FIELD_ID_FORM_FIELD)) Then
			aFormFieldComponent(N_FIELD_ID_FORM_FIELD) = CLng(aTemp(1))
		End If
	Else
		If IsEmpty(aFormFieldComponent(N_ID_FORM_FIELD)) Then
			If Len(oRequest("FormID").Item) > 0 Then
				aFormFieldComponent(N_ID_FORM_FIELD) = CLng(oRequest("FormID").Item)
			Else
				aFormFieldComponent(N_ID_FORM_FIELD) = -1
			End If
		End If

		If IsEmpty(aFormFieldComponent(N_FIELD_ID_FORM_FIELD)) Then
			If Len(oRequest("FormFieldID").Item) > 0 Then
				aFormFieldComponent(N_FIELD_ID_FORM_FIELD) = CLng(oRequest("FormFieldID").Item)
			Else
				aFormFieldComponent(N_FIELD_ID_FORM_FIELD) = -1
			End If
		End If
	End If

	If IsEmpty(aFormFieldComponent(S_DSN_FORM_FIELD)) Then
		If Len(oRequest("DSNForField").Item) > 0 Then
			aFormFieldComponent(S_DSN_FORM_FIELD) = oRequest("DSNForField").Item
		Else
			aFormFieldComponent(S_DSN_FORM_FIELD) = SIAP_DATABASE_PATH
		End If
	End If
	aFormFieldComponent(S_DSN_FORM_FIELD) = Left(aFormFieldComponent(S_DSN_FORM_FIELD), 255)

	If IsEmpty(aFormFieldComponent(N_CONNECTION_TYPE_FORM_FIELD)) Then
		If Len(oRequest("ConnectionType").Item) > 0 Then
			aFormFieldComponent(N_CONNECTION_TYPE_FORM_FIELD) = CInt(oRequest("ConnectionType").Item)
		Else
			aFormFieldComponent(N_CONNECTION_TYPE_FORM_FIELD) = iConnectionType
		End If
	End If

	If IsEmpty(aFormFieldComponent(S_DATABASE_NAME_FORM_FIELD)) Then
		If Len(oRequest("DatabaseName").Item) > 0 Then
			aFormFieldComponent(S_DATABASE_NAME_FORM_FIELD) = oRequest("DatabaseName").Item
		Else
			aFormFieldComponent(S_DATABASE_NAME_FORM_FIELD) = SIAP_DATABASE_NAME
		End If
	End If
	aFormFieldComponent(S_DATABASE_NAME_FORM_FIELD) = Left(aFormFieldComponent(S_DATABASE_NAME_FORM_FIELD), 255)

	If IsEmpty(aFormFieldComponent(S_TABLE_NAME_FORM_FIELD)) Then
		If Len(oRequest("TableName").Item) > 0 Then
			aFormFieldComponent(S_TABLE_NAME_FORM_FIELD) = oRequest("TableName").Item
		Else
			aFormFieldComponent(S_TABLE_NAME_FORM_FIELD) = "FormFields"
		End If
	End If
	aFormFieldComponent(S_TABLE_NAME_FORM_FIELD) = Left(aFormFieldComponent(S_TABLE_NAME_FORM_FIELD), 255)

	If IsEmpty(aFormFieldComponent(S_NAME_FORM_FIELD)) Then
		If Len(oRequest("FormFieldName").Item) > 0 Then
			aFormFieldComponent(S_NAME_FORM_FIELD) = oRequest("FormFieldName").Item
		Else
			aFormFieldComponent(S_NAME_FORM_FIELD) = ""
		End If
	End If
	aFormFieldComponent(S_NAME_FORM_FIELD) = Left(aFormFieldComponent(S_NAME_FORM_FIELD), 255)

	If IsEmpty(aFormFieldComponent(S_TEXT_FORM_FIELD)) Then
		If Len(oRequest("FormFieldText").Item) > 0 Then
			aFormFieldComponent(S_TEXT_FORM_FIELD) = oRequest("FormFieldText").Item
		Else
			aFormFieldComponent(S_TEXT_FORM_FIELD) = ""
		End If
	End If
	aFormFieldComponent(S_TEXT_FORM_FIELD) = Left(aFormFieldComponent(S_TEXT_FORM_FIELD), 255)

	If IsEmpty(aFormFieldComponent(N_IS_OPTIONAL_FORM_FIELD)) Then
		If Len(oRequest("IsOptional").Item) > 0 Then
			aFormFieldComponent(N_IS_OPTIONAL_FORM_FIELD) = CInt(oRequest("IsOptional").Item)
		Else
			aFormFieldComponent(N_IS_OPTIONAL_FORM_FIELD) = 0
		End If
	End If

	If IsEmpty(aFormFieldComponent(N_TYPE_ID_FORM_FIELD)) Then
		If Len(oRequest("FieldTypeID").Item) > 0 Then
			aFormFieldComponent(N_TYPE_ID_FORM_FIELD) = CLng(oRequest("FieldTypeID").Item)
		Else
			aFormFieldComponent(N_TYPE_ID_FORM_FIELD) = 5
		End If
	End If

	If IsEmpty(aFormFieldComponent(N_SIZE_FORM_FIELD)) Then
		If aFormFieldComponent(N_TYPE_ID_FORM_FIELD) = 5 Then
			If Len(oRequest("FormFieldSize").Item) > 0 Then
				aFormFieldComponent(N_SIZE_FORM_FIELD) = CInt(oRequest("FormFieldSize").Item)
			Else
				aFormFieldComponent(N_SIZE_FORM_FIELD) = 1
			End If
		Else
			aFormFieldComponent(N_SIZE_FORM_FIELD) = 1
		End If
	End If

	If IsEmpty(aFormFieldComponent(N_LIMIT_TYPE_FORM_FIELD)) Then
		If Len(oRequest("LimitTypeID").Item) > 0 Then
			aFormFieldComponent(N_LIMIT_TYPE_FORM_FIELD) = CInt(oRequest("LimitTypeID").Item)
		Else
			aFormFieldComponent(N_LIMIT_TYPE_FORM_FIELD) = 0
		End If
	End If

	If IsEmpty(aFormFieldComponent(N_MINIMUM_FORM_FIELD)) Then
		If aFormFieldComponent(N_TYPE_ID_FORM_FIELD) = 1 Then
			If Len(oRequest("MinimumYearValue").Item) > 0 Then
				aFormFieldComponent(N_MINIMUM_FORM_FIELD) = CDbl(oRequest("MinimumYearValue").Item)
			Else
				aFormFieldComponent(N_MINIMUM_FORM_FIELD) = 0
			End If
		Else
			If aFormFieldComponent(N_TYPE_ID_FORM_FIELD) = 5 Then
				If Len(oRequest("MinimumSize").Item) > 0 Then
					aFormFieldComponent(N_MINIMUM_FORM_FIELD) = CInt(oRequest("MinimumSize").Item)
				Else
					aFormFieldComponent(N_MINIMUM_FORM_FIELD) = 0
				End If
			ElseIf Len(oRequest("MinimumValue").Item) > 0 Then
				aFormFieldComponent(N_MINIMUM_FORM_FIELD) = CDbl(oRequest("MinimumValue").Item)
			Else
				aFormFieldComponent(N_MINIMUM_FORM_FIELD) = 0
			End If
		End If
	End If

	If IsEmpty(aFormFieldComponent(N_MAXIMUM_FORM_FIELD)) Then
		If aFormFieldComponent(N_TYPE_ID_FORM_FIELD) = 1 Then
			If Len(oRequest("MaximumYearValue").Item) > 0 Then
				aFormFieldComponent(N_MAXIMUM_FORM_FIELD) = CDbl(oRequest("MaximumYearValue").Item)
			Else
				aFormFieldComponent(N_MAXIMUM_FORM_FIELD) = 0
			End If
		Else
			If Len(oRequest("MaximumValue").Item) > 0 Then
				aFormFieldComponent(N_MAXIMUM_FORM_FIELD) = CDbl(oRequest("MaximumValue").Item)
			Else
				aFormFieldComponent(N_MAXIMUM_FORM_FIELD) = 0
			End If
		End If
	End If

	If IsEmpty(aFormFieldComponent(S_DSN_FOR_SOURCE_FORM_FIELD)) Then
		If Len(oRequest("DSNForSource").Item) > 0 Then
			aFormFieldComponent(S_DSN_FOR_SOURCE_FORM_FIELD) = oRequest("DSNForSource").Item
		Else
			aFormFieldComponent(S_DSN_FOR_SOURCE_FORM_FIELD) = aFormFieldComponent(S_DSN_FORM_FIELD)
		End If
	End If
	aFormFieldComponent(S_DSN_FOR_SOURCE_FORM_FIELD) = Left(aFormFieldComponent(S_DSN_FOR_SOURCE_FORM_FIELD), 255)

	If IsEmpty(aFormFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_FORM_FIELD)) Then
		If Len(oRequest("ConnectionTypeForSource").Item) > 0 Then
			aFormFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_FORM_FIELD) = CInt(oRequest("ConnectionTypeForSource").Item)
		Else
			aFormFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_FORM_FIELD) = iConnectionType
		End If
	End If

	If IsEmpty(aFormFieldComponent(S_QUERY_FOR_SOURCE_FORM_FIELD)) Then
		If Len(oRequest("QueryForSource").Item) > 0 Then
			aFormFieldComponent(S_QUERY_FOR_SOURCE_FORM_FIELD) = oRequest("QueryForSource").Item
		Else
			aFormFieldComponent(S_QUERY_FOR_SOURCE_FORM_FIELD) = ""
		End If
	End If
	aFormFieldComponent(S_QUERY_FOR_SOURCE_FORM_FIELD) = Left(aFormFieldComponent(S_QUERY_FOR_SOURCE_FORM_FIELD), 4000)

	If IsEmpty(aFormFieldComponent(S_DEFAULT_VALUE_FORM_FIELD)) Then
		If Len(oRequest("DefaultValue").Item) > 0 Then
			aFormFieldComponent(S_DEFAULT_VALUE_FORM_FIELD) = oRequest("DefaultValue").Item
		Else
			aFormFieldComponent(S_DEFAULT_VALUE_FORM_FIELD) = ""
		End If
	End If
	aFormFieldComponent(S_DEFAULT_VALUE_FORM_FIELD) = Left(aFormFieldComponent(S_DEFAULT_VALUE_FORM_FIELD), 255)

	If IsEmpty(aFormFieldComponent(S_JAVASCRIPT_CODE_FORM_FIELD)) Then
		If Len(oRequest("JavaScriptCode").Item) > 0 Then
			aFormFieldComponent(S_JAVASCRIPT_CODE_FORM_FIELD) = oRequest("JavaScriptCode").Item
		Else
			aFormFieldComponent(S_JAVASCRIPT_CODE_FORM_FIELD) = ""
		End If
	End If
	aFormFieldComponent(S_JAVASCRIPT_CODE_FORM_FIELD) = Left(aFormFieldComponent(S_JAVASCRIPT_CODE_FORM_FIELD), 255)

	If IsEmpty(aFormFieldComponent(S_DESCRIPTION_FORM_FIELD)) Then
		If Len(oRequest("FormFieldDescription").Item) > 0 Then
			aFormFieldComponent(S_DESCRIPTION_FORM_FIELD) = oRequest("FormFieldDescription").Item
		Else
			aFormFieldComponent(S_DESCRIPTION_FORM_FIELD) = ""
		End If
	End If
	aFormFieldComponent(S_DESCRIPTION_FORM_FIELD) = Left(aFormFieldComponent(S_DESCRIPTION_FORM_FIELD), 4000)

	aFormFieldComponent(B_CHECK_FOR_DUPLICATED_FORM_FIELD) = True
	aFormFieldComponent(B_IS_DUPLICATED_FORM_FIELD) = False

	aFormFieldComponent(B_COMPONENT_INITIALIZED_FORM_FIELD) = True
	InitializeFormFieldComponent = Err.number
	Err.Clear
End Function

Function AddFormField(oRequest, oADODBConnection, aFormFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new form into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aFormFieldComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddFormField"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFormFieldComponent(B_COMPONENT_INITIALIZED_FORM_FIELD)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFormFieldComponent(oRequest, aFormFieldComponent)
	End If

	If aFormFieldComponent(N_FIELD_ID_FORM_FIELD) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para el nuevo campo."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "FormFields", "FormFieldID", "(FormID=" & aFormFieldComponent(N_ID_FORM_FIELD) & ")", 1, aFormFieldComponent(N_FIELD_ID_FORM_FIELD), sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		If aFormFieldComponent(B_CHECK_FOR_DUPLICATED_FORM_FIELD) Then
			lErrorNumber = CheckExistencyOfFormField(oADODBConnection, aFormFieldComponent, sErrorDescription)
			If aFormFieldComponent(B_IS_DUPLICATED_FORM_FIELD) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un campo con el nombre '" & aFormFieldComponent(S_NAME_FORM_FIELD) & "' en la tabla '" & aFormFieldComponent(S_TABLE_NAME_FORM_FIELD) & "' para el DSN '" & aFormFieldComponent(S_DSN_FORM_FIELD) & "'."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FormFieldComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			End If
		End If

		If lErrorNumber = 0 Then
			If Not CheckFormFieldInformationConsistency(aFormFieldComponent, sErrorDescription) Then
				lErrorNumber = -1
			Else
				sErrorDescription = "No se pudo guardar la información del nuevo campo."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into FormFields (FormID, FormFieldID, DSNForField, ConnectionType, DatabaseName, TableName, FormFieldName, FormFieldText, IsOptional, FieldTypeID, FormFieldSize, LimitTypeID, MinimumValue, MaximumValue, DSNForSource, ConnectionTypeForSource, QueryForSource, DefaultValue, JavaScriptCode, FormFieldDescription) Values (" & aFormFieldComponent(N_ID_FORM_FIELD) & ", " & aFormFieldComponent(N_FIELD_ID_FORM_FIELD) & ", '" & Replace(aFormFieldComponent(S_DSN_FORM_FIELD), "'", "") & "', " & aFormFieldComponent(N_CONNECTION_TYPE_FORM_FIELD) & ", '" & Replace(aFormFieldComponent(S_DATABASE_NAME_FORM_FIELD), "'", "") & "', '" & Replace(aFormFieldComponent(S_TABLE_NAME_FORM_FIELD), "'", "") & "', '" & Replace(aFormFieldComponent(S_NAME_FORM_FIELD), "'", "") & "', '" & Replace(aFormFieldComponent(S_TEXT_FORM_FIELD), "'", "´") & "', " & aFormFieldComponent(N_IS_OPTIONAL_FORM_FIELD) & ", " & aFormFieldComponent(N_TYPE_ID_FORM_FIELD) & ", " & aFormFieldComponent(N_SIZE_FORM_FIELD) & ", " & aFormFieldComponent(N_LIMIT_TYPE_FORM_FIELD) & ", " & aFormFieldComponent(N_MINIMUM_FORM_FIELD) & ", " & aFormFieldComponent(N_MAXIMUM_FORM_FIELD) & ", '" & Replace(aFormFieldComponent(S_DSN_FOR_SOURCE_FORM_FIELD), "'", "") & "', " & aFormFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_FORM_FIELD) & ", '" & Replace(aFormFieldComponent(S_QUERY_FOR_SOURCE_FORM_FIELD), "'", "´") & "', '" & Replace(aFormFieldComponent(S_DEFAULT_VALUE_FORM_FIELD), "'", "´") & "', '" & Replace(aFormFieldComponent(S_JAVASCRIPT_CODE_FORM_FIELD), "'", "´") & "', '" & Replace(aFormFieldComponent(S_DESCRIPTION_FORM_FIELD), "'", "´") & "')", "FormFieldComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
	End If

	AddFormField = lErrorNumber
	Err.Clear
End Function

Function GetFormField(oRequest, oADODBConnection, aFormFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a form from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aFormFieldComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetFormField"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFormFieldComponent(B_COMPONENT_INITIALIZED_FORM_FIELD)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFormFieldComponent(oRequest, aFormFieldComponent)
	End If

	If (aFormFieldComponent(N_ID_FORM_FIELD) = -1) Or (aFormFieldComponent(N_FIELD_ID_FORM_FIELD) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del formulario y/o del campo para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FormFieldComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del campo."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From FormFields Where (FormID=" & aFormFieldComponent(N_ID_FORM_FIELD) & ") And (FormFieldID=" & aFormFieldComponent(N_FIELD_ID_FORM_FIELD) & ")", "FormFieldComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El campo especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FormFieldComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aFormFieldComponent(S_DSN_FORM_FIELD) = CStr(oRecordset.Fields("DSNForField").Value)
				aFormFieldComponent(N_CONNECTION_TYPE_FORM_FIELD) = CLng(oRecordset.Fields("ConnectionType").Value)
				aFormFieldComponent(S_DATABASE_NAME_FORM_FIELD) = CStr(oRecordset.Fields("DatabaseName").Value)
				aFormFieldComponent(S_TABLE_NAME_FORM_FIELD) = CStr(oRecordset.Fields("TableName").Value)
				aFormFieldComponent(S_NAME_FORM_FIELD) = CStr(oRecordset.Fields("FormFieldName").Value)
				aFormFieldComponent(S_TEXT_FORM_FIELD) = CStr(oRecordset.Fields("FormFieldText").Value)
				aFormFieldComponent(N_IS_OPTIONAL_FORM_FIELD) = CInt(oRecordset.Fields("IsOptional").Value)
				aFormFieldComponent(N_TYPE_ID_FORM_FIELD) = CLng(oRecordset.Fields("FieldTypeID").Value)
				aFormFieldComponent(N_SIZE_FORM_FIELD) = CInt(oRecordset.Fields("FormFieldSize").Value)
				aFormFieldComponent(N_LIMIT_TYPE_FORM_FIELD) = CInt(oRecordset.Fields("LimitTypeID").Value)
				aFormFieldComponent(N_MINIMUM_FORM_FIELD) = CDbl(oRecordset.Fields("MinimumValue").Value)
				aFormFieldComponent(N_MAXIMUM_FORM_FIELD) = CDbl(oRecordset.Fields("MaximumValue").Value)
				aFormFieldComponent(S_DSN_FOR_SOURCE_FORM_FIELD) = CStr(oRecordset.Fields("DSNForSource").Value)
				aFormFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_FORM_FIELD) = CLng(oRecordset.Fields("ConnectionTypeForSource").Value)
				aFormFieldComponent(S_QUERY_FOR_SOURCE_FORM_FIELD) = Replace(CStr(oRecordset.Fields("QueryForSource").Value), "´", "'")
				aFormFieldComponent(S_DEFAULT_VALUE_FORM_FIELD) = Replace(CStr(oRecordset.Fields("DefaultValue").Value), "´", "'")
				aFormFieldComponent(S_JAVASCRIPT_CODE_FORM_FIELD) = Replace(CStr(oRecordset.Fields("JavaScriptCode").Value), "´", "'")
				aFormFieldComponent(S_DESCRIPTION_FORM_FIELD) = CStr(oRecordset.Fields("FormFieldDescription").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetFormField = lErrorNumber
	Err.Clear
End Function

Function GetFormFields(oRequest, oADODBConnection, aFormFieldComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the forms from
'		  the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aFormFieldComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetFormFields"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFormFieldComponent(B_COMPONENT_INITIALIZED_FORM_FIELD)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFormFieldComponent(oRequest, aFormFieldComponent)
	End If

	If aFormFieldComponent(N_ID_FORM_FIELD) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del formulario para obtener sus campos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FormFieldComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información de los campos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select FormFields.*, FieldTypeName From FormFields, FieldTypes Where (FormFields.FieldTypeID=FieldTypes.FieldTypeID) And (FormID=" & aFormFieldComponent(N_ID_FORM_FIELD) & ") And (FormFieldID > -1) Order By DSNForField, TableName, FormFieldName", "FormFieldComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If

	GetFormFields = lErrorNumber
	Err.Clear
End Function

Function ModifyFormField(oRequest, oADODBConnection, aFormFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing form in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aFormFieldComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyFormField"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFormFieldComponent(B_COMPONENT_INITIALIZED_FORM_FIELD)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFormFieldComponent(oRequest, aFormFieldComponent)
	End If

	If (aFormFieldComponent(N_ID_FORM_FIELD) = -1) Or (aFormFieldComponent(N_FIELD_ID_FORM_FIELD) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del formulario y/o del campo para modificar su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FormFieldComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Not CheckFormFieldInformationConsistency(aFormFieldComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sErrorDescription = "No se pudo modificar la información del campo."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update FormFields Set DSNForField='" & Replace(aFormFieldComponent(S_DSN_FORM_FIELD), "'", "") & "', ConnectionType=" & aFormFieldComponent(N_CONNECTION_TYPE_FORM_FIELD) & ", DatabaseName='" & Replace(aFormFieldComponent(S_DATABASE_NAME_FORM_FIELD), "'", "") & "', TableName='" & Replace(aFormFieldComponent(S_TABLE_NAME_FORM_FIELD), "'", "") & "', FormFieldName='" & Replace(aFormFieldComponent(S_NAME_FORM_FIELD), "'", "") & "', FormFieldText='" & Replace(aFormFieldComponent(S_TEXT_FORM_FIELD), "'", "´") & "', IsOptional=" & aFormFieldComponent(N_IS_OPTIONAL_FORM_FIELD) & ", FieldTypeID=" & aFormFieldComponent(N_TYPE_ID_FORM_FIELD) & ", FormFieldSize=" & aFormFieldComponent(N_SIZE_FORM_FIELD) & ", LimitTypeID=" & aFormFieldComponent(N_LIMIT_TYPE_FORM_FIELD) & ", MinimumValue=" & aFormFieldComponent(N_MINIMUM_FORM_FIELD) & ", MaximumValue=" & aFormFieldComponent(N_MAXIMUM_FORM_FIELD) & ", DSNForSource='" & Replace(aFormFieldComponent(S_DSN_FOR_SOURCE_FORM_FIELD), "'", "") & "', ConnectionTypeForSource=" & aFormFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_FORM_FIELD) & ", QueryForSource='" & Replace(aFormFieldComponent(S_QUERY_FOR_SOURCE_FORM_FIELD), "'", "´") & "', DefaultValue='" & Replace(aFormFieldComponent(S_DEFAULT_VALUE_FORM_FIELD), "'", "´") & "', JavaScriptCode='" & Replace(aFormFieldComponent(S_JAVASCRIPT_CODE_FORM_FIELD), "'", "´") & "', FormFieldDescription='" & Replace(aFormFieldComponent(S_DESCRIPTION_FORM_FIELD), "'", "´") & "' Where (FormID=" & aFormFieldComponent(N_ID_FORM_FIELD) & ") And (FormFieldID=" & aFormFieldComponent(N_FIELD_ID_FORM_FIELD) & ")", "FormFieldComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	ModifyFormField = lErrorNumber
	Err.Clear
End Function

Function RemoveFormField(oRequest, oADODBConnection, aFormFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a form from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aFormFieldComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveFormField"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFormFieldComponent(B_COMPONENT_INITIALIZED_FORM_FIELD)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFormFieldComponent(oRequest, aFormFieldComponent)
	End If

	If (aFormFieldComponent(N_ID_FORM_FIELD) = -1) Or (aFormFieldComponent(N_FIELD_ID_FORM_FIELD) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del formulario y/o del campo a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FormFieldComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del campo."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From FormFields Where (FormID=" & aFormFieldComponent(N_ID_FORM_FIELD) & ") And (FormFieldID=" & aFormFieldComponent(N_FIELD_ID_FORM_FIELD) & ")", "FormFieldComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la información del formulario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From FormFieldsAnswers Where (FormID=" & aFormFieldComponent(N_ID_FORM_FIELD) & ") And (FormFieldID=" & aFormFieldComponent(N_FIELD_ID_FORM_FIELD) & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	RemoveFormField = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfFormField(oADODBConnection, aFormFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific form exists in the database
'Inputs:  oADODBConnection, aFormFieldComponent
'Outputs: aFormFieldComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfFormField"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFormFieldComponent(B_COMPONENT_INITIALIZED_FORM_FIELD)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFormFieldComponent(oRequest, aFormFieldComponent)
	End If

	If (aFormFieldComponent(N_ID_FORM_FIELD) = -1) Or (Len(aFormFieldComponent(S_DSN_FORM_FIELD)) = 0) Or (Len(aFormFieldComponent(S_TABLE_NAME_FORM_FIELD)) = 0) Or (Len(aFormFieldComponent(S_NAME_FORM_FIELD)) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el nombre del DSN, de la tabla o del campo para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FormFieldComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del campo en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From FormFields Where (FormID=" & aFormFieldComponent(N_ID_FORM_FIELD) & ") And (DSNForField='" & Replace(aFormFieldComponent(S_DSN_FORM_FIELD), "'", "") & "') And (TableName='" & Replace(aFormFieldComponent(S_TABLE_NAME_FORM_FIELD), "'", "") & "') And (FormFieldName='" & Replace(aFormFieldComponent(S_NAME_FORM_FIELD), "'", "") & "')", "FormFieldComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			aFormFieldComponent(B_IS_DUPLICATED_FORM_FIELD) = (Not oRecordset.EOF)
		End If
	End If
	oRecordset.Close

	Set oRecordset = Nothing
	CheckExistencyOfFormField = lErrorNumber
	Err.Clear
End Function

Function CheckFormFieldInformationConsistency(aFormFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aFormFieldComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckFormFieldInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aFormFieldComponent(N_ID_FORM_FIELD)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del formulario no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aFormFieldComponent(N_FIELD_ID_FORM_FIELD)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del campo no es un valor numérico."
		bIsCorrect = False
	End If
	If Len(aFormFieldComponent(S_DSN_FORM_FIELD)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El DSN de destino de los datos está vacío."
		bIsCorrect = False
	End If
	If Not IsNumeric(aFormFieldComponent(N_CONNECTION_TYPE_FORM_FIELD)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El tipo de conexión con la base de datos no es un valor numérico."
		bIsCorrect = False
	End If
	If Len(aFormFieldComponent(S_DATABASE_NAME_FORM_FIELD)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre de la base de datos destino está vacío."
		bIsCorrect = False
	End If
	If Len(aFormFieldComponent(S_TABLE_NAME_FORM_FIELD)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre de la tabla está vacío."
		bIsCorrect = False
	End If
	If Len(aFormFieldComponent(S_NAME_FORM_FIELD)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del campo está vacío."
		bIsCorrect = False
	End If
	If Len(aFormFieldComponent(S_TEXT_FORM_FIELD)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El descriptivo del campo está vacío."
		bIsCorrect = False
	End If
	If Not IsNumeric(aFormFieldComponent(N_IS_OPTIONAL_FORM_FIELD)) Then aFormFieldComponent(N_IS_OPTIONAL_FORM_FIELD) = 0
	If Not IsNumeric(aFormFieldComponent(N_TYPE_ID_FORM_FIELD)) Then aFormFieldComponent(N_TYPE_ID_FORM_FIELD) = 5
	If Not IsNumeric(aFormFieldComponent(N_SIZE_FORM_FIELD)) Then aFormFieldComponent(N_SIZE_FORM_FIELD) = 1
	If Not IsNumeric(aFormFieldComponent(N_LIMIT_TYPE_FORM_FIELD)) Then aFormFieldComponent(N_LIMIT_TYPE_FORM_FIELD) = 0
	If Not IsNumeric(aFormFieldComponent(N_MINIMUM_FORM_FIELD)) Then aFormFieldComponent(N_MINIMUM_FORM_FIELD) = 0
	If Not IsNumeric(aFormFieldComponent(N_MAXIMUM_FORM_FIELD)) Then aFormFieldComponent(N_MAXIMUM_FORM_FIELD) = 0
	If Len(aFormFieldComponent(S_DSN_FOR_SOURCE_FORM_FIELD)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El DSN de origen de los datos está vacío."
		bIsCorrect = False
	End If
	If Not IsNumeric(aFormFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_FORM_FIELD)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El tipo de conexión con el origen de los datos no es un valor numérico."
		bIsCorrect = False
	End If

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del campo contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FormFieldComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckFormFieldInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayFormFieldForm(oRequest, oADODBConnection, sAction, aFormFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a form from the
'		  database using a HTML FormField
'Inputs:  oRequest, oADODBConnection, sAction, aFormFieldComponent
'Outputs: aFormFieldComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFormFieldForm"
	Dim sNames
	Dim asFields
	Dim iIndex
	Dim lErrorNumber

	If (aFormFieldComponent(N_ID_FORM_FIELD) <> -1) And (aFormFieldComponent(N_FIELD_ID_FORM_FIELD) <> -1) Then
		lErrorNumber = GetFormField(oRequest, oADODBConnection, aFormFieldComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckFormFields(oForm) {" & vbNewLine
				If Len(oRequest("Delete").Item) = 0 Then
					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "if (oForm.DSNForField.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir el DSN de destino de los datos.');" & vbNewLine
							Response.Write "oForm.FormFieldName.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (oForm.DatabaseName.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir el nombre de la base de datos destino.');" & vbNewLine
							Response.Write "oForm.FormFieldName.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (oForm.TableName.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir el nombre de la tabla.');" & vbNewLine
							Response.Write "oForm.FormFieldName.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
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
			Response.Write "} // End of CheckFormFields" & vbNewLine

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
		Response.Write "<FORM NAME=""FormFieldFrm"" ID=""FormFieldFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckFormFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""FormFields"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormID"" ID=""FormIDHdn"" VALUE=""" & aFormFieldComponent(N_ID_FORM_FIELD) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormFieldID"" ID=""FormFieldIDHdn"" VALUE=""" & aFormFieldComponent(N_FIELD_ID_FORM_FIELD) & """ />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">DSN:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""DSNForField"" ID=""DSNForFieldTxt"" SIZE=""35"" MAXLENGTH=""255"" VALUE=""" & aFormFieldComponent(S_DSN_FORM_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Tipo de conexión:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><SELECT NAME=""ConnectionType"" ID=""ConnectionTypeCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE=""" & SQL_SERVER & """"
							If aFormFieldComponent(N_CONNECTION_TYPE_FORM_FIELD) = SQL_SERVER Then Response.Write " SELECTED=""1"""
						Response.Write ">SQL Server</OPTION>"
						Response.Write "<OPTION VALUE=""" & ACCESS & """"
							If aFormFieldComponent(N_CONNECTION_TYPE_FORM_FIELD) = ACCESS Then Response.Write " SELECTED=""1"""
						Response.Write ">MS Access File</OPTION>"
						Response.Write "<OPTION VALUE=""" & ACCESS_DSN & """"
							If aFormFieldComponent(N_CONNECTION_TYPE_FORM_FIELD) = ACCESS_DSN Then Response.Write " SELECTED=""1"""
						Response.Write ">MS Access (DSN)</OPTION>"
						Response.Write "<OPTION VALUE=""" & ORACLE & """"
							If aFormFieldComponent(N_CONNECTION_TYPE_FORM_FIELD) = ORACLE Then Response.Write " SELECTED=""1"""
						Response.Write ">Oracle</OPTION>"
						Response.Write "<OPTION VALUE=""" & MYSQL & """"
							If aFormFieldComponent(N_CONNECTION_TYPE_FORM_FIELD) = MYSQL Then Response.Write " SELECTED=""1"""
						Response.Write ">MySQL</OPTION>"
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Base de datos:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""DatabaseName"" ID=""DatabaseNameTxt"" SIZE=""35"" MAXLENGTH=""255"" VALUE=""" & aFormFieldComponent(S_DATABASE_NAME_FORM_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Tabla:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""TableName"" ID=""TableNameTxt"" SIZE=""35"" MAXLENGTH=""255"" VALUE=""" & aFormFieldComponent(S_TABLE_NAME_FORM_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Nombre:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""FormFieldName"" ID=""FormFieldNameTxt"" SIZE=""35"" MAXLENGTH=""255"" VALUE=""" & aFormFieldComponent(S_NAME_FORM_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Descriptivo:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""FormFieldText"" ID=""FormFieldTextTxt"" SIZE=""35"" MAXLENGTH=""255"" VALUE=""" & aFormFieldComponent(S_TEXT_FORM_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">¿Es opcional?&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><SELECT NAME=""IsOptional"" ID=""IsOptionalCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE=""0"""
							If aFormFieldComponent(N_IS_OPTIONAL_FORM_FIELD) = 0 Then Response.Write " SELECTED=""1"""
						Response.Write ">No</OPTION>"
						Response.Write "<OPTION VALUE=""1"""
							If aFormFieldComponent(N_IS_OPTIONAL_FORM_FIELD) = 1 Then Response.Write " SELECTED=""1"""
						Response.Write ">Sí</OPTION>"
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Tipo:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><SELECT NAME=""FieldTypeID"" ID=""FieldTypeIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowFieldsForFieldType(this.value)"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "FieldTypes", "FieldTypeID", "FieldTypeName", "", "FieldTypeName", aFormFieldComponent(N_TYPE_ID_FORM_FIELD), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""SizeDiv"" ID=""SizeDiv"" STYLE=""display: none"">"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Tamaño:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""FormFieldSize"" ID=""FormFieldSizeTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & aFormFieldComponent(N_SIZE_FORM_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""MinimumSizeDiv"" ID=""MinimumSizeDiv"" STYLE=""display: none"">"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Mínimo número de caracteres:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""MinimumSize"" ID=""MinimumSizeTxt"" SIZE=""3"" MAXLENGTH=""3"" VALUE="""
						If aFormFieldComponent(N_MINIMUM_FORM_FIELD) > 0 Then Response.Write aFormFieldComponent(N_MINIMUM_FORM_FIELD)
					Response.Write """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""LimitTypeDiv"" ID=""LimitTypeDiv"" STYLE=""display: none"">"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Tipo de límite:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><SELECT NAME=""LimitTypeID"" ID=""LimitTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "LimitTypes", "LimitTypeID", "LimitTypeName", "", "LimitTypeID", aFormFieldComponent(N_LIMIT_TYPE_FORM_FIELD), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""MinimumDiv"" ID=""MinimumDiv"" STYLE=""display: none"">"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Mínimo:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""MinimumValue"" ID=""MinimumValueTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & aFormFieldComponent(N_MINIMUM_FORM_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""MaximumDiv"" ID=""MaximumDiv"" STYLE=""display: none"">"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Máximo:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""MaximumValue"" ID=""MaximumValueTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & aFormFieldComponent(N_MAXIMUM_FORM_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""MinimumYearDiv"" ID=""MinimumYearDiv"" STYLE=""display: none"">"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Mostrar desde el año:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><SELECT NAME=""MinimumYearValue"" ID=""MinimumYearValueCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE=""0"">Año en curso</OPTION>"
						For iIndex = Year(Date()) To N_FORM_START_YEAR Step -1
							Response.Write "<OPTION VALUE=""" & iIndex & """"
								If iIndex = aFormFieldComponent(N_MINIMUM_FORM_FIELD) Then Response.Write " SELECTED=""1"""
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
								If iIndex = Abs(aFormFieldComponent(N_MAXIMUM_FORM_FIELD)) Then Response.Write " SELECTED=""1"""
							Response.Write ">Año en curso +" & iIndex & "</OPTION>"
						Next
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">DSN de origen de los datos:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""DSNForSource"" ID=""DSNForSourceTxt"" SIZE=""35"" MAXLENGTH=""255"" VALUE=""" & aFormFieldComponent(S_DSN_FOR_SOURCE_FORM_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><NOBR><FONT FACE=""Arial"" SIZE=""2"">Tipo de conexión con el origen de los datos:&nbsp;</FONT></NOBR></TD>"
					Response.Write "<TD><SELECT NAME=""ConnectionTypeForSource"" ID=""ConnectionTypeForSourceCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE=""" & SQL_SERVER & """"
							If aFormFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_FORM_FIELD) = SQL_SERVER Then Response.Write " SELECTED=""1"""
						Response.Write ">SQL Server</OPTION>"
						Response.Write "<OPTION VALUE=""" & ACCESS & """"
							If aFormFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_FORM_FIELD) = ACCESS Then Response.Write " SELECTED=""1"""
						Response.Write ">MS Access File</OPTION>"
						Response.Write "<OPTION VALUE=""" & ACCESS_DSN & """"
							If aFormFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_FORM_FIELD) = ACCESS_DSN Then Response.Write " SELECTED=""1"""
						Response.Write ">MS Access (DSN)</OPTION>"
						Response.Write "<OPTION VALUE=""" & ORACLE & """"
							If aFormFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_FORM_FIELD) = ORACLE Then Response.Write " SELECTED=""1"""
						Response.Write ">Oracle</OPTION>"
						Response.Write "<OPTION VALUE=""" & MYSQL & """"
							If aFormFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_FORM_FIELD) = MYSQL Then Response.Write " SELECTED=""1"""
						Response.Write ">MySQL</OPTION>"
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""NoCatalogFieldDiv"" ID=""NoCatalogFieldDiv"""
					If (aFormFieldComponent(N_TYPE_ID_FORM_FIELD) = N_CATALOG) Or (aFormFieldComponent(N_TYPE_ID_FORM_FIELD) = N_LIST) Then Response.Write " STYLE=""display: none"""
				Response.Write "><TD COLSPAN=""2"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Query de origen de los datos:</FONT><BR />"
					Response.Write "<SELECT onChange=""document.FormFieldFrm.QueryForSource.value=this.value"">"
						Response.Write "<OPTION VALUE=""""></OPTION>"
						Response.Write "<OPTION VALUE=""Select UserName From Users Where UserID=<USER_ID />"">Nombre del usuario</OPTION>"
						Response.Write "<OPTION VALUE=""Select UserLastName From Users Where UserID=<USER_ID />"">Apellidos del usuario</OPTION>"
						Response.Write "<OPTION VALUE=""Select UserName, UserLastName From Users Where UserID=<USER_ID />"">Nombre y apellidos del usuario</OPTION>"
						Response.Write "<OPTION VALUE=""Select UserEmail From Users Where UserID=<USER_ID />"">Correo electrónico del usuario</OPTION>"
					Response.Write "</SELECT><BR />"
					Response.Write "<TEXTAREA NAME=""QueryForSource"" ID=""QueryForSourceTxtArea"" ROWS=""5"" COLS=""40"" MAXLENGTH=""4000"" CLASS=""TextFields"">" & aFormFieldComponent(S_QUERY_FOR_SOURCE_FORM_FIELD) & "</TEXTAREA><BR />"
				Response.Write "</TD></TR>"
				Response.Write "<TR NAME=""CatalogFieldDiv"" ID=""CatalogFieldDiv"""
					If (aFormFieldComponent(N_TYPE_ID_FORM_FIELD) <> N_CATALOG) And (aFormFieldComponent(N_TYPE_ID_FORM_FIELD) <> N_LIST) Then Response.Write " STYLE=""display: none"""
				Response.Write "><TD COLSPAN=""2"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Query de origen de los datos:</FONT><BR />"
					Response.Write "<SELECT onChange=""if (this.value != '') {var sFieldTemp = this.value.split('" & LIST_SEPARATOR & "'); document.FormFieldFrm.Catalog_0.value=sFieldTemp[0]; document.FormFieldFrm.Catalog_1.value=sFieldTemp[1]; document.FormFieldFrm.Catalog_2.value=sFieldTemp[2]; document.FormFieldFrm.Catalog_3.value=sFieldTemp[3]; document.FormFieldFrm.Catalog_4.value=sFieldTemp[4]; document.FormFieldFrm.Catalog_5.value=sFieldTemp[5];}"">"
						Response.Write "<OPTION VALUE=""""></OPTION>"
						Response.Write "<OPTION VALUE=""States;;;StateID;;;;;;StateName;;;;;;StateName"">Catálogo de estados</OPTION>"
						Response.Write "<OPTION VALUE=""Genders;;;GenderID;;;;;;GenderName;;;;;;GenderName"">Catálogo de géneros</OPTION>"
					Response.Write "</SELECT><BR /><FONT FACE=""Arial"" SIZE=""2"">"
					If InStr(1, aFormFieldComponent(S_QUERY_FOR_SOURCE_FORM_FIELD), LIST_SEPARATOR, vbBinaryCompare) = 0 Then asFields = Split(BuildList("", LIST_SEPARATOR, 6), LIST_SEPARATOR) 
					asFields = Split(aFormFieldComponent(S_QUERY_FOR_SOURCE_FORM_FIELD), LIST_SEPARATOR, -1, vbBinaryCompare)
					If UBound(asFields) < 5 Then asFields = Split(JoinLists(aFormFieldComponent(S_QUERY_FOR_SOURCE_FORM_FIELD), BuildList("", LIST_SEPARATOR, 6 - UBound(asFields)), LIST_SEPARATOR), LIST_SEPARATOR, -1, vbBinaryCompare)
					Response.Write "Tabla:<IMG SRC=""Images/Transparent.gif"" WIDTH=""50"" HEIGHT=""1"" /><INPUT TYPE=""TEXT"" NAME=""Catalog_0"" ID=""Catalog_0Txt"" SIZE=""20"" MAXLENGTH=""100"" VALUE=""" & asFields(0) & """ CLASS=""TextFields"" /><BR />"
					Response.Write "Campo llave:<IMG SRC=""Images/Transparent.gif"" WIDTH=""11"" HEIGHT=""1"" /><INPUT TYPE=""TEXT"" NAME=""Catalog_1"" ID=""Catalog_1Txt"" SIZE=""20"" MAXLENGTH=""100"" VALUE=""" & asFields(1) & """ CLASS=""TextFields"" /><BR />"
					Response.Write "<SPAN NAME=""HierarchyFieldDiv"" ID=""HierarchyFieldDiv"">Campo padre: <INPUT TYPE=""TEXT"" NAME=""Catalog_2"" ID=""Catalog_2Txt"" SIZE=""20"" MAXLENGTH=""100"" VALUE=""" & asFields(2) & """ CLASS=""TextFields"" /><BR /></SPAN>"
					Response.Write "Campo:<IMG SRC=""Images/Transparent.gif"" WIDTH=""40"" HEIGHT=""1"" /><INPUT TYPE=""TEXT"" NAME=""Catalog_3"" ID=""Catalog_3Txt"" SIZE=""20"" MAXLENGTH=""100"" VALUE=""" & asFields(3) & """ CLASS=""TextFields"" /><BR />"
					Response.Write "Condición:<IMG SRC=""Images/Transparent.gif"" WIDTH=""24"" HEIGHT=""1"" /><INPUT TYPE=""TEXT"" NAME=""Catalog_4"" ID=""Catalog_4Txt"" SIZE=""20"" MAXLENGTH=""100"" VALUE=""" & asFields(4) & """ CLASS=""TextFields"" /><BR />"
					Response.Write "Ordenar por:<IMG SRC=""Images/Transparent.gif"" WIDTH=""13"" HEIGHT=""1"" /><INPUT TYPE=""TEXT"" NAME=""Catalog_5"" ID=""Catalog_5Txt"" SIZE=""20"" MAXLENGTH=""100"" VALUE=""" & asFields(5) & """ CLASS=""TextFields"" /><BR />"
				Response.Write "</FONT></TD></TR>"
				Response.Write "<TR><TD COLSPAN=""2"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Valor predeterminado:&nbsp;</FONT>"
					Response.Write "<INPUT TYPE=""TEXT"" NAME=""DefaultValue"" ID=""DefaultValueTxt"" SIZE=""46"" MAXLENGTH=""255"" VALUE=""" & aFormFieldComponent(S_DEFAULT_VALUE_FORM_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR><TD COLSPAN=""2"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">JavaScript:&nbsp;</FONT>"
					Response.Write "<INPUT TYPE=""TEXT"" NAME=""JavaScriptCode"" ID=""JavaScriptCodeTxt"" SIZE=""59"" MAXLENGTH=""255"" VALUE=""" & aFormFieldComponent(S_JAVASCRIPT_CODE_FORM_FIELD) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR><TD COLSPAN=""2"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Descripción:</FONT><BR />"
					Response.Write "<TEXTAREA NAME=""FormFieldDescription"" ID=""FormFieldDescriptionTxtArea"" ROWS=""5"" COLS=""40"" MAXLENGTH=""4000"" CLASS=""TextFields"">" & aFormFieldComponent(S_DESCRIPTION_FORM_FIELD) & "</TEXTAREA><BR />"
				Response.Write "</TD></TR>"
			Response.Write "</TABLE>"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "ShowFieldsForFieldType(document.FormFieldFrm.FieldTypeID.value);" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine

			Response.Write "<BR />"
			If Len(oRequest("Change").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveFormFieldWngDiv']); FormFieldFrm.Remove.focus()"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=FormFields&FormID=" & oRequest("FormID").Item & "'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveFormFieldWngDiv", "¿Está seguro que desea borrar el campo de la &nbsp;base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayFormFieldForm = lErrorNumber
	Err.Clear
End Function

Function DisplayFormFieldAsHiddenFields(oRequest, oADODBConnection, aFormFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a form using
'		  hidden form fields
'Inputs:  oRequest, oADODBConnection, aFormFieldComponent
'Outputs: aFormFieldComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFormFieldAsHiddenFields"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormID"" ID=""FormIDHdn"" VALUE=""" & aFormFieldComponent(N_ID_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormFieldID"" ID=""FormFieldIDHdn"" VALUE=""" & aFormFieldComponent(N_FIELD_ID_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DSNForField"" ID=""DSNForFieldHdn"" VALUE=""" & aFormFieldComponent(S_DSN_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConnectionType"" ID=""ConnectionTypeHdn"" VALUE=""" & aFormFieldComponent(N_CONNECTION_TYPE_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DatabaseName"" ID=""DatabaseNameHdn"" VALUE=""" & aFormFieldComponent(S_DATABASE_NAME_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TableName"" ID=""TableNameHdn"" VALUE=""" & aFormFieldComponent(S_TABLE_NAME_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormFieldName"" ID=""FormFieldNameHdn"" VALUE=""" & aFormFieldComponent(S_NAME_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormFieldText"" ID=""FormFieldTextHdn"" VALUE=""" & aFormFieldComponent(S_TEXT_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IsOptional"" ID=""IsOptionalHdn"" VALUE=""" & aFormFieldComponent(N_IS_OPTIONAL_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FieldTypeID"" ID=""FieldTypeIDHdn"" VALUE=""" & aFormFieldComponent(N_TYPE_ID_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormFieldSize"" ID=""FormFieldSizeHdn"" VALUE=""" & aFormFieldComponent(N_SIZE_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""LimitTypeID"" ID=""LimitTypeIDHdn"" VALUE=""" & aFormFieldComponent(N_LIMIT_TYPE_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""MinimumValue"" ID=""MinimumValueHdn"" VALUE=""" & aFormFieldComponent(N_MINIMUM_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""MaximumValue"" ID=""MaximumValueHdn"" VALUE=""" & aFormFieldComponent(N_MAXIMUM_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DSNForSource"" ID=""DSNForSourceHdn"" VALUE=""" & aFormFieldComponent(S_DSN_FOR_SOURCE_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConnectionTypeForSource"" ID=""ConnectionTypeForSourceHdn"" VALUE=""" & aFormFieldComponent(N_CONNECTION_TYPE_FOR_SOURCE_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""QueryForSource"" ID=""QueryForSourceHdn"" VALUE=""" & aFormFieldComponent(S_QUERY_FOR_SOURCE_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DefaultValue"" ID=""DefaultValueHdn"" VALUE=""" & aFormFieldComponent(S_DEFAULT_VALUE_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JavaScriptCode"" ID=""JavaScriptCodeHdn"" VALUE=""" & aFormFieldComponent(S_JAVASCRIPT_CODE_FORM_FIELD) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormFieldDescription"" ID=""FormFieldDescriptionHdn"" VALUE=""" & aFormFieldComponent(S_DESCRIPTION_FORM_FIELD) & """ />"

	DisplayFormFieldAsHiddenFields = Err.number
	Err.Clear
End Function

Function DisplayFormFieldsTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, aFormFieldComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the forms from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aFormFieldComponent
'Outputs: aFormFieldComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFormFieldsTable"
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

	lErrorNumber = GetFormFields(oRequest, oADODBConnection, aFormFieldComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""400"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
					asColumnsTitles = Split("&nbsp;,DSN,Tabla,Campo,Tipo,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,100,100,100,80", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("&nbsp;,DSN,Tabla,Campo,Tipo", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,130,130,120", ",", -1, vbBinaryCompare)
				End If
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				asCellAlignments = Split(",,,,,CENTER", ",", -1, vbBinaryCompare)
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
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""FormFieldIdentificator"" ID=""FormFieldIdentificatorRd"" VALUE=""" & CStr(oRecordset.Fields("FormID").Value) & "," & CStr(oRecordset.Fields("FormFieldID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""FormFieldIdentificator"" ID=""FormFieldIdentificatorChk"" VALUE=""" & CStr(oRecordset.Fields("FormID").Value) & "," & CStr(oRecordset.Fields("FormFieldID").Value) & """ />"
						Case Else
							sRowContents = sRowContents & "&nbsp;"
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("DSNForField").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("TableName").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE=""" & CleanStringForHTML(CStr(oRecordset.Fields("FormFieldText").Value)) & """>" & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("FormFieldName").Value)) & sBoldEnd & "</FONT>"
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("FieldTypeName").Value))
						If CInt(oRecordset.Fields("IsOptional").Value) = 0 Then sRowContents = sRowContents & " *"
					sRowContents = sRowContents & sBoldEnd
					If bUseLinks And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
						sRowContents = sRowContents & TABLE_SEPARATOR
							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=FormFields&FormID=" & CStr(oRecordset.Fields("FormID").Value) & "&FormFieldID=" & CStr(oRecordset.Fields("FormFieldID").Value) & "&Change=1"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>"
							End If

							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS Then
								sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=FormFields&FormID=" & CStr(oRecordset.Fields("FormID").Value) & "&FormFieldID=" & CStr(oRecordset.Fields("FormFieldID").Value) & "&Delete=1"">"
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
	DisplayFormFieldsTable = lErrorNumber
	Err.Clear
End Function
%>