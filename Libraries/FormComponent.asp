<%
Const N_ID_FORM = 0
Const S_NAME_FORM = 1
Const S_RECEIVERS_FORM = 2
Const N_MODULE_FORM = 3
Const S_FILE_FORM = 4
Const S_PRINT_FILE_FORM = 5
Const S_NEXT_URL_FORM = 6
Const S_DESCRIPTION_FORM = 7
Const N_ACTIVE_FORM = 8
Const N_ANSWER_ID_FORM = 9
Const S_ANSWER_ID_FORM = 10
Const N_FIELD_ID_FORM = 11
Const N_USER_ID_FORM = 12
Const N_INSPECTION_ID_FORM = 13
Const N_DATE_FORM = 14
Const S_COMMENTS_FORM = 15
Const N_STATUS_ID_FORM = 16
Const N_FIELDS_FORM = 17
Const B_ANSWERED_FORM = 18
Const B_CHECK_FOR_DUPLICATED_FORM = 19
Const B_IS_DUPLICATED_FORM = 20
Const B_COMPONENT_INITIALIZED_FORM = 21

Const N_FORM_COMPONENT_SIZE = 21

Dim aFormComponent()
Redim aFormComponent(N_FORM_COMPONENT_SIZE)

Function InitializeFormComponent(oRequest, aFormComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Form Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aFormComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeFormComponent"
	Redim Preserve aFormComponent(N_FORM_COMPONENT_SIZE)

	If IsEmpty(aFormComponent(N_ID_FORM)) Then
		If Len(oRequest("FormID").Item) > 0 Then
			aFormComponent(N_ID_FORM) = CLng(oRequest("FormID").Item)
		Else
			aFormComponent(N_ID_FORM) = -1
		End If
	End If

	If IsEmpty(aFormComponent(S_NAME_FORM)) Then
		If Len(oRequest("FormName").Item) > 0 Then
			aFormComponent(S_NAME_FORM) = oRequest("FormName").Item
		Else
			aFormComponent(S_NAME_FORM) = ""
		End If
	End If
	aFormComponent(S_NAME_FORM) = Left(aFormComponent(S_NAME_FORM), 255)

	If IsEmpty(aFormComponent(S_RECEIVERS_FORM)) Then
		If Len(oRequest("ReceiversID").Item) > 0 Then
			aFormComponent(S_RECEIVERS_FORM) = Replace(oRequest("ReceiversID").Item, " ", "")
		Else
			aFormComponent(S_RECEIVERS_FORM) = "-1"
		End If
	End If
	aFormComponent(S_RECEIVERS_FORM) = Left(aFormComponent(S_RECEIVERS_FORM), 255)

	If IsEmpty(aFormComponent(N_MODULE_FORM)) Then
		If Len(oRequest("ModuleID").Item) > 0 Then
			aFormComponent(N_MODULE_FORM) = CLng(oRequest("ModuleID").Item)
		Else
			aFormComponent(N_MODULE_FORM) = lCurrentModule
		End If
	End If

	If IsEmpty(aFormComponent(S_FILE_FORM)) Then
		If Len(oRequest("FormFile").Item) > 0 Then
			aFormComponent(S_FILE_FORM) = oRequest("FormFile").Item
		Else
			aFormComponent(S_FILE_FORM) = ""
		End If
	End If
	aFormComponent(S_FILE_FORM) = Left(aFormComponent(S_FILE_FORM), 255)

	If IsEmpty(aFormComponent(S_PRINT_FILE_FORM)) Then
		If Len(oRequest("PrintFile").Item) > 0 Then
			aFormComponent(S_PRINT_FILE_FORM) = oRequest("PrintFile").Item
		Else
			aFormComponent(S_PRINT_FILE_FORM) = ""
		End If
	End If
	aFormComponent(S_PRINT_FILE_FORM) = Left(aFormComponent(S_PRINT_FILE_FORM), 255)

	If IsEmpty(aFormComponent(S_NEXT_URL_FORM)) Then
		If Len(oRequest("NextURL").Item) > 0 Then
			aFormComponent(S_NEXT_URL_FORM) = oRequest("NextURL").Item
		Else
			aFormComponent(S_NEXT_URL_FORM) = ""
		End If
	End If
	aFormComponent(S_NEXT_URL_FORM) = Left(aFormComponent(S_NEXT_URL_FORM), 255)

	If IsEmpty(aFormComponent(S_DESCRIPTION_FORM)) Then
		If Len(oRequest("FormDescription").Item) > 0 Then
			aFormComponent(S_DESCRIPTION_FORM) = oRequest("FormDescription").Item
		Else
			aFormComponent(S_DESCRIPTION_FORM) = ""
		End If
	End If
	aFormComponent(S_DESCRIPTION_FORM) = Left(aFormComponent(S_DESCRIPTION_FORM), 3000)

	If IsEmpty(aFormComponent(N_ACTIVE_FORM)) Then
		If Len(oRequest("Active").Item) > 0 Then
			aFormComponent(N_ACTIVE_FORM) = CInt(oRequest("Active").Item)
		Else
			aFormComponent(N_ACTIVE_FORM) = 1
		End If
	End If

	If IsEmpty(aFormComponent(N_ANSWER_ID_FORM)) Then
		If Len(oRequest("MaintenanceID").Item) > 0 Then
			If InStr(1, oRequest("MaintenanceID").Item, ",", vbBinaryCompare) > 0 Then
				aFormComponent(N_ANSWER_ID_FORM) = -1
				aFormComponent(S_ANSWER_ID_FORM) = oRequest("MaintenanceID").Item
			Else
				aFormComponent(N_ANSWER_ID_FORM) = CLng(oRequest("MaintenanceID").Item)
				aFormComponent(S_ANSWER_ID_FORM) = oRequest("MaintenanceID").Item & "," & CLng(oRequest("FormID").Item)
			End If
		ElseIf Len(oRequest("AnswerID").Item) > 0 Then
			If InStr(1, oRequest("AnswerID").Item, ",", vbBinaryCompare) > 0 Then
				aFormComponent(N_ANSWER_ID_FORM) = -1
				aFormComponent(S_ANSWER_ID_FORM) = oRequest("AnswerID").Item
			Else
				aFormComponent(N_ANSWER_ID_FORM) = CLng(oRequest("AnswerID").Item)
				aFormComponent(S_ANSWER_ID_FORM) = oRequest("AnswerID").Item & "," & CLng(oRequest("FormID").Item)
			End If
		Else
			aFormComponent(N_ANSWER_ID_FORM) = -1
		End If
	End If

	If IsEmpty(aFormComponent(N_FIELD_ID_FORM)) Then
		If Len(oRequest("FormFieldID").Item) > 0 Then
			aFormComponent(N_FIELD_ID_FORM) = CInt(oRequest("FormFieldID").Item)
		Else
			aFormComponent(N_FIELD_ID_FORM) = -1
		End If
	End If

	If IsEmpty(aFormComponent(N_USER_ID_FORM)) Then
		If Len(oRequest("FormUserID").Item) > 0 Then
			aFormComponent(N_USER_ID_FORM) = CLng(oRequest("FormUserID").Item)
		Else
			aFormComponent(N_USER_ID_FORM) = aLoginComponent(N_USER_ID_LOGIN)
		End If
	End If

	If IsEmpty(aFormComponent(N_VEHICLE_ID_FORM)) Then
		If Len(oRequest("VehicleID").Item) > 0 Then
			aFormComponent(N_VEHICLE_ID_FORM) = CLng(oRequest("VehicleID").Item)
		Else
			aFormComponent(N_VEHICLE_ID_FORM) = -1
		End If
	End If

	If IsEmpty(aFormComponent(N_DATE_FORM)) Then
		If Len(oRequest("FormDate").Item) > 0 Then
			aFormComponent(N_DATE_FORM) = CLng(oRequest("FormDate").Item)
		Else
			aFormComponent(N_DATE_FORM) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
		End If
	End If

	If IsEmpty(aFormComponent(S_COMMENTS_FORM)) Then
		If Len(oRequest("AnswerComments").Item) > 0 Then
			aFormComponent(S_COMMENTS_FORM) = oRequest("AnswerComments").Item
		Else
			aFormComponent(S_COMMENTS_FORM) = ""
		End If
	End If
	aFormComponent(S_COMMENTS_FORM) = Left(aFormComponent(S_COMMENTS_FORM), 4000)

	If IsEmpty(aFormComponent(N_STATUS_ID_FORM)) Then
		If Len(oRequest("StatusFormID").Item) > 0 Then
			aFormComponent(N_STATUS_ID_FORM) = CLng(oRequest("StatusFormID").Item)
		Else
			aFormComponent(N_STATUS_ID_FORM) = 1
		End If
	End If

	aFormComponent(B_ANSWERED_FORM) = False
	aFormComponent(B_CHECK_FOR_DUPLICATED_FORM) = True
	aFormComponent(B_IS_DUPLICATED_FORM) = False

	aFormComponent(B_COMPONENT_INITIALIZED_FORM) = True
	InitializeFormComponent = Err.number
	Err.Clear
End Function

Function AddForm(oRequest, oADODBConnection, aFormComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new form into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aFormComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddForm"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFormComponent(B_COMPONENT_INITIALIZED_FORM)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFormComponent(oRequest, aFormComponent)
	End If

	If aFormComponent(N_ID_FORM) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para el nuevo formulario."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "Forms", "FormID", "", 1, aFormComponent(N_ID_FORM), sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		If aFormComponent(B_CHECK_FOR_DUPLICATED_FORM) Then
			lErrorNumber = CheckExistencyOfForm(oADODBConnection, aFormComponent, sErrorDescription)
			If aFormComponent(B_IS_DUPLICATED_FORM) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un formulario con el nombre " & aFormComponent(S_NAME_FORM) & " registrado en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FormComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			End If
		End If

		If lErrorNumber = 0 Then
			If Not CheckFormInformationConsistency(aFormComponent, sErrorDescription) Then
				lErrorNumber = -1
			Else
				sErrorDescription = "No se pudo guardar la información del nuevo formulario."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Forms (FormID, FormName, ReceiversID, ModuleID, FormFile, PrintFile, NextURL, FormDescription, Active) Values (" & aFormComponent(N_ID_FORM) & ", '" & Replace(aFormComponent(S_NAME_FORM), "'", "") & "', '" & aFormComponent(S_RECEIVERS_FORM) & "', " & aFormComponent(N_MODULE_FORM) & ", '" & Replace(aFormComponent(S_FILE_FORM), "'", "") & "', '" & Replace(aFormComponent(S_PRINT_FILE_FORM), "'", "") & "', '" & Replace(aFormComponent(S_NEXT_URL_FORM), "'", "") & "', '" & Replace(aFormComponent(S_DESCRIPTION_FORM), "'", "") & "', " & aFormComponent(N_ACTIVE_FORM) & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
	End If

	AddForm = lErrorNumber
	Err.Clear
End Function

Function GetForm(oRequest, oADODBConnection, aFormComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a form from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aFormComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetForm"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFormComponent(B_COMPONENT_INITIALIZED_FORM)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFormComponent(oRequest, aFormComponent)
	End If

	If aFormComponent(N_ID_FORM) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del formulario para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FormComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del formulario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Forms Where FormID=" & aFormComponent(N_ID_FORM), "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El formulario especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FormComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aFormComponent(S_NAME_FORM) = CStr(oRecordset.Fields("FormName").Value)
				aFormComponent(S_RECEIVERS_FORM) = CStr(oRecordset.Fields("ReceiversID").Value)
				aFormComponent(N_MODULE_FORM) = CInt(oRecordset.Fields("ModuleID").Value)
				aFormComponent(S_FILE_FORM) = CStr(oRecordset.Fields("FormFile").Value)
				aFormComponent(S_PRINT_FILE_FORM) = CStr(oRecordset.Fields("PrintFile").Value)
				aFormComponent(S_NEXT_URL_FORM) = CStr(oRecordset.Fields("NextURL").Value)
				aFormComponent(S_DESCRIPTION_FORM) = CStr(oRecordset.Fields("FormDescription").Value)
				aFormComponent(N_ACTIVE_FORM) = CInt(oRecordset.Fields("Active").Value)
			End If
			oRecordset.Close

			aFormComponent(N_FIELDS_FORM) = 0
			sErrorDescription = "No se pudo obtener la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) From FormFields Where FormID=" & aFormComponent(N_ID_FORM), "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then aFormComponent(N_FIELDS_FORM) = CInt(oRecordset.Fields(0).Value)
				oRecordset.Close
			End If

			sErrorDescription = "No se pudo revisar si el formulario ya fue contestado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct FormsAnswers.FormID From FormsAnswers, FormFieldsAnswers Where (FormsAnswers.AnswerID=FormFieldsAnswers.AnswerID) And (FormsAnswers.FormID=FormFieldsAnswers.FormID) And (FormsAnswers.AnswerID=" & aFormComponent(N_ANSWER_ID_FORM) & ") And (FormsAnswers.FormID=" & aFormComponent(N_ID_FORM) & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then aFormComponent(B_ANSWERED_FORM) = (Not oRecordset.EOF)
		End If
	End If

	Set oRecordset = Nothing
	GetForm = lErrorNumber
	Err.Clear
End Function

Function GetForms(oRequest, oADODBConnection, aFormComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the forms from
'		  the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aFormComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetForms"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFormComponent(B_COMPONENT_INITIALIZED_FORM)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFormComponent(oRequest, aFormComponent)
	End If

	sErrorDescription = "No se pudo obtener la información de los formularios."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Forms.*, ModuleName From Forms, Modules Where (Forms.ModuleID=Modules.ModuleID) And (FormID > -1) Order By FormName", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetForms = lErrorNumber
	Err.Clear
End Function

Function ModifyForm(oRequest, oADODBConnection, aFormComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing form in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aFormComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyForm"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFormComponent(B_COMPONENT_INITIALIZED_FORM)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFormComponent(oRequest, aFormComponent)
	End If

	If aFormComponent(N_ID_FORM) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del formulario a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FormComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Not CheckFormInformationConsistency(aFormComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sErrorDescription = "No se pudo modificar la información del formulario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Forms Set FormName='" & Replace(aFormComponent(S_NAME_FORM), "'", "") & "', ReceiversID='" & aFormComponent(S_RECEIVERS_FORM) & "', ModuleID=" & aFormComponent(N_MODULE_FORM) & ", FormFile='" & Replace(aFormComponent(S_FILE_FORM), "'", "") & "', PrintFile='" & Replace(aFormComponent(S_PRINT_FILE_FORM), "'", "") & "', NextURL='" & Replace(aFormComponent(S_NEXT_URL_FORM), "'", "") & "', FormDescription='" & Replace(aFormComponent(S_DESCRIPTION_FORM), "'", "") & "', Active=" & aFormComponent(N_ACTIVE_FORM) & " Where (FormID=" & aFormComponent(N_ID_FORM) & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	ModifyForm = lErrorNumber
	Err.Clear
End Function

Function SetActiveForForm(oRequest, oADODBConnection, aFormComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given form
'Inputs:  oRequest, oADODBConnection
'Outputs: aFormComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForForm"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFormComponent(B_COMPONENT_INITIALIZED_FORM)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFormComponent(oRequest, aFormComponent)
	End If

	If aFormComponent(N_ID_FORM) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FormComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo modificar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Forms Set Active=" & CInt(oRequest("SetActive").Item) & " Where (FormID=" & aFormComponent(N_ID_FORM) & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	SetActiveForForm = lErrorNumber
	Err.Clear
End Function

Function RemoveForm(oRequest, oADODBConnection, aFormComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a form from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aFormComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveForm"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFormComponent(B_COMPONENT_INITIALIZED_FORM)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFormComponent(oRequest, aFormComponent)
	End If

	If aFormComponent(N_ID_FORM) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el formulario a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FormComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del formulario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Forms Where (FormID=" & aFormComponent(N_ID_FORM) & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la información del formulario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From FormFields Where (FormID=" & aFormComponent(N_ID_FORM) & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la información del formulario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From FormsAnswers Where (FormID=" & aFormComponent(N_ID_FORM) & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la información del formulario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From FormFieldsAnswers Where (FormID=" & aFormComponent(N_ID_FORM) & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	RemoveForm = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfForm(oADODBConnection, aFormComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific form exists in the database
'Inputs:  oADODBConnection, aFormComponent
'Outputs: aFormComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfForm"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFormComponent(B_COMPONENT_INITIALIZED_FORM)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFormComponent(oRequest, aFormComponent)
	End If

	If Len(aFormComponent(S_NAME_FORM)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el nombre del formulario para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FormComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del formulario en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Forms Where (FormName='" & Replace(aFormComponent(S_NAME_FORM), "'", "") & "')", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			aFormComponent(B_IS_DUPLICATED_FORM) = (Not oRecordset.EOF)
		End If
	End If
	oRecordset.Close

	Set oRecordset = Nothing
	CheckExistencyOfForm = lErrorNumber
	Err.Clear
End Function

Function CheckFormInformationConsistency(aFormComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aFormComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckFormInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aFormComponent(N_ID_FORM)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del formulario no es un valor numérico."
		bIsCorrect = False
	End If
	If Len(aFormComponent(S_NAME_FORM)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del formulario está vacío."
		bIsCorrect = False
	End If
	If Len(aFormComponent(S_RECEIVERS_FORM)) = 0 Then aFormComponent(S_RECEIVERS_FORM) = "-1"
	If Not IsNumeric(aFormComponent(N_MODULE_FORM)) Then aFormComponent(N_MODULE_FORM) = -1

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del formulario contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "FormComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckFormInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayFormForm(oRequest, oADODBConnection, sAction, aFormComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a form from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aFormComponent
'Outputs: aFormComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFormForm"
	Dim sFolderContents
	Dim asFolderContents
	Dim iIndex
	Dim oRecordset
	Dim lErrorNumber

	If aFormComponent(N_ID_FORM) <> -1 Then
		lErrorNumber = GetForm(oRequest, oADODBConnection, aFormComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckFormFields(oForm) {" & vbNewLine
				If Len(oRequest("Delete").Item) = 0 Then
					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "if (oForm.FormName.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir el nombre del formulario.');" & vbNewLine
							Response.Write "oForm.FormName.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
				End If
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckFormFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""FormFrm"" ID=""FormFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckFormFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Forms"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormID"" ID=""FormIDHdn"" VALUE=""" & aFormComponent(N_ID_FORM) & """ />"
			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Nombre: </FONT>"
			Response.Write "<INPUT TYPE=""TEXT"" NAME=""FormName"" ID=""FormNameTxt"" SIZE=""35"" MAXLENGTH=""100"" VALUE=""" & aFormComponent(S_NAME_FORM) & """ CLASS=""TextFields"" /><BR /><BR />"

			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Usuarios que recibirán el formulario:<BR /></FONT>"
			Response.Write "<SELECT NAME=""ReceiversID"" ID=""ReceiversIDLst"" SIZE=""10"" MULTIPLE=""1"" CLASS=""TextFields"">"
				sErrorDescription = "No se pudieron obtener los usuarios que cuentan con permisos de administrador."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select UserID, UserName, UserLastName From Users Where (UserID >= 10) And (UserPermissions <> 0) Order By UserLastName, UserName", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						Response.Write "<OPTION VALUE=""" & CStr(oRecordset.Fields("UserID").Value) & """"
							If InStr(1, ("," & aFormComponent(S_RECEIVERS_FORM) & ","), ("," & CStr(oRecordset.Fields("UserID").Value) & ","), vbBinaryCompare) > 0 Then Response.Write " SELECTED=""1"""
						Response.Write ">" & CleanStringForHTML(CStr(oRecordset.Fields("UserLastName").Value) & ", " & CStr(oRecordset.Fields("UserName").Value)) & "</OPTION>"
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
				End If
			Response.Write "</SELECT><BR /><BR />"

			If False Then
				Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Módulo en donde aparecerá el formulario:<BR /></FONT>"
				Response.Write "<SELECT NAME=""ModuleID"" ID=""ModuleIDCmb"" SIZE=""1"" CLASS=""TextFields"">"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Modules", "ModuleID", "ModuleName", "", "ModuleName", aFormComponent(N_MODULE_FORM), "Ninguno;;;0", sErrorDescription)
				Response.Write "</SELECT><BR /><BR />"
			Else
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ModuleID"" ID=""ModuleIDHdn"" VALUE=""" & aFormComponent(N_MODULE_FORM) & """ />"
			End If

			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Plantilla: </FONT>"
			Response.Write "<SELECT NAME=""FormFile"" ID=""FormFileCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""if ((this.value == '') || (this.value.search(/\.asp/gi) != -1)) {HideDisplay(document.all['TemplateDiv']);} else {document.all['TemplateIFrame'].src = 'Templates\/' + this.value; ShowDisplay(document.all['TemplateDiv']);}"">"
				sFolderContents = ""
				lErrorNumber = GetFolderContents(Server.MapPath(TEMPLATES_PHYSICAL_PATH), False, sFolderContents, sErrorDescription)
				asFolderContents = Split(sFolderContents, LIST_SEPARATOR)
				Response.Write "<OPTION VALUE="""">Seleccione una plantilla</OPTION>"
				For iIndex=0 To UBound(asFolderContents)
					Response.Write "<OPTION VALUE=""" & asFolderContents(iIndex) & """>" & asFolderContents(iIndex) & "</OPTION>"
				Next
			Response.Write "</SELECT><BR /><BR />"

			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Impresión: </FONT>"
			Response.Write "<SELECT NAME=""PrintFile"" ID=""PrintFileCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""if ((this.value == '') || (this.value.search(/\.asp/gi) != -1)) {HideDisplay(document.all['TemplateDiv']);} else {document.all['TemplateIFrame'].src = 'Templates\/' + this.value; ShowDisplay(document.all['TemplateDiv']);}"">"
				sFolderContents = ""
				lErrorNumber = GetFolderContents(Server.MapPath(TEMPLATES_PHYSICAL_PATH), False, sFolderContents, sErrorDescription)
				asFolderContents = Split(sFolderContents, LIST_SEPARATOR)
				Response.Write "<OPTION VALUE="""">Seleccione una plantilla</OPTION>"
				For iIndex=0 To UBound(asFolderContents)
					Response.Write "<OPTION VALUE=""" & asFolderContents(iIndex) & """>" & asFolderContents(iIndex) & "</OPTION>"
				Next
			Response.Write "</SELECT><BR /><BR />"

			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">URL que se abrirá al guardar el formulario: </FONT>"
			Response.Write "<INPUT TYPE=""TEXT"" NAME=""NextURL"" ID=""NextURLTxt"" SIZE=""35"" MAXLENGTH=""255"" VALUE=""" & aFormComponent(S_NEXT_URL_FORM) & """ CLASS=""TextFields"" /><BR /><BR />"

			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Descripción:</FONT><BR />"
			Response.Write "<TEXTAREA NAME=""FormDescription"" ID=""FormDescriptionTxtArea"" ROWS=""5"" COLS=""40"" MAXLENGTH=""3000"" CLASS=""TextFields"">" & aFormComponent(S_DESCRIPTION_FORM) & "</TEXTAREA><BR /><BR />"

			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Activo:&nbsp;"
			Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" ID=""ActiveRd"" VALUE=""1"""
				If aFormComponent(N_ACTIVE_FORM) = 1 Then Response.Write " CHECKED=""1"""
			Response.Write " /> Sí&nbsp;&nbsp;&nbsp;"
			Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" ID=""ActiveRd"" VALUE=""0"""
				If aFormComponent(N_ACTIVE_FORM) = 0 Then Response.Write " CHECKED=""1"""
			Response.Write " /> No</FONT><BR /><BR />"

			Response.Write "<DIV NAME=""TemplateDiv"" ID=""TemplateDiv"" STYLE=""display: none""><IFRAME SRC="""" NAME=""TemplateIFrame"" FRAMEBORDER=""1"" WIDTH=""98%"" HEIGHT=""250""></IFRAME></DIV><BR />"

			Response.Write "<BR />"
			If Len(oRequest("Change").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveFormWngDiv']); FormFrm.Remove.focus()"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=Forms'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveFormWngDiv", "¿Está seguro que desea borrar el formulario de la &nbsp;base de datos?")
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "SelectItemByValue('" & aFormComponent(S_FILE_FORM) & "', false, document.FormFrm.FormFile);" & vbNewLine
				Response.Write "SelectItemByValue('" & aFormComponent(S_PRINT_FILE_FORM) & "', false, document.FormFrm.PrintFile);" & vbNewLine
				If InStr(1, aFormComponent(S_FILE_FORM), ".htm", vbBinaryCompare) > 0 Then
					Response.Write "document.all['TemplateIFrame'].src = 'Templates\/' + document.FormFrm.FormFile.value;" & vbNewLine
					Response.Write "ShowDisplay(document.all['TemplateDiv']);" & vbNewLine
				End If
			Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "</FORM>"
	End If

	DisplayFormForm = lErrorNumber
	Err.Clear
End Function

Function DisplayDuplicateFormForm(oRequest, oADODBConnection, sAction, aFormComponent, sErrorDescription)
'************************************************************
'Purpose: To display the form for duplicate the form's data
'Inputs:  oRequest, oADODBConnection, sAction, aFormComponent
'Outputs: aFormComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayDuplicateFormForm"
	Dim oRecordset
	Dim lErrorNumber

	lErrorNumber = GetForms(oRequest, oADODBConnection, aFormComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckDuplicateFormFields(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "if (oForm.SourceFormID.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de seleccionar el formulario a duplicar.');" & vbNewLine
						Response.Write "oForm.SourceFormID.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (oForm.FormName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de introducir el nombre del formulario.');" & vbNewLine
						Response.Write "oForm.FormName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckFormFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""DuplicateFormFrm"" ID=""DuplicateFormFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckDuplicateFormFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Forms"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormID"" ID=""FormIDHdn"" VALUE=""-1"" />"
			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Formulario a duplicar: </FONT>"
			Response.Write "<SELECT NAME=""SourceFormID"" ID=""SourceFormIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""this.form.FormName.value = GetSelectedText(this)"">"
				Response.Write "<OPTION VALUE="""">Seleccione un formulario</OPTION>"
				Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Forms", "FormID", "FormName", "", "FormName", oRequest("SourceFormID").Item, "", sErrorDescription)
			Response.Write "</SELECT><BR /><BR />"
			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Nombre del nuevo formulario: </FONT>"
			Response.Write "<INPUT TYPE=""TEXT"" NAME=""FormName"" ID=""FormNameTxt"" SIZE=""35"" MAXLENGTH=""100"" VALUE=""" & aFormComponent(S_NAME_FORM) & """ CLASS=""TextFields"" /><BR /><BR />"

			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.document.body.focus(); HidePopupItem('DuplicateDiv', document.DuplicateDiv)"" />"
		Response.Write "</FORM>"
	End If

	DisplayDuplicateFormForm = lErrorNumber
	Err.Clear
End Function

Function DisplayFormAsHiddenFields(oRequest, oADODBConnection, aFormComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a form using
'		  hidden form fields
'Inputs:  oRequest, oADODBConnection, aFormComponent
'Outputs: aFormComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFormAsHiddenFields"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormID"" ID=""FormIDHdn"" VALUE=""" & aFormComponent(N_ID_FORM) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormName"" ID=""FormNameHdn"" VALUE=""" & aFormComponent(S_NAME_FORM) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReceiversID"" ID=""ReceiversIDHdn"" VALUE=""" & aFormComponent(S_RECEIVERS_FORM) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ModuleID"" ID=""ModuleIDHdn"" VALUE=""" & aFormComponent(N_MODULE_FORM) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormFile"" ID=""FormFileHdn"" VALUE=""" & aFormComponent(S_FILE_FORM) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PrintFile"" ID=""PrintFileHdn"" VALUE=""" & aFormComponent(S_PRINT_FILE_FORM) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""NextURL"" ID=""NextURLHdn"" VALUE=""" & aFormComponent(S_NEXT_URL_FORM) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormDescription"" ID=""FormDescriptionHdn"" VALUE=""" & aFormComponent(S_DESCRIPTION_FORM) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Active"" ID=""ActiveHdn"" VALUE=""" & aFormComponent(N_ACTIVE_FORM) & """ />"

	DisplayFormAsHiddenFields = Err.number
	Err.Clear
End Function

Function DisplayFormsTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, aFormComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the forms from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aFormComponent
'Outputs: aFormComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFormsTable"
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

	lErrorNumber = GetForms(oRequest, oADODBConnection, aFormComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""450"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
					asColumnsTitles = Split("&nbsp;,Formulario,Módulo,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,250,100,80", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("&nbsp;,Formulario,Módulo", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,290,140", ",", -1, vbBinaryCompare)
				End If
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				asCellAlignments = Split(",,,CENTER", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("FormID").Value), oRequest("FormID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""FormID"" ID=""FormIDRd"" VALUE=""" & CStr(oRecordset.Fields("FormID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""FormID"" ID=""FormIDChk"" VALUE=""" & CStr(oRecordset.Fields("FormID").Value) & """ />"
						Case Else
							sRowContents = sRowContents & "&nbsp;"
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""Catalogs.asp?FormID=" & CStr(oRecordset.Fields("FormID").Value) & "&Action=FormFields"">" & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("FormName").Value)) & sBoldEnd & "</A>"
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ModuleName").Value)) & sBoldEnd
					If bUseLinks And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
						sRowContents = sRowContents & TABLE_SEPARATOR
							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Forms&FormID=" & CStr(oRecordset.Fields("FormID").Value) & "&Change=1"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
							End If

							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS Then
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Forms&FormID=" & CStr(oRecordset.Fields("FormID").Value) & "&Delete=1"">"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
							End If

							If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
								If CInt(oRecordset.Fields("Active").Value) = 0 Then
									sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Forms&FormID=" & CStr(oRecordset.Fields("FormID").Value) & "&SetActive=1""><IMG SRC=""Images/BtnActive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Activar"" BORDER=""0"" /></A>"
								Else
									sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Forms&FormID=" & CStr(oRecordset.Fields("FormID").Value) & "&SetActive=0""><IMG SRC=""Images/BtnDeactive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Desactivar"" BORDER=""0"" /></A>"
								End If
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
			sErrorDescription = "No existen formularios registrados en la base de datos."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayFormsTable = lErrorNumber
	Err.Clear
End Function

Function SaveUserAnswer(oADODBConnection, lFormID, lFormFieldID, lAnswerID, sAnswer, sErrorDescription)
'************************************************************
'Purpose: To save the given answer for the specified field
'Inputs:  oADODBConnection, lFormID, lFormFieldID, lAnswerID, sAnswer
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SaveUserAnswer"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AnswerID From FormFieldsAnswers Where (AnswerID=" & lAnswerID & ") And (FormID=" & lFormID & ") And (FormFieldID=" & lFormFieldID & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudo guardar la información del nuevo registro."
		If oRecordset.EOF Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into FormFieldsAnswers (AnswerID, FormID, FormFieldID, Answer) Values (" & lAnswerID & ", " & lFormID & ", " & lFormFieldID & ", '" & Replace(sAnswer, "'", "´") & "')", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update FormFieldsAnswers Set Answer='" & Replace(sAnswer, "'", "´") & "' Where (AnswerID=" & lAnswerID & ") And (FormID=" & lFormID & ") And (FormFieldID=" & lFormFieldID & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	SaveUserAnswer = lErrorNumber
	Err.Clear
End Function

Function SaveUserForm(oRequest, oADODBConnection, aFormComponent, sErrorDescription)
'************************************************************
'Purpose: To save the user's answers for the specified form
'Inputs:  oRequest, oADODBConnection, aFormComponent
'Outputs: aFormComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SaveUserForm"
	Dim oItem
	Dim asFormFieldIDs
	Dim sDateFields
	Dim sTempRequest
	Dim sDateRequest
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aFormComponent(B_COMPONENT_INITIALIZED_FORM)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeFormComponent(oRequest, aFormComponent)
	End If

	sDateFields = ","
	sTempRequest = ""
	sDateRequest = " "

	If aFormComponent(N_ANSWER_ID_FORM) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para el nuevo formulario."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "FormsAnswers", "AnswerID", "(FormID=" & aFormComponent(N_ID_FORM) & ")", 1, aFormComponent(N_ANSWER_ID_FORM), sErrorDescription)
		If lErrorNumber = 0 Then
			If Len(oRequest("StatusFormID__" & aFormComponent(N_ANSWER_ID_FORM) & "__" & aFormComponent(N_ID_FORM)).Item) > 0 Then aFormComponent(N_STATUS_ID_FORM) = CLng(oRequest("StatusFormID__" & aFormComponent(N_ANSWER_ID_FORM) & "__" & aFormComponent(N_ID_FORM)).Item)
			sErrorDescription = "No se pudo guardar la respuesta del usuario al campo del formulario especificado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into FormsAnswers (AnswerID, FormID, UserID, VehicleID, FormDate, AnswerComments, StatusFormID) Values (" & aFormComponent(N_ANSWER_ID_FORM) & ", " & aFormComponent(N_ID_FORM) & ", " & aFormComponent(N_USER_ID_FORM) & ", " & aFormComponent(N_VEHICLE_ID_FORM) & ", " & aFormComponent(N_DATE_FORM) & ", '" & Replace(aFormComponent(S_COMMENTS_FORM), "'", "´") & "', " & aFormComponent(N_STATUS_ID_FORM) & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	If lErrorNumber = 0 Then
		For Each oItem In oRequest
			sTempRequest = ""
			If (InStr(1, oItem, "FF__", vbBinaryCompare) > 0) And (InStr(1, oItem, sDateRequest, vbBinaryCompare) = 0) Then
				asFormFieldIDs = Split(oItem, "__", 3, vbBinaryCompare)
				If Not IsNumeric(asFormFieldIDs(2)) Then
					If (InStr(1, asFormFieldIDs(2), "Hour", vbBinaryCompare) > 0) Or (InStr(1, asFormFieldIDs(2), "Minute", vbBinaryCompare) > 0) Then
						asFormFieldIDs(2) = Replace(Replace(asFormFieldIDs(2), "Hour", ""), "Minute", "")
						sDateFields = sDateFields & Replace(Replace(oItem, "Hour", ""), "Minute", "") & ","
						sTempRequest = oRequest(Join(asFormFieldIDs, "__") & "Hour") & oRequest(Join(asFormFieldIDs, "__") & "Minute")
					Else
						asFormFieldIDs(2) = Replace(Replace(Replace(asFormFieldIDs(2), "Day", ""), "Month", ""), "Year", "")
						sDateFields = sDateFields & Replace(Replace(Replace(oItem, "Day", ""), "Month", ""), "Year", "") & ","
						sTempRequest = oRequest(Join(asFormFieldIDs, "__") & "Year") & oRequest(Join(asFormFieldIDs, "__") & "Month") & oRequest(Join(asFormFieldIDs, "__") & "Day")
					End If
					sDateRequest = Join(asFormFieldIDs, "__")
				Else
					sTempRequest = oRequest(oItem).Item
					sDateRequest = " "
				End If
				If IsNumeric(asFormFieldIDs(1)) And IsNumeric(asFormFieldIDs(2)) Then
					sErrorDescription = "No se revisar la existencia de la respuesta del usuario al campo del formulario especificado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AnswerID From FormFieldsAnswers Where (AnswerID=" & aFormComponent(N_ANSWER_ID_FORM) & ") And (FormID=" & asFormFieldIDs(1) & ") And (FormFieldID=" & asFormFieldIDs(2) & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudo guardar la respuesta del usuario al campo del formulario especificado."
						If oRecordset.EOF Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into FormFieldsAnswers (AnswerID, FormID, FormFieldID, Answer) Values (" & aFormComponent(N_ANSWER_ID_FORM) & ", " & asFormFieldIDs(1) & ", " & asFormFieldIDs(2) & ", '" & Replace(sTempRequest, "'", "´") & "')", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update FormFieldsAnswers Set Answer='" & Replace(sTempRequest, "'", "´") & "' Where (AnswerID=" & aFormComponent(N_ANSWER_ID_FORM) & ") And (FormID=" & asFormFieldIDs(1) & ") And (FormFieldID=" & asFormFieldIDs(2) & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
						oRecordset.Close
					End If
				End If
			ElseIf (InStr(1, oItem, "StatusFormID__", vbBinaryCompare) > 0) Then
				asFormFieldIDs = Split(oItem, "__", 3, vbBinaryCompare)
				sErrorDescription = "No se pudo actualizar el status del formulario especificado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update FormsAnswers Set StatusFormID=" & oRequest(oItem).Item & " Where (AnswerID=" & asFormFieldIDs(1) & ") And (FormID=" & asFormFieldIDs(2) & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
		Next
	End If

	Set oRecordset = Nothing
	SaveUserForm = lErrorNumber
	Err.Clear
End Function

Function DisplayFormForTask(oADODBConnection, sFormName, bUseTemplate, bForEmail, aFormComponent, aTaskComponent, sErrorDescription)
'************************************************************
'Purpose: To display the details for the given task and user
'Inputs:  oRequest, oADODBConnection, sFormName, bUseTemplate, bForEmail, aFormComponent, aTaskComponent
'Outputs: aFormComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFormForTask"
	Dim iCurrentForm
	Dim lFormID
	Dim sFormFile
	Dim sFormContents
	Dim sField
	Dim sQueryForSource
	Dim sUserName
	Dim sTab
	Dim asFields
	Dim asValues
	Dim sURL
	Dim iIndex
	Dim sMinType
	Dim sMaxType
	Dim bTemplate
	Dim bASP
	Dim asIDs
	Dim oFieldADODBConnection
	Dim oRecordset
	Dim oCatalogRecordset
	Dim sAnswer
	Dim asFiles
	Dim aEmailComponent
	Dim sReceiversID
	Dim lErrorNumber

	aFormComponent(N_FIELDS_FORM) = 0
	sErrorDescription = "No se pudo obtener la información del registro."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(FormFieldID) From Forms, FormFields Where (Forms.FormID=FormFields.FormID) And (Forms.Active=1)", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then aFormComponent(N_FIELDS_FORM) = CInt(oRecordset.Fields(0).Value)
		oRecordset.Close
	End If
	If aFormComponent(N_FIELDS_FORM) > 50 And (Not bForEmail) And (StrComp(GetASPFileName(""), "Export.asp", vbBinaryCompare) <> 0) Then
		Response.Write "<DONT_EXPORT><SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "var bContinue = false;" & vbNewLine
			Response.Write "var iOnlyOne = 0;" & vbNewLine
			Response.Write "var aFormFields = new Array("
				Call GenerateJavaScriptArrayFromQuery(oADODBConnection, "Forms, FormFields, Tasks", "FormFields.FormID, FormFieldID", "FieldTypeID", "(Forms.FormID=FormFields.FormID) And (Forms.FormID=Tasks.FormID) And (Forms.Active=1) And (Tasks.ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (Tasks.TaskID=" & aTaskComponent(N_ID_TASK) & ")", "FormFields.FormID, FormFieldID", sErrorDescription)
			Response.Write "['-1', '-1']);" & vbNewLine

			Response.Write "function SaveAnswers(oForm, iFormFieldID, bOnlyOne, sURL) {" & vbNewLine
				Response.Write "var oField = null;" & vbNewLine
				Response.Write "var oField2 = null;" & vbNewLine
				Response.Write "var oField3 = null;" & vbNewLine
				Response.Write "var bSave = true;" & vbNewLine
				Response.Write "if (bOnlyOne)" & vbNewLine
					Response.Write "iOnlyOne = 1;" & vbNewLine
				Response.Write "if (iFormFieldID == 0)" & vbNewLine
					Response.Write "bSave = CheckFormForModule(oForm);" & vbNewLine

				Response.Write "if (bSave) {" & vbNewLine
					Response.Write "ShowPopupItem('Wait2Div', document.Wait2Div);" & vbNewLine
					Response.Write "if (iFormFieldID < (aFormFields.length - 1)) {" & vbNewLine
						Response.Write "switch (aFormFields[iFormFieldID][2]) {" & vbNewLine
							Response.Write "case '1':" & vbNewLine
								Response.Write "oField = eval('oForm.FF__' + aFormFields[iFormFieldID][0] + '__' + aFormFields[iFormFieldID][1] + 'Year');" & vbNewLine
								Response.Write "oField2 = eval('oForm.FF__' + aFormFields[iFormFieldID][0] + '__' + aFormFields[iFormFieldID][1] + 'Month');" & vbNewLine
								Response.Write "oField3 = eval('oForm.FF__' + aFormFields[iFormFieldID][0] + '__' + aFormFields[iFormFieldID][1] + 'Day');" & vbNewLine
								Response.Write "break;" & vbNewLine
							Response.Write "case '3':" & vbNewLine
								Response.Write "oField = eval('oForm.FF__' + aFormFields[iFormFieldID][0] + '__' + aFormFields[iFormFieldID][1] + 'Hour');" & vbNewLine
								Response.Write "oField2 = eval('oForm.FF__' + aFormFields[iFormFieldID][0] + '__' + aFormFields[iFormFieldID][1] + 'Minute');" & vbNewLine
								Response.Write "break;" & vbNewLine
							Response.Write "case '10':" & vbNewLine
								Response.Write "oField = null;" & vbNewLine
								Response.Write "oField2 = null;" & vbNewLine
								Response.Write "oField3 = null;" & vbNewLine
								Response.Write "break;" & vbNewLine
							Response.Write "default:" & vbNewLine
								Response.Write "oField = eval('oForm.FF__' + aFormFields[iFormFieldID][0] + '__' + aFormFields[iFormFieldID][1]);" & vbNewLine
								Response.Write "oField2 = null;" & vbNewLine
								Response.Write "oField3 = null;" & vbNewLine
								Response.Write "break;" & vbNewLine
						Response.Write "}" & vbNewLine

						Response.Write "if (oField) {" & vbNewLine
							Response.Write "switch (aFormFields[iFormFieldID][2]) {" & vbNewLine
								Response.Write "case '0':" & vbNewLine
									Response.Write "if (oField.checked)" & vbNewLine
										Response.Write "window.document.SaveFormAnswerIFrame.location.href = 'SaveFormAnswer.asp?FormName=' + oForm.name + '&ItemID=' + iFormFieldID + '&FormID=' + aFormFields[iFormFieldID][0] + '&FormFieldID=' + aFormFields[iFormFieldID][1] + '&AnswerID=" & aFormComponent(N_ANSWER_ID_FORM) & "&Answer=1&bOnlyOne=' + iOnlyOne + '&' + sURL;" & vbNewLine
									Response.Write "else" & vbNewLine
										Response.Write "window.document.SaveFormAnswerIFrame.location.href = 'SaveFormAnswer.asp?FormName=' + oForm.name + '&ItemID=' + iFormFieldID + '&FormID=' + aFormFields[iFormFieldID][0] + '&FormFieldID=' + aFormFields[iFormFieldID][1] + '&AnswerID=" & aFormComponent(N_ANSWER_ID_FORM) & "&Answer=0&bOnlyOne=' + iOnlyOne + '&' + sURL;" & vbNewLine
									Response.Write "break;" & vbNewLine
								Response.Write "case '1':" & vbNewLine
									Response.Write "window.document.SaveFormAnswerIFrame.location.href = 'SaveFormAnswer.asp?FormName=' + oForm.name + '&ItemID=' + iFormFieldID + '&FormID=' + aFormFields[iFormFieldID][0] + '&FormFieldID=' + aFormFields[iFormFieldID][1] + '&AnswerID=" & aFormComponent(N_ANSWER_ID_FORM) & "&Answer=' + oField.value + oField2.value + oField3.value + '&bOnlyOne=' + iOnlyOne + '&' + sURL;" & vbNewLine
									Response.Write "break;" & vbNewLine
								Response.Write "case '3':" & vbNewLine
									Response.Write "window.document.SaveFormAnswerIFrame.location.href = 'SaveFormAnswer.asp?FormName=' + oForm.name + '&ItemID=' + iFormFieldID + '&FormID=' + aFormFields[iFormFieldID][0] + '&FormFieldID=' + aFormFields[iFormFieldID][1] + '&AnswerID=" & aFormComponent(N_ANSWER_ID_FORM) & "&Answer=' + oField.value + oField2.value + '&bOnlyOne=' + iOnlyOne + '&' + sURL;" & vbNewLine
									Response.Write "break;" & vbNewLine
								Response.Write "case '5':" & vbNewLine
									Response.Write "window.document.SaveFormAnswerIFrame.location.href = 'SaveFormAnswer.asp?FormName=' + oForm.name + '&ItemID=' + iFormFieldID + '&FormID=' + aFormFields[iFormFieldID][0] + '&FormFieldID=' + aFormFields[iFormFieldID][1] + '&AnswerID=" & aFormComponent(N_ANSWER_ID_FORM) & "&TextAnswer=1&bOnlyOne=' + iOnlyOne + '&' + sURL;" & vbNewLine
									Response.Write "break;" & vbNewLine
								Response.Write "case '8':" & vbNewLine
									Response.Write "window.document.SaveFormAnswerIFrame.location.href = 'SaveFormAnswer.asp?FormName=' + oForm.name + '&ItemID=' + iFormFieldID + '&FormID=' + aFormFields[iFormFieldID][0] + '&FormFieldID=' + aFormFields[iFormFieldID][1] + '&AnswerID=" & aFormComponent(N_ANSWER_ID_FORM) & "&Answer=' + GetSelectedValues(oField) + '&bOnlyOne=' + iOnlyOne + '&' + sURL;" & vbNewLine
									Response.Write "break;" & vbNewLine
								Response.Write "case '9':" & vbNewLine
									Response.Write "window.document.SaveFormAnswerIFrame.location.href = 'SaveFormAnswer.asp?FormName=' + oForm.name + '&ItemID=' + iFormFieldID + '&FormID=' + aFormFields[iFormFieldID][0] + '&FormFieldID=' + aFormFields[iFormFieldID][1] + '&AnswerID=" & aFormComponent(N_ANSWER_ID_FORM) & "&Answer=' + GetSelectedValues(oField) + '&bOnlyOne=' + iOnlyOne + '&' + sURL;" & vbNewLine
									Response.Write "break;" & vbNewLine
								Response.Write "default:" & vbNewLine
									Response.Write "window.document.SaveFormAnswerIFrame.location.href = 'SaveFormAnswer.asp?FormName=' + oForm.name + '&ItemID=' + iFormFieldID + '&FormID=' + aFormFields[iFormFieldID][0] + '&FormFieldID=' + aFormFields[iFormFieldID][1] + '&AnswerID=" & aFormComponent(N_ANSWER_ID_FORM) & "&Answer=' + escape(oField.value) + '&bOnlyOne=' + iOnlyOne + '&' + sURL;" & vbNewLine
									Response.Write "break;" & vbNewLine
							Response.Write "}" & vbNewLine
							
						Response.Write "} else {" & vbNewLine
							Response.Write "if (! bOnlyOne)" & vbNewLine
								Response.Write "SaveAnswers(oForm, iFormFieldID + 1, bOnlyOne, sURL);" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "} else {" & vbNewLine
						Response.Write "HidePopupItem('Wait2Div', document.Wait2Div);" & vbNewLine
						Response.Write "window.location.reload();" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (bOnlyOne)" & vbNewLine
						Response.Write "HidePopupItem('Wait2Div', document.Wait2Div);" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of SaveAnswers" & vbNewLine
		Response.Write "//--></SCRIPT></DONT_EXPORT>" & vbNewLine
		Response.Write "<IFRAME SRC=""SaveFormAnswer.asp"" NAME=""SaveFormAnswerIFrame"" FRAMEBORDER=""0"" WIDTH=""0"" HEIGHT=""0""></IFRAME>"
		Response.Write "<DIV ID=""Wait2Div"" CLASS=""ClassPopupItem"" STYLE=""top: 200px; visibility: none;"">"
			Response.Write "<TABLE BGCOLOR=""#FFFFFF"" WIDTH=""100%"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR><TD ALIGN=""CENTER"">"
				Response.Write "<IMG SRC=""Images/AniWait.gif"" WIDTH=""100"" HEIGHT=""100"" ALT=""Guardando información..."" /><BR /><BR />"
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Guardando información...</B></FONT>"
			Response.Write "</TD></TR></TABLE>"
		Response.Write "</DIV>"
	End If

	sErrorDescription = "No se pudo obtener el formulario para el módulo."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select FormName, FormFile, PrintFile, FormDescription, FormFields.*, ReceiversID From Forms, FormFields, Tasks Where (Forms.FormID=FormFields.FormID) And (Forms.FormID=Tasks.FormID) And (Forms.Active=1) And (Tasks.ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (Tasks.TaskID=" & aTaskComponent(N_ID_TASK) & ") Order By FormFields.FormID, FormFieldID", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		iCurrentForm = -1
		sFormFile = ""
		sFormContents = ""
		bTemplate = bUseTemplate
		If Not oRecordset.EOF Then
			lFormID = CLng(oRecordset.Fields("FormID").Value)
			sReceiversID = CStr(oRecordset.Fields("ReceiversID").Value)
			If Len(sFormName) > 0 Then Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AnswerID"" ID=""AnswerIDHdn"" VALUE=""" & aFormComponent(N_ANSWER_ID_FORM) & """ />"
			Do While Not oRecordset.EOF
				If iCurrentForm <> CLng(oRecordset.Fields("FormID").Value) Then
					If Len(sFormContents) > 0 Then
						sFormContents = Replace(sFormContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, -1))
						lErrorNumber = GetNameFromTable(oADODBConnection, "Users", aFormComponent(N_USER_ID_FORM), "", "", sUserName, sErrorDescription)
						sFormContents = Replace(sFormContents, "<USER_COMPLETE_NAME />", CleanStringForHTML(sUserName))
						Response.Write sFormContents
						sFormContents = ""
					End If

					If iCurrentForm > -1 Then Response.Write "<BR /><HR /><BR />"
					If Len(sFormName) > 0 Then
						If iCurrentForm > -1 Then Response.Write "</DIV>"
						Response.Write "<DIV NAME=""Form" & CLng(oRecordset.Fields("FormID").Value) & "Div"" ID=""Form" & CLng(oRecordset.Fields("FormID").Value) & "Div"" CLASS=""FormForModule"">"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FormID"" ID=""FormIDHdn"" VALUE=""" & CLng(oRecordset.Fields("FormID").Value) & """ />"
					End If
					sFormFile = ""
					If bForEmail Or (StrComp(GetASPFileName(""), "Export.asp", vbBinaryCompare) = 0) Then
						sFormFile = CStr(oRecordset.Fields("PrintFile").Value)
						Err.Clear
						If Len(sFormFile) = 0 Then sFormFile = CStr(oRecordset.Fields("FormFile").Value)
					Else
						sFormFile = CStr(oRecordset.Fields("FormFile").Value)
					End If
					Err.Clear
					bTemplate = bUseTemplate And (Len(sFormFile) > 0)
					If bTemplate Then
						bASP = (StrComp(Right(sFormFile, Len(".asp")), ".asp", vbBinaryCompare) = 0)
						If Not bASP Then
							sFormContents = GetFileContents(Server.MapPath(TEMPLATES_PHYSICAL_PATH & sFormFile), sErrorDescription)
							If (Len(sFormName) = 0) And (Not bForEmail) Then Call CleanDontExport(sFormContents)
							lErrorNumber = TransformXMLTags(aFormComponent, sFormContents, aTaskComponent, sErrorDescription)
						End If
					Else
						Response.Write "<CENTER><B>" & CleanStringForHTML(CStr(oRecordset.Fields("FormName").Value)) & "</B></CENTER><BR />"
						Response.Write CleanStringForHTML(CStr(oRecordset.Fields("FormDescription").Value)) & "<BR /><BR />"
					End If
					iCurrentForm = CLng(oRecordset.Fields("FormID").Value)
				End If
				sField = ""
				sQueryForSource = ""
				sQueryForSource = oRecordset.Fields("QueryForSource").Value
				Err.Clear
				If Len(sQueryForSource) > 0 Then
					Call TransformXMLTags(aFormComponent, sQueryForSource, aTaskComponent, sErrorDescription)
					lErrorNumber = CreateADODBConnection(CStr(oRecordset.Fields("DSNForSource").Value), "", "", CInt(oRecordset.Fields("ConnectionTypeForSource").Value), oFieldADODBConnection, sErrorDescription)
				End If
				If (lErrorNumber = 0) And (Len(sFormName) > 0) Then
					If Not bASP Then
						Select Case CLng(oRecordset.Fields("FieldTypeID").Value)
							Case 0 'Booleano
								sField = "<INPUT TYPE=""CHECKBOX"" NAME=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & """ ID=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Chk"" VALUE=""1"" "
									If Not IsNull(oRecordset.Fields("DefaultValue")) Then
										If Len(oRecordset.Fields("DefaultValue")) > 0 Then sField = sField & " CHECKED=""1"""
									End If
									If Not IsNull(oRecordset.Fields("JavaScriptCode")) Then sField = sField & CStr(oRecordset.Fields("JavaScriptCode").Value)
								sField = sField & " />"
								If Not bTemplate Then sField = sField & CStr(oRecordset.Fields("FormFieldText").Value) & "<BR />"
							Case 1 'Fecha
								If Not bTemplate Then sField = CStr(oRecordset.Fields("FormFieldText").Value) & ": "
								If Not IsNull(oRecordset.Fields("DefaultValue")) Then
									sAnswer = CStr(oRecordset.Fields("DefaultValue").Value)
								Else
									sAnswer = "0"
								End If
								If CLng(oRecordset.Fields("MinimumValue").Value) < 1 Then
									sField = sField & DisplayDateCombosUsingSerial(sAnswer, "FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value), Year(Date()) - Abs(CLng(oRecordset.Fields("MinimumValue").Value)), Year(Date()) + Abs(CLng(oRecordset.Fields("MaximumValue").Value)), True, (CInt(oRecordset.Fields("IsOptional").Value) = 1))
								Else
									sField = sField & DisplayDateCombosUsingSerial(sAnswer, "FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value), CLng(oRecordset.Fields("MinimumValue").Value), Year(Date()) + Abs(CLng(oRecordset.Fields("MaximumValue").Value)), True, (CInt(oRecordset.Fields("IsOptional").Value) = 1))
								End If
								If Not bTemplate Then sField = sField & "<BR />"
							Case 2, 4 'Flotante, Numérico
								If Not bTemplate Then sField = CStr(oRecordset.Fields("FormFieldText").Value) & ": "
								sField = sField & "<INPUT TYPE=""TEXT"" NAME=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & """ ID=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Txt"" SIZE=""20"" MAXLENGTH=""20"" VALUE="""
									If Not IsNull(oRecordset.Fields("DefaultValue")) Then sField = sField & CStr(oRecordset.Fields("DefaultValue").Value)
								sField = sField & """ CLASS=""TextFields"" "
									If Not IsNull(oRecordset.Fields("JavaScriptCode")) Then sField = sField & CStr(oRecordset.Fields("JavaScriptCode").Value)
								sField = sField & " />"
								If Not bTemplate Then sField = sField & "<BR />"
							Case 3 'Hora
								If Not bTemplate Then sField = CStr(oRecordset.Fields("FormFieldText").Value) & ": "
								If Not IsNull(oRecordset.Fields("DefaultValue")) Then
									sAnswer = CStr(oRecordset.Fields("DefaultValue").Value)
								Else
									sAnswer = "0"
								End If
								sField = sField & DisplayTimeCombosUsingSerial(sAnswer, "FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value), 0, 24, 1, (CInt(oRecordset.Fields("IsOptional").Value) = 1))
								If Not bTemplate Then sField = sField & "<BR />"
							Case 5 'Texto
								If Not bTemplate Then sField = CStr(oRecordset.Fields("FormFieldText").Value) & ": "
								If CLng(oRecordset.Fields("FormFieldSize").Value) > 255 Then
									If Not bTemplate Then sField = sField & "<BR />"
									sField = sField & "<TEXTAREA NAME=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & """ ID=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "TxtArea"" "
									If CLng(oRecordset.Fields("FormFieldSize").Value) < 1000 Then
										sField = sField & "ROWS=""6"""
									Else
										sField = sField & "ROWS=""20"""
									End If
									sField = sField & " COLS=""50"" MAXLENGTH=""" & CStr(oRecordset.Fields("FormFieldSize").Value) & """ CLASS=""TextFields"" "
										If Not IsNull(oRecordset.Fields("JavaScriptCode")) Then sField = sField & CStr(oRecordset.Fields("JavaScriptCode").Value)
									sField = sField & ">"
								Else
									sField = sField & "<INPUT TYPE=""TEXT"" NAME=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & """ ID=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Txt"" SIZE="""
									If CLng(oRecordset.Fields("FormFieldSize").Value) < 30 Then
										sField = sField & CStr(oRecordset.Fields("FormFieldSize").Value)
									Else
										sField = sField & "30"
									End If
									sField = sField & """ MAXLENGTH=""" & CStr(oRecordset.Fields("FormFieldSize").Value) & """ VALUE="""
								End If
									If Len(sQueryForSource) > 0 Then
										sErrorDescription = "No se pudo obtener el valor por default para el campo de texto del formulario."
										lErrorNumber = ExecuteSQLQuery(oFieldADODBConnection, sQueryForSource, "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oCatalogRecordset)
										If lErrorNumber = 0 Then
											If Not oCatalogRecordset.EOF Then
												For iIndex = 0 To oCatalogRecordset.Fields.Count - 1
													sField = sField & CStr(oCatalogRecordset.Fields(iIndex).Value)
													Err.Clear
													If iIndex < oCatalogRecordset.Fields.Count - 1 Then sField = sField & " "
												Next
											End If
										End If
									ElseIf Not IsNull(oRecordset.Fields("DefaultValue")) Then
										sField = sField & CStr(oRecordset.Fields("DefaultValue").Value)
									End If
								If CLng(oRecordset.Fields("FormFieldSize").Value) > 255 Then
									sField = sField & "</TEXTAREA>"
									If Not bTemplate Then sField = sField & "<BR />"
								Else
									sField = sField & """ CLASS=""TextFields"" "
										If Not IsNull(oRecordset.Fields("JavaScriptCode")) Then sField = sField & CStr(oRecordset.Fields("JavaScriptCode").Value)
									sField = sField & " />"
								End If
								If Not bTemplate Then sField = sField & "<BR />"
							Case 6, 8 'Catálogo, Lista
								If lErrorNumber = 0 Then
									If Not bTemplate Then
										sField = CStr(oRecordset.Fields("FormFieldText").Value) & ": "
										If CLng(oRecordset.Fields("FieldTypeID").Value) = 8 Then sField = sField & "<BR />"
									End If
									sField = sField & "<SELECT NAME=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & """ ID=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Cmb"""
									If CLng(oRecordset.Fields("FieldTypeID").Value) = 6 Then
										sField = sField & " SIZE=""1"""
									Else
										sField = sField & " SIZE=""5"" MULTIPLE=""1"""
									End If
									sField = sField & " VALUE="""" CLASS=""Lists"" "
										If Not IsNull(oRecordset.Fields("JavaScriptCode")) Then sField = sField & CStr(oRecordset.Fields("JavaScriptCode").Value)
									sField = sField & " >"
										If Len(sQueryForSource) > 0 Then
											asFields = Split(sQueryForSource, LIST_SEPARATOR, -1, vbBinaryCompare)
											sAnswer = ""
											sAnswer = CStr(oRecordset.Fields("DefaultValue").Value)
											Err.Clear
											sField = sField & GenerateListOptionsFromQuery(oADODBConnection, asFields(0), asFields(1), asFields(3), asFields(4), asFields(5), sAnswer, "", sErrorDescription)
										End If
									sField = sField & "</SELECT>"
									If Not bTemplate Then sField = sField & "<BR />"
								End If
							Case 7, 9 'Catálogo jerárquico, Lista jerárquica
								If lErrorNumber = 0 Then
									If Not bTemplate Then
										sField = CStr(oRecordset.Fields("FormFieldText").Value) & ": "
										If CLng(oRecordset.Fields("FieldTypeID").Value) = 9 Then sField = sField & "<BR />"
									End If
									sField = sField & "<SELECT NAME=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & """ ID=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Cmb"""
									If CLng(oRecordset.Fields("FieldTypeID").Value) = 7 Then
										sField = sField & " SIZE=""1"""
									Else
										sField = sField & " SIZE=""5"" MULTIPLE=""1"""
									End If
									sField = sField & " VALUE="""" CLASS=""Lists"" "
										If Not IsNull(oRecordset.Fields("JavaScriptCode")) Then sField = sField & CStr(oRecordset.Fields("JavaScriptCode").Value)
									sField = sField & " >"
										asFields = Split(sQueryForSource, LIST_SEPARATOR, -1, vbBinaryCompare)
										sAnswer = ""
										sAnswer = CStr(oRecordset.Fields("DefaultValue").Value)
										Err.Clear
										sField = sField & GenerateHierarchyListOptionsFromQuery(oFieldADODBConnection, asFields(0), asFields(1), asFields(2), asFields(3), asFields(4), -1, asFields(5), sAnswer, "", "", sField, sErrorDescription)
									sField = sField & "</SELECT>"
									If Not bTemplate Then sField = sField & "<BR />"
								End If
							Case 10 'Archivo
								sField = "<IFRAME SRC=""BrowserFile.asp?FormID=" & CStr(oRecordset.Fields("FormID").Value) & "&AnswerID=" & aFormComponent(N_ANSWER_ID_FORM) & "&FormFieldID=" & CStr(oRecordset.Fields("FormFieldID").Value) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & """ NAME=""Form_" & CStr(oRecordset.Fields("FormID").Value) & "_" & aFormComponent(N_ANSWER_ID_FORM) & "_" & CStr(oRecordset.Fields("FormFieldID").Value) & "_FilesIFrame"" FRAMEBORDER=""1"" WIDTH=""400"" HEIGHT=""150""></IFRAME>"
							Case 11 'Oculto
								sField = "<INPUT TYPE=""HIDDEN"" NAME=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & """ ID=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Hdn"" VALUE="""
									If Len(sQueryForSource) > 0 Then
										sErrorDescription = "No se pudo obtener el valor por default para el campo de texto del formulario."
										lErrorNumber = ExecuteSQLQuery(oFieldADODBConnection, sQueryForSource, "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oCatalogRecordset)
										If lErrorNumber = 0 Then
											If Not oCatalogRecordset.EOF Then
												For iIndex = 0 To oCatalogRecordset.Fields.Count - 1
													sField = sField & CStr(oCatalogRecordset.Fields(iIndex).Value)
													Err.Clear
													If iIndex < oCatalogRecordset.Fields.Count - 1 Then sField = sField & " "
												Next
											End If
										End If
									ElseIf Not IsNull(oRecordset.Fields("DefaultValue")) Then
										sField = sField & CStr(oRecordset.Fields("DefaultValue").Value)
									End If
								sField = sField & """ />"
						End Select
						If bTemplate Then
							sFormContents = Replace(sFormContents, "<" & CStr(oRecordset.Fields("FormFieldName").Value) & " />", sField)
						Else
							Response.Write sField
						End If
					Else
						Select Case CLng(oRecordset.Fields("FieldTypeID").Value)
							Case 0, 2, 4, 5, 6, 7, 8, 9 'Booleano, Flotante, Numérico, Texto, Catálogo jerárquico, Lista jerárquica, Catálogo, Lista
								sField = "<INPUT TYPE=""HIDDEN"" NAME=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & """ ID=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Hdn"" VALUE="""" />"
							Case 1 'Fecha
								sField = "<INPUT TYPE=""HIDDEN"" NAME=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Year"" ID=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "YearHdn"" VALUE="""" />"
								sField = sField & "<INPUT TYPE=""HIDDEN"" NAME=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Month"" ID=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "MonthHdn"" VALUE="""" />"
								sField = sField & "<INPUT TYPE=""HIDDEN"" NAME=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Day"" ID=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "DayHdn"" VALUE="""" />"
							Case 3 ' Hora
								sField = "<INPUT TYPE=""HIDDEN"" NAME=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Hour"" ID=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "HourHdn"" VALUE="""" />"
								sField = sField & "<INPUT TYPE=""HIDDEN"" NAME=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Minute"" ID=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "MinuteHdn"" VALUE="""" />"
								sField = sField & "<INPUT TYPE=""HIDDEN"" NAME=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Second"" ID=""FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "SecondHdn"" VALUE="""" />"
						End Select
						Response.Write sField
					End If
				End If
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close

			If Len(sFormName) = 0 Then
				sErrorDescription = "No se pudieron obtener las respuestas del usuario para el formulario."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select FormFieldsAnswers.FormID, FormFieldsAnswers.FormFieldID, FormFieldName, Answer, FieldTypeID, QueryForSource From FormFieldsAnswers, FormFields, Forms, Tasks Where (FormFieldsAnswers.FormID=FormFields.FormID) And (FormFieldsAnswers.FormFieldID=FormFields.FormFieldID) And (FormFieldsAnswers.FormID=Forms.FormID) And (Forms.FormID=Tasks.FormID) And (FormFieldsAnswers.AnswerID = " & aFormComponent(N_ANSWER_ID_FORM) & ") And (Tasks.ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (Tasks.TaskID=" & aTaskComponent(N_ID_TASK) & ") Order By FormFieldsAnswers.FormID, FormFieldsAnswers.FormFieldID", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						sAnswer = CStr(oRecordset.Fields("Answer").Value)
						Select Case CInt(oRecordset.Fields("FieldTypeID").Value)
							Case 0 'Booleano
								sFormContents = Replace(sFormContents, "<" & CStr(oRecordset.Fields("FormFieldName").Value) & " />", DisplayYesNo(CInt(sAnswer), False))
							Case 1 'Fecha
								sFormContents = Replace(sFormContents, "<" & CStr(oRecordset.Fields("FormFieldName").Value) & " />", DisplayDateAndTimeFromSerialNumber(sAnswer, ""))
							Case 3 'Hora
								sFormContents = Replace(sFormContents, "<" & CStr(oRecordset.Fields("FormFieldName").Value) & " />", DisplayTimeFromSerialNumber(sAnswer & "00"))
							Case 6, 8 'Catálogo, Lista
								asFields = Split(CStr(oRecordset.Fields("QueryForSource").Value), LIST_SEPARATOR, -1, vbBinaryCompare)
								sErrorDescription = "No se pudieron obtener las respuestas del usuario para el formulario."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & asFields(3) & " From " & asFields(0) & " Where (" & asFields(1) & "=" & sAnswer & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oCatalogRecordset)
								If lErrorNumber = 0 Then
									If Not oCatalogRecordset.EOF Then
										sFormContents = Replace(sFormContents, "<" & CStr(oRecordset.Fields("FormFieldName").Value) & " />", CStr(oCatalogRecordset.Fields(asFields(3)).Value))
									End If
								End If
							Case 7, 9 'Catálogo jerárquico, Lista jerárquica
								sFormContents = Replace(sFormContents, "<" & CStr(oRecordset.Fields("FormFieldName").Value) & " />", CleanStringForHTML(sAnswer))
							Case Else 'Flotante, Numérico, Texto
								sFormContents = Replace(sFormContents, "<" & CStr(oRecordset.Fields("FormFieldName").Value) & " />", CleanStringForHTML(sAnswer))
						End Select
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
					If Len(sURL) > 0 Then
						sURL = Left(sURL, (Len(sURL) - Len("&")))
						Response.Write "SendURLValuesToForm('" & sURL & "', document." & sFormName & ");" & vbNewLine
					End If
				End If

				sErrorDescription = "No se pudieron obtener las respuestas del usuario para el formulario."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select FormID, FormFieldID, FormFieldName From FormFields Where (FormID=" & aFormComponent(N_ID_FORM) & ") And (FieldTypeID=10) Order By FormID, FormFieldID", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						sAnswer = ""
						Call GetFolderContents(Server.MapPath(INSPECTIONS_PHYSICAL_PATH & "i" & CStr(oRecordset.Fields("FormID").Value) & "_" & aFormComponent(N_ANSWER_ID_FORM) & "_" & CStr(oRecordset.Fields("FormFieldID").Value)), False, sAnswer, sErrorDescription)
						asFiles = Split(sAnswer, LIST_SEPARATOR)
						sAnswer = ""
						For iIndex = 0 To UBound(asFiles)
							If (InStr(1, LCase(asFiles(iIndex)), ".bmp") <> 0) Or (InStr(1, LCase(asFiles(iIndex)), ".gif") <> 0) Or (InStr(1, LCase(asFiles(iIndex)), ".jpg") <> 0) Or (InStr(1, LCase(asFiles(iIndex)), ".png") <> 0) Then
								sAnswer = sAnswer & "<IMG SRC=""" & S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME & INSPECTIONS_PATH & "i" & CStr(oRecordset.Fields("FormID").Value) & "_" & aFormComponent(N_ANSWER_ID_FORM) & "_" & CStr(oRecordset.Fields("FormFieldID").Value) & "/" & asFiles(iIndex) & """ /><BR />"
							Else
								sAnswer = sAnswer & "<A HREF=""" & S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME & INSPECTIONS_PATH & "i" & CStr(oRecordset.Fields("FormID").Value) & "_" & aFormComponent(N_ANSWER_ID_FORM) & "_" & CStr(oRecordset.Fields("FormFieldID").Value) & "/" & asFiles(iIndex) & """ TARGET=""_blank"">" & asFiles(iIndex) & "</A><BR />"
							End If
						Next
						sFormContents = Replace(sFormContents, "<" & CStr(oRecordset.Fields("FormFieldName").Value) & " />", sAnswer)
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
				End If
			End If
			If bTemplate Then
				If Not bASP Then
					If Len(sFormContents) > 0 Then
						sFormContents = Replace(sFormContents, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, -1))
						lErrorNumber = GetNameFromTable(oADODBConnection, "Users", aFormComponent(N_USER_ID_FORM), "", "", sUserName, sErrorDescription)
						sFormContents = Replace(sFormContents, "<USER_COMPLETE_NAME />", CleanStringForHTML(sUserName))
						If Not bForEmail Then
							Response.Write sFormContents
						Else
							ReDim aEmailComponent(N_EMAIL_COMPONENT_SIZE)
							Call GetNameFromTable(oADODBConnection, "UsersEmail", sReceiversID, "", ";", aEmailComponent(S_TO_EMAIL), "")
							aEmailComponent(S_FROM_EMAIL) = aLoginComponent(S_USER_E_MAIL_LOGIN)
							aEmailComponent(S_SUBJECT_EMAIL) = "Formulario enviado por el Sistema de Administración Vehicular"
							aEmailComponent(S_BODY_EMAIL) = "<SPAN STYLE=""background-color: " & S_MAIN_COLOR_FOR_GUI & "; width: 100%;""><CENTER><FONT FACE=""Arial"" SIZE=""2"" COLOR=""#" & S_MENU_LINK_FOR_GUI & """><B>Para entrar al Sistema de Administración Vehicular y ver este formulario, <A HREF=""" & S_HTTP & EXT_SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME & "Project.asp?TaskID=" & aTaskComponent(N_ID_TASK) & "&Change=1&Tab=2""><FONT COLOR=""#" & S_SELECTED_LINK_FOR_GUI & """>PRESIONE AQUÍ</FONT></A></B></FONT><CENTER></SPAN><BR /><BR />"
							aEmailComponent(S_BODY_EMAIL) = aEmailComponent(S_BODY_EMAIL) & sFormContents
							lErrorNumber = SendEmail(oRequest, aEmailComponent, sErrorDescription)
						End If
						sFormContents = ""
					End If
					If Len(sFormFile) > 0 Then sFormContents = GetFileContents(Server.MapPath(TEMPLATES_PHYSICAL_PATH & sFormFile), sErrorDescription)
				Else
					If Len(sFormName) > 0 Then
						Response.Write "<IFRAME SRC=""Templates/" & sFormFile & "?FormID=" & aFormComponent(N_ID_FORM)
						Response.Write """ NAME=""PaperworkFormIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""355""></IFRAME>"
					End If
				End If
			End If
			If Len(sFormName) > 0 Then Response.Write "</DIV>"

			If Len(sFormName) > 0 Then
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					If Len(aFormComponent(S_ANSWER_ID_FORM)) > 0 Then
						asIDs = Split(aFormComponent(S_ANSWER_ID_FORM), ";")
						For iIndex = 0 To UBound(asIDs)
							asIDs(iIndex) = Split(asIDs(iIndex), ",")
							Call GetNameFromTable(oADODBConnection, "FormModule", asIDs(iIndex)(1), "", "", asIDs(iIndex)(1), "")
							If CInt(asIDs(iIndex)(1)) = aFormComponent(N_MODULE_FORM) Then aFormComponent(N_ANSWER_ID_FORM) = asIDs(iIndex)(0)
						Next
					End If
					sErrorDescription = "No se pudieron obtener las respuestas del usuario para el formulario."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select FormFieldsAnswers.FormID, FormFieldsAnswers.FormFieldID, Answer, FieldTypeID From FormFieldsAnswers, FormFields, Forms, Tasks Where (FormFieldsAnswers.FormID=FormFields.FormID) And (FormFieldsAnswers.FormFieldID=FormFields.FormFieldID) And (FormFieldsAnswers.FormID=Forms.FormID) And (Forms.FormID=Tasks.FormID) And (FormFieldsAnswers.AnswerID = " & aFormComponent(N_ANSWER_ID_FORM) & ") And (Tasks.ProjectID=" & aTaskComponent(N_PROJECT_ID_TASK) & ") And (Tasks.TaskID=" & aTaskComponent(N_ID_TASK) & ") Order By FormFieldsAnswers.FormID, FormFieldsAnswers.FormFieldID", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						aFormComponent(B_ANSWERED_FORM) = (Not oRecordset.EOF)
						sURL = ""
						Do While Not oRecordset.EOF
							sAnswer = CStr(oRecordset.Fields("Answer").Value)
							Select Case CInt(oRecordset.Fields("FieldTypeID").Value)
								Case 1
									sURL = sURL & "FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Year=" & Left(sAnswer, 4) & "&"
									sURL = sURL & "FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Month=" & Mid(sAnswer, 5, 2) & "&"
									sURL = sURL & "FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Day=" & Right(sAnswer, 2) & "&"
								Case 3
									sURL = sURL & "FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Hour=" & Left(sAnswer, 2) & "&"
									sURL = sURL & "FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "Minute=" & Right(sAnswer, 2) & "&"
								Case 8, 9
									asValues = Split(sAnswer, ", ", -1, vbBinaryCompare)
									For iIndex = 0 To UBound(asValues)
										Response.Write "SelectItemByValue('" & asValues(iIndex) & "', true, document." & sFormName & ".FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & ");" & vbNewLine
										If Err.number <> 0 Then Exit For
									Next
								Case Else
									sURL = sURL & "FF__" & CStr(oRecordset.Fields("FormID").Value) & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & "=" & CleanStringForJavaScript(sAnswer) & "&"
							End Select
							oRecordset.MoveNext
							If Err.number <> 0 Then Exit Do
						Loop
						oRecordset.Close
						If Len(sURL) > 0 Then
							sURL = Left(sURL, (Len(sURL) - Len("&")))
							Response.Write "SendURLValuesToForm('" & sURL & "', document." & sFormName & ");" & vbNewLine
						End If
					End If

					Response.Write "function CheckFormForModule(oForm) {" & vbNewLine
						sErrorDescription = "No se pudieron obtener los campos obligatorios para el formulario."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select FormFieldID, FormFieldText, FieldTypeID, FormFieldSize, LimitTypeID, MinimumValue, MaximumValue, IsOptional From FormFields Where (FormID=" & lFormID & ") Order By FormFieldID", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							Do While Not oRecordset.EOF
								Select Case CInt(oRecordset.Fields("LimitTypeID").Value)
									Case 0	'Ninguno
										sMinType = "N_NO_RANK_FLAG"
										sMaxType = "N_CLOSED_FLAG"
									Case 1 'Sólo mínimo abierto
										sMinType = "N_MINIMUM_ONLY_FLAG"
										sMaxType = "N_OPEN_FLAG"
									Case 2	'Sólo máximo abierto
										sMinType = "N_MAXIMUM_ONLY_FLAG"
										sMaxType = "N_OPEN_FLAG"
									Case 3	'Mínimo abierto y máximo abierto
										sMinType = "N_BOTH_FLAG"
										sMaxType = "N_OPEN_FLAG"
									Case 5	'Sólo mínimo cerrado
										sMinType = "N_MINIMUM_ONLY_FLAG"
										sMaxType = "N_CLOSED_FLAG"
									Case 7	'Mínimo cerrado y máximo abierto
										sMinType = "N_BOTH_FLAG"
										sMaxType = "N_MAXIMUM_OPEN_FLAG"
									Case 10	'Sólo máximo cerrado
										sMinType = "N_MAXIMUM_ONLY_FLAG"
										sMaxType = "N_CLOSED_FLAG"
									Case 11	'Mínimo abierto y máximo cerrado
										sMinType = "N_BOTH_FLAG"
										sMaxType = "N_MINIMUM_OPEN_FLAG"
									Case 15	'Mínimo cerrado y máximo cerrado
										sMinType = "N_BOTH_FLAG"
										sMaxType = "N_CLOSED_FLAG"
								End Select
								Response.Write vbTab & "if (oForm.FF__" & lFormID & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & ") {" & vbNewLine
									Select Case CInt(oRecordset.Fields("FieldTypeID").Value)
										Case 2 'Flotante
											If CInt(oRecordset.Fields("IsOptional").Value) = 1 Then Response.Write vbTab & "if (oForm.FF__" & lFormID & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value != '') {" & vbNewLine
												Response.Write vbTab & vbTab & "oForm.FF__" & lFormID & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value = oForm.FF__" & lFormID & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
												Response.Write vbTab & "if (! CheckFloatValue(oForm.FF__" & lFormID & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & ", '" & Replace(Replace(Replace(CStr(oRecordset.Fields("FormFieldText").Value), "\", "\\"), "/", "\/"), "'", "\'") & "', " & sMinType & ", " & sMaxType & ", " & CStr(oRecordset.Fields("MinimumValue").Value) & ", " & CStr(oRecordset.Fields("MaximumValue").Value) & ")) {" & vbNewLine
													Response.Write vbTab & vbTab & "oForm.FF__" & lFormID & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".focus();" & vbNewLine
													Response.Write vbTab & vbTab & "return false;" & vbNewLine
												Response.Write vbTab & "}" & vbNewLine
											If CInt(oRecordset.Fields("IsOptional").Value) = 1 Then Response.Write vbTab & "}" & vbNewLine
										Case 4 'Numérico
											If CInt(oRecordset.Fields("IsOptional").Value) = 1 Then Response.Write vbTab & "if (oForm.FF__" & lFormID & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value != '') {" & vbNewLine
												Response.Write vbTab & vbTab & "oForm.FF__" & lFormID & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value = oForm.FF__" & lFormID & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
												Response.Write vbTab & "if (! CheckIntegerValue(oForm.FF__" & lFormID & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & ", '" & Replace(Replace(Replace(CStr(oRecordset.Fields("FormFieldText").Value), "\", "\\"), "/", "\/"), "'", "\'") & "', " & sMinType & ", " & sMaxType & ", " & CStr(oRecordset.Fields("MinimumValue").Value) & ", " & CStr(oRecordset.Fields("MaximumValue").Value) & ")) {" & vbNewLine
													Response.Write vbTab & vbTab & "oForm.FF__" & lFormID & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".focus();" & vbNewLine
													Response.Write vbTab & vbTab & "return false;" & vbNewLine
												Response.Write vbTab & "}" & vbNewLine
											If CInt(oRecordset.Fields("IsOptional").Value) = 1 Then Response.Write vbTab & "}" & vbNewLine
										Case 5 'Texto
											If CInt(oRecordset.Fields("IsOptional").Value) = 0 Then
												Response.Write vbTab & "if (oForm.FF__" & lFormID & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value.length == 0) {" & vbNewLine
													Response.Write vbTab & vbTab & "alert('Favor de introducir la información para el campo " & Replace(Replace(Replace(CStr(oRecordset.Fields("FormFieldText").Value), "\", "\\"), "/", "\/"), "'", "\'") & ".');" & vbNewLine
													Response.Write vbTab & vbTab & "oForm.FF__" & lFormID & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".focus();" & vbNewLine
													Response.Write vbTab & vbTab & "return false;" & vbNewLine
												Response.Write vbTab & "}" & vbNewLine
												If CInt(oRecordset.Fields("MinimumValue").Value) > 0 Then
													Response.Write vbTab & vbTab & "if (oForm.FF__" & lFormID & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".value.length < " & CStr(oRecordset.Fields("MinimumValue").Value) & ") {" & vbNewLine
														Response.Write vbTab & vbTab & vbTab & "ShowTaskTab(2);" & vbNewLine
														Response.Write vbTab & vbTab & vbTab & "alert('El campo " & Replace(Replace(Replace(CStr(oRecordset.Fields("FormFieldText").Value), "\", "\\"), "/", "\/"), "'", "\'") & " requiere al menos " & CStr(oRecordset.Fields("MinimumValue").Value) & " caracteres.');" & vbNewLine
														Response.Write vbTab & vbTab & vbTab & "window.setTimeout('oForm.FF__" & lFormID & "__" & CStr(oRecordset.Fields("FormFieldID").Value) & ".focus()', 1000);" & vbNewLine
														Response.Write vbTab & vbTab & vbTab & "return false;" & vbNewLine
													Response.Write vbTab & vbTab & "}" & vbNewLine
												End If
											End If
									End Select
								Response.Write vbTab & "}" & vbNewLine
								oRecordset.MoveNext
								If Err.number <> 0 Then Exit Do
							Loop
						End If
						If InStr(1, sFormContents, "function CheckTemplate(", vbBinaryCompare) > 0 Then
							Response.Write vbTab & "return CheckTemplate(oForm);" & vbNewLine
						End If
						Response.Write vbTab & "return true;" & vbNewLine
					Response.Write "} // End of CheckFormForModule" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
			End If
		End If
	End If

	Set oRecordset = Nothing
	DisplayFormForTask = lErrorNumber
	Err.Clear
End Function

Function CleanDontExport(sFormContents)
'************************************************************
'Purpose: To remove the <DONT_EXPORT> tags
'Inputs:  sFormContents
'Outputs: sFormContents
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CleanDontExport"
	Dim iStartPos
	Dim iEndPos

	iStartPos = InStr(1, sFormContents, "<DONT_EXPORT>", vbBinaryCompare)
	Do While (iStartPos > 0)
		iStartPos = iStartPos - Len("<")
		iEndPos = InStr(iStartPos, sFormContents, "</DONT_EXPORT>", vbBinaryCompare)
		If iEndPos > 0 Then
			iEndPos = iEndPos + Len("</DONT_EXPORT")
			sFormContents = Left(sFormContents, iStartPos) & Right(sFormContents, (Len(sFormContents) - iEndPos))
		End If
		iStartPos = InStr(1, sFormContents, "<DONT_EXPORT>", vbBinaryCompare)
		If Err.number <> 0 Then Exit Do
	Loop

	CleanDontExport = Err.number
	Err.Clear
End Function

Function TransformXMLTags(aFormComponent, sText, aTaskComponent, sErrorDescription)
'************************************************************
'Purpose: To replace the XML tags using the entries from the
'         database
'Inputs:  aFormComponent, sText, aTaskComponent
'Outputs: sText, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "TransformXMLTags"
	Dim lFormID
	Dim iStartPos
	Dim iMidPos
	Dim iEndPos
	Dim lDate
	Dim sFormFieldName
	Dim asFields
	Dim sAnswer
	Dim sFormAnswers
	Dim sAccessKey
	Dim sPassword
	Dim sCondition
	Dim asTemp
	Dim iIndex
	Dim oRecordset
	Dim oCatalogRecordset
	Dim lErrorNumber

	sText = Replace(sText, "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME, 1, -1, vbBinaryCompare)
	sText = Replace(sText, "<EXT_SYSTEM_URL />", S_HTTP & EXT_SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME, 1, -1, vbBinaryCompare)
	sText = Replace(sText, "<SERVER_IP_FOR_LICENSE />", SERVER_IP_FOR_LICENSE, 1, -1, vbBinaryCompare)
	sText = Replace(sText, "<EXT_SERVER_IP_FOR_LICENSE />", EXT_SERVER_IP_FOR_LICENSE, 1, -1, vbBinaryCompare)
	sText = Replace(sText, "<CURRENT_DATE />", DisplayDateFromSerialNumber(Left(GetSerialNumberForDate(""), Len("00000000")), -1, -1, -1), 1, -1, vbBinaryCompare)
	sText = Replace(sText, "<CURRENT_YEAR />", Year(Date()), 1, -1, vbBinaryCompare)
	sText = Replace(sText, "<CURRENT_SERIAL_DATE />", Left(GetSerialNumberForDate(""), Len("00000000")), 1, -1, vbBinaryCompare)
	sText = Replace(sText, "<CURRENT_TIME />", DisplayTimeFromSerialNumber(""), 1, -1, vbBinaryCompare)
	sText = Replace(sText, "<START_YEAR />", Left(aTaskComponent(N_START_DATE_TASK), Len("0000")), 1, -1, vbBinaryCompare)
	sText = Replace(sText, "<START_MONTH />", Mid(aTaskComponent(N_START_DATE_TASK), Len("00000"), Len("00")), 1, -1, vbBinaryCompare)
	sText = Replace(sText, "<START_DAY />", Mid(aTaskComponent(N_START_DATE_TASK), Len("0000000"), Len("00")), 1, -1, vbBinaryCompare)
	sText = Replace(sText, "<ANSWER_ID />", aTaskComponent(N_ID_TASK), 1, -1, vbBinaryCompare)

	sErrorDescription = "No se pudo obtener la información del usuario."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Users Where (UserID=" & aFormComponent(N_USER_ID_FORM) & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sText = Replace(sText, "<USER_ID />", aFormComponent(N_USER_ID_FORM), 1, -1, vbBinaryCompare)
			sAccessKey = CStr(oRecordset.Fields("UserAccessKey").Value)
			sText = Replace(sText, "<USER_ACCESS_KEY />", sAccessKey, 1, -1, vbBinaryCompare)
			sPassword = CStr(oRecordset.Fields("UserPassword").Value)
			sText = Replace(sText, "<USER_PASSWORD />", sPassword, 1, -1, vbBinaryCompare)
			sText = Replace(sText, "<USER_COMPLETE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("UserName").Value) & " " & CStr(oRecordset.Fields("UserLastName").Value)), 1, -1, vbBinaryCompare)
			sText = Replace(sText, "<USER_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("UserName").Value)), 1, -1, vbBinaryCompare)
			sText = Replace(sText, "<USER_LAST_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("UserLastName").Value)), 1, -1, vbBinaryCompare)
			sText = Replace(sText, "<USER_EMAIL />", CleanStringForHTML(CStr(oRecordset.Fields("UserEmail").Value)), 1, -1, vbBinaryCompare)
			sText = Replace(sText, "<USER_PERMISSIONS />", CLng(oRecordset.Fields("UserPermissions").Value), 1, -1, vbBinaryCompare)
			oRecordset.Close
		End If
	End If

	iStartPos = InStr(1, sText, "<FORM_FIELD FORM_ID=""", vbBinaryCompare)
	Do While (iStartPos > 0)
		lFormID = -1
		sFormFieldName = ""
		iMidPos = InStr(iStartPos, sText, "FORM_ID=""", vbBinaryCompare) + Len("FORM_ID=""")
		iEndPos = InStr(iMidPos, sText, """", vbBinaryCompare)
		If (iMidPos > Len("FORM_ID=""")) And (iEndPos > 0) Then
			lFormID = CLng(Mid(sText, iMidPos, (iEndPos - iMidPos)))
			iMidPos = InStr(iStartPos, sText, "NAME=""", vbBinaryCompare) + Len("NAME=""")
			iEndPos = InStr(iMidPos, sText, """", vbBinaryCompare)
			If (iMidPos > Len("NAME=""")) And (iEndPos > 0) Then
				sFormFieldName = Mid(sText, iMidPos, (iEndPos - iMidPos))
				iEndPos = InStr(iMidPos, sText, "/>", vbBinaryCompare)
				iEndPos = iEndPos + Len("/>")
				If Err.number = 0 Then
					sErrorDescription = "No se pudo obtener la información del trámite."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Answer, FieldTypeID, QueryForSource From FormFieldsAnswers, FormFields Where (FormFieldsAnswers.FormID=FormFields.FormID) And (FormFieldsAnswers.FormFieldID=FormFields.FormFieldID) And (FormFieldsAnswers.FormID=" & lFormID & ") And (FormFieldName='" & sFormFieldName & "') And (AnswerID=" & aFormComponent(N_ANSWER_ID_FORM) & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						sAnswer = ""
						If Not oRecordset.EOF Then
							sAnswer = CStr(oRecordset.Fields("Answer").Value)
							Err.Clear
							Select Case CInt(oRecordset.Fields("FieldTypeID").Value)
								Case 0
									sAnswer = DisplayYesNo(CInt(sAnswer), False)
								Case 1
									sAnswer = DisplayDateFromSerialNumber(sAnswer, -1, -1, -1)
								Case 3
									sAnswer = DisplayTimeFromSerialNumber(Left(sAnswer, Len("0000")) & "00")
								Case 6, 8
									asFields = Split(CStr(oRecordset.Fields("QueryForSource").Value), LIST_SEPARATOR, -1, vbBinaryCompare)
									If Len(asFields(4)) > 0 Then sCondition = asFields(4) & " And "
									sErrorDescription = "No se pudieron obtener las respuestas del usuario para el formulario."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & asFields(3) & " From " & asFields(0) & " Where " & sCondition & "(" & asFields(1) & "=" & sAnswer & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oCatalogRecordset)
									If lErrorNumber = 0 Then
										If Not oCatalogRecordset.EOF Then
											asTemp = Split(Replace(asFields(3), " ", ""), ",")
											sAnswer = ""
											For iIndex = 0 To UBound(asTemp)
												sAnswer = sAnswer & CStr(oCatalogRecordset.Fields(asTemp(iIndex)).Value) & " "
											Next
										End If
									End If
							End Select
						End If
						sText = Replace(sText, "<FORM_FIELD FORM_ID=""" & lFormID & """ NAME=""" & sFormFieldName & """ />", CleanStringForHTML(sAnswer))
					End If
				End If
			End If
		End If
		iStartPos = InStr(1, sText, "<FORM_FIELD FORM_ID=""", vbBinaryCompare)
		If Err.number <> 0 Then Exit Do
	Loop

	iStartPos = InStr(1, sText, "<FORM_FIELD ", vbBinaryCompare)
	Do While (iStartPos > 0)
		sFormFieldName = ""
		iMidPos = InStr(iStartPos, sText, "NAME=""", vbBinaryCompare) + Len("NAME=""")
		iEndPos = InStr(iMidPos, sText, """", vbBinaryCompare)
		If (iMidPos > Len("NAME=""")) And (iEndPos > 0) Then
			sFormFieldName = Mid(sText, iMidPos, (iEndPos - iMidPos))
			iEndPos = InStr(iMidPos, sText, "/>", vbBinaryCompare)
			iEndPos = iEndPos + Len("/>")
			If Err.number = 0 Then
				sErrorDescription = "No se pudo obtener la información del trámite."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Answer, FieldTypeID From FormFieldsAnswers, FormFields Where (FormFieldsAnswers.FormID=FormFields.FormID) And (FormFieldsAnswers.FormFieldID=FormFields.FormFieldID) And (FormFields.FormID=" & aFormComponent(N_ID_FORM) & ") And (FormFieldName='" & sFormFieldName & "') And (FormFieldsAnswers.AnswerID = " & aFormComponent(N_ANSWER_ID_FORM) & ")", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					sAnswer = ""
					If Not oRecordset.EOF Then
						sAnswer = CStr(oRecordset.Fields("Answer").Value)
						Err.Clear
						Select Case CInt(oRecordset.Fields("FieldTypeID").Value)
							Case 0
								sAnswer = DisplayYesNo(CInt(sAnswer), False)
							Case 1
								sAnswer = DisplayDateFromSerialNumber(sAnswer, -1, -1, -1)
						End Select
					End If
					sText = Replace(sText, "<FORM_FIELD NAME=""" & sFormFieldName & """ />", CleanStringForHTML(sAnswer))
				End If
			End If
		End If
		iStartPos = InStr(1, sText, "<FORM_FIELD ", vbBinaryCompare)
		If Err.number <> 0 Then Exit Do
	Loop

	iStartPos = InStr(1, sText, "<FORM ", vbBinaryCompare)
	Do While (iStartPos > 0)
		lFormID = -1
		sFormFieldName = ""
		iMidPos = InStr(iStartPos, sText, "FORM_ID=""", vbBinaryCompare) + Len("FORM_ID=""")
		iEndPos = InStr(iMidPos, sText, """", vbBinaryCompare)
		If (iMidPos > Len("FORM_ID=""")) And (iEndPos > 0) Then
			lFormID = CLng(Mid(sText, iMidPos, (iEndPos - iMidPos)))
			iEndPos = InStr(iMidPos, sText, "/>", vbBinaryCompare)
			iEndPos = iEndPos + Len("/>")
			If Err.number = 0 Then
				sErrorDescription = "No se pudo obtener la información del formulario."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select FormFieldText, Answer, FieldTypeID, FormFieldSize From FormFieldsAnswers, FormFields Where (FormFieldsAnswers.FormID=FormFields.FormID) And (FormFieldsAnswers.FormFieldID=FormFields.FormFieldID) And (FormFields.FormID=" & lFormID & ") And (FormFieldName='" & sFormFieldName & "') And (FormFieldsAnswers.AnswerID = " & aFormComponent(N_ANSWER_ID_FORM) & ") Order By FormFields.FormFieldID", "FormComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					sFormAnswers = ""
					Do While Not oRecordset.EOF
						sAnswer = ""
						If Not oRecordset.EOF Then
							sAnswer = CStr(oRecordset.Fields("Answer").Value)
							Err.Clear
							Select Case CInt(oRecordset.Fields("FieldTypeID").Value)
								Case 0
									sAnswer = DisplayYesNo(CInt(sAnswer), False)
								Case 1
									sAnswer = DisplayDateFromSerialNumber(sAnswer, -1, -1, -1)
								Case 5
									If CInt(oRecordset.Fields("FormFieldSize").Value) > 100 Then
										sAnswer = "<BR />" & CleanStringForHTML(sAnswer)
									Else
										sAnswer = CleanStringForHTML(sAnswer)
									End If
								Case Else
									sAnswer = CleanStringForHTML(sAnswer)
							End Select
						End If
						sFormAnswers = sFormAnswers & "" & CleanStringForHTML(CStr(oRecordset.Fields("FormFieldText").Value)) & ":&nbsp;" & sAnswer & "<BR />"
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					sText = Left(sText, (iStartPos - Len("<"))) & sFormAnswers & Right(sText, (Len(sText) - iEndPos + Len(".")))
				End If
			End If
		End If
		iStartPos = InStr(1, sText, "<FORM ", vbBinaryCompare)
		If Err.number <> 0 Then Exit Do
	Loop

	iStartPos = InStr(1, sText, "<CURRENT_DATE ", vbBinaryCompare)
	Do While (iStartPos > 0)
		iMidPos = InStr(iStartPos, sText, "ADD=""", vbBinaryCompare) + Len("ADD=""")
		iEndPos = InStr(iMidPos, sText, """", vbBinaryCompare)
		If (iMidPos > Len("ADD=""")) And (iEndPos > 0) Then
			lDate = CLng(Mid(sText, iMidPos, (iEndPos - iMidPos)))
			lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) + lDate
			iEndPos = InStr(iMidPos, sText, "/>", vbBinaryCompare)
			iEndPos = iEndPos + Len("/>")
			sText = Left(sText, (iStartPos - Len("<"))) & DisplayDateFromSerialNumber(lDate, -1, -1, -1) & Right(sText, (Len(sText) - iEndPos + Len(".")))
		End If
		iStartPos = InStr(1, sText, "<CURRENT_DATE ", vbBinaryCompare)
		If Err.number <> 0 Then Exit Do
	Loop

	TransformXMLTags = lErrorNumber
	Set oRecordset = Nothing
	Err.Clear
End Function
%>