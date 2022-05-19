<!-- #include file="EmployeeComponentConstants.asp" -->
<!-- #include file="EmployeeDisplayFormsComponent.asp" -->
<!-- #include file="EmployeeDisplayFormsComponentB.asp" -->
<!-- #include file="EmployeeDisplayTablesComponent.asp" -->
<!-- #include file="EmployeeAddComponent.asp" -->
<!-- #include file="EmployeeGetComponent.asp" -->
<%
Function AddChild(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: Add a child of an employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddChild"
	Dim lErrorNumber
	Dim oRecordset

	sErrorDescription = "No se pudo agregar la información del empleado."'se ingresa un nuevo hijo
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select *  From EmployeesChildrenLKP Where (EmployeeID="& aEmployeeComponent(N_ID_EMPLOYEE) &") AND  ROWNUM=1 order by childID desc;", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Len(oRecordset.Fields("ChildID").Value) <> 0 Then
				aEmployeeComponent(N_ID_CHILD_EMPLOYEE) = CLng(oRecordset.Fields("ChildID").Value) + 1
			End If
		Else
			aEmployeeComponent(N_ID_CHILD_EMPLOYEE) = 1
		End If
		sErrorDescription = "No se pudo agregar la información del empleado."
		If B_UPPERCASE Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesChildrenLKP (EmployeeID, ChildID, ChildName, ChildLastName, ChildLastName2, ChildBirthDate, ChildEndDate, LevelID, RegistrationDate, UserID) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) &"," & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & ", '" & Replace(UCase(aEmployeeComponent(S_NAME_CHILD_EMPLOYEE)), "'", "´") & "', '" & Replace(UCase(aEmployeeComponent(S_LAST_NAME_CHILD_EMPLOYEE)), "'", "´") & "', '" &  Replace(UCase(aEmployeeComponent(S_LAST_NAME2_CHILD_EMPLOYEE)), "'", "´") & "', " & aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE) & ", 0, " & aEmployeeComponent(N_CHILD_LEVEL_ID_EMPLOYEE) & ",0,-1)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesChildrenLKP (EmployeeID, ChildID, ChildName, ChildLastName, ChildLastName2, ChildBirthDate, ChildEndDate, LevelID, RegistrationDate, UserID) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) &"," & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NAME_CHILD_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_LAST_NAME_CHILD_EMPLOYEE), "'", "´") & "', '" &  Replace(UCase(aEmployeeComponent(S_LAST_NAME2_CHILD_EMPLOYEE)), "'", "´") & "', " & aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE) & ", 0, " & aEmployeeComponent(N_CHILD_LEVEL_ID_EMPLOYEE) & ", 0,-1)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		End If
	End If

	AddChild = lErrorNumber
	Err.Clear

End Function

Function AuthorizeEmployeeDocument(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given employee's concept
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AuthorizeEmployeeDocument"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número de empleado para autorizar la solicitud de hoja única de servicio."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo autorizar la solicitud de hoja única de servicio del empleado."

		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesDocs Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (DocumentDate=" & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_DATE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				If StrComp(CStr(oRecordset.Fields("Authorized").Value), "-1", vbBinaryCompare) = 0 Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesDocs Set Authorized='" & aLoginComponent(N_USER_ID_LOGIN) & "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (DocumentDate=" & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_DATE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesDocs Set Authorized=Authorized+'," & aLoginComponent(N_USER_ID_LOGIN) & "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (DocumentDate=" & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_DATE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				End If
			Else
				lErrorNumber = -1
				sErrorDescription = "No existe solicitud de hoja única de servicio para autorizarla."
			End If
		Else
			lErrorNumber = -1
			sErrorDescription = "Error al obtener la solicitud de hoja única de servicio para autorizarla."
		End If
	End If

	AuthorizeEmployeeDocument = lErrorNumber
	Err.Clear
End Function

Function CalculateAmountDiferenceForAntiquityConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To cancel a employee concept
'Inputs:  oRequest, oADODBConnection, sAction
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CalculateAmountDiferenceForAntiquityConcept"
	Dim lErrorNumber
	Dim lCurrentSalary
	Dim lSuperiorAmount
	Dim bIsSuperior

	lErrorNumber = GetEmployeeSalary(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
	If lErrorNumber = 0 Then
		lCurrentSalary = aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE)
		lErrorNumber = GetEmployeeSuperiorPositionAmount(oRequest, oADODBConnection, aEmployeeComponent, lSuperiorAmount, bIsSuperior, sErrorDescription)
		If lErrorNumber = 0 Then
			If bIsSuperior Then
				aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = lSuperiorAmount - aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE)
			Else
				aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) - lSuperiorAmount
			End If
		End If
	End If

	CalculateAmountDiferenceForAntiquityConcept = lErrorNumber
	Err.Clear
End Function

Function CancelEmployeeAddSafeSeparation(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To cancel a employee concept
'Inputs:  oRequest, oADODBConnection, sAction
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CancelEmployeeAddSafeSeparation"
	Dim lErrorNumber
	Dim bComponentInitialized

	If (ExistAddSafeSeparationConcept(oADODBConnection, aEmployeeComponent, sErrorDescription)) Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set EndDate=" & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		sErrorDescription = "No se pudo cancelar el seguro de separación adicional."
	End If

	CancelEmployeeAddSafeSeparation = lErrorNumber
	Err.Clear
End Function

Function CancelEmployeeConcept(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To cancel a employee concept
'Inputs:  oRequest, oADODBConnection, sAction
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CancelEmployeeConcept"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 87) Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set EndDate=" & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		sErrorDescription = "No se pudo eliminar el registro seleccionado."
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set Active=2, ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		sErrorDescription = "No se pudo eliminar el registro seleccionado."
	End If

	CancelEmployeeConcept = lErrorNumber
	Err.Clear
End Function

Function CloseEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing concept for the employee in
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CloseEmployeeConcept"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim lConceptStartDate
	Dim lConceptEndDate
	Dim iActive
	Dim sDate
	Dim oRecordset

	sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If
	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del concepto para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del concepto del empleado."
		sQuery = "Select ConceptID, EndDate From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery , "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If oRecordset.EOF Then
			sErrorDescription = "El concepto o la fecha inicial de vigencia no son correctos"
			lErrorNumber = -1
		Else
			If (Len(oRequest("Cancel").Item) > 0) Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set Active=2, RegistrationDate=" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", ModifyDate=" & sDate & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")" , "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set EndDate=" & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", ModifyDate=" & sDate & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")" , "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				If (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 120) and (lErrorNumber = 0) Then
					sErrorDescription = "No se pudo cancelar el seguro adicional."	
					lErrorNumber = CancelEmployeeAddSafeSeparation(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
				End If
				lErrorNumber = GetEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				If aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 30000000 Then
					aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE)
				Else
					aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = AddDaysToSerialDate(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), 1)
				End If
				aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 0
				aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 2
				aEmployeeComponent(B_CANCEL_CONCEPT_FOR_EMPLOYEE) = True
				lErrorNumber = AddEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			End If
		End If
	End If

	CloseEmployeeConcept = lErrorNumber
	Err.Clear
End Function

Function DeleteChild(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a child of an employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DeleteChild"
	Dim lErrorNumber
	Dim oRecordset

	sErrorDescription = "Error al borrar la información de la base de datos"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesChildrenLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ChildID=" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

	DeleteChild = lErrorNumber
	Err.Clear
End Function

Function DropEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a concept for the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DropEmployeeConcept"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del concepto a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del concepto del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	DropEmployeeConcept = lErrorNumber
	Err.Clear
End Function

Function ExistAddSafeSeparationConcept(oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if an employee has a Credit
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ExistAddSafeSeparationConcept"
	Dim lErrorNumber
	Dim oRecordset

	sErrorDescription = "Error al verificar si el empleado tiene registrado seguro adicional de separación."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from EmployeesConceptsLKP Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID = 87) And ((EndDate = 30000000) Or (EndDate > " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ")) And (Active = 1)", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = CInt(oRecordset("ConceptID").Value)
			aEmployeeComponent(L_START_DATE_ID_EMPLOYEE) = CLng(oRecordset("StartDate").Value)
			ExistAddSafeSeparationConcept = True
		Else
			sErrorDescription = "El empleado no tiene registrado seguro adicional de separación."
			ExistAddSafeSeparationConcept = False
		End If
	Else
		sErrorDescription = "Error al verificar si el empleado tiene registrado seguro adicional de separación."
		ExistAddSafeSeparationConcept = False
	End If
	Err.Clear
End Function

Function ModifyChild(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify a child from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyChild"
	Dim lErrorNumber
	Dim oRecordset

	sErrorDescription = "Error al modificar la información de la base de datos"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChildrenLKP Set ChildName='" & Replace(UCase(aEmployeeComponent(S_NAME_CHILD_EMPLOYEE)), "'", "") & "', ChildLastName='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME_CHILD_EMPLOYEE)), "'", "") & "', ChildLastName2='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME2_CHILD_EMPLOYEE)), "'", "") & "', ChildBirthDate=" & aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE) & ", LevelID=" & aEmployeeComponent(N_CHILD_LEVEL_ID_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") and (ChildID=" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & ") ", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

	ModifyChild = lErrorNumber
	Err.Clear

End Function

Function ModifyEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing employee in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyEmployee"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sDate
	Dim lModifyDate
	Dim lRiskLevel
	Dim lRiskAmount
	Dim lExtraShift1
	Dim lExtraShift2

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aEmployeeComponent(B_CHECK_FOR_DUPLICATED_EMPLOYEE) Then
			lErrorNumber = CheckExistencyOfEmployee(aEmployeeComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aEmployeeComponent(B_IS_DUPLICATED_EMPLOYEE) And (StrComp(oRequest("Action").Item,"Jobs",vbBinaryCompare) <> 0) And (StrComp(oRequest("Modify").Item,"Modificar",vbBinaryCompare) <> 0) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un empleado con el número " & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				If Not CheckEmployeeInformationConsistency(aEmployeeComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					If Len(oRequest("ReasonID").Item) <> 0 Then
						If lReasonID = 57 Then
							lRiskLevel = oRequest("RiskLevel").Item
							lRiskAmount = CInt(oRequest("RiskLevel").Item) * 10
							lExtraShift1 = oRequest("StartHour3").Item
							lExtraShift2 = oRequest("EndHour3").Item
						End If
					End If
					If (StrComp(oRequest("Modify").Item , "Aplicar Titularidad", vbBinaryCompare) = 0) Then aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRequest("OwnerID").Item)
					sDate = Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000"))
					If Len(oRequest("ReasonID").Item) <> 0 Then
						If lReasonID <> 57 Then
							If aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 0 Then aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 30000000
						End If
					End If
					sErrorDescription = "No se pudo modificar la información del empleado."
					'If Len(aEmployeeComponent(N_START_HOUR_3_EMPLOYEE)) = 0 Then aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) = 0
					'If Len(aEmployeeComponent(N_END_HOUR_3_EMPLOYEE)) = 0 Then aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) = 0
					If B_UPPERCASE Then
						'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', EmployeeAccessKey='" & Replace(aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE), "'", "") & "', EmployeePassword='" & Replace(aEmployeeComponent(S_PASSWORD_EMPLOYEE), "'", "") & "', EmployeeName='" & Replace(UCase(aEmployeeComponent(S_NAME_EMPLOYEE)), "'", "´") & "', EmployeeLastName='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME_EMPLOYEE)), "'", "´") & "', EmployeeLastName2='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE)), "'", "´") & "', CompanyID=" & aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) & ", JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ", ServiceID=" & aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", ClassificationID=" & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & ", GroupGradeLevelID=" & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & ", IntegrationID=" & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & ", JourneyID=" & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & ", ShiftID=" & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", StartHour1=" & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & ", EndHour1=" & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & ", StartHour2=" & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & ", EndHour2=" & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & ", StartHour3=" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", EndHour3=" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & ", WorkingHours=" & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & ", LevelID=" & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) & ", EmployeeEmail='" & Replace(aEmployeeComponent(S_EMAIL_EMPLOYEE), "'", "") & "', SocialSecurityNumber='" & Replace(aEmployeeComponent(S_SSN_EMPLOYEE), "'", "") & "', BirthYear=" & CInt(Left(sDate, Len("0000"))) & ", BirthMonth=" & CInt(Mid(sDate, Len("00000"), Len("00"))) & ", BirthDay=" & CInt(Mid(sDate, Len("0000000"), Len("00"))) & ", BirthDate=" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) & ", StartDate=" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & ", StartDate2=" & aEmployeeComponent(N_START_DATE2_EMPLOYEE) & ", CountryID=" & aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) & ", RFC='" & Replace(UCase(aEmployeeComponent(S_RFC_EMPLOYEE)), "'", "") & "', CURP='" & Replace(UCase(aEmployeeComponent(S_CURP_EMPLOYEE)), "'", "") & "', GenderID=" & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & ", MaritalStatusID=" & aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', EmployeeAccessKey='" & Replace(aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE), "'", "") & "', EmployeePassword='" & Replace(aEmployeeComponent(S_PASSWORD_EMPLOYEE), "'", "") & "', EmployeeName='" & Replace(UCase(aEmployeeComponent(S_NAME_EMPLOYEE)), "'", "´") & "', EmployeeLastName='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME_EMPLOYEE)), "'", "´") & "', EmployeeLastName2='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE)), "'", "´") & "', CompanyID=" & aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) & ", JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ", ServiceID=" & aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", ClassificationID=" & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & ", GroupGradeLevelID=" & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & ", IntegrationID=" & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & ", JourneyID=" & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & ", ShiftID=" & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", StartHour1=" & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & ", EndHour1=" & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & ", StartHour2=" & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & ", EndHour2=" & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & ", WorkingHours=" & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & ", LevelID=" & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) & ", EmployeeEmail='" & Replace(aEmployeeComponent(S_EMAIL_EMPLOYEE), "'", "") & "', SocialSecurityNumber='" & Replace(aEmployeeComponent(S_SSN_EMPLOYEE), "'", "") & "', BirthYear=" & CInt(Left(sDate, Len("0000"))) & ", BirthMonth=" & CInt(Mid(sDate, Len("00000"), Len("00"))) & ", BirthDay=" & CInt(Mid(sDate, Len("0000000"), Len("00"))) & ", BirthDate=" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) & ", StartDate=" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & ", StartDate2=" & aEmployeeComponent(N_START_DATE2_EMPLOYEE) & ", CountryID=" & aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) & ", RFC='" & Replace(UCase(aEmployeeComponent(S_RFC_EMPLOYEE)), "'", "") & "', CURP='" & Replace(UCase(aEmployeeComponent(S_CURP_EMPLOYEE)), "'", "") & "', GenderID=" & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & ", MaritalStatusID=" & aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Else
						'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', EmployeeAccessKey='" & Replace(aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE), "'", "") & "', EmployeePassword='" & Replace(aEmployeeComponent(S_PASSWORD_EMPLOYEE), "'", "") & "', EmployeeName='" & Replace(aEmployeeComponent(S_NAME_EMPLOYEE), "'", "´") & "', EmployeeLastName='" & Replace(aEmployeeComponent(S_LAST_NAME_EMPLOYEE), "'", "´") & "', EmployeeLastName2='" & Replace(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE), "'", "´") & "', CompanyID=" & aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) & ", JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ", ServiceID=" & aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", ClassificationID=" & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & ", GroupGradeLevelID=" & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & ", IntegrationID=" & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & ", JourneyID=" & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & ", ShiftID=" & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", StartHour1=" & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & ", EndHour1=" & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & ", StartHour2=" & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & ", EndHour2=" & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & ", StartHour3=" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", EndHour3=" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & ", WorkingHours=" & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & ", LevelID=" & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) & ", EmployeeEmail='" & Replace(aEmployeeComponent(S_EMAIL_EMPLOYEE), "'", "") & "', SocialSecurityNumber='" & Replace(aEmployeeComponent(S_SSN_EMPLOYEE), "'", "") & "', BirthYear=" & CInt(Left(sDate, Len("0000"))) & ", BirthMonth=" & CInt(Mid(sDate, Len("00000"), Len("00"))) & ", BirthDay=" & CInt(Mid(sDate, Len("0000000"), Len("00"))) & ", BirthDate=" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) & ", StartDate=" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & ", StartDate2=" & aEmployeeComponent(N_START_DATE2_EMPLOYEE) & ", CountryID=" & aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) & ", RFC='" & Replace(aEmployeeComponent(S_RFC_EMPLOYEE), "'", "") & "', CURP='" & Replace(aEmployeeComponent(S_CURP_EMPLOYEE), "'", "") & "', GenderID=" & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & ", MaritalStatusID=" & aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', EmployeeAccessKey='" & Replace(aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE), "'", "") & "', EmployeePassword='" & Replace(aEmployeeComponent(S_PASSWORD_EMPLOYEE), "'", "") & "', EmployeeName='" & Replace(aEmployeeComponent(S_NAME_EMPLOYEE), "'", "´") & "', EmployeeLastName='" & Replace(aEmployeeComponent(S_LAST_NAME_EMPLOYEE), "'", "´") & "', EmployeeLastName2='" & Replace(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE), "'", "´") & "', CompanyID=" & aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) & ", JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ", ServiceID=" & aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", ClassificationID=" & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & ", GroupGradeLevelID=" & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & ", IntegrationID=" & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & ", JourneyID=" & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & ", ShiftID=" & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", StartHour1=" & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & ", EndHour1=" & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & ", StartHour2=" & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & ", EndHour2=" & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & ", WorkingHours=" & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & ", LevelID=" & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) & ", EmployeeEmail='" & Replace(aEmployeeComponent(S_EMAIL_EMPLOYEE), "'", "") & "', SocialSecurityNumber='" & Replace(aEmployeeComponent(S_SSN_EMPLOYEE), "'", "") & "', BirthYear=" & CInt(Left(sDate, Len("0000"))) & ", BirthMonth=" & CInt(Mid(sDate, Len("00000"), Len("00"))) & ", BirthDay=" & CInt(Mid(sDate, Len("0000000"), Len("00"))) & ", BirthDate=" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) & ", StartDate=" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & ", StartDate2=" & aEmployeeComponent(N_START_DATE2_EMPLOYEE) & ", CountryID=" & aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) & ", RFC='" & Replace(aEmployeeComponent(S_RFC_EMPLOYEE), "'", "") & "', CURP='" & Replace(aEmployeeComponent(S_CURP_EMPLOYEE), "'", "") & "', GenderID=" & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & ", MaritalStatusID=" & aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudo modificar la información del empleado."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesExtraInfo Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudo modificar la información del empleado."
							If B_UPPERCASE Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesExtraInfo (EmployeeID, EmployeeAddress, EmployeeCity, EmployeeZipCode, StateID, CountryID, EmployeePhone, OfficePhone, OfficeExt, DocumentNumber1, DocumentNumber2, DocumentNumber3, EmployeeActivityID,BirthPlace,Languages,BloodType,CellPhone, DeathBeneficiary, DeathBeneficiary2) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", '" & Replace(UCase(aEmployeeComponent(S_ADDRESS_EMPLOYEE)), "'", "´") & "', '" & Replace(UCase(aEmployeeComponent(S_CITY_EMPLOYEE)), "'", "´") & "', '" & Replace(aEmployeeComponent(S_ZIP_CODE_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_ADDRESS_COUNTRY_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_EMPLOYEE_PHONE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_OFFICE_PHONE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_EXT_OFFICE_EMPLOYEE), "'", "") & "', '" & Replace(UCase(aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE)), "'", "") & "', '" & Replace(UCase(aEmployeeComponent(S_DOCUMENT_NUMBER_2_EMPLOYEE)), "'", "") & "', '" & Replace(UCase(aEmployeeComponent(S_DOCUMENT_NUMBER_3_EMPLOYEE)), "'", "") & "', " & aEmployeeComponent(N_ACTIVITY_ID_EMPLOYEE) & ",'" & Replace(UCase(aEmployeeComponent(S_EMPLOYEE_BIRTHPLACE)), "'", "") &"','" & Replace(UCase(aEmployeeComponent(S_EMPLOYEE_LANGUAGES)), "'", "") & "','"& Replace(UCase(aEmployeeComponent(S_EMPLOYEE_BLOODTYPE)), "'", "") &"','"& aEmployeeComponent(S_EMPLOYEE_CELLPHONE) &"','" & Replace(UCase(aEmployeeComponent(S_EMPLOYEE_DEATH_BENEFICIARY)), "'", "") &"','" & Replace(UCase(aEmployeeComponent(S_EMPLOYEE_DEATH_BENEFICIARY2)), "'", "") &"')", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
							Else
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesExtraInfo (EmployeeID, EmployeeAddress, EmployeeCity, EmployeeZipCode, StateID, CountryID, EmployeePhone, OfficePhone, OfficeExt, DocumentNumber1, DocumentNumber2, DocumentNumber3, EmployeeActivityID,BirthPlace,Languages,BloodType,CellPhone, DeathBeneficiary, DeathBeneficiary2) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_ADDRESS_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_CITY_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_ZIP_CODE_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_ADDRESS_COUNTRY_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_EMPLOYEE_PHONE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_OFFICE_PHONE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_EXT_OFFICE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_DOCUMENT_NUMBER_2_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_DOCUMENT_NUMBER_3_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_ACTIVITY_ID_EMPLOYEE) & ",'" & Replace(aEmployeeComponent(S_EMPLOYEE_BIRTHPLACE), "'", "") &"','" & Replace(aEmployeeComponent(S_EMPLOYEE_LANGUAGES), "'", "") & "','"& Replace(aEmployeeComponent(S_EMPLOYEE_BLOODTYPE), "'", "") &"','"& aEmployeeComponent(S_EMPLOYEE_CELLPHONE) &"'," & Replace(UCase(aEmployeeComponent(S_EMPLOYEE_DEATH_BENEFICIARY)), "'", "") &",'" & Replace(UCase(aEmployeeComponent(S_EMPLOYEE_DEATH_BENEFICIARY2)), "'", "") &"')", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
							End If
						End If
					End If
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudo modificar la información del empleado."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesSchoolLevelsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudo modificar la información del empleado."
							If B_UPPERCASE Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesSchoolLevelsLKP (EmployeeID, RecordID, SchoolName, SchoolarShipID, SchoolarShipStatusID, StartDate, EndDate, RegisterDate, UserID, StatusID, Specialism) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", 1, '" & Replace(UCase(aEmployeeComponent(S_EMPLOYEE_SCHOOLNAME)), "'", "") & "', " & aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_ID) &  ", 1, " & aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_DATE) & ", " & aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_DATE_END)& ", 0, -1, 1, '" & Replace(UCase(aEmployeeComponent(S_EMPLOYEE_SPECIALISM)), "'", "")& "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
							Else
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesSchoolLevelsLKP (EmployeeID, RecordID, SchoolName, SchoolarShipID, SchoolarShipStatusID, StartDate, EndDate, RegisterDate, UserID, StatusID, Specialism) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", 1, '" & Replace(aEmployeeComponent(S_EMPLOYEE_SCHOOLNAME ), "'", "") & "', " & aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_ID) &  ", 1, " & aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_DATE) & ", " & aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_DATE_END)& ", 0, -1, 1, '" & Replace(aEmployeeComponent(S_EMPLOYEE_SPECIALISM), "'", "") & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
							End If
						End If
					End If
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudo modificar la información del empleado."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, ModifyDate From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudo modificar la información del empleado."
							If oRecordset.EOF Then
								If Len(oRequest("ReasonID").Item) <> 0 Then
									If lReasonID = 57 Then
										lModifyDate = oRequest("EmployeeYear").Item & oRequest("EmployeeMonth").Item & oRequest("EmployeeDay").Item
									Else
										lModifyDate = Left(GetSerialNumberForDate(""), Len("00000000"))
									End If
								Else
									lModifyDate = Left(GetSerialNumberForDate(""), Len("00000000"))
								End If
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & lModifyDate & ", " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_ZONE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & ", " & aEmployeeComponent(N_AREA_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", 0, '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								If lErrorNumber = 0 Then lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							Else
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', CompanyID=" & aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) & ", JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ", ServiceID=" & aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) & ", ZoneID=" & aEmployeeComponent(N_ZONE_ID_EMPLOYEE) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", ClassificationID=" & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & ", GroupGradeLevelID=" & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & ", IntegrationID=" & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & ", JourneyID=" & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & ", ShiftID=" & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", WorkingHours=" & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & ", AreaID=" & aEmployeeComponent(N_AREA_ID_EMPLOYEE) & ", PositionID=" & aEmployeeComponent(N_POSITION_ID_EMPLOYEE) & ", LevelID=" & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) & ", RiskLevel=" & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", PayrollDate=" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", bProcessed=0, Comments='" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								If lErrorNumber = 0 Then lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & CStr(oRecordset.Fields("ModifyDate").Value) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							End If
							oRecordset.Close
						End If
					End If
					If lErrorNumber = 0 Then
						If lRiskLevel <> 0 Then
							sErrorDescription = "No se pudo obtener la información del riesgo profesional que tiene el empleado."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesRisksLKP Where EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE), "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo eliminar la información del riesgo al empleado."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesRisksLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo agregar la información del riesgo al empleado."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesRisksLKP (EmployeeID, RiskLevel) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
									If lRiskLevel = 1 Then
										lRiskAmount = 10
									ElseIf lRiskLevel = 2 Then
										lRiskAmount = 20
									End If
									'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesConceptsLKP (EmployeeID, ConceptID, StartDate, EndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, RegistrationDate, ModifyDate, StartUserID, EndUserID, UploadedFileName, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", 4, " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", " & lRiskAmount & ",0,2,3,0,0,0,0,'',1,401,1," & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", -1, '" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "'") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								End If
							End If
							If lErrorNumber = 0 Then
								aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 4
								aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
								aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)
								aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) = oRequest("RiskLevel").Item
								aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CInt(aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE)) * 10
								aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) = 0
								aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
								aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 3
								aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) = 0
								aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) = 1
								aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) = 0
								aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) = 1
								aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = -1
								aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) = 1
								aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 1
								aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = ""
								aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = "Trámite realizado a través de FM1"
								sErrorDescription = "No se pudo agregar el concepto de riesgo al empleado."
								lErrorNumber = ModifyEmployeeConceptSp(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							End If
						End If
						If lErrorNumber = 0 Then
							If (lExtraShift1 > 0) And (lExtraShift2 > 0) Then
								sErrorDescription = "No se pudo agregar la percepción adicional del empleado."
								aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
								aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)
								aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = 3/6.5*100
								aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) = 0
								aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
								aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 3
								aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) = oRequest("StartHour3").Item
								aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) = oRequest("EndHour3").Item
								aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) = 0
								aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) = 1
								aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) = 0
								aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) = 1
								aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1,5"
								aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) = 1
								aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 1
								aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = ""
								aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = "Trámite realizado a través de FM1"
								If aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) = 1 Then
									aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 7
									If lErrorNumber = 0 Then
										lErrorNumber = ModifyEmployeeConceptSp(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudo agregar el turno opcional al empleado."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set StartHour3=" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", EndHour3=" & oRequest("EndHour3").Item & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									End If
								ElseIf (aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) = 2) And (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 0 Or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 2 Or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 3 Or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 4) Then
									aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 8
									If lErrorNumber = 0 Then
										lErrorNumber = ModifyEmployeeConceptSp(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudo agregar la percepción adicional al empleado."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set StartHour3=" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", EndHour3=" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	ModifyEmployee = lErrorNumber
	Err.Clear
End Function

Function ModifyEmployeeForSuspension(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing employee in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyEmployeeForSuspension"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sDate

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sDate = Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000"))
		If CLng(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)) <= CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) Then
			sErrorDescription = "No se pudo modificar el estatus del empleado."
			If B_UPPERCASE Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', EmployeeAccessKey='" & Replace(aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE), "'", "") & "', EmployeePassword='" & Replace(aEmployeeComponent(S_PASSWORD_EMPLOYEE), "'", "") & "', EmployeeName='" & Replace(UCase(aEmployeeComponent(S_NAME_EMPLOYEE)), "'", "´") & "', EmployeeLastName='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME_EMPLOYEE)), "'", "´") & "', EmployeeLastName2='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE)), "'", "´") & "', CompanyID=" & aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) & ", JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ", ServiceID=" & aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", ClassificationID=" & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & ", GroupGradeLevelID=" & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & ", IntegrationID=" & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & ", JourneyID=" & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & ", ShiftID=" & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", StartHour1=" & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & ", EndHour1=" & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & ", StartHour2=" & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & ", EndHour2=" & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & ", StartHour3=" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", EndHour3=" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & ", WorkingHours=" & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & ", LevelID=" & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) & ", EmployeeEmail='" & Replace(aEmployeeComponent(S_EMAIL_EMPLOYEE), "'", "") & "', SocialSecurityNumber='" & Replace(aEmployeeComponent(S_SSN_EMPLOYEE), "'", "") & "', BirthYear=" & CInt(Left(sDate, Len("0000"))) & ", BirthMonth=" & CInt(Mid(sDate, Len("00000"), Len("00"))) & ", BirthDay=" & CInt(Mid(sDate, Len("0000000"), Len("00"))) & ", BirthDate=" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) & ", StartDate=" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & ", StartDate2=" & aEmployeeComponent(N_START_DATE2_EMPLOYEE) & ", CountryID=" & aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) & ", RFC='" & Replace(UCase(aEmployeeComponent(S_RFC_EMPLOYEE)), "'", "") & "', CURP='" & Replace(UCase(aEmployeeComponent(S_CURP_EMPLOYEE)), "'", "") & "', GenderID=" & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & ", MaritalStatusID=" & aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', EmployeeAccessKey='" & Replace(aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE), "'", "") & "', EmployeePassword='" & Replace(aEmployeeComponent(S_PASSWORD_EMPLOYEE), "'", "") & "', EmployeeName='" & Replace(aEmployeeComponent(S_NAME_EMPLOYEE), "'", "´") & "', EmployeeLastName='" & Replace(aEmployeeComponent(S_LAST_NAME_EMPLOYEE), "'", "´") & "', EmployeeLastName2='" & Replace(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE), "'", "´") & "', CompanyID=" & aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) & ", JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ", ServiceID=" & aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", ClassificationID=" & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & ", GroupGradeLevelID=" & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & ", IntegrationID=" & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & ", JourneyID=" & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & ", ShiftID=" & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", StartHour1=" & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & ", EndHour1=" & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & ", StartHour2=" & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & ", EndHour2=" & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & ", StartHour3=" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", EndHour3=" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & ", WorkingHours=" & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & ", LevelID=" & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) & ", EmployeeEmail='" & Replace(aEmployeeComponent(S_EMAIL_EMPLOYEE), "'", "") & "', SocialSecurityNumber='" & Replace(aEmployeeComponent(S_SSN_EMPLOYEE), "'", "") & "', BirthYear=" & CInt(Left(sDate, Len("0000"))) & ", BirthMonth=" & CInt(Mid(sDate, Len("00000"), Len("00"))) & ", BirthDay=" & CInt(Mid(sDate, Len("0000000"), Len("00"))) & ", BirthDate=" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) & ", StartDate=" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & ", StartDate2=" & aEmployeeComponent(N_START_DATE2_EMPLOYEE) & ", CountryID=" & aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) & ", RFC='" & Replace(aEmployeeComponent(S_RFC_EMPLOYEE), "'", "") & "', CURP='" & Replace(aEmployeeComponent(S_CURP_EMPLOYEE), "'", "") & "', GenderID=" & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & ", MaritalStatusID=" & aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo modificar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, EmployeeDate From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") Order By EmployeeDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo modificar la información del empleado."
				If oRecordset.EOF Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_ZONE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & ", " & aEmployeeComponent(N_AREA_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", 0, '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					If CLng(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)) <= CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) Then
						If lErrorNumber = 0 Then lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EndDate=" & AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), -1) & " Where (EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & ") And (EmployeeDate=" & CLng(oRecordset.Fields("EmployeeDate").Value) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					If lErrorNumber = 0 Then lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_ZONE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & ", " & aEmployeeComponent(N_AREA_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", 0, '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					If CLng(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)) <= CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) Then
						If lErrorNumber = 0 Then lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
				oRecordset.Close
			End If
		End If
	End If

	Set oRecordset = Nothing
	ModifyEmployee = lErrorNumber
	Err.Clear
End Function

Function ModifyEmployeeAbsences(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the beneficiary information for the employee in
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyEmployeeAbsences"
	Dim oRecordset
	Dim lErrorNumber
	Dim sField
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o de la incidencia para modificar la información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo modificar la información de la incidencia del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set AbsenceHours=" & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", Reasons='" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "'" & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (AbsenceID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (OcurredDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	ModifyEmployeeAbsences = lErrorNumber
	Err.Clear
End Function

Function ModifyEmployeeBankAccount(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the bank account information for the employee in
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************

	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyEmployeeBankAccount"
	Dim oRecordset
	Dim lErrorNumber
	Dim sField
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador de la cuenta para modificar la información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo modificar la información de la cuenta bancaria del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update BankAccounts set BankID=" & aEmployeeComponent(N_BANK_ID_EMPLOYEE) & ", AccountNumber='" & Replace(aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE), "'", "´") & "', StartDate="& aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", EndDate="& aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & " Where AccountID=" & aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE), "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	ModifyEmployeeBankAccount = lErrorNumber
	Err.Clear
End Function

Function ModifyEmployeeBeneficiary(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the beneficiary information for the employee in
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyEmployeeBeneficiary"
	Dim oRecordset
	Dim lErrorNumber
	Dim sField
	Dim sQuery
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del beneficiario para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Not CheckEmployeeBeneficiaryInformationConsistency(aEmployeeComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sErrorDescription = "No se pudo modificar la información del beneficiario del empleado."
			If aEmployeeComponent(N_END_DATE_BENEFICIARY_EMPLOYEE) > 0 Then
				sField = "StartUserID"
			Else
				sField = "EndUserID"
			End If
			sQuery = "Update EmployeesBeneficiariesLKP" & _
					 " Set EndDate=" & aEmployeeComponent(N_END_DATE_BENEFICIARY_EMPLOYEE) & _
					 ", BeneficiaryNumber=" & aEmployeeComponent(S_NUMBER_BENEFICIARY_EMPLOYEE) & _
					 ", BeneficiaryName='" & Replace(aEmployeeComponent(S_NAME_BENEFICIARY_EMPLOYEE), "'", "´") & _
					 "', BeneficiaryLastName='" & Replace(aEmployeeComponent(S_LAST_NAME_BENEFICIARY_EMPLOYEE), "'", "´") & _
					 "', BeneficiaryLastName2='" & Replace(aEmployeeComponent(S_LAST_NAME2_BENEFICIARY_EMPLOYEE), "'", "´") & _
					 "', BeneficiaryBirthDate=" & aEmployeeComponent(N_BIRTH_DATE_BENEFICIARY_EMPLOYEE) & _
					 ", ConceptAmount=" & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & _
					 ", AlimonyTypeID=" & aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE) & _
					 ", PaymentCenterID=" & aEmployeeComponent(N_PAYMENT_CENTER_ID_BENEFICIARY_EMPLOYEE) & _
					 ", " & sField & "=" & aLoginComponent(N_USER_ID_LOGIN) & _
					 ", Comments='" & Replace(aEmployeeComponent(S_COMMENTS_BENEFICIARY_EMPLOYEE), "'", "´") & _
					 "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & _
					 ") And (BeneficiaryID=" & aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) & _
					 ") And (StartDate=" & aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE) & ")"

			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		End If
	End If

	ModifyEmployeeBeneficiary = lErrorNumber
	Err.Clear
End Function

Function ModifyEmployeeChild(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the child information for the employee in
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyEmployeeChild"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_ID_CHILD_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o de su hijo(a) para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Not CheckEmployeeChildInformationConsistency(aEmployeeComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sErrorDescription = "No se pudo modificar la información del hijo(a) del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChildrenLKP Set ChildName='" & Replace(aEmployeeComponent(S_NAME_CHILD_EMPLOYEE), "'", "´") & "', ChildLastName='" & Replace(aEmployeeComponent(S_LAST_NAME_CHILD_EMPLOYEE), "'", "´") & "', ChildLastName2='" & Replace(aEmployeeComponent(S_LAST_NAME2_CHILD_EMPLOYEE), "'", "´") & "', ChildBirthDate=" & aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE) & ", ChildEndDate=" & aEmployeeComponent(N_END_DATE_CHILD_EMPLOYEE) & ", LevelID=" & aEmployeeComponent(N_CHILD_LEVEL_ID_EMPLOYEE) & ", RegistrationDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ChildID=" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		End If
	End If

	ModifyEmployeeChild = lErrorNumber
	Err.Clear
End Function

Function ModifyEmployeeConcepts(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the employee concept information for the employee in
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyEmployeeConcepts"
	Dim oRecordset
	Dim lErrorNumber
	Dim sField
	Dim bComponentInitialized
	Dim sQuery
	Dim lStartDate
	Dim lEndDate
	Dim bExists
	Dim lConceptID

	bExists = False
	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado ni la clave del concepto para modificar la información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		Select Case aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE)
			Case 4
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1,89"
				aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 3
			Case 5, 50
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1"
			Case 7, 8
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1,5"
				aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = 3/6.5*100
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
			Case 93
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1,4,38"
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
				aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 3
			Case 120
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1,3"
				aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 3
			Case 146
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1,4,5,6,7,8,47,89,13"
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
				aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 3
			Case Else
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 1
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1"
		End Select
		If (CInt(oRequest("ReasonID").Item) = EMPLOYEES_FOR_RISK) Or (CInt(oRequest("ReasonID").Item) =  EMPLOYEES_ADDITIONALSHIFT) _
				Or (CInt(oRequest("ReasonID").Item) = EMPLOYEES_CONCEPT_08) Then
			If Len(oRequest("ConceptID").Item) > 0 Then
				lConceptID = oRequest("ConceptID").Item
			Else
				If CInt(oRequest("ReasonID").Item) = EMPLOYEES_FOR_RISK Then lConceptID = 4
				If CInt(oRequest("ReasonID").Item) = EMPLOYEES_ADDITIONALSHIFT Then lConceptID = 7
				If CInt(oRequest("ReasonID").Item) = EMPLOYEES_CONCEPT_08 Then lConceptID = 8
			End If
			sQuery = "Select StartDate, EndDate, Active From EmployeesConceptsLKP Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID = " & lConceptID & ") Order By EndDate Desc"
			If Len(oRequest("ConceptStartYear").Item) > 0 Then
				lStartDate = CLng(oRequest("ConceptStartYear").Item & oRequest("ConceptStartMonth").Item & oRequest("ConceptStartDay").Item)
				lEndDate = CLng(oRequest("ConceptEndYear").Item & oRequest("ConceptEndMonth").Item & oRequest("ConceptEndDay").Item)
			Else
				lStartDate = aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)
				lEndDate = aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE)
			End If
			If lEndDate = 0 Then lEndDate = 30000000
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					Do While Not oRecordset.EOF
						If CInt(oRecordset.Fields("Active").Value) = 1 Then
							If (CLng(oRecordset.Fields("EndDate").Value) = 30000000) Or ((CLng(oRecordset.Fields("StartDate").Value) < lStartDate) And (CLng(oRecordset.Fields("EndDate").Value) > lEndDate)) Or _
								((CLng(oRecordset.Fields("StartDate").Value) > lStartDate) And (CLng(oRecordset.Fields("StartDate").Value) < lEndDate)) Or _
								((CLng(oRecordset.Fields("StartDate").Value) < lStartDate) And (CLng(oRecordset.Fields("StartDate").Value) > lEndDate)) Or _
								((CLng(oRecordset.Fields("EndDate").Value) > lStartDate) And (CLng(oRecordset.Fields("EndDate").Value) < lEndDate)) Then
								lErrorNumber = -1
								sErrorDescription = "El empleado indicado ya tiene registrado este concepto"
								Exit Do
							End If
						End If
						If CLng(oRecordset.Fields("StartDate").Value) = lStartDate Then
							If InStr(1, ",4,7,8,", "," & CInt(oRequest("ConceptID").Item) & ",", vbBinaryCompare) <> 0 Then
								If CLng(oRecordset.Fields("StartDate").Value) = CLng(oRecordset.Fields("EndDate").Value) Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set EndDate = " & lEndDate & " Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " And ConceptID = " & oRequest("ConceptID").Item & " And StartDate = " & lStartDate & " And EndDate = " & lStartDate, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesConceptsLKP Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " And ConceptID = " & oRequest("ConceptID").Item & " And StartDate = " & AddDaysToSerialDate(lStartDate,1) & " And EndDate = 0", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								End If
							End If
							bExists = True
						End If
						oRecordset.MoveNext
					Loop
				Else
				End If
			End If
			If lErrorNumber = 0 Then
				lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					If (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 7) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 8) Then
						aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = 3/6.5*100
					Else
						If Len(oRequest("ConceptAmount").Item) <> 0 Then
							aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(oRequest("ConceptAmount").Item)
						Else
							aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(oRequest("RiskLevel").Item) * 10
						End If
					End If
					If bExists = False Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesConceptsLKP (EmployeeID, ConceptID, StartDate, EndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, RegistrationDate, ModifyDate, StartUserID, EndUserID, UploadedFileName, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ", " & lStartDate & ", " & lEndDate & ", " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) & ", '" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "', " & aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) & ", 1, " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", '" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "'") & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						If aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 4 Then
							If lErrorNumber = 0 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesRisksLKP (EmployeeID, RiskLevel) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							End If
						End If
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set Active = 1, ConceptAmount = " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & " Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " And ConceptID = " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & " And StartDate = " & lStartDate & " And EndDate = " & lEndDate, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
					If lErrorNumber = 0 Then
						If (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 7) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 8) Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set StartHour3=" & oRequest("StartHour3").Item & ",EndHour3=" & oRequest("EndHour3").Item & " Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE), "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
					End If
					If lErrorNumber <> 0 Then
						sErrorDescription = "No se pudo actualizar la información del empleado"
					End If
				Else
					sErrorDescription = "No se pudo obtener la información del empleado"
				End If
			End If
		Else
			sErrorDescription = "No se pudo modificar la información de la cuenta bancaria del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set EndDate=" & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", RegistrationDate=" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", ConceptAmount=" & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)		
		End If
	End If
	ModifyEmployeeConcepts = lErrorNumber
	Err.Clear
End Function

Function ModifyEmployeeConceptSp(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing concept for the employee in
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyEmployeeConcept"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim lConceptStartDate
	Dim lConceptEndDate
	Dim iActive
	Dim sDate

	sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del concepto para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		'If VerifyDatesForPayroll(aEmployeeComponent, sErrorDescription) Then
		'	lErrorNumber = -1
		'End If
		Select Case aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE)
			Case 4
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1,89"
				aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 3
			Case 5, 50
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1"
			Case 7, 8
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1,5"
				aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = 3/6.5*100
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
			Case 13
				aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 1
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = -1
			Case 93
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1,4,38"
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
				aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 3
			Case 120
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1,3"
				aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 3
			Case 146
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1,4,5,6,7,8,47,89,13"
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
				aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 3
			Case Else
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 1
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1"
		End Select
		Select Case aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE)
			Case 22, -68, 24, 26, 45, 46, 94, 76 'EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE,EMPLOYEES_CONCEPT_C3, EMPLOYEES_CONCEPT_C3, EMPLOYEES_MONTHAWARD, EMPLOYEES_FONAC_ADJUSTMENT
				aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)
		End Select
		If aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 0 Then
			aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 30000000
		End If
		If (Len(oRequest("Authorization").Item) = 0) Then
			Select Case aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE)
				Case 24
					If StrComp(oRequest("ModifyConcept").Item, "1", vbBinaryCompare) <> 0 Then
						If Not VerifyAnualDiferenceOfEmployeesConcept(oADODBConnection, aEmployeeComponent, True, sErrorDescription) Then
							lErrorNumber = -1
						End If
					End If
				Case 22, 94
					If StrComp(oRequest("ModifyConcept").Item, "1", vbBinaryCompare) <> 0 Then
						If Not VerifyAnualDiferenceOfEmployeesConcept(oADODBConnection, aEmployeeComponent, False, sErrorDescription) Then
							lErrorNumber = -1
						End If
					End If
				Case 87
					sErrorDescription = "No se pudo obtener la información del concepto de pago del empleado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=120) And (EndDate=30000000) And (Active = 1)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If ((lErrorNumber <> 0) And (Not oRecordset.EOF)) Then
						sErrorDescription = "Para registrar el seguro adicional, debe de estar registrado el concepto SI."
						lErrorNumber = -1						
					ElseIf ((aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) > 100) And (aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2)) Then
						sErrorDescription = "El seguro adicional no debe de ser mayor a 100 %."
						lErrorNumber = -1											
					End If
				Case 93
					If Not IsHoliday(oADODBConnection, aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), sErrorDescription) Then
						sErrorDescription = "El concepto solo se puede registrar en día festivo."
						lErrorNumber = -1
					End If
				Case 120
					If aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) < aEmployeeComponent(N_START_DATE_EMPLOYEE) Then
						sErrorDescription = "La fecha de inicio del seguro de separación individual no puedo ser menor a la fecha de ingreso del empleado al Instituto."
						lErrorNumber = -1
					End If
					If lErrorNumber = 0 Then
						If (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 2) And (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 4) And (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 5) And (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 10) Then
							sErrorDescription = "El porcentaje por seguro de separación solo puede ser 2, 4, 5, 10."
							lErrorNumber = -1
						End If
					End If
			End Select
		End If
		If lErrorNumber = 0 Then
			If Not CheckEmployeeConceptInformationConsistency(aEmployeeComponent, sErrorDescription) Then
				lErrorNumber = -1
			Else
				If aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0 Then
					sErrorDescription = "No se pudo obtener la información del concepto del empleado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, StartDate, EndDate, Active From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (Active <> 2) Order By StartDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							lConceptStartDate = CLng(oRecordset.Fields("StartDate").Value)
							lConceptEndDate = CLng(oRecordset.Fields("EndDate").Value)
							iActive = CInt(oRecordset.Fields("Active").Value)
							If iActive = 1 Then
								If CLng(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)) = lConceptStartDate Then
									sErrorDescription = "No se pudo modificar la información del concepto del empleado."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set EndDate=" & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", Comments='" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								ElseIf aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) < lConceptStartDate Then
									sErrorDescription = "No se puede agregar un concepto con fecha menor a la registrada."
									lErrorNumber = -1
								ElseIf (lConceptEndDate <> 30000000) And (aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) < lConceptEndDate) Then
									sErrorDescription = "No se puede insertar el nuevo concepto del empleado puesto que se traslapa con otro ya registrado."
									lErrorNumber = -1
								Else
									sErrorDescription = "No se pudo insertar el nuevo concepto del empleado."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesConceptsLKP (EmployeeID, ConceptID, StartDate, EndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, RegistrationDate, ModifyDate, StartUserID, EndUserID, UploadedFileName, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) & ", '" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "', " & aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", '" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								End If
							ElseIf iActive = 0 Then
								If CLng(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)) = lConceptStartDate Then
									sErrorDescription = "No se pudo modificar la información del concepto del empleado."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set ConceptAmount=" & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", EndDate=" & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", CurrencyID=" & aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) & ", ConceptQttyID=" & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & ", ConceptTypeID=" & aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) & ", ConceptMin=" & aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) & ", ConceptMinQttyID=" & aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) & ", ConceptMax=" & aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) & ", ConceptMaxQttyID=" & aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) & ", AppliesToID='" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "', AbsenceTypeID=" & aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) & ", ConceptOrder=" & aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StartUserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", UploadedFileName='" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "") & "', Comments='" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "', RegistrationDate='" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								ElseIf aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) < lConceptStartDate Then
									sErrorDescription = "No se puede agregar un concepto con fecha menor a la registrada."
									lErrorNumber = -1
								ElseIf (lConceptEndDate <> 30000000) And (aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) < lConceptEndDate) Then
									sErrorDescription = "No se puede insertar el nuevo concepto del empleado puesto que se traslapa con otro ya registrado."
									lErrorNumber = -1
								Else
									sErrorDescription = "No se pudo insertar el nuevo concepto del empleado."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesConceptsLKP (EmployeeID, ConceptID, StartDate, EndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, RegistrationDate, ModifyDate, StartUserID, EndUserID, UploadedFileName, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) & ", '" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "', " & aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", '" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								End If
							End If
						Else
							sErrorDescription = "No se pudo insertar el nuevo concepto del empleado."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesConceptsLKP (EmployeeID, ConceptID, StartDate, EndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, RegistrationDate, ModifyDate, StartUserID, EndUserID, UploadedFileName, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) & ", '" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "', " & aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", '" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
						oRecordset.Close
					End If
				Else
					sErrorDescription = "No se pudo obtener la información del concepto del empleado."
					If (Len(oRequest("Authorization").Item) > 0) And (lReasonID = 14) Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, StartDate, EndDate From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (Active <> 2) Order By StartDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesConceptsLKP (EmployeeID, ConceptID, StartDate, EndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, RegistrationDate, ModifyDate, StartUserID, EndUserID, UploadedFileName, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) / 2 & ", " & aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) & ", '" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "', " & aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", '" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, StartDate, EndDate From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate<=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (Active <> 2) Order By StartDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								lConceptStartDate = CLng(oRecordset.Fields("StartDate").Value)
								lConceptEndDate = CLng(oRecordset.Fields("EndDate").Value)
								If lConceptEndDate < aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) Then
									If aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 7 Then
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesConceptsLKP (EmployeeID, ConceptID, StartDate, EndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, RegistrationDate, ModifyDate, StartUserID, EndUserID, UploadedFileName, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", " & (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) / 2) & ", " & aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) & ", '" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "', " & aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", '" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									Else
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesConceptsLKP (EmployeeID, ConceptID, StartDate, EndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, RegistrationDate, ModifyDate, StartUserID, EndUserID, UploadedFileName, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) & ", '" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "', " & aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", '" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									End If
								Else
									If aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 7 Then
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set Active=1, ModifyDate=" & sDate & ", ConceptAmount= " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) / 2 & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & lConceptStartDate & ") And (EndDate=" & lConceptEndDate & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
									Else
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set Active=1, ModifyDate=" & sDate & ", ConceptAmount= " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & lConceptStartDate & ") And (EndDate=" & lConceptEndDate & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
									End If
								End If
								oRecordset.MoveNext
								If Not oRecordset.EOF Then
									lConceptStartDate = CLng(oRecordset.Fields("StartDate").Value)
									lConceptEndDate = CLng(oRecordset.Fields("EndDate").Value)
									If aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) <= lConceptEndDate Then
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set EndDate=" & AddDaysToSerialDate(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", ModifyDate=" & sDate & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & lConceptStartDate & ") And (EndDate=" & lConceptEndDate & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
									End If
								End If
							Else
								sErrorDescription = "No se pudo insertar el nuevo concepto del empleado."
								If aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 7 Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesConceptsLKP (EmployeeID, ConceptID, StartDate, EndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, RegistrationDate, ModifyDate, StartUserID, EndUserID, UploadedFileName, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", " & (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) / 2) & ", " & aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) & ", '" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "', " & aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", '" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								Else
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesConceptsLKP (EmployeeID, ConceptID, StartDate, EndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, RegistrationDate, ModifyDate, StartUserID, EndUserID, UploadedFileName, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) & ", '" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "', " & aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", '" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								End If
							End If
							oRecordset.Close
						End If
					End If
				End If
			End If
		End If
	End If

	ModifyEmployeeConceptSp = lErrorNumber
	Err.Clear
End Function

Function ModifyEmployeeCredit(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the credit information for the employee in
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyEmployeeCredit"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del credito a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Not CheckEmployeeCreditInformationConsistency(aEmployeeComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sErrorDescription = "No se pudo modificar la información del credito del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Credits Set CreditTypeID='" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & "', ContractNumber = '" & Replace(aEmployeeComponent(S_CREDIT_CONTRACT_NUMBER_EMPLOYEE), "'", "´") & "', AccountNumber = '" & Replace(aEmployeeComponent(S_CREDIT_ACCOUNT_NUMBER_EMPLOYEE), "'", "´") & "', PaymentsNumber = " & aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) & ", EndDate = " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", FinishDate = " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", Comments = '" & aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) & "' Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (CreditID = " & aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		End If
	End If

	ModifyEmployeeCredit = lErrorNumber
	Err.Clear
End Function

Function ModifyEmployeeCreditor(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the beneficiary information for the employee in
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyEmployeeCreditor"
	Dim oRecordset
	Dim lErrorNumber
	Dim sField
	Dim sQuery
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_ID_CREDITOR_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del beneficiario para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Not CheckEmployeeCreditorInformationConsistency(aEmployeeComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sErrorDescription = "No se pudo modificar la información del beneficiario del empleado."
			If aEmployeeComponent(N_END_DATE_CREDITOR_EMPLOYEE) > 0 Then
				sField = "StartUserID"
			Else
				sField = "EndUserID"
			End If
			sQuery = "Update EmployeesCreditorsLKP" & _
					 " Set EndDate=" & aEmployeeComponent(N_END_DATE_CREDITOR_EMPLOYEE) & _
					 ", CreditorNumber=" & aEmployeeComponent(S_NUMBER_CREDITOR_EMPLOYEE) & _
					 ", CreditorName='" & Replace(aEmployeeComponent(S_NAME_CREDITOR_EMPLOYEE), "'", "´") & _
					 "', CreditorLastName='" & Replace(aEmployeeComponent(S_LAST_NAME_CREDITOR_EMPLOYEE), "'", "´") & _
					 "', CreditorLastName2='" & Replace(aEmployeeComponent(S_LAST_NAME2_CREDITOR_EMPLOYEE), "'", "´") & _
					 "', CreditorBirthDate=" & aEmployeeComponent(N_BIRTH_DATE_CREDITOR_EMPLOYEE) & _
					 ", ConceptAmount=" & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & _
					 ", CreditorTypeID=" & aEmployeeComponent(N_CREDITOR_TYPE_ID_EMPLOYEE) & _
					 ", PaymentCenterID=" & aEmployeeComponent(N_PAYMENT_CENTER_ID_CREDITOR_EMPLOYEE) & _
					 ", " & sField & "=" & aLoginComponent(N_USER_ID_LOGIN) & _
					 ", Comments='" & Replace(aEmployeeComponent(S_COMMENTS_CREDITOR_EMPLOYEE), "'", "´") & _
					 "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & _
					 ") And (CreditorID=" & aEmployeeComponent(N_ID_CREDITOR_EMPLOYEE) & _
					 ") And (StartDate=" & aEmployeeComponent(N_START_DATE_CREDITOR_EMPLOYEE) & ")"
			
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		End If
	End If

	ModifyEmployeeCreditor = lErrorNumber
	Err.Clear
End Function

Function ModifyEmployeeDocument(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the credit information for the employee in
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyEmployeeDocument"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para modificar la solicitud de hoja única de servicio."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo modificar la información del credito del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesDocs Set DocumentTime=" & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_TIME) & ", Document2Date=" & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_DATE_2) & ", Document2Time=" & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_TIME_2) & ", DocumentNumber='" & aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE) &  "', Authorizers='" & aEmployeeComponent(S_EMPLOYEE_AUTHORIZERS) & "', Authorized='" & aEmployeeComponent(S_EMPLOYEE_AUTHORIZED) & "', DocumentTypeID=" & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_TYPE) & " Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (DocumentDate = " & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_DATE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	ModifyEmployeeDocument = lErrorNumber
	Err.Clear
End Function

Function ModifyEmployeeHistoryList(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing employee in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyEmployeeHistoryList"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sDate
	Dim lHistoryEmployeeDate
	Dim lHistoryEndDate

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If lErrorNumber = 0 Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeDate, EndDate From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID<>0) And (ReasonID<>58) Order By EmployeeDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If Not oRecordset.EOF Then
				oRecordset.MoveNext
				If Not oRecordset.EOF Then
					lHistoryEmployeeDate = CLng(oRecordset.Fields("EmployeeDate").Value)
					lHistoryEndDate = CLng(oRecordset.Fields("EndDate").Value)
					sErrorDescription = "No se pudo actualizar la información del empleado al aplicar el movimiento"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EndDate=" & AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), -1) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & lHistoryEmployeeDate & ") And (EndDate=" & lHistoryEndDate & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	ModifyEmployeeHistoryList = lErrorNumber
	Err.Clear
End Function

Function ModifyEmployeeNumber(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the employee number
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyEmployeeNumber"
	Dim oRecordset
	Dim lErrorNumber
	Dim lEmployeeID
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If
	
	lEmployeeID = aEmployeeComponent(N_ID_EMPLOYEE)
	sErrorDescription = "No se pudo guardar la información del nuevo empleado."
	If InStr(1, ",5,6,7,8,9,", "," & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ",", vbBinaryCompare) = 0 Then
		sErrorDescription = "No se pudo obtener la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From ConsecutiveIDs Where (IDType=-1)", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("CurrentID").Value) + 1
			End If
		End If
		sErrorDescription = "No se pudo guardar la información del nuevo empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConsecutiveIDs Set CurrentID=CurrentID+1 Where (IDType=-1)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	ElseIf InStr(1, ",5,6,", "," & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ",", vbBinaryCompare) > 0 Then
		sErrorDescription = "No se pudo obtener la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From ConsecutiveIDs Where (IDType=5)", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("CurrentID").Value) + 1
			End If
		End If
		sErrorDescription = "No se pudo guardar la información del nuevo empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConsecutiveIDs Set CurrentID=CurrentID+1 Where (IDType=5)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	Else
		sErrorDescription = "No se pudo obtener la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From ConsecutiveIDs Where (IDType=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("CurrentID").Value) + 1
			End If
		End If
		sErrorDescription = "No se pudo guardar la información del nuevo empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConsecutiveIDs Set CurrentID=CurrentID+1 Where (IDType=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If
	If lErrorNumber = 0 Then
		aEmployeeComponent(S_NUMBER_EMPLOYEE) = Right("000000" & aEmployeeComponent(N_ID_EMPLOYEE), Len("000000"))
		aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE) = Right("000000" & aEmployeeComponent(N_ID_EMPLOYEE), Len("000000"))
		aEmployeeComponent(S_PASSWORD_EMPLOYEE) = Right("000000" & aEmployeeComponent(N_ID_EMPLOYEE), Len("000000"))
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ", EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', EmployeeAccessKey='" & Replace(aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE), "'", "") & "', EmployeePassword='" & Replace(aEmployeeComponent(S_PASSWORD_EMPLOYEE), "'", "") & "' Where (EmployeeID =" & lEmployeeID & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		If lErrorNumber = 0 Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ", EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "' Where (EmployeeID =" & lEmployeeID & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			If lErrorNumber = 0 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " Where (EmployeeID =" & lEmployeeID & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				If lErrorNumber = 0 Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesExtraInfo Set EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " Where (EmployeeID =" & lEmployeeID & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					If lErrorNumber = 0 Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRequirementsFM1LKP Set EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " Where (EmployeeID =" & lEmployeeID & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
						If lErrorNumber = 0 Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update JobsHistoryList Set EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " Where (EmployeeID =" & lEmployeeID & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
							If lErrorNumber = 0 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesRisksLKP Set EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " Where (EmployeeID =" & lEmployeeID & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
								If lErrorNumber = 0 Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesSchoolLevelsLKP Set EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " Where (EmployeeID =" & lEmployeeID & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
									If lErrorNumber = 0 Then
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Jobs Set JobID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " Where (JobID =" & lEmployeeID & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
										If lErrorNumber = 0 Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update JobsHistoryList Set JobID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " Where (JobID =" & lEmployeeID & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	End If

	ModifyEmployeeNumber = lErrorNumber
	Err.Clear
End Function

Function ModifyEmployeePayroll(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To send the concepts for the employee to the payroll
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyEmployeePayroll"
	Dim alConceptID
	Dim adAmount
	Dim iIndex
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador de la nómina para agregar los conceptos de pago del empleado."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudieron agregar los conceptos de pago del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		If lErrorNumber = 0 Then
			If Len(oRequest("ConceptsIDs").Item) > 0 Then
				alConceptID = Split(oRequest("ConceptsIDs").Item, LIST_SEPARATOR, -1, vbBinaryCompare)
				adAmount = Split(oRequest("ConceptsAmounts").Item, LIST_SEPARATOR, -1, vbBinaryCompare)
				If UBound(adAmount) < UBound(alConceptID) Then adAmount = Split(JoinLists(oRequest("ConceptsAmounts").Item, BuildList("0", LIST_SEPARATOR, UBound(alConceptID) + 1), LIST_SEPARATOR), LIST_SEPARATOR, -1, vbBinaryCompare)
				For iIndex = 0 To UBound(alConceptID)
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayRollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & aEmployeeComponent(N_PAYROLL_ID_EMPLOYEE) & ", 1, " & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & alConceptID(iIndex) & ", 1, " & adAmount(iIndex) & ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					If lErrorNumber <> 0 Then Exit For
				Next
			End If
		End If
	End If

	ModifyEmployeePayroll = lErrorNumber
	Err.Clear
End Function

Function SetActiveForEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given employee
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForEmployee"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo modificar la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set Active=" & CInt(oRequest("SetActive").Item) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	SetActiveForEmployee = lErrorNumber
	Err.Clear
End Function

Function SetActiveForEmployeeAbsences(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given employee's concept
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForEmployeeAbsences"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim iAbscenceID

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	Select Case CLng(oRequest("ReasonID").Item)
		Case EMPLOYEES_EXTRAHOURS
			If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set Active=1 Where (AbsenceID=201) And (Active<1)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set Active=1 Where (AbsenceID=201) And (Active<1) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			End If
		Case EMPLOYEES_SUNDAYS
			If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set Active=1 Where (AbsenceID=202) And (Active<1)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set Active=1 Where (AbsenceID=202) And (Active<1) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			End If
		Case Else
			If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
				lErrorNumber = -1
				sErrorDescription = "No se especificó el identificador del concepto a modificar."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				sErrorDescription = "No se pudo modificar la información."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set Active=1 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (AbsenceID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (OcurredDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			End If
	End Select

	SetActiveForEmployeeAbsences = lErrorNumber
	Err.Clear
End Function

Function SetActiveForEmployeeAdjustment(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given employee's concept
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForEmployeeAdjustment"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del concepto a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo modificar la información del reclamo o ajuste."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAdjustmentsLKP Set Active=1 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (MissingDate=" & aEmployeeComponent(N_MISSING_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	SetActiveForEmployeeAdjustment = lErrorNumber
	Err.Clear
End Function

Function SetActiveForEmployeeBankAccount(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given employee's concept
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForEmployeeBankAccount"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim Start_date
	
	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del concepto a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AccountID From BankAccounts Where (AccountID<>" & aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) & ") And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EndDate>=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") Order by StartDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StartDate from BankAccounts where (AccountID=" & CLng(oRecordset.Fields("AccountID").Value) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		Start_date = CLng(oRecordset.Fields("StartDate").Value)
		
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From BankAccounts Where (AccountID<>" & aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) & ") And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EndDate>=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") Order by StartDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If ((lErrorNumber = 0) And (Not oRecordset.EOF)) Then
		'	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update BankAccounts Set EndDate=" & AddDaysToSerialDate(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1) & " Where (AccountID=" & CLng(oRecordset.Fields("AccountID").Value) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update BankAccounts SET StartDate = " & Start_date & " WHERE  (AccountID=" & aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "delete from BankAccounts where (AccountID=" & CLng(oRecordset.Fields("AccountID").Value) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			
			
		End If
		sErrorDescription = "No se pudo modificar la información."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update BankAccounts Set Active=1, ApliedDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & " Where (AccountID=" & aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	SetActiveForEmployeeBankAccount = lErrorNumber
	Err.Clear
End Function

Function SetActiveForEmployeeBeneficiary(oRequest, oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given employee's beneficiary
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForEmployeeBeneficiary"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	Select Case lReasonID
		Case EMPLOYEES_ADD_BENEFICIARIES
			If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) = -1) Or (aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE) = -1) Then
				lErrorNumber = -1
				sErrorDescription = "No se especificó el identificador del beneficiario de pensión alimenticia a modificar."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				sErrorDescription = "No se pudo modificar la información del beneficiario de pensión alimenticia."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesBeneficiariesLKP Set Active=1 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (BeneficiaryID=" & aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			End If
		Case EMPLOYEES_CREDITORS
			If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_ID_CREDITOR_EMPLOYEE) = -1) Or (aEmployeeComponent(N_START_DATE_CREDITOR_EMPLOYEE) = -1) Then
				lErrorNumber = -1
				sErrorDescription = "No se especificó el identificador del acreedor a modificar."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				sErrorDescription = "No se pudo modificar la información del acreedor."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesCreditorsLKP Set Active=1 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (CreditorID=" & aEmployeeComponent(N_ID_CREDITOR_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(N_START_DATE_CREDITOR_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			End If
	End Select

	SetActiveForEmployeeBeneficiary = lErrorNumber
	Err.Clear
End Function

Function SetActiveForEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given employee's concept
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForEmployeeConcept"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del concepto a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (EndDate>=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (Active=1) Order by StartDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If ((lErrorNumber = 0) And (Not oRecordset.EOF)) Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set EndDate=" & AddDaysToSerialDate(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & CLng(oRecordset.Fields("StartDate").Value) & ") And (EndDate=" & CLng(oRecordset.Fields("EndDate").Value) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		End If
		sErrorDescription = "No se pudo modificar la información del concepto."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set Active=1 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	SetActiveForEmployeeConcept = lErrorNumber
	Err.Clear
End Function

Function SetActiveForEmployeeCredit(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new credit for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForEmployeeCredit"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim iEndDate
	Dim sQuery
	
	Dim lCreditID
	Dim iPaymentNumber

	iPaymentNumber = CInt(oRequest("ConceptQttyID").Item)
	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If Not CheckEmployeeConceptInformationConsistency(aEmployeeComponent, sErrorDescription) Then
		lErrorNumber = -1
		sErrorDescription = "No se pudo validar la consistencia de la información."
	Else
		If Not CheckEmployeeCreditInformationConsistency(aEmployeeComponent, sErrorDescription) Then
			lErrorNumber = -1
			sErrorDescription = "No se pudo validar la consistencia de la información."
		Else		
			sQuery = "Insert Into Credits" & _
						" (EmployeeID, CreditID, CreditTypeID, ContractNumber, AccountNumber," & _
						" PaymentsNumber, PeriodID, StartDate, EndDate, FinishDate, QttyID," & _
						" AppliesToID, StartAmount, PaymentAmount, DebtAmount, PaymentsCounter," & _
						" Active, UploadedFileName, Comments, UploadedRecordType)" & _
						" Values" & _
						" ('" & aEmployeeComponent(N_ID_EMPLOYEE) & "', '" & aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) & "', '" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & "', '" & aEmployeeComponent(S_CREDIT_CONTRACT_NUMBER_EMPLOYEE) & "', '" & _
						aEmployeeComponent(S_CREDIT_ACCOUNT_NUMBER_EMPLOYEE) & "', '" & aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) & "', '" & aEmployeeComponent(N_CREDIT_PERIOD_ID_EMPLOYEE) & "', '" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & "', '" & _
						aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & "', '" & aEmployeeComponent(L_CREDIT_FINISH_DATE_EMPLOYEE) & "', '" & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & "', '" &aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "', '" & _
						aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & "', '" & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & "', '" & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & "', '" & _
						aEmployeeComponent(N_CREDIT_PAYMENTS_COUNTER_EMPLOYEE) & "', '" & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & "', '" & _
						aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) & "', '" & aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) & "', '" & _
						aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) & "')"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If lErrorNumber <> 0 Then
				sErrorCarga = "Error al insertar el registro en la tabla Credits para el empleado: " & aEmployeeComponent(N_ID_EMPLOYEE)
			End If
		End If
	End If
	SetActiveForEmployeeCredit = lErrorNumber
	Err.Clear
End Function

Function SetActiveForEmployeeCreditFile(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new credit for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForEmployeeCreditFile"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim iEndDate
	Dim sQuery
	
	Dim lCreditID
	Dim iPaymentNumber

	iPaymentNumber = CInt(oRequest("ConceptQttyID").Item)
	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If Not CheckEmployeeConceptInformationConsistency(aEmployeeComponent, sErrorDescription) Then
		lErrorNumber = -1
		sErrorDescription = "No se pudo validar la consistencia de la información."
	Else
		If Not CheckEmployeeCreditInformationConsistency(aEmployeeComponent, sErrorDescription) Then
			lErrorNumber = -1
			sErrorDescription = "No se pudo validar la consistencia de la información."
		Else
			lErrorNumber = GetEmployeeCredit(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			Select Case aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE)
				Case 0 ' Captura en línea
					sQuery = "Select * From Credits Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (CreditTypeID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (EndDate>=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (Active=1) Order By StartDate Desc"
					sErrorDescription = "Error al obtener los datos del crédito."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sQuery = "Update Credits Set EndDate=" & CLng(AddDaysToSerialDate(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE),-1)) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (CreditID=" & CLng(oRecordset.Fields("CreditID").Value) & ")"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						End If
					End If
					sQuery = "Update Credits" & _
						" Set Active=1, UploadedRecordType=0, UploadedFileName=' ', Comments=' '" & _
						" Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
						" And (CreditID=" & aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) & ")"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				Case 1, 2 ' Altas y Cambios
					aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = GetPayrollStartDate(CLng(oRequest("AppliedDate").Item))
					sQuery = "Select * From Credits Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (CreditTypeID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (EndDate>=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (Active=1) Order By StartDate Desc"
					sErrorDescription = "Error al obtener los datos del crédito."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sQuery = "Update Credits Set EndDate=" & CLng(AddDaysToSerialDate(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1)) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (CreditID=" & CLng(oRecordset.Fields("CreditID").Value) & ")"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						End If
					End If
					sQuery = "Update Credits" & _
						" Set StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", Active=1, UploadedRecordType=0, UploadedFileName=' ', Comments=' '" & _
						" Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
						" And (CreditID=" & aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) & ")"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				Case 3 ' Baja
					sQuery = "Select * From Credits" & _
						" Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
						" And (CreditTypeID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ")" & _
						" And (EndDate>=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")" & _
						" And (Active=1) Order By StartDate Desc"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber <> 0 Then
						sErrorDescription = "Error al actualizar el cambio de crédito."
					Else
						If Not oRecordset.EOF Then
							sQuery = "Update Credits Set CreditTypeID='" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & "'," & _
									 " ContractNumber='" & Replace(aEmployeeComponent(S_CREDIT_CONTRACT_NUMBER_EMPLOYEE), "'", "´") & "'," & _
									 " AccountNumber='" & Replace(aEmployeeComponent(S_CREDIT_ACCOUNT_NUMBER_EMPLOYEE), "'", "´") & "'," & _
									 " PaymentsNumber=" & aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) & "," & _
									 " PeriodID=" & aEmployeeComponent(N_CREDIT_PERIOD_ID_EMPLOYEE) & "," & _
									 " EndDate=" & AddDaysToSerialDate(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1) & "," & _
									 " FinishDate=" & AddDaysToSerialDate(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1) & "," & _
									 " StartAmount=" & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & "," & _
									 " UploadedRecordType=3," & _
									 " Comments='" & aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) & "'" & _
									 " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
									 " And (CreditID=" & CLng(oRecordset.Fields("CreditID").Value) & ")"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							If lErrorNumber <> 0 Then
								sErrorDescription = "Error al aplicar la baja del Credito para el empleado: " & aEmployeeComponent(N_ID_EMPLOYEE)
							Else
								sQuery = "Delete From Credits" & _
									" Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
									" And (CreditID=" & aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) & ")"
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							End If
						Else
							sQuery = "Update Credits" & _
								" Set Active=1, UploadedRecordType=0, UploadedFileName=' ', Comments=' '" & _
								" Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
								" And (CreditID=" & aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) & ")"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						End If
					End If
			End Select
			If lErrorNumber <> 0 Then
				sErrorDescription = "Error al insertar el registro en la tabla Credits para el empleado: " & aEmployeeComponent(N_ID_EMPLOYEE)
			End If
		End If
	End If
	SetActiveForEmployeeCreditFile = lErrorNumber
	Err.Clear
End Function

Function SetActiveForEmployeeGrade(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given employee's concept
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForEmployeeGrade"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número de empleado del registro a aplicar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesGrades Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EndDate>=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (Active=1) Order by StartDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If ((lErrorNumber = 0) And (Not oRecordset.EOF)) Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesGrades Set EndDate=" & AddDaysToSerialDate(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1) & " Where (EmployeeID=" & CLng(oRecordset.Fields("EmployeeID").Value) & ") And (StartDate=" & CLng(oRecordset.Fields("StartDate").Value) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		End If
		sErrorDescription = "No se pudo aplicar la calificación en proceso."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesGrades Set Active=1 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	SetActiveForEmployeeGrade = lErrorNumber
	Err.Clear
End Function

Function SetDeActiveForEmployeeAbsences(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given employee's concept
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetDeActiveForEmployeeAbsences"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del concepto a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo modificar la información."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set Active=2 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (AbsenceID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (OcurredDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	SetDeActiveForEmployeeAbsences = lErrorNumber
	Err.Clear
End Function

Function SetDeActiveForEmployeeConcepts(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given employee's concept
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetDeActiveForEmployeeConcepts"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del concepto a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo modificar la información."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set Active=2 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	SetDeActiveForEmployeeAbsences = lErrorNumber
	Err.Clear
End Function

Function RemoveEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To remove an employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployee"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el empleado a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Employees Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From BankAccounts Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Credits Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From DocumentsForLicenses Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesAbsencesLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesAdjustmentsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesAntiquitiesLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesBeneficiariesLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesChangesLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesChildrenLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesDocs Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesExtraInfo Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesFONAC Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesForTaxAdjustment Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			'sErrorDescription = "No se pudo eliminar la información del empleado."
			'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesFields Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesHandicapsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesInformation Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesKardex Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesKardex2 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesKardex4 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesReasonsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesRequirementsFM1LKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesRequirementsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesRisksLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesSchoolLevelsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesSpecialJourneys Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesSyndicatesLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PaperworkOwners Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PaperworkOwners Where (OwnerID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PaperworkOwnersLKP Where (OwnerID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payments Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PaymentsMessages Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PaymentsRecords2 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_Antiquities Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From SADE_Constancias Where (ID_Usuario=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From SADE_CursosEmpleadosLKP Where (ID_Empleado=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From SADE_EntradasCursos Where (ID_Usuario=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From SADE_NewCourse Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Jobs Set OwnerID=-1 Where (OwnerID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PaperworkComments Set OwnerID=-1 Where (OwnerID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PaperworkComments Set OwnerID=-1 Where (OwnerID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Paperworks Set OwnerID=-1 Where (OwnerID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		End If
	End If

	RemoveEmployee = lErrorNumber
	Err.Clear
End Function

Function RemoveEmployeeAdjustments(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a concept adjustment for the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployeeAdjustments"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del concepto del ajuste a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del concepto del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete from EmployeesAdjustmentsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (MissingDate=" & aEmployeeComponent(N_MISSING_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	RemoveEmployeeAdjustments = lErrorNumber
	Err.Clear
End Function

Function RemoveEmployeeBankAccount(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a bank account from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployeeBankAccount"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador de la cuenta del empleado."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información de la cuenta del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From BankAccounts Where (AccountID=" & aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	RemoveEmployeeBankAccount = lErrorNumber
	Err.Clear
End Function

Function RemoveEmployeeBeneficiary(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a concept for the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployeeBeneficiary"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del beneficiario para eliminar el registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del beneficiario del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesBeneficiariesLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (BeneficiaryID=" & aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	RemoveEmployeeBeneficiary = lErrorNumber
	Err.Clear
End Function

Function RemoveEmployeeChild(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To remove an employee's child from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployeeChild"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_ID_CHILD_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o de su hijo(a) para eliminar el registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del hijo(a) del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesChildrenLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ChildID=" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	RemoveEmployeeChild = lErrorNumber
	Err.Clear
End Function

Function RemoveEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a concept for the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployeeConcept"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del concepto a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del concepto del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set EndDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	RemoveEmployeeConcept = lErrorNumber
	Err.Clear
End Function

Function RemoveEmployeeConceptsFile(oRequest, oADODBConnection, sQuery, lReasonID, aEmployeeComponent, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a group of records for the employee into the database
'Inputs:  oRequest, oADODBConnection, lReasonID
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployeeConceptsFile"
	Dim oRecordset
	Dim lErrorNumber
	Dim sErrorQueries

	sErrorDescription = "No se pudo obtener la información de la aplicación de los conceptos de pago de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sErrorQueries=""
			Do While Not oRecordset.EOF
				Select Case lReasonID
					Case -58
						aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
						aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = CLng(oRecordset.Fields("ConceptID").Value)
						aEmployeeComponent(N_MISSING_DATE_EMPLOYEE) = CLng(oRecordset.Fields("MissingDate").Value)
						lErrorNumber = GetEmployeeAdjustments(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						If lErrorNumber = 0 Then
							lErrorNumber = RemoveEmployeeAdjustments(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
							If lErrorNumber <> 0 Then
								sErrorDescription = "No se pudo eliminar el ajuste del empleado " & CStr(aEmployeeComponent(N_ID_EMPLOYEE)) & ", del concepto " & CStr(aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE))
								sErrorQueries = sErrorQueries & "<B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
							End If
						Else
							sErrorDescription = "No se pudo eliminar el ajuste del empleado " & CStr(aEmployeeComponent(N_ID_EMPLOYEE)) & ", del concepto " & CStr(aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE))
							sErrorQueries = sErrorQueries & "<B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
						End If
					Case EMPLOYEES_THIRD_PROCESS
						'If Not IsEmpty(oRequest(CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("CreditID").Value))) Then
							aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
							aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) = CLng(oRecordset.Fields("CreditID").Value)
							lErrorNumber = GetEmployeeCredit(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							If lErrorNumber = 0 Then
								lErrorNumber = RemoveEmployeeCredit(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
								If lErrorNumber <> 0 Then
									sErrorDescription = "No se pudo eliminar la información del registro de " & CStr(aEmployeeComponent(N_ID_EMPLOYEE)) & ", con crédito " & CStr(aEmployeeComponent(N_CREDIT_ID_EMPLOYEE))
									sErrorQueries = sErrorQueries & "<B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
								End If
							Else
								sErrorDescription = "No se pudo eliminar la información del registro de " & CStr(aEmployeeComponent(N_ID_EMPLOYEE)) & ", con crédito " & CStr(aEmployeeComponent(N_CREDIT_ID_EMPLOYEE))
								sErrorQueries = sErrorQueries & "<B>ERROR: </B>" & sErrorDescription & "<BR /><BR />"
							End If
						'End If
				End Select
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
		End If
	End If
	If Len(sErrorQueries) > 0 Then
		lErrorNumber = -1
		sErrorDescription = "<BR /><B>NO SE PUDIERON AGREGAR LOS SIGUIENTES RENGLONES:</B><BR /><BR />" & sErrorQueries
	End If
	Set oRecordset = Nothing
	RemoveEmployeeConceptsFile = lErrorNumber
	Err.Clear
End Function

Function RemoveEmployeeCredit(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a concept for the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployeeCredit"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del credito a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del concepto del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete Credits Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (CreditID=" & aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	RemoveEmployeeCredit = lErrorNumber
	Err.Clear
End Function

Function RemoveEmployeeCreditor(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a concept for the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployeeCreditor"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_ID_CREDITOR_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del acreedor para eliminar el registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del acreedor del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesCreditorsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (CreditorID=" & aEmployeeComponent(N_ID_CREDITOR_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(N_START_DATE_CREDITOR_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	RemoveEmployeeCreditor = lErrorNumber
	Err.Clear
End Function

Function RemoveEmployeeCreditsRejected(oADODBConnection, sOriginalFileName, sErrorDescription)
'************************************************************
'Purpose: To add a new credit rejected report from third upload files
'		for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployeeCreditsRejected"
	Dim oRecordset
	Dim lErrorNumber


	sErrorDescription = "No se pudo eliminar la información de los registros de créditos rechazados para el archivo: " & sOriginalFileName
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete from UploadThirdCreditsRejected where UploadedFileName like '" & sOriginalFileName & "'", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, null)

	RemoveEmployeeCreditsRejected = lErrorNumber
	Err.Clear	
End Function

Function RemoveEmployeeDocument(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a document for the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployeeDocument"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para eliminar el registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del acreedor del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesDocs Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (DocumentDate=" & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_DATE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	RemoveEmployeeDocument = lErrorNumber
	Err.Clear
End Function

Function RemoveEmployeeGrade(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a grade for the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployeeGrade"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para eliminar el registro de calificación."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del acreedor del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesGrades Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	RemoveEmployeeGrade = lErrorNumber
	Err.Clear
End Function

Function RemoveEmployeeForValidationSP(oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To cancel a move to an employee
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployeeForValidation"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim oRecordset

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	sErrorDescription = "No se pudo cancelar el movimiento al empleado."
	Select Case aEmployeeComponent(N_REASON_ID_EMPLOYEE)
		Case 12
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set JobID=-1, StatusID=-2, bProcessed=0 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (bProcessed=2)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			If lErrorNumber = 0 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set JobID=-1, StatusID=-2 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			End If
		Case Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (bProcessed=2)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End Select

	RemoveEmployeeForValidation = lErrorNumber
	Err.Clear
End Function

Function RemoveEmployeeForValidation(oRequest, oADODBConnection, sAction, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To cancel a move to an employee
'Inputs:  oRequest, oADODBConnection, sAction
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployeeForValidation"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim oRecordset
	Dim iEmployeeId
	Dim iEmployeeConceptDate
	Dim iConceptId
	Dim sQuery
	Dim bProcessed
	Dim lStatusID

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	Select Case sAction
		Case "EmployeesMovements"
			sErrorDescription = "No se pudo cancelar el movimiento."
			Select Case aEmployeeComponent(N_REASON_ID_EMPLOYEE)
				Case 12
					If lErrorNumber = 0 Then lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=-2, Active=0 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					If lErrorNumber = 0 Then lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set JobID=-1, ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=-2, Active=0 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesHistoryList Where (bProcessed=2) And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ") And (EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				Case 18
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesHistoryList Where (bProcessed=2) And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ") And (EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And (EndDate=" & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				Case 13,14,17,18
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesHistoryList Where (bProcessed=2) And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ") And (EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And (EndDate=" & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					If lErrorNumber = 0 Then lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=-2, Active=0 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				Case -58
					If lErrorNumber = 0 Then lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesAdjustmentsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (MissingDate=" & aEmployeeComponent(N_MISSING_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
					sQuery = "Delete From EmployeesAbsencesLKP" & _
							" Where " & _
							" (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
							" And (AbsenceID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ")" & _
							" And (OcurredDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")"
					sErrorDescription = "No se pudo eliminar el registro seleccionado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				Case -89, EMPLOYEES_HONORARIUM_CONCEPT, EMPLOYEES_SAFE_SEPARATION, EMPLOYEES_ADD_SAFE_SEPARATION, 53, EMPLOYEES_ANTIQUITIES, EMPLOYEES_ADDITIONALSHIFT, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_MONTHAWARD, EMPLOYEES_SPORTS_HELP, EMPLOYEES_SPORTS, EMPLOYEES_CARLOAN, EMPLOYEES_CONCEPT_C3, EMPLOYEES_BENEFICIARIES, EMPLOYEES_CONCEPT_08, EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_LICENSES, EMPLOYEES_CONCEPT_16, EMPLOYEES_NON_EXCENT, EMPLOYEES_EXCENT, EMPLOYEES_MOTHERAWARD, EMPLOYEES_HELP_COMISSION, EMPLOYEES_SAFEDOWN, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_FONAC_CONCEPT, EMPLOYEES_FONAC_ADJUSTMENT, EMPLOYEES_ANTIQUITY_25_AND_30_YEARS
					sQuery = "Delete From EmployeesConceptsLKP" & _
							" Where " & _
							" (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
							" And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ")" & _
							" And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")"
					sErrorDescription = "No se pudo eliminar el registro seleccionado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				Case EMPLOYEES_BENEFICIARIES_DEBIT
					sQuery = "Delete From Credits" & _
							" Where " & _
							" (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
							" And (CreditID=" & aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) & ")" '& _
							'" And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")"
					sErrorDescription = "No se pudo eliminar el registro seleccionado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				Case 1, 5, 6, 10, 2, 3, 4, 8, 62, 63, 66
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesHistoryList Where (bProcessed=2) And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ") And (EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					If lErrorNumber = 0 Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusID From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") Order By EmployeeDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						If Not oRecordset.EOF Then
							lStatusID = CLng(oRecordset.Fields("StatusID").Value)
							If lStatusID < 0 Then
								lStatusID = 0
							End If
							If lErrorNumber = 0 Then lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=" & lStatusID & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
					End If
				Case Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesHistoryList Where (bProcessed=2) And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ") And (EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					If lErrorNumber = 0 Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusID From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") Order By EmployeeDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						If Not oRecordset.EOF Then
							lStatusID = CLng(oRecordset.Fields("StatusID").Value)
							If lErrorNumber = 0 Then lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=" & lStatusID & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
					End If
			End Select
		Case "EmployeesAdditionalShift"
	End Select

	RemoveEmployeeForValidation = lErrorNumber
	Err.Clear
End Function

Function RemoveEmployeeReasonForRejection(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a reasons for rejection for employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployeeReasonForRejection"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim oRecordset

	sErrorDescription = "No se pudo eliminar las razones de rechazo del movimiento."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesReasonsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)

	RemoveEmployeeReasonForRejection = lErrorNumber
	Err.Clear
End Function

Function RemoveUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, sUploadedFileName, sErrorDescription)
'************************************************************
'Purpose: To remove rejected third credits records
'		for a file from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveUploadThirdCreditsRejected"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim iEndDate
	Dim sQuery
	Dim lCreditID
	Dim iPaymentNumber

	iPaymentNumber = CInt(oRequest("ConceptQttyID").Item)
	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If
	If Not CheckEmployeeConceptInformationConsistency(aEmployeeComponent, sErrorDescription) Then
		lErrorNumber = -1
		sErrorDescription = "No se pudo validar la consistencia de la información."
	End If

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select COUNT(*) As Cuantos from Credits where UploadedFileName='"& sUploadedFileName& "'", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If CInt(oRecordset.Fields("Cuantos").Value) = 0 Then
			sErrorDescription = "No se pudo eliminar la información de los registros rechazados del archivo" & sUploadedFileName
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete from UploadThirdCreditsRejected where UploadedFileName='" & sUploadedFileName & "'", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		End If
	End If

	RemoveUploadThirdCreditsRejected = lErrorNumber
	Err.Clear	
End Function

Function UpdateEmployeeHistoryListRecord(oRequest, oADODBConnection, iAction, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new record for the employee's history list
'Inputs:  oRequest, oADODBConnection, iAction, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UpdateEmployeeHistoryListRecord"
	Dim lEndDate
	Dim lStartDate
	Dim sEmployeeNumber
	Dim lCompanyID
	Dim lJobID
	Dim lServiceID
	Dim lZoneID
	Dim lEmployeeTypeID
	Dim lPositionTypeID
	Dim lClassificationID
	Dim lGroupGradeLevelID
	Dim lIntegrationID
	Dim lJourneyID
	Dim lShiftID
	Dim dWorkingHours
	Dim lAreaID
	Dim lPositionID
	Dim lLevelID
	Dim lStatusID
	Dim lPaymentCenterID
	Dim lRiskLevel
	Dim lActive
	Dim lReasonID
	Dim oRecordset
	Dim lErrorNumber

	Dim lEndDateForFillDateDiference
	Dim lStartDateForFillDateDiference
	Dim sEmployeeNumberForFillDateDiference
	Dim lCompanyIDForFillDateDiference
	Dim lJobIDForFillDateDiference
	Dim lServiceIDForFillDateDiference
	Dim lZoneIDForFillDateDiference
	Dim lEmployeeTypeIDForFillDateDiference
	Dim lPositionTypeIDForFillDateDiference
	Dim lClassificationIDForFillDateDiference
	Dim lGroupGradeLevelIDForFillDateDiference
	Dim lIntegrationIDForFillDateDiference
	Dim lJourneyIDForFillDateDiference
	Dim lShiftIDForFillDateDiference
	Dim dWorkingHoursForFillDateDiference
	Dim lAreaIDForFillDateDiference
	Dim lPositionIDForFillDateDiference
	Dim lLevelIDForFillDateDiference
	Dim lStatusIDForFillDateDiference
	Dim lPaymentCenterIDForFillDateDiference
	Dim lRiskLevelForFillDateDiference
	Dim lActiveForFillDateDiference
	Dim lReasonIDForFillDateDiference

	lStartDate = CLng(oRequest("EmployeeYear").Item & Right(("0" & oRequest("EmployeeMonth").Item), Len("00")) & Right(("0" & oRequest("EmployeeDay").Item), Len("00")))
	If Len(oRequest("EmployeeEndYear").Item) > 0 Then
		lEndDate = CLng(oRequest("EmployeeEndYear").Item & Right(("0" & oRequest("EmployeeEndMonth").Item), Len("00")) & Right(("0" & oRequest("EmployeeEndDay").Item), Len("00")))
	Else
		lEndDate = CLng(oRequest("EndYear").Item & Right(("0" & oRequest("EndMonth").Item), Len("00")) & Right(("0" & oRequest("EndDay").Item), Len("00")))
	End If
	If lEndDate = 0 Then lEndDate = 30000000
	If Len(oRequest("EmployeeNumber").Item) > 0 Then
		sEmployeeNumber = Right(("000000" & oRequest("EmployeeNumber").Item), Len("000000"))
	Else
		sEmployeeNumber = Right(("000000" & oRequest("EmployeeID").Item), Len("000000"))
	End If
	lJobID = oRequest("JobID").Item
	lEmployeeTypeID = -1
	lPositionTypeID = -1
	lStatusID = oRequest("StatusID").Item
	lReasonID = oRequest("ReasonID").Item
	lRiskLevel = 0
	lActive = oRequest("Active").Item
	lCompanyID = -1
	lServiceID = -1
	lZoneID = -1
	lClassificationID = -1
	lGroupGradeLevelID = -1
	lIntegrationID = -1
	lJourneyID = -1
	lShiftID = -1
	dWorkingHours = -1
	lAreaID = -1
	lPositionID = -1
	lLevelID = -1
	lPaymentCenterID = -1

	sErrorDescription = "No se pudo agregar el registro en el historial del empleado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Jobs.*, Positions.EmployeeTypeID, Positions.PositionTypeID, Areas.ZoneID As AreaZoneID From Jobs, Positions, Areas Where (Jobs.PositionID=Positions.PositionID) And (Jobs.AreaID=Areas.AreaID) And (Jobs.JobID=" & lJobID & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			lEmployeeTypeID = CStr(oRecordset.Fields("EmployeeTypeID").Value)
			lPositionTypeID = CStr(oRecordset.Fields("PositionTypeID").Value)
			lCompanyID = CStr(oRecordset.Fields("CompanyID").Value)
			lServiceID = CStr(oRecordset.Fields("ServiceID").Value)
			lZoneID = CStr(oRecordset.Fields("AreaZoneID").Value)
			lClassificationID = CStr(oRecordset.Fields("ClassificationID").Value)
			lGroupGradeLevelID = CStr(oRecordset.Fields("GroupGradeLevelID").Value)
			lIntegrationID = CStr(oRecordset.Fields("IntegrationID").Value)
			lJourneyID = CStr(oRecordset.Fields("JourneyID").Value)
			lShiftID = CStr(oRecordset.Fields("ShiftID").Value)
			dWorkingHours = CStr(oRecordset.Fields("WorkingHours").Value)
			lAreaID = CStr(oRecordset.Fields("AreaID").Value)
			lPositionID = CStr(oRecordset.Fields("PositionID").Value)
			lLevelID = CStr(oRecordset.Fields("LevelID").Value)
			lPaymentCenterID = CStr(oRecordset.Fields("PaymentCenterID").Value)
		Else
			lErrorNumber = -1
			sErrorDescription = "La plaza indicada no se encuentra en el catálogo de plazas."
		End If
	End If
	If lErrorNumber = 0 Then
		Set oRecordset = Nothing
		Select Case iAction
			Case 0
				sErrorDescription = "No se pudo agregar el registro en el historial del empleado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & lStartDate & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
			Case 1, 2
				If lErrorNumber = 0 Then
					lStartDateForFillDateDiference = 0
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from EmployeesHistoryList where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate<" & lStartDate & ") And (EndDate>" & lEndDate & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sEmployeeNumberForFillDateDiference = Right(("000000" & oRecordset.Fields("EmployeeID").Value), Len("000000"))
							lJobIDForFillDateDiference = CStr(oRecordset.Fields("JobID").Value)
							lEmployeeTypeIDForFillDateDiference = CStr(oRecordset.Fields("EmployeeTypeID").Value)
							lPositionTypeIDForFillDateDiference = CStr(oRecordset.Fields("PositionTypeID").Value)
							lStatusIDForFillDateDiference = CStr(oRecordset.Fields("StatusID").Value)
							lReasonIDForFillDateDiference = CStr(oRecordset.Fields("ReasonID").Value)
							lRiskLevelForFillDateDiference = CStr(oRecordset.Fields("RiskLevel").Value)
							lActiveForFillDateDiference = CStr(oRecordset.Fields("Active").Value)
							lStartDateForFillDateDiference = CLng(oRecordset.Fields("EmployeeDate").Value)
							lEndDateForFillDateDiference = CLng(oRecordset.Fields("EndDate").Value)
							lCompanyIDForFillDateDiference = CStr(oRecordset.Fields("CompanyID").Value)
							lServiceIDForFillDateDiference = CStr(oRecordset.Fields("ServiceID").Value)
							lZoneIDForFillDateDiference = CStr(oRecordset.Fields("ZoneID").Value)
							lClassificationIDForFillDateDiference = CStr(oRecordset.Fields("ClassificationID").Value)
							lGroupGradeLevelIDForFillDateDiference = CStr(oRecordset.Fields("GroupGradeLevelID").Value)
							lIntegrationIDForFillDateDiference = CStr(oRecordset.Fields("IntegrationID").Value)
							lJourneyIDForFillDateDiference = CStr(oRecordset.Fields("JourneyID").Value)
							lShiftIDForFillDateDiference = CStr(oRecordset.Fields("ShiftID").Value)
							dWorkingHoursForFillDateDiference = CStr(oRecordset.Fields("WorkingHours").Value)
							lAreaIDForFillDateDiference = CStr(oRecordset.Fields("AreaID").Value)
							lPositionIDForFillDateDiference = CStr(oRecordset.Fields("PositionID").Value)
							lLevelIDForFillDateDiference = CStr(oRecordset.Fields("LevelID").Value)
							lPaymentCenterIDForFillDateDiference = CStr(oRecordset.Fields("PaymentCenterID").Value)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo agregar el registro en el historial del empleado."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & AddDaysToSerialDate(lEndDate, 1) & ", " & lEndDateForFillDateDiference & ", '" & sEmployeeNumberForFillDateDiference & "', " & lCompanyIDForFillDateDiference & ", " & lJobIDForFillDateDiference & ", " & lServiceIDForFillDateDiference & ", " & lZoneIDForFillDateDiference & ", " & lEmployeeTypeIDForFillDateDiference & ", " & lPositionTypeIDForFillDateDiference & ", " & lClassificationIDForFillDateDiference & ", " & lGroupGradeLevelIDForFillDateDiference & ", " & lIntegrationIDForFillDateDiference & ", " & lJourneyIDForFillDateDiference & ", " & lShiftIDForFillDateDiference & ", " & dWorkingHoursForFillDateDiference & ", " & lAreaIDForFillDateDiference & ", " & lPositionIDForFillDateDiference & ", " & lLevelIDForFillDateDiference & ", " & lStatusIDForFillDateDiference & ", " & lPaymentCenterIDForFillDateDiference & ", " & lRiskLevelForFillDateDiference & ", " & lActiveForFillDateDiference & ", " & lReasonIDForFillDateDiference & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & "," & oRequest("PayrollDateHdn").Item & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", 1, '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
							End If
						Else
							Set oRecordset = Nothing
							If lErrorNumber = 0 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from EmployeesHistoryList where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate>=" & lStartDate & ") And (EmployeeDate<" & lEndDate & ") Order by EmployeeDate", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										sErrorDescription = "No se pudo agregar el registro en el historial del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EmployeeDate=" & AddDaysToSerialDate(lEndDate, 1) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & CLng(oRecordset.Fields("EmployeeDate").Value) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
									End If
								End If
							End If
						End If
					End If
				End If
				If lErrorNumber = 0 Then
					Set oRecordset = Nothing
					sErrorDescription = "No se pudo agregar el registro en el historial del empleado."
					If lStartDateForFillDateDiference > 0 Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EndDate=" & AddDaysToSerialDate(lStartDate, -1) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & lStartDateForFillDateDiference & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from EmployeesHistoryList where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate<" & lStartDate & ") Order by EmployeeDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								sErrorDescription = "No se pudo agregar el registro en el historial del empleado."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EndDate=" & AddDaysToSerialDate(lStartDate, -1) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & CLng(oRecordset.Fields("EmployeeDate").Value) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
							End If
						End If
					End If
				End If
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudo agregar el registro en el historial del empleado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate>=" & lStartDate & ") And (EndDate<=" & lEndDate & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				End If
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudo agregar el registro en el historial del empleado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & lStartDate & ", " & lEndDate & ", '" & sEmployeeNumber & "', " & lCompanyID & ", " & lJobID & ", " & lServiceID & ", " & lZoneID & ", " & lEmployeeTypeID & ", " & lPositionTypeID & ", " & lClassificationID & ", " & lGroupGradeLevelID & ", " & lIntegrationID & ", " & lJourneyID & ", " & lShiftID & ", " & dWorkingHours & ", " & lAreaID & ", " & lPositionID & ", " & lLevelID & ", " & lStatusID & ", " & lPaymentCenterID & ", " & lRiskLevel & ", " & lActive & ", " & lReasonID & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", 0, " & aLoginComponent(N_USER_ID_LOGIN) & ", 1, '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				End If
		End Select
		Call UpdateEmployeeFromHistoryList(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
	Else
		lErrorNumber = -1 
	End If

	Set oRecordset = Nothing
	UpdateEmployeeHistoryListRecord = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfBankID(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific employee exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfBankID"
	Dim oRecordset
	Dim lErrorNumber

	If aEmployeeComponent(N_BANK_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el Id del banco para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del Id del banco en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Banks Where (BankID=" & aEmployeeComponent(N_BANK_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				sErrorDescription = "No existe registro de banco para la clave indicada."
				lErrorNumber = L_ERR_NO_RECORDS
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfBankID = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfBankShortName(aEmployeeComponent, sBankShortName, sErrorDescription)
'************************************************************
'Purpose: To check if a specific employee exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfBankShortName"
	Dim oRecordset
	Dim lErrorNumber

	If aEmployeeComponent(N_BANK_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el Id del banco para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del Id del banco en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Banks Where (BankShortName='" & sBankShortName & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				sErrorDescription = "No existe registro de banco para la clave " & sBankShortName
				lErrorNumber = L_ERR_NO_RECORDS
			Else
				aEmployeeComponent(N_BANK_ID_EMPLOYEE) = CInt(oRecordset.Fields("BankID").Value)
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfBankShortName = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfEmployee(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific employee exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfEmployee"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_JOB)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If Len(aEmployeeComponent(S_NUMBER_EMPLOYEE)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del empleado en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Employees Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And ((RFC='" & Replace(aEmployeeComponent(S_RFC_EMPLOYEE), "'", "") & "') Or (CURP='" & Replace(aEmployeeComponent(S_CURP_EMPLOYEE), "'", "") & "'))", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aEmployeeComponent(B_IS_DUPLICATED_EMPLOYEE) = True
				aEmployeeComponent(N_ID_EMPLOYEE) = -1
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfEmployee = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfEmployeeBeneficiary(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific employee exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfEmployeeBeneficiary"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_JOB)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If Len(aEmployeeComponent(S_NUMBER_EMPLOYEE)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del empleado en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesBeneficiariesLKP Where (BeneficiaryNumber='" & Replace(aEmployeeComponent(S_NUMBER_BENEFICIARY_EMPLOYEE), "'", "") & "') And (((StartDate >= " &  aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE) & ") And (EndDate <= " &  aEmployeeComponent(N_END_DATE_BENEFICIARY_EMPLOYEE) & ")) Or ((EndDate >= " &  aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE) & ") And (EndDate <= " &  aEmployeeComponent(N_END_DATE_BENEFICIARY_EMPLOYEE) & ")) Or ((EndDate >= " &  aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE) & ") And (StartDate <= " &  aEmployeeComponent(N_END_DATE_BENEFICIARY_EMPLOYEE) & ")))", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aEmployeeComponent(B_IS_DUPLICATED_EMPLOYEE) = True
			Else
				lErrorNumber = L_ERR_NO_RECORDS
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfEmployeeBeneficiary = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfEmployeeCreditors(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific employee exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfEmployeeCreditors"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_JOB)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If Len(aEmployeeComponent(S_NUMBER_EMPLOYEE)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del empleado en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesCreditorsLKP Where (CreditorNumber='" & Replace(aEmployeeComponent(S_NUMBER_CREDITOR_EMPLOYEE), "'", "") & "') And (((StartDate >= " &  aEmployeeComponent(N_START_DATE_CREDITOR_EMPLOYEE) & ") And (EndDate <= " &  aEmployeeComponent(N_END_DATE_CREDITOR_EMPLOYEE) & ")) Or ((EndDate >= " &  aEmployeeComponent(N_START_DATE_CREDITOR_EMPLOYEE) & ") And (EndDate <= " &  aEmployeeComponent(N_END_DATE_CREDITOR_EMPLOYEE) & ")) Or ((EndDate >= " & aEmployeeComponent(N_START_DATE_CREDITOR_EMPLOYEE) & ") And (StartDate <= " &  aEmployeeComponent(N_END_DATE_CREDITOR_EMPLOYEE) & ")))", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aEmployeeComponent(B_IS_DUPLICATED_EMPLOYEE) = True
			Else
				lErrorNumber = L_ERR_NO_RECORDS
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfEmployeeCreditors = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfEmployeeDocument(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific employee document exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfEmployeeDocument"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_JOB)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If Len(aEmployeeComponent(S_NUMBER_EMPLOYEE)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del empleado en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesDocs Where (EmployeeID=" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & ") And (DocumentDate= " &  aEmployeeComponent(N_EMPLOYEE_DOCUMENT_DATE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aEmployeeComponent(B_IS_DUPLICATED_EMPLOYEE) = True
			Else
				lErrorNumber = L_ERR_NO_RECORDS
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfEmployeeDocument = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfEmployeeAbsence(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific employee sunday or extrahour exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfEmployeeAbsence"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	If (aEmployeeComponent(N_ID_EMPLOYEE) < 0) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) < 0) Or (aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) < 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado o el identificador del concepto o la fecha de registro para validar si existe alguna concepto ya registrado."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del empleado en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID From EmployeesAbsencesLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (AbsenceID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (OcurredDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aEmployeeComponent(B_IS_DUPLICATED_EMPLOYEE) = True
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfEmployeeAbsence = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfEmployeeAdjustment(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific employee exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfEmployeeAdjustment"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sConceptShortName

	If aEmployeeComponent(N_ID_EMPLOYEE) < 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del empleado en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAdjustmentsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (MissingDate=" & aEmployeeComponent(N_MISSING_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				Call GetNameFromTable(oADODBConnection, "ShortConcepts", aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE), "", "", sConceptShortName, "")
				sErrorDescription = "Ya existe un registro del concepto " & sConceptShortName & " con fecha de omisión de pago " & DisplayDateFromSerialNumber(aEmployeeComponent(N_MISSING_DATE_EMPLOYEE), -1, -1, -1) & " para el empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & "."
				lErrorNumber = L_ERR_DUPLICATED_RECORD
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfEmployeeAdjustment = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfEmployeeBankAccount(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific bank account exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfEmployeeBankAccount"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim sEmployeeConceptType

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_BANK_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado o el identificador del banco o la fecha de inicio para validar si no existe alguna cuenta ya registrada."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia de la cuenta bancaria en la base de datos."
		sQuery = "Select * From BankAccounts Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
				 " And (((StartDate >= " &  aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (EndDate <= " &  aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & "))" & _
				 " Or ((EndDate >= " &  aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (EndDate <= " &  aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & "))" & _
				 " Or ((EndDate >= " &  aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (StartDate <= " &  aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & "))) Order By StartDate Desc"

		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aEmployeeComponent(N_CONCEPT_CREDIT_TYPE) = 2
				Call GetCrossingEmployeeConceptType(oADODBConnection, aEmployeeComponent, sEmployeeConceptType, lStartDate, lEndDate, sErrorDescription)
				Select Case sEmployeeConceptType
					Case "Left", "Right"
						CheckExistencyOfEmployeeBankAccount = False
					Case "Inner"
						CheckExistencyOfEmployeeBankAccount = False
					Case Else
						sErrorDescription = "No se puede agregar la cuenta bancaria " & aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE) & " del empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " con fecha de inicio " & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1, -1, -1) & " debido a que existe una registrada en el periodo indicado"
						CheckExistencyOfEmployeeBankAccount = True
				End Select
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function CheckExistencyOfEmployeeGrade(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific bank account exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfEmployeeGrade"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim sEmployeeConceptType

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_BANK_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado o el identificador del banco o la fecha de inicio para validar si no existe alguna cuenta ya registrada."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia de la cuenta bancaria en la base de datos."
		sQuery = "Select * From EmployeesGrades Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
				 " And (((StartDate >= " &  aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (EndDate <= " &  aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & "))" & _
				 " Or ((EndDate >= " &  aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (EndDate <= " &  aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & "))" & _
				 " Or ((EndDate >= " &  aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (StartDate <= " &  aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & "))) Order By StartDate Desc"

		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aEmployeeComponent(N_CONCEPT_CREDIT_TYPE) = 3
				Call GetCrossingEmployeeConceptType(oADODBConnection, aEmployeeComponent, sEmployeeConceptType, lStartDate, lEndDate, sErrorDescription)
				Select Case sEmployeeConceptType
					Case "Left", "Right"
						CheckExistencyOfEmployeeGrade = False
					Case "Inner"
						CheckExistencyOfEmployeeGrade = False
					Case Else
						sErrorDescription = "No se puede agregar la calificación del empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " con fecha de inicio " & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1, -1, -1) & " debido a que existe una registrada en el periodo indicado"
						CheckExistencyOfEmployeeGrade = True
				End Select
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function CheckExistencyOfEmployeeGradeXX(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific employee sunday or extrahour exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfEmployeeGrade"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	If (aEmployeeComponent(N_ID_EMPLOYEE) < 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para validar si existe alguna calificación ya registrada."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia de calificación para el empleado en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID From EmployeesGrades Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aEmployeeComponent(B_IS_DUPLICATED_EMPLOYEE) = True
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfEmployeeGrade = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific employee exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfEmployeeID"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized


	If aEmployeeComponent(N_ID_EMPLOYEE) < 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del empleado en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Employees Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				sErrorDescription = "El número de empleado no esta registrado en la base de datos."
				lErrorNumber = L_ERR_NO_RECORDS
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfEmployeeID = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfEmployeeJob(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific employee exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfEmployeeJob"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	If aEmployeeComponent(N_ID_EMPLOYEE) < 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del empleado en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (JobID<>-1)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sErrorDescription = "Este empleado en algún momento tuvo una plaza en el Instituto."
				lErrorNumber = -1
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfEmployeeJob = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfEmployeeRFC(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific employee exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfEmployeeRFC"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim oHistoryRecordset

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_JOB)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If Len(aEmployeeComponent(S_NUMBER_EMPLOYEE)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del empleado por RFC, CURP y tipo de tabulador en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, EmployeeTypeID, Active From Employees Where ((RFC='" & Replace(UCase(aEmployeeComponent(S_RFC_EMPLOYEE)), "'", "") & "') Or (CURP='" & Replace(UCase(aEmployeeComponent(S_CURP_EMPLOYEE)), "'", "") & "')) Or (RFC like '" & Left(Replace(UCase(aEmployeeComponent(S_RFC_EMPLOYEE)), "%'", ""), 10) & "') Or ((EmployeeName='" & Replace(UCase(aEmployeeComponent(S_NAME_EMPLOYEE)), "'", "´") & "') And (EmployeeLastName='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME_EMPLOYEE)), "'", "´") & "') And (EmployeeLastName2='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE)), "'", "´") & "'))", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				If (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 5) Or (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 6) _
					Or (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 7) Then
					Do While Not oRecordset.EOF
						If (CInt(oRecordset.Fields("Active").Value) = 0) Then
							aEmployeeComponent(N_ID_EMPLOYEE) = -1
							oRecordset.MoveNext
						Else
							aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
							Exit Do
						End If
					Loop
				Else
					Do While Not oRecordset.EOF
						If ((CInt(oRecordset.Fields("EmployeeTypeID").Value) = 5) Or _
							(CInt(oRecordset.Fields("EmployeeTypeID").Value) = 6) Or (CInt(oRecordset.Fields("EmployeeTypeID").Value) = 7)) _
							And (CInt(oRecordset.Fields("Active")) = 0) Then
							aEmployeeComponent(N_ID_EMPLOYEE) = -1
							oRecordset.MoveNext
						Else
							aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
							Exit Do
						End If
					Loop
				End If
			Else
				aEmployeeComponent(N_ID_EMPLOYEE) = -1
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfEmployeeRFC = lErrorNumber
	Err.Clear
End Function

Function CheckEmployeeInformationConsistency(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'         going to be added into the database
'Inputs:  aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckEmployeeInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aEmployeeComponent(N_ID_EMPLOYEE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del empleado no es un valor numérico."
		bIsCorrect = False
	End If
	If Len(aEmployeeComponent(S_NUMBER_EMPLOYEE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El número del empleado está vacío."
		bIsCorrect = False
	End If
	If Len(aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La clave de acceso del empleado está vacía."
		bIsCorrect = False
	End If
	If Len(aEmployeeComponent(S_PASSWORD_EMPLOYEE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La contraseña del empleado está vacía."
		bIsCorrect = False
	End If
	If Len(aEmployeeComponent(S_NAME_EMPLOYEE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del empleado está vacío."
		bIsCorrect = False
	End If
	If Len(aEmployeeComponent(S_LAST_NAME_EMPLOYEE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El apellido paterno del empleado está vacío."
		bIsCorrect = False
	End If
	If Not IsNumeric(aEmployeeComponent(N_COMPANY_ID_EMPLOYEE)) Then aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) = -1
	If Not IsNumeric(aEmployeeComponent(N_JOB_ID_EMPLOYEE)) Then aEmployeeComponent(N_JOB_ID_EMPLOYEE) = -1
	If Not IsNumeric(aEmployeeComponent(N_SERVICE_ID_EMPLOYEE)) Then aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) = -1
	If Not IsNumeric(aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE)) Then aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = -1
	If Not IsNumeric(aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE)) Then aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) = -1
	If Not IsNumeric(aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE)) Then aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) = -1
	If Not IsNumeric(aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE)) Then aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) = -1
	If Not IsNumeric(aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE)) Then aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) = -1
	If Not IsNumeric(aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)) Then aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) = -1
	If Not IsNumeric(aEmployeeComponent(N_SHIFT_ID_EMPLOYEE)) Then aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) = -1
	If Not IsNumeric(aEmployeeComponent(N_START_HOUR_1_EMPLOYEE)) Then aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(N_END_HOUR_1_EMPLOYEE)) Then aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(N_START_HOUR_2_EMPLOYEE)) Then aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(N_END_HOUR_2_EMPLOYEE)) Then aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(N_START_HOUR_3_EMPLOYEE)) Then aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(N_END_HOUR_3_EMPLOYEE)) Then aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE)) Then aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(N_LEVEL_ID_EMPLOYEE)) Then aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) = -1
	If Not IsNumeric(aEmployeeComponent(N_STATUS_ID_EMPLOYEE)) Then aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = -2
	If Not IsNumeric(aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE)) Then aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) = -1
	If Not IsNumeric(aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE)) Then aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)) Then aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(N_START_DATE_EMPLOYEE)) Then aEmployeeComponent(N_START_DATE_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(N_START_DATE2_EMPLOYEE)) Then aEmployeeComponent(N_START_DATE2_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE)) Then aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) = -1
	If Len(aEmployeeComponent(S_RFC_EMPLOYEE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El RFC del empleado está vacío."
		bIsCorrect = False
	End If
	If Not IsNumeric(aEmployeeComponent(N_GENDER_ID_EMPLOYEE)) Then aEmployeeComponent(N_GENDER_ID_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE)) Then aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) = 1
	If Not IsNumeric(aEmployeeComponent(N_ACTIVE_EMPLOYEE)) Then aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 1

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del empleado contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckEmployeeInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function CheckEmployeeBeneficiaryInformationConsistency(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'         going to be added into the database
'Inputs:  aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckEmployeeBeneficiaryInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del beneficiario no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE)) Then aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	If Not IsNumeric(aEmployeeComponent(N_END_DATE_BENEFICIARY_EMPLOYEE)) Then aEmployeeComponent(N_END_DATE_BENEFICIARY_EMPLOYEE) = 0
	'If Len(IsNumeric(aEmployeeComponent(S_NUMBER_BENEFICIARY_EMPLOYEE)) = 0 Then 
	'	sErrorDescription = sErrorDescription & "<BR />&nbsp;- El número del beneficiario del empleado está vacío."
	'	bIsCorrect = False
	'End If
	If Len(aEmployeeComponent(S_NAME_BENEFICIARY_EMPLOYEE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del beneficiario del empleado está vacío."
		bIsCorrect = False
	End If
	If Len(aEmployeeComponent(S_LAST_NAME_BENEFICIARY_EMPLOYEE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El apellido paterno del beneficiario del empleado está vacío."
		bIsCorrect = False
	End If
	If Not IsNumeric(aEmployeeComponent(N_BIRTH_DATE_BENEFICIARY_EMPLOYEE)) Then aEmployeeComponent(N_BIRTH_DATE_BENEFICIARY_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(D_ALIMONY_AMOUNT_BENEFICIARY_EMPLOYEE)) Then aEmployeeComponent(D_ALIMONY_AMOUNT_BENEFICIARY_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE)) Then aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE) = 1
	If Not IsNumeric(aEmployeeComponent(N_PAYMENT_CENTER_ID_BENEFICIARY_EMPLOYEE)) Then aEmployeeComponent(N_PAYMENT_CENTER_ID_BENEFICIARY_EMPLOYEE) = -1

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del beneficiario del empleado contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckEmployeeBeneficiaryInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function CheckEmployeeChildInformationConsistency(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'         going to be added into the database
'Inputs:  aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckEmployeeChildInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aEmployeeComponent(N_ID_CHILD_EMPLOYEE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del hijo(a) no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La fecha de nacimiento no es un valor numérico."
		bIsCorrect = False
	Else
		If aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE) = 0 Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- La fecha de nacimiento está vacia."
			bIsCorrect = False
		End If
	End If
	If Len(aEmployeeComponent(S_NAME_CHILD_EMPLOYEE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del hijo(a) del empleado está vacío."
		bIsCorrect = False
	End If
	If Len(aEmployeeComponent(S_LAST_NAME_CHILD_EMPLOYEE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El apellido paterno del hijo(a) del empleado está vacío."
		bIsCorrect = False
	End If

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del hijo(a) del empleado contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckEmployeeChildInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function CheckEmployeeConceptInformationConsistency(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'         going to be added into the database
'Inputs:  aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckEmployeeConceptInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aEmployeeComponent(N_ID_EMPLOYEE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del empleado no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del concepto no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)) Then aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	If Not IsNumeric(aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El monto del concepto no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE)) Then aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE)) Then aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 1
	If Not IsNumeric(aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE)) Then aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 3
	If Not IsNumeric(aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE)) Then aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE)) Then aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) = 1
	If Not IsNumeric(aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE)) Then aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE)) Then aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) = 1
	If Not IsNumeric(aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE)) Then aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) = 1
	If Not IsNumeric(aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE)) Then Call GetNewIDFromTable(oADODBConnection, "EmployeesConceptsLKP", "ConceptOrder", "(EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", 1, aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE), sErrorDescription)
	If Not IsNumeric(aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE)) Then aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 1

	'If Len(sErrorDescription) > 0 Then
	If Not bIsCorrect Then
		sErrorDescription = "La información del empleado contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckEmployeeConceptInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function CheckEmployeeCreditInformationConsistency(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'         going to be added into the database
'Inputs:  aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckEmployeeCreditInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True
	If Not IsNumeric(aEmployeeComponent(N_CREDIT_ID_EMPLOYEE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del crédito no es un valor numérico."
		bIsCorrect = False
	End If
	'If Len(aEmployeeComponent(S_CREDIT_CONTRACT_NUMBER_EMPLOYEE)) = 0 Then
	'	sErrorDescription = sErrorDescription & "<BR />&nbsp;- El número de contrato del empleado está vacío."
	'	bIsCorrect = False
	'End If
	'If Len(aEmployeeComponent(S_CREDIT_ACCOUNT_NUMBER_EMPLOYEE)) = 0 Then
	'	sErrorDescription = sErrorDescription & "<BR />&nbsp;- El número de cuenta del empleado está vacío."
	'	bIsCorrect = False
	'End If
	If Not IsNumeric(aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El número de pagos del crédito no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aEmployeeComponent(N_CREDIT_PERIOD_ID_EMPLOYEE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El número de periodo está vacio."
		bIsCorrect = False
	End If
	If Not IsNumeric(aEmployeeComponent(L_CREDIT_FINISH_DATE_EMPLOYEE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La fecha de termino no es un valor numérico."
		bIsCorrect = False
	Else
		If aEmployeeComponent(L_CREDIT_FINISH_DATE_EMPLOYEE) = 0 Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- La fecha de termino está vacia."
			bIsCorrect = False
		End If
	End If
	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del crédito del empleado contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckEmployeeCreditInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function CheckEmployeeCreditorInformationConsistency(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'         going to be added into the database
'Inputs:  aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckEmployeeCreditorInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del acreedor no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aEmployeeComponent(N_START_DATE_CREDITOR_EMPLOYEE)) Then aEmployeeComponent(N_START_DATE_CREDITOR_EMPLOYEE) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	If Not IsNumeric(aEmployeeComponent(N_END_DATE_CREDITOR_EMPLOYEE)) Then aEmployeeComponent(N_END_DATE_CREDITOR_EMPLOYEE) = 0
	If Len(aEmployeeComponent(S_NAME_CREDITOR_EMPLOYEE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del acreedor del empleado está vacío."
		bIsCorrect = False
	End If
	If Len(aEmployeeComponent(S_LAST_NAME_CREDITOR_EMPLOYEE)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El apellido paterno del acreedor del empleado está vacío."
		bIsCorrect = False
	End If
	If Not IsNumeric(aEmployeeComponent(N_BIRTH_DATE_CREDITOR_EMPLOYEE)) Then aEmployeeComponent(N_BIRTH_DATE_CREDITOR_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(D_CREDITOR_AMOUNT_EMPLOYEE)) Then aEmployeeComponent(D_CREDITOR_AMOUNT_EMPLOYEE) = 0
	If Not IsNumeric(aEmployeeComponent(N_CREDITOR_TYPE_ID_EMPLOYEE)) Then aEmployeeComponent(N_CREDITOR_TYPE_ID_EMPLOYEE) = 1
	If Not IsNumeric(aEmployeeComponent(N_PAYMENT_CENTER_ID_CREDITOR_EMPLOYEE)) Then aEmployeeComponent(N_PAYMENT_CENTER_ID_CREDITOR_EMPLOYEE) = -1

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del beneficiario del empleado contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckEmployeeCreditorInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function CheckEmployeeStatus(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific employee is in the process of
'         of staff turnover
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckEmployeeStatus"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_JOB)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If Len(aEmployeeComponent(S_NUMBER_EMPLOYEE)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del empleado en la base de datos."
		If (aEmployeeComponent(N_STATUS_ID_EMPLOYEE) <> 0) And (aEmployeeComponent(N_STATUS_ID_EMPLOYEE) <> 1) Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusName, ReasonID From StatusEmployees Where (StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					If aEmployeeComponent(N_REASON_ID_EMPLOYEE) <> CLng(oRecordset.Fields("ReasonID").Value) Then
						lErrorNumber = -1
					End If
				End If
			End If
		Else 
			lErrorNumber = 0
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckEmployeeStatus = lErrorNumber
	Err.Clear
End Function

Function CheckRequirementsOfEmployeeMovement(oRequest, EmployeeComponent, lReasonID, sErrorDescription)
'************************************************************
'Purpose: To verify that an employee meets the requirements to effect movement
'Inputs:  aEmployeeComponent, lReasonID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckRequirementsOfEmployeeMovement"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sRequirementsIDs
	Dim lStartDate
	Dim iTotalDays
	Dim dStartDate
	Dim dEndDate
	Dim lDiffDate
	Dim lAntiquity
	Dim lPeriod
	Dim lCount
	Dim lCurrentDate
	Dim sHolidays

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		dStartDate = oRequest("EmployeeYear").Item & oRequest("EmployeeMonth").Item & oRequest("EmployeeDay").Item
		dEndDate = oRequest("EmployeeEndYear").Item & oRequest("EmployeeEndMonth").Item & oRequest("EmployeeEndDay").Item
		If (CInt(oRequest("ReasonID").Item) = 31) Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "SELECT Count(EmployeeID) as lPrevious From EmployeesHistoryList WHERE (EmployeeID = " & oRequest("EmployeeID").Item & ") And reasonID = 31", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If (CLng(oRecordset.Fields("lPrevious").Value) > 0) Then
				lErrorNumber = -1
				sErrorDescription = "El empleado ya ha hizo uso de esta licencia, el trámite no puede continuar."
			End If
		End If
		If (CInt(oRequest("ReasonID").Item) >= 29 And CInt(oRequest("ReasonID").Item) <= 34) Then
			If (lErrorNumber = 0) Then
				lPeriod = DateDiff("d", GetDateFromSerialNumber(dStartDate), GetDateFromSerialNumber(dEndDate)) + 1
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Holiday From Holidays", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				If (lErrorNumber = 0) And (CInt(oRequest("ReasonID").ITem) <> 33) Then
					sHolidays = ","
					Do While Not oRecordset.EOF
						sHolidays = sHolidays & CStr(oRecordset.Fields("Holiday").Value) & ","
						oRecordset.MoveNext
					Loop
					For lCount = getDateFromSerialNumber(dStartDate) to getDateFromSerialNumber(dEndDate)
						lCurrentDate = getSerialNumberForDate(lCount)
						If (InStr(1, sHolidays, "," & Mid(lCurrentDate, 1, 8) & ",", vbBinaryCompare) > 0) And (Weekday(lCount) <> 1) And (Weekday(lCount) <> 7) Then
							lPeriod = lPeriod - 1
						End If
						If (Weekday(lCount) = 1) Or (Weekday(lCount) = 7) Then
							lPeriod = lPeriod - 1
						End If
					Next
				End If
			End If
			If (lErrorNumber = 0) Then 
				If (CInt(oRequest("ReasonID").Item) = 29) And (lPeriod > 368) Then
					lErrorNumber = -1
					sErrorDescription = "La vigencia máxima permitida para la icencia con goce de sueldo por Comisión sindical es de 1 año, verifique la fecha final"
				ElseIf (CInt(oRequest("ReasonID").Item) = 30) And (lPeriod > 90) Then
					lErrorNumber = -1
					sErrorDescription = "La vigencia máxima permitida para la Licencia con goce de sueldo por trámite de pensión es de 90 días hábiles, verifique la fecha final"
				ElseIf (CInt(oRequest("ReasonID").Item) = 31) And (lPeriod > 10) Then
					lErrorNumber = -1
					sErrorDescription = "La vigencia máxima permitida para la Licencia con goce de sueldo por contraer matrimonio es de 10 días hábiles, verifique la fecha final"
				ElseIf (CInt(oRequest("ReasonID").Item) = 32) And (lPeriod > 5) Then
					lErrorNumber = -1
					sErrorDescription = "La vigencia máxima permitida para la Licencia con goce de sueldo por fallecimiento de familiar en primer grado es de 5 días hábiles, verifique la fecha final"
				ElseIf (CInt(oRequest("ReasonID").Item) = 33) And (lPeriod > 368) Then
					lErrorNumber = -1
					sErrorDescription = "La vigencia máxima permitida para la Licencia con goce de sueldo por otorgamiento de beca es de 1 año, verifique la fecha final"
				ElseIf (CInt(oRequest("ReasonID").Item) = 34) And (lPeriod > 180) Then
					lErrorNumber = -1
					sErrorDescription = "La vigencia máxima permitida para la Licencia con goce de sueldo por práctica de servicio social es de 180 días, verifique la fecha final"
				End If
			End If
		End If
'		If (CInt(oRequest("ReasonID").Item) >= 29) And (CInt(oRequest("ReasonID").Item)<=48) And (lErrorNumber = 0) Then
'			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "SELECT AbsenceID, OcurredDate, EndDate FROM EmployeesAbsencesLKP WHERE (EmployeeID = " & oRequest("EmployeeID").Item & ") And (OcurredDate>=" & dStartDate & ") And (OcurredDate<=" & dEndDate & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
'			If (NOT oRecordset.EOF) Then
'				sErrorDescription = "La operación no pudo completarse por detectarse incidencias registradas en el periodo de la licencia."
'				lErrorNumber = -1
'			End If
'		End If
		If (lErrorNumber = 0) And CInt(oRequest("ReasonID").Item) = 43 Then
			If dEndDate = "000" Then
				lErrorNumber = -1
				sErrorDescription = "La vigencia del movimiento debe tener una fecha de término"
			End If
			If lErrorNumber = 0 Then
				lAntiquity = (CInt(oRequest("AntiquityYears").Item) * 365) + (CInt(oRequest("AntiquityMonths").Item) * 30) + CInt(oRequest("AntiquityDays").Item)
				lPeriod = dateDiff("d", getDateFromSerialNumber(dStartDate),getDateFromSerialNumber(dEndDate))
				If (lAntiquity < 180) Then
					lErrorNumber = -1
					sErrorDescription = "El empleado no tiene la antigüedad requerida para esta licencia"
				End If
				If lErrorNumber = 0 Then
					If ((lAntiquity >= 180 And lAntiquity < 365) And lPeriod > 30) Or ((lAntiquity >= 365 And lAntiquity < 1095) And lPeriod > 90) Or (lAntiquity >= 1095 And lPeriod > 180) Then
						lErrorNumber = -1
						sErrorDescription = "La vigencia del movimiento no corresponde con la antigüedad del empleado"
					End If
				End If
			End If
		End If
		If (lErrorNumber = 0) And (CInt(oRequest("ReasonID").Item) = 33) Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeDate, EndDate From EmployeesHistoryList Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " And ReasonID = 33 order by 1 Desc" , "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If (lErrorNumber = 0) Then
				lPeriod = 0
				Do While Not oRecordset.EOF
					dStartDate = GetDateFromSerialNumber(oRecordset.Fields("EmployeeDate").Value)
					dEndDate = GetDateFromSerialNumber(oRecordset.Fields("EndDate").Value)
					lPeriod = lPeriod + DateDiff("d", dStartDate, dEndDate)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				dStartDate = GetDateFromSerialNumber(oRequest("EmployeeYear").Item & oRequest("EmployeeMonth").Item & oRequest("EmployeeDay").Item)
				dEndDate = GetDateFromSerialNumber(oRequest("EmployeeEndYear").Item & oRequest("EmployeeEndMonth").Item & oRequest("EmployeeEndDay").Item)
				lPeriod = lPeriod + dateDiff("d",dStartDate,dEndDate)
				If (lPeriod > 2555) Then
					lErrorNumber = -1
					sErrorDescription = "El empleado ha acumulado más de 7 años de licencias con goce de sueldo por otorgamiento de becas"
				End If
			End If
		End If
		If (lErrorNumber = 0) And (CInt(oRequest("ReasonID").Item) = 13) Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusID From Jobs Where JobID = " & aEmployeeComponent(N_JOB_ID_EMPLOYEE), "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If (lErrorNumber = 0) Then
			End If
			If (lErrorNumber = 0) Then
				dStartDate = GetDateFromSerialNumber(oRequest("EmployeeYear").Item & oRequest("EmployeeMonth").Item & oRequest("EmployeeDay").Item)
				dEndDate = GetDateFromSerialNumber(oRequest("EmployeeEndYear").Item & oRequest("EmployeeEndMonth").Item & oRequest("EmployeeEndDay").Item)
				lPeriod = dateDiff("m",dStartDate,dEndDate)
				If (lPeriod > 6) Then
					lErrorNumber = -1
					sErrorDescription = "Los interinatos no pueden tener una duración mayor a 6 meses"
				End If
			End If
			If (lErrorNumber = 0) Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusID From Jobs Where JobID = " & aEmployeeComponent(N_JOB_ID_EMPLOYEE), "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				If (lErrorNumber = 0) Then
					If CInt(oRecordset.Fields("StatusId").Value) = 2 Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "select JobDate, EndDate from JobsHistoryList Where JobID = " & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & " And JobDate >= (select max(enddate) from EmployeesHistoryList where ReasonID in (1,2,3,4,5,6,63) and JobID = " & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & " and EndDate <> 30000000) And StatusID = 1 order by 1 Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						If (lErrorNumber = 0) Then
							If Not oRecordset.EOF Then
								lPeriod = 0
								Do While Not oRecordset.EOF
									dStartDate = GetDateFromSerialNumber(oRecordset.Fields("JobDate").Value)
									dEndDate = GetDateFromSerialNumber(oRecordset.Fields("EndDate").Value)
									lPeriod = CLng(dateDiff("d",dStartDate,dEndDate)) + lPeriod
									oRecordset.MoveNext
								Loop
								dStartDate = GetDateFromSerialNumber(oRequest("EmployeeYear").Item & oRequest("EmployeeMonth").Item & oRequest("EmployeeDay").Item)
								dEndDate = GetDateFromSerialNumber(oRequest("EmployeeEndYear").Item & oRequest("EmployeeEndMonth").Item & oRequest("EmployeeEndDay").Item)
								lPeriod = dateDiff("d",dStartDate,dEndDate) + lPeriod
								If lPeriod > 90 Then
									lErrorNumber = -1	
									sErrorDescription = "La plaza indicada acumula más de 90 días en interinatos."
								End If
							End If
						End If
					End If
				End If
			End If
			If (lErrorNumber = 0) Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "select employeedate, enddate from EmployeesHistoryList where (ReasonId = 13) And (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") order by 1 Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				If Not oRecordset.EOF Then
					dStartDate = GetDateFromSerialNumber(oRecordset.Fields("EndDate").Value)
					dEndDate = GetDateFromSerialNumber(oRequest("EmployeeEndYear").Item & oRequest("EmployeeEndMonth").Item & oRequest("EmployeeEndDay").Item)
					lPeriod = dateDiff("d",dStartDate,dEndDate)
					If (lPeriod < 15) Then
						lErrorNumber = -1
						sErrorDescription = "No han transucrrido 15 días entre el último interinato y el que se está registrando"
					End If
				End If
			End If
		End If
		If (lErrorNumber = 0) And (oRequest("ReasonID").Item = 30) Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeDate, EndDate From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (StatusID=58) And  (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			if lErrorNumber = 0 then
				iTotalDays = 0
				Do While Not oRecordset.EOF
					dStartDate = GetDateFromSerialNumber(oRecordset.Fields("EmployeeDate").Value)
					dEndDate = GetDateFromSerialNumber(oRecordset.Fields("EndDate").Value)
					lDiffDate = CLng(dateDiff("d",dStartDate,dEndDate))
					iTotalDays = iTotalDays + CLng(lDiffDate)
					oRecordset.MoveNext
				Loop
				dStartDate = GetDateFromSerialNumber(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE))
				dEndDate = GetDateFromSerialNumber(aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE))
				iTotalDays = iTotalDays + CLng(dateDiff("d",dStartDate,dEndDate))
				oRecordset.Close
				if CLng(iTotalDays) > 91 then
					sErrorDescription = "El plazo de la vigencia ha sobrepasado el límite de 3 meses 1 día"
					lErrorNumber = -1
				End If
			End If
		End If
		If lErrorNumber = 0 Then
			If oRequest("ReasonID").Item = 33 Or oRequest("ReasonID").Item = 29 Then
				If dEndDate = "000" Then
					lErrorNumber = -1
					sErrorDescription = "La fecha final de la vigencia es requerida"
				Else 
					If oRequest("ReasonID").Item = 33 And CInt(DateDiff("d",GetDateFromSerialNumber(dStartDate),GetDateFromSerialNumber(dEndDate))) > 1095 Then
						lErrorNumber = -1
						sErrorDescription = "La vigencia sobrepasa el tiempo máximo permitido para esta licencia."
					End If
					If oRequest("ReasonID").Item = 29 And CInt(DateDiff("d",GetDateFromSerialNumber(dStartDate),GetDateFromSerialNumber(dEndDate))) > 365 Then
						lErrorNumber = -1
						sErrorDescription = "La vigencia sobrepasa el tiempo máximo permitido para esta licencia."
					End If					
				End If
			End If
		End If
		If lErrorNumber = 0 Then
			If (lReasonID = 51) And (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 7) Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeDate, EndDate From EmployeesHistoryList Where EmployeeId = " & aEmployeeComponent(N_ID_EMPLOYEE) & " Order By EmployeeDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				If oRecordset.EOF Then
					lErrorNumber = -1
					sErrorDescription = "No se ha encontrado la información del contrato de honorarios del empleado indicado"
				Else
					dStartDate = CLng(oRequest("EmployeeYear").Item & oRequest("EmployeeMonth").Item & oRequest("EmployeeDay").Item)
					dEndDate = CLng(oRequest("EmployeeEndYear").Item & oRequest("EmployeeEndMonth").Item & oRequest("EmployeeEndDay").Item)
					If dEndDate = 0 Then
						lErrorNumber = -1
						sErrorDescription = "La fecha final de vigencia es necesaria para los empleados por honorarios"
					Else
						If (dStartDate < oRecordset.Fields("EmployeeDate").Value) Or (dStartDate > oRecordset.Fields("EndDate").Value) Then
							lErrorNumber = -1
							sErrorDescription = "La fecha inicial de la vigencia está fuera de los límites del contrato actual de honorarios" & ": " & dStartDate & ", " & dEndDate & ", " & oRecordset.Fields("EmployeeDate").Value & ", " & oRecordset.Fields("EndDate").Value
						End If
						If lErrorNumber = 0 Then
							If (dEndDate < oRecordset.Fields("EmployeeDate").Value) Or (dEndDate > oRecordset.Fields("EndDate").Value) Then
								lErrorNumber = -1
								sErrorDescription = "La fecha final de la vigencia está fuera de los límites del contrato actual de honorarios" & ": " & dStartDate & ", " & dEndDate & ", " & oRecordset.Fields("EmployeeDate").Value & ", " & oRecordset.Fields("EndDate").Value
							End If
						End If
					End If
				End If
			End If
		End If
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sErrorDescription = ""
				sRequirementsIDs = "," & CStr(oRecordset.Fields("ReasonRequirementIDs").Value) & ","
				If (InStr(1, sRequirementsIDs, ",5,", vbBinaryCompare) > 0) Then
					If (aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 0) Or (aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 30000000) Then
						sErrorDescription = "La fecha fin del movimiento es requerida" & "<BR />"
						lErrorNumber = -1
					End If
				End If
				If lErrorNumber = 0 Then
					If (InStr(1, sRequirementsIDs, ",7,", vbBinaryCompare) > 0) Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (bProcessed<>2) And (ReasonID<>0) And (ReasonID<>58) And (ReasonID<>28) And (EmployeeDate>=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								sErrorDescription = "La fecha de inicio del movimiento no puede ser menor a la registrada en el sistema: " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value))
								If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
									sErrorDescription = sErrorDescription & " a Indefinida" & "<BR />"
								Else
									sErrorDescription = sErrorDescription & " al " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & "<BR />"
								End If
								lErrorNumber = -1
							End If
						End If
					End If
					If lErrorNumber = 0 Then
						If (InStr(1, sRequirementsIDs, ",2,", vbBinaryCompare) > 0) Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesHistoryList Where (JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ") And (bProcessed<>2) And (Active=1) Order By EndDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									If (CLng(oRecordset.Fields("EndDate").Value) >= aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)) Then
										sErrorDescription = "Verifique la fecha de inicio de vigencia debido a que la plaza se encontró ocupada por el empleado: " & Right("000000" & CLng(oRecordset.Fields("EmployeeID").Value), Len("000000"))
										lErrorNumber = -1
									End If
								End If
							End If
						End If
						If lErrorNumber = 0 Then
							iTotalDays = 0
							if (InStr(1,sRequirementsIDs, ",8,", vbBinaryCompare) > 0) then
								lErrorNumber = ExecuteSQLQuery(oADODBConnectio, "Select EmployeeDate, EndDate From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ") And  (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								if lErrorNumber = 0 then
									Do While Not oRecordset.EOF
										dStartDate = GetDateFromSerialNumber(oRecordset.Fields("EmployeeDate").Value)
										dEndDate = GetDateFromSerialNumber(oRecordSet.Fields("EndDate").Value)
										dDiffDate = dateDiff("d",dStartDate,dEndDate)
										iTotalDays = iTotalDays + dDiffDate
										oRecordset.MoveNext
										If Err.number <> 0 Then Exit Do
									Loop
									oRecordset.Close
									if iTotalDays < 91 then
										sErrorDescription = "La vigencia del movimiento no puede ser mayor a 3 meses más 1 día"
										lErrorNumber = -1
									end if
								End if
							End if
						end if
						if lErrorNumber = 0 then
							If (InStr(1, sRequirementsIDs, ",6,", vbBinaryCompare) > 0) Then
								lStartDate = CLng(AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), -180))
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (bProcessed<>2) And (ReasonID=" & lReasonID & ") And (EndDate>" & lStartDate & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										sErrorDescription = "El empleado no tiene derecho a tomar otra licencia hasta que transcurran 6 meses a partir de su última licencia." & "<BR />"
										lErrorNumber = -1
									End If
								End If
							End If
							If lErrorNumber = 0 Then
								If (InStr(1, sRequirementsIDs, ",1,", vbBinaryCompare) > 0) Then
									lStartDate = CLng(AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), 165))
									If lStartDate < aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) Then
										sErrorDescription = "La vigencia del movimiento no puede ser mayor que 5 meses 15 días" & "<BR />"
										lErrorNumber = -1
									End If
								End If
								If lErrorNumber = 0 Then
									If (InStr(1, sRequirementsIDs, ",2,", vbBinaryCompare) > 0) Then
										aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
										lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
										If lErrorNumber = 0 Then
											If aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 0 Then aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 30000000
											If aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) < aJobComponent(N_START_DATE_JOB) Then
												sErrorDescription = sErrorDescription & "Por favor verifique la vigencia del movimiento debido a que la fecha de inicio del movimiento es menor a la fecha de inicio de la plaza." & "<BR />"
												lErrorNumber = -1
											ElseIf aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) > aJobComponent(N_END_DATE_JOB) Then
												sErrorDescription = sErrorDescription & "Por favor verifique que la vigencia del movimiento debido a que la fecha de fin del movimientos es mayor a la fecha de fin de la plaza." & "<BR />"
												lErrorNumber = -1
											ElseIf aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) > aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) Then
												sErrorDescription = sErrorDescription & "Por favor verifique la vigencia del movimiento" & "<BR />"
												lErrorNumber = -1
											End If
										End If
										If lErrorNumber = 0 Then
											If lReasonID = 13 Then
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From JobsHistoryList Where (JobID=" & aJobComponent(N_ID_JOB) & ") And (JobDate<=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ")  And (EndDate>=" & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ") And (StatusID In (2, 4))", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
												If lErrorNumber = 0 Then
													If oRecordset.EOF Then
														sErrorDescription = sErrorDescription & "Por favor verifique la vigencia del movimiento debido a que la plaza no se encuentra en estatus de vacante ni de licencia en el período señalado." & "<BR />"
														lErrorNumber = -1
													End If
												End If
											End If
										End If
									End If
									If lErrorNumber = 0 Then
										If (InStr(1, sRequirementsIDs, ",4,", vbBinaryCompare) > 0) Then
											lStartDate = CLng(AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), 180))
											If lStartDate < aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) Then
												sErrorDescription = "La vigencia no puede ser mayor que 6 meses" & "<BR />"
												lErrorNumber = -1
											End If
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			Else
				sErrorDescription = "No tiene requisitos este movimiento"
			End If
		End If
	End If

	Set oRecordset = Nothing
	CheckRequirementsOfEmployeeMovement = lErrorNumber
	Err.Clear
End Function

Function TransformXMLTagsForEmployeeForm(aEmployeeComponent, bFull, sText, sErrorDescription)
'************************************************************
'Purpose: To replace the XML tags using the entries from the
'         database
'Inputs:  aEmployeeComponent, bFull
'Outputs: sText, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "TransformXMLTagsForEmployeeForm"
	Dim lFormID
	Dim iStartPos
	Dim iMidPos
	Dim iEndPos
	Dim lDate
	Dim sFormFieldName
	Dim asFields
	Dim sAnswer
	Dim sFormAnswers
	Dim lPayrollID
	Dim sConceptShortName
	Dim sAccessKey
	Dim sPassword
	Dim lPreviousEmployeeID
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
	sText = Replace(sText, "<EMPLOYEE_DATE />", aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE))

	sText = Replace(sText, "<EMPLOYEE_ID />", aEmployeeComponent(N_ID_EMPLOYEE), 1, -1, vbBinaryCompare)
	If InStr(1, sText, " />", vbBinaryCompare) > 0 Then
		sErrorDescription = "No se pudo obtener la información del empleado."
		If bFull Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.EmployeeNumber, Employees.EmployeeName, Employees.EmployeeLastName, Employees.EmployeeLastName2, EmployeesHistoryList.*, LevelName, PositionTypeShortName, JobNumber, ServiceShortName, ServiceName, JobTypes.JobTypeID, JobTypes.JobTypeShortName, OccupationTypes.OccupationTypeID, OccupationTypes.OccupationTypeShortName, Jobs.StartDate As JobStartDate, Jobs.EndDate As JobEndDate, StatusJobs.StatusID As JobStatusID, StatusJobs.StatusShortName As JobStatusShortName, StatusJobs.StatusName As JobStatusName, Areas.*, Positions.*, EconomicZones.EconomicZoneCode, MaritalStatusName From Employees, EmployeesHistoryList, Levels, PositionTypes, Services, Jobs, JobTypes, OccupationTypes, StatusJobs, Areas, Positions, EconomicZones, MaritalStatus Where (Employees.LevelID=Levels.LevelID) And (Employees.PositionTypeID=PositionTypes.PositionTypeID) And (Employees.ServiceID=Services.ServiceID) And (Employees.JobID=Jobs.JobID) And (Jobs.JobTypeID=JobTypes.JobTypeID) And (Jobs.OccupationTypeID=OccupationTypes.OccupationTypeID) And (Jobs.StatusID=StatusJobs.StatusID) And (Jobs.AreaID=Areas.AreaID) And (Jobs.PositionID=Positions.PositionID) And (Areas.EconomicZoneID=EconomicZones.EconomicZoneID) And (Employees.MaritalStatusID=MaritalStatus.MaritalStatusID) And (Employees.EmployeeID = EmployeesHistoryList.EmployeeID) And (Employees.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Employees Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		End If
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sText = Replace(sText, "<EMPLOYEE_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)), 1, -1, vbBinaryCompare)
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
					sText = Replace(sText, "<EMPLOYEE_FULL_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value)), 1, -1, vbBinaryCompare)
				Else
					sText = Replace(sText, "<EMPLOYEE_FULL_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName").Value)), 1, -1, vbBinaryCompare)
				End If
				sText = Replace(sText, "<EMPLOYEE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_LAST_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)), 1, -1, vbBinaryCompare)
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
					sText = Replace(sText, "<EMPLOYEE_LAST_NAME2 />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)), 1, -1, vbBinaryCompare)
				Else
					sText = Replace(sText, "<EMPLOYEE_LAST_NAME2 />", " ", 1, -1, vbBinaryCompare)
				End If
				sText = Replace(sText, "<EMPLOYEE_COMPANY_ID />", CStr(oRecordset.Fields("CompanyID").Value), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_JOB_ID />", CleanStringForHTML(CStr(oRecordset.Fields("JobID").Value)), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_SERVICE_ID />", CleanStringForHTML(CStr(oRecordset.Fields("ServiceID").Value)), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_TYPE_ID />", CStr(oRecordset.Fields("EmployeeTypeID").Value), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_POSITION_TYPE_ID />", CStr(oRecordset.Fields("PositionTypeID").Value), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_GROUP_GRADE_LEVEL_ID />", CStr(oRecordset.Fields("GroupGradeLevelID").Value), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_JOURNEY_ID />", CStr(oRecordset.Fields("JourneyID").Value), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_SHIFT_ID />", CStr(oRecordset.Fields("ShiftID").Value), 1, -1, vbBinaryCompare)
				If CLng(oRecordset.Fields("StartHour1").Value) = 0 Then
					sText = Replace(sText, "<EMPLOYEE_START_HOUR_1 />", "", 1, -1, vbBinaryCompare)
				Else
					sText = Replace(sText, "<EMPLOYEE_START_HOUR_1 />", Right(("0000" & CStr(oRecordset.Fields("StartHour1").Value)), Len("0000")), 1, -1, vbBinaryCompare)
				End If
				If CLng(oRecordset.Fields("EndHour1").Value) = 0 Then
					sText = Replace(sText, "<EMPLOYEE_END_HOUR_1 />", "", 1, -1, vbBinaryCompare)
				Else
					sText = Replace(sText, "<EMPLOYEE_END_HOUR_1 />", Right(("0000" & CStr(oRecordset.Fields("EndHour1").Value)), Len("0000")), 1, -1, vbBinaryCompare)
				End If
				If CLng(oRecordset.Fields("StartHour2").Value) = 0 Then
					sText = Replace(sText, "<EMPLOYEE_START_HOUR_2 />", "", 1, -1, vbBinaryCompare)
				Else
					sText = Replace(sText, "<EMPLOYEE_START_HOUR_2 />", Right(("0000" & CStr(oRecordset.Fields("StartHour2").Value)), Len("0000")), 1, -1, vbBinaryCompare)
				End If
				If CLng(oRecordset.Fields("EndHour2").Value) = 0 Then
					sText = Replace(sText, "<EMPLOYEE_END_HOUR_2 />", "", 1, -1, vbBinaryCompare)
				Else
					sText = Replace(sText, "<EMPLOYEE_END_HOUR_2 />", Right(("0000" & CStr(oRecordset.Fields("EndHour2").Value)), Len("0000")), 1, -1, vbBinaryCompare)
				End If
				If CLng(oRecordset.Fields("StartHour3").Value) = 0 Then
					sText = Replace(sText, "<EMPLOYEE_START_HOUR_3 />", "", 1, -1, vbBinaryCompare)
				Else
					sText = Replace(sText, "<EMPLOYEE_START_HOUR_3 />", Right(("0000" & CStr(oRecordset.Fields("StartHour3").Value)), Len("0000")), 1, -1, vbBinaryCompare)
				End If
				If CLng(oRecordset.Fields("EndHour3").Value) = 0 Then
					sText = Replace(sText, "<EMPLOYEE_END_HOUR_3 />", "", 1, -1, vbBinaryCompare)
				Else
					sText = Replace(sText, "<EMPLOYEE_END_HOUR_3 />", Right(("0000" & CStr(oRecordset.Fields("EndHour3").Value)), Len("0000")), 1, -1, vbBinaryCompare)
				End If
				sText = Replace(sText, "<EMPLOYEE_LEVEL_ID />", CStr(oRecordset.Fields("LevelID").Value), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_STATUS_ID />", CStr(oRecordset.Fields("StatusID").Value), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_PAYMENT_CENTER_ID />", CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterID").Value)), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_EMAIL />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeEmail").Value)), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_SSN />", CleanStringForHTML(CStr(oRecordset.Fields("SocialSecurityNumber").Value)), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_BIRTH_DATE />", DisplayDateFromSerialNumber(CLng(oRecordset.Fields("BirthDate").Value), -1, -1, -1), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_AGE />", DateDiff("yyyy", DateSerial(CInt(oRecordset.Fields("BirthYear").Value), CInt(oRecordset.Fields("BirthMonth").Value), CInt(oRecordset.Fields("BirthDay").Value)), Date), 1, -1, vbBinaryCompare)
				If CLng(oRecordset.Fields("StartDate").Value) > 0 Then
					sText = Replace(sText, "<EMPLOYEE_START_YEAR />", Left(CStr(oRecordset.Fields("StartDate").Value), Len("0000")), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<EMPLOYEE_START_MONTH />", Mid(CStr(oRecordset.Fields("StartDate").Value), Len("00000"), Len("00")), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<EMPLOYEE_START_DAY />", Right(CStr(oRecordset.Fields("StartDate").Value), Len("00")), 1, -1, vbBinaryCompare)
				Else
					sText = Replace(sText, "<EMPLOYEE_START_YEAR />", "0000", 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<EMPLOYEE_START_MONTH />", "00", 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<EMPLOYEE_START_DAY />", "00", 1, -1, vbBinaryCompare)
				End If
				If (CLng(oRecordset.Fields("EndDate").Value) > 0) And (CLng(oRecordset.Fields("EndDate").Value) < 30000000) Then
					sText = Replace(sText, "<EMPLOYEE_END_YEAR />", Left(CStr(oRecordset.Fields("EndDate").Value), Len("0000")), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<EMPLOYEE_END_MONTH />", Mid(CStr(oRecordset.Fields("EndDate").Value), Len("00000"), Len("00")), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<EMPLOYEE_END_DAY />", Right(CStr(oRecordset.Fields("EndDate").Value), Len("00")), 1, -1, vbBinaryCompare)
				Else
					sText = Replace(sText, "<EMPLOYEE_END_YEAR />", "9999", 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<EMPLOYEE_END_MONTH />", "99", 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<EMPLOYEE_END_DAY />", "99", 1, -1, vbBinaryCompare)
				End If
				sText = Replace(sText, "<EMPLOYEE_COUNTRY_ID />", CStr(oRecordset.Fields("CountryID").Value), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_RFC />", CStr(oRecordset.Fields("RFC").Value), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_CURP />", CStr(oRecordset.Fields("CURP").Value), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_GENDER_ID />", CStr(oRecordset.Fields("GenderID").Value), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<GENDER_SHORT_NAME />", Replace(Replace(CStr(oRecordset.Fields("GenderID").Value), "0", "M"), "1", "H"), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_MARITAL_STATUS_ID />", CStr(oRecordset.Fields("MaritalStatusID").Value), 1, -1, vbBinaryCompare)
				If bFull Then
					sText = Replace(sText, "<EMPLOYEE_LEVEL_NAME />", Left(Right(("000" & CStr(oRecordset.Fields("LevelName").Value)), Len("000")), Len("00")), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<EMPLOYEE_SUBLEVEL_NAME />", Right(CStr(oRecordset.Fields("LevelName").Value), Len("0")), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<SERVICE_SHORT_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("ServiceShortName").Value)), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<SERVICE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("ServiceName").Value)), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<JOB_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("JobNumber").Value)), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<JOB_TYPE_SHORT_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("JobTypeShortName").Value)), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<OCCUPATION_TYPE_SHORT_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("OccupationTypeShortName").Value)), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<JOB_STATUS_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("JobStatusName").Value)), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<JOB_STATUS_SHORT_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("JobStatusShortName").Value)), 1, -1, vbBinaryCompare)
					If CLng(oRecordset.Fields("JobStartDate").Value) > 0 Then
						sText = Replace(sText, "<JOB_START_YEAR />", Left(CStr(oRecordset.Fields("JobStartDate").Value), Len("0000")), 1, -1, vbBinaryCompare)
						sText = Replace(sText, "<JOB_START_MONTH />", Mid(CStr(oRecordset.Fields("JobStartDate").Value), Len("00000"), Len("00")), 1, -1, vbBinaryCompare)
						sText = Replace(sText, "<JOB_START_DAY />", Right(CStr(oRecordset.Fields("JobStartDate").Value), Len("00")), 1, -1, vbBinaryCompare)
					Else
						sText = Replace(sText, "<JOB_START_YEAR />", "0000", 1, -1, vbBinaryCompare)
						sText = Replace(sText, "<JOB_START_MONTH />", "00", 1, -1, vbBinaryCompare)
						sText = Replace(sText, "<JOB_START_DAY />", "00", 1, -1, vbBinaryCompare)
					End If
					If (CLng(oRecordset.Fields("JobEndDate").Value) > 0) And (CLng(oRecordset.Fields("JobEndDate").Value) < 30000000) Then
						sText = Replace(sText, "<JOB_END_YEAR />", Left(CStr(oRecordset.Fields("JobEndDate").Value), Len("0000")), 1, -1, vbBinaryCompare)
						sText = Replace(sText, "<JOB_END_MONTH />", Mid(CStr(oRecordset.Fields("JobEndDate").Value), Len("00000"), Len("00")), 1, -1, vbBinaryCompare)
						sText = Replace(sText, "<JOB_END_DAY />", Right(CStr(oRecordset.Fields("JobEndDate").Value), Len("00")), 1, -1, vbBinaryCompare)
					Else
						sText = Replace(sText, "<JOB_END_YEAR />", "9999", 1, -1, vbBinaryCompare)
						sText = Replace(sText, "<JOB_END_MONTH />", "99", 1, -1, vbBinaryCompare)
						sText = Replace(sText, "<JOB_END_DAY />", "99", 1, -1, vbBinaryCompare)
					End If
					sText = Replace(sText, "<AREA_NAME />", CStr(oRecordset.Fields("AreaName").Value), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<AREA_CODE />", CStr(oRecordset.Fields("AreaCode").Value), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<AREA_SHORT_NAME />", CStr(oRecordset.Fields("AreaShortName").Value), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<ECONOMIC_ZONE_ID />", CStr(oRecordset.Fields("EconomicZoneID").Value), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<ECONOMIC_ZONE_CODE />", CStr(oRecordset.Fields("EconomicZoneCode").Value), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<POSITION_NAME />", CStr(oRecordset.Fields("PositionName").Value), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<POSITION_SHORT_NAME />", CStr(oRecordset.Fields("PositionShortName").Value), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<POSITION_TYPE_SHORT_NAME />", CStr(oRecordset.Fields("PositionTypeShortName").Value), 1, -1, vbBinaryCompare)
					Call GetNameFromTable(oADODBConnection, "LastPayrollID", "-1", "", "", lPayrollID, "")
					sText = Replace(sText, "<PAYROLL_YEAR />", Left(CStr(lPayrollID), Len("0000")), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<PAYROLL_NUMBER />", GetPayrollNumber(lPayrollID), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<MARITAL_STATUS_NAME />", CStr(oRecordset.Fields("MaritalStatusName").Value), 1, -1, vbBinaryCompare)
				End If
				oRecordset.Close
			End If
		End If

		If InStr(1, sText, "<EMPLOYEE_JOB_NUMBER />", vbBinaryCompare) > 0 Then
			sErrorDescription = "No se pudo obtener la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobNumber From Employees, Jobs Where (Employees.JobID=Jobs.JobID) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sText = Replace(sText, "<EMPLOYEE_JOB_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("JobNumber").Value)), 1, -1, vbBinaryCompare)
					oRecordset.Close
				End If
			End If
		End If

		If InStr(1, sText, "<EMPLOYEE_PAYMENT_CENTER_NUMBER />", vbBinaryCompare) > 0 Then
			sErrorDescription = "No se pudo obtener la información del empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaShortName From Employees, Areas As PaymentCenters Where (Employees.PaymentCenterID=PaymentCenters.AreaID) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sText = Replace(sText, "<EMPLOYEE_PAYMENT_CENTER_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value)), 1, -1, vbBinaryCompare)
					oRecordset.Close
				End If
			End If
		End If

		sErrorDescription = "No se pudo obtener la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesExtraInfo.*, StateName From EmployeesExtraInfo, States Where (EmployeesExtraInfo.StateID=States.StateID) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sText = Replace(sText, "<EMPLOYEE_ADDRESS />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeAddress").Value)), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_CITY />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeCity").Value)), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_ZIP_CODE />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeZipCode").Value)), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_STATE_ID />", CStr(oRecordset.Fields("StateID").Value), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<EMPLOYEE_STATE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("StateName").Value)), 1, -1, vbBinaryCompare)
				oRecordset.Close
			End If
		End If

		sErrorDescription = "No se pudo obtener la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesSchoolLevelsLKP.*, SchoolarshipName, StatusName From EmployeesSchoolLevelsLKP, Schoolarships, StatusSchoolarships Where (EmployeesSchoolLevelsLKP.SchoolarshipID=Schoolarships.SchoolarshipID) And (EmployeesSchoolLevelsLKP.StatusID=StatusSchoolarships.StatusID) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (StatusSchoolarships.StatusID=1) Order By EmployeesSchoolLevelsLKP.SchoolarshipID Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sText = Replace(sText, "<EMPLOYEE_SCHOOLARSHIP_ID />", CStr(oRecordset.Fields("SchoolarshipID").Value), 1, -1, vbBinaryCompare)
				sText = Replace(sText, "<SCHOOLARSHIP_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("SchoolarshipName").Value)), 1, -1, vbBinaryCompare)
				oRecordset.Close
			End If
		End If

		sErrorDescription = "No se pudo obtener la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(*) As ChildrenCount From EmployeesChildrenLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ChildEndDate In (0,30000000))", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				If Not IsNull(oRecordset.Fields("ChildrenCount").Value) Then
					sText = Replace(sText, "<EMPLOYEE_CHILDREN />", CStr(oRecordset.Fields("ChildrenCount").Value), 1, -1, vbBinaryCompare)
				Else
					sText = Replace(sText, "<EMPLOYEE_CHILDREN />", "0", 1, -1, vbBinaryCompare)
				End If
				oRecordset.Close
			Else
				sText = Replace(sText, "<EMPLOYEE_CHILDREN />", "0", 1, -1, vbBinaryCompare)
			End If
		End If

		If InStr(1, sText, "<HAS_CONCEPT_", vbBinaryCompare) > 0 Then
			iStartPos = InStr(1, sText, "<HAS_CONCEPT_", vbBinaryCompare)
			Do While (iStartPos > 0)
				sConceptShortName = ""
				iStartPos = iStartPos + Len("<HAS_CONCEPT_")
				iEndPos = InStr(iStartPos, sText, " ", vbBinaryCompare)
				If (iEndPos > 0) Then
					sConceptShortName = Mid(sText, iStartPos, (iEndPos - iStartPos))
					If Err.Number = 0 Then
						sErrorDescription = "No se pudo obtener la información de la nómina."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID From Payroll_" & lPayrollID & ", Concepts Where (Payroll_" & lPayrollID & ".ConceptID=Concepts.ConceptID) And (Payroll_" & lPayrollID & ".EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptShortName='" & sConceptShortName & "')", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								sText = Replace(sText, "<HAS_CONCEPT_" & sConceptShortName & " />", "*")
							End If
							oRecordset.Close
						End If
					End If
				End If
				sText = Replace(sText, "<HAS_CONCEPT_" & sConceptShortName & " />", "&nbsp;")
				iStartPos = InStr(1, sText, "<HAS_CONCEPT_", vbBinaryCompare)
				If (Err.Number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
		End If

		If InStr(1, sText, "<PREVIOUS_EMPLOYEE_", vbBinaryCompare) > 0 Then
			sErrorDescription = "No se pudo obtener la información del empleado anterior."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, EmployeeDate From EmployeesHistoryList Where (JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ") And (EmployeeID<>" & aEmployeeComponent(N_ID_EMPLOYEE) & ") Order By EmployeeDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lPreviousEmployeeID = CStr(oRecordset.Fields("EmployeeID").Value)
					sText = Replace(sText, "<PREVIOUS_JOB_START_YEAR />", Left(CStr(oRecordset.Fields("EmployeeDate").Value), Len("0000")), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<PREVIOUS_JOB_START_MONTH />", Mid(CStr(oRecordset.Fields("EmployeeDate").Value), Len("00000"), Len("00")), 1, -1, vbBinaryCompare)
					sText = Replace(sText, "<PREVIOUS_JOB_START_DAY />", Right(CStr(oRecordset.Fields("EmployeeDate").Value), Len("00")), 1, -1, vbBinaryCompare)
					oRecordset.Close
					sErrorDescription = "No se pudo obtener la información del empleado anterior."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Employees Where (EmployeeID=" & lPreviousEmployeeID & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sText = Replace(sText, "<PREVIOUS_EMPLOYEE_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)), 1, -1, vbBinaryCompare)
							sText = Replace(sText, "<PREVIOUS_EMPLOYEE_LAST_NAME />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)), 1, -1, vbBinaryCompare)
							If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								sText = Replace(sText, "<PREVIOUS_EMPLOYEE_LAST_NAME2 />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)), 1, -1, vbBinaryCompare)
							Else
								sText = Replace(sText, "<PREVIOUS_EMPLOYEE_LAST_NAME2 />", " ", 1, -1, vbBinaryCompare)
							End If
							sText = Replace(sText, "<PREVIOUS_EMPLOYEE_RFC />", CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)), 1, -1, vbBinaryCompare)
							sText = Replace(sText, "<PPREVIOUS_EMPLOYEE_CURP />", CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)), 1, -1, vbBinaryCompare)
							sText = Replace(sText, "<PREVIOUS_EMPLOYEE_NUMBER />", CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)), 1, -1, vbBinaryCompare)
							oRecordset.Close
							If InStr(1, sText, "<PREVIOUS_JOB_END_", vbBinaryCompare) > 0 Then
								sErrorDescription = "No se pudo obtener la información del empleado anterior."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeDate From EmployeesHistoryList Where (JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ") And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") Order By EmployeeDate", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										sText = Replace(sText, "<PREVIOUS_JOB_END_YEAR />", Left(CStr(oRecordset.Fields("EmployeeDate").Value), Len("0000")), 1, -1, vbBinaryCompare)
										sText = Replace(sText, "<PREVIOUS_JOB_END_MONTH />", Mid(CStr(oRecordset.Fields("EmployeeDate").Value), Len("00000"), Len("00")), 1, -1, vbBinaryCompare)
										sText = Replace(sText, "<PREVIOUS_JOB_END_DAY />", Right(CStr(oRecordset.Fields("EmployeeDate").Value), Len("00")), 1, -1, vbBinaryCompare)
									End If
									oRecordset.Close
								End If
							End If
						Else
							oRecordset.Close
						End If
					End If
				Else
					oRecordset.Close
				End If
			End If
		End If

		sText = Replace(sText, "<EMPLOYEE_NUMBER />", "", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_FULL_NAME />", "", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_NAME />", "", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_LAST_NAME />", "", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_LAST_NAME2 />", "", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_COMPANY_ID />", "-1", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_JOB_ID />", "-1", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_SERVICE_ID />", "-1", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_TYPE_ID />", "-1", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_POSITION_TYPE_ID />", "-1", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_GROUP_GRADE_LEVEL_ID />", "-1", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_JOURNEY_ID />", "-1", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_SHIFT_ID />", "-1", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_LEVEL_ID />", "-1", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_STATUS_ID />", "-1", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_PAYMENT_CENTER_ID />", "-1", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_EMAIL />", "", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_COUNTRY_ID />", "0", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_GENDER_ID />", "0", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_MARITAL_STATUS_ID />", "0", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_SCHOOLARSHIP_ID />", "-1", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_JOB_NUMBER />", "", 1, -1, vbBinaryCompare)
		sText = Replace(sText, "<EMPLOYEE_PAYMENT_CENTER_NUMBER />", "", 1, -1, vbBinaryCompare)

		iStartPos = InStr(1, sText, "<EMPLOYEE_FIELD ", vbBinaryCompare)
		Do While (iStartPos > 0)
			sFormFieldName = ""
			iMidPos = InStr(iStartPos, sText, "NAME=""", vbBinaryCompare) + Len("NAME=""")
			iEndPos = InStr(iMidPos, sText, """", vbBinaryCompare)
			If (iMidPos > Len("NAME=""")) And (iEndPos > 0) Then
				sFormFieldName = Mid(sText, iMidPos, (iEndPos - iMidPos))
				iEndPos = InStr(iMidPos, sText, "/>", vbBinaryCompare)
				iEndPos = iEndPos + Len("/>")
				If Err.Number = 0 Then
					sErrorDescription = "No se pudo obtener la información del empleado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Answer, FieldTypeID, QueryForSource From EmployeesInformation, EmployeeFields Where (EmployeesInformation.FormFieldID=EmployeeFields.FormFieldID) And (FormFieldName='" & sFormFieldName & "') And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
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
									sErrorDescription = "No se pudieron obtener las respuestas para el formulario del empleado."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & asFields(3) & " From " & asFields(0) & " Where " & sCondition & "(" & asFields(1) & "=" & sAnswer & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oCatalogRecordset)
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
						oRecordset.Close
						sText = Replace(sText, "<EMPLOYEE_FIELD NAME=""" & sFormFieldName & """ />", CleanStringForHTML(sAnswer))
					End If
				End If
			End If
			iStartPos = InStr(1, sText, "<EMPLOYEE_FIELD ", vbBinaryCompare)
			If (Err.Number <> 0) Or (lErrorNumber <> 0) Then Exit Do
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
			If Err.Number <> 0 Then Exit Do
		Loop
	End If

	TransformXMLTagsForEmployeeForm = lErrorNumber
	Set oRecordset = Nothing
	Err.Clear
End Function

Function ModifyEmployeeSafeSeparationConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing concept for the employee in
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyEmployeeSafeSeparationConcept"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim iEndDate
	
	'lErrorNumber = 0
	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del concepto para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		Select Case aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE)
			Case 24, 45, 46, 94
				iEndDate = aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)
			Case Else
				iEndDate = 30000000
		End Select
		If Not CheckEmployeeConceptInformationConsistency(aEmployeeComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			Select Case aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) ' Validaciones por concepto
				Case 120
					If (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 2) And (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 4) And (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 5) And (aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 10) Then
						sErrorDescription = "El porcentaje por seguro de separación solo puede ser 2, 4, 5, 10."
						lErrorNumber = -1
					End If
				Case 87
					sQuery = "Select ConceptID From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=120) And (StartDate>" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (EndDate=30000000)"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If ((lErrorNumber <> 0) And (Not oRecordset.EOF)) Then
						sErrorDescription = "Para capturar el seguro adicional, debe de estar registrado el concepto SI."
						lErrorNumber = -1						
					ElseIf ((aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) > 100) And (aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2)) Then
						sErrorDescription = "El seguro adicional no debe de ser mayor a 100 %."
						lErrorNumber = -1											
					End If
			End Select
			If lErrorNumber = 0 Then
				sQuery = "Select ConceptID From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudo modificar la información del concepto del empleado."
					If oRecordset.EOF Then
						sQuery = "Update EmployeesConceptsLKP Set EndDate=" & AddDaysToSerialDate(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE),-1) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", RegistrationDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (EndDate=30000000)"
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
						If lErrorNumber = 0 Then
							aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
							sErrorDescription = "No se pudo modificar la información del concepto del empleado."
							sQuery = "Insert Into EmployeesConceptsLKP (EmployeeID, ConceptID, StartDate, EndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, RegistrationDate, ModifyDate, StartUserID, EndUserID, UploadedFileName, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", " & iEndDate & ", " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) & ", '" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "', " & aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aLoginComponent(N_USER_ID_LOGIN) &", '" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "')"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
					Else
						sQuery = "Update EmployeesConceptsLKP Set ConceptAmount=" & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", CurrencyID=" & aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) & ", ConceptQttyID=" & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & ", ConceptTypeID=" & aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) & ", ConceptMin=" & aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) & ", ConceptMinQttyID=" & aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) & ", ConceptMax=" & aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) & ", ConceptMaxQttyID=" & aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) & ", AppliesToID='" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "', AbsenceTypeID=" & aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) & ", ConceptOrder=" & aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & ", StartUserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", UploadedFileName='" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "") & "', Comments='" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "', RegistrationDate='" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")"
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)				
					End If
					oRecordset.Close
				End If
			Else
				sErrorDescription = "Para registrar el seguro adicional de separación debe tener registrado el seguro de separación SI."
				lErrorNumber = -1
			End If
			oRecordset.Close
		End If
	End If
	ModifyEmployeeSafeSeparationConcept = lErrorNumber
	Err.Clear
End Function

Function MoveEmployeeBeneficiaryUp(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the beneficiary information for the employee in
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "MoveEmployeeBeneficiaryUp"
	Dim oRecordset
	Dim lErrorNumber
	Dim sField
	Dim sQuery
	Dim bComponentInitialized
	Dim iTempID

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del beneficiario para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else			
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from EmployeesBeneficiariesLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (BeneficiaryID<" & aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) & ") Order By BeneficiaryID Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			iTempID = CInt(oRecordset.Fields("BeneficiaryID").Value)
			'aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE)
			'CInt(oRecordset.Fields("BeneficiaryID").Value)
			sErrorDescription = "No se pudo actualizar la información del beneficiario."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesBeneficiariesLKP Set BeneficiaryID=-1 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (BeneficiaryID=" & CLng(oRecordset.Fields("BeneficiaryID").Value) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesBeneficiariesLKP Set BeneficiaryID=" & iTempID & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (BeneficiaryID=" & aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesBeneficiariesLKP Set BeneficiaryID=" & aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (BeneficiaryID=-1)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				End If
			End If
		Else
			sErrorDescription = "No se pudo actualizar el nivel de prioridad del beneficiario."
		End If
	End If

	MoveEmployeeBeneficiaryUp = lErrorNumber
	Err.Clear
End Function

Function UpdateEmployeeFromHistoryList(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the beneficiary information for the employee in
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UpdateEmployeeFromHistoryList"
	Dim oRecordset
	Dim lErrorNumber
	Dim sField
	Dim sQuery
	Dim bComponentInitialized
	Dim iTempID
	Dim sUpdateQuery

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") Order By EmployeeDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If (Not oRecordset.EOF) Then
				sUpdateQuery = "Update Employees Set CompanyID=" & CStr(oRecordset.Fields("CompanyID").Value) & ", " & _
							   "JobID=" & CStr(oRecordset.Fields("JobID").Value) & ", " & _
							   "ServiceID=" & CStr(oRecordset.Fields("ServiceID").Value) & ", " & _
							   "EmployeeTypeID=" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & ", " & _
							   "PositionTypeID=" & CStr(oRecordset.Fields("PositionTypeID").Value) & ", " & _
							   "ClassificationID=" & CStr(oRecordset.Fields("ClassificationID").Value) & ", " & _
							   "GroupGradeLevelID=" & CStr(oRecordset.Fields("GroupGradeLevelID").Value) & ", " & _
							   "IntegrationID=" & CStr(oRecordset.Fields("IntegrationID").Value) & ", " & _
							   "JourneyID=" & CStr(oRecordset.Fields("JourneyID").Value) & ", " & _
							   "ShiftID=" & CStr(oRecordset.Fields("ShiftID").Value) & ", " & _
							   "WorkingHours=" & CStr(oRecordset.Fields("WorkingHours").Value) & ", " & _
							   "LevelID=" & CStr(oRecordset.Fields("LevelID").Value) & ", " & _
							   "StatusID=" & CStr(oRecordset.Fields("StatusID").Value) & ", " & _
							   "PaymentCenterID=" & CStr(oRecordset.Fields("PaymentCenterID").Value) & ", " & _
							   "Active=" & CStr(oRecordset.Fields("Active").Value) & _
							   " Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE)
				sErrorDescription = "No se pudo actualizar la información del empleado con los datos del último movimiento registrado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sUpdateQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			End If
		Else
			sErrorDescription = "No se pudo actualizar la información del empleado con los datos del último movimiento registrado."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	UpdateEmployeeFromHistoryList = lErrorNumber
	Err.Clear
End Function

Function VerifyAnualDiferenceOfEmployeesConcept(oADODBConnection, aEmployeeComponent, bAnualCalendar, sErrorDescription)
'************************************************************
'Purpose: To verify if employee concept exist with diference minimum of one year
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyAnualDiferenceOfEmployeesConcept"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim lStartDate
	Dim lEndDate

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (Active = 1) Order By StartDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If bAnualCalendar Then
				lStartDate = CLng(Left(CStr(oRecordset.Fields("StartDate").Value), Len("0000"))& "0101")
				lEndDate = CLng(Left(CStr(oRecordset.Fields("StartDate").Value), Len("0000"))& "1231")
				If (aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) >= lStartDate) And (aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) <= lEndDate) Then
					sErrorDescription = "El concepto ya fué registrado en este año calendario."
					VerifyAnualDiferenceOfEmployeesConcept = False
				Else
					VerifyAnualDiferenceOfEmployeesConcept = True
				End If
			Else
				lEndDate = CLng(AddDaysToSerialDate(CLng(oRecordset.Fields("StartDate").Value), 365))
				If aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) > lEndDate Then
					VerifyAnualDiferenceOfEmployeesConcept = True
				Else
					sErrorDescription = "El concepto no cubre el plazo de un año de haberse otorgado."
					VerifyAnualDiferenceOfEmployeesConcept = False
				End If
			End If
		Else
			VerifyAnualDiferenceOfEmployeesConcept = True
		End If
	Else
		sErrorDescription = "Error al verificar si el concepto cubre el plazo de un año de haberse otorgado."
		VerifyAnualDiferenceOfEmployeesConcept = False
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyEmployeeJourneyForConcepts(oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To verify journey requirements for distinct concepts
'Inputs:  oADODBConnection, lReasonID, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyEmployeeJourneyForConcepts"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim iJourneyID

	sQuery = "Select * from Employees" & _
			 " Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
	
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			iJourneyID = CInt(oRecordset.Fields("JourneyID").Value)
			Select Case lReasonID
				Case EMPLOYEES_NIGHTSHIFTS
					If (iJourneyID <> 21) And (iJourneyID <> 22) And (iJourneyID <> 23) Then
						sErrorDescription = "El concepto sólo es válido para los turnos 21, 22 y 23"
						VerifyEmployeeJourneyForConcepts = False
					Else
						VerifyEmployeeJourneyForConcepts = True
					End If
				Case Else
					VerifyEmployeeJourneyForConcepts = True
			End Select
		Else
			sErrorDescription = "Error al verificar el turno del empleado."
			VerifyEmployeeJourneyForConcepts = False
		End If
	Else
		sErrorDescription = "Error al verificar el turno del empleado."
		VerifyEmployeeJourneyForConcepts = False
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyEmployeePositionTypeForConcepts(oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To verify position types requirements for distinct concepts
'Inputs:  oADODBConnection, lReasonID, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyEmployeePositionTypeForConcepts"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim iPositionTypeID

	sQuery = "Select * From Employees Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			iPositionTypeID = CInt(oRecordset.Fields("PositionTypeID").Value)
			Select Case lReasonID
				Case EMPLOYEES_ANTIQUITIES, EMPLOYEES_CHILDREN_SCHOOLARSHIPS, 21, 26, 29, 30, 33, 43, 44, 50
					If iPositionTypeID <> 1 Then
						sErrorDescription = "El concepto es válido sólo para el personal de base"
						VerifyEmployeePositionTypeForConcepts = False
					Else
						VerifyEmployeePositionTypeForConcepts = True
					End If
				CASE EMPLOYEES_FOR_RISK, EMPLOYEES_ADDITIONALSHIFT
					If (iPositionTypeID <> 1) And (iPositionTypeID <> 4) Then
						sErrorDescription = "El concepto es válido sólo para el personal de base"
						VerifyEmployeePositionTypeForConcepts = False
					Else
						VerifyEmployeePositionTypeForConcepts = True
					End If
				Case EMPLOYEES_CONCEPT_08
					If iPositionTypeID <> 2 Then
						sErrorDescription = "El concepto es válido sólo para el personal de confianza"
						VerifyEmployeePositionTypeForConcepts = False
					Else
						VerifyEmployeePositionTypeForConcepts = True
					End If
				Case Else
					VerifyEmployeePositionTypeForConcepts = True
			End Select
		Else
			sErrorDescription = "Error al verificar el tipo de puesto del empleado para registrar el concepto."
			VerifyEmployeePositionTypeForConcepts = False
		End If
	Else
		sErrorDescription = "Error al verificar el tipo de puesto del empleado para registrar el concepto."
		VerifyEmployeePositionTypeForConcepts = False
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyEmployeeTypeForConcepts(oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To verify employee types requirements for distinct concepts
'Inputs:  oADODBConnection, lReasonID, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyEmployeeTypeForConcepts"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim iEmployeeTypeID
	Dim iStatusEmployeeID

	sQuery = "Select * From Employees Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			iEmployeeTypeID = CInt(oRecordset.Fields("EmployeeTypeID").Value)
			iStatusEmployeeID = CInt(oRecordset.Fields("StatusID").Value)
			Select Case lReasonID
				Case EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE
					If (iEmployeeTypeID = 1) Or (iEmployeeTypeID = 7) Or (iStatusEmployeeID = 1) Then
						sErrorDescription = "El concepto no es válido para Internos, Honorarios ni Funcionarios"
						VerifyEmployeeTypeForConcepts = False
					Else
						VerifyEmployeeTypeForConcepts = True
					End If
				Case EMPLOYEES_GLASSES, EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS, EMPLOYEES_CONCEPT_08
					If (iEmployeeTypeID = 1) Then
						sErrorDescription = "El concepto no es válido para Funcionarios"
						VerifyEmployeeTypeForConcepts = False
					Else
						VerifyEmployeeTypeForConcepts = True
					End If
				Case -61, -62, -95
					If (iEmployeeTypeID <> 1) Then
						sErrorDescription = "El concepto es sólo válido para Funcionarios"
						VerifyEmployeeTypeForConcepts = False
					Else
						VerifyEmployeeTypeForConcepts = True
					End If
				Case EMPLOYEES_HONORARIUM_CONCEPT
					If (iEmployeeTypeID <> 7) Then
						sErrorDescription = "El concepto sólamente es válido para personal por Honorarios"
						VerifyEmployeeTypeForConcepts = False
					Else
						VerifyEmployeeTypeForConcepts = True
					End If
                Case Else
					VerifyEmployeeTypeForConcepts = True
			End Select
		Else
			sErrorDescription = "Error al verificar el tipo de empleado para registrar el concepto."
			VerifyEmployeeTypeForConcepts = False
		End If
	Else
		sErrorDescription = "Error al verificar el tipo de empleado para registrar el concepto."
		VerifyEmployeeTypeForConcepts = False
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyJobOwner(oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: Ty verify if the employee and the owner of the job 
'are the same
'Inputs:  oADODBConnection, lReasonID, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyJobOwner"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim sOwner
	Dim nJobId
	
	lErrorNumber = 0
	VerifyJobOwner = True
	
	sQuery = "Select OwnerID From Jobs, Employees Where (Employees.JobID = Jobs.JobID) And (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If CLng(oRecordset.Fields("OwnerID").Value) <> CLng(aEmployeeComponent(N_ID_EMPLOYEE)) Then
			sErrorDescription = "El empleado actual no es titular de la plaza indicada"
			VerifyJobOwner = False
		End If
	Else
		sErrorDescription = "No se pudo verificar la titularidad del empleado"
		VerifyJobOwner = False
	End If	

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyEmployeeStatus(oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To verify employee status requirements to register absences
'Inputs:  oADODBConnection, lReasonID, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyEmployeeStatus"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim iEmployeeTypeID
	Dim iStatusEmployeeID
	Dim sStatusEmployee

	sQuery = "Select * From Employees Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			iEmployeeTypeID = CInt(oRecordset.Fields("EmployeeTypeID").Value)
			iStatusEmployeeID = CInt(oRecordset.Fields("StatusID").Value)
			If (CInt(oRequest("ReasonID").Item) = EMPLOYEES_MOTHERAWARD) And (iStatusEmployeeID <> 0) Then
				Call GetNameFromTable(oADODBConnection, "StatusEmployees", iStatusEmployeeID, "", "", sStatusEmployee, sErrorDescription)
				sErrorDescription = "Solamente se puede registrar este tipo de incidencias al personal con estatus activo. El status actual del empleado es " & sStatusEmployee
				VerifyEmployeeStatus = False
			ElseIf (iStatusEmployeeID <> 0) And (iStatusEmployeeID <> 1) Then
				sErrorDescription = "El empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " no está activo."
				VerifyEmployeeStatus = False
			Else
				VerifyEmployeeStatus = True
			End If
		Else
			sErrorDescription = "Error al verificar si el empleado está activo."
			VerifyEmployeeStatus = False
		End If
	Else
		sErrorDescription = "No existe un registro del empleado indicado."
		VerifyEmployeeStatus = False
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyEmployeeStatusInHistoryList(oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To verify employee status requirements to register absences
'Inputs:  oADODBConnection, lReasonID, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyEmployeeStatusInHistoryList"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim iStatusEmployeeID
	Dim sStatusEmployee

	If CLng(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)) = CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) Then
		VerifyEmployeeStatusInHistoryList = True
	Else
		sQuery = "Select * From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
				 " And (EmployeeDate<=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And (EndDate>=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ")"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				iStatusEmployeeID = CInt(oRecordset.Fields("StatusID").Value)
				If (iStatusEmployeeID <> 0) And (iStatusEmployeeID <> 1) Then
					Select Case iStatusEmployeeID
						Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 27, 28, 29, 31, 32, 33, 35, 36, 37, 39, 40, 41, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 119, 120, 121, 123, 124, 125, 126, 127, 128, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 145, 146, 147, 149, 150, 151, 152, 153, 154, 155, 156, 157, 158
						Case Else
							lErrorNumber = -1
					End Select
					If lErrorNumber = 0 Then
						VerifyEmployeeStatusInHistoryList = True
					Else
						Call GetNameFromTable(oADODBConnection, "StatusEmployees", iStatusEmployeeID, "", "", sStatusEmployee, sErrorDescription)
						sErrorDescription = "El empleado no estará activo en el periodo indicado. El estatus del empleado en la fecha indicada es " & sStatusEmployee
						VerifyEmployeeStatusInHistoryList = False
					End If
				Else
					VerifyEmployeeStatusInHistoryList = True
				End If
			Else
				sErrorDescription = "No existen registros en el historial de movimientos para obtener el estatus del empleado."
				VerifyEmployeeStatusInHistoryList = False
			End If
		Else
			sErrorDescription = "Error al verificar el estatus del empleado."
			VerifyEmployeeStatusInHistoryList = False
		End If
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyExistenceOfConceptAudit(oADODBConnection, aAuditComponent, sAuditType, sAuditOperation, sErrorDescription)
'************************************************************
'Purpose: To verify if employee exist in EmployeesChildrenLKP table
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistenceOfConceptAudit"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery

	sQuery = "Select * From AuditOperationsLKP, AuditTypes, AuditOperationTypes" & _
			 " Where AuditOperationsLKP.AuditTypeID = AuditTypes.AuditTypeID" & _
			 " And AuditOperationsLKP.AuditOperationTypeID = AuditOperationTypes.AuditOperationTypeID" & _
			 " And AuditTypeShortName='" & sAuditType & "'" & _
			 " And AuditOperationShortName='" & sAuditOperation & "'"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			aAuditComponent(N_AUDIT_CONCEPT_TYPE_ID) = CLng(oRecordset("AuditTypeID"))
			aAuditComponent(N_AUDIT_OPERATION_TYPE) = CLng(oRecordset("AuditOperationID"))
			VerifyExistenceOfConceptAudit = True
		Else
			sErrorDescription = "Error al verificar si el concepto tiene registro de auditoria."
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			VerifyExistenceOfConceptAudit = False
		End If
	Else
		sErrorDescription = "Error al verificar si el concepto tiene registro de auditoria."
		VerifyExistenceOfConceptAudit = False
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyExistenceOfEmployeeAbsences(oADODBConnection, aAbsenceComponent, sAbsenceIDs, sErrorDescription)
'************************************************************
'Purpose: To verify if an absence already exist in database
'Inputs:  oADODBConnection, aAbsenceComponent, sAbsenceIDs, bIsForPeriod
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistenceOfEmployeeAbsences"
	Dim lErrorNumber
	Dim oRecordset
	Dim oRecordset1
	Dim sErrorDescription1
	Dim sQuery
	Dim sCondition
	Dim lOcurredDate
	Dim bComponentInitialized

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (InStr(1, sAbsenceIDs, "9999", vbBinaryCompare) > 0) Then
		VerifyExistenceOfEmployeeAbsences = True
	Else
		sQuery = "Select * From EmployeesAbsencesLKP Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
		'Case 10,11,12,13,14,16,17,82,83,84,85,86,87,29,30,31,32,33,34,35,37,38
		sCondition = " And (AbsenceID In (" & sAbsenceIDs & ",201,202))"
		'sCondition = " And (AbsenceID In (" & sAbsenceIDs & "))"
		sQuery = sQuery & sCondition & _
				 " And (((OcurredDate >= " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (OcurredDate <= " &  aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & "))" & _
				 " Or ((EndDate >= " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (EndDate <= " &  aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & "))" & _
				 " Or ((EndDate >= " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (OcurredDate <= " &  aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ")))"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
					sErrorDescription1 = ""
					lOcurredDate = CLng(oRecordset.Fields("OcurredDate").Value)
					Do While Not oRecordset.EOF
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from Absences Where (AbsenceID = " & oRecordset.Fields("AbsenceID").Value & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset1)
						sErrorDescription1 = sErrorDescription1 & " " & CStr(oRecordset1.Fields("AbsenceName").Value) & ", con fecha de inicio del " &  DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value), -1, -1, -1) & " y fecha de término del " &  DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1) & ";"
						oRecordset1.Close
						oRecordset.MoveNext
					Loop
					sErrorDescription = "Para poder registrar el concepto no debe de estar registrado alguno de los siguientes: " & sErrorDescription1 & " verifique que así sea"
					VerifyExistenceOfEmployeeAbsences = False
			Else
				VerifyExistenceOfEmployeeAbsences = True
			End If
			oRecordset.Close
		Else
			sErrorDescription = "Error al verificar si esta registrado otro concepto o incidencia."
			VerifyExistenceOfEmployeeAbsences = False
		End If
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyExistenceOfEmployeesBeneficiary(oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if employee exist in EmployeesChildrenLKP table
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistenceOfEmployeesBeneficiary"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery

	sQuery = "Select * From EmployeesBeneficiariesLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EndDate>" & Left(GetSerialNumberForDate(""), Len("00000000")) & ")"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	
	If lErrorNumber = 0 Then
		VerifyExistenceOfEmployeesBeneficiary = (Not oRecordset.EOF)
	Else
		sErrorDescription = "Error al verificar si el empleado tiene registrado beneficiarios."
		VerifyExistenceOfEmployeesBeneficiary = False
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyExistenceOfEmployeesConcept(oADODBConnection, aEmployeeComponent, iExistenceType, sErrorDescription)
'************************************************************
'Purpose: To verify if a concept already exist in database
'Inputs:  oADODBConnection, aAbsenceComponent, sAbsenceIDs, bIsForPeriod
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistenceOfEmployeesConcept"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim sEmployeeConceptType
	Dim lStartDate
	Dim lEndDate
	Dim sConceptShortName

	Select Case aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE)
		Case 22, 24, 45, 46, 50, 72, 73, 93, 94, 26, 100, 32, 93, 13
			sQuery = "Select * From EmployeesConceptsLKP Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID =" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ")" & _
					 " And (StartDate >= " &  aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (EndDate <= " &  aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ")" & _
					 " And (EmployeesConceptsLKP.Active<>2)"
		Case Else
			sQuery = "Select * From EmployeesConceptsLKP Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID =" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ")" & _
					 " And (((StartDate >= " &  aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (EndDate <= " &  aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & "))" & _
					 " Or ((EndDate >= " &  aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (EndDate <= " &  aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & "))" & _
					 " Or ((EndDate >= " &  aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (StartDate <= " &  aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ")))" & _
					 " And (EmployeesConceptsLKP.Active<>2)"
	End Select

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Call GetNameFromTable(oADODBConnection, "ShortConcepts", aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE), "", "", sConceptShortName, "")
			Select Case aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE)
				Case 22, 24, 45, 46, 94, 26, 100, 32, 93
					Select Case CInt(oRecordset.Fields("Active").Value)
						Case 0
							iExistenceType = 0
							sErrorDescription = "Existe registrada una prestación " & sConceptShortName & " en proceso para el empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " en la fecha " & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1, -1, -1) & "."
						Case 1
							iExistenceType = 1
							sErrorDescription = "Existe registrada una prestación " & sConceptShortName & "  activa para el empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " en la fecha " & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1, -1, -1) & "."
						Case 2
							iExistenceType = 2
							sErrorDescription = "Existe registrada una prestación " & sConceptShortName & " cancelada para el empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " en la fecha " & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1, -1, -1) & "."
					End Select
					VerifyExistenceOfEmployeesConcept = True
				Case Else
					aEmployeeComponent(N_CONCEPT_CREDIT_TYPE) = 1
					Call GetCrossingEmployeeConceptType(oADODBConnection, aEmployeeComponent, sEmployeeConceptType, lStartDate, lEndDate, sErrorDescription)
					Select Case sEmployeeConceptType
						Case "Left", "Right"
							VerifyExistenceOfEmployeesConcept = False
						Case "Inner"
							VerifyExistenceOfEmployeesConcept = False
						Case Else
							sErrorDescription = "No se puede agregar el registro de " & sConceptShortName & " para el empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " con fecha de inicio " & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1, -1, -1) & " debido a que existe uno en el periodo indicado. Puede agregar registros con fecha de inicio mayor a la del registro existente; si se aplica, se cierra el efecto a la fecha de inicio del nuevo."
							VerifyExistenceOfEmployeesConcept = True
					End Select
			End Select
		Else
			VerifyExistenceOfEmployeesConcept = False
		End If
	Else
		sErrorDescription = "Error al verificar si ya esta registrada la prestación."
		VerifyExistenceOfEmployeesConcept = True
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyExistenceOfEmployeesCreditSp(oADODBConnection, aEmployeeComponent, lActive, sErrorDescription)
'************************************************************
'Purpose: To verify if an employee has a credit
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistenceOfEmployeesCreditSp"
	Dim lErrorNumber
	Dim oRecordset

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Credits Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (CreditTypeID = " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (EndDate >= " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (Active = " & lActive & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		VerifyExistenceOfEmployeesCreditSp = (Not oRecordset.EOF)
	Else
		sErrorDescription = "Error al verificar si el empleado tiene registrado el crédito."
		aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) = -1
		VerifyExistenceOfEmployeesCreditSp = False
	End If
	Err.Clear
End Function

Function VerifyExistenceOfEmployeesCredit(oADODBConnection, aEmployeeComponent, iExistenceType, sErrorDescription)
'************************************************************
'Purpose: To verify if an employee has a credit already exist in database
'Inputs:  oADODBConnection, aAbsenceComponent, sAbsenceIDs, bIsForPeriod
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistenceOfEmployeesCredit"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim sEmployeeConceptType
	Dim lStartDate
	Dim lEndDate
	Dim sCreditType

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		VerifyExistenceOfEmployeesCredit = False
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador del tipo de crédito para validar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sQuery = "Select * from Credits Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (CreditTypeID =" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ")" & _
				 " And (((StartDate >= " &  aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (EndDate <= " &  aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & "))" & _
				 " Or ((EndDate >= " &  aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (EndDate <= " &  aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & "))" & _
				 " Or ((EndDate >= " &  aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") And (StartDate <= " &  aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ")))" & _
				 " And (Credits.Active<>2)"

		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				Call GetNameFromTable(oADODBConnection, "CreditTypes", aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE), "", "", sCreditType, sErrorDescription)
				aEmployeeComponent(N_CONCEPT_CREDIT_TYPE) = 0
				Call GetCrossingEmployeeConceptType(oADODBConnection, aEmployeeComponent, sEmployeeConceptType, lStartDate, lEndDate, sErrorDescription)
				Select Case sEmployeeConceptType
					Case "Left", "Right", "Cross"
						sErrorDescription = "No se puede agregar el crédito " & sCreditType & " del empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " porque la fecha de inicio " & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1, -1, -1) & " es anterior a la de uno ya registrado."
						VerifyExistenceOfEmployeesCredit = False
						iExistenceType = 0
					Case "Inner"
						'sErrorDescription = "No se puede insertar el nuevo crédito del empleado puesto que se traslapa con otro ya registrado."
						'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Credits Set EndDate=" & AddDaysToSerialDate(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (CreditTypeID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & lStartDate & ") And (EndDate=" & lEndDate & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
						Select Case aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE)
							Case 61,82, 58, 64, 83
								VerifyExistenceOfEmployeesCredit = False
							Case Else
								sErrorDescription = "No se puede agregar el crédito " & sCreditType & " del empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " con fecha de inicio " & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), -1, -1, -1) & " debido a que existe un registro en el periodo indicado"
								VerifyExistenceOfEmployeesCredit = True
						End Select
						iExistenceType = 1
					Case Else
						sErrorDescription = "No se puede agregar el crédito debido a que se traslapa con un registro en el periodo indicado"
						VerifyExistenceOfEmployeesCredit = True
						iExistenceType = 3
				End Select
			Else
				VerifyExistenceOfEmployeesCredit = False
			End If
		Else
			sErrorDescription = "Error al verificar si ya esta registrado el tipo de crédito."
			VerifyExistenceOfEmployeesCredit = True
		End If
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyExistenceOnEmployeesChildren(oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if employee exist in EmployeesChildrenLKP table
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistenceOnEmployeesChildren"
	Dim lErrorNumber
	Dim oRecordset

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Employees, EmployeesChildrenLKP Where (Employees.EmployeeID=EmployeesChildrenLKP.EmployeeID) And (GenderID=0) And (Employees.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			VerifyExistenceOnEmployeesChildren = True
		Else
			sErrorDescription = "La empleada no esta registrada en el padrón de madres."
			VerifyExistenceOnEmployeesChildren = False
		End If
	Else
		sErrorDescription = "Error al verificar si la empleada esta registrada en el padrón de madres."
		VerifyExistenceOnEmployeesChildren = False
	End If
	Err.Clear
End Function

Function VerifyRequerimentsToEmployeesAbsences(oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if employee exist in EmployeesChildrenLKP table
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyRequerimentsToEmployeesAbsences"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Then
		sErrorDescription = "No se especificó el identificador del empleado para agregar la información del beneficiario."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
		VerifyRequerimentsToEmployeesAbsences = False
	Else
		lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			If VerifyUserPermissionOnEmployee(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
				Select Case lReasonID
					Case EMPLOYEES_BANK_ACCOUNTS
						If Not VerifyEmployeeStatus(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
							VerifyRequerimentsToEmployeesAbsences = False
						Else
							VerifyRequerimentsToEmployeesAbsences = True
						End If
					Case 43,44,45,46,47,48
						lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						If lErrorNumber = 0 Then
							If VerifyJobOwner(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
								VerifyRequerimentsToEmployeesAbsences = True
							Else
								VerifyRequerimentsToEmployeesAbsences = False
							End If
						Else
							If Len(sErrorDescription) = 0 Then ErrorDescription = "No tiene permisos para realizar movimientos a empleados que pertenecen a otro centro de trabajo."
							VerifyRequerimentsToEmployeesAbsences = False
						End If
					Case EMPLOYEES_BENEFICIARIES_DEBIT
						If VerifyExistenceOfEmployeesBeneficiary(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
							VerifyRequerimentsToEmployeesAbsences = True
						Else
							sErrorDescription = "El empleado indicado no tiene registrados beneficiarios para calcular el concepto"
							VerifyRequerimentsToEmployeesAbsences = False
						End If
					Case EMPLOYEES_MOTHERAWARD
						If Not VerifyEmployeeStatus(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
							VerifyRequerimentsToEmployeesAbsences = False
						Else
							VerifyRequerimentsToEmployeesAbsences = True
						End If
					Case Else
						If VerifyEmployeeJourneyForConcepts(oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription) Then
							If VerifyEmployeePositionTypeForConcepts(oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription) Then
								If VerifyEmployeeTypeForConcepts(oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription) Then
									VerifyRequerimentsToEmployeesAbsences = True
								Else
									VerifyRequerimentsToEmployeesAbsences = False
								End If
							Else
								VerifyRequerimentsToEmployeesAbsences = False
							End If
						Else
							VerifyRequerimentsToEmployeesAbsences = False
						End If
				End Select
			Else
				VerifyRequerimentsToEmployeesAbsences = False
			End If
		Else
			VerifyRequerimentsToEmployeesAbsences = False
		End If
	End If
	Err.Clear
End Function

Function VerifyRequerimentsForEmployeesConcepts(oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if employee exist in EmployeesChildrenLKP table
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyRequerimentsForEmployeesConcepts"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Then
		sErrorDescription = "No se especificó el identificador del empleado para agregar la información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
		VerifyRequerimentsForEmployeesConcepts = False
	Else
		lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			If VerifyUserPermissionOnEmployee(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
				If VerifyRecordIntegrity(oADODBConnection,aEmployeeComponent,lReasonID,sErrorDescription) Then
					Select Case lReasonID
						Case EMPLOYEES_HONORARIUM_CONCEPT
							If Not VerifyEmployeeStatus(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
								VerifyRequerimentsForEmployeesConcepts = False
							Else
								VerifyRequerimentsForEmployeesConcepts = True
							End If
						Case EMPLOYEES_BANK_ACCOUNTS
							If Not VerifyEmployeeStatus(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
								VerifyRequerimentsForEmployeesConcepts = False
							Else
								VerifyRequerimentsForEmployeesConcepts = True
							End If
						Case 43,44,45,46,47,48
							lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							If lErrorNumber = 0 Then
								If VerifyJobOwner(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
									VerifyRequerimentsForEmployeesConcepts = True
								Else
									VerifyRequerimentsForEmployeesConcepts = False
								End If
							Else
								If Len(sErrorDescription) = 0 Then ErrorDescription = "No tiene permisos para realizar movimientos a empleados que pertenecen a otro centro de trabajo."
								VerifyRequerimentsForEmployeesConcepts = False
							End If
						Case EMPLOYEES_BENEFICIARIES_DEBIT
							If VerifyExistenceOfEmployeesBeneficiary(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
								VerifyRequerimentsForEmployeesConcepts = True
							Else
								sErrorDescription = "El empleado indicado no tiene registrados beneficiarios para calcular el concepto"
								VerifyRequerimentsForEmployeesConcepts = False
							End If
						Case EMPLOYEES_MOTHERAWARD
							If Not VerifyEmployeeStatus(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
								VerifyRequerimentsForEmployeesConcepts = False
							Else
								VerifyRequerimentsForEmployeesConcepts = True
							End If
						Case Else
							If VerifyEmployeeJourneyForConcepts(oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription) Then
								If VerifyEmployeePositionTypeForConcepts(oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription) Then
									If VerifyEmployeeTypeForConcepts(oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription) Then
										VerifyRequerimentsForEmployeesConcepts = True
									Else
										VerifyRequerimentsForEmployeesConcepts = False
									End If
								Else
									VerifyRequerimentsForEmployeesConcepts = False
								End If
							Else
								VerifyRequerimentsForEmployeesConcepts = False
							End If
					End Select
				Else
					VerifyRequerimentsForEmployeesConcepts = False
				End IF
			Else
				VerifyRequerimentsForEmployeesConcepts = False
			End If
		Else
			VerifyRequerimentsForEmployeesConcepts = False
		End If
	End If
	Err.Clear
End Function

Function VerifyRequerimentsForEmployeesBankAccounts(oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if employee exist in EmployeesChildrenLKP table
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyRequerimentsForEmployeesBankAccounts"
	Dim lErrorNumber
	Dim lBankLength
	Dim oRecordset
	Dim sQuery
	Dim bComponentInitialized
	Dim iStatusEmployeeID

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If
	VerifyRequerimentsForEmployeesBankAccounts = True

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para obtener sus cuentas bancarias."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		Select Case lReasonID
			Case EMPLOYEES_ADD_BENEFICIARIES
				lErrorNumber = CheckExistencyOfEmployeeBeneficiary(aEmployeeComponent, sErrorDescription)
			Case EMPLOYEES_CREDITORS
				lErrorNumber = CheckExistencyOfEmployeeCreditors(aEmployeeComponent, sErrorDescription)
			Case Else
				lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
		End Select
		If lErrorNumber = 0 Then
			If (lReasonID = EMPLOYEES_ADD_BENEFICIARIES) Or (lReasonID = EMPLOYEES_CREDITORS) Then
				VerifyRequerimentsForEmployeesBankAccounts = True
			Else
				If VerifyUserPermissionOnEmployee(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
					If Not VerifyEmployeeStatus(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
						sErrorDescription = "Solamente se pueden registrar cuentas bancarias al personal activo."
						VerifyRequerimentsForEmployeesBankAccounts = False
					Else
						If (InStr(1, aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE), ".") > 0) Then
							VerifyRequerimentsForEmployeesBankAccounts = True
						ElseIf False Then
							lBankLength = Len(CStr(aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE)))
							sErrorDescription = "Verifique la longitud de la cuenta indicada."
							Select Case aEmployeeComponent(N_BANK_ID_EMPLOYEE)
								Case 1 'BBVA Bancomer
									VerifyRequerimentsForEmployeesBankAccounts = (lBankLength = 16)
								Case 3 'Banamex
									VerifyRequerimentsForEmployeesBankAccounts = (lBankLength = 16)
								Case 14 'Banorte
									VerifyRequerimentsForEmployeesBankAccounts = (lBankLength = 18)
								Case 17 'Santander Serfín
									VerifyRequerimentsForEmployeesBankAccounts = (lBankLength = 11)
								Case 18 'Scotiabank
									VerifyRequerimentsForEmployeesBankAccounts = (lBankLength = 10)
								Case 24 'HSBC
									VerifyRequerimentsForEmployeesBankAccounts = (lBankLength = 10)
								Case Else
									VerifyRequerimentsForEmployeesBankAccounts = False
							End Select
						End If
					End If
				Else
					VerifyRequerimentsForEmployeesBankAccounts = False
				End If
			End If
		Else
			VerifyRequerimentsForEmployeesBankAccounts = False
		End If
	End If
	Err.Clear
End Function

Function VerifyRequerimentsForEmployeesGrades(oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if employee exist in EmployeesChildrenLKP table
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyRequerimentsForEmployeesGrades"
	Dim lErrorNumber
	Dim lBankLength
	Dim oRecordset
	Dim sQuery
	Dim bComponentInitialized
	Dim iStatusEmployeeID

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para obtener sus cuentas bancarias."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			If VerifyUserPermissionOnEmployee(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
				If Not VerifyEmployeeStatus(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
					sErrorDescription = "Solamente se pueden registrar calificación al personal activo."
					VerifyRequerimentsForEmployeesGrades = False
				Else
					VerifyRequerimentsForEmployeesGrades = True
				End If
			Else
				VerifyRequerimentsForEmployeesGrades = False
			End If
		Else
			VerifyRequerimentsForEmployeesGrades = False
		End If
	End If
	Err.Clear
End Function

Function VerifyIfArePayrollsOpen(sErrorDescription)
'************************************************************
'Purpose: Check for open payrolls
'Inputs:
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyIfArePayrollsOpen"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la nóminas abiertas."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PayrollID From Payrolls Where (IsClosed<>1)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			lErrorNumber = L_ERR_NO_RECORDS
		Else
			Do While Not oRecordset.EOF
				sPayrollsIDs = sPayrollsIDs & CStr(oRecordset.Fields("PayrollID").Value) & ","
				oRecordset.MoveNext
			Loop
			If (InStr(Right(sPayrollsIDs,1),",") > 0) Then
				sPayrollsIDs = Left(sPayrollsIDs, Len(sPayrollsIDs) -1)
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	VerifyIfArePayrollsOpen = lErrorNumber
	Err.Clear
End Function

Function VerifyIfEmployeeIsInLocalZone(oADODBConnection, aEmployeeComponent, lOccurredDate, bZoneFlag, sErrorDescription)
'************************************************************
'Purpose: Verify if employee is in local or foreing zone in a specific date
'Inputs:  oADODBConnection, aEmployeeComponent, lOccurredDate, bZoneFlag
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyIfEmployeeIsInLocalZone"
	Dim oRecordset
	Dim lErrorNumber
	Dim sQuery

	sQuery = "Select ParentZones.ZoneName, AreasZones.ZoneName, Areas.AreaName, ParentAreas.AreaName, EmployeeID, Jobs.JobID" & _
			 " From Zones As ParentZones, Zones As AreasZones, Areas, Areas As ParentAreas, Employees, Jobs" & _
			 " Where (Employees.JobID=Jobs.JobID)" & _
			 " And (Jobs.AreaID=Areas.AreaID) And (Areas.ParentID=ParentAreas.AreaID)" & _
			 " And (AreasZones.ParentID=ParentZones.ZoneID) And (Areas.ZoneID=AreasZones.ZoneID)"
			If bZoneFlag Then
				sQuery = sQuery & " And (AreasZones.ZonePath Like '%,9,%')"
			Else
				sQuery = sQuery & _
						 " And ((ParentZones.ParentID In (1,2,3,4,7,8,5,6,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,38))" & _
							" Or (Areas.AreaPath Like '%,38,%')" & _
						 ")"
			End If
	sQuery = sQuery & " And (Areas.StartDate<=" & lOccurredDate & ") And (Areas.EndDate>=" & lOccurredDate & ")" & _
			 " And (ParentZones.StartDate<=" & lOccurredDate & ") And (ParentZones.EndDate>=" & lOccurredDate & ")" & _
			 " And (Employees.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			If bZoneFlag Then
				sErrorDescription = "El empleado no pertenece a alguna zona local."
			Else
				sErrorDescription = "El empleado no pertenece a alguna zona foranea."
			End If
			VerifyIfEmployeeIsInLocalZone = False
		Else
			VerifyIfEmployeeIsInLocalZone = True
		End If
		oRecordset.Close
	Else
		sErrorDescription = "Error al validar la zona del empleado."
		VerifyIfEmployeeIsInLocalZone = False
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyPayrollIsActive(oADODBConnection, lPayrollID, lPayrollType, sErrorDescription)
'************************************************************
'Purpose: To verify if employee exist in EmployeesChildrenLKP table
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyPayrollIsActive"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim sPayrollTypeCondition
	Dim bComponentInitialized

	Select Case lPayrollType
		Case N_PAYROLL_FOR_MOVEMENTS
			sPayrollTypeCondition = "IsActive_1"
		Case N_PAYROLL_FOR_ABSENCES
			sPayrollTypeCondition = "IsActive_2"
		Case N_PAYROLL_FOR_MOTHER
			sPayrollTypeCondition = "IsActive_3"
		Case N_PAYROLL_FOR_BANK
			sPayrollTypeCondition = "IsActive_4"
		Case N_PAYROLL_FOR_PROVAC
			sPayrollTypeCondition = "IsActive_5"
		Case N_PAYROLL_FOR_FONAC
			sPayrollTypeCondition = "IsActive_6"
		Case N_PAYROLL_FOR_FEATURES
			sPayrollTypeCondition = "IsActive_7"
		Case N_PAYROLL_FOR_8
			sPayrollTypeCondition = "IsActive_8"
		Case N_PAYROLL_FOR_9
			sPayrollTypeCondition = "IsActive_9"
		Case N_PAYROLL_FOR_10
			sPayrollTypeCondition = "IsActive_10"
	End Select

	If lPayrollType = 0 Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PayrollID From Payrolls Where (PayrollID=" & lPayrollID & ") And (IsClosed<>1)", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PayrollID From Payrolls Where (PayrollID=" & lPayrollID & ") And (" & sPayrollTypeCondition & "=1) And (IsClosed<>1)", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			VerifyPayrollIsActive = True
		Else
			sErrorDescription = "La nómina " & CStr(GetDateFromSerialNumber(lPayrollID)) & " indicada no está activa."
			VerifyPayrollIsActive = False
		End If
	Else
		sErrorDescription = "Error al verificar si la nómina " & CStr(GetDateFromSerialNumber(lPayrollID)) & " indicada está activa."
		VerifyPayrollIsActive = False
	End If
	Err.Clear
End Function

Function VerifyDatesForPayroll(aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if employee exist in EmployeesChildrenLKP table
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyDatesForPayroll"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim sPayrollTypeCondition
	Dim bComponentInitialized

	If (aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) < GetPayrollStartDate(aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE))) Or (aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) > CLng(aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE))) Then
		sErrorDescription = "La fecha de inicio no esta dentro del rango de la quincena de aplicación."
		VerifyDatesForPayroll = False
	Else
		If (aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) < GetPayrollStartDate(aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE))) Or (aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) > CLng(aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE))) Then
			sErrorDescription = "La fecha de fin no esta dentro del rango de la quincena de aplicación."
			VerifyDatesForPayroll = False
		Else
			VerifyDatesForPayroll = True
		End If
	End If
	Err.Clear
End Function

Function VerifyUserPermissionOnEmployee(oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if employee exist in EmployeesChildrenLKP table
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyUserPermissionOnEmployee"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim bComponentInitialized

	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) = 0 Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Employees Where EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE), "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.* From Employees, Jobs Where (Employees.JobID=Jobs.JobID) And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Jobs.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))) And (Employees.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	End If

	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			sErrorDescription = "No tiene permisos para realizar movimientos a empleados que pertenecen a otro centro de trabajo."
			VerifyUserPermissionOnEmployee = False
		Else
			VerifyUserPermissionOnEmployee = True
		End If
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyJobJourneyForEmployeeShift(oADODBConnection, oRequest, sErrorDescription)
'************************************************************
'Purpose: Verify the shift of the employee with the jourey of the job
'Inputs:  oADODBConnection, aEmployeeComponent, oRequest
'Outputs: sErrorDescription
'************************************************************
	Const S_FUNCTION_NAME = "VerifyJobJourneyForEmployeeShift"
	Dim lErrorNumber
	Dim sQuery
	Dim oRecordset
	Dim lJourneyID
	Dim lJourneyShift
	Dim lJobID

	If InStr(oRequest.Item,"JobID") > 0 Then
		lJobID = oRequest("JobID").Item
	Else
		lJobID = oEmployeeComponent(N_JOB_ID_EMPLOYEE)
	End If
	sQuery = "Select JourneyID From Jobs Where JobID = " & lJobID
	sErrorDescription = "No se encontró la información de la plaza"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		lJourneyID = oRecordset.Fields("JourneyID").Value
		sQuery = "Select JourneyID From Shifts Where ShiftID = " & oRequest("ShiftID").Item
		sErrorDescription = "No se pudo obtener la información del horario del empleado"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			lJourneyShift = oRecordset.Fields("JourneyID").Value
			If CLng(lJourneyID) <> CLng(lJourneyShift) Then
				lErrorNumber = -1
				sErrorDescription = "El horario del empleado no corresponde al turno de la plaza"
			End If
		End If
	End If

	VerifyJobJourneyForEmployeeShift = lErrorNumber

	Set oRecordset = Nothing
	Err.Clear
End Function

Function CalculateEmployeeAntiquity(oADODBConnection, aEmployeeComponent, lDate, sEmployeeAntiquity, lAntiquityYears, lAntiquityMonths, lAntiquityDays, sErrorDescription)
'************************************************************
'Purpose: Calculate the antiquity of an employee
'Inputs:  oADODBConnection, aEmployeeComponent, lDate, lAntiquityYears, lAntiquityMonths, lAntiquityDays
'Outputs: sErrorDescription, sEmployeeAntiquity
'************************************************************
	Const S_FUNCTION_NAME = "CalculateEmployeeAntiquity"
	Dim lErrorNumber
	Dim lCurrentDate
	Dim lStartDate
	Dim lEndDate
	Dim oRecordset
	Dim sQuery
	Dim iDays
	Dim iInactive
	Dim iDiff

	iInactive = 0
	iDays = 0
	lAntiquityYears = 0
	lAntiquityMonths = 0
	lAntiquityDays = 0
	sEmployeeAntiquity = ""
	If lDate = 0 Then
		lCurrentDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	Else
		lCurrentDate = lDate
	End If
	If (Len(aEmployeeComponent(N_START_DATE_EMPLOYEE)) = 0) Or (aEmployeeComponent(N_START_DATE_EMPLOYEE) = 0) Then aEmployeeComponent(N_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)

	sErrorDescription = "No se pudo obtener la información de los registros."
	If aEmployeeComponent(N_JOB_ID_EMPLOYEE) <> -3 Then
		sQuery = "Select EmployeesHistoryList.StatusID, StatusEmployees.Active, Reasons.ActiveEmployeeID, EmployeesHistoryList.EmployeeDate, " & _
			"EmployeesHistoryList.EndDate, EmployeesHistoryList.Comments, StatusName, EmployeesHistoryList.ReasonID,  " & _
			"ReasonName, EmployeesHistoryList.JobID " & _
			"From EmployeesHistoryList, StatusEmployees, Reasons, Employees, Jobs, Areas " & _
			"Where (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) " & _
			"And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) " & _
			"And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) " & _
			"And (Employees.JobID=Jobs.JobID) " & _
			"And (Jobs.AreaID=Areas.AreaID) " & _
			"And (EmployeesHistoryList.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") " & _
			"And (EmployeesHistoryList.EmployeeDate<=" & lCurrentDate & ") " & _
			"And (EmployeesHistoryList.EndDate>100) " & _
			"Order By EmployeesHistoryList.EndDate Desc"
	Else
		sQuery = "Select EmployeesHistoryList.StatusID, StatusEmployees.Active, Reasons.ActiveEmployeeID, EmployeesHistoryList.EmployeeDate, " & _
			"EmployeesHistoryList.EndDate, EmployeesHistoryList.Comments, StatusName, EmployeesHistoryList.ReasonID,  " & _
			"ReasonName, EmployeesHistoryList.JobID " & _
			"From EmployeesHistoryList, StatusEmployees, Reasons, Employees" & _
			"Where (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) " & _
			"And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) " & _
			"And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) " & _
			"And (EmployeesHistoryList.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") " & _
			"And (EmployeesHistoryList.EmployeeDate<=" & lCurrentDate & ") " & _
			"And (EmployeesHistoryList.EndDate>100) " & _
			"Order By EmployeesHistoryList.EndDate Desc"
	End If
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
					lEndDate = lCurrentDate	
				Else
					lEndDate = oRecordset.Fields("EndDate").Value
				End If
				lStartDate = oRecordset.Fields("EmployeeDate").Value
				iDiff = DateDiff("d", GetDateFromSerialNumber(lStartDate), GetDateFromSerialNumber(lEndDate)) + 1
				iDays = iDays + iDiff
				If (CInt(oRecordset.Fields("ActiveEmployeeID").Value) = 0) Or (CLng(oRecordset.Fields("JobID").Value) = -3) Then
					iInactive = iInactive + iDiff
				End If
				oRecordset.MoveNext
			Loop
		End If
		oRecordset.Close
	End If

	iDays = iDays - iInactive
	lAntiquityYears = Int(iDays / 365)
	iDays = iDays Mod 365
	lAntiquityMonths = Int(iDays / 30.4)
	lAntiquityDays = Int(iDays - (lAntiquityMonths * 30.4))

	If lAntiquityYears > 0 Then sEmployeeAntiquity = lAntiquityYears & " año(s) "
	If lAntiquityMonths > 0 Then sEmployeeAntiquity = sEmployeeAntiquity & lAntiquityMonths & " mes(es) "
	If lAntiquityDays > 0 Then sEmployeeAntiquity = sEmployeeAntiquity & lAntiquityDays & " día(s)" 

	Set oRecordset = Nothing
	CalculateEmployeeAntiquity = lErrorNumber
	Err.Clear
End Function

Function CalculateEmployeeAge(oADODBConnection, aEmployeeComponent, lEmployeeAge, sErrorDescription)
'************************************************************
'Purpose: Calculate the age of an employee
'Inputs:  oADODBConnection, aEmployeeComponent, lAntiquityAge
'Outputs: sErrorDescription
'************************************************************
	Const S_FUNCTION_NAME = "CalculateEmployeeAntiquity"
	Dim lErrorNumber
	Dim sCondition
	Dim iYears
	Dim iMonths
	Dim iDays
	Dim lCurrentDate
	Dim lStartDate
	Dim oRecordset
	Dim sQuery

	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = GetEmployeeStartDate(oADODBConnection, aEmployeeComponent, sErrorDescription)

	lEmployeeAge = 0
	lCurrentDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))

	Call GetAntiquityFromSerialDates(aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE),lCurrentDate, iYears, iMonths, iDays)

	lEmployeeAge = iYears

	CalculateEmployeeAge = lEmployeeAge
	Err.Clear
End Function

Function GetEmployeeDataFromCatalog(oADODBConnection, sCatalog, sFields, sFieldID, lValueID, lPeriod, oRecordsetCatalog, sErrorDescription)
'************************************************************
'Purpose: Get the short name and name of an ID within a catalog
'			according a specified period
'Inputs:  oADODBConnection, sCatalog, sFields, sFieldID, 
'		  lValueID, lPeriod
'Outputs: oRecordsetCatalog, sErrorDescription
'************************************************************
	Const S_FUNCTION_NAME = "CalculateEmployeeAntiquity"
	Dim lErrorNumber
	Dim sQuery
	Dim oRecordsetCatalog2

	lErrorNumber = 0
	sErrorDescription = ""

	sQuery = "Select " & sFields & _
			 " From " & sCatalog & _
			 " Where (" & sFieldID & " = " & lValueID & ")" & _
			 " And (((StartDate <= "& Mid(lPeriod,1,8) & ") And (EndDate >= " & Mid(lPeriod,10) & "))" & _
				" Or ((StartDate <= " & Mid(lPeriod,1,8) & ") And (EndDate <= " & Mid(lPeriod,10) & ") And (EndDate >= " & Mid(lPeriod,1,8) & "))" & _
				" Or ((StartDate >= " & Mid(lPeriod,1,8) & ") And (EndDate <= " & Mid(lPeriod,10) & "))" & _
				" Or ((StartDate >= " & Mid(lPeriod,1,8) & ") And (EndDate >= " & Mid(lPeriod,10) & ") And (StartDate <= " & Mid(lPeriod,10) & ")))"

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordsetCatalog)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			sErrorDescription = "No Disponible"
		End If
	Else
		sErrorDescription = "No se pudieron leer los catálogos del sistema"
	End If

	GetEmployeeDataFromCatalog = lErrorNumber
	Set oRecordsetCatalog = Nothing
	Err.Clear
End Function

Function CheckForFractionatedPeriod(oRequest, oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To check if current movement starts-ends out of
'			payroll limit dates
'Inputs:  oRequest, oADODBConnection, lReasonID, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	Const S_FUNCTION_NAME = "CheckForFractionatedPeriod"
	Dim lErrorNumber
	Dim sQuery
	Dim lPayrollStartDate
	Dim lEmployeeDate
	Dim lEndDate
	Dim lPayrollID
	Dim lDays
	Dim oRecordset
	Dim oHistoryRecordset
	Dim asHistoryRecord
	Dim asIxEmployee
	Dim jIndex
	Dim bStatus

	bStatus = True
	'Índices de EmployeeComponent para extraer datos y comparar contra registros en EHL
	asIxEmployee = Split("8,9,11,12,13,14,15,16,17,18,25,26,27,40,41,67,68,69",",")

	lErrorNumber = 0

	lPayrollID = CLng(aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE))
	lEndDate = AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE),-1)
	lPayrollStartDate = CLng(GetPayrollStartDate(lPayrollID))

	sQuery = "Select EmployeeDate From EmployeesHistoryList Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EndDate =" & lEndDate & ")"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If Not oRecordset.EOF Then
		lEmployeeDate = CLng(oRecordset.Fields("EmployeeDate"))

		If lEmployeeDate >= lPayrollStartDate Then
			lDays = lEndDate - lEmployeeDate + 1
		Else
			lDays = lEndDate - lPayrollStartDate + 1
		End If

		sQuery = "Select PayrollID, EmployeeID, EmployeeDate, DayCounter From EmployeesSpecialCases Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (PayrollID = " & lPayrollID & ") Order By EmployeeDate Asc"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				sQuery = "Insert Into EmployeesSpecialCases (PayrollID, EmployeeID, EmployeeDate, DayCounter) Values (" & lPayrollID & "," & aEmployeeComponent(N_ID_EMPLOYEE) & "," & lEmployeeDate & "," & lDays & ")"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Else
				Do While Not oRecordset.EOF
					sQuery = "Select CompanyID, JobID, ServiceID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, LevelID, PaymentCenterID, StatusID, ReasonID, ZoneID, AreaID, PositionID, EmployeeDate From EmployeesHistoryList Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (PayrollDate = " & lPayrollID & ") And (EndDate > 100) And (Active = 1) And (StatusID <> -2) And (ReasonID <> 0) And (EmployeeDate <> " & oRecordset.Fields("EmployeeDate").Value & ")"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oHistoryRecordset)
					If lErrorNumber = 0 Then
						If Not oHistoryRecordset.EOF Then
							jIndex = 0
							bStatus = True
							For jIndex = 0 To oHistoryRecordset.Fields.Count - 1
								If CLng(oHistoryRecordset.Fields(jIndex).Value) = CLng(aEmployeeComponent(asIxEmployee(jIndex))) Then
									bStatus = True
								Else
									bStatus = False
									Exit For
								End If
							Next
							If bStatus = True Then
								lDays = lDays + CInt(asSpecialCases(iIndex,3))
								sQuery = "Delete EmployeesSpecialCases Where (PayrollID = " & asSpecialCases(iIndex,0) & ") And (EmployeeID = " & asSpecialCases(iIndex,1) & ") And (EmployeeDate = " & asSpecialCases(iIndex,2)  & ")"
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								Exit Do
							End If
						End If
					End If
					oRecordset.MoveNext
				Loop
				If lErrorNumber = 0 Then
					sQuery = "Insert Into EmployeesSpecialCases (PayrollID, EmployeeID, EmployeeDate, DayCounter) Values (" & lPayrollID & "," & aEmployeeComponent(N_ID_EMPLOYEE) & "," & lEmployeeDate & "," & lDays & ")"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
		End If
	Else
		lDays = lEndDate - lPayrollStartDate + 1
		sQuery = "Insert Into EmployeesSpecialCases (PayrollID, EmployeeID, EmployeeDate, DayCounter) Values (" & lPayrollID & "," & aEmployeeComponent(N_ID_EMPLOYEE) & "," & lEmployeeDate & "," & lDays & ")"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	CheckForFractionatedPeriod = lErrorNumber
	Set oRecordsetCatalog = Nothing
	Err.Clear
End Function

Function CheckFor0708Concepts(oADODBConnection, lEmployeeID)
'************************************************************
'Purpose: To check if current movement starts-ends out of
'			payroll limit dates
'Inputs:  oRequest, oADODBConnection, lReasonID, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	Const S_FUNCTION_NAME = "CheckFor0708Concepts"
	Dim lErrorNumber
	Dim sQuery
	Dim lCurrentDate
	Dim oRecordset
	Dim lConceptID

	lCurrentDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	lConceptID = 0

	sQuery = "Select ConceptID From EmployeesConceptsLKP Where (EmployeeID = " & lEmployeeID & ") And (StartDate <" & lCurrentDate & ") And (EndDate >" & lCurrentDate & ") And (ConceptID In (7,8))"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If Not oRecordset.EOF Then
		If CLng(oRecordset.Fields("ConceptID").Value) <> 0 Then
			lConceptID = CLng(oRecordset.Fields("ConceptID").Value)
		End If
	End IF

	CheckFor0708Concepts = lConceptID
	Set oRecordset = Nothing
	Err.Clear

End Function

Function VerifyRecordIntegrity(oADODBConnection,aEmployeeComponent,lReasonID,sErrorDescription)
'************************************************************
'Purpose: To check if the information of ther current record
'			is complete (JobID, PositionID, Service ...)
'Inputs:  oRequest, oADODBConnection, lReasonID, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************	
	Const S_FUNCTION_NAME = "VerifyRecordIntegrity"
	Dim sQuery
	Dim oRecordset
	Dim sErrDescription
	Dim lErrorNumber

	If InStr(oRequest.Item,"ErrorDescription") = 0 Then
		lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
		sErrDescription = "La operación no puede realizarse por la(s) siguiente(s) razón(es):<UL>"
		VerifyRecordIntegrity = true
		If InStr(1, ",12,13,14,17,18,28,57,68,82,", "," & lReasonID & ",", vbBinaryCompare) = 0 Then
			If aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) = -1 Then
				sErrDescription = sErrDescription & "<LI>No se tiene registrada la empresa del empleado.</LI>"
				VerifyRecordIntegrity = false
			End If
			If aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) = -1 Then
				sErrDescription = sErrDescription & "<LI>El empleado no cuenta con centro de pago.</LI>"
				VerifyRecordIntegrity = false
			End If
			If (aEmployeeComponent(N_JOB_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_JOB_ID_EMPLOYEE) = 0) Then
				sErrDescription = sErrDescription & "<LI>El empleado no tiene plaza asignada.</LI>"
				VerifyRecordIntegrity = false
			Else
				aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
				lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					If aJobComponent(N_POSITION_ID_JOB) = -1 Then
						sErrDescription = sErrDescription & "<LI>La plaza (" & aJobComponent(N_ID_JOB) & ") del empleado indicado no tiene puesto.</LI>"
						VerifyRecordIntegrity = false
					End IF
					If aJobComponent(N_JOB_TYPE_ID_JOB) = -1 Then
						sErrDescription = sErrDescription & "<LI>La plaza (" & aJobComponent(N_ID_JOB) & ") del empleado indicado no tiene tipo de plaza asignado.</LI>"
						VerifyRecordIntegrity = false
					End IF
					
					If aJobComponent(N_POSITION_TYPE_ID_JOB) = -1 Then
						sErrDescription = sErrDescription & "<LI>La plaza (" & aJobComponent(N_ID_JOB) & ") del empleado indicado no tiene tipo de puesto asignado.</LI>"
						VerifyRecordIntegrity = false
					End IF

				Else
					VerifyRecordIntegrity = false
					sErrDescription = ""
				End IF
			End If
			sErrDescription = sErrDescription & "</UL>"
		End If
		If Len(sErrDescription) > 0 Then sErrorDescription = sErrDescription
		Set oRecordset = Nothing
		Err.Clear
	Else
		VerifyRecordIntegrity = true
	End If
	Err.Clear

End Function

Function CheckExistencyOfActiveAccount(oADODBConnection,aEmployeeComponent)
'************************************************************
'Purpose: To check if the current employee has a previous active
'			bankaccount registered. If exists and its enddate
'			is earlier than current's movement one then fix it.
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************	
	Const S_FUNCTION_NAME = "CheckExistencyOfActiveAccount"
	Dim sQuery
	Dim oRecordset

	CheckExistencyOfActiveAccount = False
	sQuery = "Select AccountID, EndDate From BankAccounts Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EndDate > " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ")"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If Not oRecordset.EOF Then
		If CLng(oRecordset.Fields("EndDate").Value) < aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) Then
			sQuery = "Update BankAccounts Set EndDate = " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & " Where (AccountID = " & oRecordset.Fields("AccountID").Value & ")"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		'CheckExistencyOfActiveAccount = True
	End If
	Set oRecordset = Nothing
	Err.Clear

End Function
%>