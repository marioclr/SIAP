<%
Const N_EMPLOYEE_ID_ABSENCE = 0
Const N_ABSENCE_ID_ABSENCE = 1
Const N_OCURRED_DATE_ABSENCE = 2
Const N_END_DATE_ABSENCE = 3
Const N_REGISTRATION_DATE_ABSENCE = 4
Const S_DOCUMENT_NUMBER_ABSENCE = 5
Const N_HOURS_ABSENCE = 6
Const N_JUSTIFICATION_ID_ABSENCE = 7
Const N_APPLIES_FOR_PUNCTUALITY_ABSENCE = 8
Const S_REASONS_ABSENCE = 9
Const N_ADD_USER_ID_ABSENCE = 10
Const N_APPLIED_DATE_ABSENCE = 11
Const N_REMOVED_ABSENCE = 12
Const N_REMOVE_USER_ID_ABSENCE = 13
Const N_REMOVED_DATE_ABSENCE = 14
Const N_APPLIED_REMOVE_DATE_ABSENCE = 15
Const N_ACTIVE_ABSENCE = 16
Const B_IS_DUPLICATED_ABSENCE = 17
Const B_COMPONENT_INITIALIZED_ABSENCE = 18
Const S_QUERY_CONDITION_ABSENCE = 19
Const N_VACATION_PERIOD_ABSENCE = 20
Const S_ABSENCE_SHORT_NAME_ABSENCE = 21
Const N_FOR_JUSTIFICATION_ID_ABSENCE = 22

Const N_ABSENCE_COMPONENT_SIZE = 23

Dim aAbsenceComponent()
Redim aAbsenceComponent(N_ABSENCE_COMPONENT_SIZE)

Function InitializeAbsenceComponent(oRequest, aAbsenceComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Absence Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aAbsenceComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeAbsenceComponent"
	Dim oItem
	Redim Preserve aAbsenceComponent(N_EMPLOYEE_COMPONENT_SIZE)

	If IsEmpty(aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE)) Then
		If Len(oRequest("EmployeeID").Item) > 0 Then
			aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = CLng(oRequest("EmployeeID").Item)
		Else
			aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1
		End If
	End If

	If IsEmpty(aAbsenceComponent(N_ABSENCE_ID_ABSENCE)) Then
		If Len(oRequest("AbsenceID").Item) > 0 Then
			aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = CLng(oRequest("AbsenceID").Item)
			Call GetNameFromTable(oADODBConnection, "Absences", aAbsenceComponent(N_ABSENCE_ID_ABSENCE), "", "", aAbsenceComponent(S_ABSENCE_SHORT_NAME_ABSENCE), sErrorDescription)
		Else
			aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1
		End If
	End If

	If IsEmpty(aAbsenceComponent(N_OCURRED_DATE_ABSENCE)) Then
		If Len(oRequest("OcurredYear").Item) > 0 Then
			aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = CInt(oRequest("OcurredYear").Item) & Right(("0" & oRequest("OcurredMonth").Item), Len("00")) & Right(("0" & oRequest("OcurredDay").Item), Len("00"))
		ElseIf Len(oRequest("OcurredDate").Item) > 0 Then
			aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = CLng(oRequest("OcurredDate").Item)
		Else
			aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
		End If
	End If

	If IsEmpty(aAbsenceComponent(N_END_DATE_ABSENCE)) Then
		If Len(oRequest("EndYear").Item) > 0 Then
			aAbsenceComponent(N_END_DATE_ABSENCE) = CInt(oRequest("EndYear").Item) & Right(("0" & oRequest("EndMonth").Item), Len("00")) & Right(("0" & oRequest("EndDay").Item), Len("00"))
		ElseIf Len(oRequest("EndDate").Item) > 0 Then
			aAbsenceComponent(N_END_DATE_ABSENCE) = CLng(oRequest("EndDate").Item)
		Else
			aAbsenceComponent(N_END_DATE_ABSENCE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
		End If
	End If

	If IsEmpty(aAbsenceComponent(N_REGISTRATION_DATE_ABSENCE)) Then
		If Len(oRequest("RegistrationYear").Item) > 0 Then
			aAbsenceComponent(N_REGISTRATION_DATE_ABSENCE) = CInt(oRequest("RegistrationYear").Item) & Right(("0" & oRequest("RegistrationMonth").Item), Len("00")) & Right(("0" & oRequest("RegistrationDay").Item), Len("00"))
		ElseIf Len(oRequest("RegistrationDate").Item) > 0 Then
			aAbsenceComponent(N_REGISTRATION_DATE_ABSENCE) = CLng(oRequest("RegistrationDate").Item)
		Else
			aAbsenceComponent(N_REGISTRATION_DATE_ABSENCE) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
		End If
	End If

	If IsEmpty(aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE)) Then
		If Len(oRequest("DocumentNumber").Item) > 0 Then
			aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE) = oRequest("DocumentNumber").Item
		Else
			aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE) = ""
		End If
	End If
	aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE) = Left(aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE), 50)

	If IsEmpty(aAbsenceComponent(N_HOURS_ABSENCE)) Then
		If Len(oRequest("AbsenceHours").Item) > 0 Then
			aAbsenceComponent(N_HOURS_ABSENCE) = CInt(oRequest("AbsenceHours").Item)
		Else
			aAbsenceComponent(N_HOURS_ABSENCE) = 0
		End If
	End If

	If IsEmpty(aAbsenceComponent(N_JUSTIFICATION_ID_ABSENCE)) Then
		If Len(oRequest("JustificationID").Item) > 0 Then
			aAbsenceComponent(N_JUSTIFICATION_ID_ABSENCE) = CLng(oRequest("JustificationID").Item)
		Else
			aAbsenceComponent(N_JUSTIFICATION_ID_ABSENCE) = -1
		End If
	End If

	If IsEmpty(aAbsenceComponent(N_APPLIES_FOR_PUNCTUALITY_ABSENCE)) Then
		If Len(oRequest("AppliesForPunctuality").Item) > 0 Then
			aAbsenceComponent(N_APPLIES_FOR_PUNCTUALITY_ABSENCE) = CInt(oRequest("AppliesForPunctuality").Item)
		Else
			aAbsenceComponent(N_APPLIES_FOR_PUNCTUALITY_ABSENCE) = 1
		End If
	End If

	If IsEmpty(aAbsenceComponent(S_REASONS_ABSENCE)) Then
		If Len(oRequest("ForReasons").Item) > 0 Then
			If Len(oRequest("Reasons").Item) > 0 Then
				aAbsenceComponent(S_REASONS_ABSENCE) = oRequest("Reasons").Item
			Else
				aAbsenceComponent(S_REASONS_ABSENCE) = ""
			End If
		Else
			aAbsenceComponent(S_REASONS_ABSENCE) = ""
		End If
	End If
	aAbsenceComponent(S_REASONS_ABSENCE) = Left(aAbsenceComponent(S_REASONS_ABSENCE), 2000)

	If IsEmpty(aAbsenceComponent(N_ADD_USER_ID_ABSENCE)) Then
		If Len(oRequest("AddUserID").Item) > 0 Then
			aAbsenceComponent(N_ADD_USER_ID_ABSENCE) = CLng(oRequest("AddUserID").Item)
		Else
			aAbsenceComponent(N_ADD_USER_ID_ABSENCE) = aLoginComponent(N_USER_ID_LOGIN)
		End If
	End If

	If IsEmpty(aAbsenceComponent(N_APPLIED_DATE_ABSENCE)) Then
		If Len(oRequest("AppliedYear").Item) > 0 Then
			aAbsenceComponent(N_APPLIED_DATE_ABSENCE) = CLng(oRequest("AppliedYear").Item & Right(("0" & oRequest("AppliedMonth").Item), Len("00")) & Right(("0" & oRequest("AppliedDay").Item), Len("00")))
		ElseIf Len(oRequest("AppliedDate").Item) > 0 Then
			aAbsenceComponent(N_APPLIED_DATE_ABSENCE) = CLng(oRequest("AppliedDate").Item)
		Else
			aAbsenceComponent(N_APPLIED_DATE_ABSENCE) = 0
		End If
	End If

	If IsEmpty(aAbsenceComponent(N_REMOVED_ABSENCE)) Then
		If Len(oRequest("Removed").Item) > 0 Then
			aAbsenceComponent(N_REMOVED_ABSENCE) = CInt(oRequest("Removed").Item)
		Else
			aAbsenceComponent(N_REMOVED_ABSENCE) = 0
		End If
	End If

	If IsEmpty(aAbsenceComponent(N_REMOVE_USER_ID_ABSENCE)) Then
		If Len(oRequest("RemoveUserID").Item) > 0 Then
			aAbsenceComponent(N_REMOVE_USER_ID_ABSENCE) = CLng(oRequest("RemoveUserID").Item)
		'ElseIf Len(oRequest("Remove").Item) > 0 Then
		'	aAbsenceComponent(N_REMOVE_USER_ID_ABSENCE) = aLoginComponent(N_USER_ID_LOGIN)
		Else
			aAbsenceComponent(N_REMOVE_USER_ID_ABSENCE) = aLoginComponent(N_USER_ID_LOGIN)
		End If
	End If

	If IsEmpty(aAbsenceComponent(N_REMOVED_DATE_ABSENCE)) Then
		If Len(oRequest("RemovedYear").Item) > 0 Then
			aAbsenceComponent(N_REMOVED_DATE_ABSENCE) = CLng(oRequest("RemovedYear").Item & Right(("0" & oRequest("RemovedMonth").Item), Len("00")) & Right(("0" & oRequest("RemovedDay").Item), Len("00")))
		ElseIf Len(oRequest("RemovedDate").Item) > 0 Then
			aAbsenceComponent(N_REMOVED_DATE_ABSENCE) = CLng(oRequest("RemovedDate").Item)
		Else
			aAbsenceComponent(N_REMOVED_DATE_ABSENCE) = 0
		End If
	End If

	If IsEmpty(aAbsenceComponent(N_APPLIED_REMOVE_DATE_ABSENCE)) Then
		If Len(oRequest("AppliedRemoveYear").Item) > 0 Then
			aAbsenceComponent(N_APPLIED_REMOVE_DATE_ABSENCE) = CLng(oRequest("AppliedRemoveYear").Item & Right(("0" & oRequest("AppliedRemoveMonth").Item), Len("00")) & Right(("0" & oRequest("AppliedRemoveDay").Item), Len("00")))
		ElseIf Len(oRequest("AppliedRemoveDate").Item) > 0 Then
			aAbsenceComponent(N_APPLIED_REMOVE_DATE_ABSENCE) = CLng(oRequest("AppliedRemoveDate").Item)
		Else
			aAbsenceComponent(N_APPLIED_REMOVE_DATE_ABSENCE) = 0
		End If
	End If

	If IsEmpty(aAbsenceComponent(N_ACTIVE_ABSENCE)) Then
		If Len(oRequest("Active").Item) > 0 Then
			aAbsenceComponent(N_ACTIVE_ABSENCE) = CInt(oRequest("Active").Item)
		Else
			aAbsenceComponent(N_ACTIVE_ABSENCE) = 0
		End If
	End If

	If IsEmpty(aAbsenceComponent(N_VACATION_PERIOD_ABSENCE)) Then
		If Len(oRequest("YearID").Item) > 0 Then
			Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
				Case 35, 37, 38
					aAbsenceComponent(N_VACATION_PERIOD_ABSENCE) = CLng(oRequest("YearID").Item) & CInt(oRequest("PeriodVacationID").Item)
				Case 39, 40
					aAbsenceComponent(N_VACATION_PERIOD_ABSENCE) = CLng(CStr(oRequest("YearID").Item) & CStr(Right("0" & oRequest("PeriodVacationID").Item, Len("MM"))))
				Case Else
					aAbsenceComponent(N_VACATION_PERIOD_ABSENCE) = 0
			End Select
		Else
			aAbsenceComponent(N_VACATION_PERIOD_ABSENCE) = 0
		End If
	End If

	If IsEmpty(aAbsenceComponent(N_FOR_JUSTIFICATION_ID_ABSENCE)) Then
		If Len(oRequest("ForJustificationID").Item) > 0 Then
			aAbsenceComponent(N_FOR_JUSTIFICATION_ID_ABSENCE) = CLng(oRequest("ForJustificationID").Item)
		Else
			aAbsenceComponent(N_FOR_JUSTIFICATION_ID_ABSENCE) = -1
		End If
	End If

	aAbsenceComponent(B_IS_DUPLICATED_ABSENCE) = False
	aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE) = True
	aAbsenceComponent(S_QUERY_CONDITION_ABSENCE) = ""
	InitializeAbsenceComponent = Err.number
	Err.Clear
End Function

Function AddAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new absence for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddAbsence"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sAbsenceIDs
	Dim sAbsenceShortName
	Dim lDate
	Dim bIsForPeriod
	Dim iJourneyTypeID

	bIsForPeriod = True
	lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador de la incidencia y/o la fecha para agregar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If (Not VerifyAbsencesForPeriod(oADODBConnection, aAbsenceComponent, sErrorDescription) Or ((aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)=21) Or (aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)=22) Or (aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE)=23))) Then
			bIsForPeriod = False
			Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
				Case 34, 35, 37, 38, 39
				Case Else
					If aAbsenceComponent(N_OCURRED_DATE_ABSENCE) >= lDate Then
						lErrorNumber = -1
						sErrorDescription = "La fecha de registro " & GetDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE)) & " de la incidencia no puede ser mayor o igual a la fecha del día"
					End If
			End Select
			aAbsenceComponent(N_END_DATE_ABSENCE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
			aAbsenceComponent(N_HOURS_ABSENCE) = 1
			Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
				Case 35, 37, 38, 39, 40
				Case Else
					aAbsenceComponent(N_VACATION_PERIOD_ABSENCE) = 0
			End Select
		Else
			Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
				Case 29, 30, 35, 37, 38
					Call GetEmployeeJourneyType(oRequest, oADODBConnection, aEmployeeComponent, iJourneyTypeID, sErrorDescription)
					'Select Case iJourneyTypeID
					'	Case 1
					'		aAbsenceComponent(N_HOURS_ABSENCE) = GetWorkingDaysOfAbsencesPeriod(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), aAbsenceComponent(N_END_DATE_ABSENCE), 1)
					'	Case 2, 3
					'		aAbsenceComponent(N_HOURS_ABSENCE) = GetWorkingDaysOfAbsencesPeriod(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), aAbsenceComponent(N_END_DATE_ABSENCE), 2)
					'	Case 4
					'		aAbsenceComponent(N_HOURS_ABSENCE) = GetWorkingDaysOfAbsencesPeriod(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), aAbsenceComponent(N_END_DATE_ABSENCE), 0)
					'End Select
					aAbsenceComponent(N_HOURS_ABSENCE) = GetWorkingDaysOfAbsencesPeriod(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), aAbsenceComponent(N_END_DATE_ABSENCE), iJourneyTypeID)
					If aAbsenceComponent(N_HOURS_ABSENCE) = 0 Then
						lErrorNumber = -1
						sErrorDescription = "En las fechas introducidas el empleado no tiene días laborables para ser registrados como vacaciones."
					End If
				Case Else
					If aAbsenceComponent(N_END_DATE_ABSENCE) = 30000000 Then
						aAbsenceComponent(N_HOURS_ABSENCE) = 1
					Else
						aAbsenceComponent(N_HOURS_ABSENCE) = DateDiff("d", GetDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE)), GetDateFromSerialNumber(aAbsenceComponent(N_END_DATE_ABSENCE))) + 1
					End If
			End Select
		End If
		If lErrorNumber = 0 Then
			If aAbsenceComponent(N_END_DATE_ABSENCE) < aAbsenceComponent(N_OCURRED_DATE_ABSENCE) Then
				lErrorNumber = -1
				sErrorDescription = "La fecha de fin (" & DisplayDateFromSerialNumber(aAbsenceComponent(N_END_DATE_ABSENCE), -1, -1, -1) & ") no debe de ser menor a la fecha de inicio (" & DisplayDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), -1, -1, -1) & ")"
			End If
		End If
		If lErrorNumber = 0 Then
			If VerifyRequerimentsForEmployeesAbsences(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
				aAbsenceComponent(B_IS_DUPLICATED_ABSENCE) = False
				lErrorNumber = CheckExistencyOfAbsence(aAbsenceComponent, bIsForPeriod, sErrorDescription)
				If lErrorNumber = 0 Then
					If aAbsenceComponent(B_IS_DUPLICATED_ABSENCE) Then
						lErrorNumber = L_ERR_DUPLICATED_RECORD
						Call GetNameFromTable(oADODBConnection, "Absences1", aAbsenceComponent(N_ABSENCE_ID_ABSENCE), "", "", sAbsenceShortName, "")
						sErrorDescription = "Ya existe un registro de la clave " & sAbsenceShortName & " el día " & DisplayDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), -1, -1, -1) & " para el empleado " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE)
						Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
					Else
						If Not CheckAbsenceInformationConsistency(aAbsenceComponent, sErrorDescription) Then
							lErrorNumber = -1
						Else
							lErrorNumber = GetAbsenceAppliesToID(oRequest, oADODBConnection, aAbsenceComponent, sAbsenceIDs, sErrorDescription)
							If lErrorNumber = 0 Then
								If VerifyExistenceOfAbsences(oADODBConnection, aAbsenceComponent, sAbsenceIDs, bIsForPeriod, sErrorDescription) Then
									sErrorDescription = "No se pudo guardar la información del registro."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesAbsencesLKP (EmployeeID, AbsenceID, OcurredDate, EndDate, RegistrationDate, DocumentNumber, AbsenceHours, JustificationID, AppliesForPunctuality, Reasons, AddUserID, AppliedDate, Removed, RemoveUserID, RemovedDate, AppliedRemoveDate, Active, VacationPeriod) Values (" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ", " & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ", " & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ", " & aAbsenceComponent(N_END_DATE_ABSENCE) & ", " & aAbsenceComponent(N_REGISTRATION_DATE_ABSENCE) & ", '" & Replace(aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE), "'", "´") & "', " & aAbsenceComponent(N_HOURS_ABSENCE) & ", " & aAbsenceComponent(N_JUSTIFICATION_ID_ABSENCE) & ", " & aAbsenceComponent(N_APPLIES_FOR_PUNCTUALITY_ABSENCE) & ", '" & Replace(aAbsenceComponent(S_REASONS_ABSENCE), "'", "´") & "', " & aAbsenceComponent(N_ADD_USER_ID_ABSENCE) & ", " & aAbsenceComponent(N_APPLIED_DATE_ABSENCE) & ", " & aAbsenceComponent(N_REMOVED_ABSENCE) & ", " &  aAbsenceComponent(N_REMOVE_USER_ID_ABSENCE) & ", " &  aAbsenceComponent(N_REMOVED_DATE_ABSENCE) & ", " & aAbsenceComponent(N_APPLIED_REMOVE_DATE_ABSENCE) & ", " & aAbsenceComponent(N_ACTIVE_ABSENCE) & ", " & aAbsenceComponent(N_VACATION_PERIOD_ABSENCE) & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								Else
									lErrorNumber = -1
								End If
							Else
								sErrorDescription = "Error al validar las incidencias registradas."
							End If
						End If
					End If
				End If
			Else
				lErrorNumber = -1
			End If
		End If
	End If

	AddAbsence = lErrorNumber
	Err.Clear
End Function

Function AddAbsenceFile(oRequest, oADODBConnection, sQuery, lReasonID, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new absence for the employee into the database
'Inputs:  oRequest, oADODBConnection, sQuery, lReasonID
'Outputs: aAbsenceComponent, aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddAbsenceFile"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información de la aplicación de incidencias masivos de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				If Not IsEmpty(oRequest(CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("AbsenceID").Value) & CStr(oRecordset.Fields("OcurredDate").Value))) Then
					aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = CLng(oRecordset.Fields("EmployeeID").Value)
					aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = CLng(oRecordset.Fields("AbsenceID").Value)
					aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = CLng(oRecordset.Fields("OcurredDate").Value)
					aAbsenceComponent(N_APPLIED_DATE_ABSENCE) = CLng(oRequest("AppliedDate").Item)
					lErrorNumber = SetActiveForEmployeeAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
				End If
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
		End If
	End If

	Set oRecordset = Nothing
	AddAbsenceFile = lErrorNumber
	Err.Clear
End Function

Function AddJustification(oRequest, oADODBConnection, iAbsenceID, iActiveOriginal, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To justify an existent absence for the employee into the database
'Inputs:  oRequest, oADODBConnection, iAbsenceID, iActiveOriginal, aAbsenceComponent
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddJustification"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sAbsenceIDs
	Dim lDate
	Dim bIsForPeriod

	bIsForPeriod = True
	lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador de la incidencia y/o la fecha para agregar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo justificar la incidencia del día " + CStr(GetDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE)))
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set DocumentNumber='" & aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE) & "', JustificationID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ", AppliesForPunctuality=" & aAbsenceComponent(N_APPLIES_FOR_PUNCTUALITY_ABSENCE) & ", Removed=1, RemoveUserID=" & aAbsenceComponent(N_REMOVE_USER_ID_ABSENCE) & ", RemovedDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", AppliedRemoveDate=" & aAbsenceComponent(N_APPLIED_REMOVE_DATE_ABSENCE) & " Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & iAbsenceID & ") And (OcurredDate=" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	AddJustification = lErrorNumber
	Err.Clear
End Function

Function AddSuspension(oRequest, oADODBConnection, iAbsenceID, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new absence for the employee into the database
'Inputs:  oRequest, oADODBConnection, iAbsenceID, aAbsenceComponent
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddSuspension"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sAbsenceIDs
	Dim lDate
	Dim bIsForPeriod

	bIsForPeriod = True
	lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Or (aAbsenceComponent(N_END_DATE_ABSENCE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador de la incidencia y/o la fecha para agregar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		GetEmployee()
		sErrorDescription = "No se pudo justificar la incidencia del día " + aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set JustificationID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ", AppliesForPunctuality=" & aAbsenceComponent(N_APPLIES_FOR_PUNCTUALITY_ABSENCE) & ", Removed=" & aAbsenceComponent(N_REMOVED_ABSENCE) & ", RemoveUserID=" & aAbsenceComponent(N_REMOVE_USER_ID_ABSENCE) & ", RemovedDate=" & aAbsenceComponent(N_REGISTRATION_DATE_ABSENCE) & ", AppliedRemoveDate=" & aAbsenceComponent(N_APPLIED_DATE_ABSENCE) & ", Active=" & aAbsenceComponent(N_ACTIVE_ABSENCE) & " Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & iAbsenceID & ") And (OcurredDate=" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	AddSuspension = lErrorNumber
	Err.Clear
End Function

Function ApplyAbsencesInProcess(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To apply pending absences for the employees
'Inputs:  oRequest, oADODBConnection, sQuery, lReasonID
'Outputs: aAbsenceComponent, aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ApplyAbsencesInProcess"
	Dim oRecordset
	Dim lErrorNumber
	Dim sQuery

	'Aplicar Suspensiones
	sQuery = "Update EmployeesAbsencesLKP set Active = 1, AppliedDate =" & CLng(oRequest("AppliedDate").Item) & _
		" where (Active = 0)"
	If aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1 Then
		sQuery = sQuery & " And (AbsenceID < 100)"
	Else
		sQuery = sQuery & " And (AbsenceID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ")"
	End If 
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber <> 0 Then
		sErrorDescription = "Error al aplicar las incidencias en proceso."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	End If

	Set oRecordset = Nothing
	ApplyAbsencesInProcess = lErrorNumber
	Err.Clear
End Function

Function ApplyAttendanceControlInProcess(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To apply pending suspensions for the employees
'Inputs:  oRequest, oADODBConnection, sQuery, lReasonID
'Outputs: aAbsenceComponent, aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ApplyAttendanceControlInProcess"
	Dim oRecordset
	Dim oRecordset1
	Dim lErrorNumber
	Dim sQuery
	Dim sAttendanceID

	sAttendanceID = "50, 51, 52, 53, 54, 55, 56"
	sQuery = "Select * from EmployeesAbsencesLKP" & _
			 " where (Active = 0) And (AbsenceID IN (" & sAttendanceID & "))"

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				If (CInt(oRecordset.Fields("AbsenceID").Value) = 52) Or (CInt(oRecordset.Fields("AbsenceID").Value) = 53) Then
					'1) Si existe algun tipo de registro de asistencia en el mismo periodo eliminarlo puesto que se insertara el nuevo
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAbsencesLKP Where (EmployeeID=" & oRecordset.Fields("EmployeeID").Value & ") And (AbsenceID IN(" & sAttendanceID & ")) And (OcurredDate>=" & oRecordset.Fields("OcurredDate").Value & ") And (EndDate<=" & oRecordset.Fields("EndDate").Value & ") And (Active=1)", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset1)
					If ((lErrorNumber = 0) And (Not oRecordset1.EOF)) Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete EmployeesAbsencesLKP Where (EmployeeID=" & oRecordset1.Fields("EmployeeID").Value & ") And (AbsenceID IN(" & sAttendanceID & ")) And (OcurredDate>=" & oRecordset.Fields("OcurredDate").Value & ") And (EndDate<=" & oRecordset.Fields("EndDate").Value & ") And (Active=1)", "AbscenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					End If
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAbsencesLKP Where (EmployeeID=" & oRecordset.Fields("EmployeeID").Value & ") And (AbsenceID IN(" & sAttendanceID & ")) And (OcurredDate<" & oRecordset.Fields("OcurredDate").Value & ") And (EndDate>" & oRecordset.Fields("EndDate").Value & ") And (Active=1) Order by OcurredDate Desc", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset1)
					If ((lErrorNumber = 0) And (Not oRecordset1.EOF)) Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesAbsencesLKP (EmployeeID, AbsenceID, OcurredDate, EndDate, RegistrationDate, DocumentNumber, AbsenceHours, JustificationID, AppliesForPunctuality, Reasons, AddUserID, AppliedDate, Removed, RemoveUserID, RemovedDate, AppliedRemoveDate, Active, VacationPeriod) Values (" & oRecordset1.Fields("EmployeeID").Value & ", " & oRecordset1.Fields("AbsenceID").Value & ", " & AddDaysToSerialDate(CLng(oRecordset.Fields("EndDate").Value), 1) & ", " & oRecordset1.Fields("EndDate").Value & ", " & oRecordset1.Fields("RegistrationDate").Value & ", '" & CStr(oRecordset1.Fields("DocumentNumber").Value) & "', " & oRecordset1.Fields("AbsenceHours").Value & ", " & oRecordset1.Fields("JustificationID").Value & ", " & oRecordset1.Fields("AppliesForPunctuality").Value & ", '" & CStr(oRecordset1.Fields("Reasons").Value) & "', " & oRecordset1.Fields("AddUserID").Value & ", " & oRecordset1.Fields("AppliedDate").Value & ", " & oRecordset1.Fields("Removed").Value & ", " & oRecordset1.Fields("RemoveUserID").Value & ", " &  oRecordset1.Fields("RemovedDate").Value & ", " &  oRecordset1.Fields("AppliedRemoveDate").Value & ", " & oRecordset1.Fields("Active").Value & ", " & oRecordset1.Fields("VacationPeriod").Value & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set EndDate=" & AddDaysToSerialDate(CLng(oRecordset.Fields("OcurredDate").Value), -1) & " Where (EmployeeID=" & oRecordset1.Fields("EmployeeID").Value & ") And (AbsenceID=" & oRecordset1.Fields("AbsenceID").Value & ") And (OcurredDate=" & CLng(oRecordset1.Fields("OcurredDate").Value) & ")", "AbscenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					End If
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAbsencesLKP Where (EmployeeID=" & oRecordset.Fields("EmployeeID").Value & ") And (AbsenceID IN(" & sAttendanceID & ")) And (OcurredDate<" & oRecordset.Fields("OcurredDate").Value & ") And (EndDate<=" & oRecordset.Fields("EndDate").Value & ") And (Active=1) Order by OcurredDate Desc", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset1)
					If ((lErrorNumber = 0) And (Not oRecordset1.EOF)) Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set EndDate=" & AddDaysToSerialDate(CLng(oRecordset.Fields("OcurredDate").Value), -1) & " Where (EmployeeID=" & oRecordset1.Fields("EmployeeID").Value & ") And (AbsenceID=" & oRecordset1.Fields("AbsenceID").Value & ") And (OcurredDate=" & CLng(oRecordset1.Fields("OcurredDate").Value) & ")", "AbscenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					End If
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAbsencesLKP Where (EmployeeID=" & oRecordset.Fields("EmployeeID").Value & ") And (AbsenceID IN(" & sAttendanceID & ")) And (OcurredDate>" & oRecordset.Fields("OcurredDate").Value & ") And (OcurredDate<" & oRecordset.Fields("EndDate").Value & ") And (EndDate>" & oRecordset.Fields("EndDate").Value & ") And (Active=1) Order by OcurredDate Desc", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset1)
					If ((lErrorNumber = 0) And (Not oRecordset1.EOF)) Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set OcurredDate=" & AddDaysToSerialDate(CLng(oRecordset.Fields("EndDate").Value), 1) & " Where (EmployeeID=" & oRecordset1.Fields("EmployeeID").Value & ") And (AbsenceID=" & oRecordset1.Fields("AbsenceID").Value & ") And (OcurredDate=" & CLng(oRecordset1.Fields("OcurredDate").Value) & ")", "AbscenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					End If
					sErrorDescription = "No se pudo aplicar la suspensión del empleado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set Active=1 Where (EmployeeID=" & oRecordset.Fields("EmployeeID").Value & ") And (AbsenceID=" & oRecordset.Fields("AbsenceID").Value & ") And (OcurredDate=" & oRecordset.Fields("OcurredDate").Value & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAbsencesLKP Where (EmployeeID=" & oRecordset.Fields("EmployeeID").Value & ") And (AbsenceID IN(" & sAttendanceID & ")) And (EndDate>=" & oRecordset.Fields("OcurredDate").Value & ") And (Active=1) Order by OcurredDate Desc", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset1)
					If ((lErrorNumber = 0) And (Not oRecordset1.EOF)) Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set EndDate=" & AddDaysToSerialDate(CLng(oRecordset.Fields("OcurredDate").Value), -1) & " Where (EmployeeID=" & oRecordset1.Fields("EmployeeID").Value & ") And (AbsenceID=" & oRecordset1.Fields("AbsenceID").Value & ") And (OcurredDate=" & CLng(oRecordset1.Fields("OcurredDate").Value) & ")", "AbscenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					End If
					sErrorDescription = "No se pudo aplicar la suspensión del empleado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set Active=1 Where (EmployeeID=" & oRecordset.Fields("EmployeeID").Value & ") And (AbsenceID=" & oRecordset.Fields("AbsenceID").Value & ") And (OcurredDate=" & oRecordset.Fields("OcurredDate").Value & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				End If
				oRecordset.MoveNext
				If Err.Number <> 0 Then Exit Do
			Loop
		End If
	End If
	If lErrorNumber <> 0 Then
		sErrorDescription = "Error al aplicar las suspensiones en proceso."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	End If

	Set oRecordset = Nothing
	ApplyAttendanceControlInProcess = lErrorNumber
	Err.Clear
End Function

Function ApplySuspensionsInProcess(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To apply pending suspensions for the employees
'Inputs:  oRequest, oADODBConnection, sQuery, lReasonID
'Outputs: aAbsenceComponent, aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ApplySuspensionsInProcess"
	Dim oRecordset
	Dim oRecordset1
	Dim lErrorNumber
	Dim sQuery
	Dim sSuspensionsID

	sSuspensionsID = "41,42,43,44,45,46,47,48,49,57,58"
	sQuery = "Select * from EmployeesAbsencesLKP" & _
			 " where (Active = 0) And (AbsenceID IN (" & sSuspensionsID & "))"

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAbsencesLKP Where (EmployeeID=" & oRecordset.Fields("EmployeeID").Value & ") And (AbsenceID IN(" & sSuspensionsID & ")) And (OcurredDate>=" & oRecordset.Fields("OcurredDate").Value & ") And (EndDate<=" & oRecordset.Fields("EndDate").Value & ") And (Active=1)", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset1)
				If ((lErrorNumber = 0) And (Not oRecordset1.EOF)) Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete EmployeesAbsencesLKP Where (EmployeeID=" & oRecordset1.Fields("EmployeeID").Value & ") And (AbsenceID IN(" & sSuspensionsID & ")) And (OcurredDate>=" & oRecordset.Fields("OcurredDate").Value & ") And (EndDate<=" & oRecordset.Fields("EndDate").Value & ") And (Active=1)", "AbscenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				End If
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAbsencesLKP Where (EmployeeID=" & oRecordset.Fields("EmployeeID").Value & ") And (AbsenceID IN(" & sSuspensionsID & ")) And (OcurredDate<" & oRecordset.Fields("OcurredDate").Value & ") And (EndDate>" & oRecordset.Fields("EndDate").Value & ") And (Active=1) Order by OcurredDate Desc", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset1)
				If ((lErrorNumber = 0) And (Not oRecordset1.EOF)) Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesAbsencesLKP (EmployeeID, AbsenceID, OcurredDate, EndDate, RegistrationDate, DocumentNumber, AbsenceHours, JustificationID, AppliesForPunctuality, Reasons, AddUserID, AppliedDate, Removed, RemoveUserID, RemovedDate, AppliedRemoveDate, Active, VacationPeriod) Values (" & oRecordset1.Fields("EmployeeID").Value & ", " & oRecordset1.Fields("AbsenceID").Value & ", " & AddDaysToSerialDate(CLng(oRecordset.Fields("EndDate").Value), 1) & ", " & oRecordset1.Fields("EndDate").Value & ", " & oRecordset1.Fields("RegistrationDate").Value & ", '" & CStr(oRecordset1.Fields("DocumentNumber").Value) & "', " & oRecordset1.Fields("AbsenceHours").Value & ", " & oRecordset1.Fields("JustificationID").Value & ", " & oRecordset1.Fields("AppliesForPunctuality").Value & ", '" & CStr(oRecordset1.Fields("Reasons").Value) & "', " & oRecordset1.Fields("AddUserID").Value & ", " & oRecordset1.Fields("AppliedDate").Value & ", " & oRecordset1.Fields("Removed").Value & ", " & oRecordset1.Fields("RemoveUserID").Value & ", " &  oRecordset1.Fields("RemovedDate").Value & ", " &  oRecordset1.Fields("AppliedRemoveDate").Value & ", " & oRecordset1.Fields("Active").Value & ", " & oRecordset1.Fields("VacationPeriod").Value & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set EndDate=" & AddDaysToSerialDate(CLng(oRecordset.Fields("OcurredDate").Value), -1) & " Where (EmployeeID=" & oRecordset1.Fields("EmployeeID").Value & ") And (AbsenceID=" & oRecordset1.Fields("AbsenceID").Value & ") And (OcurredDate=" & CLng(oRecordset1.Fields("OcurredDate").Value) & ")", "AbscenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				End If
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAbsencesLKP Where (EmployeeID=" & oRecordset.Fields("EmployeeID").Value & ") And (AbsenceID IN(" & sSuspensionsID & ")) And (OcurredDate<" & oRecordset.Fields("OcurredDate").Value & ") And (EndDate<=" & oRecordset.Fields("EndDate").Value & ") And (Active=1) Order by OcurredDate Desc", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset1)
				If ((lErrorNumber = 0) And (Not oRecordset1.EOF)) Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set EndDate=" & AddDaysToSerialDate(CLng(oRecordset.Fields("OcurredDate").Value), -1) & " Where (EmployeeID=" & oRecordset1.Fields("EmployeeID").Value & ") And (AbsenceID=" & oRecordset1.Fields("AbsenceID").Value & ") And (OcurredDate=" & CLng(oRecordset1.Fields("OcurredDate").Value) & ")", "AbscenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				End If
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAbsencesLKP Where (EmployeeID=" & oRecordset.Fields("EmployeeID").Value & ") And (AbsenceID IN(" & sSuspensionsID & ")) And (OcurredDate>" & oRecordset.Fields("OcurredDate").Value & ") And (OcurredDate<" & oRecordset.Fields("EndDate").Value & ") And (EndDate>" & oRecordset.Fields("EndDate").Value & ") And (Active=1) Order by OcurredDate Desc", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset1)
				If ((lErrorNumber = 0) And (Not oRecordset1.EOF)) Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set OcurredDate=" & AddDaysToSerialDate(CLng(oRecordset.Fields("EndDate").Value), 1) & " Where (EmployeeID=" & oRecordset1.Fields("EmployeeID").Value & ") And (AbsenceID=" & oRecordset1.Fields("AbsenceID").Value & ") And (OcurredDate=" & CLng(oRecordset1.Fields("OcurredDate").Value) & ")", "AbscenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				End If
				sErrorDescription = "No se pudo aplicar la suspensión del empleado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set Active=1 Where (EmployeeID=" & oRecordset.Fields("EmployeeID").Value & ") And (AbsenceID=" & oRecordset.Fields("AbsenceID").Value & ") And (OcurredDate=" & oRecordset.Fields("OcurredDate").Value & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				oRecordset.MoveNext
				If Err.Number <> 0 Then Exit Do
			Loop
		End If
	End If
	If lErrorNumber <> 0 Then
		sErrorDescription = "Error al aplicar las suspensiones en proceso."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	End If

	Set oRecordset = Nothing
	ApplySuspensionsInProcess = lErrorNumber
	Err.Clear
End Function

Function CancelJustification(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To remove an absence for the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CancelJustification"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim iActive
	Dim iActiveStatus

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador del concepto y/o la fecha para eliminar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If VerifyExistenceOfJustificationForCancel(oADODBConnection, aAbsenceComponent, iActiveStatus, sErrorDescription) Then
			If iActiveStatus = -1 Then
				iActive = 0
			Else
				iActive = 1
			End If
			sErrorDescription = "No se pudo cancelar la incidencia del día " + CStr(GetDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE)))
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set JustificationID=-1, DocumentNumber='', Removed=0, RemoveUserID=-1, RemovedDate=0, AppliedRemoveDate=0, Active=" & iActive & " Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ") And (OcurredDate=" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Else
			sErrorDescription = "No existe incidencia justificada que se pueda cancelar con los criterios seleccionados."
			lErrorNumber = -1
		End If
	End If

	CancelJustification = lErrorNumber
	Err.Clear
End Function

Function GetAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about an absence for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAbsence"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador del concepto y/o la fecha para obtener la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAbsencesLKP Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ") And (OcurredDate=" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aAbsenceComponent(N_END_DATE_ABSENCE) = CLng(oRecordset.Fields("EndDate").Value)
				aAbsenceComponent(N_REGISTRATION_DATE_ABSENCE) = CLng(oRecordset.Fields("RegistrationDate").Value)
				aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE) = CStr(oRecordset.Fields("DocumentNumber").Value)
				aAbsenceComponent(N_HOURS_ABSENCE) = CInt(oRecordset.Fields("AbsenceHours").Value)
				aAbsenceComponent(N_JUSTIFICATION_ID_ABSENCE) = CLng(oRecordset.Fields("JustificationID").Value)
				aAbsenceComponent(N_APPLIES_FOR_PUNCTUALITY_ABSENCE) = CInt(oRecordset.Fields("AppliesForPunctuality").Value)
				aAbsenceComponent(S_REASONS_ABSENCE) = CStr(oRecordset.Fields("Reasons").Value)
				aAbsenceComponent(N_ADD_USER_ID_ABSENCE) = CLng(oRecordset.Fields("AddUserID").Value)
				aAbsenceComponent(N_APPLIED_DATE_ABSENCE) = CLng(oRecordset.Fields("AppliedDate").Value)
				aAbsenceComponent(N_REMOVED_ABSENCE) = CInt(oRecordset.Fields("Removed").Value)
				aAbsenceComponent(N_REMOVE_USER_ID_ABSENCE) = CLng(oRecordset.Fields("RemoveUserID").Value)
				aAbsenceComponent(N_REMOVED_DATE_ABSENCE) = CLng(oRecordset.Fields("RemovedDate").Value)
				aAbsenceComponent(N_APPLIED_REMOVE_DATE_ABSENCE) = CLng(oRecordset.Fields("AppliedRemoveDate").Value)
				aAbsenceComponent(N_ACTIVE_ABSENCE) = CLng(oRecordset.Fields("Active").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetAbsence = lErrorNumber
	Err.Clear
End Function

Function GetAbsences(oRequest, oADODBConnection, aAbsenceComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the absences for
'         the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAbsenceComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAbsences"
	Dim sTables
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		sTables = ", Jobs"
		sCondition = "And (Employees.JobID=Jobs.JobID) And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Jobs.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	Else
		sTables = ""
		sCondition = ""
	End If
	If CLng(oRequest("AbsenceID").Item) > 0 Then
		sCondition = sCondition & " And (Absences.AbsenceID=" & CStr(oRequest("AbsenceID").Item) & ")"
	End If
	Call GetStartAndEndDatesFromURL("FilterStart", "FilterEnd", "OcurredDate", False, sCondition)
	sCondition = sCondition & aAbsenceComponent(S_QUERY_CONDITION_ABSENCE)
	If Len(sCondition ) > 0 Then
		If InStr(1, sCondition , "And ", vbBinaryCompare) = 0 Then sCondition  = "And " & sCondition
	End If

	sQuery = "Select EmployeesAbsencesLKP.*, Absences.JustificationID As WithJustification, AbsenceShortName, AbsenceName, JustificationShortName, JustificationName, Users.UserName, Users.UserLastName, RemoveUsers.UserName As RemoveUserName, RemoveUsers.UserLastName As RemoveUserLastName From EmployeesAbsencesLKP, Absences, Justifications, Users, Users As RemoveUsers, Employees" & sTables & " Where (Employees.EmployeeID = EmployeesAbsencesLKP.EmployeeID) And (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And (EmployeesAbsencesLKP.JustificationID=Justifications.JustificationID) And (EmployeesAbsencesLKP.AddUserID=Users.UserID) And (EmployeesAbsencesLKP.RemoveUserID=RemoveUsers.UserID)"
	sErrorDescription = "No se pudo obtener la información de los registros."
	If CInt(Request.Cookies("SIAP_SectionID")) <> 7 Then  ' Dif. de Desc.
		If CInt(Request.Cookies("SIAP_SubSectionID")) = 22 Then  ' Igual a Prestaciones e incidencias
			sCondition = sCondition & " And (EmployeesAbsencesLKP.AbsenceID=" &	aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ")"
		Else ' Igual a Inf. - Emp. - Inci
			sCondition = sCondition & " And (EmployeesAbsencesLKP.AbsenceID Not IN (201, 202))"
		End If
	Else ' Igual a Desc.
		If CInt(Request.Cookies("SIAP_SubSectionID")) = 721 Then  ' Igual a Prestaciones e incidencias
			sCondition = sCondition & " And (EmployeesAbsencesLKP.AbsenceID=" &	aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ")"
		Else
			sCondition = sCondition & " And (EmployeesAbsencesLKP.AbsenceID Not IN (201, 202))"
		End If
	End If
	If aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1 Then
		If CInt(Request.Cookies("SIAP_SubSectionID")) = 22 Then
			If aAbsenceComponent(N_ACTIVE_ABSENCE) = 1 Then
				sCondition = sCondition & " And (EmployeesAbsencesLKP.EmployeeID=0)"
			End If
		Else
			sCondition = sCondition & " And (EmployeesAbsencesLKP.EmployeeID=0)"
		End If
	Else
		sCondition = sCondition & " And (EmployeesAbsencesLKP.EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ")"
	End If
	If aAbsenceComponent(N_ACTIVE_ABSENCE) = 0 Then
		sCondition = sCondition & " And (EmployeesAbsencesLKP.Active<=" & aAbsenceComponent(N_ACTIVE_ABSENCE) & ")"
	Else
		sCondition = sCondition & " And (EmployeesAbsencesLKP.Active>=" & aAbsenceComponent(N_ACTIVE_ABSENCE) & ")"
	End If
	sQuery = sQuery & sCondition & " Order By OcurredDate Desc, RegistrationDate Desc"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & sQuery & """ />"

	GetAbsences = lErrorNumber
	Err.Clear
End Function

Function GetAbsencesDates(oRequest, oADODBConnection, aAbsenceComponent, lReasonID, sDates, sErrorDescription)
'************************************************************
'Purpose: To get the dates for all the absences for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAbsenceComponent, sDates, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAbsencesDates"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sCondition

	Select Case lReasonID
		Case EMPLOYEES_EXTRAHOURS
			sCondition = " And (AbsenceID=201)"
		Case EMPLOYEES_SUNDAYS
			sCondition = " And (AbsenceID=202)"
		Case Else
			sCondition = " And (AbsenceID NOT IN (201, 202))"
	End Select
	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct OcurredDate From EmployeesAbsencesLKP Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ")" & sCondition & " Order By OcurredDate", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		sDates = ""
		Do While Not oRecordset.EOF
			sDates = sDates & CStr(oRecordset.Fields("OcurredDate").Value) & ","
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
	End If

	GetAbsencesDates = lErrorNumber
	Err.Clear
End Function

Function GetAbsenceAppliesToID(oRequest, oADODBConnection, aAbsenceComponent, sAbsenceIDs, sErrorDescription)
'************************************************************
'Purpose: To get the absences requerids for insert te
'         absence for employee
'Inputs:  oRequest, oADODBConnection, aAbsenceComponent
'Outputs: sAbsenceIDs, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAbsenceAppliesToID"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información de la incidencia."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Absences Where (AbsenceID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "El tipo de incidencia especificada no se encuentra en el sistema."
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "AbsenceComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
		Else
			sAbsenceIDs = CStr(oRecordset.Fields("AppliesToID").Value)
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GetAbsenceAppliesToID = lErrorNumber
	Err.Clear
End Function

Function GetAbsenceIDsForPeriod(sAbsenceIDs, sErrorDescription)
'************************************************************
'Purpose: To get the absences requerids for insert te
'         absence for employee
'Inputs:  oRequest, oADODBConnection, aAbsenceComponent
'Outputs: sAbsenceIDs, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAbsenceIDsForPeriod"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información de la incidencia."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Absences Where (IsForPeriod=1)", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen incidencias por periodos en el sistema."
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "AbsenceComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
		Else
			Do While Not oRecordset.EOF
				sAbsenceIDs = sAbsenceIDs & CStr(oRecordset.Fields("AbsenceID").Value) & ","
				oRecordset.MoveNext
			Loop
			If (InStr(Right(sAbsenceIDs,1),",") > 0) Then
				sAbsenceIDs = Left(sAbsenceIDs, Len(sAbsenceIDs) -1)
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GetAbsenceIDsForPeriod = lErrorNumber
	Err.Clear
End Function

Function GetCrossingAbsenceType(oADODBConnection, aAbsenceComponent, sAbsenceIDs, sAbsenceCrossType, lAbsenceID, lStartDate, lEndDate, lVacationPeriod, lDays, sErrorDescription)
'************************************************************
'Purpose: To get the type of crossing absence for the
'         absence to insert
'Inputs:  oRequest, oADODBConnection, aAbsenceComponent
'Outputs: sAbsenceIDs, lStartDate, lEndDate, lVacationPeriod, lDays, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetCrossingAbsenceType"
	Dim oRecordset
	Dim lErrorNumber
	Dim sQuery

	If (InStr(1, sAbsenceIDs, "-1", vbBinaryCompare) > 0) And (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 34) Then
		sAbsenceIDs = "35, 37"
		sQuery = "Select * from EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_ID_EMPLOYEE) & ") And (AbsenceID IN (" & sAbsenceIDs & "))" & _
				 " And (((OcurredDate <= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (EndDate >= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & "))" & _
				 " And ((OcurredDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & ") And (EndDate >= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & ")" & _
				 " Or (OcurredDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & ") And (EndDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & ")))"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sAbsenceCrossType = "Inner"
				lAbsenceID = CInt(oRecordset.Fields("AbsenceID").Value)
				lStartDate = CLng(oRecordset.Fields("OcurredDate").Value)
				lEndDate = CLng(oRecordset.Fields("EndDate").Value)
				lVacationPeriod = CLng(oRecordset.Fields("VacationPeriod").Value)
				lDays = CInt(oRecordset.Fields("AbsenceHours").Value)
			Else
				sQuery = "Select * from EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_ID_EMPLOYEE) & ") And (AbsenceID IN (" & sAbsenceIDs & "))" & _
						 " And (((OcurredDate > " & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (EndDate > " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & "))" & _
						 " And ((OcurredDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & ") And (EndDate > " &  aAbsenceComponent(N_END_DATE_ABSENCE) & ")))"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						sAbsenceCrossType = "Left"
						lAbsenceID = CInt(oRecordset.Fields("AbsenceID").Value)
						lStartDate = CLng(oRecordset.Fields("OcurredDate").Value)
						lEndDate = CLng(oRecordset.Fields("EndDate").Value)
						lVacationPeriod = CLng(oRecordset.Fields("VacationPeriod").Value)
						lDays = CInt(oRecordset.Fields("AbsenceHours").Value)
					Else
						sQuery = "Select * from EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_ID_EMPLOYEE) & ") And (AbsenceID IN (" & sAbsenceIDs & "))" & _
								 " And (((OcurredDate < " & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (EndDate >= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & "))" & _
								 " And ((OcurredDate < " &  aAbsenceComponent(N_END_DATE_ABSENCE) & ") And (EndDate < " &  aAbsenceComponent(N_END_DATE_ABSENCE) & ")))"
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								sAbsenceCrossType = "Right"
								lAbsenceID = CInt(oRecordset.Fields("AbsenceID").Value)
								lStartDate = CLng(oRecordset.Fields("OcurredDate").Value)
								lEndDate = CLng(oRecordset.Fields("EndDate").Value)
								lVacationPeriod = CLng(oRecordset.Fields("VacationPeriod").Value)
								lDays = CInt(oRecordset.Fields("AbsenceHours").Value)
							End If
						Else
							sErrorDescription = "No se pudo obtener la información de la ausencia, para verificar que no se empalme con otra."
						End If
					End If
				Else
					sErrorDescription = "No se pudo obtener la información de la ausencia, para verificar que no se empalme con otra."
				End If
			End If
		Else
			sErrorDescription = "No se pudo obtener la información de la ausencia, para verificar que no se empalme con otra."
		End If
	Else
		If InStr(1, ",50,51,52,53,54,55,56,", "," & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ",", vbBinaryCompare) > 0 Then
			sAbsenceIDs = "50,51,52,53,54,55,56"
		Else
			sAbsenceIDs = "41,42,43,44,45,46,47,48,49,57,58"
		End If
		sQuery = "Select * From EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_ID_EMPLOYEE) & ")" & _
				 " And (AbsenceID IN (" & sAbsenceIDs & "))" & _
				 " And (OcurredDate>=" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ")" & _
				 " And (EndDate>=OcurredDate)"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sAbsenceCrossType = "Cross"
				lAbsenceID = CInt(oRecordset.Fields("AbsenceID").Value)
				lStartDate = CLng(oRecordset.Fields("OcurredDate").Value)
				lEndDate = CLng(oRecordset.Fields("EndDate").Value)
			Else
				sQuery = "Select * From EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_ID_EMPLOYEE) & ")" & _
						 " And (AbsenceID IN (" & sAbsenceIDs & "))" & _
						 " And (OcurredDate<" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ")" & _
						 " And (EndDate>=" &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ")"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						sAbsenceCrossType = "Inner"
						lAbsenceID = CInt(oRecordset.Fields("AbsenceID").Value)
						lStartDate = CLng(oRecordset.Fields("OcurredDate").Value)
						lEndDate = CLng(oRecordset.Fields("EndDate").Value)
					End If
				Else
					sErrorDescription = "No se pudo obtener la información de la ausencia, para verificar que no se empalme con otra."
				End If
			End If
		Else
			sErrorDescription = "No se pudo obtener la información de la ausencia, para verificar que no se empalme con otra."
		End If
	End If

	Set oRecordset = Nothing
	GetCrossingAbsenceType = lErrorNumber
	Err.Clear
End Function

Function GetReasonIDfromAbsence(iReasonID, sErrorDescription)
'************************************************************
'Purpose: To get the ReasonID from suspension
'         absence for employee
'Inputs:  oRequest, oADODBConnection, aAbsenceComponent
'Outputs: sAbsenceIDs, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetReasonIDfromAbsence"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener el motivo de la suspensión."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Reasons.* From Reasons, Absences Where (Reasons.ReasonShortName=Absences.AbsenceShortName) And (Absences.AbsenceID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existe el motivo de la suspensión."
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "AbsenceComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
		Else
			iReasonID = CInt(oRecordset.Fields("ReasonID").Value)
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GetReasonIDfromAbsence = lErrorNumber
	Err.Clear
End Function

Function GetShortNamesForApplyAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To get the dates for all the absences for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAbsenceComponent, sDates, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAbsencesDates"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct OcurredDate From EmployeesAbsencesLKP Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") Order By OcurredDate", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		sDates = ""
		Do While Not oRecordset.EOF
			sDates = sDates & CStr(oRecordset.Fields("OcurredDate").Value) & ","
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
	End If

	GetAbsencesDates = lErrorNumber
	Err.Clear
End Function

Function GetVacationDates(oRequest, oADODBConnection, aAbsenceComponent, lStartDate, lEndDate, sErrorDescription)
'************************************************************
'Purpose: To get the dates for all the absences for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAbsenceComponent, sDates, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetVacationDates"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	sErrorDescription = "No se pudo obtener la información de las vacaciones."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAbsencesLKP Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") and (AbsenceID IN () and " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") Order By OcurredDate", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		sDates = ""
		Do While Not oRecordset.EOF
			sDates = sDates & CStr(oRecordset.Fields("OcurredDate").Value) & ","
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
	End If

	GetVacationDates = lErrorNumber
	Err.Clear
End Function

Function JustifyAbsence(oRequest, oADODBConnection, iAbsenceId, iJustificationID, iActiveOriginal, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To justify an existing absence for the employee in
'         the database
'Inputs:  oRequest, oADODBConnection, iAbsenceId, iJustificationID, iActiveOriginal
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "JustifyAbsence"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	sErrorDescription = "No se pudo actualizar la información del registro."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set DocumentNumber='" & aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE) & "', JustificationID=" & iJustificationID & ", Removed=1, RemoveUserID=" & aAbsenceComponent(N_REMOVE_USER_ID_ABSENCE) & ", RemovedDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", AppliedRemoveDate=" & aAbsenceComponent(N_APPLIED_REMOVE_DATE_ABSENCE) & ", Active=" & iActiveOriginal & " Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & iAbsenceID & ") And (OcurredDate=" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

	JustifyAbsence = lErrorNumber
	Err.Clear
End Function

Function ModifyAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing absence for the employee in
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyAbsence"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador del concepto y/o la fecha para modificar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If lErrorNumber = 0 Then
			If aAbsenceComponent(B_IS_DUPLICATED_ABSENCE) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un registro para el " & DisplayDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), -1, -1, -1) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				If Not CheckAbsenceInformationConsistency(aAbsenceComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					sErrorDescription = "No se pudo modificar la información del registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set DocumentNumber='" & Replace(aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE), "'", "´") & "', AbsenceHours=" & aAbsenceComponent(N_HOURS_ABSENCE) & ", JustificationID=" & aAbsenceComponent(N_JUSTIFICATION_ID_ABSENCE) & ", AppliesForPunctuality=" & aAbsenceComponent(N_APPLIES_FOR_PUNCTUALITY_ABSENCE) & ", Removed=" & aAbsenceComponent(N_REMOVED_ABSENCE) & ", RemoveUserID=" & aAbsenceComponent(N_REMOVE_USER_ID_ABSENCE) & ", RemovedDate=" & aAbsenceComponent(N_REMOVED_DATE_ABSENCE) & ", AppliedRemoveDate=" & aAbsenceComponent(N_APPLIED_REMOVE_DATE_ABSENCE) & ", Active=" & aAbsenceComponent(N_ACTIVE_ABSENCE) & " Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ") And (OcurredDate=" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
		End If
	End If

	ModifyAbsence = lErrorNumber
	Err.Clear
End Function

Function CancelAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To remove an absence for the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CancelAbsence"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador del concepto y/o la fecha para eliminar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If IsEmpty(iAbsenceID) Then iAbsenceID = aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
		sErrorDescription = "No se pudo cancelar la incidencia del día " + aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set Removed=1, DocumentNumber='" & aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE) & "', RemoveUserID=" & aAbsenceComponent(N_REMOVE_USER_ID_ABSENCE) & ", RemovedDate=" & aAbsenceComponent(N_REMOVED_DATE_ABSENCE) & ", AppliedRemoveDate=" & aAbsenceComponent(N_APPLIED_REMOVE_DATE_ABSENCE) & ", Active=" & aAbsenceComponent(N_ACTIVE_ABSENCE) & " Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ") And (OcurredDate=" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	CancelAbsence = lErrorNumber
	Err.Clear
End Function

Function RemoveAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To remove an absence for the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveAbsence"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador del concepto y/o la fecha para eliminar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesAbsencesLKP Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ") And (OcurredDate=" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set Removed=1, RemoveUserID=" & aAbsenceComponent(N_REMOVE_USER_ID_ABSENCE) & ", RemovedDate=" & aAbsenceComponent(N_REMOVED_DATE_ABSENCE) & ", AppliedRemoveDate=0 Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ") And (OcurredDate=" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (AppliedDate>0)", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	RemoveAbsence = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfAbsence(aAbsenceComponent, bIsForPeriod, sErrorDescription)
'************************************************************
'Purpose: To check if a specific absence for the employee
'         exists in the database
'Inputs:  aAbsenceComponent
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfAbsence"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."

		sQuery = "Select * From EmployeesAbsencesLKP Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ")"
		If bIsForPeriod Then
			sQuery = sQuery & _
					 " And (((OcurredDate >= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (OcurredDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & "))" & _
					 " Or ((EndDate >= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (EndDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & "))" & _
					 " Or ((EndDate >= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (OcurredDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & ")))"
		Else
			sQuery = sQuery & " And (OcurredDate=" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ")"
		End If

		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aAbsenceComponent(B_IS_DUPLICATED_ABSENCE) = True
			End If
		Else
			If aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = 5 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAbsencesLKP Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=3) And (OcurredDate=" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete from EmployeesAbsencesLKP Where (EmployeeId = " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID = 3) And (OcurredDate = " & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					End If
				End If
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfAbsence = lErrorNumber
	Err.Clear
End Function

Function CheckAbsenceInformationConsistency(aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aAbsenceComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckAbsenceInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del empleado no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aAbsenceComponent(N_ABSENCE_ID_ABSENCE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del concepto no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aAbsenceComponent(N_OCURRED_DATE_ABSENCE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La fecha de ocurrencia del registro no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aAbsenceComponent(N_END_DATE_ABSENCE)) Then aAbsenceComponent(N_END_DATE_ABSENCE) = aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
	If Not IsNumeric(aAbsenceComponent(N_REGISTRATION_DATE_ABSENCE)) Then aAbsenceComponent(N_REGISTRATION_DATE_ABSENCE) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	If Len(aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE)) = 0 Then aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE) = " "
	If Not IsNumeric(aAbsenceComponent(N_HOURS_ABSENCE)) Then aAbsenceComponent(N_HOURS_ABSENCE) = 0
	If Not IsNumeric(aAbsenceComponent(N_JUSTIFICATION_ID_ABSENCE)) Then aAbsenceComponent(N_JUSTIFICATION_ID_ABSENCE) = -1
	If Not IsNumeric(aAbsenceComponent(N_APPLIES_FOR_PUNCTUALITY_ABSENCE)) Then aAbsenceComponent(N_APPLIES_FOR_PUNCTUALITY_ABSENCE) = 1
	If Not IsNumeric(aAbsenceComponent(N_ADD_USER_ID_ABSENCE)) Then aAbsenceComponent(N_ADD_USER_ID_ABSENCE) = aLoginComponent(N_USER_ID_LOGIN)

	If Not IsNumeric(aAbsenceComponent(N_APPLIED_DATE_ABSENCE)) Then aAbsenceComponent(N_APPLIED_DATE_ABSENCE) = 0
	If Not IsNumeric(aAbsenceComponent(N_REMOVED_ABSENCE)) Then aAbsenceComponent(N_REMOVED_ABSENCE) = 0
	If Not IsNumeric(aAbsenceComponent(N_REMOVE_USER_ID_ABSENCE)) Then aAbsenceComponent(N_REMOVE_USER_ID_ABSENCE) = -1
	If Not IsNumeric(aAbsenceComponent(N_REMOVED_DATE_ABSENCE)) Then aAbsenceComponent(N_REMOVED_DATE_ABSENCE) = 0
	If Not IsNumeric(aAbsenceComponent(N_APPLIED_REMOVE_DATE_ABSENCE)) Then aAbsenceComponent(N_APPLIED_REMOVE_DATE_ABSENCE) = 0

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del registro contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckAbsenceInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DayIsVacation(aAbsenceComponent, lDate, sErrorDescription)
'************************************************************
'Purpose: To add a new absence for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DayIsVacation"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bIsForPeriod = True
	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (lDate = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador de la incidencia y/o la fecha para verificar si existen vacaciones."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sQuery = "Select * from EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_ID_EMPLOYEE) & ") And (AbsenceID IN (35,37,38))" & _
				 " And (((OcurredDate >= " & lDate & ") And (OcurredDate <= " & lDate & "))" & _
				 " Or ((EndDate >= " & lDate & ") And (EndDate <= " & lDate & "))" & _
				 " Or ((EndDate >= " & lDate & ") And (OcurredDate <= " & lDate & ")))"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				DayIsVacation = True
			Else
				DayIsVacation = False
			End If
		Else
			DayIsVacation = False
		End If
		oRecordset.Close
	End If
	Set oRecordset = Nothing
	Err.Clear
End Function

Function DisplayAbsenceForm(oRequest, oADODBConnection, sAction, lReasonID, sExtraURL, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an absence for the
'         employee from the database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, sAbsenceIDs, sExtraURL, aAbsenceComponent
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAbsenceForm"
	Dim sNames
	Dim aRelatedAbsences
	Dim iIndex
	Dim oRecordset
	Dim lErrorNumber
	Dim sAbsenceIDs
	Dim sCaseOptions

	If (Len(oRequest("AbsenceReview").Item) = 0) And (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) <> -1) And (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) <> -1) And (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) <> 0) Then
		lErrorNumber = GetAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function UpdateText(elementID, newText){" & vbNewLine
				Response.Write "var objetoSPAN = document.getElementById(elementID);" & vbNewLine
					Response.Write "objetoSPAN.innerHTML = newText;" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "}" & vbNewLine

			Response.Write "function CheckAbsenceFields(oForm) {" & vbNewLine
				If Len(oRequest("Delete").Item) = 0 Then
					Response.Write "if (oForm) {" & vbNewLine
						If Not B_ISSSTE Then
							Response.Write "if (oForm.DocumentNumber.value.length == 0) {" & vbNewLine
								Response.Write "alert('Favor de introducir el número de folio.');" & vbNewLine
								Response.Write "oForm.DocumentNumber.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Else
							If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
								'If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
									Response.Write "if ((oForm.EmployeeID.value.length == 0) || (oForm.EmployeeID.value == '-1')) {" & vbNewLine
										Response.Write "alert('Favor de especificar el número de empleado.');" & vbNewLine
										Response.Write "oForm.EmployeeNumber.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
								'End If
							Else
								If Len(oRequest("AbsenceChange").Item) > 0 Then
									Response.Write "if (parseInt(oForm.AppliedRemoveDate.value)==-1) {" & vbNewLine
										Response.Write "alert('No existen nóminas abiertas para el registro de movimientos.');" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (oForm.AbsenceID.value == '-1') {" & vbNewLine
										Response.Write "alert('Favor de seleccionar la justificación de la incidencia.');" & vbNewLine
										Response.Write "oForm.JustificationID.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
								Else
									Response.Write "if (parseInt(oForm.AbsenceID.value)==-1) {" & vbNewLine
										Response.Write "alert('Seleccione la clave para el registro de incidencias');" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if (parseInt(oForm.AppliedDate.value)==-1) {" & vbNewLine
										Response.Write "alert('No existen nóminas abiertas para el registro de incidencias.');" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if ((oForm.AbsenceIDCmb.value == 52 || oForm.AbsenceIDCmb.value == 53) && (oForm.OcurredDates.length != 2)) {" & vbNewLine
										Response.Write "alert('Favor de indicar una fecha de inicio y una fecha de fin para la incidencia que se esta registrando.');" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if ((oForm.AbsenceIDCmb.value == 35 || oForm.AbsenceIDCmb.value == 37 || oForm.AbsenceIDCmb.value == 38) && (oForm.PeriodVacationID.value == 0)) {" & vbNewLine
										Response.Write "alert('Favor de indicar el periodo de las vacaciones que se estan registrando.');" & vbNewLine
										Response.Write "oForm.PeriodVacationID.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if ((oForm.AbsenceIDCmb.value == 39 || oForm.AbsenceIDCmb.value == 40) && (oForm.PeriodVacationID.value == 0)) {" & vbNewLine
										Response.Write "alert('Favor de indicar el periodo del estimulo que se estan registrando.');" & vbNewLine
										Response.Write "oForm.PeriodVacationID.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if ((oForm.AbsenceIDCmb.value == 8 || oForm.AbsenceIDCmb.value == 9 || oForm.AbsenceIDCmb.value == 15) && (oForm.ForJustificationID.value == -1)) {" & vbNewLine
										Response.Write "alert('Favor de indicar la clave de la incidencia que desea justificar.');" & vbNewLine
										Response.Write "oForm.ForJustificationID.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if ((oForm.AbsenceIDCmb.value == 52 || oForm.AbsenceIDCmb.value == 32 ) && (oForm.CheckGender.value == '')) {" & vbNewLine
										Response.Write "if (oForm.AbsenceIDCmb.value == 52) {" & vbNewLine
											Response.Write "alert('Favor de validar el sexo del empleado para registrar la tolerancia de lactancia.');" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "if (oForm.AbsenceIDCmb.value == 32) {" & vbNewLine
											Response.Write "alert('Favor de validar el sexo del empleado para registrar la incapacidad por gravidez.');" & vbNewLine
										Response.Write "}" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if ((oForm.AbsenceIDCmb.value == 32) && (oForm.CheckGender.value == '1')) {" & vbNewLine
										Response.Write "alert('No puede registrar la incapacidad por gravidez a un empleado de sexo másculino.');" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "if ((oForm.AbsenceIDCmb.value == 52) && (oForm.CheckGender.value == '1')) {" & vbNewLine
										Response.Write "alert('Esta registrando la tolerancia de lactancia a un empleado de sexo másculino.');" & vbNewLine
									Response.Write "}" & vbNewLine
									Response.Write "SelectAllItemsFromList(oForm.OcurredDates);" & vbNewLine
								End If
							End If
						End If
					Response.Write "}" & vbNewLine
				End If
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckAbsenceFields" & vbNewLine
		If aEmployeeComponent(N_ID_EMPLOYEE) <> -1 Then
			Response.Write "function ShowHideAbsencesFields(sValue) {" & vbNewLine
				Response.Write "var oForm = document.AbsencesFrm" & vbNewLine
				If Not B_ISSSTE Then
					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "if (sValue == 0) {" & vbNewLine
							Response.Write "HideDisplay(document.all['AbsenceHoursDiv']);" & vbNewLine
						Response.Write "} else {" & vbNewLine
							Response.Write "ShowDisplay(document.all['AbsenceHoursDiv']);" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
				End If
			Response.Write "} // End of ShowHideAbsencesFields" & vbNewLine
			Response.Write "function ShowMoveAbsencesWng(sValue) {" & vbNewLine
				Response.Write "var oForm = document.AbsencesFrm" & vbNewLine
					Response.Write "switch (sValue) {" & vbNewLine
						Response.Write "case '34':" & vbNewLine
							Response.Write "ShowDisplay(document.all['ApplyAbsenceWngDiv']);" & vbNewLine
							Response.Write "return false;" & vbNewLine
							Response.Write "break;" & vbNewLine
						Response.Write "default:" & vbNewLine
							Response.Write "HideDisplay(document.all['ApplyAbsenceWngDiv']);" & vbNewLine
							Response.Write "return false;" & vbNewLine
							Response.Write "break;" & vbNewLine
					Response.Write "}" & vbNewLine
			Response.Write "} // End of ShowMoveAbsencesWng" & vbNewLine
			Response.Write "function ClearOcurredDates() {" & vbNewLine
				'Response.Write "alert('Funcion ClearOcurredDates ' + document.AbsencesFrm.AbsenceIDCmb.value);" & vbNewLine
				Response.Write "HideDisplay(document.all['EmployeeGenderDiv'])" & vbNewLine
				Response.Write "HideDisplay(document.all['AbsencesForPeriodDiv'])" & vbNewLine
				Response.Write "HideDisplay(document.all['EmployeeVacationDiv'])" & vbNewLine
				Response.Write "if (IsAbsencesForPeriod(document.AbsencesFrm.AbsenceIDCmb.value)){" & vbNewLine
					Response.Write "if (lJourneyID!=21 && lJourneyID!=22 && lJourneyID!=23) {" & vbNewLine
						Response.Write "ShowDisplay(document.all['AbsencesForPeriodDiv'])" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "else {" & vbNewLine
						'Response.Write "alert('lJourneyID ' + lJourneyID);" & vbNewLine
						Response.Write "ShowDisplay(document.all['AbsencesForPeriod1Div'])" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "RemoveAllItemsFromList(null, document.AbsencesFrm.OcurredDates);" & vbNewLine
					Response.Write "if (document.AbsencesFrm.AbsenceIDCmb.value == 35 || document.AbsencesFrm.AbsenceIDCmb.value == 37 || document.AbsencesFrm.AbsenceIDCmb.value == 38) {"  & vbNewLine
						Response.Write "if (document.AbsencesFrm.AbsenceIDCmb.value == 35) {"  & vbNewLine
							Response.Write "RemoveAllItemsFromList(null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
							Response.Write "AddItemToList(0, 0, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
							Response.Write "AddItemToList(1, 1, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
							Response.Write "AddItemToList(2, 2, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (document.AbsencesFrm.AbsenceIDCmb.value == 37) {"  & vbNewLine
							Response.Write "RemoveAllItemsFromList(null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
							Response.Write "AddItemToList(0, 0, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
							Response.Write "AddItemToList(1, 1, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
							Response.Write "AddItemToList(2, 2, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
							Response.Write "AddItemToList(3, 3, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
							Response.Write "AddItemToList(4, 4, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (document.AbsencesFrm.AbsenceIDCmb.value == 38) {"  & vbNewLine
							Response.Write "RemoveAllItemsFromList(null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
							Response.Write "AddItemToList(1, 1, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "ShowDisplay(document.all['EmployeeVacationDiv'])" & vbNewLine
						Response.Write "UpdateText('Period', 'Periodo:');" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (document.AbsencesFrm.AbsenceIDCmb.value == 32) {"  & vbNewLine
						Response.Write "ShowDisplay(document.all['EmployeeGenderDiv'])" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "else if (document.AbsencesFrm.AbsenceIDCmb.value == 52) {"  & vbNewLine
						Response.Write "ShowDisplay(document.all['EmployeeGenderDiv'])" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "else {" & vbNewLine
					Response.Write "if (document.AbsencesFrm.AbsenceIDCmb.value == 52) {"  & vbNewLine
						Response.Write "ShowDisplay(document.all['EmployeeGenderDiv'])" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (document.AbsencesFrm.AbsenceIDCmb.value == 39 || document.AbsencesFrm.AbsenceIDCmb.value == 40) {"  & vbNewLine
						Response.Write "RemoveAllItemsFromList(null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
						Response.Write "AddItemToList('0', 0, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
						Response.Write "AddItemToList('1', 1, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
						Response.Write "AddItemToList('2', 2, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
						Response.Write "AddItemToList('3', 3, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
						Response.Write "AddItemToList('4', 4, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
						Response.Write "AddItemToList('5', 5, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
						Response.Write "AddItemToList('6', 6, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
						Response.Write "AddItemToList('7', 7, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
						Response.Write "AddItemToList('8', 8, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
						Response.Write "AddItemToList('9', 9, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
						Response.Write "AddItemToList('10', 10, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
						Response.Write "AddItemToList('11', 11, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
						Response.Write "AddItemToList('12', 12, null, document.AbsencesFrm.PeriodVacationID);" & vbNewLine
						Response.Write "ShowDisplay(document.all['EmployeeVacationDiv'])" & vbNewLine
						Response.Write "UpdateText('Period', 'Mes:');" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "else {" & vbNewLine
					Response.Write "HideDisplay(document.all['AbsencesForPeriodDiv'])" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (document.AbsencesFrm.AbsenceIDCmb.value == 8 || document.AbsencesFrm.AbsenceIDCmb.value == 9 || document.AbsencesFrm.AbsenceIDCmb.value == 15) {"  & vbNewLine
						Response.Write "if (document.AbsencesFrm.AbsenceIDCmb.value == 8) {"  & vbNewLine
							Response.Write "RemoveAllItemsFromList(null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList('Ninguna', -1, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList('0818', 18, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList('0819', 19, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList('0820', 20, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList('0822', 22, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList('0825', 23, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (document.AbsencesFrm.AbsenceIDCmb.value == 9) {"  & vbNewLine
							Response.Write "RemoveAllItemsFromList(null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList('Ninguna', -1, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList(0801, 1, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList(0802, 2, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList(0803, 3, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList(0805, 5, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList(0821, 21, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList(0823, 91, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (document.AbsencesFrm.AbsenceIDCmb.value == 15) {"  & vbNewLine
							Response.Write "RemoveAllItemsFromList(null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList('Ninguna', -1, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList(0810, 10, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList(0811, 11, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList(0816, 16, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
							Response.Write "AddItemToList(0827, 25, null, document.AbsencesFrm.ForJustificationID);" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "ShowDisplay(document.all['JustificationDiv'])" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "} // End of IsAbsencesForPeriod" & vbNewLine
			Response.Write "} // End of ClearOcurredDates" & vbNewLine
			Response.Write "function IsAbsencesForPeriod(sValue) {" & vbNewLine
				lErrorNumber = GetAbsenceIDsForPeriod(sAbsenceIDs, sErrorDescription)
				If (lErrorNumber = L_ERR_NO_RECORDS) Then
					Response.Write "return false;" & vbNewLine
					sErrorDescription = ""
					lErrorNumber = 0
				Else
					sCaseOptions = Split(sAbsenceIDs, "," , -1, vbBinaryCompare)
					Response.Write "switch (sValue) {" & vbNewLine
						For iIndex = 0 To UBound(sCaseOptions)
							Response.Write "case '" & CInt(sCaseOptions(iIndex)) & "':" & vbNewLine
						Next
							Response.Write "return true;" & vbNewLine
						Response.Write "default:" & vbNewLine
							'Response.Write "alert('default: ' + sValue);" & vbNewLine
							Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				End If
			Response.Write "} // End of IsAbsencesForPeriod" & vbNewLine
		End If
		Response.Write "//--></SCRIPT>" & vbNewLine
			sNames = oRequest("Action").Item
			If Len(sNames) = 0 Then sNames = "Absences"
				If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
					Response.Write "<FORM NAME=""AbsencesFrm"" ID=""AbsencesFrm"" ACTION=""" & sAction & """ METHOD=""GET"" onSubmit=""return CheckAbsenceFields(this)"">"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Absences"" />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE=""" & lReasonID & """ />"
					Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
							Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
					Response.Write "</TABLE>"
					Response.Write "<BR /><BR />"
					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then
						Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""AbscenceMovement"" ID=""AbscenceMovementBtn"" VALUE=""Buscar empleado"" CLASS=""Buttons"" />"
					End If
					Response.Write "</FORM>"
				Else
					lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						If lErrorNumber = 0 Then
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
								Response.Write "var lJourneyID=" & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & ";" & vbNewLine
							Response.Write "//--></SCRIPT>" & vbNewLine
							If CInt(Request.Cookies("SIAP_SectionID")) <> 1 Then
								Response.Write "<FORM NAME=""AnotherAbsencesFrm"" ID=""AnotherAbsencesFrm"" ACTION=""" & sAction & """ METHOD=""GET"">"
									Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
										Response.Write "<TR>"
											Response.Write "<TD VALIGN=""TOP"" WIDTH=""500"">&nbsp;</TD>"
											Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><IMG SRC=""Images/MnLeftArrows.gif"" WIDTH=""32"" HEIGHT=""32"" ALT=""Incidencias"" BORDER=""0"" /><BR /></TD>"
											Response.Write "<TD VALIGN=""TOP"" WIDTH=""290""><FONT FACE=""Arial"" SIZE=""2""><B>Otro empleado</B><BR /></FONT>"
											Response.Write "<DIV CLASS=""MenuOverflow""><FONT FACE=""Arial"" SIZE=""2"">Registre las incidencias a un empleado diferente.</FONT></DIV></TD>"
										Response.Write "</TR>"
										Response.Write "<TR>"
											Response.Write "<TD VALIGN=""TOP"" WIDTH=""500"">&nbsp;</TD>"
											Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;</FONT></TD>"
											Response.Write "<TD VALIGN=""TOP"" WIDTH=""290""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;Número del empleado:&nbsp;</FONT><INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" /></TD>"
											Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Absences"" />"
											Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE=""" & lReasonID & """ />"
										Response.Write "</TR>"
										Response.Write "<TR>"
											Response.Write "<TD VALIGN=""TOP"" WIDTH=""500"">&nbsp;</TD>"
											Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;</FONT></TD>"
											Response.Write "<TD VALIGN=""TOP"" WIDTH=""290""><INPUT TYPE=""SUBMIT"" NAME=""AbscenceMovement"" ID=""AbscenceMovementBtn"" VALUE=""Buscar empleado"" CLASS=""Buttons"" /></TD>"
										Response.Write "</TR>"
									Response.Write "</TABLE>"
								Response.Write "</FORM>"
							End If
							Response.Write "<DIV NAME=""AbsencesForPeriodDiv"" ID=""AbsencesForPeriodDiv"" STYLE=""display: none"">"
								Response.Write "<FORM NAME=""UploadValidateInfoFrm"" ID=""UploadValidateInfoFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
									Call DisplayErrorMessage("Advertencia", "Para este tipo de incidencia puede capturar periodos. </BR> La primer fecha seleccionada es la fecha de inicio y la segunda fecha seleccionada es la fecha de fin. En caso de elegir <B>sólo una fecha de registro</B>, la incidencia será de un solo día, con fecha de inicio y fin igual al día seleccionado.")
								Response.Write "</FORM>"
							Response.Write "</DIV>"
							Response.Write "<DIV NAME=""AbsencesForPeriod1Div"" ID=""AbsencesForPeriod1Div"" STYLE=""display: none"">"
								Response.Write "<FORM NAME=""UploadValidateInfoFrm"" ID=""UploadValidateInfoFrm"" METHOD=""POST"" onSubmit=""return bReady"">"
									Call DisplayErrorMessage("Advertencia", "Para registrar vacaciones a empleados con tipo de jornada 2 debe de registrar los días de forma individual.")
								Response.Write "</FORM>"
							Response.Write "</DIV>"
							Response.Write "<FORM NAME=""AbsencesFrm"" ID=""AbsencesFrm"" ACTION=""" & sAction & """ METHOD=""GET"" onSubmit=""return CheckAbsenceFields(this)"">"
								' Aquí se ubicaron los botones cuando los solicitaron en la parte superior
								'Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""150"" HEIGHT=""1"" />"
								'If aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1 Then
								'	Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=Absences'"" />"
								'Else
								'	Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=Absences&EmployeeId=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & "'"" />"
								'End If
								Response.Write "<BR /><BR />"
							If Len(oRequest("AbsenceChange").Item) > 0 Then
								lErrorNumber = GetAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AbsenceChange"" ID=""AbsenceChangeHdn"" VALUE=""" & oRequest("AbsenceChange").Item & """ />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RemoveUserID"" ID=""RemoveUserIDHdn"" VALUE=""" & aLoginComponent(N_USER_ID_LOGIN) & """ />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RemovedDate"" ID=""RemovedDateHdn"" VALUE=""" & Left(GetSerialNumberForDate(""), Len("00000000")) & """ />"
								If Len(oRequest("CancelAbsence").Item) > 0 Then
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CancelAbsence"" ID=""CancelAbsenceHdn"" VALUE=""1"" />"
								ElseIf Len(oRequest("CancelJustification").Item) > 0 Then
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CancelJustification"" ID=""CancelJustificationHdn"" VALUE=""1"" />"
								Else
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Justification"" ID=""JustificationHdn"" VALUE=""1"" />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Active"" ID=""ActiveHdn"" VALUE=""" & oRequest("Active").Item & """ />"
									aAbsenceComponent(N_JUSTIFICATION_ID_ABSENCE) = CLng(oRequest("JustificationID").Item)
								End If
								aAbsenceComponent(N_ACTIVE_ABSENCE) = 0
							End If
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Absences"" />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(N_ID_EMPLOYEE) & """ />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JourneyID"" ID=""JourneyIDHdn"" VALUE=""" & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & """ />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""4"" />"
							If (Len(oRequest("AbsenceChange").Item) = 0) And ((aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Or (Len(oRequest("AbsenceReview").Item) > 0)) Then
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OcurredDate"" ID=""OcurredDateHdn"" VALUE="""
									Response.Write aAbsenceComponent(N_OCURRED_DATE_ABSENCE)
								Response.Write """ />"
							End If
							If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) > -1) And (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) > -1) And (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) > 0) Then
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RegistrationDate"" ID=""RegistrationDateHdn"" VALUE=""" & aAbsenceComponent(N_REGISTRATION_DATE_ABSENCE) & """ />"
							End If
							Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Información general</B></FONT>"
							Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
							Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
								Response.Write "<TR>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
									If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
										Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""SpecialTextFields"" onFocus=""document.EmployeeFrm.EmployeeName.focus()"" /></TD>"
									Else
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "</FONT><INPUT TYPE=""HIDDEN"" NAME=""EmployeeNumber"" ID=""EmployeeNumberHdn"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ /></TD>"
									End If
								If (Len(oRequest("Success").Item) > 0) Then
									Response.Write "&nbsp;&nbsp;&nbsp;<TD VALIGN=""TOP"" ALIGN=""LEFT"" WIDTH=""60%"" ROWSPAN=""9"">"
									If CInt(oRequest("Success").Item) = 1 Then
										Call DisplayErrorMessage("Confirmación", "La operación con la incidencia " & CStr(oRequest("AbsenceShortName").Item) & " fué ejecutada exitosamente.")
									Else
										Call DisplayErrorMessage("Error al realizar la operación con la incidencia " & CStr(oRequest("AbsenceShortName").Item), CStr(oRequest("ErrorDescription").Item))
									End If
									Response.Write "</TD>"
								End If
								Response.Write "</TR>"
								Response.Write "<TR>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre(s):&nbsp;</FONT></TD>"
									Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeName"" ID=""EmployeeNameTxt"" VALUE=""" & aEmployeeComponent(S_NAME_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
								Response.Write "</TR>"
								Response.Write "<TR>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno:&nbsp;</FONT></TD>"
									Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName"" ID=""EmployeeLastNameTxt"" VALUE=""" & aEmployeeComponent(S_LAST_NAME_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
								Response.Write "</TR>"
								Response.Write "<TR>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido materno:&nbsp;</FONT></TD>"
									Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName2"" ID=""EmployeeLastName2Txt"" VALUE=""" & aEmployeeComponent(S_LAST_NAME2_EMPLOYEE) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
								Response.Write "</TR>"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate"" ID=""StartDateHdn"" VALUE=""" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & """ />"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de tabulador:&nbsp;</FONT></TD>"
								Call GetNameFromTable(oADODBConnection, "EmployeeTypes", aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE), "", "", sNames, "")
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de empleado:&nbsp;</FONT></TD>"
								Call GetNameFromTable(oADODBConnection, "PositionTypes", aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE), "", "", sNames, "")
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
								Call GetNameFromTable(oADODBConnection, "Positions", aEmployeeComponent(N_POSITION_ID_EMPLOYEE), "", "", sNames, "")
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Turno:&nbsp;</FONT></TD>"
								Call GetNameFromTable(oADODBConnection, "Journeys", aJobComponent(N_JOURNEY_ID_JOB), "", "", sNames, sErrorDescription)
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
							Response.Write "</TR>"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Shifts.ShiftID, JourneyTypeID From Shifts, (Select ShiftID From Employees Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") As Empleado Where Shifts.ShiftID=Empleado.ShiftID", "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horarios:&nbsp;</FONT></TD>"
								Call GetNameFromTable(oADODBConnection, "Shifts", aEmployeeComponent(N_SHIFT_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de Jornada:&nbsp;</FONT></TD>"
								Response.Write "<TD>"
									Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & oRecordset.Fields("JourneyTypeID").Value & "</FONT>"
								Response.Write "</TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estatus:&nbsp;</FONT></TD>"
								Call GetNameFromTable(oADODBConnection, "StatusEmployees", aEmployeeComponent(N_STATUS_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "</TABLE>"
							Response.Write "<DIV NAME=""EmployeeGenderDiv"" ID=""EmployeeGenderDiv"" STYLE=""display: none"">"
								Response.Write "<BR />"
								Response.Write " <B>Valide al empleado para esta incidencia:</B> <A HREF=""javascript: SearchRecord(document.AbsencesFrm.EmployeeID.value, 'EmployeeGender', 'SearchEmployeeGenderIFrame', 'AbsencesFrm.CheckGender')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>"
									Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeGenderIFrame"" FRAMEBORDER=""0"" WIDTH=""400"" HEIGHT=""16""></IFRAME>"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CheckGender"" ID=""CheckGenderHdn"" VALUE="""" />"
								Response.Write "<BR />"
							Response.Write "</DIV>"
							Response.Write "<DIV NAME=""EmployeeVacationDiv"" ID=""EmployeeVacationDiv"" STYLE=""display: none"">"
								Response.Write "<BR />"
									Response.Write "<FONT FACE=""Arial"" SIZE=""2""><SPAN ID=""Period"">Periodo:</SPAN></FONT>"
									Response.Write "&nbsp;&nbsp;<SELECT NAME=""PeriodVacationID"" ID=""PeriodVacationIDCmb"" SIZE=""1"" CLASS=""Lists"">"
										Response.Write "<OPTION VALUE=""0"">0</OPTION>"
										Response.Write "<OPTION VALUE=""1"">1</OPTION>"
										Response.Write "<OPTION VALUE=""2"">2</OPTION>"
									Response.Write "</SELECT>&nbsp;&nbsp;"
									Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Año:</FONT>"
									Response.Write "&nbsp;&nbsp;<SELECT NAME=""YearID"" ID=""YearIDCmb"" SIZE=""1"" CLASS=""Lists"">"
										For iIndex = (Year(Date()) - 2) To Year(Date()) + 2
											Response.Write "<OPTION VALUE=""" & iIndex & """>" & iIndex & "</OPTION>"
										Next
									Response.Write "</SELECT><BR />"
								Response.Write "<BR />"
							Response.Write "</DIV>"
							If Len(oRequest("AbsenceChange").Item) = 0 Then
								Response.Write "<DIV NAME=""JustificationDiv"" ID=""JustificationDiv"" STYLE=""display: none"">"
									Response.Write "<BR />"
										Response.Write "<FONT COLOR=""RED"" FACE=""Arial"" SIZE=""2"">Incidencia a justificar:</FONT>"
										Response.Write "&nbsp;&nbsp;<SELECT NAME=""ForJustificationID"" ID=""ForJustificationIDCmb"" SIZE=""1"" CLASS=""Lists"">"
											Response.Write "<OPTION VALUE=""-1"">Ninguna</OPTION>"
										Response.Write "</SELECT>&nbsp;&nbsp;"
									Response.Write "<BR />"
								Response.Write "</DIV>"
							End If
							Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
							Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
							If Len(oRequest("AbsenceChange").Item) > 0 Then
								Response.Write "<TR><TD COLSPAN=""2""><BR /></TD></TR>"
								Response.Write "<TR><TD COLSPAN=""2"">"
								If Len(oRequest("CancelAbsence").Item) > 0 Then
									Call DisplayInstructionsMessage("Cancelar Incidencia", "Seleccione la quincena de aplicación para la cancelación y capture el número de oficio, posteriormente de clic en el botón Modificar.")
								Else
									If CInt(oRequest("JustificationID").Item) <> -1 Then
										Call DisplayInstructionsMessage("Justificar Incidencia", "Seleccione la quincena de aplicación para la justificación y capture el número de oficio, posteriormente de clic en el botón Modificar.")
									End If
								End If
							End If
							Response.Write "</TABLE>"
							Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
							If Len(oRequest("AbsenceChange").Item) > 0 Then
								Response.Write "<TR><TD><BR /></TD></TR>"
								Response.Write "<TR NAME=""AppliedRemoveDateDiv"" ID=""AppliedRemoveDateDiv"">"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Quincena de aplicación de la justificación:&nbsp;</NOBR></FONT></TD>"
									Response.Write "<TD><SELECT NAME=""AppliedRemoveDate"" ID=""AppliedRemoveDate"" SIZE=""1"" CLASS=""Lists"">"
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "((IsClosed<>1) And (IsActive_2=1) And (PayrollTypeID=1)) Or (PayrollID=" & aAbsenceComponent(N_APPLIED_DATE_ABSENCE) & ")", "PayrollID Desc", aAbsenceComponent(N_APPLIED_DATE_ABSENCE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
									Response.Write "</SELECT>&nbsp;"
									Response.Write "</TD>"
								Response.Write "</TR>"
							Else
								Response.Write "<TR NAME=""AppliedDateDiv"" ID=""AppliedDateDiv"">"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Quincena de aplicación:&nbsp;</NOBR></FONT></TD>"
									Response.Write "<TD><SELECT NAME=""AppliedDate"" ID=""AppliedDate"" SIZE=""1"" CLASS=""Lists"">"
										If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
											Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_2=1) And (PayrollTypeID=1)", "PayrollID Desc", aAbsenceComponent(N_APPLIED_DATE_ABSENCE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
										Else
											Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID=1)", "PayrollID Desc", aAbsenceComponent(N_APPLIED_DATE_ABSENCE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
										End If
									Response.Write "</SELECT>&nbsp;"
									Response.Write "</TD>"
								Response.Write "</TR>"
							End If
							If Not B_ISSSTE Then
								Response.Write "<TR>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha:&nbsp;</FONT></TD>"
									Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""AbsenceDate"" ID=""AbsenceDateTxt"" SIZE=""50"" VALUE="""" CLASS=""SpecialTextFields"" onFocus=""document.AbsencesFrm.AbsenceID.focus()"" /></TD>"
								Response.Write "</TR>"
							Else
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AbsenceDate"" ID=""AbsenceDateHdn"" VALUE="""" />"
							End If
								Response.Write "<TR>"
									If Len(oRequest("AbsenceChange").Item) > 0 Then
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Incidencia cancelar o justificar:&nbsp;</FONT></TD>"
										Response.Write "<TD>"
											Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ForJustificationID"" ID=""ConceptIDHdn"" VALUE=""" & aAbsenceComponent(N_FOR_JUSTIFICATION_ID_ABSENCE) & """ />"
											Call GetNameFromTable(oADODBConnection, "Absences", aAbsenceComponent(N_FOR_JUSTIFICATION_ID_ABSENCE), "", "", sNames, sErrorDescription)
											Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT>"
										Response.Write "</TD>"
									Else
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Incidencia a capturar:&nbsp;</FONT></TD>"
										Response.Write "<TD><SELECT NAME=""AbsenceID"" ID=""AbsenceIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ClearOcurredDates()"">"
											Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Absences", "AbsenceID", "AbsenceShortName, AbsenceName", "(AbsenceID<100) And (Active=1)", "AbsenceShortName", "", "Ninguno;;;-1", sErrorDescription)
										Response.Write "</SELECT></TD>"
									End If
								Response.Write "</TR>"
								If Len(oRequest("AbsenceChange").Item) > 0 Then
									Response.Write "<TR>"
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), -1, -1, -1) & "</FONT></TD>"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OcurredDate"" ID=""OcurredDateHdn"" VALUE=""" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & """ />"
									Response.Write "</TR>"
								End If
								If Len(oRequest("AbsenceChange").Item) > 0 Then
									If B_ISSSTE Then
										Response.Write "<TR>"
											Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>No. de oficio:&nbsp;</NOBR></FONT></TD>"
											Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""DocumentNumber"" ID=""DocumentNumberTxt"" SIZE=""30"" MAXLENGTH=""50"" VALUE=""" & aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE) & """ CLASS=""TextFields"" /></TD>"
										Response.Write "</TR>"
									Else
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DocumentNumber"" ID=""DocumentNumberHdn"" VALUE=""."" />"
									End If
								End If
								Response.Write "<TR>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Observaciones:&nbsp;</NOBR></FONT></TD>"
									Response.Write "<TD><INPUT TYPE=""CHECKBOX"" NAME=""ForReasons"" ID=""ForReasonsTxt"" SIZE=""2"" MAXLENGTH=""2"" onclick=""if (this.checked) {ShowDisplay(document.all['ReasonsDiv']) } else {HideDisplay(document.all['ReasonsDiv'])}; "" /></TD>"
								Response.Write "</TR>"
							Response.Write "</TABLE>"
							Response.Write "<DIV NAME=""ReasonsDiv"" ID=""ReasonsDiv"" STYLE=""display: none"">"
								Response.Write "<TEXTAREA NAME=""Reasons"" ID=""ReasonsTxtArea"" ROWS=""5"" COLS=""50"" MAXLENGTH=""2000"" VALUE="""" CLASS=""TextFields"">" & aAbsenceComponent(S_REASONS_ABSENCE) & "</TEXTAREA>"
							Response.Write "</DIV>"
							If B_ISSSTE Then
								If (Len(oRequest("AbsenceChange").Item) = 0) And ((aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Or (Len(oRequest("AbsenceReview").Item) > 0)) Then
									If InStr(1, sAction, "Employees") = 0 Then
										Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
											Response.Write "<TR>"
												Response.Write "<TD VALIGN=""TOP"" ALIGN=""LEFT"" WIDTH=""72%"">"
													Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
															Response.Write "<TR><BR />"
																Response.Write "<TD VALIGN=""TOP"">"
																	Response.Write "<IFRAME SRC=""BrowserMonthForPayments.asp?EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & "&EmployeeDate=" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & "&FromArrow=1&ReasonID=1&" & sExtraURL & """ NAME=""BrowserMonthIFrame"" FRAMEBORDER=""0"" WIDTH=""330"" HEIGHT=""130""></IFRAME>"
																Response.Write "</TD>"
																Response.Write "<TD>&nbsp;&nbsp;&nbsp;</TD>"
																Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fechas de las incidencias:&nbsp;</FONT></TD>"
																Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""OcurredDates"" ID=""OcurredDatesLst"" SIZE=""7"" MULTIPLE=""1""></SELECT></TD>"
																Response.Write "<TD VALIGN=""BOTTOM""><A HREF=""javascript: RemoveSelectedItemsFromList(null, document.AbsencesFrm.OcurredDates)""><IMG SRC=""Images/BtnCrclDelete.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Eliminar"" BORDER=""0"" HSPACE=""5"" /></A></TD>"
															Response.Write "</TR>"
													Response.Write "</TABLE>"
											Response.Write "</TD>"
											Response.Write "<TD ALIGN=""CENTER"" VALIGN=""BOTTOM"" WIDTH=""28%"">"
											If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Or (Len(oRequest("AbsenceReview").Item) > 0) Then
												If InStr(1, sAction, "Employees") = 0 Then
													If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
												End If
											ElseIf Len(oRequest("Delete").Item) > 0 Then
												If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveAbsenceWngDiv']); AbsencesFrm.Remove.focus()"" />"
											Else
												If CInt(oRequest("JustificationID").Item) = -1 Then
													Call DisplayErrorMessage("Advertencia", "Este tipo de incidencia no es justificable por ninguna clave.")
													Response.Write "<BR /><BR />"
												Else
													'If (CInt(Request.Cookies("SIAP_SectionID")) <> 7) And aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" onClick=""ShowDisplay(document.all['ModifyAbsenceWngDiv']);"" />"
													If (CInt(Request.Cookies("SIAP_SectionID")) <> 7) And aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""AddBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
												End If
											End If
											Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""150"" HEIGHT=""1"" />"
											If aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1 Then
												Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=Absences'"" />"
											Else
												Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=Absences&EmployeeId=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & "'"" />"
											End If
											Response.Write "<BR /><BR />"
											Call DisplayWarningDiv("ModifyAbsenceWngDiv", "¿Está seguro que desea modificar el registro en la base de datos?")
											Call DisplayWarningDiv("RemoveAbsenceWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
											Call DisplayWarningForMoveAbsences("ApplyAbsenceWngDiv", "El empleado tiene registradas vacaciones, las cuales se recorreran al registrar la incapacidad ¿Está seguro que desea proceder con el registro?")
											Response.Write "</TD>"
											Response.Write "</TR>"
										Response.Write "</TABLE>"
									End If
								Else
									If Len(oRequest("AbsenceChange").Item) > 0 Then
										Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
											Response.Write "<TR NAME=""JustificationDiv"" ID=""JustificationDiv"">"
												If Len(oRequest("CancelAbsence").Item) > 0 Then
													Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cancelación de la incidencia:&nbsp;</FONT></TD>"
													Response.Write "<TD><SELECT NAME=""AbsenceID"" ID=""AbsenceID"" SIZE=""1"" CLASS=""Lists"">"
															aAbsenceComponent(N_JUSTIFICATION_ID_ABSENCE) = 0
															Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Justifications", "JustificationID", "JustificationShortName, JustificationName", "(JustificationID=" & aAbsenceComponent(N_JUSTIFICATION_ID_ABSENCE) & ")", "JustificationShortName", aAbsenceComponent(N_JUSTIFICATION_ID_ABSENCE), "Ninguno;;;-1", sErrorDescription)
													Response.Write "</SELECT>&nbsp;"
													Response.Write "</TD>"
												Else
													Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Justificación de la incidencia:&nbsp;</FONT></TD>"
													Response.Write "<TD><SELECT NAME=""AbsenceID"" ID=""AbsenceID"" SIZE=""1"" CLASS=""Lists"">"
															Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Justifications", "JustificationID", "JustificationShortName, JustificationName", "(JustificationID=" & aAbsenceComponent(N_JUSTIFICATION_ID_ABSENCE) & ")", "JustificationShortName", aAbsenceComponent(N_JUSTIFICATION_ID_ABSENCE), "Ninguno;;;-1", sErrorDescription)
													Response.Write "</SELECT>&nbsp;"
													Response.Write "</TD>"
												End If
											Response.Write "</TR>"
											Response.Write "<TR>"
												Response.Write "<TD>"
													Response.Write "<BR />"
													If Len(oRequest("CancelAbsence").Item) > 0 Then
														If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" />"
													Else
														If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Justificar"" CLASS=""Buttons"" />"
													End If
												Response.Write "</TD>"
											Response.Write "</TR>"
										Response.Write "</TABLE><BR />"
									End If
								End If
							End If
							Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""340"" HEIGHT=""1"" /><BR /><BR />"
							Response.Write "</FORM>"
							Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
							If Len(oRequest("AbsenceChange").Item) = 0 Then
								Response.Write "ShowHideAbsencesFields(document.AbsencesFrm.AbsenceID.value);" & vbNewLine
							End If
							Response.Write "//--></SCRIPT>" & vbNewLine
						End If
					End If
			End If
	End If

	DisplayAbsenceForm = lErrorNumber
	Err.Clear
End Function

Function DisplayAbsenceAsHiddenFields(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an absence for the
'         employee using hidden form fields
'Inputs:  oRequest, oADODBConnection, aAbsenceComponent
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAbsenceAsHiddenFields"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AbsenceID"" ID=""AbsenceIDHdn"" VALUE=""" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OcurredDate"" ID=""OcurredDateHdn"" VALUE=""" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EndDate"" ID=""EndDateHdn"" VALUE=""" & aAbsenceComponent(N_END_DATE_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RegistrationDate"" ID=""RegistrationDateHdn"" VALUE=""" & aAbsenceComponent(N_REGISTRATION_DATE_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DocumentNumber"" ID=""DocumentNumberHdn"" VALUE=""" & aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AbsenceHours"" ID=""AbsenceHoursHdn"" VALUE=""" & aAbsenceComponent(N_HOURS_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JustificationID"" ID=""JustificationIDHdn"" VALUE=""" & aAbsenceComponent(N_JUSTIFICATION_ID_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AppliesForPunctuality"" ID=""AppliesForPunctualityHdn"" VALUE=""" & aAbsenceComponent(N_APPLIES_FOR_PUNCTUALITY_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Reasons"" ID=""ReasonsHdn"" VALUE=""" & aAbsenceComponent(S_REASONS_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AddUserID"" ID=""AddUserIDHdn"" VALUE=""" & aAbsenceComponent(N_ADD_USER_ID_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AppliedDate"" ID=""AppliedDateHdn"" VALUE=""" & aAbsenceComponent(N_APPLIED_DATE_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Removed"" ID=""RemovedHdn"" VALUE=""" & aAbsenceComponent(N_REMOVED_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RemoveUserID"" ID=""RemoveUserIDHdn"" VALUE=""" & aAbsenceComponent(N_REMOVE_USER_ID_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RemovedDate"" ID=""RemovedDateHdn"" VALUE=""" & aAbsenceComponent(N_REMOVED_DATE_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AppliedRemoveDate"" ID=""AppliedRemoveDateHdn"" VALUE=""" & aAbsenceComponent(N_APPLIED_REMOVE_DATE_ABSENCE) & """ />"

	DisplayAbsenceAsHiddenFields = Err.number
	Err.Clear
End Function

Function DisplayAbsencesTable(oRequest, oADODBConnection, lIDColumn, bForExport, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To display the absences for the given absence for
'		  the employee from the database in a table
'Inputs:  oRequest, oADODBConnection, bForExport, aAbsenceComponent
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAbsencesTable"
	Dim oRecordset
	Dim iRecordCounter
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sFontBegin
	Dim sFontEnd
	Dim sBoldBegin
	Dim sBoldEnd
	Dim sNames
	Dim lErrorNumber
	Dim oStartDate
	Dim lDate
	Dim sConcept

	oStartDate = Now()
	lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	lErrorNumber = GetAbsences(oRequest, oADODBConnection, aAbsenceComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetch(oRequest, CInt(oRequest("StartPage").Item), ROWS_REPORT, oRecordset)
			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If (Not bForExport) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Or (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
					Select Case CLng(oRequest("ReasonID").Item)
						Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
							asColumnsTitles = Split("Acciones,Empleado,Concepto,Fecha de ocurrencia,Horas/Días,Fecha de registro,Nómina en que se aplica,Usuario que Registró", ",", -1, vbBinaryCompare)
							asCellWidths = Split("200,100,200,200,100,100,200,200", ",", -1, vbBinaryCompare)
							asCellAlignments = Split("CENTER,,,,CENTER,CENTER,CENTER,", ",", -1, vbBinaryCompare)
						Case Else
							If Len(oRequest("Tab").Item) <> 0 Then
								asColumnsTitles = Split("Empleado,Incidencia,F. de incidencia,F. de termino,Días,F. de registro,Nómina en que se aplica,Justificada,F. de alta,Registró,F. de justif.,Justificó", ",", -1, vbBinaryCompare)
								asCellWidths = Split("100,400,200,200,100,200,200,200,200,200,200,200,200", ",", -1, vbBinaryCompare)
								asCellAlignments = Split(",,,,CENTER,CENTER,,,,", ",", -1, vbBinaryCompare)
							Else
								asColumnsTitles = Split("Acciones,Empleado,Incidencia,F. de incidencia,F. de termino,Días,F. de registro,Nómina en que se aplica,Justificada,F. de alta,Registró,F. de justif.,Justificó", ",", -1, vbBinaryCompare)
								asCellWidths = Split("200,100,400,200,200,100,200,200,200,200,200,200,200,200", ",", -1, vbBinaryCompare)
								asCellAlignments = Split("CENTER,,,,,CENTER,CENTER,,,,", ",", -1, vbBinaryCompare)
							End If
					End Select
				Else
					Select Case CLng(oRequest("ReasonID").Item)
						Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
							asColumnsTitles = Split("Empleado,Concepto,Fecha de ocurrencia,Horas/Días,Fecha de registro,Nómina en que se aplica,Usuario que Registró", ",", -1, vbBinaryCompare)
							asCellWidths = Split("100,200,200,100,100,200,200", ",", -1, vbBinaryCompare)
							asCellAlignments = Split(",,,,CENTER,CENTER,,,", ",", -1, vbBinaryCompare)
						Case Else
							asColumnsTitles = Split("Empleado,Incidencia,F. de incidencia,F. de termino,Días,F. de registro,Nómina en que se aplica,Justificada,F. de alta,Registró,F. de justif.,Justificó", ",", -1, vbBinaryCompare)
							asCellWidths = Split("20,100,400,200,200,100,200,200,200,200,200,200,200,200", ",", -1, vbBinaryCompare)
							asCellAlignments = Split(",,,,CENTER,CENTER,,,,", ",", -1, vbBinaryCompare)
					End Select
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
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If (StrComp(CStr(oRecordset.Fields("AbsenceID").Value), oRequest("AbsenceID").Item, vbBinaryCompare) = 0) And (StrComp(CStr(oRecordset.Fields("OcurredDate").Value), oRequest("OcurredDate").Item, vbBinaryCompare) = 0) Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sFontBegin = ""
					sFontEnd = ""
					If CInt(oRecordset.Fields("Removed").Value) = 1 Then
						sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
						sFontEnd = "</FONT>"
					End If
					sRowContents = ""
					If (Not bForExport) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Or (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "&nbsp;"
						If aAbsenceComponent(N_ACTIVE_ABSENCE) >= 1 Then
							If (CInt(oRecordset.Fields("Removed").Value) = 0) And (CInt(oRecordset.Fields("JustificationID").Value) = -1) Then
								Select Case CLng(oRequest("ReasonID").Item)
									Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
										If CInt(oRecordset.Fields("Active").Value) = 1 Then
											sRowContents = sRowContents & "<A HREF=""" & "UploadInfo.asp" & "?Action=EmployeesMovements&SaveEmployeesMovements=1&DeactiveConcept=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&ReasonID=" & lReasonID & """>"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnDeactive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Desactivar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
										Else
											sRowContents = sRowContents & "<A HREF=""" & "UploadInfo.asp" & "?Action=EmployeesMovements&SaveEmployeesMovements=1&ActiveConcept=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&ReasonID=" & lReasonID & """>"
												sRowContents = sRowContents & "<IMG SRC=""Images/BtnActive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Activar"" BORDER=""0"" />"
											sRowContents = sRowContents & "</A>&nbsp;"
										End If
									Case Else
										If Len(oRequest("Tab").Item) = 0 Then
											If	CInt(Request.Cookies("SIAP_SectionID")) <> 7 Then
												If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
													Select Case CInt(oRecordset.Fields("AbsenceID").Value)
														Case 41,42,43,44,45,46,47,48,49,57,58,50,51,52,53,54,55,56
															sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""8"" />"
															sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
														Case Else
															sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Absences&EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & "&ForJustificationID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&OcurredDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&RegistrationDate=" & CStr(oRecordset.Fields("RegistrationDate").Value) & "&AbsenceChange=1&JustificationID=" & CStr(oRecordset.Fields("WithJustification").Value) & "&Active=" & CStr(oRecordset.Fields("Active").Value) & """>"
																sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Justificar incidencia"" BORDER=""0"" />"
															sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
															sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Absences&EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & "&ForJustificationID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&OcurredDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&RegistrationDate=" & CStr(oRecordset.Fields("RegistrationDate").Value) & "&AbsenceChange=1&CancelAbsence=1"">"
																sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar incidencia"" BORDER=""0"" />"
															sRowContents = sRowContents & "</A>"
													End Select
												Else
													sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""8"" />"
													sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
												End If
											Else
												sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""8"" />"
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
											End If
										End If
								End Select
							Else
								If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
									sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=Absences&EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & "&AbsenceID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&OcurredDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&AppliedDate=" & CStr(oRecordset.Fields("AppliedDate").Value) & "&RegistrationDate=" & CStr(oRecordset.Fields("RegistrationDate").Value) &  "&FilterStartYear=" & oRequest("FilterStartYear").Item & "&FilterStartMonth=" & oRequest("FilterStartMonth").Item & "&FilterStartDay=" & oRequest("FilterStartDay").Item & "&FilterEndYear=" & oRequest("FilterEndYear").Item & "&FilterEndMonth=" & oRequest("FilterEndMonth").Item & "&FilterEndDay=" & oRequest("FilterEndDay").Item & "&Tab=4&Modify=1&CancelJustification=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar Justificación/Cancelación"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>"
								Else
									sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""8"" />"
									sRowContents = sRowContents & "&nbsp;"
								End If
							End If
						Else
							If (CInt(oRecordset.Fields("Removed").Value) = 0) And (CInt(oRecordset.Fields("JustificationID").Value) = -1) Then
								If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
									Select Case CLng(oRequest("ReasonID").Item)
										Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
											sRowContents = sRowContents & "<A HREF=""" & "UploadInfo.asp" & "?Action=EmployeesMovements&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&ReasonID=" & CStr(oRequest("ReasonID").Item) & """>"
										Case Else
											sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=Absences&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&AbsenceID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&OcurredDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&RegistrationDate=" & CStr(oRecordset.Fields("RegistrationDate").Value) & "&FilterStartYear=" & oRequest("FilterStartYear").Item & "&FilterStartMonth=" & oRequest("FilterStartMonth").Item & "&FilterStartDay=" & oRequest("FilterStartDay").Item & "&FilterEndYear=" & oRequest("FilterEndYear").Item & "&FilterEndMonth=" & oRequest("FilterEndMonth").Item & "&FilterEndDay=" & oRequest("FilterEndDay").Item & "&Tab=4&Remove=1"">"
									End Select
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar registro"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>"
								Else
									sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""2"" HEIGHT=""8"" />"
									sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
								End If
							Else
								'If VerifyPayrollIsActive(oADODBConnection, CLng(oRecordset.Fields("AppliedRemoveDate").Value), N_PAYROLL_FOR_ABSENCES, sErrorDescription) Then
									If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
										Select Case CLng(oRequest("ReasonID").Item)
											Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
												sRowContents = sRowContents & "<A HREF=""" & "UploadInfo.asp" & "?Action=EmployeesMovements&SaveEmployeesMovements=1&CancelMotion=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&ReasonID=" & CStr(oRequest("ReasonID").Item) & """>"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar"" BORDER=""0"" />"
											Case Else
												sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=Absences&EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & "&AbsenceID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&OcurredDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&AppliedDate=" & CStr(oRecordset.Fields("AppliedDate").Value) & "&RegistrationDate=" & CStr(oRecordset.Fields("RegistrationDate").Value) &  "&FilterStartYear=" & oRequest("FilterStartYear").Item & "&FilterStartMonth=" & oRequest("FilterStartMonth").Item & "&FilterStartDay=" & oRequest("FilterStartDay").Item & "&FilterEndYear=" & oRequest("FilterEndYear").Item & "&FilterEndMonth=" & oRequest("FilterEndMonth").Item & "&FilterEndDay=" & oRequest("FilterEndDay").Item & "&Tab=4&Modify=1&CancelJustification=1"">"
													sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar Justificación"" BORDER=""0"" />"
										End Select
										sRowContents = sRowContents & "</A>"
									Else
										sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""2"" HEIGHT=""8"" />"
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
									End If
								'Else
								'	sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""2"" HEIGHT=""8"" />"
								'	sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
								'End If
							End If
							If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
								If (CInt(Request.Cookies("SIAP_SectionID")) <> 7) And (CInt(Request.Cookies("SIAP_SubSectionID")) <> 422) Then
									Select Case CLng(oRequest("ReasonID").Item)
										Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
											sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=EmployeesMovements&SaveEmployeesMovements=1&Authorization=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&ConceptAmount=" & CStr(oRecordset.Fields("AbsenceHours").Value) & "&ReasonID=" & CStr(oRequest("ReasonID").Item) &""">"
										Case Else
											sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SetActive=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & """>"
									End Select
										sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Enviar a validación"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								Else
									sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""2"" HEIGHT=""8"" />"
									sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
								End If
							End If
							If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) And (iActive=0) Then
								Select Case CLng(oRequest("ReasonID").Item)
									Case EMPLOYEES_EXTRAHOURS, EMPLOYEES_SUNDAYS
										sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""2"" HEIGHT=""8"" />"
										sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
									Case Else
										If (CInt(Request.Cookies("SIAP_SectionID")) <> 7) And (CInt(Request.Cookies("SIAP_SubSectionID")) <> 422) Then
											sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("AbsenceID").Value) & CStr(oRecordset.Fields("OcurredDate").Value) & """ ID=""" & CStr(oRecordset.Fields("EmployeeID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("EmployeeID").Value) & """ CHECKED=""1"" &/>"
										Else
											sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""2"" HEIGHT=""8"" />"
											sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
										End If
								End Select
							End If
						End If
						If Len(oRequest("Tab").Item) = 0 Then
							sRowContents = sRowContents & "&nbsp;" & TABLE_SEPARATOR
						End If
					End If
					sRowContents = sRowContents & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeID").Value)) & sBoldEnd & sFontEnd
					If (CLng(oRecordset.Fields("AbsenceID").Value) = 35) Or (CLng(oRecordset.Fields("AbsenceID").Value) = 37) Or (CLng(oRecordset.Fields("AbsenceID").Value) = 38) Or (CLng(oRecordset.Fields("AbsenceID").Value) = 39) Or (CLng(oRecordset.Fields("AbsenceID").Value) = 40) Then
						If (CLng(oRecordset.Fields("AbsenceID").Value) = 39) Or (CLng(oRecordset.Fields("AbsenceID").Value) = 40) Then
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceShortName").Value) & ". " & CStr(oRecordset.Fields("AbsenceName").Value) & " - Periodo " & Left(CStr(oRecordset.Fields("VacationPeriod").Value), 4) & "-" & Right(CStr(oRecordset.Fields("VacationPeriod").Value), 2)) & sBoldEnd & sFontEnd
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceShortName").Value) & ". " & CStr(oRecordset.Fields("AbsenceName").Value) & " - Periodo " & Left(CStr(oRecordset.Fields("VacationPeriod").Value), 4) & "-" & Right(CStr(oRecordset.Fields("VacationPeriod").Value), 1)) & sBoldEnd & sFontEnd
						End If
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceShortName").Value) & ". " & CStr(oRecordset.Fields("AbsenceName").Value)) & sBoldEnd & sFontEnd
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
					If (CLng(oRequest("ReasonID").Item) <> EMPLOYEES_EXTRAHOURS) And (CLng(oRequest("ReasonID").Item) <> EMPLOYEES_SUNDAYS) Then
						If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("A la fecha") & sBoldEnd & sFontEnd
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
						End If
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceHours").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("RegistrationDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
						If CLng(oRecordset.Fields("AppliedDate").Value) = 0 Then
							sRowContents = sRowContents & CleanStringForHTML("Ninguna")
						Else
							Call GetNameFromTable(oADODBConnection, "Payrolls", CStr(oRecordset.Fields("AppliedDate").Value), "", "", sNames, sErrorDescription)
							If Len(sNames) > 0 Then
								sRowContents = sRowContents & CleanStringForHTML(sNames)
							Else
								sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("AppliedDate").Value), -1, -1, -1)
							End If
						End If
					sRowContents = sRowContents & sBoldEnd & sFontEnd
					If (CLng(oRequest("ReasonID").Item) <> EMPLOYEES_EXTRAHOURS) And (CLng(oRequest("ReasonID").Item) <> EMPLOYEES_SUNDAYS) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("JustificationShortName").Value) & ". " & CStr(oRecordset.Fields("JustificationName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("RegistrationDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("UserName").Value) & " " & CStr(oRecordset.Fields("UserLastName").Value)) & sBoldEnd & sFontEnd
					If (CLng(oRequest("ReasonID").Item) <> EMPLOYEES_EXTRAHOURS) And (CLng(oRequest("ReasonID").Item) <> EMPLOYEES_SUNDAYS) Then
						If CLng(oRecordset.Fields("RemovedDate").Value) = 0 Then
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "---" & sBoldEnd & sFontEnd
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "---" & sBoldEnd & sFontEnd
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("RemovedDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
							sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("RemoveUserName").Value) & " " & CStr(oRecordset.Fields("RemoveUserLastName").Value)) & sBoldEnd & sFontEnd
						End If
					End If
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					iRecordCounter = iRecordCounter + 1
					If (Not bForExport) And (iRecordCounter >= ROWS_REPORT) Then Exit Do
					If Err.Number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE></DIV>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			If CInt(Request.Cookies("SIAP_SectionID")) <> 7 Then  ' Dif. de Desc.
				If CInt(Request.Cookies("SIAP_SubSectionID")) = 22 Then  ' Igual a Prestaciones e incidencias
					Select Case CLng(oRequest("ReasonID").Item)
						Case EMPLOYEES_EXTRAHOURS
							sConcept = "horas extras"
						Case EMPLOYEES_SUNDAYS
							sConcept = "primas dominicales"
					End Select
				Else ' Igual a Inf. - Emp. - Inci
					sConcept = "incidencias"
				End If
			Else
				' Igual a Desc.
				If CInt(Request.Cookies("SIAP_SubSectionID")) = 721 Then  ' Igual a Prestaciones e incidencias
					Select Case CLng(oRequest("ReasonID").Item)
						Case EMPLOYEES_EXTRAHOURS
							sConcept = "horas extras"
						Case EMPLOYEES_SUNDAYS
							sConcept = "primas dominicales"
					End Select
				Else ' Igual a Inf. - Emp. - Inci
					sConcept = "incidencias"
				End If
			End If
			If aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1 Then
				If CInt(Request.Cookies("SIAP_SubSectionID")) = 22 Then  ' Igual a Prestaciones e incidencias
					If aAbsenceComponent(N_ACTIVE_ABSENCE) Then
						sErrorDescription = "Introduzca un número de empleado para consultar sus " & sConcept
					Else
						sErrorDescription = "No existen registros de " & sConcept & " en proceso de aplicación"
					End If
				Else
					sErrorDescription = "Introduzca un número de empleado para consultar sus " & sConcept
				End If
			Else
				If aAbsenceComponent(N_ACTIVE_ABSENCE) Then
					sErrorDescription = "No existen registros de " & sConcept & " para este empleado."
				Else
					sErrorDescription = "No se han registrado " & sConcept & " en proceso de aplicación para este empleado."
				End If
			End If
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayAbsencesTable = lErrorNumber
	Err.Clear
End Function

Function DisplayAbsencesForApplyTable(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: Reporte de totales de incidencias por aplicar
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAbsencesForApplyTable"
	Dim oRecordset
	Dim sCondition
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim iAbsencesCount

	iAbsencesCount=0
	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.RegistrationDate, EmployeesAbsencesLKP.AppliedDate, EmployeesAbsencesLKP.AbsenceID, Absences.AbsenceShortName, Absences.AbsenceName, COUNT(*) As Registros, SUM(AbsenceHours) As Dias From EmployeesAbsencesLKP, Absences Where (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And (EmployeesAbsencesLKP.Active=0) And (Absences.AbsenceID<100) Group By EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.RegistrationDate, EmployeesAbsencesLKP.AppliedDate, EmployeesAbsencesLKP.AbsenceID, Absences.AbsenceShortName, Absences.AbsenceName Order by EmployeesAbsencesLKP.OcurredDate, Absences.AbsenceShortName", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("Clave de la incidencia,Fecha de aplicación,Fecha de registro,Descripción,Registros,No. de días", ",", -1, vbBinaryCompare)
				asCellWidths = Split(",,,,,,", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If
				asCellAlignments = Split(",,,,,,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					iAbsencesCount=iAbsencesCount + 1
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("AbsenceShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(DisplayDateFromSerialNumber(CDbl(oRecordset.Fields("OcurredDate").Value), -1, -1, -1)))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(DisplayDateFromSerialNumber(CDbl(oRecordset.Fields("RegistrationDate").Value), -1, -1, -1)))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Registros").Value))
					If oRecordset.Fields("Dias").Value < 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("<CENTER>---</ CENTER>")
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Dias").Value))
					End If
					
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				Response.Write "</TABLE><BR /><BR />"
				Call DisplayInstructionsMessage("Número de registros", "Existen:&nbsp;" & iAbsencesCount & " claves distintas por aplicar.")
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existen registros de incidencias en proceso para ser aplicados."
			End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayAbsencesForApplyTable = lErrorNumber
	Err.Clear
End Function

Function DisplayPendingEmployeesAbscencesTable(oRequest, oADODBConnection, iActive, bForExport, lReasonID, sAction, aEmployeeComponent, sErrorDescription)
'*****************************************************************
'Purpose: To display the employees' credits of employees that were
'         captured by the users
'Inputs:  oRequest, oADODBConnection, bForExport, lStatusID
'Outputs: aEmployeeComponent, sErrorDescription
'*****************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPendingEmployeesAbscencesTable"
	Dim asFields
	Dim asKeyFields
	Dim sTabsDone
	Dim sCurrentTab
	Dim iIndex
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sFontBegin
	Dim sFontEnd
	Dim sBoldBegin
	Dim sBoldEnd
	Dim lErrorNumber
	Dim iEmployeeTypeID
	Dim sCondition
	Dim sFields
	Dim lStatusID
	Dim sQuery
	Dim sConceptID
	Dim sAccion

	If CLng(oRequest("AbsenceID").Item) > 0 Then
		sCondition = sCondition & " And (Absences.AbsenceID=" & CStr(oRequest("AbsenceID").Item) & ")"
	End If
	Call GetStartAndEndDatesFromURL("FilterStart", "FilterEnd", "OcurredDate", False, sCondition)
	If Len(sCondition) > 0 Then
		If InStr(1, sCondition , "And ", vbBinaryCompare) = 0 Then sCondition  = "And " & sCondition
	End If

	sQuery = "Select EA.EmployeeID, EA.AbsenceID, EmployeeName + ' ' + EmployeeLastName + ' ' + EmployeeLastName2 As EmployeeFullName," &_
			" EA.OcurredDate, EA.EndDate, EA.RegistrationDate, EA.DocumentNumber, EA.AbsenceHours, A.AbsenceShortName," & _
			" A.AbsenceName, J.JustificationShortName, EA.Reasons, EA.Removed, EA.AppliedDate," & _
			" EA.JustificationID As AbsenceJustified, A.IsJustified, A.JustificationID As WithJustification" & _
			" From Employees As E, EmployeesAbsencesLKP As EA, Absences As A, Justifications As J, Users As U" & _
			" Where (E.EmployeeID=EA.EmployeeID)" & _
			" And (EA.JustificationID=J.JustificationID)" & _
			" And (EA.AbsenceID=A.AbsenceID)" & _
			" And (EA.AddUserID=U.UserID) And (EA.AbsenceID Not IN (201, 202))"
			If aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) > 0 Then
				sQuery = sQuery & " And (EA.EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ")"
			Else
				sQuery = sQuery & " And (EA.EmployeeID=0)"
			End If
			If iActive = 0 Then
				sQuery = sQuery & " And (EA.Active<= " & iActive & ")"
			Else
				sQuery = sQuery & " And (EA.Active=" & iActive & ")"
			End If
			sQuery = sQuery & sCondition & " Order By EA.OcurredDate, EA.AbsenceID"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & sQuery & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE="&lReasonID&" />"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeDisplayTablesComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
			If Not bForExport Then
				Response.Write "0"
			Else
				Response.Write "1"
			End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			'If bForExport Or (iActive = 1) Then
			If bForExport Then
				If (InStr(1, sAction, "Pending", vbBinaryCompare) > 0) Or (InStr(1, sAction, "Rejected", vbBinaryCompare) > 0) Then
					If (lReasonID = 0) Then
						If (Len(oRequest("ShortFormat").Item) > 0) Then
							asColumnsTitles = Split("Empleado, F. inicio, F. término, F. de registro, Clave, Descripción, Justificada", ",", -1, vbBinaryCompare)
						Else
							asColumnsTitles = Split("No. Empleado, Nombre, Fecha de Ocurrencia, Fecha de término, Fecha de registro, Documento, Días, Clave, Descripción, Justificada, Observaciones", ",", -1, vbBinaryCompare)
						End If
					Else
						asColumnsTitles = Split("No. Empleado, Nombre, Fecha de inicio, Fecha de termino, Fecha de registro, Documento, Días, Clave, Descripción, Justificada, Observaciones", ",", -1, vbBinaryCompare)
					End If
					asCellWidths = Split(",,,,,,,,,", ",", -1, vbBinaryCompare)
					asCellAlignments = Split(",,,,,,,CENTER", ",", -1, vbBinaryCompare)
				Else
					If (lReasonID = 0) Then
						If (Len(oRequest("ShortFormat").Item) > 0) Then
							asColumnsTitles = Split("Empleado, F. inicio, F. término, F. de registro, Clave, Descripción, Justificada", ",", -1, vbBinaryCompare)
						Else
							asColumnsTitles = Split("No. Empleado, Nombre, Fecha de Ocurrencia, Fecha de término, Fecha de registro, Documento, Días, Clave, Descripción, Justificada, Observaciones", ",", -1, vbBinaryCompare)
						End If
					Else
						asColumnsTitles = Split("No. Empleado, Nombre, Fecha de inicio, Fecha de termino, Fecha de registro, Documento, Días, Clave, Descripción, Justificada, Observaciones", ",", -1, vbBinaryCompare)
					End If
					asCellWidths = Split(",,,,,,,,,", ",", -1, vbBinaryCompare)
					asCellAlignments = Split(",,,,,,CENTER", ",", -1, vbBinaryCompare)
				End If
			Else
				If (InStr(1, sAction, "Pending", vbBinaryCompare) > 0) Or (InStr(1, sAction, "Rejected", vbBinaryCompare) > 0) Then
					If (lReasonID = 0) Then
						asColumnsTitles = Split("No. Empleado, Nombre, Fecha de Ocurrencia, Fecha de término, Fecha de registro, Documento, Días, Clave, Descripción, Justificada, Observaciones, Acciones", ",", -1, vbBinaryCompare)
					Else
						asColumnsTitles = Split("No. Empleado, Nombre, Fecha de inicio, Fecha de termino, Fecha de registro, Documento, Días, Clave, Descripción, Justificada, Observaciones, Acciones", ",", -1, vbBinaryCompare)
					End If
					asCellWidths = Split(",,,,,,,,,,,",",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,,,,,,,CENTER", ",", -1, vbBinaryCompare)
				Else
					If (lReasonID = 0) Then
						asColumnsTitles = Split("No. Empleado, Nombre, Fecha de Ocurrencia, Fecha de término, Fecha de registro, Documento, Días, Clave, Descripción, Justificada, Observaciones, Acciones", ",", -1, vbBinaryCompare)
					Else
						asColumnsTitles = Split("No. Empleado, Nombre, Fecha de inicio, Fecha de termino, Fecha de registro, Documento, Días, Clave, Descripción, Justificada, Observaciones, Acciones", ",", -1, vbBinaryCompare)
					End If
					asCellWidths = Split(",,,,,,,,,,,", ",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,,,,,,,CENTER", ",", -1, vbBinaryCompare)
				End If
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
			sBoldBegin = "<B>"
			sBoldEnd = "</B>"
			sFontBegin = ""
			sFontEnd = ""
			Do While Not oRecordset.EOF
				If bForExport Then
					sRowContents = "=T(""" & Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000")) & """)"
				Else
					sRowContents = Right("000000" & CStr(oRecordset.Fields("EmployeeID").Value), Len("000000"))
				End If
				If (Len(oRequest("ShortFormat").Item) = 0) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeFullName").Value))
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value))
				If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("A la fecha")
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value))
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("RegistrationDate").Value))
				If (Len(oRequest("ShortFormat").Item) = 0) Then
					If Len(CStr(oRecordset.Fields("DocumentNumber").Value)) <= 1 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("NA")
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value))
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceHours").Value))
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceShortName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AbsenceName").Value))
				sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("JustificationShortName").Value))
				If (Len(oRequest("ShortFormat").Item) = 0) Then
					If Len(CStr(oRecordset.Fields("Reasons").Value)) <= 1 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("Ninguna")
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("Reasons").Value))
					End If
				End If
				If Not bForExport Then
					If iActive = 0 Then
						If (CInt(oRecordset.Fields("Removed").Value) = 0) And (CInt(oRecordset.Fields("AbsenceJustified").Value) = -1) Then
							If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&Remove=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&AbsenceID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&OcurredDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar registro"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
							End If
						Else
							If VerifyPayrollIsActive(oADODBConnection, CLng(oRecordset.Fields("AppliedRemoveDate").Value), N_PAYROLL_FOR_ABSENCES, sErrorDescription) Then
								If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
									sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & GetASPFileName("") & "?Action=Absences&EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & "&AbsenceID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&OcurredDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&AppliedDate=" &  CStr(oRecordset.Fields("AppliedDate").Value) & "&RegistrationDate=" & CStr(oRecordset.Fields("RegistrationDate").Value) & "&FilterStartYear=" & oRequest("FilterStartYear").Item & "&FilterStartMonth=" & oRequest("FilterStartMonth").Item & "&FilterStartDay=" & oRequest("FilterStartDay").Item & "&FilterEndYear=" & oRequest("FilterEndYear").Item & "&FilterEndMonth=" & oRequest("FilterEndMonth").Item & "&FilterEndDay=" & oRequest("FilterEndDay").Item & "&Tab=4&Modify=1&CancelJustification=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar Justificación"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>"
								Else
									sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""8"" />"
									sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
								End If
							Else
								sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""8"" />"
								sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"								
							End If
						End If
						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) And (CInt(Request.Cookies("SIAP_SectionID")) <> 7) Then
								sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&SetActive=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&ConceptID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&ConceptStartDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Enviar a validación"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
							End If
						End If
						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) And (iActive=0) Then
							If (CInt(Request.Cookies("SIAP_SectionID")) <> 7) Then
								sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("AbsenceID").Value) & CStr(oRecordset.Fields("OcurredDate").Value) & """ ID=""" & CStr(oRecordset.Fields("EmployeeID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("EmployeeID").Value) & """ CHECKED=""1"" &/>"
							End If
						End If
					Else
						If (CInt(oRecordset.Fields("Removed").Value) = 0) And (CInt(oRecordset.Fields("AbsenceJustified").Value) = -1) Then
							If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&AbsenceChange=1&CancelAbsence=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&AbsenceID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&OcurredDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar incidencia"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
							End If
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR
								sRowContents = sRowContents & "<IMG SRC=""Images/Transparent.gif"" WIDTH=""10"" HEIGHT=""8"" BORDER=""0"" />"
							sRowContents = sRowContents & "&nbsp;"
						End If
						If (CInt(oRecordset.Fields("IsJustified").Value) <> 0) And (CInt(oRecordset.Fields("AbsenceJustified").Value) = -1) Then
							If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
								sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "UploadInfo.asp" & "?Action=" & sAction & "&AbsenceChange=1&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&AbsenceID=" & CStr(oRecordset.Fields("AbsenceID").Value) & "&OcurredDate=" & CStr(oRecordset.Fields("OcurredDate").Value) & "&JustificationID=" & CStr(oRecordset.Fields("WithJustification").Value) & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Justificar incidencia"" BORDER=""0"" />"
								sRowContents = sRowContents & "</A>&nbsp;"
							End If
						End If
					End If
					sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
					sFontEnd = "</FONT>"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
				Else
					sFontBegin = "<FONT COLOR=""#" & S_WARNING_FOR_GUI & """>"
					sFontEnd = "</FONT>"
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				End If
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			Response.Write "</TABLE><BR /><BR />"
		Else
			If CInt(aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE)) > 0 Then
				lErrorNumber = L_ERR_NO_RECORDS
				If iActive Then
					sErrorDescription = "El empleado seleccionado no tiene registros de incidencias."
				Else
					sErrorDescription = "El empleado seleccionado no tiene registros de incidencias en proceso."
				End If
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "Seleccione un número de empleado para buscar sus registros."
			End If
		End If
	End If

	Set oRecordset = Nothing
	DisplayPendingEmployeesAbscencesTable = lErrorNumber
	Err.Clear
End Function

Function SetActiveForEmployeeAbsence(oRequest, oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given employee's concept
'Inputs:  oRequest, oADODBConnection
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForEmployeeConcept"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador de la incidencia y/o la fecha para agregar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo modificar la información del concepto."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set Active=1, AppliedDate=" & aAbsenceComponent(N_APPLIED_DATE_ABSENCE) & " Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ") And (OcurredDate=" & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	SetActiveForEmployeeConcept = lErrorNumber
	Err.Clear
End Function

Function VerifyAbsencesForPeriod(oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if an absence apply for period
'Inputs:  oADODBConnection, aAbsenceComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyAbsencesForPeriod"
	Dim lErrorNumber
	Dim oRecordset
	Dim bComponentInitialized

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from Absences Where (AbsenceID = " & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If CLng(oRecordset.Fields("IsForPeriod").Value) = 1 Then
				VerifyAbsencesForPeriod = True
			Else
				sErrorDescription = "La incidencia no se registra por periodo."
				VerifyAbsencesForPeriod = False
			End If
			oRecordset.Close
		Else
			sErrorDescription = "No se encontro la incidencia en el catálogo."
			VerifyAbsencesForPeriod = False
		End If
	Else
		sErrorDescription = "Error al verificar si la incidencia se registra por periodo."
		VerifyAbsencesForPeriod = False
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyAbsenceIsJustification(oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if an absence apply for period
'Inputs:  oADODBConnection, aAbsenceComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyAbsenceIsJustification"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador de la incidencia para validar si es justificación."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If InStr(1, ",8,9,15,47,", "," & CStr(aAbsenceComponent(N_ABSENCE_ID_ABSENCE)) & ",", vbBinaryCompare) > 0 Then
			VerifyAbsenceIsJustification = True
		Else
			sErrorDescription = "La incidencia seleccionada no es justificación."
			VerifyAbsenceIsJustification = False
		End If
	End If
End Function

Function VerifyAbsenceOccurredDateLimit(oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To get the dates for all the absences for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAbsenceComponent, sDates, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyAbsenceOccurredDateLimit"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim lDateLimit

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	lDateLimit = aAbsenceComponent(N_APPLIED_DATE_ABSENCE)
	lDateLimit = GetPayrollStartDate(lDateLimit)
	lDateLimit = AddMonthsToSerialDate(lDateLimit, -1)

	If aAbsenceComponent(N_OCURRED_DATE_ABSENCE) < lDateLimit Then
		VerifyAbsenceOccurredDateLimit = False
	Else
		VerifyAbsenceOccurredDateLimit = True
	End If
End Function

Function VerifyAbsenceType(oADODBConnection, aAbsenceComponent, sAbsenceType, sErrorDescription)
'************************************************************
'Purpose: To verify if an absence apply for period
'Inputs:  oADODBConnection, aAbsenceComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyAbscenceIsSuspension"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador de la incidencia para validar su tipo."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If InStr(1, ",8,9,15,47,", "," & CStr(aAbsenceComponent(N_ABSENCE_ID_ABSENCE)) & ",", vbBinaryCompare) > 0 Then
			sAbsenceType = "Justification"
		ElseIf InStr(1, ",41,42,43,44,45,46,47,48,49,57,58,", "," & CStr(aAbsenceComponent(N_ABSENCE_ID_ABSENCE)) & ",", vbBinaryCompare) > 0 Then
			sAbsenceType = "Suspension"
		ElseIf InStr(1, ",35,37,38,", "," & CStr(aAbsenceComponent(N_ABSENCE_ID_ABSENCE)) & ",", vbBinaryCompare) > 0 Then
			sAbsenceType = "Vacation"
		ElseIf InStr(1, ",30, 83, 89, 90,", "," & CStr(aAbsenceComponent(N_ABSENCE_ID_ABSENCE)) & ",", vbBinaryCompare) > 0 Then
			sAbsenceType = "Licence"
		End If
	End If
End Function

Function VerifyAnualDiferenceOfAbsences(oADODBConnection, aAbsenceComponent, bAnualCalendar, sErrorDescription)
'************************************************************
'Purpose: To verify if employee absences exist with diference minimum of one year
'		or if the 0853 Abscence don't exceed 8 days in a year
'Inputs:  oADODBConnection, aAbsenceComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyAnualDiferenceOfAbsences"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim lStartDate
	Dim lEndDate
	Dim iTotalDays

	If (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 31) Then
		lStartDate = CLng(Left(CStr(aAbsenceComponent(N_OCURRED_DATE_ABSENCE)), Len("0000"))& "0101")
		lEndDate = CLng(Left(CStr(aAbsenceComponent(N_OCURRED_DATE_ABSENCE)), Len("0000"))& "1231")
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select SUM(AbsenceHours) As TotalDays From EmployeesAbsencesLKP Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ") And (OcurredDate>=" & lStartDate & ") And (EndDate<=" & lEndDate & ")" , "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAbsencesLKP Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ") And (Active = 1) Order By StartDate Desc", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 31) Then
				If IsNull(oRecordset.Fields("TotalDays").Value) Then
					iTotalDays = 0
				Else
					iTotalDays = CInt(oRecordset.Fields("TotalDays").Value)
				End If
				If iTotalDays + CInt(aAbsenceComponent(N_HOURS_ABSENCE)) <= 8 Then
					VerifyAnualDiferenceOfAbsences = True
				Else
					sErrorDescription = "Solamente puede agregar " & CInt(8 - iTotalDays) & " días para registrar los días por cuidados maternos en el año indicado."
					VerifyAnualDiferenceOfAbsences = False
				End If
			Else
				If bAnualCalendar Then
					lStartDate = CLng(Left(CStr(oRecordset.Fields("OcurredDate").Value), Len("0000"))& "0101")
					lEndDate = CLng(Left(CStr(oRecordset.Fields("OcurredDate").Value), Len("0000"))& "1231")
					If (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) >= lStartDate) And (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) <= lEndDate) Then
						sErrorDescription = "La incidencia ya fué registrada en este año calendario"
						VerifyAnualDiferenceOfAbsences = False
					Else
						VerifyAnualDiferenceOfAbsences = True
					End If
				Else
					lEndDate = CLng(AddDaysToSerialDate(CLng(oRecordset.Fields("OcurredDate").Value), 365))
					If aAbsenceComponent(N_OCURRED_DATE_ABSENCE) > lEndDate Then
						VerifyAnualDiferenceOfAbsences = True
					Else
						sErrorDescription = "La incidencia no cubre el plazo de un año de haberse registrado"
						VerifyAnualDiferenceOfAbsences = False
					End If
				End If
			End If
		Else
			VerifyAnualDiferenceOfAbsences = True
		End If
		oRecordset.Close
	Else
		If (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 31) Then
			sErrorDescription = "Error al verificar si la incidencia de días por cuidados maternos cubre el límite de 8 días por año"
		Else
			sErrorDescription = "Error al verificar si la incidencia cubre el plazo de un año de haberse registrado"
		End If
		VerifyAnualDiferenceOfAbsences = False
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyExistenceOfAbsences(oADODBConnection, aAbsenceComponent, sAbsenceIDs, bIsForPeriod, sErrorDescription)
'************************************************************
'Purpose: To verify if an absence already exist in database
'Inputs:  oADODBConnection, aAbsenceComponent, sAbsenceIDs, bIsForPeriod
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistenceOfAbsences"
	Dim lErrorNumber
	Dim oRecordset
	Dim oRecordset1
	Dim sErrorDescription1
	Dim sQuery
	Dim sCondition
	Dim lOcurredDate
	Dim sAbsenceCrossType
	Dim sAbsenceShortName
	Dim sExistingAbsenceShortNames
	Dim lAbsenceID
	Dim lStartDate
	Dim lStartDate1
	Dim lEndDate
	Dim lVacationPeriod
	Dim iDays
	Dim iActiveOriginal
	Dim bComponentInitialized
	Dim iDay

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (InStr(1, sAbsenceIDs, "9999", vbBinaryCompare) > 0) Then
		VerifyExistenceOfAbsences = True
	Else
		If bIsForPeriod Then
			If InStr(1, sAbsenceIDs, "9999", vbBinaryCompare) > 0 Then
				VerifyExistenceOfAbsences = True
			Else
				sQuery = "Select * From EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_ID_EMPLOYEE) & ")"
				If InStr(1, sAbsenceIDs, "-1", vbBinaryCompare) > 0 Then
					If InStr(1, ",41,42,43,44,45,46,47,48,49,57,58,", "," & CStr(aAbsenceComponent(N_ABSENCE_ID_ABSENCE)) & ",", vbBinaryCompare) > 0 Then
						sCondition = " And (AbsenceID IN (41,42,43,44,45,46,47,48,49,57,58))"
					End If
				Else
					'Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
					'	Case 10,11,12,13,14,16,17,82,83,84,85,86,87,29,30,31,32,33,34,35,37,38
					'		sCondition = " And (AbsenceID In (" & sAbsenceIDs & ",201,202))"
					'	Case Else
							sCondition = " And (AbsenceID IN (" & sAbsenceIDs & "))"
					'End Select
				End If
				sCondition = sCondition & " And (AbsenceID NOT IN (201, 202))"
				If InStr(1, ",50,51,52,53,54,55,56,", "," & CStr(aAbsenceComponent(N_ABSENCE_ID_ABSENCE)) & ",", vbBinaryCompare) = 0 Then
					sCondition = sCondition & " And (AbsenceID Not In (50,51,52,53,54,55,56))"
				End If
				sQuery = sQuery & sCondition & _
						 " And (((OcurredDate >= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (OcurredDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & "))" & _
						 " Or ((EndDate >= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (EndDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & "))" & _
						 " Or ((EndDate >= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (OcurredDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & ")))"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						If aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = ABSENCE_MEDICAL_LICENSE Then
							Call GetCrossingAbsenceType(oADODBConnection, aAbsenceComponent, sAbsenceIDs, sAbsenceCrossType, lAbsenceID, lStartDate, lEndDate, lVacationPeriod, iDays, sErrorDescription)
							Select Case sAbsenceCrossType
								Case "Left"
									Call MoveVacationsToLeft(oADODBConnection, aAbsenceComponent, lAbsenceId, lStartDate, lEndDate, lVacationPeriod, iDays, sErrorDescription)
									VerifyExistenceOfAbsences = True
								'Case "Right"
								'	Call MoveVacationsToRight(oADODBConnection, aAbsenceComponent, lAbsenceId, lStartDate, lEndDate, lVacationPeriod, iDays, sErrorDescription)
								'	VerifyExistenceOfAbsences = True
								Case "Inner", "Right"
									Call MoveVacationsToInner(oADODBConnection, aAbsenceComponent, lAbsenceID, lStartDate, lEndDate, lVacationPeriod, iDays, sErrorDescription)
									VerifyExistenceOfAbsences = True
								Case Else
									sQuery = "Delete EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_ID_EMPLOYEE) & ")" & _
											 " And (((OcurredDate >= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (OcurredDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & "))" & _
											 " Or ((EndDate >= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (EndDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & "))" & _
											 " Or ((EndDate >= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (OcurredDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & ")))"
									'Call ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									VerifyExistenceOfAbsences = True
							End Select
						ElseIf InStr(1, ",41,42,43,44,45,46,47,48,49,57,58,", "," & CStr(aAbsenceComponent(N_ABSENCE_ID_ABSENCE)) & ",", vbBinaryCompare) > 0 Then
							Call GetCrossingAbsenceType(oADODBConnection, aAbsenceComponent, sAbsenceIDs, sAbsenceCrossType, lAbsenceID, lStartDate, lEndDate, lVacationPeriod, iDays, sErrorDescription)
							Select Case sAbsenceCrossType
								Case "Left", "Right"
									VerifyExistenceOfAbsences = False
								Case "Inner"
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set EndDate=" & AddDaysToSerialDate(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), -1) & ", RemoveUserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (AbsenceID=" & lAbsenceID & ") And (OcurredDate=" & lStartDate & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
									VerifyExistenceOfAbsences = False
								Case Else
									sErrorDescription = "No se puede agregar el registro debido a que existe uno en el periodo indicado"
									VerifyExistenceOfAbsences = True
							End Select
						ElseIf InStr(1, ",50,51,52,53,54,55,56,", "," & CStr(aAbsenceComponent(N_ABSENCE_ID_ABSENCE)) & ",", vbBinaryCompare) > 0 Then
							Call GetCrossingAbsenceType(oADODBConnection, aAbsenceComponent, sAbsenceIDs, sAbsenceCrossType, lAbsenceID, lStartDate, lEndDate, lVacationPeriod, iDays, sErrorDescription)
							Select Case sAbsenceCrossType
								Case "Left", "Inner"
									VerifyExistenceOfAbsences = True
								Case Else
									sErrorDescription = "No se puede agregar el registro debido a que existe uno en el periodo indicado"
									VerifyExistenceOfAbsences = False
							End Select
						Else
							sErrorDescription1 = ""
							lOcurredDate = CLng(oRecordset.Fields("OcurredDate").Value)
							Do While Not oRecordset.EOF
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from Absences Where (AbsenceID = " & oRecordset.Fields("AbsenceID").Value & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset1)
								sErrorDescription1 = sErrorDescription1 & " " & CStr(oRecordset1.Fields("AbsenceName").Value) & ", con fecha de inicio del " &  DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value), -1, -1, -1) & " y fecha de término del " &  DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1) & ";"
								oRecordset1.Close
								oRecordset.MoveNext
							Loop
							sErrorDescription = "Para poder registrar la incidencia no debe de estar registrada ya alguna de las siguientes: " & sErrorDescription1 & " verifique que así sea"
							VerifyExistenceOfAbsences = False
						End If
					Else
						VerifyExistenceOfAbsences = True
					End If
					oRecordset.Close
				Else
					sErrorDescription = "Error al verificar si esta registrada otra incidencia."
					VerifyExistenceOfAbsences = False
				End If
			End If
		Else
			If (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 29) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 30) Then
				sQuery = "Select * From EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_ID_EMPLOYEE) & ")"
				sCondition = sCondition & " And (AbsenceID NOT IN (201, 202)) "
				If InStr(1, ",50,51,52,53,54,55,56,", "," & CStr(aAbsenceComponent(N_ABSENCE_ID_ABSENCE)) & ",", vbBinaryCompare) = 0 Then
					sCondition = sCondition & " And (AbsenceID Not In (50,51,52,53,54,55,56))"
				End If
				sQuery = sQuery & sCondition & _
						 " And (AbsenceID Not In (201, 202))" & _
						 " And (((OcurredDate >= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (OcurredDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & "))" & _
						 " Or ((EndDate >= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (EndDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & "))" & _
						 " Or ((EndDate >= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (OcurredDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & ")))"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						Call GetNameFromTable(oADODBConnection, "Absences1", oRecordset.Fields("AbsenceID").Value, "", "", sAbsenceShortName, "")
						sErrorDescription1 = sAbsenceShortName & ", con fecha de inicio del " &  DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value), -1, -1, -1) & " y fecha de término del " &  DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
						sErrorDescription = "Para poder registrar la incidencia no debe de estar registrada " & sErrorDescription1
						oRecordset.Close
						VerifyExistenceOfAbsences = False
					Else
						VerifyExistenceOfAbsences = True
					End If
				Else
					sErrorDescription = "Error al verificar si esta registrada otra incidencia."
					VerifyExistenceOfAbsences = False
				End If
			Else
				sQuery = "Select * from EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_ID_EMPLOYEE) & ")"
				If InStr(1, sAbsenceIDs, "-1", vbBinaryCompare) > 0 Then
					Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
						Case 39, 40
							sCondition = sCondition & " And (AbsenceID Not In (10, 12, 13, 14, 17, 92, 93, 94, 82, 83, 84, 85, 86, 87, 88, 89, 90, 35, 37, 38))"
					End Select
				Else
					Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
						Case 10,11,12,13,14,16,17,82,83,84,85,86,87,29,30,31,32,33,34,35,37,38
							' La siguiente condición se agrego para que estas claves no se capturaran cn H.Extras o P.Dominical pero solicitaron que no se aplicara
							'sCondition = " And (AbsenceID In (" & sAbsenceIDs & ",201,202))"
							sCondition = " And (AbsenceID In (" & sAbsenceIDs & "))"
						'Case 39, 40
						'	sCondition = sCondition & " And (AbsenceID Not In (10, 12, 13, 14, 17, 92, 93, 94, 82, 83, 84, 85, 86, 87, 88, 89, 90, 35, 37, 38))"
						Case Else
							sCondition = " And (AbsenceID In (" & sAbsenceIDs & "))"
					End Select
				End If
				If InStr(1, ",50,51,52,53,54,55,56,", "," & CStr(aAbsenceComponent(N_ABSENCE_ID_ABSENCE)) & ",", vbBinaryCompare) = 0 Then
					sCondition = sCondition & " And (AbsenceID Not In (50,51,52,53,54,55,56))"
				End If
				sCondition = sCondition & " And (AbsenceID Not In (201, 202))"
				Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
					Case 39, 40
						lStartDate1 = CLng(Left(CStr(aAbsenceComponent(N_VACATION_PERIOD_ABSENCE)), 4) & Right(CStr(aAbsenceComponent(N_VACATION_PERIOD_ABSENCE)), 2) & "01")
						Call GetLastDayFromMonth(lStartDate1, iDay)
						lEndDate = CLng(Left(CStr(aAbsenceComponent(N_VACATION_PERIOD_ABSENCE)), 4) & Right(CStr(aAbsenceComponent(N_VACATION_PERIOD_ABSENCE)), 2) & iDay)
						sCondition = sCondition & _
									 " And (((OcurredDate >= " & lStartDate1 & ") And (OcurredDate <= " & lEndDate & "))" & _
									 " Or ((EndDate >= " & lStartDate1 & ") And (EndDate <= " & lEndDate & "))" & _
									 " Or ((EndDate >= " & lStartDate1 & ") And (OcurredDate <= " & lEndDate & ")))"
					Case Else
						sCondition = sCondition & _
									 " And (((OcurredDate >= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (OcurredDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & "))" & _
									 " Or ((EndDate >= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (EndDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & "))" & _
									 " Or ((EndDate >= " &  aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (OcurredDate <= " &  aAbsenceComponent(N_END_DATE_ABSENCE) & ")))"
				End Select
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery & sCondition, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						sErrorDescription1 = ""
						lOcurredDate = CLng(oRecordset.Fields("OcurredDate").Value)
						Do While Not oRecordset.EOF
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from Absences Where (AbsenceID = " & oRecordset.Fields("AbsenceID").Value & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset1)
							sExistingAbsenceShortNames = sExistingAbsenceShortNames & CStr(oRecordset1.Fields("AbsenceShortName").Value) & ","
							If (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 4) And (CInt(oRecordset1.Fields("AbsenceID").Value) = 3) Then
								If CInt(oRecordset1.Fields("Active").Value) = 0 Then
									iActiveOriginal = -1
								Else
									iActiveOriginal = -2
								End If
							Else
								sErrorDescription1 = sErrorDescription1 & " " & CStr(oRecordset1.Fields("AbsenceName").Value) & ", con fecha de inicio del " &  DisplayDateFromSerialNumber(lOcurredDate, -1, -1, -1) & ";"
							End If
							oRecordset.MoveNext
						Loop
						oRecordset1.Close
						Set oRecordset1 = Nothing
						If (InStr(Right(sExistingAbsenceShortNames, 1), ",") > 0) Then
							sExistingAbsenceShortNames = Left(sExistingAbsenceShortNames, Len(sExistingAbsenceShortNames) -1)
						End If
						If (InStr(Right(sErrorDescription1, 1), ";") > 0) Then
							sErrorDescription1 = Left(sErrorDescription1, Len(sErrorDescription1) -1)
						End If
						sErrorDescription = "Para poder registrar la incidencia no debe de estar registrada ya alguna de las siguientes: "
						If aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 4 Then
							If (InStr(sExistingAbsenceShortNames, "0803") > 0) And (InStr(sExistingAbsenceShortNames, ",") > 0) Then
								sErrorDescription = sErrorDescription & sErrorDescription1 & " verifique que así sea"
								VerifyExistenceOfAbsences = False
							Else
								If (InStr(sExistingAbsenceShortNames, "0803") > 0) And (InStr(sExistingAbsenceShortNames, ",") = 0) Then
									lErrorNumber = JustifyAbsence(oRequest, oADODBConnection, 3, 9, iActiveOriginal, aAbsenceComponent, sErrorDescription)
									VerifyExistenceOfAbsences = True
								Else
									sErrorDescription = sErrorDescription & sErrorDescription1 & sErrorDescription1 & " verifique que así sea"
									VerifyExistenceOfAbsences = False
								End If
							End If
						Else
							sErrorDescription = sErrorDescription & sErrorDescription1 & " verifique que así sea"
							VerifyExistenceOfAbsences = False
						End If
					Else
						VerifyExistenceOfAbsences = True
					End If
				Else
					sErrorDescription = "Error al verificar si esta registrada otra incidencia."
					VerifyExistenceOfAbsences = False
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyExistenceOfAbsencesInPeriod(sAbsenceID, lPeriodDate, sErrorDescription)
'************************************************************
'Purpose: To verify if an absence already exist in database
'Inputs:  oADODBConnection, aAbsenceComponent, sAbsenceIDs, bIsForPeriod
'Outputs: iAbsenceID, iActiveStatus, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistenceOfAbsencesInPeriod"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim bComponentInitialized
	Dim sAbsenceShortName
	Dim sConditionForDates
	Dim lStartDate
	Dim lEndDate
	Dim sAbsenceIDs
	Dim iDay
	Dim iCount

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID In (" & sAbsenceID & ")) And (VacationPeriod = " & lPeriodDate & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Call GetNameFromTable(oADODBConnection, "Absences1", CInt(oRecordset.Fields("AbsenceID").Value), "", "", sAbsenceShortName, "")
			sErrorDescription = "No puede registrarse este tipo de incidencia debido a que esta registrada una " & sAbsenceShortName & " el día " & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value), -1, -1, -1) & " para el empleado " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE)
			VerifyExistenceOfAbsencesInPeriod = True
		Else
			If aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 39 Then
				lStartDate = CLng(Left(CStr(aAbsenceComponent(N_VACATION_PERIOD_ABSENCE)), 4) & Right(CStr(aAbsenceComponent(N_VACATION_PERIOD_ABSENCE)), 2) & "01")
				Call GetLastDayFromMonth(lStartDate, iDay)
				lEndDate = CLng(Left(CStr(aAbsenceComponent(N_VACATION_PERIOD_ABSENCE)), 4) & Right(CStr(aAbsenceComponent(N_VACATION_PERIOD_ABSENCE)), 2) & iDay)
				sConditionForDates = " And (((OcurredDate >= " & lStartDate & ") And (OcurredDate <= " & lEndDate & "))" & _
									 " Or ((EndDate >= " & lStartDate & ") And (EndDate <= " & lEndDate & "))" & _
									 " Or ((EndDate >= " & lStartDate & ") And (OcurredDate <= " & lEndDate & ")))"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select SUM(AbsenceHours) As Total from EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID In (29, 30))" & sConditionForDates, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						iCount = 0
						iCount = CInt(oRecordset.Fields("Total").Value)
						If (Not IsEmpty(iCount)) And (iCount > 10) Then
							Call GetNameFromTable(oADODBConnection, "Absences1", CInt(oRecordset.Fields("AbsenceID").Value), "", "", sAbsenceShortName, "")
							sErrorDescription = "No puede registrarse este tipo de incidencia debido a que existen más de 10 días de claves '0840' y '0841' registradas en el periodo del " & DisplayDateFromSerialNumber(lStartDate, -1, -1, -1) & " al " & DisplayDateFromSerialNumber(lEndDate, -1, -1, -1) & " para el empleado " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE)
							VerifyExistenceOfAbsencesInPeriod = True
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select SUM(AbsenceHours) As Total from EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID In (31))" & sConditionForDates, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									iCount = 0
									iCount = CInt(oRecordset.Fields("Total").Value)
									If (Not IsEmpty(iCount)) And (iCount > 3) Then
										Call GetNameFromTable(oADODBConnection, "Absences1", CInt(oRecordset.Fields("AbsenceID").Value), "", "", sAbsenceShortName, "")
										sErrorDescription = "No puede registrarse este tipo de incidencia debido a que existen más de 3 días de claves '0847' registradas en el periodo del " & DisplayDateFromSerialNumber(lStartDate, -1, -1, -1) & " al " & DisplayDateFromSerialNumber(lEndDate, -1, -1, -1) & " para el empleado " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE)
										VerifyExistenceOfAbsencesInPeriod = True
									Else
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select SUM(AbsenceHours) As Total from EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID In (34))" & sConditionForDates, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										If lErrorNumber = 0 Then
											If Not oRecordset.EOF Then
												iCount = 0
												iCount = CInt(oRecordset.Fields("Total").Value)
												If (Not IsEmpty(iCount)) And (iCount > 3) Then
													Call GetNameFromTable(oADODBConnection, "Absences1", CInt(oRecordset.Fields("AbsenceID").Value), "", "", sAbsenceShortName, "")
													sErrorDescription = "No puede registrarse este tipo de incidencia debido a que existen más de 3 días de claves '0855' registradas en el periodo del " & DisplayDateFromSerialNumber(lStartDate, -1, -1, -1) & " al " & DisplayDateFromSerialNumber(lEndDate, -1, -1, -1) & " para el empleado " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE)
													VerifyExistenceOfAbsencesInPeriod = True
												Else
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select SUM(AbsenceHours) As Total from EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID In (31, 34))" & sConditionForDates, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
													If lErrorNumber = 0 Then
														If Not oRecordset.EOF Then
															iCount = 0
															iCount = CInt(oRecordset.Fields("Total").Value)
															If (Not IsEmpty(iCount)) And (iCount > 3) Then
																Call GetNameFromTable(oADODBConnection, "Absences1", CInt(oRecordset.Fields("AbsenceID").Value), "", "", sAbsenceShortName, "")
																sErrorDescription = "No puede registrarse este tipo de incidencia debido a que existen más de 3 días de claves '0847' y '0855' registradas en el periodo del " & DisplayDateFromSerialNumber(lStartDate, -1, -1, -1) & " al " & DisplayDateFromSerialNumber(lEndDate, -1, -1, -1) & " para el empleado " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE)
																VerifyExistenceOfAbsencesInPeriod = True
															Else
																lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select SUM(AbsenceHours) As Total from EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID In (54))" & sConditionForDates, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
																If lErrorNumber = 0 Then
																	If Not oRecordset.EOF Then
																		iCount = 0
																		iCount = CInt(oRecordset.Fields("Total").Value)
																		If (Not IsEmpty(iCount)) And (iCount > 0) Then
																			Call GetNameFromTable(oADODBConnection, "Absences1", CInt(oRecordset.Fields("AbsenceID").Value), "", "", sAbsenceShortName, "")
																			sErrorDescription = "No puede registrarse este tipo de incidencia debido a que está registrada la clave '0905' en el periodo del " & DisplayDateFromSerialNumber(lStartDate, -1, -1, -1) & " al " & DisplayDateFromSerialNumber(lEndDate, -1, -1, -1) & " para el empleado " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE)
																			VerifyExistenceOfAbsencesInPeriod = True
																		Else
																			If VerifyLicencesDaysInDates(oADODBConnection, 10, aAbsenceComponent, sErrorDescription) Then
																				VerifyExistenceOfAbsencesInPeriod = True
																			Else
																				VerifyExistenceOfAbsencesInPeriod = False
																			End If
																		End If
																	End If
																Else
																	sErrorDescription = "Error al verificar si esta registrada una incidencia en el periodo " & lPeriodDate
																	VerifyExistenceOfAbsencesInPeriod = True
																End If
															End If
														End If
													Else
														sErrorDescription = "Error al verificar si esta registrada una incidencia en el periodo " & lPeriodDate
														VerifyExistenceOfAbsencesInPeriod = True
													End If
												End If
											End If
										Else
											sErrorDescription = "Error al verificar si esta registrada una incidencia en el periodo " & lPeriodDate
											VerifyExistenceOfAbsencesInPeriod = True
										End If
									End If
								End If
							Else
								sErrorDescription = "Error al verificar si esta registrada una incidencia en el periodo " & lPeriodDate
								VerifyExistenceOfAbsencesInPeriod = True
							End If
						End If
					End If
				Else
					sErrorDescription = "Error al verificar si esta registrada una incidencia en el periodo " & lPeriodDate
					VerifyExistenceOfAbsencesInPeriod = True
				End If
			Else
				VerifyExistenceOfAbsencesInPeriod = False
			End If
		End If
	Else
		sErrorDescription = "Error al verificar si esta registrada una incidencia en el periodo " & lPeriodDate
		VerifyExistenceOfAbsencesInPeriod = True
	End If
	oRecordset.Close

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyExistenceOfAbsencesForJustification(oADODBConnection, aAbsenceComponent, sAbsenceIDs, bIsForPeriod, iAbsenceID, iActiveStatus, sErrorDescription)
'************************************************************
'Purpose: To verify if an absence already exist in database
'Inputs:  oADODBConnection, aAbsenceComponent, sAbsenceIDs, bIsForPeriod
'Outputs: iAbsenceID, iActiveStatus, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistenceOfAbsencesForJustification"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim bComponentInitialized

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If bIsForPeriod Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_ID_EMPLOYEE) & ") And (OcurredDate = " & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (EndDate = " & aAbsenceComponent(N_END_DATE_ABSENCE) & ") And (AbsenceID IN (" & sAbsenceIDs & "))", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_ID_EMPLOYEE) & ") And (OcurredDate = " & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (AbsenceID IN (" & sAbsenceIDs & "))", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			iAbsenceID = CInt(oRecordset.Fields("AbsenceID").Value)
			iActiveStatus = CInt(oRecordset.Fields("Active").Value)
			VerifyExistenceOfAbsencesForJustification = True
		Else
			sErrorDescription = "No existen registradas incidencias para justificarlas."
			VerifyExistenceOfAbsencesForJustification = False
		End If
	Else
		sErrorDescription = "Error al verificar si esta registrada la incidencia."
		VerifyExistenceOfAbsencesForJustification = False
	End If
	oRecordset.Close

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyExistenceOfEmployeeAdditionalFeatures(oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if employee have additional perception or additional shift
'Inputs:  oADODBConnection, aAbsenceComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistenceOfEmployeeAdditionalFeatures"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim lStartDate
	Dim lEndDate

	lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
	If lErrorNumber = 0 Then
		If (aJobComponent(N_POSITION_TYPE_ID_JOB) = 1 Or aJobComponent(N_POSITION_TYPE_ID_JOB) = 2) And aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 1 Then
			If aJobComponent(N_POSITION_TYPE_ID_JOB) = 1 Then
				aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 7
			Else
				aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 8
			End If
			lErrorNumber = GetEmployeeSpecificConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				If aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) > 0 Then
					VerifyExistenceOfEmployeeAdditionalFeatures = True
				Else
					If aJobComponent(N_POSITION_TYPE_ID_JOB) = 1 Then
						sErrorDescription = "El empleado no tiene registrado el turno opcional."
						VerifyExistenceOfEmployeeAdditionalFeatures = False
					Else
						sErrorDescription = "El empleado no tiene registrada la percepción adicional."
						VerifyExistenceOfEmployeeAdditionalFeatures = False
					End If
				End If
			Else
				sErrorDescription = "No se pudo validar la existencia del turno opcional o la percepción adicional."
				VerifyExistenceOfEmployeeAdditionalFeatures = False
			End If
		Else
			sErrorDescription = "No se pudo validar la existencia del turno opcional o la percepción adicional."
			VerifyExistenceOfEmployeeAdditionalFeatures = False
		End If
	Else
		sErrorDescription = "No se pudo validar la existencia del turno opcional o la percepción adicional."
		VerifyExistenceOfEmployeeAdditionalFeatures = False
	End If
	Err.Clear
End Function

Function VerifyExistenceOfJustificationForCancel(oADODBConnection, aAbsenceComponent, iActiveStatus, sErrorDescription)
'************************************************************
'Purpose: To verify if an absence already exist in database
'Inputs:  oADODBConnection, aAbsenceComponent, sAbsenceIDs
'Outputs: iAbsenceID, iActiveStatus, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistenceOfJustificationForCancel"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim bComponentInitialized

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * from EmployeesAbsencesLKP Where (EmployeeID = " & aAbsenceComponent(N_ID_EMPLOYEE) & ") And (OcurredDate = " & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (EndDate = " & aAbsenceComponent(N_END_DATE_ABSENCE) & ") And (JustificationID<>-1)", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			iActiveStatus = CInt(oRecordset.Fields("Active").Value)
			VerifyExistenceOfJustificationForCancel = True
		Else
			sErrorDescription = "No existen registradas incidencias justificadas."
			VerifyExistenceOfJustificationForCancel = False
		End If
	Else
		sErrorDescription = "No existen registradas incidencias justificadas."
		VerifyExistenceOfJustificationForCancel = False
	End If
	oRecordset.Close

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyExistenceOfVacations(oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if an absence already exist in database
'Inputs:  oADODBConnection, aAbsenceComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistenceOfVacations"
	Dim lErrorNumber
	Dim oRecordset
	Dim sErrorDescription1
	Dim sAbsenceIDs
	Dim bComponentInitialized

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	sAbsenceIDs = "35, 37"
	If VerifyExistenceOfAbsences(oADODBConnection, aAbsenceComponent, sAbsenceIDs, sErrorDescription) Then
			sErrorDescription = "El empleado tiene registradas vacaciones."
			VerifyExistenceOfVacations = True
	Else
		VerifyExistenceOfVacations = False
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyFortnightlyExistenceOfAbsences(oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if employee absences exist with diference minimum of one month
'Inputs:  oADODBConnection, aAbsenceComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyFortnightlyExistenceOfAbsences"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim lStartDate
	Dim lEndDate

	lStartDate = GetPayrollStartDate(aAbsenceComponent(N_APPLIED_DATE_ABSENCE))
	lEndDate = aAbsenceComponent(N_APPLIED_DATE_ABSENCE)

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAbsencesLKP Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ") And (OcurredDate>=" & lStartDate & ") And (OcurredDate<=" & lEndDate &  ") Order By OcurredDate Desc", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			VerifyFortnightlyExistenceOfAbsences = False
		Else
			sErrorDescription = "La incidencia no cubre el plazo de una quincena de haberse registrado"
			VerifyFortnightlyExistenceOfAbsences = True
		End If
		oRecordset.Close
	Else
		sErrorDescription = "Error al verificar si la incidencia ya fué registrada en esta quincena"
		VerifyFortnightlyExistenceOfAbsences = True
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyLicencesDaysInDates(oADODBConnection, iLicencesDays, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if employee absences exist with diference minimum of one month
'Inputs:  oADODBConnection, aAbsenceComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyLicencesDaysInDates"
	Dim lErrorNumber
	Dim oRecordset
	Dim dStartDate
	Dim dEndDate
	Dim lDaysCount
	Dim sDatesCondition

	sDatesCondition = " And (((EmployeeDate >= " & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (EmployeeDate <= " & aAbsenceComponent(N_END_DATE_ABSENCE) & "))" & _
					  " Or ((EndDate >= " & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (EndDate <= " & aAbsenceComponent(N_END_DATE_ABSENCE) & "))" & _
					  " Or ((EndDate >= " & aAbsenceComponent(N_OCURRED_DATE_ABSENCE) & ") And (EmployeeDate <= " & aAbsenceComponent(N_END_DATE_ABSENCE) & ")))"

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesHistoryList Where (EmployeeID = " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (StatusID=150)" & sDatesCondition, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			lDaysCount = 0
			Do While Not oRecordset.EOF
				dStartDate = GetDateFromSerialNumber(oRecordset.Fields("EmployeeDate").value)
				If CLng(oRecordset.Fields("EndDate").value) = 30000000 Then
					lErrorNumber = -1
					dEndDate = GetDateFromSerialNumber(CLng(20200000))
				Else
					dEndDate = GetDateFromSerialNumber(oRecordset.Fields("EndDate").value)
				End If
				lDaysCount = lDaysCount + DateDiff("d", dStartDate, dEndDate)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Or (lDaysCount>iLicencesDays) Then Exit Do
			Loop
		End If
		oRecordset.Close
	End If
	If (lDaysCount > iLicencesDays) Then
		sErrorDescription = "El empleado " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & " tiene registradas licencias con goce de sueldo por más de 10 días"
		VerifyLicencesDaysInDates = True
	Else
		VerifyLicencesDaysInDates = False
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyMensualExistenceOfAbsences(oADODBConnection, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if employee absences exist with diference minimum of one month
'Inputs:  oADODBConnection, aAbsenceComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyMensualExistenceOfAbsences"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim lStartDate
	Dim lEndDate
	Dim iDay

	lStartDate = CLng(Left(CStr(aAbsenceComponent(N_OCURRED_DATE_ABSENCE)), Len("000000"))& "01")
	Call GetLastDayFromMonth(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), iDay)
	lEndDate = CLng(Left(CStr(aAbsenceComponent(N_OCURRED_DATE_ABSENCE)), Len("000000"))& iDay)

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAbsencesLKP Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ") And (OcurredDate>=" & lStartDate & ") And (OcurredDate<=" & lEndDate &  ") And (Active = 1) Order By OcurredDate Desc", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			VerifyMensualExistenceOfAbsences = False
		Else
			sErrorDescription = "La incidencia no cubre el plazo de un mes de haberse registrado"
			VerifyMensualExistenceOfAbsences = True
		End If
	Else
		sErrorDescription = "Error al verificar si la incidencia cubre el plazo de un mes de haberse registrado"
		VerifyMensualExistenceOfAbsences = True
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyMinimumTimeInJob(oADODBConnection, iJobID, aAbsenceComponent, sErrorDescription)
'************************************************************
'Purpose: To verify if employee absences exist with diference minimum of one month
'Inputs:  oADODBConnection, aAbsenceComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyMinimumTimeInJob"
	Dim lErrorNumber
	Dim oRecordset
	Dim dStartDate
	Dim dEndDate
	Dim lDaysCount

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesHistoryList Where (EmployeeID = " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (PositionTypeID=1)", "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			lDaysCount = 0
			Do While Not oRecordset.EOF
				dStartDate = GetDateFromSerialNumber(oRecordset.Fields("EmployeeDate").value)
				If CLng(oRecordset.Fields("EndDate").value) = 30000000 Then
					dEndDate = GetDateFromSerialNumber(CLng(Left(GetSerialNumberForDate(""), Len("00000000"))))
				Else
					dEndDate = GetDateFromSerialNumber(oRecordset.Fields("EndDate").value)
				End If
				lDaysCount = lDaysCount + DateDiff("d", dStartDate, dEndDate)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Or (lDaysCount>=181) Then Exit Do
			Loop
		End If
		oRecordset.Close
	End If
	If (lDaysCount < 181) Then
		sErrorDescription = "El empleado señalado no cumple con el mínimo de 6 meses en plaza de base"
		VerifyMinimumTimeInJob = False
	Else
		VerifyMinimumTimeInJob = True
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyRequerimentsForEmployeesAbsences(oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To verify employee status requirements to register absences
'Inputs:  oADODBConnection, lReasonID, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyRequerimentsForEmployeesAbsences"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim iJobID
	Dim iEmployeeTypeID
	Dim iStatusEmployeeID
	Dim sStatusEmployee
	Dim iPositionTypeID
	Dim iShiftID
	Dim sShiftName
	Dim iServiceID
	Dim sServiceShortName
	Dim iJourneyTypeID
	Dim bComponentInitialized
	Dim iDay
	Dim sAbsenceID
	Dim lPerioDate
	Dim sAbsenceShortName

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If
	VerifyRequerimentsForEmployeesAbsences = True
	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Then
		sErrorDescription = "No se especificó el identificador del empleado para agregar incidencias."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
		VerifyRequerimentsForEmployeesAbsences = False
	Else
		aEmployeeComponent(N_ID_EMPLOYEE) = aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE)
		lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			If VerifyUserPermissionOnEmployee(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
				sQuery = "Select * From Employees Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
				sErrorDescription = "No se pudieron obtener los datos del empleado para validar los requisitos de la incidencia."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						iJobID = CLng(oRecordset.Fields("JobID").Value)
						iEmployeeTypeID = CInt(oRecordset.Fields("EmployeeTypeID").Value)
						iStatusEmployeeID = CInt(oRecordset.Fields("StatusID").Value)
						iPositionTypeID = CInt(oRecordset.Fields("PositionTypeID").Value)
						iShiftID = CInt(oRecordset.Fields("ShiftID").Value)
						iServiceID = CInt(oRecordset.Fields("ServiceID").Value)
						oRecordset.Close
						sQuery = "Select * From Shifts Where (ShiftID=" & iShiftID & ")"
						sErrorDescription = "No se pudo obtener la jornada del empleado para validar los requisitos de la incidencia."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						sShiftName = CStr(oRecordset.Fields("ShiftName").Value)
						If Len(sShiftName) > 0 Then
							sShiftName = LCase(sShiftName)
						Else
							sShiftName = "indeterminada"
						End If
						iJourneyTypeID = CInt(oRecordset.Fields("JourneyTypeID").Value)
						oRecordset.Close
						sQuery = "Select * From Services Where (ServiceID=" & iServiceID & ") And (StartDate<=" & CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) & ") And (EndDate>=" & CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) & ")"
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						sErrorDescription = "No se pudo obtener el servicio del empleado para validar los requisitos de la incidencia."
						sServiceShortName = CStr(oRecordset.Fields("ServiceShortName").Value)
						If Len(sServiceShortName) > 0 Then
							sShiftName = UCase(sShiftName)
						Else
							sShiftName = "indeterminada"
						End If
						oRecordset.Close
						If (InStr(1, ",29,30,31,35,37,38,", "," & CStr(aAbsenceComponent(N_ABSENCE_ID_ABSENCE)) & ",", vbBinaryCompare) > 0) And (iStatusEmployeeID <> 0) Then
							Call GetNameFromTable(oADODBConnection, "StatusEmployees", iStatusEmployeeID, "", "", sStatusEmployee, sErrorDescription)
							sErrorDescription = "Solamente se puede registrar este tipo de incidencias al personal con estatus activo. El status actual del empleado es " & sStatusEmployee
							VerifyRequerimentsForEmployeesAbsences = False
						ElseIf (iStatusEmployeeID <> 0) And (iStatusEmployeeID <> 1) Then
							Select Case iStatusEmployeeID
								Case 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 27, 28, 29, 31, 32, 33, 35, 36, 37, 39, 40, 41, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 119, 120, 121, 123, 124, 125, 126, 127, 128, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 145, 146, 147, 149, 150, 151, 152, 153, 154, 155, 156, 157, 158
								Case Else
									lErrorNumber = -1
							End Select
						End If
						If lErrorNumber <> 0 Then
							sErrorDescription = "Solamente se pueden registrar este tipo de incidencias al personal con estatus activo"
							VerifyRequerimentsForEmployeesAbsences = False
						Else
							If iEmployeeTypeID = 7 Then
								sErrorDescription = "No se pueden registrar incidencias a los empleados por honorarios"
								VerifyRequerimentsForEmployeesAbsences = False
							Else
								If (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) <> -1) And (CInt(Request.Cookies("SIAP_SectionID")) = 7) And (Not VerifyAbsenceOccurredDateLimit(oADODBConnection, aAbsenceComponent, sErrorDescription)) Then
									Call GetNameFromTable(oADODBConnection, "Absences1", aAbsenceComponent(N_ABSENCE_ID_ABSENCE), "", "", sAbsenceShortName, "")
									sErrorDescription = "No se puede registrar " & sAbsenceShortName & " con fecha de inicio " & DisplayDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), -1, -1, -1) & " debido a que en desconcentrados solamente puede registrar incidencias con fecha de ocurrencia de hasta dos quincenas atras"
									VerifyRequerimentsForEmployeesAbsences = False
								Else
									Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
										Case 1, 2, 3, 4, 7, 8, 9, 10, 12, 15, 16, 18, 19, 21, 22, 91, 23, 24, 27, 28, 31, 32, 35, 37, 79 ' Aplica para todo tipo de personal y que sean los 4 tipos de jornada
											If (iJourneyTypeID = 1) Or (iJourneyTypeID = 2) Or (iJourneyTypeID = 3) Or (iJourneyTypeID = 4) Then
												If aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 12 Then
													If VerifyIfEmployeeIsInLocalZone(oADODBConnection, aEmployeeComponent, aAbsenceComponent(N_OCURRED_DATE_ABSENCE), False, sErrorDescription) Then
														VerifyRequerimentsForEmployeesAbsences = True
													Else
														VerifyRequerimentsForEmployeesAbsences = False
													End If
												ElseIf (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 35) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 37) Then
													If VerifyRequerimentsForEmployeesVacations(oADODBConnection, aAbsenceComponent, sServiceShortName, iJourneyTypeID, sErrorDescription) Then
														VerifyRequerimentsForEmployeesAbsences = True
													Else
														VerifyRequerimentsForEmployeesAbsences = False
													End If
												ElseIf aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 28 Then
													If VerifyExistenceOfEmployeeAdditionalFeatures(oADODBConnection, aAbsenceComponent, sErrorDescription) Then
														VerifyRequerimentsForEmployeesAbsences = True
													Else
														VerifyRequerimentsForEmployeesAbsences = False
													End If
												'ElseIf aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 31 Then
												'	If VerifyAnualDiferenceOfAbsences(oADODBConnection, aAbsenceComponent, False, sErrorDescription) Then
												'		VerifyRequerimentsForEmployeesAbsences = True
												'	Else
												'		VerifyRequerimentsForEmployeesAbsences = False
												'	End If
												ElseIf aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 32 Then
													If DateDiff("d", GetDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE)), GetDateFromSerialNumber(aAbsenceComponent(N_END_DATE_ABSENCE))) + 1 > 90 Then
														sErrorDescription = "El periodo de la incapacidad por gravidez no debe de ser mayor a 90 días calendario"
														VerifyRequerimentsForEmployeesAbsences = False
													Else
														VerifyRequerimentsForEmployeesAbsences = True
													End If
												ElseIf aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 79 Then
													If Not VerifyFortnightlyExistenceOfAbsences(oADODBConnection, aAbsenceComponent, sErrorDescription) Then
														Call GetLastDayFromMonth(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), iDay)
														If (CInt(Right(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), Len("00"))) = 15) Or (CInt(Right(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), Len("00"))) = iDay) Then
															VerifyRequerimentsForEmployeesAbsences = True
														Else
															sErrorDescription = "Esta incidencia solamente se puede registrar los días 15 o últimos de cada mes"
															VerifyRequerimentsForEmployeesAbsences = False
														End If
													Else
														VerifyRequerimentsForEmployeesAbsences = False
													End If
												Else
													VerifyRequerimentsForEmployeesAbsences = True
												End If
											Else
												sErrorDescription = "No se puede registrar el tipo de incidencia a los empleados con la jornada " & sShiftName
												VerifyRequerimentsForEmployeesAbsences = False
											End If
										Case 5, 11, 92 ' Caso en el que sea todo tipo de personal y que jornada sea igual a jornada ordinaria 1
											If iJourneyTypeID = 1 Then
												If (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 5 Or aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 11) And (VerifyExistenceOfEmployeeAdditionalFeatures(oADODBConnection, aAbsenceComponent, sErrorDescription)) Then
													VerifyRequerimentsForEmployeesAbsences = True
												Else
													VerifyRequerimentsForEmployeesAbsences = False
												End If
											Else
												sErrorDescription = "No se puede registrar el tipo de incidencia a los empleados con la jornada " & sShiftName
												VerifyRequerimentsForEmployeesAbsences = False
											End If
										Case 20, 25, 26
											If VerifyExistenceOfEmployeeAdditionalFeatures(oADODBConnection, aAbsenceComponent, sErrorDescription) Then
												VerifyRequerimentsForEmployeesAbsences = True
											Else
												VerifyRequerimentsForEmployeesAbsences = False
											End If
										Case 13 ' Caso en el que sea personal de base y que sean los 3 tipos de jornada
											If iPositionTypeID <> 1 Then
												sErrorDescription = "La incidencia es válido sólo para el personal de base"
												VerifyRequerimentsForEmployeesAbsences = False
											Else
												If (iJourneyTypeID = 1) Or (iJourneyTypeID = 2) Or (iJourneyTypeID = 3) Or (iJourneyTypeID = 4) Then
													VerifyRequerimentsForEmployeesAbsences = True
												Else
													sErrorDescription = "No se puede registrar el tipo de incidencia a los empleados con la jornada " & sShiftName
													VerifyRequerimentsForEmployeesAbsences = False
												End If
											End If
										Case 17, 39, 40
											If iPositionTypeID <> 1 Then
												sErrorDescription = "La incidencia es válido sólo para el personal de base"
												VerifyRequerimentsForEmployeesAbsences = False
											Else
												If (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 39) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 40) Then
													Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
														Case 39
															sAbsenceID = 40
														Case 40
															sAbsenceID = 39
													End Select
													Call GetLastDayFromMonth(aAbsenceComponent(N_VACATION_PERIOD_ABSENCE) & "01", iDay)
													If CLng(Left(CStr(aAbsenceComponent(N_VACATION_PERIOD_ABSENCE)), 4) & Right(CStr(aAbsenceComponent(N_VACATION_PERIOD_ABSENCE)), 2) & iDay) > aAbsenceComponent(N_OCURRED_DATE_ABSENCE) Then
														sErrorDescription = "El periodo capturado no puede ser mayor que la fecha de ocurrencias de la incidencia"
														VerifyRequerimentsForEmployeesAbsences = False
													Else
														If CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
															If VerifyExistenceOfAbsencesInPeriod(sAbsenceID, aAbsenceComponent(N_VACATION_PERIOD_ABSENCE), sErrorDescription) Then
																VerifyRequerimentsForEmployeesAbsences = False
															Else
																If Not VerifyMensualExistenceOfAbsences(oADODBConnection, aAbsenceComponent, sErrorDescription) Then
																	Call GetLastDayFromMonth(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), iDay)
																	If Right(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), Len("00")) = Right(iDay, Len("00")) Then
																		VerifyRequerimentsForEmployeesAbsences = True
																	Else
																		sErrorDescription = "Esta incidencia solamente se puede registrar el último día del mes"
																		VerifyRequerimentsForEmployeesAbsences = False
																	End If
																Else
																	VerifyRequerimentsForEmployeesAbsences = False
																End If
															End If
														Else
															If DateDiff("m", GetDateFromSerialNumber(CLng(Left(CStr(aAbsenceComponent(N_VACATION_PERIOD_ABSENCE)), 4) & Right(CStr(aAbsenceComponent(N_VACATION_PERIOD_ABSENCE)), 2) & iDay)), GetDateFromSerialNumber(aAbsenceComponent(N_APPLIED_DATE_ABSENCE))) <= 1 Then
																If VerifyExistenceOfAbsencesInPeriod(sAbsenceID, aAbsenceComponent(N_VACATION_PERIOD_ABSENCE), sErrorDescription) Then
																	VerifyRequerimentsForEmployeesAbsences = False
																Else
																	If Not VerifyMensualExistenceOfAbsences(oADODBConnection, aAbsenceComponent, sErrorDescription) Then
																		Call GetLastDayFromMonth(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), iDay)
																		If Right(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), Len("00")) = Right(iDay, Len("00")) Then
																			VerifyRequerimentsForEmployeesAbsences = True
																		Else
																			sErrorDescription = "Esta incidencia solamente se puede registrar el último día del mes"
																			VerifyRequerimentsForEmployeesAbsences = False
																		End If
																	Else
																		VerifyRequerimentsForEmployeesAbsences = False
																	End If
																End If
															Else
																sErrorDescription = "El periodo capturado con respecto a la quincena de aplicación tiene más de dos quincenas de diferencia"
																VerifyRequerimentsForEmployeesAbsences = False
															End If
														End If
													End If
												End If
											End If
										Case 29, 30
											Dim lAntiquityYears, lAntiquityMonths, lAntiquityDays, lEmployeDateForAntiquity
											Dim iDaysForAbsence, iTotalDaysForAbsence
											Dim sEmployeeAntiquity
											If (iPositionTypeID <> 1) And (iPositionTypeID <> 5) Then ' PROVISIONAL
												sErrorDescription = "La incidencia es válido sólo para el personal de base o residentes"
												VerifyRequerimentsForEmployeesAbsences = False
											Else
												If VerifyMinimumTimeInJob(oADODBConnection, iJobID, aAbsenceComponent, sErrorDescription) Then
													If CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) <= CLng(Year(Date()) & "0531") Then
														lEmployeDateForAntiquity = CLng(Year(Date()) & "0531")
													Else
														lEmployeDateForAntiquity = CLng(Year(Date()) & "1231")
													End If
													lErrorNumber = CalculateEmployeeAntiquity(oADODBConnection, aEmployeeComponent, lEmployeDateForAntiquity, sEmployeeAntiquity, lAntiquityYears, lAntiquityMonths, lAntiquityDays, sErrorDescription)
													If (lAntiquityYears >= 0) And (lAntiquityYears <= 5) Then
														iTotalDaysForAbsence = 21
													ElseIf (lAntiquityYears > 5) And (lAntiquityYears <= 10) Then
														iTotalDaysForAbsence = 26
													ElseIf (lAntiquityYears > 10) And (lAntiquityYears <= 15) Then
														iTotalDaysForAbsence = 31
													ElseIf (lAntiquityYears > 15) And (lAntiquityYears <= 20) Then
														iTotalDaysForAbsence = 36
													ElseIf (lAntiquityYears > 20) Then
														iTotalDaysForAbsence = 41
													End If
													If VerifyRequerimentsForEmployeesLicences(oADODBConnection, aAbsenceComponent, iTotalDaysForAbsence, sErrorDescription) Then
														VerifyRequerimentsForEmployeesAbsences = True
													Else
														VerifyRequerimentsForEmployeesAbsences = False
													End If
												Else
													sErrorDescription = "El empleado no tiene el tiempo minimo de ocupación de la plaza"
													VerifyRequerimentsForEmployeesAbsences = False
												End If
											End If
										Case 38 ' Caso en el que sea personal de base y confianza "B" y que sean los 3 tipos de jornada
											If (iPositionTypeID <> 1) And (iPositionTypeID <> 2) Then
												sErrorDescription = "La incidencia es válido sólo para el personal de base o confianza"
												VerifyRequerimentsForEmployeesAbsences = False
											Else
												If (iJourneyTypeID = 1) Or (iJourneyTypeID = 2) Or (iJourneyTypeID = 3) Or (iJourneyTypeID = 4) Then
													If aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 31 Then
														If VerifyAnualDiferenceOfAbsences(oADODBConnection, aAbsenceComponent, False, sErrorDescription) Then
															VerifyRequerimentsForEmployeesAbsences = True
														Else
															VerifyRequerimentsForEmployeesAbsences = False
														End If
													ElseIf aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 38 Then
														If VerifyRequerimentsForEmployeesVacations(oADODBConnection, aAbsenceComponent, sErrorDescription) Then
															VerifyRequerimentsForEmployeesAbsences = True
														Else
															VerifyRequerimentsForEmployeesAbsences = False
														End If
													End If
												Else
													sErrorDescription = "No se puede registrar el tipo de incidencia a los empleados con la jornada " & sShiftName
													VerifyRequerimentsForEmployeesAbsences = False
												End If
											End If
										Case 40, 14 ' Caso en el que sea personal de base y confianza "B" sin importar el tipo de jornada
											If (iPositionTypeID <> 1) And (iPositionTypeID <> 2) Then
												sErrorDescription = "La incidencia es válido sólo para el personal de base o confianza"
												VerifyRequerimentsForEmployeesAbsences = False
											Else
												If aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = 40 Then
													If Not VerifyMensualExistenceOfAbsences(oADODBConnection, aAbsenceComponent, sErrorDescription) Then
														Call GetLastDayFromMonth(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), iDay)
														If Right(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), Len("00")) = Right(iDay, Len("00")) Then
															VerifyRequerimentsForEmployeesAbsences = True
														Else
															sErrorDescription = "Esta incidencia solamente se puede registrar el último día del mes"
															VerifyRequerimentsForEmployeesAbsences = False
														End If
													Else
														VerifyRequerimentsForEmployeesAbsences = False
													End If
												Else
													VerifyRequerimentsForEmployeesAbsences = True
												End If
											End If
										Case 33, 36
											VerifyRequerimentsForEmployeesAbsences = True
										Case Else
											VerifyRequerimentsForEmployeesAbsences = True
									End Select
								End If
							End If
						End If
						oRecordset.Close
					Else
						sErrorDescription = "Error al verificar el status del empleado para registrar la incidencia"
						VerifyRequerimentsForEmployeesAbsences = False
					End If
				Else
					sErrorDescription = "Error al verificar el status del empleado para registrar la incidencia"
					VerifyRequerimentsForEmployeesAbsences = False
				End If
			Else
				VerifyRequerimentsForEmployeesAbsences = False
			End If
		Else
			VerifyRequerimentsForEmployeesAbsences = False
		End If
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyRequerimentsForEmployeesLicences(oADODBConnection, aAbsenceComponent, iDaysPerPeriod, sErrorDescription)
'************************************************************
'Purpose: To verify employee status requirements to register absences
'Inputs:  oADODBConnection, aAbsenceComponent, iDaysPerPeriod
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyRequerimentsForEmployeesLicences"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim bComponentInitialized
	Dim iTotalDays
	Dim lStartDate
	Dim sPeriodCondition

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Then
		sErrorDescription = "No se especificó el identificador del empleado para agregar incidencias."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
		VerifyRequerimentsForEmployeesLicences = False
	Else
		If iDaysPerPeriod >= 0 Then
			sPeriodCondition = "(OcurredDate>=" & CLng(Year(Date()) & "0101") & ") And (EndDate<=" & CLng(Year(Date()) & "1231") & ")"
			sQuery = "Select SUM(AbsenceHours) As TotalDays From EmployeesAbsencesLKP Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ") And " & sPeriodCondition
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					If IsNull(oRecordset.Fields("TotalDays").Value) Then
						iTotalDays = 0
					Else
						iTotalDays = CInt(oRecordset.Fields("TotalDays").Value)
					End If
					If (aAbsenceComponent(N_HOURS_ABSENCE) + iTotalDays > iDaysPerPeriod) Then
						If (iDaysPerPeriod - iTotalDays) = 0 Then
							sErrorDescription = "La cantidad de permisos/licencias con goce de sueldo capturadas en el periodo con fecha de inicio " & CStr(GetDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE))) & ", exceden el número de jornadas disponibles para el empleado, según su antigüedad: Ninguna jornada disponible"
						Else
							sErrorDescription = "La cantidad de permisos/licencias con goce de sueldo capturadas en el periodo con fecha de inicio " & CStr(GetDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE))) & ", exceden el número de jornadas disponibles para el empleado, según su antigüedad: " & iDaysPerPeriod - iTotalDays & " jornadas disponibles"
						End If
						VerifyRequerimentsForEmployeesLicences = False
					Else
						VerifyRequerimentsForEmployeesLicences = True
					End If
					oRecordset.Close
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function VerifyRequerimentsForEmployeesVacations(oADODBConnection, aAbsenceComponent, sServiceShortName, iJourneyTypeID, sErrorDescription)
'************************************************************
'Purpose: To verify employee status requirements to register absences
'Inputs:  oADODBConnection, lReasonID, aAbsenceComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyRequerimentsForEmployeesVacations"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim bComponentInitialized
	Dim iDaysPerPeriod
	Dim iTotalDays

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Then
		sErrorDescription = "No se especificó el identificador del empleado para agregar incidencias."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
		VerifyRequerimentsForEmployeesVacations = False
	Else
		iDaysPerPeriod = -1
		Select Case aAbsenceComponent(N_ABSENCE_ID_ABSENCE)
			Case 35 ' Clave 60: Vacaciones
				Select Case sServiceShortName
					Case "09200", "09210", "20600", "17824", "70910", "11150"
						sErrorDescription = "No se pueden registrar vacaciones a empleados que tengan las siguientes claves de servicio: '09200', '09210', '20600', '17824', '70910', '11150'. A estos les debe de registrar vacaciones por emanaciones radiactivas."
						VerifyRequerimentsForEmployeesVacations = False
					Case Else
						Select Case iJourneyTypeID
							Case 1
								iDaysPerPeriod = 10
							Case 2, 3
								iDaysPerPeriod = 5
							Case 4
								iDaysPerPeriod = 2
							Case Else
								iDaysPerPeriod = 0
						End Select
				End Select
			Case 37 ' Clave 62: Vacaciones por emanaciones radiactivas
				Select Case sServiceShortName
					Case "09200", "09210", "20600", "17824", "70910", "11150"
						Select Case iJourneyTypeID
							Case 1
								iDaysPerPeriod = 20
							Case 2, 3
								iDaysPerPeriod = 10
							Case 4
								iDaysPerPeriod = 4
							Case Else
								iDaysPerPeriod = 0
						End Select
					Case Else
						sErrorDescription = "Solamente se pueden registrar vacaciones por emanaciones radiactivas a empleados que tengan las siguientes claves de servicio: '09200', '09210', '20600', '17824', '70910', '11150'. A los empleados con claves de servicio distintas les debe de registrar vacaciones normales."
						VerifyRequerimentsForEmployeesVacations = False
				End Select
			Case 38 ' Clave 63: Vacaciones extraordinarias por premios, estimulos y recompensas
				Select Case iJourneyTypeID
					Case 1
						iDaysPerPeriod = 10
					Case 2, 3
						iDaysPerPeriod = 5
					Case 4
						iDaysPerPeriod = 2
					Case Else
						iDaysPerPeriod = 0
				End Select
			Case Else
				VerifyRequerimentsForEmployeesVacations = False
		End Select
		If iDaysPerPeriod >= 0 Then
			If CInt(aAbsenceComponent(N_VACATION_PERIOD_ABSENCE)) < 20120 Then
				If (aAbsenceComponent(N_HOURS_ABSENCE) > iDaysPerPeriod) Then
					sErrorDescription = "La cantidad de vacaciones capturadas en el periodo, con fecha de inicio " & CStr(GetDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE))) & ", exceden el número de jornadas disponibles para el empleado: " & iDaysPerPeriod & " jornadas disponibles"
					VerifyRequerimentsForEmployeesVacations = False
				Else
					VerifyRequerimentsForEmployeesVacations = True
				End If
			Else
				sQuery = "Select SUM(AbsenceHours) As TotalDays From EmployeesAbsencesLKP Where (EmployeeID=" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & aAbsenceComponent(N_ABSENCE_ID_ABSENCE) & ") And (VacationPeriod=" & aAbsenceComponent(N_VACATION_PERIOD_ABSENCE) & ")"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						If IsNull(oRecordset.Fields("TotalDays").Value) Then
							iTotalDays = 0
						Else
							iTotalDays = CInt(oRecordset.Fields("TotalDays").Value)
						End If
						oRecordset.Close
						If (aAbsenceComponent(N_HOURS_ABSENCE) + iTotalDays > iDaysPerPeriod) Then
							If (iDaysPerPeriod - iTotalDays) = 0 Then
								sErrorDescription = "La cantidad de vacaciones capturadas, en el periodo con fecha de inicio " & CStr(GetDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE))) & ", exceden el número de jornadas disponibles para el empleado: Ninguna jornada disponible"
							Else
								sErrorDescription = "La cantidad de vacaciones capturadas, en el periodo con fecha de inicio " & CStr(GetDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE))) & ", exceden el número de jornadas disponibles para el empleado: " & iDaysPerPeriod - iTotalDays & " jornadas disponibles"
							End If
							VerifyRequerimentsForEmployeesVacations = False
						Else
							VerifyRequerimentsForEmployeesVacations = True
						End If
					End If
				End If
			End If
		'Else
		'	VerifyRequerimentsForEmployeesVacations = False
		End If
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function MoveVacationsToInner(oADODBConnection, aAbsenceComponent, lAbsenceId, lStartDate, lEndDate, lVacationPeriod, iDays, sErrorDescription)
'************************************************************
'Purpose: To register medical licence into vacations period
'         append registry of them
'Inputs:  oADODBConnection, aAbsenceComponent, lStartDate, lEndDate
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "MoveVacationsToInner"
	Dim lErrorNumber
	Dim oRecordset
	Dim bComponentInitialized
	Dim lDaysFromMedicalLicence
	Dim oDate
	Dim sQuery
	Dim iJourneyTypeID
	Dim updateVacationsDays   ' Las vacaciones se parten en dos, la primera parte se trunca y se actualizan días de vacaciones
	Dim lNextStartDateForVacations ' Siguiente día habíl para regstrar las vacaciones
	Dim lNextEndDateForVacations   ' Siguiente día habíl para regstrar el fin de las vacaciones que se truncan
	Dim lFinalDateForVacations   ' Día de fin de las vacaciones
	Dim iNextDaysForVacations ' Total de días que cuentan como vacaciones
	Dim iDaysAddForVacations  ' Días que se cuentan en caso de recorrerse por ser Sabado o Domingo

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Or (aAbsenceComponent(N_END_DATE_ABSENCE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador de la incidencia y/o la fecha para agregar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	End If

	iDaysAddForVacations = 0
	updateVacationsDays = 0
	lDaysFromMedicalLicence = DateDiff("d", GetDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE)), GetDateFromSerialNumber(aAbsenceComponent(N_END_DATE_ABSENCE))) + 1
	lNextStartDateForVacations = GetNextStartDateForVacations(AddDaysToSerialDate(aAbsenceComponent(N_END_DATE_ABSENCE), 1), 1, iDaysAddForVacations)

	Select Case lAbsenceId
		Case 35, 37, 38
			Call GetEmployeeJourneyType(oRequest, oADODBConnection, aEmployeeComponent, iJourneyTypeID, sErrorDescription)
			If lStartDate <> aAbsenceComponent(N_OCURRED_DATE_ABSENCE) Then
				' Para truncarse las vacaciones se tiene como límite la fecha de inicio de la lic. médica menos 1 día 
				lNextEndDateForVacations = GetNextEndDateForVacations(AddDaysToSerialDate(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), -1), iJourneyTypeID)
				' Sabiendo la fecha de inicio y la fecha de fin se deducen los días habiles para el primer periódo
				updateVacationsDays = GetWorkingDaysOfAbsencesPeriod(lStartDate, lNextEndDateForVacations, iJourneyTypeID)
			End If
			lFinalDateForVacations = AddDaysToSerialDateForVacations(oADODBConnection, lNextStartDateForVacations, iDays - updateVacationsDays -1, iJourneyTypeID)
			' También hay que saber los días habiles para la segunda parte de las vacaciones
			iNextDaysForVacations = GetWorkingDaysOfAbsencesPeriod(lNextStartDateForVacations, lFinalDateForVacations, iJourneyTypeID)
		Case Else
			updateVacationsDays = DateDiff("d", GetDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE)), GetDateFromSerialNumber(aAbsenceComponent(N_END_DATE_ABSENCE))) + 1
			aAbsenceComponent(N_HOURS_ABSENCE) = DateDiff("d", GetDateFromSerialNumber(AddDaysToSerialDate(oADODBConnection, aAbsenceComponent(N_END_DATE_ABSENCE), 1)), GetDateFromSerialNumber(AddDaysToSerialDate(oADODBConnection, lEndDate, lDaysFromMedicalLicence))) + 1
	End Select

	If lStartDate <> aAbsenceComponent(N_OCURRED_DATE_ABSENCE) Then
		sQuery = "Update EmployeesAbsencesLKP set EndDate = " & AddDaysToSerialDate(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), -1) & _
				 " , AbsenceHours = " & updateVacationsDays & _
				 " where EmployeeId = " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & _
				 " and AbsenceID = " & lAbsenceId & _
				 " and OcurredDate = " & lStartDate
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Else
		sQuery = "Delete from EmployeesAbsencesLKP" & _
				 " where EmployeeId = " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & _
				 " and AbsenceID = " & lAbsenceId & _
				 " and OcurredDate = " & lStartDate
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If

	If lErrorNumber = 0 Then
		sQuery = "Insert Into EmployeesAbsencesLKP (EmployeeID, AbsenceID, OcurredDate, EndDate, RegistrationDate, " & _
				 " DocumentNumber, AbsenceHours, JustificationID, AppliesForPunctuality, Reasons, AddUserID, AppliedDate, " & _
				 " Removed, RemoveUserID, RemovedDate, AppliedRemoveDate, Active, VacationPeriod) Values (" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ", " & _
				 lAbsenceId & ", " & lNextStartDateForVacations & ", " & _
				 lFinalDateForVacations & ", " & aAbsenceComponent(N_REGISTRATION_DATE_ABSENCE) & ", '" & _
				 Replace(aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE), "'", "´") & "', " & iNextDaysForVacations & ", " & _
				 aAbsenceComponent(N_JUSTIFICATION_ID_ABSENCE) & ", " & aAbsenceComponent(N_APPLIES_FOR_PUNCTUALITY_ABSENCE) & ", '" & _
				 Replace(aAbsenceComponent(S_REASONS_ABSENCE), "'", "´") & "', " & aAbsenceComponent(N_ADD_USER_ID_ABSENCE) & ", " & _
				 aAbsenceComponent(N_APPLIED_DATE_ABSENCE) & ", " & aAbsenceComponent(N_REMOVED_ABSENCE) & ", " & _
				 aAbsenceComponent(N_REMOVE_USER_ID_ABSENCE) & ", " &  aAbsenceComponent(N_REMOVED_DATE_ABSENCE) & ", " & _
				 aAbsenceComponent(N_APPLIED_REMOVE_DATE_ABSENCE) & ", " & aAbsenceComponent(N_ACTIVE_ABSENCE) & ", " & _
				 lVacationPeriod & ")"

		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber <> 0 Then
			sErrorDescription = "No se pudieron modificar las vacaciones del empleado."
		End If
	Else
		sErrorDescription = "No se pudieron modificar las vacaciones del empleado."
	End If
	Set oRecordset = Nothing
	MoveVacationsToInner = lErrorNumber
	Err.Clear
End Function

Function MoveVacationsToLeft(oADODBConnection, aAbsenceComponent, lAbsenceId, lStartDate, lEndDate, lVacationPeriod, iDays, sErrorDescription)
'************************************************************
'Purpose: To register medical licence before vacations period
'			append registry of them
'Inputs:  oADODBConnection, aAbsenceComponent, lStartDate, lEndDate
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "MoveVacationsToLeft"
	Dim lErrorNumber
	Dim oRecordset
	Dim bComponentInitialized
	Dim lDaysFromMedicalLicence
	Dim lDaysToMove
	Dim lVacationEndDate
	Dim oDate
	Dim sQuery
	Dim iJourneyTypeID
	Dim updateVacationsDays

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador de la incidencia y/o la fecha para agregar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	End If

	lDaysToMove = DateDiff("d", GetDateFromSerialNumber(lStartDate), GetDateFromSerialNumber(aAbsenceComponent(N_END_DATE_ABSENCE))) + 1
	lDaysFromMedicalLicence = DateDiff("d", GetDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE)), GetDateFromSerialNumber(aAbsenceComponent(N_END_DATE_ABSENCE))) + 1

	Select Case lAbsenceId
		Case 35, 37, 38
			Call GetEmployeeJourneyType(oRequest, oADODBConnection, aEmployeeComponent, iJourneyTypeID, sErrorDescription)
			lVacationEndDate = AddDaysToSerialDateForVacations(oADODBConnection, lEndDate, lDaysToMove, iJourneyTypeID)
			updateVacationsDays = GetWorkingDaysOfAbsencesPeriod(AddDaysToSerialDate(lStartDate, lDaysToMove), lVacationEndDate, iJourneyTypeID)
		Case Else
			updateVacationsDays = DateDiff("d", GetDateFromSerialNumber(AddDaysToSerialDate(lStartDate, lDaysToMove)), GetDateFromSerialNumber(AddDaysToSerialDate(lEndDate, lDaysToMove))) + 1
	End Select

	sQuery = "Update EmployeesAbsencesLKP set OcurredDate = " & AddDaysToSerialDate(lStartDate, lDaysToMove) & _
			 " , EndDate = " & lVacationEndDate & _
			 " , AbsenceHours = " & updateVacationsDays & _
			 " where EmployeeId = " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & _
			 " and AbsenceID = " & lAbsenceId & _
			 " and OcurredDate = " & lStartDate
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	
	If lErrorNumber <> 0 Then
		sErrorDescription = "No se pudieron modificar las vacaciones del empleado."
	End If

	Set oRecordset = Nothing
	MoveVacationsToLeft = lErrorNumber
	Err.Clear
End Function

Function MoveVacationsToRight(oADODBConnection, aAbsenceComponent, lAbsenceId, lStartDate, lEndDate, lVacationPeriod, iDays, sErrorDescription)
'************************************************************
'Purpose: To register medical licence after vacations period
'			append registry of them
'Inputs:  oADODBConnection, aAbsenceComponent, lStartDate, lEndDate
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "MoveVacationsToRight"
	Dim lErrorNumber
	Dim oRecordset
	Dim bComponentInitialized
	Dim lDaysToMove
	Dim oDate
	Dim sQuery
	Dim iJourneyTypeID
	Dim updateVacationsDays
	Dim lNextStartDateForVacations
	Dim lNextEndDateForVacations   ' Siguiente día habíl para regstrar el fin de las vacaciones que se truncan
	Dim lFinalDateForVacations   ' Día de fin de las vacaciones
	Dim iNextDaysForVacations
	Dim iDaysAddForVacations

	bComponentInitialized = aAbsenceComponent(B_COMPONENT_INITIALIZED_ABSENCE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAbsenceComponent(oRequest, aAbsenceComponent)
	End If

	If (aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aAbsenceComponent(N_OCURRED_DATE_ABSENCE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador de la incidencia y/o la fecha para agregar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AbsenceComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	End If

	iDaysAddForVacations = 0
	Call GetEmployeeJourneyType(oRequest, oADODBConnection, aEmployeeComponent, iJourneyTypeID, sErrorDescription)
	lDaysToMove = DateDiff("d", GetDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE)), GetDateFromSerialNumber(lEndDate)) + 1
	lDaysFromMedicalLicence = DateDiff("d", GetDateFromSerialNumber(aAbsenceComponent(N_OCURRED_DATE_ABSENCE)), GetDateFromSerialNumber(aAbsenceComponent(N_END_DATE_ABSENCE))) + 1
	lNextStartDateForVacations = GetNextStartDateForVacations(AddDaysToSerialDate(aAbsenceComponent(N_END_DATE_ABSENCE), 1), iJourneyTypeID, iDaysAddForVacations)

	Select Case lAbsenceId
		Case 35, 37, 38
			lNextEndDateForVacations = GetNextEndDateForVacations(AddDaysToSerialDate(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), -1), iJourneyTypeID)
			updateVacationsDays = GetWorkingDaysOfAbsencesPeriod(lStartDate, lNextEndDateForVacations, iJourneyTypeID)
			lFinalDateForVacations = AddDaysToSerialDateForVacations(oADODBConnection, lNextStartDateForVacations, iDays - updateVacationsDays -1, iJourneyTypeID)
			iNextDaysForVacations = GetWorkingDaysOfAbsencesPeriod(lNextStartDateForVacations, lFinalDateForVacations, iJourneyTypeID)
		Case Else
			updateVacationsDays = DateDiff("d", lStartDate, AddDaysToSerialDate(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), -1)) + 1
			aAbsenceComponent(N_HOURS_ABSENCE) = DateDiff("d", GetDateFromSerialNumber(AddDaysToSerialDate(aAbsenceComponent(N_END_DATE_ABSENCE), 1)), GetDateFromSerialNumber(AddDaysToSerialDate(AddDaysToSerialDate(aAbsenceComponent(N_END_DATE_ABSENCE), 1), lDaysToMove))) + 1
	End Select

	sQuery = "Update EmployeesAbsencesLKP set EndDate = " & AddDaysToSerialDate(aAbsenceComponent(N_OCURRED_DATE_ABSENCE), -1) & _
			 " , AbsenceHours = " & updateVacationsDays & _
			 " where EmployeeId = " & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & _
			 " and AbsenceID = " & lAbsenceId & _
			 " and OcurredDate = " & lStartDate
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	If lErrorNumber = 0 Then
		sQuery = "Insert Into EmployeesAbsencesLKP (EmployeeID, AbsenceID, OcurredDate, EndDate, RegistrationDate, " & _
				 " DocumentNumber, AbsenceHours, JustificationID, AppliesForPunctuality, Reasons, AddUserID, AppliedDate, " & _
				 " Removed, RemoveUserID, RemovedDate, AppliedRemoveDate, Active, VacationPeriod) Values (" & aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE) & ", " & _
				 lAbsenceId & ", " & AddDaysToSerialDate(aAbsenceComponent(N_END_DATE_ABSENCE), 1) & ", " & _
				 AddDaysToSerialDate(AddDaysToSerialDate(aAbsenceComponent(N_END_DATE_ABSENCE), 1), lDaysToMove) & ", " & aAbsenceComponent(N_REGISTRATION_DATE_ABSENCE) & ", '" & _
				 Replace(aAbsenceComponent(S_DOCUMENT_NUMBER_ABSENCE), "'", "´") & "', " & aAbsenceComponent(N_HOURS_ABSENCE) & ", " & _
				 aAbsenceComponent(N_JUSTIFICATION_ID_ABSENCE) & ", " & aAbsenceComponent(N_APPLIES_FOR_PUNCTUALITY_ABSENCE) & ", '" & _
				 Replace(aAbsenceComponent(S_REASONS_ABSENCE), "'", "´") & "', " & aAbsenceComponent(N_ADD_USER_ID_ABSENCE) & ", " & _
				 aAbsenceComponent(N_APPLIED_DATE_ABSENCE) & ", " & aAbsenceComponent(N_REMOVED_ABSENCE) & ", " & _
				 aAbsenceComponent(N_REMOVE_USER_ID_ABSENCE) & ", " &  aAbsenceComponent(N_REMOVED_DATE_ABSENCE) & ", " & _
				 aAbsenceComponent(N_APPLIED_REMOVE_DATE_ABSENCE) & ", " & aAbsenceComponent(N_ACTIVE_ABSENCE) & ", " & _
				 lVacationPeriod & ")"

		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "AbsenceComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber <> 0 Then
			sErrorDescription = "No se pudieron modificar las vacaciones del empleado."
		End If
	Else
		sErrorDescription = "No se pudieron modificar las vacaciones del empleado."
	End If

	Set oRecordset = Nothing
	MoveVacationsToRight = lErrorNumber
	Err.Clear
End Function
%>