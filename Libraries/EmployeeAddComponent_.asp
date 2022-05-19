<%
Function AddDocumentsForLicenses(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new document for license into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddDocumentsForLicenses"
	Dim iDocumentDateMonth
	Dim iDocumentDateYear
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_DOCUMENT_FOR_LICENSE_ID_EMPLOYEE) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para el nuevo documento sindical."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "DocumentsForLicenses", "DocumentForLicenseID", "", 1, aEmployeeComponent(N_DOCUMENT_FOR_LICENSE_ID_EMPLOYEE), sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo guardar la información del nuevo documento sindical."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into DocumentsForLicenses (DocumentForLicenseID, DocumentForLicenseNumber, DocumentForCancelLicenseNumber, DocumentTemplate, RequestNumber, EmployeeID, LicenseSyndicateTypeID, DocumentLicenseDate, LicenseStartDate, LicenseEndDate, LicenseCancelDate, UserID) Values (" & aEmployeeComponent(N_DOCUMENT_FOR_LICENSE_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_DOCUMENT_FOR_LICENSE_NUMBER_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_DOCUMENT_FOR_CANCEL_LICENSE_NUMBER_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_DOCUMENT_TEMPLATE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_REQUEST_NUMBER_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_SYNDICATE_TYPE_ID_LICENSE_EMPLOYEE) & ", " & aEmployeeComponent(N_DATE_LICENSE_DOCUMENT_EMPLOYEE) & ", " & aEmployeeComponent(N_LICENSE_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_LICENSE_END_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_CANCEL_DATE_LICENSE_DOCUMENT_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	AddDocumentsForLicenses = lErrorNumber
	Err.Clear
End Function

Function AddEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployee"
	Dim sDate
	Dim iIndex
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sGenderID

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para el nuevo empleado."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "Employees", "EmployeeID", "", 1, aEmployeeComponent(N_ID_EMPLOYEE), sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		If aEmployeeComponent(B_CHECK_FOR_DUPLICATED_EMPLOYEE) Then
			lErrorNumber = CheckExistencyOfEmployee(aEmployeeComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aEmployeeComponent(B_IS_DUPLICATED_EMPLOYEE) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un empleado con el número " & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeAddComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				If Not CheckEmployeeInformationConsistency(aEmployeeComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					sDate = Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000"))
					If aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 0 Then aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 30000000
					sErrorDescription = "No se pudo guardar la información del nuevo empleado."
					If Len(aEmployeeComponent(S_CURP_EMPLOYEE)) = 18 Then
						sGenderID  = Mid(aEmployeeComponent(S_CURP_EMPLOYEE), Len("00000000000"), Len("0"))
						If (InStr(1, sGenderID, "M", vbBinaryCompare) > 0) Then	
							aEmployeeComponent(N_GENDER_ID_EMPLOYEE) = 0
						Else
							If (InStr(1, sGenderID, "H", vbBinaryCompare) > 0) Then	
								aEmployeeComponent(N_GENDER_ID_EMPLOYEE) = 1
							End If
						End If
					End If
					If B_UPPERCASE Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Employees (EmployeeID, EmployeeNumber, EmployeeAccessKey, EmployeePassword, EmployeeName, EmployeeLastName, EmployeeLastName2, CompanyID, JobID, ServiceID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, StartHour1, EndHour1, StartHour2, EndHour2, StartHour3, EndHour3, WorkingHours, LevelID, StatusID, PaymentCenterID, EmployeeEmail, SocialSecurityNumber, BirthYear, BirthMonth, BirthDay, BirthDate, StartDate, StartDate2, CountryID, RFC, CURP, GenderID, MaritalStatusID, AntiquityID, Antiquity2ID, Antiquity3ID, Antiquity4ID, ModifyDate, Active) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_PASSWORD_EMPLOYEE), "'", "") & "', '" & Replace(UCase(aEmployeeComponent(S_NAME_EMPLOYEE)), "'", "´") & "', '" & Replace(UCase(aEmployeeComponent(S_LAST_NAME_EMPLOYEE)), "'", "´") & "', '" & Replace(UCase(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE)), "'", "´") & "', " & aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & ", " & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & ", " & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & ", " & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & ", " & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", " & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & ", " & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & ", " & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_EMAIL_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_SSN_EMPLOYEE), "'", "") & "', " & CInt(Left(sDate, Len("0000"))) & ", " & CInt(Mid(sDate, Len("00000"), Len("00"))) & ", " & CInt(Mid(sDate, Len("0000000"), Len("00"))) & ", " & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_START_DATE2_EMPLOYEE) & ", " & aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) & ", '" & Replace(UCase(aEmployeeComponent(S_RFC_EMPLOYEE)), "'", "") & "', '" & Replace(UCase(aEmployeeComponent(S_CURP_EMPLOYEE)), "'", "") & "', " & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) & ", 0, 0, 0, 0, " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Employees (EmployeeID, EmployeeNumber, EmployeeAccessKey, EmployeePassword, EmployeeName, EmployeeLastName, EmployeeLastName2, CompanyID, JobID, ServiceID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, StartHour1, EndHour1, StartHour2, EndHour2, StartHour3, EndHour3, WorkingHours, LevelID, StatusID, PaymentCenterID, EmployeeEmail, SocialSecurityNumber, BirthYear, BirthMonth, BirthDay, BirthDate, StartDate, StartDate2, CountryID, RFC, CURP, GenderID, MaritalStatusID, AntiquityID, Antiquity2ID, Antiquity3ID, Antiquity4ID, ModifyDate, Active) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_PASSWORD_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_NAME_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_LAST_NAME_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE), "'", "´") & "', " & aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & ", " & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & ", " & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & ", " & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & ", " & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", " & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & ", " & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & ", " & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_EMAIL_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_SSN_EMPLOYEE), "'", "") & "', " & CInt(Left(sDate, Len("0000"))) & ", " & CInt(Mid(sDate, Len("00000"), Len("00"))) & ", " & CInt(Mid(sDate, Len("0000000"), Len("00"))) & ", " & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_START_DATE2_EMPLOYEE) & ", " & aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_RFC_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_CURP_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) & ", 0, 0, 0, 0, " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudo guardar la información del nuevo empleado."
						If InStr(1, ",5,6,7,8,9,", "," & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ",", vbBinaryCompare) = 0 Then
							If aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 12 Or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 13 Then
                            	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConsecutiveIDs Set CurrentID=CurrentID+1 Where (IDType=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
                            Else
								If aEmployeeComponent(N_ID_EMPLOYEE) >= 1000000 Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConsecutiveIDs2 Set CurrentID=CurrentID+1 Where (IDType=-1)", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
								Else
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConsecutiveIDs Set CurrentID=CurrentID+1 Where (IDType=-1)", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
								End If
							End If
                        ElseIf InStr(1, ",5,6,", "," & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ",", vbBinaryCompare) > 0 Then
							If aEmployeeComponent(N_ID_EMPLOYEE) >= 1000000 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConsecutiveIDs2 Set CurrentID=CurrentID+1 Where (IDType=5)", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
							Else
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConsecutiveIDs Set CurrentID=CurrentID+1 Where (IDType=5)", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
							End If
						Else
							If aEmployeeComponent(N_ID_EMPLOYEE) >= 1000000 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConsecutiveIDs2 Set CurrentID=CurrentID+1 Where (IDType=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
							Else
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConsecutiveIDs Set CurrentID=CurrentID+1 Where (IDType=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
							End If
						End If
					End If
					If (aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 12) Or (aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 13) Or (aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 14) Or (aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 17) Or (aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 18) Or (aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 26) Or (aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 57) Or (aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 58) Then
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudo guardar la información del nuevo empleado."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesExtraInfo Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo modificar la información del empleado."
								If B_UPPERCASE Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesExtraInfo (EmployeeID, EmployeeAddress, EmployeeCity, EmployeeZipCode, StateID, CountryID, EmployeePhone, OfficePhone, OfficeExt, DocumentNumber1, DocumentNumber2, DocumentNumber3, EmployeeActivityID) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", '" & Replace(UCase(aEmployeeComponent(S_ADDRESS_EMPLOYEE)), "'", "´") & "', '" & Replace(UCase(aEmployeeComponent(S_CITY_EMPLOYEE)), "'", "´") & "', '" & Replace(aEmployeeComponent(S_ZIP_CODE_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_ADDRESS_COUNTRY_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_EMPLOYEE_PHONE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_OFFICE_PHONE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_EXT_OFFICE_EMPLOYEE), "'", "") & "', '" & Replace(UCase(aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE)), "'", "") & "', '" & Replace(UCase(aEmployeeComponent(S_DOCUMENT_NUMBER_2_EMPLOYEE)), "'", "") & "', '" & Replace(UCase(aEmployeeComponent(S_DOCUMENT_NUMBER_3_EMPLOYEE)), "'", "") & "', " & aEmployeeComponent(N_ACTIVITY_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
								Else
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesExtraInfo (EmployeeID, EmployeeAddress, EmployeeCity, EmployeeZipCode, StateID, CountryID, EmployeePhone, OfficePhone, OfficeExt, DocumentNumber1, DocumentNumber2, DocumentNumber3, EmployeeActivityID) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_ADDRESS_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_CITY_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_ZIP_CODE_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_ADDRESS_COUNTRY_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_EMPLOYEE_PHONE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_OFFICE_PHONE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_EXT_OFFICE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_DOCUMENT_NUMBER_2_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_DOCUMENT_NUMBER_3_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_ACTIVITY_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
								End If
							End If
						End If
					End If
					If lErrorNumber = 0 Then
						If aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 0 Then aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
						sErrorDescription = "No se pudo guardar la información del nuevo empleado."
						If lReasonID = 0 Then
							aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 0
						End If
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_ZONE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & ", " & aEmployeeComponent(N_AREA_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", 0, '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
					If lErrorNumber = 0 Then
						For iIndex = 2000 To 2100
							sErrorDescription = "No se pudo guardar la información del nuevo empleado."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesForTaxAdjustment (EmployeeID, PayrollYear, bTaxAdjustment) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & iIndex & ", 1)", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						Next
					End If
				End If
			End If
		End If
	End If

	AddEmployee = lErrorNumber
	Err.Clear
End Function

Function AddEmployeeAbsences(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new absence for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeAbsences"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim iEndDate
	Dim sEmployeeAbsenceIDs
	Dim sConceptShortName

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	Select Case aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE)
		Case 10 ' Horas extras
			aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 201
			sConceptShortName = "horas extras"
		Case 16, 17 ' Prima dominical
			aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 202
			aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = 1
			sConceptShortName = "prima dominical"
	End Select

	aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
	aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) = 1
	aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)

	If VerifyRequerimentsToEmployeesAbsences(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
		aEmployeeComponent(B_IS_DUPLICATED_EMPLOYEE) = False
		lErrorNumber = CheckExistencyOfEmployeeAbsence(aEmployeeComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			If aEmployeeComponent(B_IS_DUPLICATED_EMPLOYEE) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un registro de " & sConceptShortName & " para el empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " en el día " & DisplayDateFromSerialNumber(aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE), -1, -1, -1) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeAddComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				If Not CheckEmployeeConceptInformationConsistency(aEmployeeComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					lErrorNumber = GetEmployeeAbsencesAppliesToID(oRequest, oADODBConnection, aEmployeeComponent, sEmployeeAbsenceIDs, sErrorDescription)
					If lErrorNumber = 0 Then
						If VerifyExistenceOfEmployeeAbsences(oADODBConnection, aAbsenceComponent, sEmployeeAbsenceIDs, sErrorDescription) Then
							sErrorDescription = "No se pudo guardar la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesAbsencesLKP (EmployeeID, AbsenceID, OcurredDate, EndDate, RegistrationDate, DocumentNumber, AbsenceHours, JustificationID, AppliesForPunctuality, Reasons, AddUserID, AppliedDate, Removed, RemoveUserID, RemovedDate, AppliedRemoveDate, Active, VacationPeriod) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000"))  & ", '.', " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_JUSTIFICATION_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_FOR_PUNCTUALITY_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "', " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", 0, -1, 0, 0, 0, 0)", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						Else
							lErrorNumber = -1
						End If
					Else
						sErrorDescription = "Error al validar las incidencias registradas."
					End If
				End If
			End If
		End If
	End If

	AddEmployeeAbsences = lErrorNumber
	Err.Clear
End Function

Function AddEmployeeAdjustment(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new payment claims adjustments and deductions for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeAdjustment"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim lDate
	Dim iForPayrollIsActiveConstant

	lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para agregar la información del beneficiario."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeAddComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "El número de empleado no existe."
		lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			sErrorDescription = "El concepto ya fue registrado con anterioridad."
			If aEmployeeComponent(N_MISSING_DATE_EMPLOYEE) < lDate Then
				If CInt(Request.Cookies("SIAP_SectionID")) = 1 Then
					iForPayrollIsActiveConstant = N_PAYROLL_FOR_MOVEMENTS
				ElseIf (CInt(Request.Cookies("SIAP_SectionID")) = 2) Or (CInt(Request.Cookies("SIAP_SectionID")) = 7) Then
					iForPayrollIsActiveConstant = N_PAYROLL_FOR_FEATURES
				ElseIf CInt(Request.Cookies("SIAP_SectionID")) = 4 Then
					iForPayrollIsActiveConstant = 0
				End If
				If VerifyPayrollIsActive(oADODBConnection, aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE), iForPayrollIsActiveConstant, sErrorDescription) Then
					lErrorNumber = CheckExistencyOfEmployeeAdjustment(aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudo agregar la información del reclamo de pago por ajustes y deducciones del empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " con fecha de reclamo " & DisplayDateFromSerialNumber(aEmployeeComponent(N_MISSING_DATE_EMPLOYEE), -1, -1, -1) & "."
						If lErrorNumber = 0 Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesAdjustmentsLKP (EmployeeID, ConceptID, ConceptAmount, MissingDate, PaymentDate, ModifyDate, PayrollDate, BeneficiaryName, UserID, Active, AdjustmentType) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", " & aEmployeeComponent(N_MISSING_DATE_EMPLOYEE) & ", 0, " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NAME_BENEFICIARY_EMPLOYEE), "'", "") & "', " & aLoginComponent(N_USER_ID_LOGIN) & ", 0, 0)", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
						End If
					End If
				Else
					lErrorNumber = -1
				End If
			Else
				lErrorNumber = -1
				sErrorDescription = "La fecha de omisión para el reclamo de pago no puede ser mayor o igual a la fecha actual."
			End If
		End If
	End If

	Set oRecordset = Nothing
	AddEmployeeAdjustment = lErrorNumber
	Err.Clear
End Function

Function AddEmployeeBankAccount(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new beneficiary for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeBankAccount"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim lEndHistoryDate
	Dim lStartDate
	Dim aZonePath

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para agregar la cuenta bancaria del empleado."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeAddComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If lErrorNumber = 0 Then
			If aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 0 Then aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 30000000
			If VerifyRequerimentsForEmployeesBankAccounts(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
				If (lReasonID <> EMPLOYEES_ADD_BENEFICIARIES) And (lReasonID <> EMPLOYEES_CREDITORS) And (CheckExistencyOfEmployeeBankAccount(aEmployeeComponent, sErrorDescription)) Then
					lErrorNumber = L_ERR_DUPLICATED_RECORD
				Else
					If aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) = -1 Then
						If CheckExistencyOfActiveAccount(oADODBConnection,aEmployeeComponent) = False Then
							sErrorDescription = "No se pudo obtener un identificador para la nueva cuenta bancaria."
							lErrorNumber = GetNewIDFromTable(oADODBConnection, "BankAccounts", "AccountID", "", 1, aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE), sErrorDescription)
							If lErrorNumber = 0 Then
								'aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 1
								sErrorDescription = "No se pudo agregar la información de la cuenta bancaria del empleado."
								If InStr(1,",12,13,14,17,18,68,","," & lReasonID & ",",vbBinaryCompare) <> 0 Then
									aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE) = "."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ZonePath From Zones Where ZoneID = (Select ZoneID From Jobs Where JobID = " & aJobComponent(N_ID_JOB) & ")" , "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									aZonePath = Split(oRecordset.Fields("ZonePath").Value,",")
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Bankid From BankAccounts Where StateID = " & aZonePath(2) , "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									aEmployeeComponent(N_BANK_ID_EMPLOYEE) = oRecordset.Fields("BankID").Value
								Else
									aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE) = -1
								End If
								If InStr(1,",-85,-99,", "," & lReasonID & ",", vbBinaryCompare) <> 0 Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From BankAccounts Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									If ((lErrorNumber = 0) And (Not oRecordset.EOF)) Then
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From BankAccounts Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
									End If
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaPath From Areas Where (AreaID=" & aEmployeeComponent(N_PAYMENT_CENTER_ID_BENEFICIARY_EMPLOYEE) & ")" , "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									aZonePath = Split(oRecordset.Fields("AreaPath").Value,",")
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select BankId From BankAccounts Where StateID = " & aZonePath(2) , "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									aEmployeeComponent(N_BANK_ID_EMPLOYEE) = oRecordset.Fields("BankID").Value
									aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 1
								End If
								If (lReasonID = EMPLOYEES_BANK_ACCOUNTS) Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into BankAccounts (AccountID, EmployeeID, BankID, AccountNumber, StateID, StartDate, EndDate, RegistrationDate, UserID, Active) Values (" & aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_BANK_ID_EMPLOYEE) & ", '" & aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE) & "', " & aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE) &", " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
								Else
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into BankAccounts (AccountID, EmployeeID, BankID, AccountNumber, StateID, StartDate, EndDate, RegistrationDate, UserID, Active) Values (" & aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_BANK_ID_EMPLOYEE) & ", '" & aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE) & "', " & aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE) &", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
								End If
								If InStr(1,",12,13,14,17,18,68,","," & lReasonID & ",",vbBinaryCompare) <> 0 Then
									If lErrorNumber = 0 Then
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AccountID, StartDate, EndDate From BankAccounts Where AccountID <> " & aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) & "And EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " Order By StartDate Desc" , "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
										lEndHistoryDate = AddDaysToSerialDate(oRecordset.Fields("StartDate").Value, -1)
										oRecordset.MoveNext
										Do While Not oRecordset.EOF
											lStartDate = oRecordset.Fields("StartDate").Value
											If oRecordset.Fields("EndDate") = 30000000 Then
												If lEndHistoryDate <= lStartDate Then
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update BankAccounts Set EndDate = " & lEndHistoryDate & " Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EndDate = " & oRecordset.Fields("EndDate").Value & ") And (AccountID <> " & aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
												Else
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update BankAccounts Set EndDate = " & lStartDate & " Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EndDate = " & oRecordset.Fields("EndDate").Value & ") And (AccountID <> " & aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
												End If
											End If
											lEndHistoryDate = AddDaysToSerialDate(oRecordset.Fields("StartDate").Value, -1)
											oRecordset.MoveNext
										Loop
									End If
								End If
							End If
						Else
							
						End If
					Else
						sErrorDescription = "No se pudo agregar la información de la cuenta bancaria del empleado."
						lErrorNumber = -1
					End If
				End If
			Else
				lErrorNumber = -1
			End If
		End If
	End If

	Set oRecordset = Nothing
	AddEmployeeBankAccount = lErrorNumber
	Err.Clear
End Function

Function AddEmployeeBeneficiary(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new beneficiary for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeBeneficiary"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim lTotalAmount

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para agregar la información del beneficiario."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeAddComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aEmployeeComponent(N_END_DATE_BENEFICIARY_EMPLOYEE) = 0 Then aEmployeeComponent(N_END_DATE_BENEFICIARY_EMPLOYEE) = 30000000
		If aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) = -1 Then
			sErrorDescription = "No se pudo obtener un identificador para el nuevo beneficiario."
			lErrorNumber = CheckExistencyOfEmployeeBeneficiary(aEmployeeComponent, sErrorDescription)
			If lErrorNumber = L_ERR_NO_RECORDS Then
				Call GetAlimonyTypesTotalAmountForEmployee(lTotalAmount, sErrorDescription)
				If lTotalAmount + aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) > 100 Then
					lErrorNumber = -1
					sErrorDescription = "La cantidad total de beneficiarios que aplican por porcentaje exceden el 100 % en los efectos del registro con agregado, con fecha de inicio " & DisplayDateFromSerialNumber(aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE), -1, -1, -1) & "."
				Else
					lErrorNumber = GetNewIDFromTable(oADODBConnection, "EmployeesBeneficiariesLKP", "BeneficiaryID", "(EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", 1, aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE), sErrorDescription)
					If lErrorNumber = 0 Then
						If Not CheckEmployeeBeneficiaryInformationConsistency(aEmployeeComponent, sErrorDescription) Then
							lErrorNumber = -1
						Else
							sErrorDescription = "No se pudo agregar la información del beneficiario del empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " con fecha de inicio " & DisplayDateFromSerialNumber(aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE), -1, -1, -1) & "."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesBeneficiariesLKP (EmployeeID, BeneficiaryID, StartDate, EndDate, BeneficiaryNumber, BeneficiaryName, BeneficiaryLastName, BeneficiaryLastName2, BeneficiaryBirthDate, AlimonyTypeID, ConceptAmount, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, PaymentCenterID, StartUserID, EndUserID, Comments, Active) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) & ", " & aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE) & ", " & aEmployeeComponent(N_END_DATE_BENEFICIARY_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_BENEFICIARY_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_NAME_BENEFICIARY_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_LAST_NAME_BENEFICIARY_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_LAST_NAME2_BENEFICIARY_EMPLOYEE), "'", "´") & "', " & aEmployeeComponent(N_BIRTH_DATE_BENEFICIARY_EMPLOYEE) & ", " & aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", " &  aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE)  & ", " & aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE)  & ", " &  aEmployeeComponent(N_PAYMENT_CENTER_ID_BENEFICIARY_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", -1, '" & Replace(aEmployeeComponent(S_COMMENTS_BENEFICIARY_EMPLOYEE), "'", "´") & "', 0)", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
						End If
					End If
					aEmployeeComponent(N_ID_EMPLOYEE) = aEmployeeComponent(S_NUMBER_BENEFICIARY_EMPLOYEE)
					aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = 0
					aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 30000000
					aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE) = "."
					lErrorNumber = AddEmployeeBankAccount(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				End If
			Else
				lErrorNumber = -1
				sErrorDescription = "Ya existe un beneficiario de pensión alimenticia con el mismo número que el indicado."
			End If
		End If
	End If

	Set oRecordset = Nothing
	AddEmployeeBeneficiary = lErrorNumber
	Err.Clear
End Function

Function AddEmployeeChild(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new child for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeChild"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para agregar la información de su hijo(a)."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeAddComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aEmployeeComponent(N_ID_CHILD_EMPLOYEE) = -1 Then
			sErrorDescription = "No se pudo obtener un identificador para el nuevo hijo(a)."
			lErrorNumber = GetNewIDFromTable(oADODBConnection, "EmployeesChildrenLKP", "ChildID", "(EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", 1, aEmployeeComponent(N_ID_CHILD_EMPLOYEE), sErrorDescription)
			If lErrorNumber = 0 Then
				If aEmployeeComponent(N_ID_CHILD_EMPLOYEE) = 0 Then aEmployeeComponent(N_ID_CHILD_EMPLOYEE) = 1
				If Not CheckEmployeeChildInformationConsistency(aEmployeeComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					sErrorDescription = "No se pudo agregar la información del hijo(a) del empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " con nombre " & aEmployeeComponent(S_NAME_CHILD_EMPLOYEE) & " " & aEmployeeComponent(S_LAST_NAME_CHILD_EMPLOYEE) & " " & aEmployeeComponent(S_LAST_NAME2_CHILD_EMPLOYEE) & "."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesChildrenLKP (EmployeeID, ChildID, ChildName, ChildLastName, ChildLastName2, ChildBirthDate, ChildEndDate, LevelID, RegistrationDate, UserID) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NAME_CHILD_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_LAST_NAME_CHILD_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_LAST_NAME2_CHILD_EMPLOYEE), "'", "´") & "', " & aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE) & ", " & aEmployeeComponent(N_END_DATE_CHILD_EMPLOYEE) & ", " & aEmployeeComponent(N_CHILD_LEVEL_ID_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	AddEmployeeChild = lErrorNumber
	Err.Clear
End Function

Function AddEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new concept for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeConcept"
	Dim oRecordset
	Dim lErrorNumber
	Dim iExistenceType
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para agregar la información de su hijo(a)."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeAddComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		Select Case aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE)
			Case 4
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1,89"
				aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 3
			Case 5
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 1
				aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 1
				aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 1
			Case 50
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1"
			Case 7, 8
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1,5"
				aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = 3/6.5*100
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
			Case 26
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "-1"
				aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = 12
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 13
			Case 87
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
			Case 22, 24, 45, 46, 50, 93, 94, 32, 76 ' EMPLOYEES_CHILDREN_SCHOOLARSHIPS, EMPLOYEES_GLASSES, EMPLOYEES_FAMILY_DEATH, EMPLOYEES_PROFESSIONAL_DEGREE, EMPLOYEES_MONTHAWARD, EMPLOYEES_NIGHTSHIFTS, EMPLOYEES_NON_EXCENT, EMPLOYEES_EXCENT, EMPLOYEES_CONCEPT_C3, EMPLOYEES_MOTHERAWARD, -89, EMPLOYEES_ANUAL_AWARD, EMPLOYEES_FONAC_ADJUSTMENT
				aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)
			Case Else
				If (aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 0) And (Not aEmployeeComponent(B_CANCEL_CONCEPT_FOR_EMPLOYEE)) Then
					aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 30000000
				End If
		End Select
		'aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
		aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) = 1
		If Not CheckEmployeeConceptInformationConsistency(aEmployeeComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			If VerifyRequerimentsForEmployeesConcepts(oADODBConnection, CInt(oRequest("ReasonID").Item), aEmployeeComponent, sErrorDescription) Then
				If lErrorNumber = 0 Then
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
						'Case 93
						'	If Not IsHoliday(oADODBConnection, aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), sErrorDescription) Then
						'		sErrorDescription = "El concepto solo se puede registrar en día festivo."
						'		lErrorNumber = -1
						'	End If
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
					If lErrorNumber = 0 Then
						If VerifyExistenceOfEmployeesConcept(oADODBConnection, aEmployeeComponent, iExistenceType, sErrorDescription) Then
							lErrorNumber = L_ERR_DUPLICATED_RECORD
						Else
							sErrorDescription = "No se pudo guardar la información del concepto para el empleado."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesConceptsLKP (EmployeeID, ConceptID, StartDate, EndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, RegistrationDate, ModifyDate, StartUserID, EndUserID, UploadedFileName, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) & ", '" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "', " & aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", -1, '" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							If aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 32 Then
								If lErrorNumber <> 0  Then
									lErrorNumber = -1
									sErrorDescription = "No se pudo agregar el registro del concepto para el empleado " & aEmployeeComponent(N_ID_EMPLOYEE)
								Else
									lErrorNumber = SetActiveForEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								End If
							End If
						End If
					End If
				End If
			Else
				lErrorNumber = -1
			End If
		End If
	End If

	AddEmployeeConcept = lErrorNumber
	Err.Clear
End Function

Function AddEmployeeConceptsFile(oRequest, oADODBConnection, sQuery, lReasonID, aEmployeeComponent, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new movement for the employee into the database
'Inputs:  oRequest, oADODBConnection, lReasonID
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeConceptsFile"
	Dim oRecordset
	Dim lErrorNumber
	Dim iRecordCount

	sErrorDescription = "No se pudo obtener la información para la aplicación de los conceptos de pago de los empleados."
	iRecordCount = 0
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				Select Case lReasonID
					Case -58
						'If Not IsEmpty(oRequest(CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("ConceptID").Value) & CStr(oRecordset.Fields("MissingDate").Value))) Then
							aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = CLng(oRecordset.Fields("ConceptID").Value)
							aEmployeeComponent(N_MISSING_DATE_EMPLOYEE) = CLng(oRecordset.Fields("MissingDate").Value)
							lErrorNumber = SetActiveForEmployeeAdjustment(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							If lErrorNumber = 0 Then
								iRecordCount = iRecordCount + 1
							End If
						'End If
					Case EMPLOYEES_ADD_BENEFICIARIES
						If Not IsEmpty(oRequest(CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("BeneficiaryID").Value) & CStr(oRecordset.Fields("StartDate").Value))) Then
							aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
							aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) = CLng(oRecordset.Fields("BeneficiaryID").Value)
							aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE) = CLng(oRecordset.Fields("StartDate").Value)
							lErrorNumber = SetActiveForEmployeeBeneficiary(oRequest, oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription)
							If lErrorNumber = 0 Then
								iRecordCount = iRecordCount + 1
							End If
						End If
					Case EMPLOYEES_BANK_ACCOUNTS
						If Not IsEmpty(oRequest(CStr(oRecordset.Fields("AccountID").Value))) Then
							aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) = CLng(oRecordset.Fields("AccountID").Value)
							aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
							aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = CLng(oRecordset.Fields("StartDate").Value)
							lErrorNumber = SetActiveForEmployeeBankAccount(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							If lErrorNumber = 0 Then
								iRecordCount = iRecordCount + 1
							End If
						End If
					Case EMPLOYEES_BENEFICIARIES_DEBIT
						If Not IsEmpty(oRequest(CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("CreditID").Value))) Then
							aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
							aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) = CLng(oRecordset.Fields("CreditID").Value)
							lErrorNumber = GetEmployeeCredit(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							If lErrorNumber = 0 Then
								lErrorNumber = SetActiveForEmployeeCreditFile(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								If lErrorNumber = 0 Then
									iRecordCount = iRecordCount + 1
								End If
							End If
						End If
					Case EMPLOYEES_CREDITORS
						If Not IsEmpty(oRequest(CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("CreditorID").Value) & CStr(oRecordset.Fields("StartDate").Value))) Then
							aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
							aEmployeeComponent(N_ID_CREDITOR_EMPLOYEE) = CLng(oRecordset.Fields("CreditorID").Value)
							aEmployeeComponent(N_START_DATE_CREDITOR_EMPLOYEE) = CLng(oRecordset.Fields("StartDate").Value)
							lErrorNumber = SetActiveForEmployeeBeneficiary(oRequest, oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription)
							If lErrorNumber = 0 Then
								iRecordCount = iRecordCount + 1
							End If
						End If
					Case EMPLOYEES_THIRD_CONCEPT, EMPLOYEES_THIRD_PROCESS, EMPLOYEES_BENEFICIARIES_DEBIT
						aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
						aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) = CLng(oRecordset.Fields("CreditID").Value)
						lErrorNumber = GetEmployeeCredit(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
						If lErrorNumber = 0 Then
							lErrorNumber = SetActiveForEmployeeCreditFile(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							If lErrorNumber = 0 Then
								iRecordCount = iRecordCount + 1
							End If
						End If
					Case EMPLOYEES_GRADE
						If Not IsEmpty(oRequest(CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("StartDate").Value))) Then
							aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
							aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = CLng(oRecordset.Fields("StartDate").Value)
							lErrorNumber = SetActiveForEmployeeGrade(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							If lErrorNumber = 0 Then
								iRecordCount = iRecordCount + 1
							End If
						End If
					Case Else
						If Not IsEmpty(oRequest(CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("ConceptID").Value) & CStr(oRecordset.Fields("StartDate").Value))) Then
							aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = CLng(oRecordset.Fields("ConceptID").Value)
							aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = CLng(oRecordset.Fields("StartDate").Value)
							lErrorNumber = SetActiveForEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							If lErrorNumber = 0 Then
								iRecordCount = iRecordCount + 1
							End If
						End If
				End Select
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
		End If
	End If
	sErrorDescription = "Se aplicaron " & iRecordCount & " registros"

	Set oRecordset = Nothing
	AddEmployeeConceptsFile = lErrorNumber
	Err.Clear
End Function

Function AddEmployeeCreditForValidation(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new credit for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeCreditForValidation"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim iEndDate
	Dim sQuery
	Dim lCreditID
	Dim iPaymentNumber
	Dim iExistenceType

	iPaymentNumber = CInt(oRequest("ConceptQttyID").Item)
	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador del tipo de crédito para validar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeAddComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Not CheckEmployeeConceptInformationConsistency(aEmployeeComponent, sErrorDescription) Then
			lErrorNumber = -1
			sErrorDescription = "No se pudo validar la consistencia de la información."
		Else
			If Not CheckEmployeeCreditInformationConsistency(aEmployeeComponent, sErrorDescription) Then
				lErrorNumber = -1
				sErrorDescription = "No se pudo validar la consistencia de la información para cargar el crédito."
			Else
				If aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) = -1 Then
					sErrorDescription = "No se pudo obtener un identificador para cargar el nuevo crédito."
					lErrorNumber = GetNewIDFromTable(oADODBConnection, "Credits", "CreditID", "(EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", 1, aEmployeeComponent(N_CREDIT_ID_EMPLOYEE), sErrorDescription)
					If lErrorNumber <> 0 Then
						sErrorDescription = "No se pudo obtener un identificador para el nuevo crédito."
					Else
						Select Case aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE)
							Case 64
								aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1,6"
								aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
							Case Else
								aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "\"
								aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 1
						End Select
						Select Case aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE)
							Case 1 ' Nuevo
								If VerifyExistenceOfEmployeesCredit(oADODBConnection, aEmployeeComponent, iExistenceType, sErrorDescription) Then
									If (CInt(oRequest("ReasonID").Item) = EMPLOYEES_THIRD_PROCESS) Then
										Select Case aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE)
											Case 59, 61, 64, 81, 82, 83, 126
												If (iExistenceType=1) Then
													aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = "EXISTE UN REGISTRO ABIERTO PARA ESTE TIPO DE CREDITO, SE CERRARAN SUS EFECTOS CUANDO APLIQUE EL NUEVO REGISTRO. "
												Else
													lErrorNumber = -1
												End If
											Case Else
												lErrorNumber = -1
										End Select
									End If
								End If
								If aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = 0 Then
									sErrorDescription = "Para poder dar de alta el registro, el monto de la cuota fija no puede ser igual a cero."
									lErrorNumber = -1
								Else
									If aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) = 0 Then
										aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) & "EL NUMERO DE PAGOS DEL REGISTRO ES 0. "
									End If
								End If
							Case 2 ' Modificación
								If VerifyExistenceOfEmployeesCredit(oADODBConnection, aEmployeeComponent, iExistenceType, sErrorDescription) Then
									'aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = "EXISTE UN REGISTRO ABIERTO PARA ESTE CREDITO, EL CUAL SE MODIFICARA CUANDO APLIQUE EL CAMBIO. "
									If iExistenceType <> 1 Then
										aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = "NO EXISTE EL CREDITO A MODIFICAR, ESTE SE APLICARA COMO ALTA CUANDO LO ACTIVE. "
									Else
										aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = "EXISTE UN REGISTRO ABIERTO PARA ESTE CREDITO, EL CUAL SE MODIFICARA CUANDO APLIQUE EL CAMBIO. "
									End If
								'Else
								'	aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = "NO EXISTE EL CREDITO A MODIFICAR, ESTE SE APLICARA COMO ALTA CUANDO LO ACTIVE. "
								End If
								If aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = 0 Then
									aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) & "EL IMPORTE DE LA CUOTA FIJA DEL REGISTRO ES 0. "
								End If
								If aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) = 0 Then
									aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) & "EL NUMERO DE PAGOS DEL REGISTRO ES 0. "
								End If
							Case 3 ' Baja
								If VerifyExistenceOfEmployeesCredit(oADODBConnection, aEmployeeComponent, iExistenceType, sErrorDescription) Then
									If iExistenceType <> 1 Then
										aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = "NO EXISTE EL CREDITO PARA DARLO DE BAJA, VERIFICAR LOS DATOS. "
									End If
								End If
								If aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) <> 0 Then
									aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) & "EL IMPORTE DE LA CUOTA FIJA DEL REGISTRO ES DISTINTO DE 0. "
								End If
								If aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) <> 0 Then
									aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) & "EL NUMERO DE PAGOS DEL REGISTRO ES DISTINTO DE 0. "
								End If
							Case Else
								If VerifyExistenceOfEmployeesCredit(oADODBConnection, aEmployeeComponent, iExistenceType, sErrorDescription) Then
									If iExistenceType <> 1 Then
										lErrorNumber = -1
									End If
								End If
						End Select
						If lErrorNumber = 0 Then
							If aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = -1 Then aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 0
							sErrorDescription = "No se pudo agregar la información del crédito para el empleado: " & aEmployeeComponent(N_ID_EMPLOYEE)
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Credits (EmployeeID, CreditID, CreditTypeID, ContractNumber, AccountNumber, PaymentsNumber, PeriodID, StartDate, EndDate, FinishDate, QttyID, AppliesToID, StartAmount, PaymentAmount, DebtAmount, PaymentsCounter, Active, UploadedFileName, Comments, UploadedRecordType, BeneficiaryID, UserID, RegistrationDate) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_CREDIT_CONTRACT_NUMBER_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_CREDIT_ACCOUNT_NUMBER_EMPLOYEE), "'", "´") & "', " & aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) & ", " & aEmployeeComponent(N_CREDIT_PERIOD_ID_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CREDIT_FINISH_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & ", '" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "', " & aEmployeeComponent(D_CREDIT_START_AMOUNT_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", " & aEmployeeComponent(D_CREDIT_START_AMOUNT_EMPLOYEE) & ", " & aEmployeeComponent(N_CREDIT_PAYMENTS_COUNTER_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "', " & aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) & ", -1, " & aLoginComponent(N_USER_ID_LOGIN) & ", " & CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						End If
					End If
				Else
					lErrorNumber = GetNewIDFromTable(oADODBConnection, "Credits", "CreditID", "(EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", 1, aEmployeeComponent(N_CREDIT_ID_EMPLOYEE), sErrorDescription)
					If lErrorNumber <> 0 Then
						sErrorDescription = "No se pudo obtener un identificador para el nuevo crédito."
					Else
						If VerifyExistenceOfEmployeesCredit(oADODBConnection, aEmployeeComponent, 1, sErrorDescription) Then
							lErrorNumber = -1
							aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = "NO EXISTE EL CREDITO PARA DARLO DE BAJA, VERIFICAR LOS DATOS. "
						End If
						sErrorDescription = "No se pudo agregar la información del crédito para el empleado: " & aEmployeeComponent(N_ID_EMPLOYEE)
 						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Credits (EmployeeID, CreditID, CreditTypeID, ContractNumber, AccountNumber, PaymentsNumber, PeriodID, StartDate, EndDate, FinishDate, QttyID, AppliesToID, StartAmount, PaymentAmount, DebtAmount, PaymentsCounter, Active, UploadedFileName, Comments, UploadedRecordType, BeneficiaryID, UserID, RegistrationDate) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_CREDIT_CONTRACT_NUMBER_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_CREDIT_ACCOUNT_NUMBER_EMPLOYEE), "'", "´") & "', " & aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) & ", " & aEmployeeComponent(N_CREDIT_PERIOD_ID_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CREDIT_FINISH_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & ", '" &aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "', " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", " & aEmployeeComponent(N_CREDIT_PAYMENTS_COUNTER_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "', " & aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) & ", " & aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
 					End If
				End If
			End If
		End If
	End If
	AddEmployeeCreditForValidation = lErrorNumber
	Err.Clear
End Function

Function AddEmployeeCreditors(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new creditor for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeCreditors"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para agregar la información del beneficiario."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeAddComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aEmployeeComponent(N_END_DATE_CREDITOR_EMPLOYEE) = 0 Then aEmployeeComponent(N_END_DATE_CREDITOR_EMPLOYEE) = 30000000
		If aEmployeeComponent(N_ID_CREDITOR_EMPLOYEE) = -1 Then
			sErrorDescription = "No se pudo obtener un identificador para el nuevo acreedor."
			lErrorNumber = CheckExistencyOfEmployeeCreditors(aEmployeeComponent, sErrorDescription)
			If lErrorNumber = L_ERR_NO_RECORDS Then
				lErrorNumber = GetNewIDFromTable(oADODBConnection, "EmployeesCreditorsLKP", "CreditorID", "(EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", 1, aEmployeeComponent(N_ID_CREDITOR_EMPLOYEE), sErrorDescription)
				If lErrorNumber = 0 Then
					If Not CheckEmployeeCreditorInformationConsistency(aEmployeeComponent, sErrorDescription) Then
						lErrorNumber = -1
					Else
						sErrorDescription = "No se pudo agregar la información del acreedor del empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " con fecha de inicio " & DisplayDateFromSerialNumber(aEmployeeComponent(N_START_DATE_CREDITOR_EMPLOYEE), -1, -1, -1) & "."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesCreditorsLKP (EmployeeID, CreditorID, StartDate, EndDate, CreditorNumber, CreditorName, CreditorLastName, CreditorLastName2, CreditorBirthDate, CreditorTypeID, ConceptAmount, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, PaymentCenterID, StartUserID, EndUserID, Comments, Active) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_ID_CREDITOR_EMPLOYEE) & ", " & aEmployeeComponent(N_START_DATE_CREDITOR_EMPLOYEE) & ", " & aEmployeeComponent(N_END_DATE_CREDITOR_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_CREDITOR_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_NAME_CREDITOR_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_LAST_NAME_CREDITOR_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_LAST_NAME2_CREDITOR_EMPLOYEE), "'", "´") & "', " & aEmployeeComponent(N_BIRTH_DATE_CREDITOR_EMPLOYEE) & ", " & aEmployeeComponent(N_CREDITOR_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", " &  aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE)  & ", " & aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE)  & ", " &  aEmployeeComponent(N_PAYMENT_CENTER_ID_CREDITOR_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", -1, '" & Replace(aEmployeeComponent(S_COMMENTS_CREDITOR_EMPLOYEE), "'", "´") & "', 0)", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					End If
				End If
			Else
				lErrorNumber = -1
				sErrorDescription = "Ya existe un acreedor con el mismo número que el indicado."
			End If
		End If
	End If

	Set oRecordset = Nothing
	AddEmployeeCreditors = lErrorNumber
	Err.Clear
End Function

Function AddEmployeeDocument(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new creditor for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeDocument"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sSign1
	Dim sSign2

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para agregar la información de la solicitud de hoja única de servicio."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeAddComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		'If aEmployeeComponent(N_END_DATE_CREDITOR_EMPLOYEE) = 0 Then aEmployeeComponent(N_END_DATE_CREDITOR_EMPLOYEE) = 30000000
		If aEmployeeComponent(N_EMPLOYEE_DOCUMENT_ID) = -1 Then
			sErrorDescription = "No se pudo obtener un identificador para el nuevo acreedor."
			lErrorNumber = CheckExistencyOfEmployeeDocument(aEmployeeComponent, sErrorDescription)
			If lErrorNumber = L_ERR_NO_RECORDS Then
				lErrorNumber = GetNewIDFromTable(oADODBConnection, "EmployeesDocs", "RecordID", "", 1, aEmployeeComponent(N_EMPLOYEE_DOCUMENT_ID), sErrorDescription)
				If lErrorNumber = 0 Then
					'If Not CheckEmployeeDocumentInformationConsistency(aEmployeeComponent, sErrorDescription) Then
					If False Then
						lErrorNumber = -1
					Else
						sSign1 = GenerateRandomCharactersSecuence(176) & "="
						sSign2 = GenerateRandomHexadecimalSecuence(8) & "-" & GenerateRandomHexadecimalSecuence(4) & "-" & GenerateRandomHexadecimalSecuence(4) & "-" & GenerateRandomHexadecimalSecuence(4) & "-" & GenerateRandomHexadecimalSecuence(12)
						sErrorDescription = "No se pudo agregar la información de la solicitud de la hoja única de servicio del empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & "."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesDocs (RecordID, EmployeeID, DocumentDate, DocumentTime, Document2Date, Document2Time, Document3Date, Document3Time, DocumentNumber, DocumentTypeID, Authorizers, Authorized, Sign1, Sign2, bPrinted, UserID, Comments) Values (" & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_ID) & ", " & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_DATE) & ", " & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_TIME) & ", " & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_DATE_2) & ", " & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_TIME_2) & ", " & "1" & ", " & "1" & ", '" & Replace(aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE), "'", "´") & "', " & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_TYPE) & ", '" & aEmployeeComponent(S_EMPLOYEE_AUTHORIZERS) & "', '" & aEmployeeComponent(S_EMPLOYEE_AUTHORIZED) & "', '" & sSign1 & "', '" & sSign2 & "', 0, " & aLoginComponent(N_USER_ID_LOGIN) & ", '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					End If
				End If
			Else
				lErrorNumber = -1
				sErrorDescription = "Ya existe una solicitud de Hoja única de servicio registrada en esa fecha para el empleado indicado."
			End If
		End If
	End If

	Set oRecordset = Nothing
	AddEmployeeDocument = lErrorNumber
	Err.Clear
End Function

Function AddEmployeeGrade(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new absence for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeGrade"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim iEndDate
	Dim sEmployeeAbsenceIDs
	Dim sConceptShortName

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
	aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) =  CLng(aEmployeeComponent(N_CALIFICATION_YEAR) & "0101")
	aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 30000000

	If VerifyRequerimentsForEmployeesGrades(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
		aEmployeeComponent(B_IS_DUPLICATED_EMPLOYEE) = False
		If CheckExistencyOfEmployeeGrade(aEmployeeComponent, sErrorDescription) Then
			lErrorNumber = L_ERR_DUPLICATED_RECORD
		Else
			If lErrorNumber = 0 Then
				'If aEmployeeComponent(B_IS_DUPLICATED_EMPLOYEE) Then
				'	lErrorNumber = L_ERR_DUPLICATED_RECORD
				'	sErrorDescription = "Ya existe un registro para el empleado " & aEmployeeComponent(N_ID_EMPLOYEE) & " en el año " & DisplayDateFromSerialNumber(aEmployeeComponent(N_CALIFICATION_YEAR), -1, -1, -1) & "."
				'	Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeAddComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
				'Else
					If Not CheckEmployeeConceptInformationConsistency(aEmployeeComponent, sErrorDescription) Then
						lErrorNumber = -1
					Else
						sErrorDescription = "No se pudo guardar la información de la calificación del empleado " & aEmployeeComponent(N_ID_EMPLOYEE)
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesGrades (EmployeeID, StartDate, EndDate, PayrollID, EmployeeGrade, GradePercentage, ModifyDate, UserID, Active) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", '" & aEmployeeComponent(S_EMPLOYEE_GRADE) & "', " & aEmployeeComponent(N_EMPLOYEE_GRADE_PORCENTAGE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				'End If
			End If
		End If
	Else
		lErrorNumber = -1
	End If

	AddEmployeeGrade = lErrorNumber
	Err.Clear
End Function

Function AddEmployeeMovement(oRequest, oADODBConnection, lReasonID, aEmployeeComponent, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new movement for the employee into the database
'Inputs:  oRequest, oADODBConnection, lReasonID
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeMovement"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim iStatusReasonID
	Dim lJob1
	Dim lJob2
	Dim lStartDate
	Dim lEndDate
	Dim lEmployeeID1
	Dim lEmployeeID2
	Dim lJobID1
	Dim lJobID2
	Dim lHistoryEmployeeDate
	Dim lHistoryEndDate
	Dim bProcessed
	Dim iStatusJob
	Dim lAreaID
	Dim lPaymentCenterID
	Dim lServiceID
	Dim lJourneyID
	Dim lShiftID
	Dim sDate
	Dim lHistoryJobDate
	Dim lStatusJobID
	Dim lCompanyID
	Dim lJobID
	Dim lServiceID2
	Dim lZoneID
	Dim lPositionTypeID
	Dim lClassificationID
	Dim lGroupGradeLevelID
	Dim lIntegrationID
	Dim lJourneyID2
	Dim lShiftID2
	Dim lWorkingHours
	Dim lAreaID2
	Dim lPositionID
	Dim lLevelID
	Dim lPaymentCenterID2
	Dim lEmployeeTypeID
	Dim lRiskLevel
	Dim lMovementDate
	Dim lOwnerID1
	Dim lOwnerID2
	Dim lStatusID
	Dim lActive
	Dim lRiskAmount
	Dim sComments
	Dim lOldReasonID
	Dim oConceptsRecordset
	Dim lAppliesToID
	Dim lExtraShift1
	Dim lExtraShift2
	Dim oNextRecordset
	Dim lNextEndDate
	Dim aEmployee1Concepts
	Dim aEmployee2Concepts
	Dim iIndex
	Dim sQuery

	iStatusJob = 1
	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para agregar el movimiento al empleado."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeAddComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 0 Then aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 30000000
			If InStr(1,",12,13,14,17,18,28,68," , "," & lReasonID & "," ,vbBinaryCompare) <> 0 Then
			If Len(oRequest("EmployeeYear").Item) <> 0 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Active From EmployeesHistoryList Where (EmployeeID="& aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate <= " & CLng(oRequest("EmployeeYear").Item & oRequest("EmployeeMonth").Item & oRequest("EmployeeDay").Item) & ") And (EndDate >= " & CLng(oRequest("EmployeeYear").Item & oRequest("EmployeeMonth").Item & oRequest("EmployeeDay").Item) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Active From EmployeesHistoryList Where (EmployeeID="& aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate <= " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And (EndDate >= " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			End If
			If Not oRecordset.EOF Then
				If (CLng(oRecordset.Fields("Active").Value) = 1) Then
					lErrorNumber = -1
					sErrorDescription = "El empleado ya aparece activo en el periodo indicado, verifique la vigencia del movimiento."
				End If
			End If
		End If
'	Validación bloqueada temporalmente hasta homologación de horarios y turnos
'		If lErrorNumber = 0 Then
'			If InStr(1,",12,13,17,68,18,28,21,50,57," , "," & lReasonID & "," , vbBinaryCompare) <> 0 Then
'				sErrorDescription = "El horario del empleado no corresponde con el turno de la plaza"
'				lErrorNumber = VerifyJobJourneyForEmployeeShift(oADODBConnection, oRequest, sErrorDescription)
'			End If
'		End If
		If lErrorNumber = 0 Then
			lErrorNumber = CheckRequirementsOfEmployeeMovement(oRequest, aEmployeeComponent, lReasonID, sErrorDescription)
		End If
		If lErrorNumber = 0 Then
			lErrorNumber = AddEmployeesRequirements(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				iStatusReasonID = aEmployeeComponent(N_STATUS_REASON_ID_EMPLOYEE)
				If (Len(oRequest("Register").Item) > 0) Then
					iStatusReasonID = 2
				ElseIf (Len(oRequest("Validate").Item) > 0) Then
					iStatusReasonID = 3
				ElseIf (Len(oRequest("Authorization").Item) > 0) Then
					iStatusReasonID = 0
				ElseIf (Len(oRequest("AuthorizationFile").Item) > 0) Then
					iStatusReasonID = 0
				ElseIf Len(oRequest("SaveChanges").Item) Then
					iStatusReasonID = 1
				Else
					iStatusReasonID = 1
				End If
				sDate = Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000"))
				If iStatusReasonID <> 0 Then
					bProcessed = 2
					Select Case lReasonID
						'Case -105
						'	lErrorNumber = RemoveEmployeeConceptForValidation(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
                        Case 14 'HonoraryEmployees
							Call InitializeJobComponent(oRequest, aJobComponent)
							aEmployeeComponent(N_AREA_ID_EMPLOYEE) = aJobComponent(N_AREA_ID_JOB)
							lErrorNumber = GetZoneByArea(oADODBConnection, aEmployeeComponent, sErrorDescription)
							aJobComponent(N_ID_JOB) = aEmployeeComponent(N_ID_EMPLOYEE)
							aJobComponent(N_ID_EMPLOYEE_JOB) = aEmployeeComponent(N_ID_EMPLOYEE)
							aJobComponent(N_ID_OWNER_JOB) = -1
							aJobComponent(S_NUMBER_JOB) = Right("000000" & aEmployeeComponent(N_ID_EMPLOYEE),6)
							aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) = aJobComponent(N_COMPANY_ID_JOB)
							aJobComponent(N_ZONE_ID_JOB) = aEmployeeComponent(N_ZONE_ID_EMPLOYEE)
							aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) = aJobComponent(N_PAYMENT_CENTER_ID_JOB)
							aJobComponent(N_POSITION_ID_JOB) = L_HONORARY_POSITION_ID
							aJobComponent(N_JOB_TYPE_ID_JOB) = 4
							aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) = 3
							aJobComponent(N_POSITION_TYPE_ID_JOB) = 3
							aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) = aJobComponent(N_JOURNEY_ID_JOB)
							aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) = aJobComponent(N_SERVICE_ID_JOB)
							aJobComponent(N_START_DATE_JOB) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
							aJobComponent(N_JOB_DATE_JOB) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
							aJobComponent(N_END_DATE_JOB) = aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)
							aJobComponent(N_END_DATE_HISTORY_JOB) = aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)
							aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) = aJobComponent(N_SHIFT_ID_JOB)
							aJobComponent(N_STATUS_ID_JOB) = 1
							aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) = -1
							aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) = 8
							aJobComponent(D_WORKING_HOURS_JOB) = aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE)

							sErrorDescription = "No se pudo guardar la información de la nueva plaza de honorarios."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Jobs Where (JobID=" & aJobComponent(N_ID_JOB) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
								Else
									lErrorNumber = AddJob(oRequest, oADODBConnection, aJobComponent, True, sErrorDescription)
								End If
							End If
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo obtener la información del registro."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusID, Active From StatusEmployees Where (ReasonID=" & lReasonID & ") And (StatusReasonID=" & iStatusReasonID & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = CLng(oRecordset.Fields("StatusID").Value)
										aEmployeeComponent(N_ACTIVE_EMPLOYEE) = CLng(oRecordset.Fields("Active").Value)
										lStatusID = CLng(oRecordset.Fields("StatusID").Value)
										lActive = CLng(oRecordset.Fields("Active").Value)
									End If
								End If
								bProcessed = 2
								sErrorDescription = "No se pudo obtener la información del registro."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, ReasonID From EmployeesHistoryList Where (bProcessed=" & bProcessed & ") And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If oRecordset.EOF Then
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & "," & bProcessed & ", '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudo actualizar la información del empleado."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									Else
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", EndDate=" & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", ZoneID=" & aJobComponent(N_ZONE_ID_JOB) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aJobComponent(N_SHIFT_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", AreaID=" & aJobComponent(N_AREA_ID_JOB) & ", PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", RiskLevel=" & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", PayrollDate=" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", Comments='" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (bProcessed=" & bProcessed & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudo actualizar la información del empleado."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									End If
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', EmployeeAccessKey='" & Replace(aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE), "'", "") & "', EmployeePassword='" & Replace(aEmployeeComponent(S_PASSWORD_EMPLOYEE), "'", "") & "', EmployeeName='" & Replace(UCase(aEmployeeComponent(S_NAME_EMPLOYEE)), "'", "´") & "', EmployeeLastName='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME_EMPLOYEE)), "'", "´") & "', EmployeeLastName2='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE)), "'", "´") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aJobComponent(N_SHIFT_ID_JOB) & ", StartHour1=" & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & ", EndHour1=" & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & ", StartHour2=" & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & ", EndHour2=" & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & ", StartHour3=" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", EndHour3=" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", EmployeeEmail='" & Replace(aEmployeeComponent(S_EMAIL_EMPLOYEE), "'", "") & "', SocialSecurityNumber='" & Replace(aEmployeeComponent(S_SSN_EMPLOYEE), "'", "") & "', BirthYear=" & CInt(Left(sDate, Len("0000"))) & ", BirthMonth=" & CInt(Mid(sDate, Len("00000"), Len("00"))) & ", BirthDay=" & CInt(Mid(sDate, Len("0000000"), Len("00"))) & ", BirthDate=" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) & ", StartDate=" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & ", StartDate2=" & aEmployeeComponent(N_START_DATE2_EMPLOYEE) & ", CountryID=" & aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) & ", RFC='" & Replace(UCase(aEmployeeComponent(S_RFC_EMPLOYEE)), "'", "") & "', CURP='" & Replace(UCase(aEmployeeComponent(S_CURP_EMPLOYEE)), "'", "") & "', GenderID=" & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & ", MaritalStatusID=" & aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									End If
								End If
							End If
							If lErrorNumber = 0 Then
								aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 13
								aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
								aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)
								aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) = 0
								aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 1
								aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 3
								aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) = 0
								aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) = 1
								aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) = 0
								aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) = 1
								aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = 11
								aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) = 1
								aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
								aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = ""
								aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = "Trámite realizado por ingreso"
								aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
								lErrorNumber = ModifyEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							End If
						Case 26 'Change jobs between two employees
							lStartDate = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
							lEndDate = aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)
							lEmployeeID1 = aEmployeeComponent(N_ID_EMPLOYEE)
							lEmployeeID2 = aEmployeeComponent(N_ID_EMPLOYEE_2)
							If lEmployeeID1 <> lEmployeeID2 Then
								aEmployeeComponent(N_ID_EMPLOYEE) = lEmployeeID1
								lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								lJobID1 = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
								aEmployeeComponent(N_ID_EMPLOYEE) = lEmployeeID2
								lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								lJobID2 = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
								aEmployeeComponent(N_ID_EMPLOYEE) = lEmployeeID1
								lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = lStartDate
								aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = lEndDate
								aEmployeeComponent(N_JOB_ID_EMPLOYEE) = lJobID2
								aJobComponent(N_ID_JOB) = lJobID2
								If aJobComponent(N_ID_JOB) <> -1 Then
									lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
									If aJobComponent(N_POSITION_TYPE_ID_JOB) <> 1 Then
										sErrorDescription = "Los cambios por permuta de plazas solo se pueden otorgar a personal con puesto de base"
										lErrorNumber = -1
									End If
								End If
							Else
								sErrorDescription = "Los cambios por permuta de plazas solo se puede realizar entre diferentes empleados."
								lErrorNumber = -1
							End If
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo obtener la información del registro."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusID, Active From StatusEmployees Where (ReasonID=" & lReasonID & ") And (StatusReasonID=" & iStatusReasonID & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = CLng(oRecordset.Fields("StatusID").Value)
										aEmployeeComponent(N_ACTIVE_EMPLOYEE) = CLng(oRecordset.Fields("Active").Value)
									End If
								End If
								sErrorDescription = "No se pudo obtener la información del registro."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, ReasonID From EmployeesHistoryList Where (bProcessed = 2) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If oRecordset.EOF Then
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & "," & bProcessed & ", '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudo actualizar la información del empleado."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									Else
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", EndDate=" & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", ZoneID=" & aJobComponent(N_ZONE_ID_JOB) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aJobComponent(N_SHIFT_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", AreaID=" & aJobComponent(N_AREA_ID_JOB) & ", PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", RiskLevel=" & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", PayrollDate=" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", Comments='" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (bProcessed=" & bProcessed & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudo actualizar la información del empleado."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									End If
								End If
								aEmployeeComponent(N_ID_EMPLOYEE) = lEmployeeID2
								lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = lStartDate
								aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = lEndDate
								aEmployeeComponent(N_JOB_ID_EMPLOYEE) = lJobID1
								aJobComponent(N_ID_JOB) = lJobID1
								If aJobComponent(N_ID_JOB) <> -1 Then
									lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
								End If
								sErrorDescription = "No se pudo obtener la información del registro."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusID, Active From StatusEmployees Where (ReasonID=" & lReasonID & ") And (StatusReasonID=" & iStatusReasonID & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = CLng(oRecordset.Fields("StatusID").Value)
										aEmployeeComponent(N_ACTIVE_EMPLOYEE) = CLng(oRecordset.Fields("Active").Value)
									End If
								End If
								sErrorDescription = "No se pudo actualizar la información del registro."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, ReasonID From EmployeesHistoryList Where (bProcessed = 2) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If oRecordset.EOF Then
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & "," & bProcessed & ", '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudo actualizar la información del empleado."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									Else
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", EndDate=" & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", ZoneID=" & aJobComponent(N_ZONE_ID_JOB) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aJobComponent(N_SHIFT_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", AreaID=" & aJobComponent(N_AREA_ID_JOB) & ", PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", RiskLevel=" & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", PayrollDate=" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", Comments='" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (bProcessed=" & bProcessed & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudo actualizar la información del empleado."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									End If
								End If
							End If

							aEmployeeComponent(N_ID_EMPLOYEE)= lEmployeeID1
							aEmployeeComponent(N_ID_EMPLOYEE_2) = lEmployeeID2
						Case 51
							sErrorDescription = "No se pudo obtener la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusID, Active From StatusEmployees Where (ReasonID=" & lReasonID & ") And (StatusReasonID=" & iStatusReasonID & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = CLng(oRecordset.Fields("StatusID").Value)
									aEmployeeComponent(N_ACTIVE_EMPLOYEE) = CLng(oRecordset.Fields("Active").Value)
								End If
							End If
							lAreaID = CLng(oRequest("AreaID").Item)
							lPaymentCenterID = CLng(oRequest("PaymentCenterID").Item)
							lServiceID = CLng(oRequest("ServiceID").Item)
							lJourneyID = CLng(oRequest("JourneyID").Item)
							lShiftID = CLng(oRequest("NewShiftID").Item)
							aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
							If aJobComponent(N_ID_JOB) <> -1 Then
								lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
							End If
							aJobComponent(N_AREA_ID_JOB) = lAreaID
							aJobComponent(N_PAYMENT_CENTER_ID_JOB) = lPaymentCenterID
							aJobComponent(N_SERVICE_ID_JOB)  = lServiceID
							aJobComponent(N_JOURNEY_ID_JOB)  = lJourneyID
							aJobComponent(N_SHIFT_ID_JOB)  = lShiftID
							sErrorDescription = "No se pudo obtener la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, ReasonID From EmployeesHistoryList Where (bProcessed=" & bProcessed & ") And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If oRecordset.EOF Then
									sErrorDescription = "No se pudo actualizar la información del empleado."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & "," & bProcessed & ", '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									End If
								Else
									sErrorDescription = "No se pudo actualizar la información del empleado."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", EndDate=" & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", ZoneID=" & aJobComponent(N_ZONE_ID_JOB) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aJobComponent(N_SHIFT_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", AreaID=" & aJobComponent(N_AREA_ID_JOB) & ", PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", RiskLevel=" & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", PayrollDate=" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", Comments='" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (bProcessed=" & bProcessed & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									End If
								End If
							End If
						Case Else
							aEmployeeComponent(N_REASON_ID_EMPLOYEE) = lReasonID
							If lReasonID = 28 Then 'Reanudación de labores
								If aEmployeeComponent(N_ACTIVE_EMPLOYEE) = 0 Then
									Select Case iStatusReasonID
										Case 1
											aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = 2
										Case 2
											aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = 3
										Case 3
											aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = 4
									End Select
								Else
									Select Case iStatusReasonID
										Case 1
											aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = 123
										Case 2
											aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = 124
										Case 3
											aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = 125
									End Select
								End If
							Else
								sErrorDescription = "No se pudo obtener la información del registro."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusID, Active From StatusEmployees Where (ReasonID=" & lReasonID & ") And (StatusReasonID=" & iStatusReasonID & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = CLng(oRecordset.Fields("StatusID").Value)
										aEmployeeComponent(N_ACTIVE_EMPLOYEE) = CLng(oRecordset.Fields("Active").Value)
										lStatusID = aEmployeeComponent(N_STATUS_ID_EMPLOYEE)
										lActive = aEmployeeComponent(N_ACTIVE_EMPLOYEE)
									End If
								End If
							End If
							aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
							If aJobComponent(N_ID_JOB) <> -1 Then
								lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
								aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)
							End If
							sErrorDescription = "No se pudo actualizar la información del registro."
							If (lReasonID = EMPLOYEES_FOR_RISK) Or (lReasonID = EMPLOYEES_ADDITIONALSHIFT) Or _
								(lReasonID = EMPLOYEES_CONCEPT_08) Or (lReasonID = EMPLOYEES_HONORARIUM_CONCEPT) Then
								lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
								If lErrorNumber = 0 Then
									If Len(oRequest("ModifyConcept").Item) > 0 Then
										lErrorNumber = GetEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
									End If
								End If
								aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = oRequest("ConceptStartYear").Item & oRequest("ConceptStartMonth").Item & oRequest("ConceptStartDay").Item
								aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = oRequest("ConceptEndYear").Item & oRequest("ConceptEndMonth").Item & oRequest("ConceptEndDay").Item
								'aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = oRequest("ConceptStartYear").Item & oRequest("ConceptStartMonth").Item & oRequest("ConceptStartDay").Item
								'aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = oRequest("ConceptEndYear").Item & oRequest("ConceptEndMonth").Item & oRequest("ConceptEndDay").Item
								aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = lStatusID
								aEmployeeComponent(N_ACTIVE_EMPLOYEE) = lActive
								If lReasonID = EMPLOYEES_FOR_RISK Then 
									If CInt(oRequest("ConceptAmount").Item) = 10 Then
										aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) = 1
									ElseIf CInt(oRequest("ConceptAmount").Item) = 20 Then
										aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) = 2 
									Else
										aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) = -1
									End If
								ElseIf (lReasonID = EMPLOYEES_ADDITIONALSHIFT) Or (lReasonID = EMPLOYEES_CONCEPT_08) Then
									aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = 3/6.5*100
								Else
									aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(oRequest("ConceptAmount").Item)
								End If
								If aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 0 Then aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 30000000
								If aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 0 Then aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 30000000
								'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, StartDate, EndDate, ConceptID, ConceptAmount From EmployeesConceptsLKP Where StartDate = " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & " And ConceptID = " & oRequest("ConceptEndDay").Item & " And EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE), "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								'If oRecordset.EOF Then
								'	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesConceptsLKP (EmployeeID, ConceptID, StartDate, EndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, RegistrationDate, ModifyDate, StartUserID, EndUserID, UploadedFileName, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ", " & aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) & ", " & aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) & ", '" & aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) & "', " & aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) & ", 0, " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", -1, '" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								'Else
								'	sErrorDescription = "El empleado ya tenía el concepto activado para el periodo indicado"
								'	lErrorNumber = -1
								'End If
								Select Case iStatusReasonID
									Case 1
                                        If Len(oRequest("ModifyConcept").Item) > 0 Then
                                            aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 1
                                            lErrorNumber = SetActiveForEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
                                        Else
                                            aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
                                            lErrorNumber = AddEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
											If lErrorNumber = 0 Then
												aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 1
												lErrorNumber = SetActiveForEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
											End If
                                        End If
                                    Case Else
                                        aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
                                        lErrorNumber = AddEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
                                End Select
							End If
							If lErrorNumber = 0 Then
							    lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, ReasonID From EmployeesHistoryList Where (bProcessed=" & bProcessed & ") And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							End If
							If lErrorNumber = 0 And  (lReasonID <> EMPLOYEES_FOR_RISK) And _
									(lReasonID <> EMPLOYEES_ADDITIONALSHIFT) And (lReasonID <> EMPLOYEES_CONCEPT_08) And _
                                    (lReasonID <> EMPLOYEES_HONORARIUM_CONCEPT) Then
								If oRecordset.EOF Then
									sErrorDescription = "No se pudo actualizar la información del empleado."
									If aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 0 Then aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 30000000
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & "," & bProcessed & ", '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", StartHour1=" & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & ", EndHour1=" & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & ", StartHour2=" & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & ", EndHour2=" & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & ", StartHour3=" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", EndHour3=" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									End If
								Else
									sErrorDescription = "No se pudo actualizar la información del empleado."
									lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
									lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
									If aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 0 Then aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 30000000
									If (lReasonID = EMPLOYEES_FOR_RISK) Or (lReasonID = EMPLOYEES_ADDITIONALSHIFT) Or _
										(lReasonID = EMPLOYEES_CONCEPT_08) Then
										lStartDate = oRequest("ConceptStartYear").Item & oRequest("ConceptStartMonth").Item & oRequest("ConceptStartDay").Item
										lEndDate = oRequest("ConceptEndYear").Item & oRequest("ConceptEndMonth").Item & oRequest("ConceptEndDay").Item
									Else
										lStartDate = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
										lEndDate = aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)
									End If
									If lEndDate = 0  Then lEndDate = 30000000
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EmployeeDate=" & lStartDate & ", EndDate=" & lEndDate & ", EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", ZoneID=" & aJobComponent(N_ZONE_ID_JOB) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aJobComponent(N_SHIFT_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", AreaID=" & aJobComponent(N_AREA_ID_JOB) & ", PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", RiskLevel=" & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", PayrollDate=" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", Comments='" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (bProcessed=" & bProcessed & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", StartHour1=" & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & ", EndHour1=" & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & ", StartHour2=" & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & ", EndHour2=" & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & ", StartHour3=" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", EndHour3=" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									End If
								End If
								oRecordset.Close
								If lErrorNumber = 0 Then
									If (lReasonID = 12) Or (lReasonID = 13) Or (lReasonID = 14) Or (lReasonID = 17) Or (lReasonID = 18) Or (lReasonID = 26) Or (lReasonID = 57) Or (lReasonID = 58) Then
										aEmployeeComponent(N_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
										aEmployeeComponent(N_START_DATE2_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
										sErrorDescription = "No se pudo modificar la información del empleado."
										If B_UPPERCASE Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set EmployeeName='" & Replace(UCase(aEmployeeComponent(S_NAME_EMPLOYEE)), "'", "´") & "', EmployeeLastName='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME_EMPLOYEE)), "'", "´") & "', EmployeeLastName2='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE)), "'", "´") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", StartHour1=" & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & ", EndHour1=" & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & ", StartHour2=" & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & ", EndHour2=" & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & ", StartHour3=" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", EndHour3=" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", EmployeeEmail='" & Replace(aEmployeeComponent(S_EMAIL_EMPLOYEE), "'", "") & "', SocialSecurityNumber='" & Replace(aEmployeeComponent(S_SSN_EMPLOYEE), "'", "") & "', BirthYear=" & CInt(Left(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("0000"))) & ", BirthMonth=" & CInt(Mid(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("00000"), Len("00"))) & ", BirthDay=" & CInt(Mid(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("0000000"), Len("00"))) & ", BirthDate=" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) & ", StartDate=" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & ", StartDate2=" & aEmployeeComponent(N_START_DATE2_EMPLOYEE) & ", CountryID=" & aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) & ", RFC='" & Replace(aEmployeeComponent(S_RFC_EMPLOYEE), "'", "") & "', CURP='" & Replace(aEmployeeComponent(S_CURP_EMPLOYEE), "'", "") & "', GenderID=" & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & ", MaritalStatusID=" & aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										Else
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set EmployeeName='" & Replace(aEmployeeComponent(S_NAME_EMPLOYEE), "'", "´") & "', EmployeeLastName='" & Replace(aEmployeeComponent(S_LAST_NAME_EMPLOYEE), "'", "´") & "', EmployeeLastName2='" & Replace(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE), "'", "´") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", StartHour1=" & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & ", EndHour1=" & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & ", StartHour2=" & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & ", EndHour2=" & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & ", StartHour3=" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", EndHour3=" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", EmployeeEmail='" & Replace(aEmployeeComponent(S_EMAIL_EMPLOYEE), "'", "") & "', SocialSecurityNumber='" & Replace(aEmployeeComponent(S_SSN_EMPLOYEE), "'", "") & "', BirthYear=" & CInt(Left(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("0000"))) & ", BirthMonth=" & CInt(Mid(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("00000"), Len("00"))) & ", BirthDay=" & CInt(Mid(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("0000000"), Len("00"))) & ", BirthDate=" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) & ", StartDate=" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & ", StartDate2=" & aEmployeeComponent(N_START_DATE2_EMPLOYEE) & ", CountryID=" & aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) & ", RFC='" & Replace(aEmployeeComponent(S_RFC_EMPLOYEE), "'", "") & "', CURP='" & Replace(aEmployeeComponent(S_CURP_EMPLOYEE), "'", "") & "', GenderID=" & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & ", MaritalStatusID=" & aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									End If
								End If
								If lErrorNumber = 0 Then
									If (lReasonID = 12) Or (lReasonID = 13) Or (lReasonID = 14) Or (lReasonID = 17) Or (lReasonID = 18) Then
										If aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) <> 0 Then
											sErrorDescription = "No se pudo obtener la información del riesgo profesional que tiene el empleado."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesRisksLKP Where EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE), "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
											If lErrorNumber = 0 Then
												sErrorDescription = "No se pudo eliminar la información del riesgo al empleado."
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesRisksLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
												If lErrorNumber = 0 Then
													sErrorDescription = "No se pudo agregar la información del riesgo al empleado."
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesRisksLKP (EmployeeID, RiskLevel) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
												End If
											End If
											aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 4
											aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
											aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)
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
											aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
											aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = ""
											aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = "Trámite realizado a través de FM1"
											sErrorDescription = "No se pudo agregar el concepto de riesgo al empleado."
											If lErrorNumber = 0 Then
												lErrorNumber = ModifyEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
											End If
										Else
											sErrorDescription = "No se pudo eliminar la información del riesgo profesional que tiene el empleado."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesRisksLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
											If lErrorNumber = 0 Then
												aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 4
												aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
												lErrorNumber = DropEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
											End If
										End If
									End If
									If lErrorNumber = 0 Then
										If (lReasonID = 12) Or (lReasonID = 13) Or (lReasonID = 14) Or (lReasonID = 17) Or (lReasonID = 18) Then
											If (aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) > 0) And (aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) > 0) Then
												sErrorDescription = "No se pudo agregar la percepción adicional del empleado."
												aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
												aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)
												aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = .4615
												aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) = 0
												aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
												aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 3
												aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) = 0
												aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) = 1
												aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) = 0
												aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) = 1
												aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = "1,5"
												aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) = 1
												aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 0
												aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = ""
												aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = "Trámite realizado a través de FM1"
												If aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) = 1 Then
													aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 7
													If lErrorNumber = 0 Then
														lErrorNumber = ModifyEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
														If lErrorNumber = 0 Then
															sErrorDescription = "No se pudo agregar el turno opcional al empleado."
															lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set StartHour3=" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", EndHour3=" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
														End If
													End If
												ElseIf (aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) = 2)  And (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 0 Or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 2 Or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 3 Or aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = 4) Then
													aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 8
													If lErrorNumber = 0 Then
														lErrorNumber = ModifyEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
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
					End Select
					If lErrorNumber = 0 Then
						If (lReasonID = 12) Or (lReasonID = 13) Or (lReasonID = 14) Or (lReasonID = 17) Or (lReasonID = 18) Or (lReasonID = 57) Then
							sErrorDescription = "No se pudo guardar la información extra del nuevo empleado."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesExtraInfo Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo modificar la información extra del empleado."
								If B_UPPERCASE Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesExtraInfo (EmployeeID, EmployeeAddress, EmployeeCity, EmployeeZipCode, StateID, CountryID, EmployeePhone, OfficePhone, OfficeExt, DocumentNumber1, DocumentNumber2, DocumentNumber3, EmployeeActivityID) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", '" & Replace(UCase(aEmployeeComponent(S_ADDRESS_EMPLOYEE)), "'", "´") & "', '" & Replace(UCase(aEmployeeComponent(S_CITY_EMPLOYEE)), "'", "´") & "', '" & Replace(aEmployeeComponent(S_ZIP_CODE_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_ADDRESS_COUNTRY_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_EMPLOYEE_PHONE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_OFFICE_PHONE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_EXT_OFFICE_EMPLOYEE), "'", "") & "', '" & Replace(UCase(aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE)), "'", "") & "', '" & Replace(UCase(aEmployeeComponent(S_DOCUMENT_NUMBER_2_EMPLOYEE)), "'", "") & "', '" & Replace(UCase(aEmployeeComponent(S_DOCUMENT_NUMBER_3_EMPLOYEE)), "'", "") & "', " & aEmployeeComponent(N_ACTIVITY_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
								Else
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesExtraInfo (EmployeeID, EmployeeAddress, EmployeeCity, EmployeeZipCode, StateID, CountryID, EmployeePhone, OfficePhone, OfficeExt, DocumentNumber1, DocumentNumber2, DocumentNumber3, EmployeeActivityID) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_ADDRESS_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_CITY_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_ZIP_CODE_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_ADDRESS_COUNTRY_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_EMPLOYEE_PHONE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_OFFICE_PHONE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_EXT_OFFICE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_DOCUMENT_NUMBER_2_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_DOCUMENT_NUMBER_3_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_ACTIVITY_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
								End If
							End If
						End If
					End If
				Else
					Select Case lReasonID
						Case 14 'HonoraryEmployees
							Call InitializeJobComponent(oRequest, aJobComponent)
							aEmployeeComponent(N_AREA_ID_EMPLOYEE) = aJobComponent(N_AREA_ID_JOB)
							lErrorNumber = GetZoneByArea(oADODBConnection, aEmployeeComponent, sErrorDescription)
							aJobComponent(N_ID_JOB) = aEmployeeComponent(N_ID_EMPLOYEE)
							aJobComponent(N_ID_EMPLOYEE_JOB) = aEmployeeComponent(N_ID_EMPLOYEE)
							aJobComponent(S_NUMBER_JOB) = Right("000000" & aEmployeeComponent(N_ID_EMPLOYEE),6)
							aJobComponent(N_ID_OWNER_JOB) = -1
							aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) = aJobComponent(N_COMPANY_ID_JOB)
							aJobComponent(N_ZONE_ID_JOB) = aEmployeeComponent(N_ZONE_ID_EMPLOYEE)
							aEmployeeComponent(N_AREA_ID_EMPLOYEE) = aJobComponent(N_AREA_ID_JOB)
							aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) = aJobComponent(N_PAYMENT_CENTER_ID_JOB)
							aJobComponent(N_POSITION_ID_JOB) = L_HONORARY_POSITION_ID
							aJobComponent(N_JOB_TYPE_ID_JOB) = 4
							aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) = 3
							aJobComponent(N_POSITION_TYPE_ID_JOB) = 3
							aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) = aJobComponent(N_JOURNEY_ID_JOB)
							aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) = aJobComponent(N_SERVICE_ID_JOB)
							aJobComponent(N_START_DATE_JOB) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
							aJobComponent(N_END_DATE_JOB) = aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)
							aJobComponent(N_JOB_DATE_JOB) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
							aJobComponent(N_END_DATE_HISTORY_JOB) = aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)
							aJobComponent(N_STATUS_ID_JOB) = 1
							aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) = -1
							aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) = 8
							aJobComponent(N_OCCUPATION_TYPE_ID_JOB) = 0
							aJobComponent(D_WORKING_HOURS_JOB) = aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE)
							sErrorDescription = "No se pudo guardar la información de la nueva plaza."
							sErrorDescription = "Error al crear la plaza para el empleado de honorarios"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Jobs Where (JobID=" & aJobComponent(N_ID_JOB) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									sErrorDescription = "No se pudo eliminar la información de la plaza."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From JobsHistoryList Where (JobID=" & aJobComponent(N_ID_JOB) & ") And (EndDate = 30000000)", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									If lErrorNumber = 0 Then
										lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
									End If
								Else
									lErrorNumber = AddJob(oRequest, oADODBConnection, aJobComponent, True, sErrorDescription)
								End If
							End If
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo obtener la información del registro."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusID, Active From StatusEmployees Where (ReasonID=0) And (StatusReasonID=" & iStatusReasonID & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = CLng(oRecordset.Fields("StatusID").Value)
										aEmployeeComponent(N_ACTIVE_EMPLOYEE) = CLng(oRecordset.Fields("Active").Value)
									End If
								End If
								sErrorDescription = "No se pudo obtener la información del registro."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, ReasonID From EmployeesHistoryList Where (bProcessed=2) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If oRecordset.EOF Then
										bProcessed = 0
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) & ", " & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & "," & bProcessed & ", '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudo actualizar la información del registro."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									Else
										bProcessed = 2
										sErrorDescription = "No se pudo actualizar la información del registro."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", EndDate=" & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", ZoneID=" & aJobComponent(N_ZONE_ID_JOB) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aJobComponent(N_SHIFT_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", AreaID=" & aJobComponent(N_AREA_ID_JOB) & ", PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", RiskLevel=" & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", PayrollDate=" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", bProcessed=0, Comments='" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (bProcessed=" & bProcessed & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudo actualizar la información del registro."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									End If
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudo actualizar la información del registro."
										aEmployeeComponent(N_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
										If B_UPPERCASE Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set EmployeeName='" & Replace(UCase(aEmployeeComponent(S_NAME_EMPLOYEE)), "'", "´") & "', EmployeeLastName='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME_EMPLOYEE)), "'", "´") & "', EmployeeLastName2='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE)), "'", "´") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", StartHour1=" & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & ", EndHour1=" & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & ", StartHour2=" & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & ", EndHour2=" & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & ", StartHour3=" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", EndHour3=" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", EmployeeEmail='" & Replace(aEmployeeComponent(S_EMAIL_EMPLOYEE), "'", "") & "', SocialSecurityNumber='" & Replace(aEmployeeComponent(S_SSN_EMPLOYEE), "'", "") & "', BirthYear=" & CInt(Left(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("0000"))) & ", BirthMonth=" & CInt(Mid(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("00000"), Len("00"))) & ", BirthDay=" & CInt(Mid(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("0000000"), Len("00"))) & ", BirthDate=" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) & ", StartDate=" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & ", StartDate2=" & aEmployeeComponent(N_START_DATE2_EMPLOYEE) & ", CountryID=" & aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) & ", RFC='" & Replace(aEmployeeComponent(S_RFC_EMPLOYEE), "'", "") & "', CURP='" & Replace(aEmployeeComponent(S_CURP_EMPLOYEE), "'", "") & "', GenderID=" & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & ", MaritalStatusID=" & aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										Else
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set EmployeeName='" & Replace(aEmployeeComponent(S_NAME_EMPLOYEE), "'", "´") & "', EmployeeLastName='" & Replace(aEmployeeComponent(S_LAST_NAME_EMPLOYEE), "'", "´") & "', EmployeeLastName2='" & Replace(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE), "'", "´") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", StartHour1=" & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & ", EndHour1=" & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & ", StartHour2=" & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & ", EndHour2=" & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & ", StartHour3=" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", EndHour3=" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", EmployeeEmail='" & Replace(aEmployeeComponent(S_EMAIL_EMPLOYEE), "'", "") & "', SocialSecurityNumber='" & Replace(aEmployeeComponent(S_SSN_EMPLOYEE), "'", "") & "', BirthYear=" & CInt(Left(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("0000"))) & ", BirthMonth=" & CInt(Mid(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("00000"), Len("00"))) & ", BirthDay=" & CInt(Mid(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("0000000"), Len("00"))) & ", BirthDate=" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) & ", StartDate=" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & ", StartDate2=" & aEmployeeComponent(N_START_DATE2_EMPLOYEE) & ", CountryID=" & aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) & ", RFC='" & Replace(aEmployeeComponent(S_RFC_EMPLOYEE), "'", "") & "', CURP='" & Replace(aEmployeeComponent(S_CURP_EMPLOYEE), "'", "") & "', GenderID=" & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & ", MaritalStatusID=" & aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									End If
								End If
							End If
							If lErrorNumber = 0 Then
								aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = 13
								aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
								aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)
								aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) = 0
								aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 1
								aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 1
								aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) = 0
								aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) = 1
								aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) = 0
								aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) = 1
								aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = -1
								aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) = 1
								aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 1
								aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = ""
								aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = "Trámite realizado por ingreso"
								lErrorNumber = ModifyEmployeeConceptSp(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							End If
						Case 26 'Change jobs between two employees
							lStartDate = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
							lEndDate = aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)
							lEmployeeID1 = aEmployeeComponent(N_ID_EMPLOYEE)
							lEmployeeID2 = aEmployeeComponent(N_ID_EMPLOYEE_2)
							aEmployeeComponent(N_ID_EMPLOYEE) = lEmployeeID1
							lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							lJobID1 = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
							aEmployeeComponent(N_ID_EMPLOYEE) = lEmployeeID2
							lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							lJobID2 = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
							aEmployeeComponent(N_ID_EMPLOYEE) = lEmployeeID1
							lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = lStartDate
							aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = lEndDate
							aEmployeeComponent(N_JOB_ID_EMPLOYEE) = lJobID2
							aJobComponent(N_ID_JOB) = lJobID2
							If aJobComponent(N_ID_JOB) <> -1 Then
								lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
							End If
							sErrorDescription = "No se pudo obtener la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusJob3, StatusEmployeeID, ActiveEmployeeID From Reasons Where ReasonID=" & lReasonID, "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							If Not oRecordset.EOF Then
								aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = CLng(oRecordset.Fields("StatusEmployeeID").Value)
								aEmployeeComponent(N_ACTIVE_EMPLOYEE) = CLng(oRecordset.Fields("ActiveEmployeeID").Value)
							End If
							sErrorDescription = "No se pudo obtener la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, ReasonID From EmployeesHistoryList Where (bProcessed = 2) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ") Order By EmployeeDate Desc", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If oRecordset.EOF Then
										bProcessed = 0
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & "," & bProcessed & ", '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudo actualizar la información del empleado."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
										If lErrorNumber = 0 Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeDate, EndDate From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") Order By EmployeeDate Desc", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
											If Not oRecordset.EOF Then
												oRecordset.MoveNext
												If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EndDate = " & AddDaysToSerialDate(lStartDate, -1) & " Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " And EndDate = 30000000 And EmployeeDate = " & oRecordset.Fields("EmployeeDate").Value, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
												End If
											End If
										End If
									Else
										bProcessed = 2
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", EndDate=" & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", ZoneID=" & aJobComponent(N_ZONE_ID_JOB) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aJobComponent(N_SHIFT_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", AreaID=" & aJobComponent(N_AREA_ID_JOB) & ", PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", RiskLevel=" & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", PayrollDate=" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", bProcessed=0, Comments='" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (bProcessed=" & bProcessed & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudo actualizar la información del empleado."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									End If
								End If
							If lErrorNumber = 0 Then
								aEmployeeComponent(N_ID_EMPLOYEE) = lEmployeeID2
								lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = lStartDate
								aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = lEndDate
								aEmployeeComponent(N_JOB_ID_EMPLOYEE) = lJobID1
								aJobComponent(N_ID_JOB) = lJobID1
								If aJobComponent(N_ID_JOB) <> -1 Then
									lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
								End If
								sErrorDescription = "No se pudo obtener la información del registro."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusJob3, StatusEmployeeID, ActiveEmployeeID From Reasons Where ReasonID=" & lReasonID, "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								If Not oRecordset.EOF Then
									aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = CLng(oRecordset.Fields("StatusEmployeeID").Value)
									aEmployeeComponent(N_ACTIVE_EMPLOYEE) = CLng(oRecordset.Fields("ActiveEmployeeID").Value)
								End If
								sErrorDescription = "No se pudo obtener la información del registro."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, ReasonID From EmployeesHistoryList Where (bProcessed = 2) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ") Order By EmployeeDate Desc", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If oRecordset.EOF Then
										bProcessed = 0
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & "," & bProcessed & ", '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudo actualizar la información del empleado."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
										If lErrorNumber = 0 Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeDate, EndDate From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") Order By EmployeeDate Desc", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
											If Not oRecordset.EOF Then
												oRecordset.MoveNext
												If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EndDate = " & AddDaysToSerialDate(lStartDate, -1) & " Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " And EndDate = 30000000 And EmployeeDate = " & oRecordset.Fields("EmployeeDate").Value, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
												End If
											End If
										End If
									Else
										bProcessed = 2
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", EndDate=" & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", ZoneID=" & aJobComponent(N_ZONE_ID_JOB) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aJobComponent(N_SHIFT_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", AreaID=" & aJobComponent(N_AREA_ID_JOB) & ", PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", RiskLevel=" & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", PayrollDate=" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", bProcessed=0 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (bProcessed=" & bProcessed & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudo actualizar la información del empleado."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									End If
								End If
								aEmployeeComponent(N_ID_EMPLOYEE)= lEmployeeID1
								aEmployeeComponent(N_ID_EMPLOYEE_2) = lEmployeeID2
							End If
						Case 51
							sErrorDescription = "No se pudo obtener la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusJob3, StatusEmployeeID, ActiveEmployeeID From Reasons Where ReasonID=" & lReasonID, "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							If Not oRecordset.EOF Then
								aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = CLng(oRecordset.Fields("StatusEmployeeID").Value)
								aEmployeeComponent(N_ACTIVE_EMPLOYEE) = CLng(oRecordset.Fields("ActiveEmployeeID").Value)
							End If
							lAreaID = CLng(oRequest("AreaID").Item)
							lPaymentCenterID = CLng(oRequest("PaymentCenterID").Item)
							lServiceID = CLng(oRequest("ServiceID").Item)
							lJourneyID = CLng(oRequest("JourneyID").Item)
							lShiftID = CLng(oRequest("ShiftID").Item)
							aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
							If aJobComponent(N_ID_JOB) <> -1 Then
								lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
							End If
							aJobComponent(N_AREA_ID_JOB) = lAreaID
							aJobComponent(N_PAYMENT_CENTER_ID_JOB) = lPaymentCenterID
							aJobComponent(N_SERVICE_ID_JOB)  = lServiceID
							aJobComponent(N_JOURNEY_ID_JOB)  = lJourneyID
							aJobComponent(N_SHIFT_ID_JOB)  = lShiftID
							sErrorDescription = "No se pudo obtener la información del registro."
							If lReasonID = 51 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, ReasonID From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							Else
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, ReasonID From EmployeesHistoryList Where (bProcessed=2) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							End If
							If lErrorNumber = 0 Then
								If oRecordset.EOF Then
									sErrorDescription = "No se pudo actualizar la información del empleado."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", 0, '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", ShiftID=" & lShiftID & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									End If
								Else
									sErrorDescription = "No se pudo actualizar la información del empleado."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", EndDate=" & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", ZoneID=" & aJobComponent(N_ZONE_ID_JOB) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aJobComponent(N_SHIFT_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", AreaID=" & aJobComponent(N_AREA_ID_JOB) & ", PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", RiskLevel=" & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", PayrollDate=" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", bProcessed=0, Comments='" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", ShiftID=" & lShiftID & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									End If
								End If
							End If
							sErrorDescription = "No se pudo modificar la información del empleado."
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo obtener la información del registro."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeDate, EndDate From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID <> 0) And (ReasonID Not In (54,55)) Order By EmployeeDate Desc", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								If Not oRecordset.EOF Then
									oRecordset.MoveNext
									If Not oRecordset.EOF Then
										lHistoryEmployeeDate = CLng(oRecordset.Fields("EmployeeDate").Value)
										lHistoryEndDate = CLng(oRecordset.Fields("EndDate").Value)
										If CLng(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)) < lHistoryEndDate Then
											sErrorDescription = "No se pudo actualizar la información del empleado al aplicar el movimiento"
											If (AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), -1) > lHistoryEmployeeDate) Then
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EndDate=" & AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), -1) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & lHistoryEmployeeDate & ") And (EndDate=" & lHistoryEndDate & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
											Else
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(EndDate) EndDate from EmployeesHistoryList Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EndDate < 1000) And (EmployeeDate = " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oNextRecordset)
												If Not(IsNull(oNextRecordset.Fields("EndDate").Value)) Then
													If (CLng(oNextRecordset.Fields("EndDate").Value) > 0) Then
														lNextEndDate = oNextRecordset.Fields("EndDate").Value + 1
													Else
														lNextEndDate = 0
													End If
												Else
													lNextEndDate = 0
												End If
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EndDate=" & lNextEndDate & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & lHistoryEmployeeDate & ") And (EndDate=" & lHistoryEndDate & ") And (ReasonID <> " & lReasonID & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
											End If
										End If
										If lErrorNumber = 0 Then
											lErrorNumber = RemoveEmployeeReasonForRejection(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
										End If
									End If
								End If
							End If
							If lErrorNumber = 0 And iStatusReasonID = 0 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ServiceID From Employees Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) = oRecordset.Fields("ServiceID").Value
								If aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) <> CInt(oRequest("ServiceID").Item) Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, ReasonID From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=54) And (EmployeeDate = " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										If oRecordset.EOF Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", 0, '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ID_JOB) & ", " & CInt(oRequest("ServiceID").Item) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", 54, " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", 0, '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										Else
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", EndDate=" & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & CInt(oRequest("ServiceID").Item) & ", ZoneID=" & aJobComponent(N_ZONE_ID_JOB) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aJobComponent(N_SHIFT_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", AreaID=" & aJobComponent(N_AREA_ID_JOB) & ", PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", RiskLevel=" & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", PayrollDate=" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", bProcessed=0, Comments='" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (bProcessed=2) And (ReasonID=54)", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ServiceID=" & CInt(oRequest("ServiceID").Item) & " Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									End If
								End If
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JourneyID From Employees Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) = oRequest.Fields("JourneyID").Value
								If aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) <> CInt(oRequest("JourneyID").Item) Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, ReasonID From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=55) And (EmployeeDate = " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										If oRecordset.EOF Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", 0, '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ID_JOB) & ", " & CInt(oRequest("ServiceID").Item) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", 55, " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", 0, '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										Else
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", EndDate=" & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & CInt(oRequest("serviceID").Item) & ", ZoneID=" & aJobComponent(N_ZONE_ID_JOB) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aJobComponent(N_SHIFT_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", AreaID=" & aJobComponent(N_AREA_ID_JOB) & ", PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", RiskLevel=" & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", PayrollDate=" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", bProcessed=0, Comments='" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (bProcessed=2) And (ReasonID=55)", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set JourneyID=" & CInt(oRequest("JourneyID").Item) & " Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									End If
								End If
							End If
							If lErrorNumber = 0 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID From EmployeesHistoryList Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " And (((EmployeeDate >= " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And (EmployeeDate <= " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ")) Or ((EndDate >= " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And (EndDate <= " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & "))) And EndDate <> 0", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If Not oRecordset.EOF Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set ServiceID = " & lServiceID & ", PaymentCenterID = " & lPaymentCenterID & ", AreaID = " & lAreaID & ", JourneyID = " & lJourneyID & " Where EmployeeID = 16660 And (((EmployeeDate >= " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And (EmployeeDate <= " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ")) Or ((EndDate >= " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And (EndDate <= " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & "))) And EndDate <> 0", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									If lErrorNumber <> 0 Then
										sErrorDescription = "No se pudo actualizar el historial con la información de la nueva adscripción"
									End If
								End If
							End If
						Case Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ReasonID From Reasons Where StatusEmployeeID = " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE), "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									lOldReasonID = oRecordset.Fields("ReasonID").Value
								End If
							End If
							sErrorDescription = "No se pudo obtener la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusJob3, StatusEmployeeID, ActiveEmployeeID From Reasons Where ReasonID=" & lReasonID, "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							iStatusJob = 1
							If Not oRecordset.EOF Then
								If CLng(oRecordset.Fields("StatusJob3").Value) <> -2 Then
									iStatusJob = CLng(oRecordset.Fields("StatusJob3").Value)
								End If
								aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = CLng(oRecordset.Fields("StatusEmployeeID").Value)
								aEmployeeComponent(N_ACTIVE_EMPLOYEE) = CLng(oRecordset.Fields("ActiveEmployeeID").Value)
							End If
							If (aEmployeeComponent(N_JOB_ID_EMPLOYEE) = -3) Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobId From JobsHistoryList Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " Order By EndDate Desc", "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, StatusID, JobDate, EndDate From JobsHistoryList Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " Order By EndDate Desc", "EmployeeDisplayFormsComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										If CInt(oRecordset.Fields("StatusID").Value) = 1 Then
											If CLng(oRecordset.Fields("EmployeeID").Value) <> aEmployeeComponent(N_ID_EMPLOYEE) Then
												If oRecordset.Fields("EndDate") > CLng(oRequest("EmployeeYear").Item & oRequest("EmployeeMonth").Item & oRequest("EmployeeDay").Item) Then
													lErrorNumber = -1
													sErrorDescription = "La última plaza que el empleado ocupó no está vacante actualmente."
												End If
											End If
										End If
									End If
								End If
								aEmployeeComponent(N_JOB_ID_EMPLOYEE) = oRecordset.Fields("JobID").Value
							End If
							If lErrorNumber = 0 Then
								aJobComponent(N_ID_JOB) = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
								If aJobComponent(N_ID_JOB) <> -1 Then
									lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
									aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) = aJobComponent(N_POSITION_TYPE_ID_JOB)
									aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)
								End If
								bProcessed = 2
								sErrorDescription = "No se pudo obtener la información del registro."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, ReasonID From EmployeesHistoryList Where (bProcessed=" & bProcessed & ") And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							End If
							If lErrorNumber = 0 Then
								If oRecordset.EOF Then
									bProcessed = 0
									sErrorDescription = "No se pudo actualizar la información del empleado."
									Select Case lReasonID
										Case 1, 2, 3, 4, 5, 6, 7, 8, 10, 62, 63, 66, 78, 79, 80
											aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), 1)
											If lReasonID = 66 Then 
												aEmployeeComponent(S_COMMENTS_EMPLOYEE) = oRequest("DropReasonName").Item
												aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = oRequest("EmployeeComments").Item
												If (Len(aEmployeeComponent(S_COMMENTS_EMPLOYEE)) > 0) And (Len(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE)) > 0) Then
													sComments = Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "Þ" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "")
												ElseIf Len(aEmployeeComponent(S_COMMENTS_EMPLOYEE)) > 0 Then
													sComments = Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "")
												Else
													sComments = Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "")
												End If
											End If
										Case Else
											sComments = Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "")
									End Select

									lCompanyID = aJobComponent(N_COMPANY_ID_JOB) 
									If Len(lCompanyID) = 0 Then lCompanyID = -1
									lJobID = aEmployeeComponent(N_JOB_ID_EMPLOYEE)
									If Len(lJobID) = 0 Then lJobID = -1
									lServiceID2 = aJobComponent(N_SERVICE_ID_JOB)
									If Len(lServiceID2) = 0 Then lServiceID2 = 1
									lZoneID = aJobComponent(N_ZONE_ID_JOB)
									If Len(lZoneID) = 0 Then lZoneID = -1
									lPositionTypeID = aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE)
									If Len(lPositionTypeID) = 0 Then lPositionTypeID = -1
									lClassificationID = aJobComponent(N_CLASSIFICATION_ID_JOB)
									If Len(lClassificationID) = 0 Then lClassificationID = -1
									lGroupGradeLevelID = aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB)
									If Len(lGroupGradeLevelID) = 0 Then lGroupGradeLevelID = -1
									lIntegrationID = aJobComponent(N_INTEGRATION_ID_JOB)
									If Len(lIntegrationID) = 0 Then lIntegrationID = -1
									lJourneyID2 = aJobComponent(N_JOURNEY_ID_JOB)
									If Len(lJourneyID2) = 0 Then lJourneyID2 = -1
									If Len(oRequest("ShiftID").Item) > 0 Then
										lShiftID = CLng(oRequest("ShiftID").Item)
									Else
										lShiftID = aJobComponent(N_SHIFT_ID_JOB)
									End If
									If Len(lShiftID2) = 0 Then 
										If Len(oRequest("ShiftID").Item) > 0 Then
											lShiftID2 = CLng(oRequest("ShiftID").Item)
										Else
											lShiftID2 = -1
										End If
									End If
									lWorkingHours = aJobComponent(D_WORKING_HOURS_JOB)
									If Len(lWorkingHours) = 0 Then lWorkingHours = 0
									lAreaID2 = aJobComponent(N_AREA_ID_JOB)
									If Len(lAreaID2) = 0 Then lAreaID2 = -1
									lPositionID = aJobComponent(N_POSITION_ID_JOB)
									If Len(lPositionID) = 0 Then lPositionID = -1
									lLevelID = aJobComponent(N_LEVEL_ID_JOB)
									If Len(lLevelID) = 0 Then lLevelID = -1
									lPaymentCenterID2 = aJobComponent(N_PAYMENT_CENTER_ID_JOB)
									If Len(lPaymentCenterID2) = 0 Then lPaymentCenterID2 = -1
									If Len(oRequest("RiskLevel").Item) > 0 Then
										lRiskLevel = CInt(oRequest("RiskLevel").Item)
									Else
										lRiskLevel = aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE)
									End If
									If Len(lRiskLevel) = 0 Then lRiskLevel = -1
									lEmployeeTypeID = aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)
									If Len(lEmployeeTypeID) = 0 Then lEmployeeTypeID = -1
									If Len(oRequest("StartHour3").Item) > 0 Then
										lExtraShift1 = CLng(oRequest("StartHour3").Item)
									Else
										lExtraShift1 = 0
									End If
									If Len(oRequest("EndHour3").Item) > 0 Then
										lExtraShift2 = CLng(oRequest("EndHour3").Item)
									Else
										lExtraShift2 = 0
									End If

									If lReasonID <> EMPLOYEES_FOR_RISK And lReasonID <> EMPLOYEES_ADDITIONALSHIFT And _
										lReasonID <> EMPLOYEES_CONCEPT_08 Then
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', " & lCompanyID & ", " & lJobID & ", " & lServiceID2 & ", " & lZoneID & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & lPositionTypeID & ", " & lClassificationID & ", " & lGroupGradeLevelID & ", " & lIntegrationID & ", " & lJourneyID2 & ", " & lShiftID2 & ", " & lWorkingHours & ", " & lAreaID2 & ", " & lPositionID & ", " & lLevelID & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & lPaymentCenterID2 & ", " & lRiskLevel * 10 & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & "," & bProcessed & ", '" & Replace(sComments, "'", "") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									End If
									If (aEmployeeComponent(N_REASON_ID_EMPLOYEE) >= 29 And aEmployeeComponent(N_REASON_ID_EMPLOYEE) <= 36) Then aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 28
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", EmployeeTypeID=" & lEmployeeTypeID & ", ShiftID= " & lShiftID & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									End If
								Else
									bProcessed = 2
									sErrorDescription = "No se pudo actualizar la información del empleado."
									aEmployeeComponent(N_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
									Select Case lReasonID
										Case 1, 2, 3, 4, 5, 6, 7, 8, 10, 62, 63, 66, 78, 79, 80, 81
											aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = CLng(AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), 1))
									End Select
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", EndDate=" & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", EmployeeNumber='" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", ZoneID=" & aJobComponent(N_ZONE_ID_JOB) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aJobComponent(N_SHIFT_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", AreaID=" & aJobComponent(N_AREA_ID_JOB) & ", PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", RiskLevel=" & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", PayrollDate=" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & ", bProcessed=0, Comments='" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "' Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (bProcessed=" & bProcessed & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudo actualizar la información del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									End If
								End If
								oRecordset.Close
								sErrorDescription = "No se pudo modificar la información del empleado."
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo obtener la información del registro."
									If ((lReasonID = 13) Or (lReasonID = 14) Or (lReasonID = 17) Or (lReasonID = 18)) Then
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeDate, EndDate, ReasonID, StatusID From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID <> 0) And (ReasonID <>" & lReasonID & ") And (EndDate <> 0) And (ReasonID <> 57) Order By EmployeeDate Desc", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									Else
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeDate, EndDate, ReasonID, StatusID From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID <> 0) And (EndDate <> 0) And (ReasonID <> 57) Order By EmployeeDate Desc", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									End If
									If lErrorNumber = 0 Then
										If Not oRecordset.EOF Then
											lHistoryEmployeeDate = CLng(oRecordset.Fields("EmployeeDate").Value)
											lHistoryEndDate = CLng(oRecordset.Fields("EndDate").Value)
											If CLng(oRecordset.Fields("ReasonID").Value) = 28 Then
												If CLng(oRecordset.Fields("StatusID").Value) = 58 Then
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & lHistoryEmployeeDate & ") And (EndDate=" & lHistoryEndDate & ") And (ReasonID=28)", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
													oRecordset.MoveNext
												End If
											ElseIf (lReasonID <> 1) And (lReasonID <> 2) And (lReasonID <> 3) And (lReasonID <> 4) And (lReasonID <> 5) And (lReasonID <> 6) And (lReasonID <> 8) And (lReasonID <> 10) And (lReasonID <> 62) And (lReasonID <> 63) And (lReasonID <> 66) And (lReasonID <> 78) And (lReasonID <> 79) And (lReasonID <> 80) And (lReasonID <> 81) Then
												Select Case CLng(oRecordset.Fields("ReasonID").Value)
													Case 1, 2, 3, 4, 5, 6, 7, 8, 10, 62, 63, 66, 78, 79, 80, 81
														If CLng(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)) = lHistoryEmployeeDate Then
															lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(EndDate) EndDate from EmployeesHistoryList Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EndDate < 1000) And (EmployeeDate = " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oNextRecordset)
															If Not(IsNull(oNextRecordset.Fields("EndDate").Value)) Then
																If (CLng(oNextRecordset.Fields("EndDate").Value) > 0) Then
																	lNextEndDate = oNextRecordset.Fields("EndDate").Value + 1
																Else
																	lNextEndDate = 0
																End If
															Else
																lNextEndDate = 0
															End If
															lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EndDate=" & lNextEndDate & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & lHistoryEmployeeDate & ") And (EndDate=" & lHistoryEndDate & ") And (ReasonID=" & CLng(oRecordset.Fields("ReasonID").Value) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
														Else
															lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EndDate=" & AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), -1) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & lHistoryEmployeeDate & ") And (EndDate=" & lHistoryEndDate & ") And (ReasonID=" & CLng(oRecordset.Fields("ReasonID").Value) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
														End If
												End Select
											End If
											If (lErrorNumber = 0) And (lReasonID = 17) Then
												lHistoryEmployeeDate = CLng(AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), -1))
												If CLng(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)) = lHistoryEmployeeDate Then
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(EndDate) EndDate from EmployeesHistoryList Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EndDate < 1000) And (EmployeeDate = " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oNextRecordset)
													If (CLng(oNextRecordset.Fields("EndDate").Value) > 0) Then
														lNextEndDate = oNextRecordset.Fields("EndDate").Value + 1
													Else
														lNextEndDate = 0
													End If
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EndDate=" & lNextEndDate & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & oRecordset.Fields("EmployeeDate").Value & ") And (EndDate=" & lHistoryEndDate & ") And (ReasonID=" & CLng(oRecordset.Fields("ReasonID").Value) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
												Else
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EndDate=" & lHistoryEmployeeDate & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & oRecordset.Fields("EmployeeDate").Value & ") And (EndDate=" & lHistoryEndDate & ") And (ReasonID=" & CLng(oRecordset.Fields("ReasonID").Value) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
												End If
											End If
											If (lErrorNumber = 0) And (lReasonID <> EMPLOYEES_FOR_RISK) And (lReasonID <> EMPLOYEES_ADDITIONALSHIFT) And (lReasonID <> EMPLOYEES_CONCEPT_08) Then
												oRecordset.MoveNext
												If Not oRecordset.EOF Then
													lHistoryEmployeeDate = CLng(oRecordset.Fields("EmployeeDate").Value)
													lHistoryEndDate = CLng(oRecordset.Fields("EndDate").Value)
													If CLng(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)) < lHistoryEndDate Then
														If lHistoryEmployeeDate >= AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), -1) Then
															lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EndDate=0 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & lHistoryEmployeeDate & ") And (EndDate=" & lHistoryEndDate & ") And (ReasonID <> " & lReasonID &")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
														Else
															lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set EndDate=" & AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), -1) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & lHistoryEmployeeDate & ") And (EndDate=" & lHistoryEndDate & ") And (ReasonID <> " & lReasonID &")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
														End If
													End If
												End If
											End If
										End If
									End If
								End If
								If lErrorNumber = 0 Then
									lErrorNumber = RemoveEmployeeReasonForRejection(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
								End If
								If lErrorNumber = 0 Then
									If (lReasonID = 12) Or (lReasonID = 13) Or (lReasonID = 14) Or (lReasonID = 17) Or (lReasonID = 18) Or (lReasonID = 21) Or (lReasonID = 28) Or (lReasonID = 50) Or (lReasonID = 51) Then
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
										'Else
										'	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=4) And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
										End If
									End If
								End If
								If lErrorNumber = 0 Then
									If (lReasonID = 12) Or (lReasonID = 13) Or (lReasonID = 14) Or (lReasonID = 17) Or (lReasonID = 18) Or (lReasonID = 21) Or (lReasonID = 28) Or (lReasonID = 50) Or (lReasonID = 51) Then
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
										'Else
										'	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID In(7,8)) And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
										End If
									End If
								End If
							End If
					End Select
					If lErrorNumber = 0 Then
						If lErrorNumber = 0 Then
							If iStatusJob <> - 2 Then
								aJobComponent(N_STATUS_ID_JOB) = iStatusJob
							End If
							If lErrorNumber = 0 Then
								aJobComponent(N_JOB_DATE_JOB) = CLng(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE))
								aJobComponent(N_END_DATE_HISTORY_JOB) = CLng(aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE))
								aJobComponent(N_ID_EMPLOYEE_JOB) = CLng(aEmployeeComponent(N_ID_EMPLOYEE))
								aJobComponent(N_JOB_TYPE_ID_JOB) = 1
								Select Case lReasonID
									Case 1, 2, 3, 4, 5, 6, 7, 8, 10, 62, 63, 78, 79, 80, 81
										aJobComponent(N_JOB_DATE_JOB) = CLng(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE))
										If (lReasonId <> 10) Then
											aJobComponent(N_ID_OWNER_JOB) = -1
										End If
										If aJobComponent(N_STATUS_ID_JOB) <> 7 Then
											If (lReasonId = 10) And (aJobComponent(N_ID_OWNER_JOB) <> -1) Then
												aJobComponent(N_STATUS_ID_JOB) = 4
											Else
												aJobComponent(N_STATUS_ID_JOB) = 2
											End If
										End If
										aJobComponent(N_ID_EMPLOYEE_JOB) = -1
										aJobComponent(N_ACTIVE_JOB) = 1
										sErrorDescription = "No se pudo actualizar la información de la plaza."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set EndDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EndDate>=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And ConceptID Not In (" & CONCEPTS_NOT_EXPIRE & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
										If lErrorNumber = 0 Then
											aJobComponent(N_END_DATE_HISTORY_JOB) = CLng(aJobComponent(N_END_DATE_JOB))
											lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
										End If
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudo actualizar la información del empleado."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", Active=0 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									Case 66
										aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = CLng(AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), -1))
										sErrorDescription = "No se pudo actualizar la información de la plaza."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set EndDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EndDate>=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And ConceptID Not In (" & CONCEPTS_NOT_EXPIRE & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
										If lErrorNumber = 0 Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Jobs Set StatusID=2, EndDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & " Where (JobID=" & aJobComponent(N_ID_JOB) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
											If lErrorNumber = 0 Then
												sErrorDescription = "No se pudo actualizar la información de la plaza."
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update JobsHistoryList Set EndDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & " Where (JobID=" & aJobComponent(N_ID_JOB) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
												If lErrorNumber = 0 Then
													sErrorDescription = "No se pudo guardar la información histórica de la plaza."
													aJobComponent(N_STATUS_ID_JOB) = 2
													aJobComponent(N_JOB_DATE_JOB) = AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE), 1)
													aJobComponent(N_END_DATE_HISTORY_JOB) = 30000000
													lErrorNumber = ModifyJobHistoryList(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
													'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_JOB_DATE_JOB) & ", " & aJobComponent(N_END_DATE_HISTORY_JOB) & ", " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
												End If
											End If
										End If
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudo actualizar la información del empleado."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", Active=0 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									Case 14
										If lErrorNumber = 0 Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Jobs Set OwnerID=" & aEmployeeComponent(N_ID_EMPLOYEE) & " Where (JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									Case 26
										If lErrorNumber = 0 Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From Jobs Where (JobID=" & lJobID1 & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
											lOwnerID1 = oRecordset.Fields("OwnerID").Value
										End If
										If lErrorNumber = 0 Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From Jobs Where (JobID=" & lJobID2 & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
											lOwnerID2 = oRecordset.Fields("OwnerID").Value
										End If
										If lErrorNumber = 0 Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Jobs Set OwnerID=" & lOwnerID1 & " Where (JobID=" & lJobID2 & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
										If lErrorNumber = 0 Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Jobs Set OwnerID=" & lOwnerID2 & " Where (JobID=" & lJobID1 & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
										If lErrorNumber = 0 Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobID, OwnerID, JobNumber, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, WorkingHours, StartDate, EndDate, StatusID, ModifyDate From Jobs Where JobID =" & lJobID1, "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										End If
										If lErrorNumber = 0 Then
											aJobComponent(N_ID_JOB) = lJobID2
											aJobComponent(N_ID_EMPLOYEE_JOB) = lEmployeeID1
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsHistoryList (JobID, OwnerID, JobDate, EndDate, CompanyID, EmployeeID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, WorkingHours, StatusID, UserID, ModifyDate) Values (" & lJobID2 & "," & lEmployeeID1 & "," & lStartDate & "," & lEndDate & "," & oRecordset.Fields("CompanyID").Value & "," & lEmployeeID1 & "," & oRecordset.Fields("ZoneID").Value & "," & oRecordset.Fields("AreaID").Value & "," & oRecordset.Fields("PaymentCenterID").Value & "," & oRecordset.Fields("PositionID").Value & "," & oRecordset.Fields("JobTypeID").Value & "," & oRecordset.Fields("ShiftID").Value & "," & oRecordset.Fields("JourneyID").Value & "," & oRecordset.Fields("ClassificationID").Value & "," & oRecordset.Fields("GroupGradeLevelID").Value & "," & oRecordset.Fields("IntegrationID").Value & "," & oRecordset.Fields("OccupationTypeID").Value & "," & oRecordset.Fields("ServiceID").Value & "," & oRecordset.Fields("LevelID").Value & "," & oRecordset.Fields("WorkingHours").Value & "," & oRecordset.Fields("StatusID").Value & "," & aLoginComponent(N_USER_ID_LOGIN) & "," & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
											If lErrorNumber = 0 Then
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobDate, EndDate From JobsHistoryList Where JobID =" & lJobID2 & " Order By JobDate Desc", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
												If Not oRecordset.EOF Then
													oRecordset.MoveNext
													If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update JobsHistoryList Set EndDate = " & AddDaysToSerialDate(lStartDate, -1) & " Where JobId = " & lJobID2 & " And EndDate = 30000000 And JobDate = " & oRecordset.Fields("JobDate").Value,"JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
													End If
												End If
											End If
										End If
										If lErrorNumber = 0 Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobID, OwnerID, JobNumber, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, WorkingHours, StartDate, EndDate, StatusID, ModifyDate From Jobs Where JobID =" & lJobID2, "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										End If
										If lErrorNumber = 0 Then
											aJobComponent(N_ID_JOB) = lJobID1
											aJobComponent(N_ID_EMPLOYEE_JOB) = lEmployeeID2
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsHistoryList (JobID, OwnerID, JobDate, EndDate, CompanyID, EmployeeID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, WorkingHours, StatusID, UserID, ModifyDate) Values (" & lJobID1 & "," & lEmployeeID2 & "," & lStartDate & "," & lEndDate & "," & oRecordset.Fields("CompanyID").Value & "," & lEmployeeID2 & "," & oRecordset.Fields("ZoneID").Value & "," & oRecordset.Fields("AreaID").Value & "," & oRecordset.Fields("PaymentCenterID").Value & "," & oRecordset.Fields("PositionID").Value & "," & oRecordset.Fields("JobTypeID").Value & "," & oRecordset.Fields("ShiftID").Value & "," & oRecordset.Fields("JourneyID").Value & "," & oRecordset.Fields("ClassificationID").Value & "," & oRecordset.Fields("GroupGradeLevelID").Value & "," & oRecordset.Fields("IntegrationID").Value & "," & oRecordset.Fields("OccupationTypeID").Value & "," & oRecordset.Fields("ServiceID").Value & "," & oRecordset.Fields("LevelID").Value & "," & oRecordset.Fields("WorkingHours").Value & "," & oRecordset.Fields("StatusID").Value & "," & aLoginComponent(N_USER_ID_LOGIN) & "," & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
											If lErrorNumber = 0 Then
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobDate, EndDate From JobsHistoryList Where JobID =" & lJobID1 & " Order By JobDate Desc", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
												If Not oRecordset.EOF Then
													oRecordset.MoveNext
													If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update JobsHistoryList Set EndDate = " & AddDaysToSerialDate(lStartDate, -1) & " Where JobId = " & lJobID1 & " And EndDate = 30000000 And JobDate = " & oRecordset.Fields("JobDate").Value,"JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
													End If
												End If
											End If
										End If
										If lErrorNumber = 0 Then
											sQuery = "Select ConceptID, ConceptAmount, EmployeesConceptsLKP.StartDate, EmployeesConceptsLKP.EndDate, AppliesToID, StartHour3, EndHour3 From EmployeesConceptsLKP, Employees Where (EmployeesConceptsLKP.EmployeeID = Employees.EmployeeID) And (ConceptID in (4,7,8)) And (EmployeesConceptsLKP.EmployeeID = " & lEmployeeID1 & ") And (EmployeesConceptsLKP.EndDate >= " & lStartDate & ")"
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										End If
										If lErrorNumber = 0 Then
											If Not oRecordset.EOF Then
												Do While Not oRecordset.EOF
													sQuery = "Update EmployeesConceptsLKP Set EndDate = " & AddDaysToSerialDate(lStartDate,-1) & " Where (EmployeeID = " & lEmployeeID1 & ") And (ConceptID = " & oRecordset.Fields("ConceptID").Value & ") And (StartDate = " & oRecordset.Fields("StartDate").Value & ") And (EndDate = " & oRecordset.Fields("EndDate").Value & ")"
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
													sQuery = "Insert Into EmployeesConceptsLKP (EmployeeID, ConceptID, StartDate, EndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, RegistrationDate, ModifyDate, StartUserID, EndUserID, UploadedFileName, Comments) Values (" & lEmployeeID2 & ", " & oRecordset.Fields("ConceptID").Value & ", " & lStartDate & ", " & lEndDate & ", " & oRequest("RiskLevel").Item * 10 & ",0,2,3,0,0,0,0, '" & oRecordset.Fields("AppliesToID") & "',1,401,1, " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", '" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "')"
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
													If CInt(oRecordset.Fields("ConceptID").Value) = 7 Or CInt(oRecordset.Fields("ConceptID").Value) = 8 Then
														sQuery = "Update Employees Set StartHour3 = " & oRecordset.Fields("StartHour3").Value & ", EndHour3 = " & oRecordset.Fields("EndHour3").Value & " Where (EmployeeID = " & lEmployeeID2 & ")"
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
													End If
													oRecordset.MoveNext
												Loop
											End If
										End If
										If lErrorNumber = 0 Then
											sQuery = "Select ConceptID, ConceptAmount, EmployeesConceptsLKP.StartDate, EmployeesConceptsLKP.EndDate, AppliesToID, StartHour3, EndHour3 From EmployeesConceptsLKP, Employees Where (EmployeesConceptsLKP.EmployeeID = Employees.EmployeeID) And (ConceptID in (4,7,8)) And (EmployeesConceptsLKP.EmployeeID = " & lEmployeeID2 & ") And (EmployeesConceptsLKP.EndDate >= " & lStartDate & ") And (EmployeesConceptsLKP.StartDate <" & lStartDate & ")"
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										End If
										If lErrorNumber = 0 Then
											If Not oRecordset.EOF Then
												Do While Not oRecordset.EOF
													sQuery = "Update EmployeesConceptsLKP Set EndDate = " & AddDaysToSerialDate(lStartDate,-1) & " Where (EmployeeID = " & lEmployeeID2 & ") And (ConceptID = " & oRecordset.Fields("ConceptID").Value & ") And (StartDate = " & oRecordset.Fields("StartDate").Value & ") And (EndDate = " & oRecordset.Fields("EndDate").Value & ")"
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
													sQuery = "Insert Into EmployeesConceptsLKP (EmployeeID, ConceptID, StartDate, EndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, AppliesToID, AbsenceTypeID, ConceptOrder, Active, RegistrationDate, ModifyDate, StartUserID, EndUserID, UploadedFileName, Comments) Values (" & lEmployeeID1 & ", " & oRecordset.Fields("ConceptID").Value & ", " & lStartDate & ", " & lEndDate & ", " & oRecordset.Fields("ConceptAmount").Value & ",0,2,3,0,0,0,0, '" & oRecordset.Fields("AppliesToID") & "',1,401,1, " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", '" & Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE), "'", "´") & "')"
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
													If CInt(oRecordset.Fields("ConceptID").Value) = 7 Or CInt(oRecordset.Fields("ConceptID").Value) = 8 Then
														sQuery = "Update Employees Set StartHour3 = " & oRecordset.Fields("StartHour3").Value & ", EndHour3 = " & oRecordset.Fields("EndHour3").Value & " Where (EmployeeID = " & lEmployeeID1 & ")"
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
														sQuery = "Update Employees Set StartHour3 = 0, EndHour3 = 0 Where (EmployeeID = " & lEmployeeID2 & ")"
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
													End If
													oRecordset.MoveNext
												Loop
											End If
										End If
									Case 37, 38, 39, 41 'Prórroga de licencias
										aJobComponent(N_ID_EMPLOYEE_JOB) = -1
										aJobComponent(N_STATUS_ID_JOB) = 4
										aJobComponent(N_ACTIVE_JOB) = 1
										sErrorDescription = "No se pudo guardar la información de la nueva plaza."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From JobsHistoryList Where (JobID=" & aJobComponent(N_ID_JOB) & ") Order by JobDate Desc", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										If Not oRecordset.EOF Then
											sErrorDescription = "No se pudo guardar la información de la  plaza."
											lHistoryJobDate = CLng(oRecordset.Fields("JobDate").Value)
											lHistoryEndDate = CLng(oRecordset.Fields("EndDate").Value)
											lMovementDate = oRequest("EmployeeEndYear")&oRequest("EmployeeEndMonth")&oRequest("EmployeeEndDay")
											lStatusJobID =  CLng(oRecordset.Fields("StatusID").Value)
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update JobsHistoryList Set EndDate=" & lMovementDate & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & " Where (JobID=" & aJobComponent(N_ID_JOB) & ") And (JobDate=" & lHistoryJobDate & ") And (EndDate=" & lHistoryEndDate & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										End If
									Case 36, 40, 43, 44, 45, 46, 47, 48
										bProcessed = 0
										aJobComponent(N_ID_EMPLOYEE_JOB) = -1
										aJobComponent(N_STATUS_ID_JOB) = 4
										aJobComponent(N_ACTIVE_JOB) = 1
										lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
									Case 29, 30, 31, 32, 33, 34
										bProcessed = 0
										aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 28
										aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = 0
										sErrorDescription = "No se pudo actualizar la información del empleado."
										If lReasonID <> 30 Then
											aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE), 1)
											aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 30000000
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryList (EmployeeID, EmployeeDate, EndDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, Active, ReasonID, ModifyDate, PayrollDate, UserID, bProcessed, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_NUMBER_EMPLOYEE), "'", "") & "', " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) & ", " & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & ", " & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & "," & bProcessed & ", '" & Replace(aEmployeeComponent(S_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
											If lErrorNumber = 0 Then
												sErrorDescription = "No se pudo actualizar la información del empleado."
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set StatusID=0 Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
											End If
										Else
											aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)
											aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 30000000
											Response.Redirect "UploadInfo.asp?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&EmployeeNumber=" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "&ReasonID=6&SaveEmployeesMovements=1&ModifyConcept=&EmployeeTypeID=" & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & "&JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & "&EmployeeDay=" & Mid(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE),7,2) & "&EmployeeMonth=" & Mid(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE),5,2) & "&EmployeeYear=" & Mid(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE),1,4) & "&EmployeePayrollDate=" & aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) & "&Authorization=+++++Aplicar+Movimiento++++"
										End If
									Case 21, 50
										If lErrorNumber = 0 Then
											'If lReasonID = 21 Then
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobID From Employees Where EmployeeID = " & oRequest("EmployeeID").Item, "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
												lJobID1 = oRecordset.Fields("JobID").Value
												lJobID2 = CLng(oRequest("JobID").Item)
											'End If
											aJobComponent(N_ACTIVE_JOB) = 1
											aJobComponent(N_JOB_DATE_JOB) = CLng(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE))
											aJobComponent(N_END_DATE_HISTORY_JOB) = CLng(aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE))
											aJobComponent(N_ID_EMPLOYEE_JOB) = CLng(aEmployeeComponent(N_ID_EMPLOYEE))
											aJobComponent(N_ACTIVE_JOB) = 1
											aJobComponent(N_STATUS_ID_JOB) = 1
											lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
											If lErrorNumber = 0 Then
												sErrorDescription = "No se pudo actualizar la información del empleado."
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
											End If
										End If
										'If lReasonID = 21 Then
											If lErrorNumber = 0 Then
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Jobs Set OwnerID=-1 Where (JobID = " & lJobID1 & ")" , "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
												If lErrorNumber = 0 Then
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Jobs Set OwnerID=" & oRequest("EmployeeID").Item & " Where (JobID = " & lJobID2 & ")" , "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
												End If
											End If
										'End If
										sErrorDescription = "No se pudo obtener la información del registro."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeDate, EndDate, JobID From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID<>0) And (ReasonID<>58) And (ReasonID <> 57) And (EndDate <> 0) Order By EmployeeDate Desc", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
										If Not oRecordset.EOF Then
											oRecordset.MoveNext
											If Not oRecordset.EOF Then
												aJobComponent(N_ID_JOB) = CLng(oRecordset.Fields("JobID").Value)
												'If lReasonID = 21 Then lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
												lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
												aJobComponent(N_STATUS_ID_JOB) = 2
												If aJobComponent(N_ID_JOB) <> -1 Then
													If CLng(aJobComponent(N_END_DATE_JOB)) > CLng(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)) Then
														aJobComponent(N_JOB_DATE_JOB) = CLng(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE))
														aJobComponent(N_END_DATE_HISTORY_JOB) = CLng(aJobComponent(N_END_DATE_JOB))
														aJobComponent(N_ID_EMPLOYEE_JOB) = -1
														aJobComponent(N_ACTIVE_JOB) = 1
														aJobComponent(N_STATUS_ID_JOB) = 2
														aJobComponent(B_CHECK_FOR_DUPLICATED_JOB) = False
														sErrorDescription = "No se pudo guardar la información de la plaza."
														lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
													End If
												End If
											End If
										End If
									Case 51
										If lErrorNumber = 0 Then
											aJobComponent(N_ACTIVE_JOB) = 1
											aJobComponent(N_AREA_ID_JOB) = lAreaID
											aJobComponent(N_PAYMENT_CENTER_ID_JOB) = lPaymentCenterID
											aJobComponent(N_SERVICE_ID_JOB)  = lServiceID
											aJobComponent(N_JOURNEY_ID_JOB)  = lJourneyID
											aJobComponent(N_SHIFT_ID_JOB)  = lShiftID
											aJobComponent(N_JOB_DATE_JOB) = CLng(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE))
											aJobComponent(N_END_DATE_HISTORY_JOB) = CLng(aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE))
											aJobComponent(N_STATUS_ID_JOB) = 1
											lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
											If CInt(oRequest("RiskLevel").Item) > 0 Then
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
											If CInt(oRequest("StartHour3").Item) > 0 And CInt(oRequest("EndHour3").Item) > 0 Then
												sErrorDescription = "No se pudo agregar la percepción adicional del empleado."
												aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
												aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)
												aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = 3/6.5*100
												aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) = oRequest("StartHour3").Item
												aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) = oRequest("EndHour3").Item
												aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) = 0
												aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 2
												aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 3
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
															lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set StartHour3=" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", EndHour3=" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
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
									Case Else
										If aJobComponent(N_ID_EMPLOYEE_JOB) <> -1 And lReasonID <> EMPLOYEES_FOR_RISK And _
											lReasonID <> EMPLOYEES_ADDITIONALSHIFT And lReasonID <> EMPLOYEES_CONCEPT_08 Then
											aJobComponent(N_JOB_DATE_JOB) = CLng(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE))
											aJobComponent(N_END_DATE_HISTORY_JOB) = CLng(aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE))
											aJobComponent(N_ID_EMPLOYEE_JOB) = CLng(aEmployeeComponent(N_ID_EMPLOYEE))
											aJobComponent(N_ACTIVE_JOB) = 1
											lErrorNumber = ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
										End If
								End Select
							End If
						End If
					End If
					If lErrorNumber = 0 Then
						If InStr(1, ",12,13,14,17,18,57,68,", "," & lReasonID  & ",", vbBinaryCompare) > 0 Then
							sErrorDescription = "No se pudo actualizar la información del empleado."
							If B_UPPERCASE Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set EmployeeName='" & Replace(UCase(aEmployeeComponent(S_NAME_EMPLOYEE)), "'", "´") & "', EmployeeLastName='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME_EMPLOYEE)), "'", "´") & "', EmployeeLastName2='" & Replace(UCase(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE)), "'", "´") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", StartHour1=" & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & ", EndHour1=" & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & ", StartHour2=" & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & ", EndHour2=" & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & ", StartHour3=" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", EndHour3=" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", EmployeeEmail='" & Replace(aEmployeeComponent(S_EMAIL_EMPLOYEE), "'", "") & "', SocialSecurityNumber='" & Replace(aEmployeeComponent(S_SSN_EMPLOYEE), "'", "") & "', BirthYear=" & CInt(Left(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("0000"))) & ", BirthMonth=" & CInt(Mid(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("00000"), Len("00"))) & ", BirthDay=" & CInt(Mid(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("0000000"), Len("00"))) & ", BirthDate=" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) & ", StartDate=" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & ", StartDate2=" & aEmployeeComponent(N_START_DATE2_EMPLOYEE) & ", CountryID=" & aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) & ", RFC='" & Replace(aEmployeeComponent(S_RFC_EMPLOYEE), "'", "") & "', CURP='" & Replace(aEmployeeComponent(S_CURP_EMPLOYEE), "'", "") & "', GenderID=" & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & ", MaritalStatusID=" & aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							Else
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set EmployeeName='" & Replace(aEmployeeComponent(S_NAME_EMPLOYEE), "'", "´") & "', EmployeeLastName='" & Replace(aEmployeeComponent(S_LAST_NAME_EMPLOYEE), "'", "´") & "', EmployeeLastName2='" & Replace(aEmployeeComponent(S_LAST_NAME2_EMPLOYEE), "'", "´") & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", JobID=" & aJobComponent(N_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", PositionTypeID=" & aJobComponent(N_POSITION_TYPE_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ShiftID=" & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & ", StartHour1=" & aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) & ", EndHour1=" & aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) & ", StartHour2=" & aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) & ", EndHour2=" & aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) & ", StartHour3=" & aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) & ", EndHour3=" & aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", EmployeeEmail='" & Replace(aEmployeeComponent(S_EMAIL_EMPLOYEE), "'", "") & "', SocialSecurityNumber='" & Replace(aEmployeeComponent(S_SSN_EMPLOYEE), "'", "") & "', BirthYear=" & CInt(Left(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("0000"))) & ", BirthMonth=" & CInt(Mid(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("00000"), Len("00"))) & ", BirthDay=" & CInt(Mid(Right(("00000000" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE)), Len("00000000")), Len("0000000"), Len("00"))) & ", BirthDate=" & aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) & ", StartDate=" & aEmployeeComponent(N_START_DATE_EMPLOYEE) & ", StartDate2=" & aEmployeeComponent(N_START_DATE2_EMPLOYEE) & ", CountryID=" & aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) & ", RFC='" & Replace(aEmployeeComponent(S_RFC_EMPLOYEE), "'", "") & "', CURP='" & Replace(aEmployeeComponent(S_CURP_EMPLOYEE), "'", "") & "', GenderID=" & aEmployeeComponent(N_GENDER_ID_EMPLOYEE) & ", MaritalStatusID=" & aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) & ", Active=" & aEmployeeComponent(N_ACTIVE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							End If
							If lReasonID = 12 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set StartDate=" & oRequest("EmployeeYear").Item & oRequest("EmployeeMonth").Item & oRequest("EmployeeDay").Item & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							End If
							sErrorDescription = "No se pudo guardar la información extra del nuevo empleado."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesExtraInfo Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo modificar la información extra del empleado."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesExtraInfo (EmployeeID, EmployeeAddress, EmployeeCity, EmployeeZipCode, StateID, CountryID, EmployeePhone, OfficePhone, OfficeExt, DocumentNumber1, DocumentNumber2, DocumentNumber3, EmployeeActivityID) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_ADDRESS_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_CITY_EMPLOYEE), "'", "´") & "', '" & Replace(aEmployeeComponent(S_ZIP_CODE_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_ADDRESS_COUNTRY_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_EMPLOYEE_PHONE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_OFFICE_PHONE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_EXT_OFFICE_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_DOCUMENT_NUMBER_2_EMPLOYEE), "'", "") & "', '" & Replace(aEmployeeComponent(S_DOCUMENT_NUMBER_3_EMPLOYEE), "'", "") & "', " & aEmployeeComponent(N_ACTIVITY_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
							End If
						End If
					End If
					'Cerrar los conceptos 4, 7 y 8 cuando el empleado obtiene una licencia sin sueldo
					If lErrorNumber = 0 Then
						Select Case lReasonID
							Case 37, 38, 39, 40, 41, 43, 44, 45, 46, 47, 48
								sErrorDescription = "No se pudo actualizar la información de la plaza."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesConceptsLKP Set EndDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EndDate>=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And (ConceptID Not In (" & CONCEPTS_NOT_EXPIRE & "))", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						End Select
					End If
					' Reasignar el número de empleado a un número empleado válido
					If lErrorNumber = 0 Then
						If (lReasonID = 12) Or (lReasonID = 13) Or (lReasonID = 14) Then
							If aEmployeeComponent(N_ID_EMPLOYEE) >= 1000000 Then
								lErrorNumber = ModifyEmployeeNumber(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							End If
						End If
					End If
				End If
				If (lErrorNumber = 0) Then
					If Len(oRequest("Authorization").Item) > 0 Then
						If InStr(1,",12,13,14,17,18,68,","," & lReasonID & ",",vbBinaryCompare) <> 0 Then
							lErrorNumber = AddEmployeeBankAccount(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							If lErrorNumber <> 0 then sErrorDescription = "No se pudo incluir la información de la cuenta de pago"
						End If
					End If
				End If
				If lErrorNumber = 0 Then
					If lReasonID = 13 Then
						If CLng(aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)) < CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) Then
							aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = CLng(oRequest("EmployeeEndYear").Item & oRequest("EmployeeEndMonth").Item & oRequest("EmployeeEndDay").Item)
							aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)
							Response.Redirect "UploadInfo.asp?EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & "&ReasonID=10&SaveEmployeesMovements=1&ModifyConcept=&JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & "&EmployeeNumber=" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "&EmployeeDay=" & oRequest("EmployeeEndDay").Item & "&EmployeeMonth=" & oRequest("EmployeeEndMonth").Item & "&EmployeeYear=" & oRequest("EmployeeEndYear").Item & "&EmployeePayrollDate=" & oRequest("EmployeePayrollDate").Item & "&Authorization=+++++Aplicar+Movimiento++++"
						End If
					End If
				End If
			End If
		End If
	End If

	If lErrorNumber = 0 Then
		Select Case iStatusReasonID
			Case 0
				sErrorDescription = "El movimiento fue aplicado exitosamente."
			Case 2
				sErrorDescription = "El movimiento fue registrado para validación exitosamente."
			Case 3
				sErrorDescription = "El movimiento fue registrado para autorización exitosamente."
			Case Else
				sErrorDescription = "El movimiento fue guardado exitosamente."
		End Select
	End If

	Set oRecordset = Nothing
	AddEmployeeMovement = lErrorNumber
	Err.Clear
End Function

Function AddEmployeeMovementFile(oRequest, oADODBConnection, sQuery, lReasonID, aEmployeeComponent, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new movement for the employee into the database
'Inputs:  oRequest, oADODBConnection, sQuery, lReasonID
'Outputs: aEmployeeComponent, aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeMovementFile"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información de la aplicación de movimientos masivos de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 0
				If Len(CStr(oRecordset.Fields("EmployeeID").Value)) > 0 Then
					aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
					lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
						If lErrorNumber = 0 Then
							If lReasonID = EMPLOYEES_FOR_RISK Then
								lErrorNumber = ModifyEmployeeConcepts(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							Else
								lErrorNumber = AddEmployeeMovement(oRequest, oADODBConnection, lReasonID, aEmployeeComponent, aJobComponent, sErrorDescription)
							End If
						End If
					End If
				End If
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
		End If
	End If

	Set oRecordset = Nothing
	AddEmployeeMovementFile = lErrorNumber
	Err.Clear
End Function

Function AddEmployeeReasonForRejection(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add a reasons for rejection for employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeReasonForRejection"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim oRecordset

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	sErrorDescription = "No se pudo obtener la información del registro."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusID, Active From StatusEmployees Where (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ") And (StatusReasonID=1)", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = CLng(oRecordset.Fields("StatusID").Value)
		End If
	End If
	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudo guardar la información de la razón de rechazo para el movimiento empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesReasonsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo guardar la información de la razón de rechazo para el movimiento empleado."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesReasonsLKP (EmployeeID, EmployeeDate, ReasonID, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ", '" & Replace(aEmployeeComponent(S_REASON_FOR_REJECTION_COMMENTS_EMPLOYEE), "'", "") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			If lErrorNumber = 0 Then
				If aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 0 Then aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 30000000
				sErrorDescription = "No se pudo guardar la información de la razón de rechazo para el movimiento empleado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ") And (EndDate=" & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudo guardar la información de la razón de rechazo para el movimiento empleado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & " Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
		End If
		oRecordset.Close
	End If

	AddEmployeeReasonForRejection = lErrorNumber
	Err.Clear
End Function

Function AddEmployeesRequirements(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To add the attachments employee movement into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeesRequirements"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) And (aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = -1) And (aEmployeeComponent(N_REASON_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado, tipo y fecha de movimiento para agregar la información del movimiento."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeAddComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeRequirementID From EmployeesRequirements Where (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sErrorDescription = "No se pudo actualizar la información del empleado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesRequirementsFM1LKP Where (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ") And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (RequirementDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And (ValidatedUserID=-1)", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
				Do While Not oRecordset.EOF
					If Not IsEmpty(oRequest(CStr(oRecordset.Fields("EmployeeRequirementID").Value))) Then
						sErrorDescription = "No se pudo actualizar la información del empleado."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesRequirementsFM1LKP (EmployeeID, EmployeeRequirementID, ReasonID, RequirementDate, RegisteredUserID, ValidatedUserID) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & CInt(oRecordset.Fields("EmployeeRequirementID").Value) & ", " & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", -1)", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	AddEmployeesRequirements = lErrorNumber
	Err.Clear
End Function

Function AddUploadThirdCreditsRejected(oRequest, oADODBConnection, aEmployeeComponent, iRejectType, sErrorDescription)
'************************************************************
'Purpose: To add a new credit rejected report from third upload files
'		for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddUploadThirdCreditsRejected"
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

	If lErrorNumber = 0 Then
		If aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = -1 Then aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = 0
		sErrorDescription = "No se pudo agregar la información del registro de crédito rechazado para el empleado: " & aEmployeeComponent(N_ID_EMPLOYEE)
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into UploadThirdCreditsRejected (EmployeeID, CreditTypeID, UploadedFileName, UploadedRecordType, UploadedRejectType, UploadedRecordLine, Comments) Values (" & aEmployeeComponent(N_ID_EMPLOYEE) & ", " & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ", '" &  Replace(aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE), "'", "´") & "', " & aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) & ", " &  iRejectType & ", " & aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_LINE) & ", '" & Replace(aEmployeeComponent(S_CREDIT_UPLOADED_REJECT_COMMENTS), "'", "´") & "')", "EmployeeAddComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	End If

	AddUploadThirdCreditsRejected = lErrorNumber
	Err.Clear	
End Function

Function AddEmployeesSpecialJourney(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To validate and register a special journey 
'			if the budget is sufficient
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeesSpecialJourney"
	Dim oRecordset
	Dim lErrorNumber
	Dim sQuery
	Dim lUca
	Dim lUR
	Dim lCT
	Dim lAux
	Dim lAppliedDate
	Dim dTotal
	Dim lEmployeeTypeID
	Dim lRecordID
	'Dim sQuery2

	'sQuery2 = "Insert Into EmployeesSpecialJourneys " & _
'		"(RecordID, EmployeeID, EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, " & _
'		"RFC, CURP, OriginalEmployeeID, PositionID, AreaID, ServiceID, LevelID, WorkingHours, " & _
'		"ShiftID, RiskLevelID, SpecialJourneyID, DocumentNumber, StartDate, EndDate, StartHour, " & _
'		"EndHour, JourneyID, WorkedHours, MovementID, FactorID, ReasonID, Comments, " & _
'		"ConceptAmount, AddUserID, AddDate, AppliedDate, Removed, RemoveUserID, RemovedDate, " & _
'		"AppliedRemoveDate, Active) Values (" & lRecordID & "," & oRequest("EmployeeID").Item & ",'" & _
'		oRequest("EmployeeNumber").Item & "','" & oRequest("EmployeeName").Item & "','" & oRequest("EmployeeLastName").Item & "','" & _
'		oRequest("EmployeeLastName2").Item & "','" & oRequest("rfc").Item & "','" & oRequest("curp").Item & "'," & _
'		oRequest("OriginalEmployeeID").Item & "," & oRequest("PositionID").Item & "," & oRequest("AreaID").Item & "," & _
'		oRequest("ServiceID").Item & "," & oRequest("LevelID").Item & "," & oRequest("WorkingHours").Item & "," & _
'		oRequest("ShiftID").Item & "," & oRequest("RiskLevelID").Item & "," & oRequest("SpecialJourneyID").Item & ",'" & _
'		oRequest("DocumentNumber").Item & "'," & _
'		oRequest("StartDateYear").Item & oRequest("StartDateMonth").Item & oRequest("StartDateDay").Item & "," & _
'		oRequest("EndDateYear").Item & oRequest("EndDateMonth").Item & oRequest("EndDateDay").Item & "," & _
'		oRequest("StartHour").Item & "," & oRequest("EndHour").Item & "," & oRequest("JourneyID").Item & "," & _
'		oRequest("WorkedHours").Item & "," & oRequest("MovementID").Item & "," & oRequest("FactorID").Item & "," & _
'		oRequest("ReasonID").Item & ",'" & oRequest("Comments").Item & "'," & oRequest("ConceptAmount").Item & ","& _
'		oRequest("AddUserID").Item & "," & oRequest("AddDate").Item & "," & oRequest("AppliedDate").Item & "," & _
'		oRequest("Removed").Item & "," & oRequest("RemoveUserID").Item & "," & oRequest("RemovedDate").Item & "," & _
'		oRequest("AppliedRemoveDate").Item & "," & oRequest("Active").Item & ")"

	'Response.Write sQuery2

	If CLng(oRequest("EmployeeID").Item) < 800000 Then
		lEmployeeTypeID = 0 'Personal Interno
	Else
		lEmployeeTypeID = 1 'Personal Externo
	End If
	
	sErrorDescription = "No se pudo obtener el centro de pago correspondiente"
	If CLng(oRequest("SectionID").Item) = 423 Then 
		If lEmployeeTypeID = 0 Then
			sQuery = "Select URCTAUX, PositionTypeID From Areas, Positions, (Select AreaID, PositionID From Jobs Where JobID = (Select JobID From Employees Where EmployeeID = " & oRequest("EmployeeID").Item & ")) J Where (Areas.AreaID = J.AreaID) And (Positions.PositionID = J.PositionID)"
		Else
			sQuery = "Select URCTAUX, PositionTypeID From Areas, Positions Where (AreaID = " & oRequest("AreaID").Item & ") And (PositionID = " & oRequest("PositionID").Item & ")"
		End If
	ElseIf CLng(oRequest("SectionID").Item) = 424 Then 
			sQuery = "Select URCTAUX, PositionTypeID From Areas, Positions, (Select AreaID, PositionID From Jobs Where JobID = (Select JobID From Employees Where EmployeeID = " & oRequest("OriginalEmployeeID").Item & ")) J Where (Areas.AreaID = J.AreaID) And (Positions.PositionID = J.PositionID)"
	End If
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeAddComponent.asp", "_root", 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			lUca = oRecordset.Fields("URCTAUX").Value
			lUR = CInt(Mid(lUca,1,3))
			lCT = CInt(Mid(lUca,5,2))
			lAux = CInt(Mid(lUca,8))
			lAppliedDate = oRequest("AppliedDate").Item
			dTotal = CDbl(oRequest("ConceptAmount").Item)
			'sQuery = "Select OriginalAmount From BudgetsMoney Where (BudgetUR = " & lUR & ") And (BudgetCT = " & lCT & ") And (BudgetAux = " & lAux & ") And (BudgetYear = " & Mid(lAppliedDate,1,4) & ") And (BudgetMonth = " & Mid(lAppliedDate,5,2) & ") And (BudgetID3 In (20036, 20037, 20039)) Order By BudgetCT Asc;"
			'Consulta temporal a tabla Budgets_Short hasta respuesta del área de Presupuestos.
			sQuery = "Select OriginalAmount From Budgets_Short Where (ZoneID = " & lUR & ") And (BudgetYear = " & Mid(lAppliedDate,1,4) & ") And (BudgetMonth = " & Mid(lAppliedDate,5,2) & ") And (BudgetEmployeeTypeID = " & lEmployeeTypeID & ")"
			sErrorDescription = "No se pudieron obtener los montos autorizados para el presupuesto correspondiente"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeAddComponent.asp", "_root", 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					If CDbl(oRecordset.Fields("OriginalAmount").Value) - dTotal < 0 Then
						sErrorDescription = "El presupuesto no es suficiente para registrar este movimiento. La operación no puede continuar"
						lErrorNumber = -1
					Else
						If lErrorNumber = 0 Then
							sErrorDescription =  "No se pudo obtener el número de registro actual"
							sQuery = "Select Max(RecordID) NextID From EmployeesSpecialJourneys"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeAddComponent.asp", "_root", 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									lRecordID = CLng(oRecordset.Fields("NextID").Value) + 1
								Else
									lRecordID = 1
								End If
								sErrorDescription =  "No se pudo registrar el movimiento"
								sQuery = "Insert Into EmployeesSpecialJourneys " & _
									"(RecordID, EmployeeID, EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, " & _
									"RFC, CURP, OriginalEmployeeID, PositionID, AreaID, ServiceID, LevelID, WorkingHours, " & _
									"ShiftID, RiskLevelID, SpecialJourneyID, DocumentNumber, StartDate, EndDate, StartHour, " & _
									"EndHour, JourneyID, WorkedHours, MovementID, FactorID, ReasonID, Comments, " & _
									"ConceptAmount, AddUserID, AddDate, AppliedDate, Removed, RemoveUserID, RemovedDate, " & _
									"AppliedRemoveDate, Active) Values (" & lRecordID & "," & oRequest("EmployeeID").Item & ",'" & _
									oRequest("EmployeeNumber").Item & "','" & oRequest("EmployeeName").Item & "','" & oRequest("EmployeeLastName").Item & "','" & _
									oRequest("EmployeeLastName2").Item & "','" & oRequest("rfc").Item & "','" & oRequest("curp").Item & "'," & _
									oRequest("OriginalEmployeeID").Item & "," & oRequest("PositionID").Item & "," & oRequest("AreaID").Item & "," & _
									oRequest("ServiceID").Item & "," & oRequest("LevelID").Item & "," & oRequest("WorkingHours").Item & "," & _
									oRequest("ShiftID").Item & "," & oRequest("RiskLevelID").Item & "," & oRequest("SpecialJourneyID").Item & ",'" & _
									oRequest("DocumentNumber").Item & "'," & _
									oRequest("StartDateYear").Item & oRequest("StartDateMonth").Item & oRequest("StartDateDay").Item & "," & _
									oRequest("EndDateYear").Item & oRequest("EndDateMonth").Item & oRequest("EndDateDay").Item & "," & _
									oRequest("StartHour").Item & "," & oRequest("EndHour").Item & "," & oRequest("JourneyID").Item & "," & _
									oRequest("WorkedHours").Item & "," & oRequest("MovementID").Item & "," & oRequest("FactorID").Item & "," & _
									oRequest("ReasonID").Item & ",'" & oRequest("Comments").Item & "'," & oRequest("ConceptAmount").Item & ","& _
									oRequest("AddUserID").Item & "," & oRequest("AddDate").Item & "," & oRequest("AppliedDate").Item & "," & _
									oRequest("Removed").Item & "," & oRequest("RemoveUserID").Item & "," & oRequest("RemovedDate").Item & "," & _
									oRequest("AppliedRemoveDate").Item & "," & oRequest("Active").Item & ")"
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeAddComponent.asp", "_root", 000, sErrorDescription, Null)
								If lErrorNumber = 0 Then
									'sQuery = "Update BudgetsMoney Set OriginalAmount = OriginalAmount -" & oRequest("ConceptAmount").Item & ", ModifiedAmount = ModifiedAmount + " & oRequest("ConceptAmount").Item & " Where (BudgetUR = " & lUR & ") And (BudgetCT = " & lCT & ") And (BudgetAux = " & lAux & ") And (BudgetYear = " & Mid(lAppliedDate,1,4) & ") And (BudgetMonth = " & Mid(lAppliedDate,5,2) & ") And (BudgetID3 In (20036, 20037, 20039))"
									'Consulta temporal a tabla Budgets_Short hasta respuesta del área de Presupuestos.
									sQuery = "Update Budgets_Short Set OriginalAmount = OriginalAmount - " & oRequest("ConceptAmount").Item & ", ModifiedAmount = ModifiedAmount + " & oRequest("ConceptAmount").Item & " Where (ZoneID = " & lUR & ") And (BudgetYear = " & Mid(lAppliedDate,1,4) & ") And (BudgetMonth = " & Mid(lAppliedDate,5,2) & ") And (BudgetEmployeeTypeID = " & lEmployeeTypeID & ")"
									sErrorDescription = "Los presupuestos original y modificados no pudieron actualziarse."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeAddComponent.asp", "_root", 000, sErrorDescription, Null)
								End If
							End If
						End If
					End If
				Else
					lErrorNumber = -1
					sErrorDescription = "No se encontró la partida presupuestal correspondiente."
				End If
			End If
		Else
			lErrorNumber = -1
		End If
	End If

	Set oRecordset = Nothing
	AddEmployeesSpecialJourney = lErrorNumber
	Err.Clear
End Function

%>