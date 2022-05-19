<%
Function GetConceptAmount(oADODBConnection, lEmployeeID, iConceptID, lStartDate, lAmount, sErrorDescription)
'************************************************************
'Purpose: To get the concept amount defined for the given employee
'Inputs:  oADODBConnection, lEmployeeID, iConceptID, lStartDate
'Outputs: lAmount, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetConceptAmount"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo eliminar la información del concepto del empleado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesConceptsLKP Where (EmployeeID=" & lEmployeeID & ") And (ConceptID=" & iConceptID & ") And (StartDate<=" & lStartDate & ") And (EndDate>=" & lStartDate & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			lAmount = 0
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "El concepto del empleado especificado no se encuentra en el sistema."
		Else
			lAmount = CLng(oRecordset.Fields("ConceptAmount").Value)
		End If
	End If

	GetConceptAmount = lErrorNumber
	Err.Clear
End Function

Function GetCrossingEmployeeConceptType(oADODBConnection, aEmployeeComponent, sEmployeeConceptType, lStartDate, lEndDate, sErrorDescription)
'************************************************************
'Purpose: To get the type of crossing employee concept for the
'         record to insert
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sEmployeeConceptType, lStartDate, lEndDate, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetCrossingEmployeeConceptType"
	Dim oRecordset
	Dim lErrorNumber
	Dim sQuery

	Select Case aEmployeeComponent(N_CONCEPT_CREDIT_TYPE)
		Case 0
			sQuery = "Select * From Credits Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
					 " And (CreditTypeID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ")" & _
					 " And (StartDate>=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")" & _
					 " And (EndDate>=StartDate)"
		Case 1
			sQuery = "Select * from EmployeesConceptsLKP Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
					 " And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ")" & _
					 " And (StartDate>=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")" & _
					 " And (EndDate>=StartDate) Order By StartDate Desc"
		Case 2
			sQuery = "Select * from BankAccounts Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
					 " And (StartDate>=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")" & _
					 " And (EndDate>=StartDate) Order By StartDate Desc"
		Case 3
			sQuery = "Select * from EmployeesGrades Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
					 " And (StartDate>=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")" & _
					 " And (EndDate>=StartDate) Order By StartDate Desc"
	End Select
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sEmployeeConceptType = "Cross"
			lStartDate = CLng(oRecordset.Fields("StartDate").Value)
			lEndDate = CLng(oRecordset.Fields("EndDate").Value)
		Else
			Select Case aEmployeeComponent(N_CONCEPT_CREDIT_TYPE)
				Case 0
					sQuery = "Select * From Credits Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
							 " And (CreditTypeID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ")" & _
							 " And (StartDate<" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")" & _
							 " And (EndDate>=" &  aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")"
				Case 1
					sQuery = "Select * from EmployeesConceptsLKP Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
							 " And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ")" & _
							 " And (StartDate<" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")" & _
							 " And (EndDate>" &  aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") Order By StartDate Desc"
				Case 2
					sQuery = "Select * from BankAccounts Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
							 " And (StartDate<" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")" & _
							 " And (EndDate>" &  aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") Order By StartDate Desc"
				Case 3
					sQuery = "Select * from EmployeesGrades Where (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ")" & _
							 " And (StartDate<" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")" & _
							 " And (EndDate>" &  aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ") Order By StartDate Desc"
			End Select
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sEmployeeConceptType = "Inner"
					lStartDate = CLng(oRecordset.Fields("StartDate").Value)
					lEndDate = CLng(oRecordset.Fields("EndDate").Value)
				End If
			Else
				Select Case aEmployeeComponent(N_CONCEPT_CREDIT_TYPE)
					Case 0
						sErrorDescription = "No se pudo obtener la información del tercero."
					Case 1
						sErrorDescription = "No se pudo obtener la información del concepto."
					Case 2
						sErrorDescription = "No se pudo obtener la información de la cuenta bancaria."
					Case 3
						sErrorDescription = "No se pudo obtener la información de la calificación del empleado."
				End Select
			End If
		End If
	Else
		Select Case aEmployeeComponent(N_CONCEPT_CREDIT_TYPE)
			Case 0
				sErrorDescription = "No se pudo obtener la información del tercero."
			Case 1
				sErrorDescription = "No se pudo obtener la información del concepto."
			Case 2
				sErrorDescription = "No se pudo obtener la información de la cuenta bancaria."
			Case 3
				sErrorDescription = "No se pudo obtener la información de la calificación del empleado."
		End Select
	End If

	Set oRecordset = Nothing
	GetCrossingEmployeeConceptType = lErrorNumber
	Err.Clear
End Function

Function GetDocumentsForLicenses(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a employee's document license from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetDocumentsForLicenses"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From DocumentsForLicenses Where DocumentForLicenseID=" & aEmployeeComponent(N_DOCUMENT_FOR_LICENSE_ID_EMPLOYEE), "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El empleado especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
				aEmployeeComponent(S_DOCUMENT_FOR_LICENSE_NUMBER_EMPLOYEE) = CStr(oRecordset.Fields("DocumentForLicenseNumber").Value)
				aEmployeeComponent(S_DOCUMENT_FOR_CANCEL_LICENSE_NUMBER_EMPLOYEE) = CStr(oRecordset.Fields("DocumentForCancelLicenseNumber").Value)
				aEmployeeComponent(S_DOCUMENT_TEMPLATE_EMPLOYEE) = CStr(oRecordset.Fields("DocumentTemplate").Value)
				aEmployeeComponent(S_REQUEST_NUMBER_EMPLOYEE) = CStr(oRecordset.Fields("RequestNumber").Value)
				aEmployeeComponent(N_SYNDICATE_TYPE_ID_LICENSE_EMPLOYEE) = CStr(oRecordset.Fields("LicenseSyndicateTypeID").Value)
				aEmployeeComponent(N_DATE_LICENSE_DOCUMENT_EMPLOYEE) = CStr(oRecordset.Fields("DocumentLicenseDate").Value)
				aEmployeeComponent(N_LICENSE_START_DATE_EMPLOYEE) = CStr(oRecordset.Fields("LicenseStartDate").Value)
				aEmployeeComponent(N_LICENSE_END_DATE_EMPLOYEE) = CStr(oRecordset.Fields("LicenseEndDate").Value)
				aEmployeeComponent(N_CANCEL_DATE_LICENSE_DOCUMENT_EMPLOYEE) = CStr(oRecordset.Fields("LicenseCancelDate").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetDocumentsForLicenses = lErrorNumber
	Err.Clear
End Function

Function GetEndDateFromCredit(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the end date from credit with start date and 
'         period numbers to pay the credit
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEndDateFromCredit"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim iPeriodsToDays
	Dim lEndDate

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó la fecha inicial del crédito o el número de pagos para liquidarlo."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) = 0 Then
			Select Case aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE)
				Case 58, 64, 83
					aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 30000000
				Case Else
					aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE)
			End Select
		Else
			iPeriodsToDays = aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) * 15
			lEndDate = GetPayrollEndDate(AddDaysToSerialDate(aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE), iPeriodsToDays))
			sErrorDescription = "No se pudo obtener la fecha de fin del crédito."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Top 1 ForPayrollDate from Payrolls Where PayrollDate =" & lEndDate, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If oRecordset.EOF Then
					aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = CLng(lEndDate)
				Else
					aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = CLng(oRecordset.Fields("ForPayrollDate").Value)
				End If
				oRecordset.Close
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetEndDateFromCredit = lErrorNumber
	Err.Clear
End Function

Function GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about an employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployee"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim lStatusReason
	Dim lConceptID

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) And (aEmployeeComponent(N_JOB_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del empleado."
		If aEmployeeComponent(N_ID_EMPLOYEE) <> -1 Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Employees Where EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE), "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If Not oRecordset.EOF Then
				aEmployeeComponent(N_JOB_ID_EMPLOYEE) = CLng(oRecordset.Fields("JobID").Value)
			Else
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El empleado especificado no se encuentra en el sistema."
			End If
			If lErrorNumber = 0 Then
				If aEmployeeComponent(N_JOB_ID_EMPLOYEE) <> -1 Then
					If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) = 0 Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Employees Where EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE), "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.* From Employees, Jobs Where (Employees.JobID=Jobs.JobID) And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Jobs.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))) And (Employees.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					End If
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Employees Where EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE), "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				End If
			End If
		Else
			If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) = 0 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Employees Where JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE), "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.* From Employees, Jobs Where (Employees.JobID=Jobs.JobID) And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Jobs.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))) And (JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			End If
		End If	
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No tiene permisos para realizar movimientos a empleados que pertenecen a otro centro de trabajo."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
				aEmployeeComponent(S_NUMBER_EMPLOYEE) = CStr(oRecordset.Fields("EmployeeNumber").Value)
				aEmployeeComponent(S_ACCESS_KEY_EMPLOYEE) = CStr(oRecordset.Fields("EmployeeAccessKey").Value)
				aEmployeeComponent(S_PASSWORD_EMPLOYEE) = CStr(oRecordset.Fields("EmployeePassword").Value)
				aEmployeeComponent(S_NAME_EMPLOYEE) = CStr(oRecordset.Fields("EmployeeName").Value)
				aEmployeeComponent(S_LAST_NAME_EMPLOYEE) = CStr(oRecordset.Fields("EmployeeLastName").Value)
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
					aEmployeeComponent(S_LAST_NAME2_EMPLOYEE) = CStr(oRecordset.Fields("EmployeeLastName2").Value)
				Else
					aEmployeeComponent(S_LAST_NAME2_EMPLOYEE) = " "
				End If
				aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) = CLng(oRecordset.Fields("CompanyID").Value)
				aEmployeeComponent(N_JOB_ID_EMPLOYEE) = CLng(oRecordset.Fields("JobID").Value)
				aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) = CLng(oRecordset.Fields("ServiceID").Value)
				aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeTypeID").Value)
				aEmployeeComponent(N_POSITION_TYPE_ID_EMPLOYEE) = CLng(oRecordset.Fields("PositionTypeID").Value)
				aEmployeeComponent(N_CLASSIFICATION_ID_EMPLOYEE) = CLng(oRecordset.Fields("ClassificationID").Value)
				aEmployeeComponent(N_GROUP_GRADE_LEVEL_ID_EMPLOYEE) = CLng(oRecordset.Fields("GroupGradeLevelID").Value)
				aEmployeeComponent(N_INTEGRATION_ID_EMPLOYEE) = CLng(oRecordset.Fields("IntegrationID").Value)
				aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) = CLng(oRecordset.Fields("JourneyID").Value)
				aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) = CLng(oRecordset.Fields("ShiftID").Value)
				aEmployeeComponent(N_START_HOUR_1_EMPLOYEE) = CInt(oRecordset.Fields("StartHour1").Value)
				aEmployeeComponent(N_END_HOUR_1_EMPLOYEE) = CInt(oRecordset.Fields("EndHour1").Value)
				aEmployeeComponent(N_START_HOUR_2_EMPLOYEE) = CInt(oRecordset.Fields("StartHour2").Value)
				aEmployeeComponent(N_END_HOUR_2_EMPLOYEE) = CInt(oRecordset.Fields("EndHour2").Value)
				aEmployeeComponent(N_START_HOUR_3_EMPLOYEE) = CInt(oRecordset.Fields("StartHour3").Value)
				aEmployeeComponent(N_END_HOUR_3_EMPLOYEE) = CInt(oRecordset.Fields("EndHour3").Value)
				aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) = CDbl(oRecordset.Fields("WorkingHours").Value)
				aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) = CLng(oRecordset.Fields("LevelID").Value)
				aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = CLng(oRecordset.Fields("StatusID").Value)
				aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) = CLng(oRecordset.Fields("PaymentCenterID").Value)
				aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) = CLng(oRecordset.Fields("RiskLevel").Value)
				aEmployeeComponent(S_EMAIL_EMPLOYEE) = CStr(oRecordset.Fields("EmployeeEmail").Value)
				aEmployeeComponent(S_SSN_EMPLOYEE) = CStr(oRecordset.Fields("SocialSecurityNumber").Value)
				aEmployeeComponent(N_BIRTH_DATE_EMPLOYEE) = CLng(oRecordset.Fields("BirthDate").Value)
				aEmployeeComponent(N_START_DATE_EMPLOYEE) = CLng(oRecordset.Fields("StartDate").Value)
				aEmployeeComponent(N_START_DATE2_EMPLOYEE) = CLng(oRecordset.Fields("StartDate2").Value)
				aEmployeeComponent(N_COUNTRY_ID_EMPLOYEE) = CLng(oRecordset.Fields("CountryID").Value)
				aEmployeeComponent(S_RFC_EMPLOYEE) = CStr(oRecordset.Fields("RFC").Value)
				aEmployeeComponent(S_CURP_EMPLOYEE) = CStr(oRecordset.Fields("CURP").Value)
				aEmployeeComponent(N_GENDER_ID_EMPLOYEE) = CInt(oRecordset.Fields("GenderID").Value)
				aEmployeeComponent(N_MARITAL_STATUS_ID_EMPLOYEE) = CLng(oRecordset.Fields("MaritalStatusID").Value)
				aEmployeeComponent(N_ANTIQUITY_EMPLOYEE) = CLng(oRecordset.Fields("AntiquityID").Value)
				aEmployeeComponent(N_ANTIQUITY2_EMPLOYEE) = CLng(oRecordset.Fields("Antiquity2ID").Value)
				aEmployeeComponent(N_ANTIQUITY3_EMPLOYEE) = CLng(oRecordset.Fields("Antiquity3ID").Value)
				aEmployeeComponent(N_ANTIQUITY4_EMPLOYEE) = CLng(oRecordset.Fields("Antiquity4ID").Value)
				aEmployeeComponent(N_ACTIVE_EMPLOYEE) = CInt(oRecordset.Fields("Active").Value)
			End If
			oRecordset.Close
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo obtener la información del empleado."
				If aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 10 And aEmployeeComponent(N_STATUS_ID_EMPLOYEE) = 1 Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=13) Order By EmployeeDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = CLng(oRecordset.Fields("EndDate").Value)
						End If
						oRecordset.Close
					End If
				ElseIf aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = 0  Or aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 28 Then
					If aEmployeeComponent(N_REASON_ID_EMPLOYEE) > 0 Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ") Order By EmployeeDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If oRecordset.EOF Then
								If (aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 28) Or (aEmployeeComponent(N_REASON_ID_EMPLOYEE) >= 36) And (aEmployeeComponent(N_REASON_ID_EMPLOYEE) <= 41) Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID<>0) And (ReasonID<>58) Order By EmployeeDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										If Not oRecordset.EOF Then
											aEmployeeComponent(N_REASON_ID_EMPLOYEE) = CLng(oRecordset.Fields("ReasonID").Value)
											aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE) = CLng(oRecordset.Fields("JobID").Value)
											aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeDate").Value)
											aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = CLng(oRecordset.Fields("EndDate").Value)
											aEmployeeComponent(S_COMMENTS_EMPLOYEE) = CStr(oRecordset.Fields("Comments").Value)
										End If
										oRecordset.Close
									End If
								End If
							End If
						End If
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesHistoryList Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID<>0) And (ReasonID<>58) Order By EmployeeDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					End If
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							aEmployeeComponent(N_REASON_ID_EMPLOYEE) = CLng(oRecordset.Fields("ReasonID").Value)
							aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE) = CLng(oRecordset.Fields("JobID").Value)
							aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeDate").Value)
							aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) = CLng(oRecordset.Fields("EndDate").Value)
							aEmployeeComponent(S_COMMENTS_EMPLOYEE) = CStr(oRecordset.Fields("Comments").Value)
							If aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 51 Then
								aEmployeeComponent(N_SERVICE_ID_EMPLOYEE) = CLng(oRecordset.Fields("ServiceID").Value)
								aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) = CLng(oRecordset.Fields("JourneyID").Value)
								aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) = CLng(oRecordset.Fields("ShiftID").Value)
							End If
						End If
						oRecordset.Close
						If (aEmployeeComponent(N_REASON_ID_EMPLOYEE) = 26) And (aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE) <> -1) Then
							sErrorDescription = "No se pudo obtener el empleado con el que se permutará la plaza."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Employees Where (JobID=" & aEmployeeComponent(N_JOB_ID_HISTORY_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									aEmployeeComponent(N_ID_EMPLOYEE_2) = CLng(oRecordset.Fields("EmployeeID").Value)
								End If
								oRecordset.Close
							End If
						Else
							aEmployeeComponent(N_ID_EMPLOYEE_2) = ""
						End If
					End If
				End If

				sErrorDescription = "No se pudo obtener la información del empleado."
				If CLng(oRequest("ReasonID").Item) > 0 Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Reasons Where (ReasonID=" & CLng(oRequest("ReasonID").Item) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							aEmployeeComponent(N_REASON_TYPE_ID_EMPLOYEE) = CLng(oRecordset.Fields("ReasonTypeID").Value)
						End If
						oRecordset.Close
					End If
				End If

				If (aEmployeeComponent(N_STATUS_ID_EMPLOYEE) <> 0) And (aEmployeeComponent(N_STATUS_ID_EMPLOYEE) <> 1) Then
					sErrorDescription = "No se pudo obtener la información del empleado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From StatusEmployees Where (StatusID=" & aEmployeeComponent(N_STATUS_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							aEmployeeComponent(N_STATUS_REASON_ID_EMPLOYEE) = CLng(oRecordset.Fields("StatusReasonID").Value)
						End If
						oRecordset.Close
					End If
				Else
					aEmployeeComponent(N_STATUS_REASON_ID_EMPLOYEE) = -1
				End If

				sErrorDescription = "No se pudo obtener la información del empleado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesExtraInfo Where EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE), "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						aEmployeeComponent(S_ADDRESS_EMPLOYEE) = CStr(oRecordset.Fields("EmployeeAddress").Value)
						aEmployeeComponent(S_CITY_EMPLOYEE) = CStr(oRecordset.Fields("EmployeeCity").Value)
						aEmployeeComponent(S_ZIP_CODE_EMPLOYEE) = CStr(oRecordset.Fields("EmployeeZipCode").Value)
						aEmployeeComponent(N_ADDRESS_STATE_ID_EMPLOYEE) = CLng(oRecordset.Fields("StateID").Value)
						aEmployeeComponent(N_ADDRESS_COUNTRY_ID_EMPLOYEE) = CLng(oRecordset.Fields("CountryID").Value)
						aEmployeeComponent(S_EMPLOYEE_PHONE_EMPLOYEE) = CStr(oRecordset.Fields("EmployeePhone").Value)
						aEmployeeComponent(S_OFFICE_PHONE_EMPLOYEE) = CStr(oRecordset.Fields("OfficePhone").Value)
						aEmployeeComponent(S_EXT_OFFICE_EMPLOYEE) = CStr(oRecordset.Fields("OfficeExt").Value)
						aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE) = CStr(oRecordset.Fields("DocumentNumber1").Value)
						aEmployeeComponent(S_DOCUMENT_NUMBER_2_EMPLOYEE) = CStr(oRecordset.Fields("DocumentNumber2").Value)
						aEmployeeComponent(S_DOCUMENT_NUMBER_3_EMPLOYEE) = CStr(oRecordset.Fields("DocumentNumber3").Value)
						aEmployeeComponent(N_ACTIVITY_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeActivityID").Value)
						aEmployeeComponent(S_EMPLOYEE_BIRTHPLACE) = CStr(oRecordset.Fields("BirthPlace").Value)
						aEmployeeComponent(S_EMPLOYEE_LANGUAGES) = CStr(oRecordset.Fields("Languages").Value)
						aEmployeeComponent(S_EMPLOYEE_BLOODTYPE) = CStr(oRecordset.Fields("BloodType").Value)
						aEmployeeComponent(S_EMPLOYEE_CELLPHONE) = CStr(oRecordset.Fields("CellPhone").Value)
						aEmployeeComponent(S_EMPLOYEE_DEATH_BENEFICIARY) = CStr(oRecordset.Fields("DeathBeneficiary").Value)
						aEmployeeComponent(S_EMPLOYEE_DEATH_BENEFICIARY2) = CStr(oRecordset.Fields("DeathBeneficiary2").Value)
					End If
					oRecordset.Close
				End If
				sErrorDescription = "No se pudo obtener la información del empleado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesSchoolLevelsLKP Where EmployeeID =" & aEmployeeComponent(N_ID_EMPLOYEE), "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_ID) = CLng(oRecordset.Fields("SchoolarShipID").Value)
						aEmployeeComponent(S_EMPLOYEE_SCHOOLNAME) = CStr(oRecordset.Fields("SchoolName").Value)
						aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_DATE) = CLng(oRecordset.Fields("StartDate").Value)
						aEmployeeComponent(N_EMPLOYEE_SCHOOLARSHIP_DATE_END) = CLng(oRecordset.Fields("EndDate").Value)
						aEmployeeComponent(S_EMPLOYEE_SPECIALISM) = CStr(oRecordset.Fields("Specialism").Value)
					End If
					oRecordset.Close
				End If
				sErrorDescription = "No se pudo obtener la información del empleado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesReasonsLKP Where EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & " And (EmployeeDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And (ReasonID=" & aEmployeeComponent(N_REASON_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						aEmployeeComponent(S_REASON_FOR_REJECTION_COMMENTS_EMPLOYEE) = CStr(oRecordset.Fields("Comments").Value)
					End If
					oRecordset.Close
				End If

				sErrorDescription = "No se pudo obtener la información del riesgo profesional que tiene el empleado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesRisksLKP Where EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE), "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						aEmployeeComponent(N_RISK_LEVEL_EMPLOYEE) = CLng(oRecordset.Fields("RiskLevel").Value)
					End If
					oRecordset.Close
				End If

				sErrorDescription = "No se pudo obtener la información del empleado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas.AreaID, EconomicZoneID, Zones.ZoneID, ZoneTypeID, Jobs.StatusID, Positions.PositionID, PositionTypeID From Areas, Jobs, Zones, Positions Where (Areas.AreaID=Jobs.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Jobs.PositionID=Positions.PositionID) And (JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						aEmployeeComponent(N_ECONOMIC_ZONE_ID_EMPLOYEE) = CLng(oRecordset.Fields("EconomicZoneID").Value)
						aEmployeeComponent(N_GEOGRAPHICAL_ZONE_ID_EMPLOYEE) = CLng(oRecordset.Fields("ZoneTypeID").Value)
						aEmployeeComponent(N_POSITION_TYPE2_ID_EMPLOYEE) = CLng(oRecordset.Fields("PositionTypeID").Value)
						aEmployeeComponent(N_JOB_STATUS_ID_EMPLOYEE) = CLng(oRecordset.Fields("StatusID").Value)
						aEmployeeComponent(N_ZONE_ID_EMPLOYEE) = CLng(oRecordset.Fields("ZoneID").Value)
						aEmployeeComponent(N_AREA_ID_EMPLOYEE) = CLng(oRecordset.Fields("AreaID").Value)
						aEmployeeComponent(N_POSITION_ID_EMPLOYEE) = CLng(oRecordset.Fields("PositionID").Value)
					End If
					oRecordset.Close
				End If

				Select Case lReasonID
					Case 53
						lConceptID = 4
						sErrorDescription = "No se pudo obtener la información del empleado de honorarios."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & lConceptID & ") Order by StartDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = lConceptID
								aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = CLng(oRecordset.Fields("StartDate").Value)
								aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = CLng(oRecordset.Fields("EndDate").Value)
								aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) = CLng(oRecordset.Fields("CurrencyID").Value)
								aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = CInt(oRecordset.Fields("ConceptQttyID").Value)
								aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = CInt(oRecordset.Fields("ConceptTypeID").Value)
								aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) = CDbl(oRecordset.Fields("ConceptMin").Value)
								aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) = CInt(oRecordset.Fields("ConceptMinQttyID").Value)
								aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) =  CDbl(oRecordset.Fields("ConceptMax").Value)
								aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) = CInt(oRecordset.Fields("ConceptMaxQttyID").Value)
								aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = CLng(oRecordset.Fields("AppliesToID").Value)
								aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) = CInt(oRecordset.Fields("AbsenceTypeID").Value)
								aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = CInt(oRecordset.Fields("Active").Value)
								aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = CStr(oRecordset.Fields("UploadedFileName").Value)
								aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = CStr(oRecordset.Fields("Comments").Value)
								aEmployeeComponent(S_REASON_FOR_REJECTION_COMMENTS_EMPLOYEE) = CStr(oRecordset.Fields("Comments").Value)
								aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(oRecordset.Fields("ConceptAmount").Value)
							End If
							oRecordset.Close
						End If
'					Case Else
'						lConceptID = 13
'						sErrorDescription = "No se pudo obtener la información del empleado de honorarios."
'						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & lConceptID & ") And (StartDate=" & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ") And (EndDate= " & aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
'						If lErrorNumber = 0 Then
'							If Not oRecordset.EOF Then
'								aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = lConceptID
'								aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE)
'								aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = aEmployeeComponent(N_EMPLOYEE_END_DATE_EMPLOYEE)
'								aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) = 0
'								aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 1
'								aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = 3
'								aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) = 0
'								aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) = 1
'								aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) = 0
'								aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) = 1
'								aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = 11
'								aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) = 1
'								aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = 1
'								aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = ""
'								aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = "Trámite realizado por ingreso."
'								aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(oRecordset.Fields("ConceptAmount").Value)
'							End If
'							oRecordset.Close
'						End If
				End Select

				If aEmployeeComponent(N_JOB_ID_EMPLOYEE) <> -1 Then
					sErrorDescription = "No se pudo obtener la información de la plaza."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Jobs Where (JobID=" & aEmployeeComponent(N_JOB_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							aJobComponent(S_NUMBER_JOB) = CStr(oRecordset.Fields("JobNumber").Value)
							aJobComponent(N_ID_OWNER_JOB) = CLng(oRecordset.Fields("OwnerID").Value)
							aJobComponent(N_COMPANY_ID_JOB) = CLng(oRecordset.Fields("CompanyID").Value)
							aJobComponent(N_ZONE_ID_JOB) = CLng(oRecordset.Fields("ZoneID").Value)
							aJobComponent(N_AREA_ID_JOB) = CLng(oRecordset.Fields("AreaID").Value)
							aJobComponent(N_PAYMENT_CENTER_ID_JOB) = CLng(oRecordset.Fields("PaymentCenterID").Value)
							aJobComponent(N_POSITION_ID_JOB) = CLng(oRecordset.Fields("PositionID").Value)
							aJobComponent(N_JOB_TYPE_ID_JOB) = CLng(oRecordset.Fields("JobTypeID").Value)
							aJobComponent(N_SHIFT_ID_JOB) = CLng(oRecordset.Fields("ShiftID").Value)
							aJobComponent(N_JOURNEY_ID_JOB) = CLng(oRecordset.Fields("JourneyID").Value)
							aJobComponent(N_CLASSIFICATION_ID_JOB) = CLng(oRecordset.Fields("ClassificationID").Value)
							aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) = CLng(oRecordset.Fields("GroupGradeLevelID").Value)
							aJobComponent(N_INTEGRATION_ID_JOB) = CLng(oRecordset.Fields("IntegrationID").Value)
							aJobComponent(N_OCCUPATION_TYPE_ID_JOB) = CLng(oRecordset.Fields("OccupationTypeID").Value)
							aJobComponent(N_SERVICE_ID_JOB) = CLng(oRecordset.Fields("ServiceID").Value)
							aJobComponent(N_LEVEL_ID_JOB) = CLng(oRecordset.Fields("LevelID").Value)
							aJobComponent(D_WORKING_HOURS_JOB) = CDbl(oRecordset.Fields("WorkingHours").Value)
							aJobComponent(N_START_DATE_JOB) = CLng(oRecordset.Fields("StartDate").Value)
							aJobComponent(N_END_DATE_JOB) = CLng(oRecordset.Fields("EndDate").Value)
							aJobComponent(N_STATUS_ID_JOB) = CLng(oRecordset.Fields("StatusID").Value)
							aJobComponent(N_ACTIVE_JOB) = CInt(oRecordset.Fields("Active").Value)
						End If
						oRecordset.Close
						If lReasonID = 51 Then
							sErrorDescription = "No se pudo obtener la información de la plaza."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesHistoryList Where (bProcessed=2) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ReasonID=" & lReasonID & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									aEmployeeComponent(N_AREA_ID_EMPLOYEE) = CLng(oRecordset.Fields("AreaID").Value)
									aEmployeeComponent(N_PAYMENT_CENTER_ID_EMPLOYEE) = CLng(oRecordset.Fields("PaymentCenterID").Value)
								End If
							End If
						End If
					End If

					If Len(aJobComponent(N_POSITION_ID_JOB)) > 0 Then
						If aJobComponent(N_POSITION_ID_JOB) <> -1 Then
							sErrorDescription = "No se pudo obtener el tipo de empleado del puesto."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Positions Where (PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									aJobComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeTypeID").Value)
									aJobComponent(N_POSITION_TYPE_ID_JOB) = CLng(oRecordset.Fields("PositionTypeID").Value)
								End If
								oRecordset.Close
							End If
						End If
					End If
				End If

				If aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) > -1 Then
					sErrorDescription = "No se pudo obtener la información del empleado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesBeneficiariesLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (BeneficiaryID=" & aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							aEmployeeComponent(N_END_DATE_BENEFICIARY_EMPLOYEE) = CLng(oRecordset.Fields("EndDate").Value)
							aEmployeeComponent(S_NUMBER_BENEFICIARY_EMPLOYEE) = CLng(oRecordset.Fields("BeneficiaryNumber").Value)
							aEmployeeComponent(S_NAME_BENEFICIARY_EMPLOYEE) = CStr(oRecordset.Fields("BeneficiaryName").Value)
							aEmployeeComponent(S_LAST_NAME_BENEFICIARY_EMPLOYEE) = CStr(oRecordset.Fields("BeneficiaryLastName").Value)
							aEmployeeComponent(S_LAST_NAME2_BENEFICIARY_EMPLOYEE) = CStr(oRecordset.Fields("BeneficiaryLastName2").Value)
							aEmployeeComponent(N_BIRTH_DATE_BENEFICIARY_EMPLOYEE) = CLng(oRecordset.Fields("BeneficiaryBirthDate").Value)
							aEmployeeComponent(D_ALIMONY_AMOUNT_BENEFICIARY_EMPLOYEE) = CDbl(oRecordset.Fields("AlimonyAmount").Value)
							aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE) = CLng(oRecordset.Fields("AlimonyTypeID").Value)
							aEmployeeComponent(N_PAYMENT_CENTER_ID_BENEFICIARY_EMPLOYEE) = CLng(oRecordset.Fields("PaymentCenterID").Value)
							aEmployeeComponent(S_COMMENTS_BENEFICIARY_EMPLOYEE) = CStr(oRecordset.Fields("Comments").Value)
						End If
						oRecordset.Close
					End If
				End If

				'If aEmployeeComponent(N_ID_CHILD_EMPLOYEE) > -1 Then
					sErrorDescription = "No se pudo obtener la información del empleado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesChildrenLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ChildID=" & aEmployeeComponent(N_ID_CHILD_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							aEmployeeComponent(N_ID_CHILD_EMPLOYEE) = CLng(oRecordset.Fields("ChildID").Value)
							aEmployeeComponent(S_NAME_CHILD_EMPLOYEE) = CStr(oRecordset.Fields("ChildName").Value)
							aEmployeeComponent(S_LAST_NAME_CHILD_EMPLOYEE) = CStr(oRecordset.Fields("ChildLastName").Value)
							aEmployeeComponent(S_LAST_NAME2_CHILD_EMPLOYEE) = CStr(oRecordset.Fields("ChildLastName2").Value)
							aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE) = CLng(oRecordset.Fields("ChildBirthDate").Value)
							aEmployeeComponent(N_END_DATE_CHILD_EMPLOYEE) = CLng(oRecordset.Fields("ChildEndDate").Value)
							aEmployeeComponent(N_CHILD_LEVEL_ID_EMPLOYEE) = CLng(oRecordset.Fields("LevelID").Value)
						End If
						oRecordset.Close
					End If
				'End If

				If aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) > -1 Then
					sErrorDescription = "No se pudo obtener la información del crédito del empleado."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Credits Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (CreditID=" & aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = CLng(oRecordset.Fields("CreditTypeID").Value)
							aEmployeeComponent(S_CREDIT_CONTRACT_NUMBER_EMPLOYEE) = CStr(oRecordset.Fields("ContractNumber").Value)
							aEmployeeComponent(S_CREDIT_ACCOUNT_NUMBER_EMPLOYEE) = CStr(oRecordset.Fields("AccountNumber").Value)
							aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) = CStr(oRecordset.Fields("PaymentsNumber").Value)
							aEmployeeComponent(N_CREDIT_PERIOD_ID_EMPLOYEE) = CLng(oRecordset.Fields("PeriodID").Value)
							aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = CLng(oRecordset.Fields("StartDate").Value)
							aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = CLng(oRecordset.Fields("EndDate").Value)
							aEmployeeComponent(L_CREDIT_FINISH_DATE_EMPLOYEE) = CLng(oRecordset.Fields("FinishDate").Value)
							aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(oRecordset.Fields("StartAmount").Value)
						End If
						oRecordset.Close
					End If
				End If
			End If
		End If

		'aEmployeeComponent(S_RELATED_EMPLOYEE) = ""
		'sErrorDescription = "No se pudo obtener la información del empleado."
		'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RelatedEmployeeID From EmployeesLKP Where EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE), "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		'If lErrorNumber = 0 Then
		'	Do While Not oRecordset.EOF
		'		aEmployeeComponent(S_RELATED_EMPLOYEE) = aEmployeeComponent(S_RELATED_EMPLOYEE) & CStr(oRecordset.Fields("RelatedEmployeeID").Value) & ","
		'		oRecordset.MoveNext
		'		If Err.number <> 0 Then Exit Do
		'	Loop
		'	oRecordset.Close
		'	If Len(aEmployeeComponent(S_RELATED_EMPLOYEE)) > 0 Then aEmployeeComponent(S_RELATED_EMPLOYEE) = Left(aEmployeeComponent(S_RELATED_EMPLOYEE), (Len(aEmployeeComponent(S_RELATED_EMPLOYEE)) - Len(",")))
		'End If
	End If

	Set oRecordset = Nothing
	GetEmployee = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeAbsencesAppliesToID(oRequest, oADODBConnection, aEmployeeComponent, sEmployeeAbsenceIDs, sErrorDescription)
'************************************************************
'Purpose: To get the absences requerids for insert the
'         absence for employee
'Inputs:  oRequest, oADODBConnection, aAbsenceComponent
'Outputs: sAbsenceIDs, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeAbsencesAppliesToID"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información de la incidencia."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Absences Where (AbsenceID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ")", "AbsenceComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "El tipo de incidencia especificada no se encuentra en el sistema."
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
		Else
			sEmployeeAbsenceIDs = CStr(oRecordset.Fields("AppliesToID").Value)
		End If
		oRecordset.Close
	Else
		sErrorDescription = "Error al validar las incidencias registradas."
	End If

	Set oRecordset = Nothing
	GetEmployeeAbsencesAppliesToID = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeAdjustments(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a concept adjustment for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeAdjustments"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del concepto para obtener la información del reclamo."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del crédito del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAdjustmentsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (MissingDate=" & aEmployeeComponent(N_MISSING_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El crédito especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(oRecordset.Fields("ConceptAmount").Value)
				aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) = CDbl(oRecordset.Fields("PayrollDate").Value)
				aEmployeeComponent(S_NAME_BENEFICIARY_EMPLOYEE) = CStr(oRecordset.Fields("BeneficiaryName").Value)
				aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = CInt(oRecordset.Fields("Active").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetEmployeeAdjustments = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeBankAccount(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a Sundays for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeBankAccount"
	Dim oRecordset
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
		sErrorDescription = "No se pudo obtener la información de la cuenta del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From BankAccounts Where (AccountID=" & aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "La cuenta especificada no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
				aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = CLng(oRecordset.Fields("StartDate").Value)
				aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = CLng(oRecordset.Fields("EndDate").Value)
				aEmployeeComponent(N_BANK_ID_EMPLOYEE) = CDbl(oRecordset.Fields("BankID").Value)
				aEmployeeComponent(S_ACCOUNT_NUMBER_EMPLOYEE) = CStr(oRecordset.Fields("AccountNumber").Value)
				aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = CInt(oRecordset.Fields("Active").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetEmployeeBankAccount = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeBeneficiary(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a concept for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeBeneficiary"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) = -1) Or (aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del beneficiario su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del beneficiario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesBeneficiariesLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (BeneficiaryID=" & aEmployeeComponent(N_ID_BENEFICIARY_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(N_START_DATE_BENEFICIARY_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El beneficiario especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aEmployeeComponent(N_END_DATE_BENEFICIARY_EMPLOYEE)	= CLng(oRecordset.Fields("EndDate").Value)
				aEmployeeComponent(S_NUMBER_BENEFICIARY_EMPLOYEE) = CLng(oRecordset.Fields("BeneficiaryNumber").Value)
				aEmployeeComponent(S_NAME_BENEFICIARY_EMPLOYEE)	= CStr(oRecordset.Fields("BeneficiaryName").Value)
				aEmployeeComponent(S_LAST_NAME_BENEFICIARY_EMPLOYEE) = CStr(oRecordset.Fields("BeneficiaryLastName").Value)
				aEmployeeComponent(S_LAST_NAME2_BENEFICIARY_EMPLOYEE) = CStr(oRecordset.Fields("BeneficiaryLastName2").Value)
				aEmployeeComponent(N_BIRTH_DATE_BENEFICIARY_EMPLOYEE) = CLng(oRecordset.Fields("BeneficiaryBirthDate").Value)
				aEmployeeComponent(N_ALIMONY_TYPE_ID_BENEFICIARY_EMPLOYEE) = CInt(oRecordset.Fields("AlimonyTypeID").Value)
				aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(oRecordset.Fields("ConceptAmount").Value)
				aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) = CDbl(oRecordset.Fields("ConceptMin").Value)
				aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) = CInt(oRecordset.Fields("ConceptMinQttyID").Value)
				aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) = CDbl(oRecordset.Fields("ConceptMax").Value)
				aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) = CInt(oRecordset.Fields("ConceptMaxQttyID").Value)
				aEmployeeComponent(N_PAYMENT_CENTER_ID_BENEFICIARY_EMPLOYEE) = (oRecordset.Fields("PaymentCenterID").Value)
				aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE)	= CStr(oRecordset.Fields("Comments").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetEmployeeBeneficiary = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeByStatus(oRequest, oADODBConnection, sStatusID, sAction, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a
'         employee from the database
'Inputs:  oRequest, oADODBConnection, lStatusID
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeByStatus"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información del empleado con el estatus indicado."
	If (InStr(1, sAction, "EmployeesNew", vbBinaryCompare) > 0) Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.* From Employees, StatusEmployees Where (Employees.StatusID=StatusEmployees.StatusID) And (Employees.StatusID In (-2,-3,-4)) And (StatusEmployees.StatusReasonID In (" & sStatusID & "))", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	ElseIf (InStr(1, sAction, "ForValidation", vbBinaryCompare) > 0) Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.* From Employees, StatusEmployees Where (Employees.StatusID=StatusEmployees.StatusID) And (StatusEmployees.StatusReasonID In (" & sStatusID & ")) And (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.* From Employees, StatusEmployees Where (Employees.StatusID=StatusEmployees.StatusID) And (StatusEmployees.StatusReasonID In (" & sStatusID & "))", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	End If
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			lErrorNumber = L_ERR_NO_RECORDS
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GetEmployeeByStatus = lErrorNumber
	Err.Clear
End Function

Function GetEmployeesGrades(oRequest, oADODBConnection, aEmployeeComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the employees from
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeesGrades"
	Dim sSort
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim iActive

	iActive = aEmployeeComponent(N_ACTIVE_EMPLOYEE)
	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If
	aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = ""
	If Len(aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE)) > 0 Then
		If InStr(1, aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE), "And ", vbBinaryCompare) <> 1 Then aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = "And " & aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE)
	End If
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) & " And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Areas.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	If aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) <> -1 Then aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) & " And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) & "," & S_WILD_CHAR & "')"
	sErrorDescription = "No se pudo obtener la información de los empleados."
	sQuery = "Select Employees.EmployeeID, EmployeeNumber, Employees.PaymentCenterID, EmployeeName + ' ' + EmployeeLastName  + ' ' + EmployeeLastName2 As EmployeeFullName," & _
			 " EmployeesGrades.StartDate, EmployeesGrades.EndDate, EmployeesGrades.PayrollID, EmployeesGrades.EmployeeGrade, Users.UserLastName + ' ' + Users.UserName As UserFullName," & _
			 " PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath" & _
			 " From Employees, EmployeesGrades, Users, Areas, Areas As PaymentCenters, Jobs, Zones As AreasZones, Zones As ParentZones, Zones, Companies" & _
			 " Where (Employees.JobID=Jobs.JobID)" & _
			 " And (Employees.PaymentCenterID=PaymentCenters.AreaID)" & _
			 " And (Jobs.AreaID=Areas.AreaID)" & _
			 " And (Areas.ZoneID=AreasZones.ZoneID)" & _
			 " And (AreasZones.ParentID=ParentZones.ZoneID)" & _
			 " And (PaymentCenters.ZoneID=Zones.ZoneID)" & _
			 " And (Employees.CompanyID=Companies.CompanyID)" & _
			 " And (Employees.PaymentCenterID=PaymentCenters.AreaID)" & _
			 " And (Employees.EmployeeID=EmployeesGrades.EmployeeID)" & _
			 " And (Users.UserID=EmployeesGrades.UserID)" & _
			 " And (Employees.CompanyID=Companies.CompanyID)"
	If CInt(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
		sQuery = sQuery & " And (EmployeesGrades.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
	Else
		If iActive Then
			sQuery = sQuery & " And (EmployeesGrades.EmployeeID=0)"
		End If
	End If
	sQuery = sQuery & aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) & " And (EmployeesGrades.Active=" & iActive & ") Order By Employees.EmployeeID, EmployeesGrades.StartDate"
	sErrorDescription = "No se pudieron obtener los registros de calificación para el empleado " & aEmployeeComponent(N_ID_EMPLOYEE)
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine
	If iActive = 0 Then Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & sQuery & """ />"

	GetEmployeesGrades = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeChildren(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about an employee's children from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeChildren"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Employees Where EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE), "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El empleado especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aEmployeeComponent(S_NUMBER_EMPLOYEE) = CStr(oRecordset.Fields("EmployeeNumber").Value)
				aEmployeeComponent(S_NAME_EMPLOYEE) = CStr(oRecordset.Fields("EmployeeName").Value)
				aEmployeeComponent(S_LAST_NAME_EMPLOYEE) = CStr(oRecordset.Fields("EmployeeLastName").Value)
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
					aEmployeeComponent(S_LAST_NAME2_EMPLOYEE) = CStr(oRecordset.Fields("EmployeeLastName2").Value)
				Else
					aEmployeeComponent(S_LAST_NAME2_EMPLOYEE) = " "
				End If
			End If
			oRecordset.Close

			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo obtener la información del empleado."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesChildrenLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						aEmployeeComponent(N_ID_CHILD_EMPLOYEE) = CLng(oRecordset.Fields("ChildID").Value)
						aEmployeeComponent(S_NAME_CHILD_EMPLOYEE) = CStr(oRecordset.Fields("ChildName").Value)
						aEmployeeComponent(S_LAST_NAME_CHILD_EMPLOYEE) = CStr(oRecordset.Fields("ChildLastName").Value)
						aEmployeeComponent(S_LAST_NAME2_CHILD_EMPLOYEE) = CStr(oRecordset.Fields("ChildLastName2").Value)
						aEmployeeComponent(N_BIRTH_DATE_CHILD_EMPLOYEE) = CLng(oRecordset.Fields("ChildBirthDate").Value)
						aEmployeeComponent(N_END_DATE_CHILD_EMPLOYEE) = CLng(oRecordset.Fields("ChildEndDate").Value)
						aEmployeeComponent(N_CHILD_LEVEL_ID_EMPLOYEE) = CLng(oRecordset.Fields("LevelID").Value)
					End If
					oRecordset.Close
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetEmployeeChildren = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a concept for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeConcept"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del concepto para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El concepto especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = CLng(oRecordset.Fields("StartDate").Value)
				aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = CLng(oRecordset.Fields("EndDate").Value)
				aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(oRecordset.Fields("ConceptAmount").Value)
				aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) = CLng(oRecordset.Fields("CurrencyID").Value)
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = CInt(oRecordset.Fields("ConceptQttyID").Value)
				aEmployeeComponent(N_CONCEPT_TYPE_ID_EMPLOYEE) = CInt(oRecordset.Fields("ConceptTypeID").Value)
				aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) = CDbl(oRecordset.Fields("ConceptMin").Value)
				aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) = CInt(oRecordset.Fields("ConceptMinQttyID").Value)
				aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) = CDbl(oRecordset.Fields("ConceptMax").Value)
				aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) = CInt(oRecordset.Fields("ConceptMaxQttyID").Value)
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = CStr(oRecordset.Fields("AppliesToID").Value)
				aEmployeeComponent(N_CONCEPT_ABSENCE_TYPE_ID_EMPLOYEE) = CInt(oRecordset.Fields("AbsenceTypeID").Value)
				aEmployeeComponent(N_CONCEPT_ORDER_EMPLOYEE) = CInt(oRecordset.Fields("ConceptOrder").Value)
				aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = CInt(oRecordset.Fields("Active").Value)
				aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = CStr(oRecordset.Fields("UploadedFileName").Value)
				aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = CStr(oRecordset.Fields("Comments").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetEmployeeConcept = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeCredit(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a concept for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeCredit"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del crédito para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del crédito del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Credits Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (CreditID=" & aEmployeeComponent(N_CREDIT_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El crédito especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = CLng(oRecordset.Fields("CreditTypeID").Value)
				aEmployeeComponent(S_CREDIT_CONTRACT_NUMBER_EMPLOYEE) = CStr(oRecordset.Fields("ContractNumber").Value)
				aEmployeeComponent(S_CREDIT_ACCOUNT_NUMBER_EMPLOYEE) = CStr(oRecordset.Fields("AccountNumber").Value)
				aEmployeeComponent(N_CREDIT_PAYMENTS_NUMBER_EMPLOYEE) = CInt(oRecordset.Fields("PaymentsNumber").Value)
				aEmployeeComponent(N_CREDIT_PERIOD_ID_EMPLOYEE) = CInt(oRecordset.Fields("PeriodID").Value)
				aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = CLng(oRecordset.Fields("StartDate").Value)
				aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE)	= CLng(oRecordset.Fields("EndDate").Value)
				aEmployeeComponent(L_CREDIT_FINISH_DATE_EMPLOYEE) = CLng(oRecordset.Fields("FinishDate").Value)
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = CInt(oRecordset.Fields("QttyID").Value)
				aEmployeeComponent(N_CONCEPT_APPLIES_TO_ID_EMPLOYEE) = CInt(oRecordset.Fields("AppliesToID").Value)
				aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(oRecordset.Fields("PaymentAmount").Value)
				aEmployeeComponent(N_CREDIT_PAYMENTS_COUNTER_EMPLOYEE) = CInt(oRecordset.Fields("PaymentsCounter").Value)
				aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = CInt(oRecordset.Fields("Active").Value)
				aEmployeeComponent(S_CONCEPT_FILE_NAME_EMPLOYEE) = CStr(oRecordset.Fields("UploadedFileName").Value)
				aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE)	= CStr(oRecordset.Fields("Comments").Value)
				aEmployeeComponent(N_CREDIT_UPLOADED_RECORD_TYPE) = CInt(oRecordset.Fields("UploadedRecordType").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetEmployeeCredit = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeCreditor(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a concept for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeCreditor"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_ID_CREDITOR_EMPLOYEE) = -1) Or (aEmployeeComponent(N_START_DATE_CREDITOR_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del beneficiario su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del beneficiario."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesCreditorsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (CreditorID=" & aEmployeeComponent(N_ID_CREDITOR_EMPLOYEE) & ") And (StartDate=" & aEmployeeComponent(N_START_DATE_CREDITOR_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El beneficiario especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aEmployeeComponent(N_END_DATE_CREDITOR_EMPLOYEE)	= CLng(oRecordset.Fields("EndDate").Value)
				aEmployeeComponent(S_NUMBER_CREDITOR_EMPLOYEE) = CLng(oRecordset.Fields("CreditorNumber").Value)
				aEmployeeComponent(S_NAME_CREDITOR_EMPLOYEE)	= CStr(oRecordset.Fields("CreditorName").Value)
				aEmployeeComponent(S_LAST_NAME_CREDITOR_EMPLOYEE) = CStr(oRecordset.Fields("CreditorLastName").Value)
				aEmployeeComponent(S_LAST_NAME2_CREDITOR_EMPLOYEE) = CStr(oRecordset.Fields("CreditorLastName2").Value)
				aEmployeeComponent(N_BIRTH_DATE_CREDITOR_EMPLOYEE) = CLng(oRecordset.Fields("CreditorBirthDate").Value)
				aEmployeeComponent(N_CREDITOR_TYPE_ID_EMPLOYEE) = CInt(oRecordset.Fields("CreditorTypeID").Value)
				aEmployeeComponent(D_CREDITOR_AMOUNT_EMPLOYEE) = CDbl(oRecordset.Fields("ConceptAmount").Value)
				aEmployeeComponent(D_CONCEPT_MIN_EMPLOYEE) = CDbl(oRecordset.Fields("ConceptMin").Value)
				aEmployeeComponent(N_CONCEPT_MIN_QTTY_ID_EMPLOYEE) = CInt(oRecordset.Fields("ConceptMinQttyID").Value)
				aEmployeeComponent(D_CONCEPT_MAX_EMPLOYEE) = CDbl(oRecordset.Fields("ConceptMax").Value)
				aEmployeeComponent(N_CONCEPT_MAX_QTTY_ID_EMPLOYEE) = CInt(oRecordset.Fields("ConceptMaxQttyID").Value)
				aEmployeeComponent(N_PAYMENT_CENTER_ID_CREDITOR_EMPLOYEE) = (oRecordset.Fields("PaymentCenterID").Value)
				aEmployeeComponent(S_COMMENTS_CREDITOR_EMPLOYEE)	= CStr(oRecordset.Fields("Comments").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetEmployeeCreditor = lErrorNumber
	Err.Clear
End Function

Function GetEmployeesDocument(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a document for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeesDocument"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para obtener la solicitud de hoja ùnica"
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información de la cuenta del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesDocs Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (DocumentDate=" & aEmployeeComponent(N_EMPLOYEE_DOCUMENT_DATE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "La cuenta especificada no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aEmployeeComponent(N_EMPLOYEE_DOCUMENT_TIME) = CLng(oRecordset.Fields("DocumentTime").Value)
				aEmployeeComponent(N_EMPLOYEE_DOCUMENT_DATE_2) = CLng(oRecordset.Fields("Document2Date").Value)
				aEmployeeComponent(N_EMPLOYEE_DOCUMENT_TIME_2) = CLng(oRecordset.Fields("Document2Time").Value)
				aEmployeeComponent(S_DOCUMENT_NUMBER_1_EMPLOYEE) = CStr(oRecordset.Fields("DocumentNumber").Value)
				aEmployeeComponent(S_EMPLOYEE_AUTHORIZERS) = CStr(oRecordset.Fields("Authorizers").Value)
				aEmployeeComponent(S_EMPLOYEE_AUTHORIZED) = CStr(oRecordset.Fields("Authorized").Value)
				aEmployeeComponent(N_EMPLOYEE_DOCUMENT_TYPE) = CInt(oRecordset.Fields("DocumentTypeID").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetEmployeesDocument = lErrorNumber
	Err.Clear
End Function

Function GetEmployeesDocuments(oRequest, oADODBConnection, aEmployeeComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the employees from
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeesDocuments"
	Dim sSort
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim iActive

	iActive = aEmployeeComponent(N_ACTIVE_EMPLOYEE)
	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If
	aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = ""
	If Len(aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE)) > 0 Then
		If InStr(1, aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE), "And ", vbBinaryCompare) <> 1 Then aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = "And " & aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE)
	End If
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) & " And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Areas.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	If aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) <> -1 Then aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) & " And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) & "," & S_WILD_CHAR & "')"
	sErrorDescription = "No se pudo obtener la información de los empleados."
	sQuery = "Select Employees.EmployeeID, EmployeeNumber, Employees.PaymentCenterID, EmployeeName + ' ' + EmployeeLastName  + ' ' + EmployeeLastName2 As EmployeeFullName," & _
			 " EmployeesDocs.DocumentDate, EmployeesDocs.Document2Date, EmployeesDocs.Document3Date, EmployeesDocs.DocumentNumber, Authorizers, Authorized, bPrinted, DocumentTypeID," & _
			 " ReportName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath, UserName + ' ' + UserLastName As UserFullName" & _
			 " From Employees, EmployeesDocs, Areas, Areas As PaymentCenters, Jobs, Zones As AreasZones, Zones As ParentZones, Zones, Companies, Users" & _
			 " Where (Employees.JobID=Jobs.JobID)" & _
			 " And (Employees.PaymentCenterID=PaymentCenters.AreaID)" & _
			 " And (Jobs.AreaID=Areas.AreaID)" & _
			 " And (Areas.ZoneID=AreasZones.ZoneID)" & _
			 " And (AreasZones.ParentID=ParentZones.ZoneID)" & _
			 " And (PaymentCenters.ZoneID=Zones.ZoneID)" & _
			 " And (Employees.CompanyID=Companies.CompanyID)" & _
			 " And (Employees.PaymentCenterID=PaymentCenters.AreaID)" & _
			 " And (Employees.EmployeeID=EmployeesDocs.EmployeeID)" & _
			 " And (Employees.CompanyID=Companies.CompanyID)" & _
			 " And (EmployeesDocs.UserID=Users.UserID)"
	If CInt(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
		sQuery = sQuery & " And (EmployeesDocs.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
	Else
		sQuery = sQuery & " And (EmployeesDocs.EmployeeID=0)"
	End If
	'sQuery = sQuery & " And (EmployeesDocs.StatusID=" & iActive & ")"
	sQuery = sQuery & aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) & " Order By Employees.EmployeeID, EmployeesDocs.DocumentDate"
	sErrorDescription = "No se pudieron obtener los registros de calificación para el empleado " & aEmployeeComponent(N_ID_EMPLOYEE)
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	'Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine
	'If iActive = 0 Then Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & sQuery & """ />"

	GetEmployeesDocuments = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeSalary(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a specific concept for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeSalary"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o la compañía o tipo de tabulador para obtener su salario."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del empleado."
		If aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 7 and aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 12  and aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) <> 13 Then
			sQuery = "Select C.ConceptAmount, J.PositionID, E.*" & _
				 " from Employees E, Jobs J, ConceptsValues C, Areas A" & _
				 " where E.JobID=J.JobID" & _
				 " and J.PositionID=C.PositionID" & _
				 " and E.GroupGradeLevelID=C.GroupGradeLevelID" & _
				 " and E.IntegrationID=C.IntegrationID" & _
				 " and E.ClassificationID=C.ClassificationID" & _
				 " and E.LevelID=C.LevelID" & _
				 " and E.WorkingHours=C.WorkingHours" & _
				 " and J.AreaID=A.AreaID" & _
				 " and C.EconomicZoneID=A.EconomicZoneID" & _
				 " and C.ConceptID=1" & _
				 " and C.CompanyID IN (-1," & aEmployeeComponent(N_COMPANY_ID_EMPLOYEE) & ")" & _
				 " and C.EmployeeTypeID In (-1," & aEmployeeComponent(N_EMPLOYEE_TYPE_ID_EMPLOYEE) & ")" & _
				 " and C.StartDate<" & Left(GetSerialNumberForDate(""), Len("00000000")) & _
				 " and C.EndDate>" & Left(GetSerialNumberForDate(""), Len("00000000")) & _
				 " and EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE)
		Else
			sQuery = "Select ConceptAmount From EmployeesConceptsLKP Where EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & " Order By EndDate Desc"
		End If
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = -1
				sErrorDescription = "No se pudo obtener el sueldo quincenal del empleado"
			Else
				aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(oRecordset.Fields("ConceptAmount").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetEmployeeSalary = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeSpecificConcept(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a specific concept for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeSpecificConcept"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o del concepto para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesConceptsLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (ConceptID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (StartDate<=" & CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) & ") And (EndDate <> 0) Order By StartDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = 0
				aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = 0
				aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = 0
				aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) = 0
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = 0
			Else
				aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) = CLng(oRecordset.Fields("StartDate").Value)
				aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = CLng(oRecordset.Fields("EndDate").Value)
				aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(oRecordset.Fields("ConceptAmount").Value)
				aEmployeeComponent(N_CONCEPT_CURRENCY_ID_EMPLOYEE) = CLng(oRecordset.Fields("CurrencyID").Value)
				aEmployeeComponent(N_CONCEPT_QTTY_ID_EMPLOYEE) = CInt(oRecordset.Fields("ConceptQttyID").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetEmployeeSpecificConcept = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeSuperiorPositionAmount(oRequest, oADODBConnection, aEmployeeComponent, lSuperiorAmount, bIsSuperior, sErrorDescription)
'************************************************************
'Purpose: To get the salary of superior position of employee
'Inputs:  oRequest, oADODBConnection, aEmployeeComponent
'Outputs: lSuperiorAmount, bIsSuperior, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeSuperiorPositionAmount"
	Dim oRecordset
	Dim lErrorNumber
	Dim sQuery
	Dim iBranchID
	Dim iSubBranchID
	Dim iHierarchyID
	Dim sWorkingHours

	sQuery = "Select * from Positions where (PositionID=" & aEmployeeComponent(N_POSITION_ID_EMPLOYEE) & ")"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			iBranchID = CInt(oRecordset.Fields("BranchID").Value)
			iSubBranchID = CInt(oRecordset.Fields("SubBranchID").Value)
			sWorkingHours = CInt(oRecordset.Fields("WorkingHours").Value)
			iHierarchyID = CInt(oRecordset.Fields("HierarchyID").Value)
			sQuery = "Select * from Positions where (BranchID=" & iBranchID & ")" & _
					 " And (SubBranchID=" & iSubBranchID & ")" & _
					 " And (WorkingHours=" & sWorkingHours & ")" & _
					 " And (HierarchyID>" & iHierarchyID & ")" & _
					 " Order By HierarchyID"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					aEmployeeComponent(N_POSITION_ID_EMPLOYEE) = CInt(oRecordset.Fields("PositionID").Value)
					lErrorNumber = GetEmployeeSalary(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
					If lErrorNumber = 0 Then
						If oRecordset.EOF Then
							lErrorNumber = -1
							sErrorDescription = "No se pudo encontrar el sueldo del puesto superior"
						Else
							bIsSuperior = True
						End If
					End If
				Else
					sQuery = "Select * from Positions where (BranchID=" & iBranchID & ")" & _
							 " And (SubBranchID=" & iSubBranchID & ")" & _
							 " And (WorkingHours=" & sWorkingHours & ")" & _
							 " And (HierarchyID<" & iHierarchyID & ")" & _
							 " Order By HierarchyID"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							aEmployeeComponent(N_POSITION_ID_EMPLOYEE) = CInt(oRecordset.Fields("PositionID").Value)
							lErrorNumber = GetEmployeeSalary(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
							If lErrorNumber = 0 Then
								If oRecordset.EOF Then
									lErrorNumber = -1
									sErrorDescription = "No se pudo encontrar el sueldo del puesto inferior"
								Else
									bIsSuperior = True
								End If
							End If
						Else
							lErrorNumber = -1
							sErrorDescription = "No se pudo obtener el puesto inmediato superior o inferior en jerarquia para el empleado " & aEmployeeComponent(N_ID_EMPLOYEE)
						End If
					End If					
				End If
			End If
		End If
	End If
	Set oRecordset = Nothing
	GetEmployeeSuperiorPositionAmount = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeJourneyType(oRequest, oADODBConnection, aEmployeeComponent, iJourneyTypeID, sErrorDescription)
'************************************************************
'Purpose: To get the information about an employee's children from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeJourneyType"
	Dim oRecordset
	Dim lErrorNumber
	Dim sQuery

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		aEmployeeComponent(N_ID_EMPLOYEE) = aAbsenceComponent(N_EMPLOYEE_ID_ABSENCE)
		lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
		sErrorDescription = "No existe empleado con el número indicado."
		If lErrorNumber = 0 Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Employees Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sErrorDescription = "No se pudo obtener la información de la jornada."
					'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Shifts Where (ShiftID=" & CInt(oRecordset.Fields("ShiftID").Value) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					iJourneyTypeID = CInt(oRecordset.Fields("JourneyID").Value)
					If (iJourneyTypeID > 0) And (iJourneyTypeID < 20) Then
						iJourneyTypeID = 1
					End If
				End If
				oRecordset.Close
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetEmployeeJourneyType = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeNumberFromRFC(oRequest, oADODBConnection, bUseLike, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the employee number from a employee with the
'         employee RFC from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeNumberFromRFC"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim sCondition

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If Len(aEmployeeComponent(S_RFC_EMPLOYEE)) <= 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el RFC del empleado."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sQuery = "Select * From Employees"
		If Len(aEmployeeComponent(S_RFC_EMPLOYEE)) = 13 Then
			sCondition = " Where RFC='" & aEmployeeComponent(S_RFC_EMPLOYEE) & "'"
		Else
			If bUseLike Then 
				sCondition = " Where RFC like '" & aEmployeeComponent(S_RFC_EMPLOYEE) & "%'"
			Else
				sCondition = " Where RFC='" & aEmployeeComponent(S_RFC_EMPLOYEE) & "'"
				lErrorNumber = -1
				sErrorDescription = "El RFC proporcionado no es de 13 posiciones."
			End If
		End If
		sQuery = sQuery & sCondition
		sErrorDescription = "No se pudo obtener el número del empleado con RFC = " & aEmployeeComponent(S_RFC_EMPLOYEE)
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "No existe algún empleado registrado con RFC = " & aEmployeeComponent(S_RFC_EMPLOYEE)
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aEmployeeComponent(N_ID_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeID").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetEmployeeNumberFromRFC = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeStartDate(oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the date of entry to the Institute
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeStartDate"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información del empleado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeDate From EmployeesHistoryList Where (ReasonID <> 0) And (ReasonID <> 58) And (EmployeesHistoryList.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) &") Order By EmployeeDate Asc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			aEmployeeComponent(N_START_DATE_EMPLOYEE) = CLng(oRecordset.Fields("EmployeeDate").Value)
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GetEmployeeStartDate = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeStatusToValidateTheMovement(oADODBConnection, lReasonID, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get employee status to validate the movement
'Inputs:  oRequest, oADODBConnection, lStatusID
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeStatusToValidateTheMovement"
	Dim oRecordset
	Dim lErrorNumber
	Dim sStatusEmployeesIDs

	sErrorDescription = "No se pudo obtener el estatus requisito que debe tener el empleado para el movimiento solicitado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusEmployeesIDs From Reasons Where (ReasonID = " & lReasonID & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sStatusEmployeesIDs = "," & CStr(oRecordset.Fields("StatusEmployeesIDs").Value) & ","
			If (InStr(1, sStatusEmployeesIDs, "," & CStr(aEmployeeComponent(N_STATUS_ID_EMPLOYEE)) & ",", vbBinaryCompare) = 0) Then
				lErrorNumber = -1
			End If
		End If
	End If
	
	Set oRecordset = Nothing
	GetEmployeeStatusToValidateTheMovement = lErrorNumber
	Err.Clear
End Function

Function GetEmployeeSundays(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a Sundays for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeeSundays"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Or (aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado o de la fecha de ocurrencia del domingo."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesAbsencesLKP Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (AbsenceID=" & aEmployeeComponent(N_CONCEPT_ID_EMPLOYEE) & ") And (OcurredDate=" & aEmployeeComponent(L_CONCEPT_START_DATE_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El concepto especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aEmployeeComponent(L_CONCEPT_END_DATE_EMPLOYEE) = CLng(oRecordset.Fields("EndDate").Value)
				aEmployeeComponent(D_CONCEPT_AMOUNT_EMPLOYEE) = CDbl(oRecordset.Fields("AbsenceHours").Value)
				aEmployeeComponent(N_CONCEPT_ACTIVE_EMPLOYEE) = CInt(oRecordset.Fields("Active").Value)
				aEmployeeComponent(S_CONCEPT_COMMENTS_EMPLOYEE) = CStr(oRecordset.Fields("Reasons").Value)
				aEmployeeComponent(N_EMPLOYEE_PAYROLL_DATE_EMPLOYEE) = CLng(oRecordset.Fields("AppliedDate").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetEmployeeSundays = lErrorNumber
	Err.Clear
End Function

Function GetEmployees(oRequest, oADODBConnection, aEmployeeComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the employees from
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployees"
	Dim sSort
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = Trim(aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE))
	If Len(aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE)) > 0 Then
		If InStr(1, aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE), "And ", vbBinaryCompare) <> 1 Then aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = "And " & aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE)
	End If
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) & " And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Areas.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	If aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) <> -1 Then aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) & " And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) & "," & S_WILD_CHAR & "')"
	sSort = aEmployeeComponent(S_SORT_COLUMN_EMPLOYEE)
	If aEmployeeComponent(B_SORT_DESCENDING_EMPLOYEE) Then sSort = Replace(sSort, ", ", " Desc, ") & " Desc"
	sErrorDescription = "No se pudo obtener la información de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Employees.*, CompanyName, JobNumber, ZoneName, AreaName, PositionShortName, PositionName, LevelName, StatusName From Employees, Companies, Jobs, Zones, Areas, Positions, Levels, StatusEmployees Where (Employees.CompanyID=Companies.CompanyID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Jobs.PositionID=Positions.PositionID) And (Employees.LevelID=Levels.LevelID) And (Employees.StatusID=StatusEmployees.StatusID) " & aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) & " Order By " & sSort, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)

	GetEmployees = lErrorNumber
	Err.Clear
End Function

Function GetEmployeesBankAccounts(oRequest, oADODBConnection, aEmployeeComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the employees from
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeesBankAccounts"
	Dim sSort
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim iActive

	iActive = aEmployeeComponent(N_ACTIVE_EMPLOYEE)
	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If
	aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = ""
	If Len(aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE)) > 0 Then
		If InStr(1, aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE), "And ", vbBinaryCompare) <> 1 Then aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = "And " & aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE)
	End If
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) & " And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Areas.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	If aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) <> -1 Then aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) = aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) & " And (Zones.ZonePath Like '" & S_WILD_CHAR & "," & aLoginComponent(N_PERMISSION_ZONE_ID_LOGIN) & "," & S_WILD_CHAR & "')"
	sErrorDescription = "No se pudo obtener la información de los empleados."
	sQuery = "Select Employees.EmployeeID, EmployeeNumber, Employees.PaymentCenterID, EmployeeName + ' ' + EmployeeLastName  + ' ' + EmployeeLastName2 As EmployeeFullName, BankName," & _
			 " BankAccounts.BankID, AccountID, AccountNumber, BankAccounts.StartDate, BankAccounts.EndDate, Users.UserLastName + ' ' + Users.UserName As UserFullName, BankAccounts.Active," & _
			 " PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, Zones.ZonePath" & _
			 " From Employees, BankAccounts, Banks, Users, Areas, Areas As PaymentCenters, Jobs, Zones As AreasZones, Zones As ParentZones, Zones, Companies" & _
			 " Where (Employees.JobID=Jobs.JobID)" & _
			 " And (Employees.PaymentCenterID=PaymentCenters.AreaID)" & _
			 " And (Jobs.AreaID=Areas.AreaID)" & _
			 " And (Areas.ZoneID=AreasZones.ZoneID)" & _
			 " And (AreasZones.ParentID=ParentZones.ZoneID)" & _
			 " And (PaymentCenters.ZoneID=Zones.ZoneID)" & _
			 " And (Employees.CompanyID=Companies.CompanyID)" & _
			 " And (Employees.PaymentCenterID=PaymentCenters.AreaID)" & _
			 " And (Employees.EmployeeID=BankAccounts.EmployeeID)" & _
			 " And (Banks.BankID=BankAccounts.BankID)" & _
			 " And (Users.UserID=BankAccounts.UserID)" & _
			 " And (Employees.CompanyID=Companies.CompanyID)"
	If CInt(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
		sQuery = sQuery & " And (BankAccounts.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
	Else
		If iActive Then
			sQuery = sQuery & " And (BankAccounts.EmployeeID=0)"
		End If
	End If
	sQuery = sQuery & aEmployeeComponent(S_QUERY_CONDITION_EMPLOYEE) & " And (BankAccounts.Active=" & iActive & ") Order By Employees.EmployeeID, BankAccounts.StartDate"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine
	If iActive = 0 Then Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & sQuery & """ />"

	GetEmployeesBankAccounts = lErrorNumber
	Err.Clear
End Function

Function GetLastEmployeeBankAccount(oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a Sundays for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetLastEmployeeBankAccount"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aEmployeeComponent(B_COMPONENT_INITIALIZED_EMPLOYEE)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeEmployeeComponent(oRequest, aEmployeeComponent)
	End If

	If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado para obtener sus cuentas bancarias."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información de la cuenta más reciente del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From BankAccounts Where (EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ") and Active = 1 Order By StartDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El empleado especificado no tiene cuentas bancarias registradas en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aEmployeeComponent(N_ACCOUNT_ID_EMPLOYEE) = CLng(oRecordset.Fields("AccountID").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetLastEmployeeBankAccount = lErrorNumber
	Err.Clear
End Function

Function GetSectionsToShow(oADODBConnection, aEmployeeComponent, lReasonID, sErrorDescription)
'************************************************************
'Purpose: To get the sections to show in a employee form
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetSectionsToShow"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información del empleado con el estatus indicado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sections From Reasons Where (ReasonID=" & lReasonID & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE) = CStr(oRecordset.Fields("Sections").Value)
		Else
			aEmployeeComponent(S_SECTIONS_TO_SHOW_EMPLOYEE) = ",1,"
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GetSectionsToShow = lErrorNumber
	Err.Clear
End Function

Function GetPayrollsEnableToApplyMovements(sPayrollsIDs, iSection, sErrorDescription)
'************************************************************
'Purpose: To get the payrolls who are enabled to
'         apply employees movements
'Inputs:
'Outputs: sPayrollsIDs, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPayrollsEnableToApplyMovements"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información de las nóminas habilitadas para el registro de movimientos."
	If iSection = 1 Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PayrollID From Payrolls Where (IsClosed<>1) And (IsActive_1=1)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PayrollID From Payrolls Where (IsClosed<>1) And (IsActive_1=1)", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	End If
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
	GetPayrollsEnableToApplyMovements = lErrorNumber
	Err.Clear
End Function

Function GetZoneByArea(oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To get ZoneID by AREAID
'Inputs:  oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetZoneByArea"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la clave de la zona solicitada."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CompanyID, ZoneID From Areas Where (AreaID = " & aEmployeeComponent(N_AREA_ID_EMPLOYEE) & ")", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			aEmployeeComponent(N_ZONE_ID_EMPLOYEE) = CLng(oRecordset.Fields("ZoneID").Value)
		End If
	End If

	Set oRecordset = Nothing
	GetZoneByArea = lErrorNumber
	Err.Clear
End Function
%>