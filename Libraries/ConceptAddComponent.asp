<%
Function AddConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new concept value into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddConcept"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If Not CheckExistencyOfConcept(aConceptComponent, sErrorDescription) Then
		lErrorNumber = L_ERR_DUPLICATED_RECORD
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aConceptComponent(N_ID_CONCEPT) = -1 Then
			sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
			lErrorNumber = GetNewIDFromTable(oADODBConnection, "Concepts", "ConceptID", "", 1, aConceptComponent(N_ID_CONCEPT), sErrorDescription)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo guardar la información del nuevo registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Concepts (ConceptID, StartDate, EndDate, ConceptShortName, ConceptName, BudgetID, PayrollTypeID, PeriodID, IsDeduction, ForAlimony, OnLeave, OrderInList, TaxAmount, TaxCurrencyID, TaxQttyID, TaxMin, TaxMinQttyID, TaxMax, TaxMaxQttyID, ExemptAmount, ExemptCurrencyID, ExemptQttyID, ExemptMin, ExemptMinQttyID, ExemptMax, ExemptMaxQttyID, StartUserID, EndUserID, StatusID) Values (" & aConceptComponent(N_ID_CONCEPT) & ", " & aConceptComponent(N_START_DATE_CONCEPT) & ", " & aConceptComponent(N_END_DATE_CONCEPT) & ", '" & Replace(aConceptComponent(S_SHORT_NAME_CONCEPT), "'", "") & "', '" & Replace(aConceptComponent(S_NAME_CONCEPT), "'", "´") & "', " & aConceptComponent(N_BUDGET_ID_CONCEPT) & ", " & aConceptComponent(N_PAYROLL_TYPE_ID_CONCEPT) & ", " & aConceptComponent(N_PERIOD_ID_CONCEPT) & ", " & aConceptComponent(N_IS_DEDUCTION_CONCEPT) & ", " & aConceptComponent(N_FOR_ALIMONY_CONCEPT) & ", " & aConceptComponent(N_ON_LEAVE_CONCEPT) & ", " & aConceptComponent(N_ORDER_IN_LIST_CONCEPT) & ", " & aConceptComponent(D_TAX_AMOUNT_CONCEPT) & ", " & aConceptComponent(N_TAX_CURRENCY_ID_CONCEPT) & ", " & aConceptComponent(N_TAX_QTTY_ID_CONCEPT) & ", " & aConceptComponent(D_TAX_MIN_CONCEPT) & ", " & aConceptComponent(N_TAX_MIN_QTTY_ID_CONCEPT) & ", " & aConceptComponent(D_TAX_MAX_CONCEPT) & ", " & aConceptComponent(N_TAX_MAX_QTTY_ID_CONCEPT) & ", " & aConceptComponent(D_EXEMPT_AMOUNT_CONCEPT) & ", " & aConceptComponent(N_EXEMPT_CURRENCY_ID_CONCEPT) & ", " & aConceptComponent(N_EXEMPT_QTTY_ID_CONCEPT) & ", " & aConceptComponent(D_EXEMPT_MIN_CONCEPT) & ", " & aConceptComponent(N_EXEMPT_MIN_QTTY_ID_CONCEPT) & ", " & aConceptComponent(D_EXEMPT_MAX_CONCEPT) & ", " & aConceptComponent(N_EXEMPT_MAX_QTTY_ID_CONCEPT) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", 0, " & aConceptComponent(N_STATUS_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			If aConceptComponent(N_IS_CREDIT) > 0 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From CreditTypes Where (CreditTypeID=" & aConceptComponent(N_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into CreditTypes (CreditTypeID, CreditTypeShortName, CreditTypeName, IsOther, Active) Values (" & aConceptComponent(N_ID_CONCEPT) & ", '" & Replace(aConceptComponent(S_SHORT_NAME_CONCEPT), "'", "") & "', '" & Replace(aConceptComponent(S_NAME_CONCEPT), "'", "´") & "', " & aConceptComponent(N_IS_OTHER) & ", 1)", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
	End If

	AddConcept = lErrorNumber
	Err.Clear
End Function

Function AddConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new concept value into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddConceptValue"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If Not CheckExistencyOfConceptValue(aConceptComponent, sErrorDescription) Then
		lErrorNumber = L_ERR_DUPLICATED_RECORD
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
			sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
			lErrorNumber = GetNewIDFromTable(oADODBConnection, "ConceptsValues", "RecordID", "", 1, aConceptComponent(N_RECORD_ID_CONCEPT), sErrorDescription)
		End If
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo guardar la información del nuevo registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into ConceptsValues (RecordID, ConceptID, CompanyID, EmployeeTypeID, PositionTypeID, EmployeeStatusID, JobStatusID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, WorkingHours, AdditionalShift, LevelID, EconomicZoneID, ServiceID, AntiquityID, Antiquity2ID, Antiquity3ID, Antiquity4ID, ForRisk, GenderID, HasChildren, SchoolarshipID, HasSyndicate, StartDate, EndDate, RegistrationStartDate, AuthorizationDate, RegistrationEndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, AppliesToID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, PositionID, StartUserID, EndUserID, StatusID) Values (" & aConceptComponent(N_RECORD_ID_CONCEPT) & ", " & aConceptComponent(N_ID_CONCEPT) & ", " & aConceptComponent(N_COMPANY_ID_CONCEPT) & ", " & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ", " & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ", " & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ", " & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ", " & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ", " & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ", " & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ", " & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ", " & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ", " & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ", " & aConceptComponent(N_LEVEL_ID_CONCEPT) & ", " & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ", " & aConceptComponent(N_SERVICE_ID_CONCEPT) & ", " & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ", " & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ", " & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ", " & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ", " & aConceptComponent(N_FOR_RISK_CONCEPT) & ", " & aConceptComponent(N_GENDER_ID_CONCEPT) & ", " & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ", " & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ", " & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ", " & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ", " & aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", 0, 0, " & aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT) & ", " & aConceptComponent(N_CURRENCY_ID_CONCEPT) & ", " & aConceptComponent(N_CONCEPT_QTTY_ID_CONCEPT) & ", " & aConceptComponent(N_CONCEPT_TYPE_ID_CONCEPT) & ", '" & Replace(aConceptComponent(S_APPLIES_ID_CONCEPT), "'", "") & "', " & aConceptComponent(D_CONCEPT_MIN_CONCEPT) & ", " & aConceptComponent(N_CONCEPT_MIN_QTTY_ID_CONCEPT) & ", " & aConceptComponent(D_CONCEPT_MAX_CONCEPT) & ", " & aConceptComponent(N_CONCEPT_MAX_QTTY_ID_CONCEPT) & ", " & aConceptComponent(N_POSITION_ID_CONCEPT) & ", " & aConceptComponent(N_START_USER_ID_CONCEPT) & ", " & aConceptComponent(N_END_USER_ID_CONCEPT) & ", " & aConceptComponent(N_STATUS_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	AddConceptValue = lErrorNumber
	Err.Clear
End Function

Function AddConceptsValuesFile(oRequest, oADODBConnection, sQuery, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To add concepts values for employee type into the database
'Inputs:  oRequest, oADODBConnection, sQuery
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddConceptsValuesFile"
	Dim oRecordset
	Dim lErrorNumber
	Dim asRecordID
	Dim sRecordIDChk

	sErrorDescription = "No se pudo obtener la información de la aplicación de tabuladores de forma masiva."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				aConceptComponent(N_RECORD_ID_CONCEPT) = CLng(oRecordset.Fields("RecordID").Value)
				lErrorNumber = SetActiveForConceptsValues(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
		End If
	End If

	Set oRecordset = Nothing
	AddConceptsValuesFile = lErrorNumber
	Err.Clear
End Function

Function AddPositionsSpecialJourneysFile(oRequest, oADODBConnection, sQuery, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To add concepts values for employee type into the database
'Inputs:  oRequest, oADODBConnection, sQuery
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddPositionsSpecialJourneysFile"
	Dim oRecordset
	Dim lErrorNumber
	Dim asRecordID
	Dim sRecordIDChk

	sErrorDescription = "No se pudo obtener la información para la aplicación de puestos para guardias y suplencias que están en proceso."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				aConceptComponent(N_RECORD_ID_CONCEPT) = CLng(oRecordset.Fields("RecordID").Value)
				lErrorNumber = SetActiveForPositionsSpecialJourneys(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
		End If
	End If

	Set oRecordset = Nothing
	AddPositionsSpecialJourneysFile = lErrorNumber
	Err.Clear
End Function

Function AddPositionsSpecialJourneysLKP(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new concept value into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddPositionsSpecialJourneysLKP"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If aConceptComponent(N_POSITION_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del puesto para guardar su valor."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Not CheckExistencyOfPositionsSpecialJourneysLKP(aConceptComponent, sErrorDescription) Then
			lErrorNumber = L_ERR_DUPLICATED_RECORD
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
		Else
			If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
				sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
				lErrorNumber = GetNewIDFromTable(oADODBConnection, "PositionsSpecialJourneysLKP", "RecordID", "", 1, aConceptComponent(N_RECORD_ID_CONCEPT), sErrorDescription)
			End If
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo guardar la información del nuevo registro."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PositionsSpecialJourneysLKP (RecordID, StartDate, EndDate, PositionID, LevelID, WorkingHours, ServiceID, CenterTypeID, IsActive1, IsActive2, IsActive3, IsActive4, Active) Values (" & aConceptComponent(N_RECORD_ID_CONCEPT) & ", " & aConceptComponent(N_START_DATE_CONCEPT) & ", " & aConceptComponent(N_END_DATE_CONCEPT) & ", " & aConceptComponent(N_POSITION_ID_CONCEPT) & ", " & aConceptComponent(N_LEVEL_ID_CONCEPT) & ", " & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ", " & aConceptComponent(N_SERVICE_ID_CONCEPT) & ", " & aConceptComponent(N_CENTER_TYPE_ID) & ", " & aConceptComponent(N_IS_ACTIVE1) & ", " & aConceptComponent(N_IS_ACTIVE2) & ", " & aConceptComponent(N_IS_ACTIVE3) & ", " & aConceptComponent(N_IS_ACTIVE4) & ", " & aConceptComponent(N_ACTIVE) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
	End If

	AddPositionsSpecialJourneysLKP = lErrorNumber
	Err.Clear
End Function
%>