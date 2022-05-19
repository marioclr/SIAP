<%
Const N_ID_JOB = 0
Const N_ID_EMPLOYEE_JOB = 1
Const N_ID_OWNER_JOB = 2
Const S_NUMBER_JOB = 3
Const N_COMPANY_ID_JOB = 4
Const N_ZONE_ID_JOB = 5
Const N_AREA_ID_JOB = 6
Const N_PAYMENT_CENTER_ID_JOB = 7
Const N_POSITION_ID_JOB = 8
Const N_JOB_TYPE_ID_JOB = 9
Const N_SHIFT_ID_JOB = 10
Const N_JOURNEY_ID_JOB = 11
Const N_CLASSIFICATION_ID_JOB = 12
Const N_GROUP_GRADE_LEVEL_ID_JOB = 13
Const N_INTEGRATION_ID_JOB = 14
Const N_OCCUPATION_TYPE_ID_JOB = 15
Const N_SERVICE_ID_JOB = 16
Const N_LEVEL_ID_JOB = 17
Const D_WORKING_HOURS_JOB = 18
Const N_START_DATE_JOB = 19
Const N_END_DATE_JOB = 20
Const N_STATUS_ID_JOB = 21
Const N_ACTIVE_JOB = 22
Const S_BUDGETS_ID_JOB = 23
Const N_JOB_DATE_JOB = 24
Const N_END_DATE_HISTORY_JOB = 25
Const N_EMPLOYEE_TYPE_ID_JOB = 26
Const N_POSITION_TYPE_ID_JOB = 27
Const N_ECONOMIC_ZONE_ID_JOB = 28

Const N_SHOW_BY_JOB = 29
Const B_SEND_TO_IFRAME_JOB = 30
Const S_QUERY_CONDITION_JOB = 31
Const B_CHECK_FOR_DUPLICATED_JOB = 32
Const B_IS_DUPLICATED_JOB = 33
Const B_COMPONENT_INITIALIZED_JOB = 34

Const N_JOB_COMPONENT_SIZE = 34

Const N_SHOW_BY_AREA = 1
Const N_SHOW_BY_POSITION = 2

Dim aJobComponent()
Redim aJobComponent(N_JOB_COMPONENT_SIZE)

Function InitializeJobComponent(oRequest, aJobComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Job
'         Component using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aJobComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeJobComponent"
	Redim Preserve aJobComponent(N_JOB_COMPONENT_SIZE)
	Dim oItem

	If IsEmpty(aJobComponent(N_ID_JOB)) Then
		If Len(oRequest("JobID").Item) > 0 Then
			aJobComponent(N_ID_JOB) = CLng(oRequest("JobID").Item)
		Else
			aJobComponent(N_ID_JOB) = -1
		End If
	End If
	
	If IsEmpty(aJobComponent(N_ID_EMPLOYEE_JOB)) Then
		If Len(oRequest("EmployeeID").Item) > 0 Then
			aJobComponent(N_ID_EMPLOYEE_JOB) = CLng(oRequest("EmployeeID").Item)
		Else
			aJobComponent(N_ID_EMPLOYEE_JOB) = -1
		End If
	End If

	If IsEmpty(aJobComponent(N_ID_OWNER_JOB)) Then
		If Len(oRequest("OwnerID").Item) > 0 Then
			aJobComponent(N_ID_OWNER_JOB) = CLng(oRequest("OwnerID").Item)
		Else
			aJobComponent(N_ID_OWNER_JOB) = -1
		End If
	End If

	If IsEmpty(aJobComponent(S_NUMBER_JOB)) Then
		If Len(oRequest("JobNumber").Item) > 0 Then
			aJobComponent(S_NUMBER_JOB) = oRequest("JobNumber").Item
		Else
			aJobComponent(S_NUMBER_JOB) = ""
		End If
	End If
	aJobComponent(S_NUMBER_JOB) = Left(aJobComponent(S_NUMBER_JOB), 50)

	If IsEmpty(aJobComponent(N_COMPANY_ID_JOB)) Then
		If Len(oRequest("CompanyID").Item) > 0 Then
			aJobComponent(N_COMPANY_ID_JOB) = CLng(oRequest("CompanyID").Item)
		Else
			aJobComponent(N_COMPANY_ID_JOB) = -1
		End If
	End If

    If aJobComponent(N_ID_JOB) = -1 Then aJobComponent(N_ID_JOB) = CLng(aJobComponent(S_NUMBER_JOB))

	If IsEmpty(aJobComponent(N_ZONE_ID_JOB)) Then
		If Len(oRequest("ZoneID").Item) > 0 Then
			aJobComponent(N_ZONE_ID_JOB) = CLng(oRequest("ZoneID").Item)
		Else
			aJobComponent(N_ZONE_ID_JOB) = -1
		End If
	End If

	If IsEmpty(aJobComponent(N_AREA_ID_JOB)) Then
		If Len(oRequest("AreaID").Item) > 0 Then
			aJobComponent(N_AREA_ID_JOB) = CLng(oRequest("AreaID").Item)
		Else
			aJobComponent(N_AREA_ID_JOB) = -1
		End If
	End If
	
	If IsEmpty(aJobComponent(N_PAYMENT_CENTER_ID_JOB)) Then
		If Len(oRequest("PaymentCenterID").Item) > 0 Then
			aJobComponent(N_PAYMENT_CENTER_ID_JOB) = CLng(oRequest("PaymentCenterID").Item)
		Else
			aJobComponent(N_PAYMENT_CENTER_ID_JOB) = -1
		End If
	End If

	If IsEmpty(aJobComponent(N_POSITION_ID_JOB)) Then
		If Len(oRequest("PositionID").Item) > 0 Then
			aJobComponent(N_POSITION_ID_JOB) = CLng(oRequest("PositionID").Item)
		Else
			aJobComponent(N_POSITION_ID_JOB) = -1
		End If
	End If

	If IsEmpty(aJobComponent(N_JOB_TYPE_ID_JOB)) Then
		If Len(oRequest("JobTypeID").Item) > 0 Then
			aJobComponent(N_JOB_TYPE_ID_JOB) = CLng(oRequest("JobTypeID").Item)
		Else
			aJobComponent(N_JOB_TYPE_ID_JOB) = 1
		End If
	End If

	If IsEmpty(aJobComponent(N_SHIFT_ID_JOB)) Then
		If Len(oRequest("ShiftID").Item) > 0 Then
			aJobComponent(N_SHIFT_ID_JOB) = CLng(oRequest("ShiftID").Item)
		Else
			aJobComponent(N_SHIFT_ID_JOB) = -1
		End If
	End If

	If IsEmpty(aJobComponent(N_JOURNEY_ID_JOB)) Then
		If Len(oRequest("JourneyID").Item) > 0 Then
			aJobComponent(N_JOURNEY_ID_JOB) = CLng(oRequest("JourneyID").Item)
		Else
			aJobComponent(N_JOURNEY_ID_JOB) = -1
		End If
	End If

	If IsEmpty(aJobComponent(N_CLASSIFICATION_ID_JOB)) Then
		If Len(oRequest("ClassificationID").Item) > 0 Then
			aJobComponent(N_CLASSIFICATION_ID_JOB) = CLng(oRequest("ClassificationID").Item)
		Else
			aJobComponent(N_CLASSIFICATION_ID_JOB) = -1
		End If
	End If

	If IsEmpty(aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB)) Then
		If Len(oRequest("GroupGradeLevelID").Item) > 0 Then
			aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) = CLng(oRequest("GroupGradeLevelID").Item)
		Else
			aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) = -1
		End If
	End If

	If IsEmpty(aJobComponent(N_INTEGRATION_ID_JOB)) Then
		If Len(oRequest("IntegrationID").Item) > 0 Then
			aJobComponent(N_INTEGRATION_ID_JOB) = CLng(oRequest("IntegrationID").Item)
		Else
			aJobComponent(N_INTEGRATION_ID_JOB) = -1
		End If
	End If

	If IsEmpty(aJobComponent(N_OCCUPATION_TYPE_ID_JOB)) Then
		If Len(oRequest("OccupationTypeID").Item) > 0 Then
			aJobComponent(N_OCCUPATION_TYPE_ID_JOB) = CLng(oRequest("OccupationTypeID").Item)
		Else
			aJobComponent(N_OCCUPATION_TYPE_ID_JOB) = -1
		End If
	End If

	If IsEmpty(aJobComponent(N_SERVICE_ID_JOB)) Then
		If Len(oRequest("ServiceID").Item) > 0 Then
			aJobComponent(N_SERVICE_ID_JOB) = CLng(oRequest("ServiceID").Item)
		Else
			aJobComponent(N_SERVICE_ID_JOB) = -1
		End If
	End If

	If IsEmpty(aJobComponent(N_LEVEL_ID_JOB)) Then
		If Len(oRequest("LevelID").Item) > 0 Then
			aJobComponent(N_LEVEL_ID_JOB) = CLng(oRequest("LevelID").Item)
		Else
			aJobComponent(N_LEVEL_ID_JOB) = -1
		End If
	End If

	If IsEmpty(aJobComponent(D_WORKING_HOURS_JOB)) Then
		If Len(oRequest("WorkingHours").Item) > 0 Then
			aJobComponent(D_WORKING_HOURS_JOB) = CLng(oRequest("WorkingHours").Item)
		Else
			aJobComponent(D_WORKING_HOURS_JOB) = -1
		End If
	End If

	If IsEmpty(aJobComponent(N_START_DATE_JOB)) Then
		If Len(oRequest("StartYear").Item) > 0 Then
			aJobComponent(N_START_DATE_JOB) = CLng(oRequest("StartYear").Item & Right(("0" & oRequest("StartMonth").Item), Len("00")) & Right(("0" & oRequest("StartDay").Item), Len("00")))
		ElseIf Len(oRequest("StartDate").Item) > 0 Then
			aJobComponent(N_START_DATE_JOB) = CLng(oRequest("StartDate").Item)
		Else
			aJobComponent(N_START_DATE_JOB) = 0
		End If
	End If

	If IsEmpty(aJobComponent(N_END_DATE_JOB)) Then
		If Len(oRequest("EndYear").Item) > 0 Then
			aJobComponent(N_END_DATE_JOB) = CLng(oRequest("EndYear").Item & Right(("0" & oRequest("EndMonth").Item), Len("00")) & Right(("0" & oRequest("EndDay").Item), Len("00")))
		ElseIf Len(oRequest("EndDate").Item) > 0 Then
			aJobComponent(N_END_DATE_JOB) = CLng(oRequest("EndDate").Item)
		Else
			aJobComponent(N_END_DATE_JOB) = 30000000
		End If
	End If

	If IsEmpty(aJobComponent(N_JOB_DATE_JOB)) Then
		If Len(oRequest("JobYear").Item) > 0 Then
			aJobComponent(N_JOB_DATE_JOB) = CLng(oRequest("JobYear").Item & Right(("0" & oRequest("JobMonth").Item), Len("00")) & Right(("0" & oRequest("JobDay").Item), Len("00")))
		ElseIf Len(oRequest("JobDate").Item) > 0 Then
			aJobComponent(N_JOB_DATE_JOB) = CLng(oRequest("JobDate").Item)
		Else
			aJobComponent(N_JOB_DATE_JOB) = aJobComponent(N_START_DATE_JOB)
		End If
	End If

	If IsEmpty(aJobComponent(N_END_DATE_HISTORY_JOB)) Then
		If Len(oRequest("JobEndYear").Item) > 0 Then
			aJobComponent(N_END_DATE_HISTORY_JOB) = CLng(oRequest("JobEndYear").Item & Right(("0" & oRequest("JobEndMonth").Item), Len("00")) & Right(("0" & oRequest("JobEndDay").Item), Len("00")))
		ElseIf Len(oRequest("JobEndDate").Item) > 0 Then
			aJobComponent(N_END_DATE_HISTORY_JOB) = CLng(oRequest("JobEndDate").Item)
		Else
			aJobComponent(N_END_DATE_HISTORY_JOB) = aJobComponent(N_END_DATE_JOB)
		End If
	End If

	If IsEmpty(aJobComponent(N_STATUS_ID_JOB)) Then
		If Len(oRequest("StatusID").Item) > 0 Then
			aJobComponent(N_STATUS_ID_JOB) = CLng(oRequest("StatusID").Item)
		Else
			aJobComponent(N_STATUS_ID_JOB) = 0
		End If
	End If

	If IsEmpty(aJobComponent(N_ACTIVE_JOB)) Then
		If Len(oRequest("Active").Item) > 0 Then
			aJobComponent(N_ACTIVE_JOB) = CInt(oRequest("Active").Item)
		Else
			aJobComponent(N_ACTIVE_JOB) = 1
		End If
	End If

	If IsEmpty(aJobComponent(S_BUDGETS_ID_JOB)) Then
		If Len(oRequest("BudgetID").Item) > 0 Then
			aJobComponent(S_BUDGETS_ID_JOB) = ""
			For Each oItem In oRequest("BudgetID")
				aJobComponent(S_BUDGETS_ID_JOB) = aJobComponent(S_BUDGETS_ID_JOB) & oItem & ","
			Next
			If Len(aJobComponent(S_BUDGETS_ID_JOB)) > 0 Then aJobComponent(S_BUDGETS_ID_JOB) = Left(aJobComponent(S_BUDGETS_ID_JOB), (Len(aJobComponent(S_BUDGETS_ID_JOB)) - Len(",")))
		ElseIf Len(oRequest("BudgetsID").Item) > 0 Then
			aJobComponent(S_BUDGETS_ID_JOB) = oRequest("BudgetsID").Item
		Else
			aJobComponent(S_BUDGETS_ID_JOB) = ""
		End If
	End If

	aJobComponent(N_SHOW_BY_JOB) = 0
	aJobComponent(B_SEND_TO_IFRAME_JOB) = False
	aJobComponent(B_CHECK_FOR_DUPLICATED_JOB) = True
	aJobComponent(B_IS_DUPLICATED_JOB) = False

	aJobComponent(B_COMPONENT_INITIALIZED_JOB) = True
	InitializeJobComponent = Err.number
	Err.Clear
End Function

Function AddJob(oRequest, oADODBConnection, aJobComponent, bAddConsecutive, sErrorDescription)
'************************************************************
'Purpose: To add a new job into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddJob"
	Dim alBudgetsID
	Dim oRecordset
	Dim iIndex
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aJobComponent(B_COMPONENT_INITIALIZED_JOB)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeJobComponent(oRequest, aJobComponent)
	End If

	If aJobComponent(N_ID_JOB) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para la nueva plaza."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "Jobs", "JobID", "", 1, aJobComponent(N_ID_JOB), sErrorDescription)
		aJobComponent(S_NUMBER_JOB) = Right(("000000" & aJobComponent(N_ID_JOB)), Len("000000"))
	End If

	If lErrorNumber = 0 Then
		If aJobComponent(B_CHECK_FOR_DUPLICATED_JOB) Then
			lErrorNumber = CheckExistencyOfJob(aJobComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aJobComponent(B_IS_DUPLICATED_JOB) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un plaza con el número " & aJobComponent(S_NUMBER_JOB) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "JobComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				If (aJobComponent(N_POSITION_ID_JOB) > 0) And (aJobComponent(N_POSITION_ID_JOB) <> L_HONORARY_POSITION_ID) Then
					lErrorNumber = GetPosition(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
				End If
				If lErrorNumber = 0 Then
					If Not CheckJobInformationConsistency(aJobComponent, sErrorDescription) Then
						lErrorNumber = -1
					Else
						sErrorDescription = "No se pudo guardar la información de la nueva plaza."
						If Len(oRequest("JobStartDateYear").Item) > 0 Then
							aJobComponent(N_START_DATE_JOB) = oRequest("JobStartDateYear").Item & oRequest("JobStartDateMonth").Item & oRequest("JobStartDateDay").Item
							aJobComponent(N_END_DATE_JOB) = oRequest("JobEndDateYear").Item & oRequest("JobEndDateMonth").Item & oRequest("JobEndDateDay").Item
						End If
						aJobComponent(S_NUMBER_JOB) = Right(("000000" & aJobComponent(N_ID_JOB)), Len("000000"))
						If aJobComponent(N_END_DATE_JOB) = 0 Then aJobComponent(N_END_DATE_JOB) = 30000000
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Jobs (JobID, OwnerID, JobNumber, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StartDate, EndDate, StatusID, Active, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", '" & Replace(aJobComponent(S_NUMBER_JOB), "'", "") & "', " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_START_DATE_JOB) & ", " & aJobComponent(N_END_DATE_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aJobComponent(N_ACTIVE_JOB) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						If bAddConsecutive = True  Then
							If aJobComponent(N_ID_JOB) < L_HONORARY_JOB_START_ID Then
								Call UpdateConsecutiveID(oADODBConnection, 100, aJobComponent(N_ID_JOB), "")
							End If
						End If
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudo guardar la información de la nueva plaza."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From JobsBudgetsLKP Where (JobID=" & aJobComponent(N_ID_JOB) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
						If lErrorNumber = 0 Then
							alBudgetsID = Split(aJobComponent(S_BUDGETS_ID_JOB), ",", -1, vbBinaryCompare)
							For iIndex = 0 To UBound(alBudgetsID)
								sErrorDescription = "No se pudo guardar la información de la nueva plaza."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsBudgetsLKP (JobID, BudgetID, StartDate, EndDate, StartUserID, EndUserID) Values (" & aJobComponent(N_ID_JOB) & ", " & alBudgetsID(iIndex) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", 0, " & aLoginComponent(N_USER_ID_LOGIN) & ", -1)", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
							Next
						End If
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudo guardar la información de la nueva plaza."
							lErrorNumber = ModifyJobHistoryList(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
							'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_START_DATE_JOB) & ", " & aJobComponent(N_END_DATE_JOB) & ", " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
					End If
				End If
			End If
		End If
	End If

	AddJob = lErrorNumber
	Err.Clear
End Function

Function AddJobFile(oRequest, oADODBConnection, sQuery, lReasonID, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new job into the database
'Inputs:  oRequest, oADODBConnection, sQuery, lReasonID
'Outputs: aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddJobFile"
	Dim oRecordset
	Dim lErrorNumber

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				If Not IsEmpty(oRequest(CStr(oRecordset.Fields("JobID").Value))) Then
					aJobComponent(N_ID_JOB) = CLng(oRecordset.Fields("JobID").Value)
					lErrorNumber = SetActiveForJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
				End If
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
		End If
	End If

	Set oRecordset = Nothing
	AddJobFile = lErrorNumber
	Err.Clear
End Function

Function GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a job from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetJob"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aJobComponent(B_COMPONENT_INITIALIZED_JOB)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeJobComponent(oRequest, aJobComponent)
	End If

	If aJobComponent(N_ID_JOB) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador de la plaza para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "JobComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información de la plaza."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Jobs Where JobID=" & aJobComponent(N_ID_JOB), "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "La plaza especificada no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "JobComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
				oRecordset.Close
			Else
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
				oRecordset.Close

				If aJobComponent(N_POSITION_ID_JOB) <> -1 Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, JobDate, EndDate From JobsHistoryList Where JobID=" & aJobComponent(N_ID_JOB) & " Order by JobDate Desc", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							aJobComponent(N_ID_EMPLOYEE_JOB) = CStr(oRecordset.Fields("EmployeeID").Value)
							aJobComponent(N_JOB_DATE_JOB) = CLng(oRecordset.Fields("JobDate").Value)
							aJobComponent(N_END_DATE_HISTORY_JOB) = CLng(oRecordset.Fields("EndDate").Value)
							oRecordset.Close
						End If
					End If
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Positions Where PositionID=" & aJobComponent(N_POSITION_ID_JOB), "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							If aJobComponent(N_ID_JOB) < 600000 Then
								aJobComponent(N_COMPANY_ID_JOB) = CLng(oRecordset.Fields("CompanyID").Value)
							End If
							aJobComponent(N_EMPLOYEE_TYPE_ID_JOB) = CLng(oRecordset.Fields("EmployeeTypeID").Value)
							aJobComponent(N_POSITION_TYPE_ID_JOB) = CLng(oRecordset.Fields("PositionTypeID").Value)
							oRecordset.Close
						End If
					End If
				End If

				If aJobComponent(N_AREA_ID_JOB) <> -1 Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EconomicZoneID From Areas Where AreaID=" & aJobComponent(N_AREA_ID_JOB), "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							aJobComponent(N_ECONOMIC_ZONE_ID_JOB) = CLng(oRecordset.Fields("EconomicZoneID").Value)
							oRecordset.Close
						End If
					End If
				End If

				aJobComponent(S_BUDGETS_ID_JOB) = ""
				sErrorDescription = "No se pudo obtener la información de la plaza."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select BudgetID From JobsBudgetsLKP Where (JobID=" & aJobComponent(N_ID_JOB) & ") And (EndDate=0) Order By BudgetID", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						aJobComponent(S_BUDGETS_ID_JOB) = aJobComponent(S_BUDGETS_ID_JOB) & CStr(oRecordset.Fields("BudgetID").Value) & ","
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					If Len(aJobComponent(S_BUDGETS_ID_JOB)) > 0 Then aJobComponent(S_BUDGETS_ID_JOB) = Left(aJobComponent(S_BUDGETS_ID_JOB), (Len(aJobComponent(S_BUDGETS_ID_JOB)) - Len(",")))
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetJob = lErrorNumber
	Err.Clear
End Function

Function GetJobs(oRequest, oADODBConnection, aJobComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the jobs from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aJobComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetJobs"
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aJobComponent(B_COMPONENT_INITIALIZED_JOB)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeJobComponent(oRequest, aJobComponent)
	End If

	If Len(aJobComponent(S_QUERY_CONDITION_JOB)) > 0 Then
		sCondition = Trim(aJobComponent(S_QUERY_CONDITION_JOB))
		If InStr(1, sCondition, "And ", vbTextCompare) <> 1 Then
			sCondition = "And " & sCondition
		End If
	End If
	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		sCondition = sCondition & " And (Areas.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
	End If
	sErrorDescription = "No se pudo obtener la información de las plazas."
	'sQuery = "Select Jobs.*, ZoneName, AreaShortName, AreaCode, AreaName, PositionShortName, PositionName, JobTypeName, OccupationTypeName, ServiceShortName, ServiceName, LevelName, StatusName, JobDate, JobsHistoryList.EndDate, EmployeesHistoryList.ReasonID, ReasonName From Jobs, Zones, Areas, Positions, JobTypes, OccupationTypes, Services, Levels, StatusJobs, JobsHistoryList, EmployeesHistoryList, Reasons Where (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Jobs.PositionID=Positions.PositionID) And (Jobs.JobTypeID=JobTypes.JobTypeID) And (Jobs.OccupationTypeID=OccupationTypes.OccupationTypeID) And (Jobs.ServiceID=Services.ServiceID) And (Jobs.LevelID=Levels.LevelID) And (Jobs.StatusID=StatusJobs.StatusID) And (jobs.JobID>-1) " & sCondition & " And Jobs.JobID = JobsHistoryList.JobID And Jobs.JobID = EmployeesHistoryList.JobID And JobsHistoryList.JobID = EmployeesHistoryList.JobID And JobsHistoryList.JobDate = (Select MAX(JobDate) As maxDate From JobsHistoryList Where jobID = Jobs.JobId) And EmployeesHistoryList.EmployeeDate = (Select MAX(JobDate) As maxDate From JobsHistoryList Where jobID = Jobs.JobId) And EmployeesHistoryList.ReasonID = Reasons.ReasonID Order By JobNumber"
	sQuery = "Select Jobs.*, ZoneName, AreaShortName, AreaCode, AreaName, PositionShortName, PositionName, JobTypeName, OccupationTypeName, ServiceShortName, ServiceName, LevelName, StatusName, JobDate, JobsHistoryList.EndDate From Jobs, Zones, Areas, Positions, JobTypes, OccupationTypes, Services, Levels, StatusJobs, JobsHistoryList Where (Jobs.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Jobs.PositionID=Positions.PositionID) And (Jobs.JobTypeID=JobTypes.JobTypeID) And (Jobs.OccupationTypeID=OccupationTypes.OccupationTypeID) And (Jobs.ServiceID=Services.ServiceID) And (Jobs.LevelID=Levels.LevelID) And (Jobs.StatusID=StatusJobs.StatusID) And (jobs.JobID>-1) " & sCondition & " And Jobs.JobID = JobsHistoryList.JobID And JobsHistoryList.JobDate = (Select MAX(JobDate) As maxDate From JobsHistoryList Where (jobID = Jobs.JobId) And (JobsHistoryList.EndDate <> 0)) Order By JobNumber"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: " & sQuery & " -->" & vbNewLine

	GetJobs = lErrorNumber
	Err.Clear
End Function

Function GetJobsNotAdded(oRequest, oADODBConnection, aJobComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the jobs from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aJobComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetJobsNotAdded"
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aJobComponent(B_COMPONENT_INITIALIZED_JOB)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeJobComponent(oRequest, aJobComponent)
	End If

	If Len(aJobComponent(S_QUERY_CONDITION_JOB)) > 0 Then
		sCondition = Trim(aJobComponent(S_QUERY_CONDITION_JOB))
		If InStr(1, sCondition, "And ", vbTextCompare) <> 1 Then
			sCondition = "And " & sCondition
		End If
	End If
	sErrorDescription = "No se pudo obtener la información de las plazas."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreasPositionsLKP.AreaID, AreaShortName, AreaName, AreasPositionsLKP.PositionID, PositionShortName, PositionName, AreasPositionsLKP.Jobs From AreasPositionsLKP, Areas, Positions Where (AreasPositionsLKP.AreaID=Areas.AreaID) And (AreasPositionsLKP.PositionID=Positions.PositionID) " & Replace(sCondition, "Jobs.", "AreasPositionsLKP.") & " Order By AreaName, PositionShortName, PositionName", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetJobsNotAdded = lErrorNumber
	Err.Clear
End Function

Function GetPosition(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a position from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPosition"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aJobComponent(B_COMPONENT_INITIALIZED_JOB)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeJobComponent(oRequest, aJobComponent)
	End If

	If aJobComponent(N_ID_JOB) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador de la plaza para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "JobComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del puesto."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Positions Where PositionID=" & aJobComponent(N_POSITION_ID_JOB), "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El puesto especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "JobComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
				oRecordset.Close
			Else
				aJobComponent(N_COMPANY_ID_JOB) = CLng(oRecordset.Fields("CompanyID").Value)
				aJobComponent(N_CLASSIFICATION_ID_JOB) = CLng(oRecordset.Fields("ClassificationID").Value)
				aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) = CLng(oRecordset.Fields("GroupGradeLevelID").Value)
				aJobComponent(N_INTEGRATION_ID_JOB) = CLng(oRecordset.Fields("IntegrationID").Value)
				aJobComponent(N_LEVEL_ID_JOB) = CLng(oRecordset.Fields("LevelID").Value)
				aJobComponent(D_WORKING_HOURS_JOB) = CDbl(oRecordset.Fields("WorkingHours").Value)
				aJobComponent(N_EMPLOYEE_TYPE_ID_JOB) = CLng(oRecordset.Fields("EmployeeTypeID").Value)
				If CLng(oRecordset.Fields("PositionTypeID").Value) <> 6 Then
					aJobComponent(N_POSITION_TYPE_ID_JOB) = CLng(oRecordset.Fields("PositionTypeID").Value)
				Else
					aJobComponent(N_POSITION_TYPE_ID_JOB) = 2
				End If
				oRecordset.Close
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetPosition = lErrorNumber
	Err.Clear
End Function

Function ModifyJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing job in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyJob"
	Dim alBudgetsID
	Dim sDate
	Dim iIndex
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim lHistoryJobDate
	Dim lHistoryEndDate
	Dim lStatusID
	Dim dStartDate
	Dim dEndDate
	Dim dCurrencyDate
	Dim sQuery
	Dim bIsCurrent
	Dim lReasonID
	
	bIsCurrent = True
	bComponentInitialized = aJobComponent(B_COMPONENT_INITIALIZED_JOB)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeJobComponent(oRequest, aJobComponent)
	End If
	lReasonID = -1
	If Len(oRequest("ReasonID").Item) <> 0 Then 
		lReasonID = oRequest("ReasonID").Item
		If InStr(1, ",1,2,3,4,5,6,8,10,62,63,66,78,79,80,81,", "," & lReasonID & ",", vbBinaryCompare) <> 0 Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobID, JobDate, EndDate, EmployeeID, OwnerID From JobsHistoryList Where (JobID=" & aJobComponent(N_ID_JOB) & ") Order By EndDate Desc", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					If Not IsEmpty(oRequest("EmployeeID").Item) Then
						If (CLng(oRecordset.Fields("EmployeeID").Value) <> CLng(oRequest("EmployeeID").Item)) Then
							bIsCurrent = False
						End If
					ElseIf Not IsEmpty(oRequest("AuthorizationFile").Item) Then
						If (CLng(aEmployeeComponent(N_ID_EMPLOYEE)) <> CLng(oRecordset.Fields("EmployeeID").Value)) Then
							bIsCurrent = False
						End If
					End If
				End If
			End If
		End If
	End If
	If lReasonID = 10 Then
		sQuery = "Select JobId, EmployeeID, JobDate, EndDate, StatusId From JobsHistoryList Where (JobID = " & aJobComponent(N_ID_JOB) & ") And (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") Order By JobDate Desc"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If Not oRecordset.EOF Then
			lHistoryEndDate = CLng(oRecordset.Fields("EndDate").Value)
			sErrorDescription = "No se pudo actualizar el historial de la plaza"
			sQuery = "Update JobsHistoryList Set EndDate = " & AddDaysToSerialDate(aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE),-1) & " Where (JobID = " & aJobComponent(N_ID_JOB) & ") And (EmployeeID = " & aEmployeeComponent(N_ID_EMPLOYEE) & ") And (JobDate = " & aJobComponent(N_START_DATE_JOB) & ")"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo actualizar el historial de la plaza"
				sQuery = "Select * From JobsHistoryList Where (JobID = " & aJobComponent(N_ID_JOB) & ") And (JobDate = " & AddDaysToSerialDate(lHistoryEndDate,1) & ") And (StatusID In  (2,5))"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If Not oRecordset.EOF Then
					sQuery = "Update JobsHistoryList Set JobDate = " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", EndDate = 30000000 Where (JobID = " & aJobComponent(N_ID_JOB) & ") And (JobDate = " & AddDaysToSerialDate(lHistoryEndDate,1) & ")"
				Else
					sQuery = "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", 30000000, " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")"
				End If
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		Else
			sQuery = "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & aEmployeeComponent(N_EMPLOYEE_DATE_EMPLOYEE) & ", 30000000, " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		bIsCurrent = False
	End If
	If bIsCurrent Then
		If aJobComponent(N_ID_JOB) = -1 Then
			lErrorNumber = -1
			sErrorDescription = "No se especificó el identificador de la plaza a modificar."
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "JobComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
		Else
			If aJobComponent(B_CHECK_FOR_DUPLICATED_JOB) Then
				lErrorNumber = CheckExistencyOfJob(aJobComponent, sErrorDescription)
			End If
			If lErrorNumber = 0 Then
				If aJobComponent(B_IS_DUPLICATED_JOB) Then
					lErrorNumber = L_ERR_DUPLICATED_RECORD
					sErrorDescription = "Ya existe un plaza con el número " & aJobComponent(S_NUMBER_JOB) & "."
					Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "JobComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
				Else
					If Not CheckJobInformationConsistency(aJobComponent, sErrorDescription) Then
						lErrorNumber = -1
					Else
						sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
						sErrorDescription = "No se pudo modificar la información de la plaza."
						aJobComponent(S_NUMBER_JOB) = Right("000000" & aJobComponent(N_ID_JOB), 6)
						If (StrComp(oRequest("Modify").Item , "Aplicar Titularidad", vbBinaryCompare) = 0) Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID From Employees Where (EmployeeID=" & oRequest("OwnerID").Item & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If oRecordset.EOF Then
									lErrorNumber = -1
									sErrorDescription = "El empleado indicado no está registrado"
								End If
							End If
'Condición bloqueada temporalmente
'							If lErrorNumber = 0 Then
'								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeTypeID From Employees Where (EmployeeID=" & oRequest("OwnerID").Item & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
'							End If
'							If lErrorNumber = 0 Then
'								If CInt(oRecordset.Fields("EmployeeTypeID").Value) = 1 Then
'									lErrorNumber = -1
'								End If
'							End If
							If lErrorNumber = 0 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobID, OwnerID From Jobs Where (JobID <> " & aJobComponent(N_ID_JOB) & ") And OwnerID = " & oRequest("OwnerID").Item, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If Not oRecordset.EOF Then
									lErrorNumber = -1
									sErrorDescription = "El empleado propuesto ya es titular en la plaza " & oRecordset.Fields("JobID").Value
								End If
							End If
							If lErrorNumber = 0 Then 
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select OwnerID From Jobs Where (JobID=" & aJobComponent(N_ID_JOB) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If (lErrorNumber = 0) Then
									If (Len(oRecordset.Fields("OwnerID").Value) > 0) And ((CLng(oRecordset.Fields("OwnerID").Value) <> 0) And (CLng(oRecordset.Fields("OwnerID").Value) <> -1)) Then
										lErrorNumber = -1
										sErrorDescription = "La plaza indicada ya tiene un titular asignado"
									Else
'										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID From Employees Where (JobID=" & aJobComponent(N_ID_JOB) & ") And (Active = 1)", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
'		Condición bloqueada temporalmente
'										If (StrComp(CStr(oRecordset.Fields("EmployeeID").Value),CStr(oRequest("OwnerID").Item),vbBinaryCompare) <> 0) Then
'											lErrorNumber = -1
'											sErrorDescription = "El empleado propuesto como titular no ocupa la plaza actualmente"
'										Else 
											'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobDate From JobsHistoryList Where (EmployeeID = " & aJobComponent(N_ID_OWNER_JOB) & ") and (JobID = " & aJobComponent(N_ID_JOB) & ") Order By EndDate Desc", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
'Modificación solicitada para permitir asignar titularidades atrasadas.
											dCurrencyDate = CLng(oRequest("OwnerJobYear").Item & oRequest("OwnerJobMonth").Item & oRequest("OwnerJobDay").Item)
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobDate From JobsHistoryList Where (JobID = " & aJobComponent(N_ID_JOB) & ") And (JobDate <" & dCurrencyDate & ") Order By JobDate Desc", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
											dStartDate = GetDateFromSerialNumber(oRecordset.Fields("JobDate").Value)
											dCurrencyDate = GetDateFromSerialNumber(CLng(oRequest("OwnerJobYear").Item & oRequest("OwnerJobMonth").Item & oRequest("OwnerJobDay").Item))
											If (DateDiff("d", dStartDate, dCurrencyDate) < 181) Then
												sErrorDescription = "El empleado señalado no ha completado el tiempo necesario"
												lErrorNumber = -1
											End If
'										End If
									End If
									If lErrorNumber = 0 Then
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Jobs Set OwnerID=" & aJobComponent(N_ID_OWNER_JOB) & " Where (JobID=" & aJobComponent(N_ID_JOB) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										If lErrorNumber = 0 Then
											lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
											aJobComponent(N_JOB_DATE_JOB) = oRequest("OwnerJobYear").Item & oRequest("OwnerJobMonth").Item & oRequest("OwnerJobDay").Item
											aJobComponent(N_END_DATE_JOB) = 0
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_JOB_DATE_JOB) & ", " & aJobComponent(N_END_DATE_JOB) & ", " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, NULL)
											If lErrorNumber = 0 Then
												sErrorDescription = "No se pudo leer el historial del empleado"
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeDate, EndDate, StatusID From EmployeesHistoryList Where (EmployeeID = " & aJobComponent(N_ID_OWNER_JOB) & ") And (StatusId In (1,2,3,4,5,6,7,26,50,62,63,79,80,81)) Order By EmployeeDate Desc", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
												If lErrorNumber = 0 Then
													If Not oRecordset.EOF Then
														dEndDate = AddDaysToSerialDate(oRecordset.Fields("EmployeeDate").Value, -1)
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update JobsHistoryList Set OwnerId = " & aJobComponent(N_ID_OWNER_JOB) & " Where (JobID = " & aJobComponent(N_ID_JOB) & ") And (JobDate > " & aJobComponent(N_JOB_DATE_JOB) & ") And (JobDate < " & dEndDate & ")  And ((OwnerID <> 0) Or (OwnerID <> -1))" , "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
													Else
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update JobsHistoryList Set OwnerId = " & aJobComponent(N_ID_OWNER_JOB) & " Where (JobID = " & aJobComponent(N_ID_JOB) & ") And (JobDate > " & aJobComponent(N_JOB_DATE_JOB) & ")  And ((OwnerID <> 0) Or (OwnerID <> -1))", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
													End If
													If lErrorNumber <> 0 Then
														sErrorDescription = "La titularidad no pudo aplicarse en el historial de la plaza."
													End IF
												End If
												sErrorDescription = "La titularidad se asignó correctamente"
											Else
												sErrorDescription = "No se pudo actualizar el historial de la plaza"
											End If
										Else
											sErrorDescription = "No se pudo asignar la titularidad."
										End If
									End If
								Else
									If Len(sErrorDescription) = 0 Then
										sErrorDescription = "No se pudo obtener la información de la plaza"
									End If
								End If
							Else
								If Len(sErrorDescription) = 0 Then
									sErrorDescription = "No se puede asignar la titularidad a un funcionario."
								End If
							End If
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobID, JobDate, EndDate, EmployeeID, OwnerID From JobsHistoryList Where (JobID=" & aJobComponent(N_ID_JOB) & ") Order By EndDate Desc", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CompanyID, ClassificationID, GroupGradeLevelID, IntegrationID, LevelID, WorkingHours From Positions Where PositionID = " & aJobComponent(N_POSITION_ID_JOB), "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If oRecordset.EOF Then
									lErrorNumber = -1
									sErrorDescription = "No se encontró la información del puesto relacionado"
								Else
									 aJobComponent(N_COMPANY_ID_JOB) = oRecordset.Fields("CompanyID").Value
									 aJobComponent(N_CLASSIFICATION_ID_JOB) = oRecordset.Fields("ClassificationID").Value
									 aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) = oRecordset.Fields("GroupGradeLevelID").Value
									 aJobComponent(N_INTEGRATION_ID_JOB) = oRecordset.Fields("IntegrationID").Value
									 aJobComponent(N_LEVEL_ID_JOB) = oRecordset.Fields("LevelID").Value
									 aJobComponent(D_WORKING_HOURS_JOB) = oRecordset.Fields("Workinghours").Value
								End If
								If lErrorNumber = 0 Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Jobs Set JobNumber='" & Right(("000000" & Replace(aJobComponent(S_NUMBER_JOB), "'", "")), Len("000000"))  & "', CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", ZoneID=" & aJobComponent(N_ZONE_ID_JOB) & ", AreaID=" & aJobComponent(N_AREA_ID_JOB) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ", JobTypeID=" & aJobComponent(N_JOB_TYPE_ID_JOB) & ", ShiftID=" & aJobComponent(N_SHIFT_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", OccupationTypeID=" & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", StartDate=" & aJobComponent(N_START_DATE_JOB) & ", EndDate=" & aJobComponent(N_END_DATE_JOB) & ", StatusID=" & aJobComponent(N_STATUS_ID_JOB) & ", Active=" & aJobComponent(N_ACTIVE_JOB) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & " Where (JobID=" & aJobComponent(N_ID_JOB) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								End If
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo modificar la información de la plaza."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From JobsBudgetsLKP Where (JobID=" & aJobComponent(N_ID_JOB) & ") And (StartDate=" & sDate & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudo modificar la información de la plaza."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update JobsBudgetsLKP Set EndDate=" & sDate & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (JobID=" & aJobComponent(N_ID_JOB) & ") And (EndDate=0)", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										If lErrorNumber = 0 Then
											alBudgetsID = Split(aJobComponent(S_BUDGETS_ID_JOB), ",", -1, vbBinaryCompare)
											For iIndex = 0 To UBound(alBudgetsID)
												sErrorDescription = "No se pudo modificar la información de la plaza."
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsBudgetsLKP (JobID, BudgetID, StartDate, EndDate, StartUserID, EndUserID) Values (" & aJobComponent(N_ID_JOB) & ", " & alBudgetsID(iIndex) & ", " & sDate & ", 0, " & aLoginComponent(N_USER_ID_LOGIN) & ", -1)", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
												If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
											Next
										End If
									End If
								End If
							End If
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo guardar la información de la nueva plaza."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From JobsHistoryList Where (JobID=" & aJobComponent(N_ID_JOB) & ") And (EndDate <> 0) Order by JobDate Desc", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If oRecordset.EOF Then
										sErrorDescription = "No se pudo guardar la información de la nueva plaza."
										lErrorNumber = ModifyJobHistoryList(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
										'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_START_DATE_JOB) & ", " & aJobComponent(N_END_DATE_HISTORY_JOB) & ", " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & "," & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									Else
										sErrorDescription = "No se pudo guardar la información de la nueva plaza."
										If (CLng(aJobComponent(N_ID_EMPLOYEE_JOB)) = CLng(oRecordset.Fields("EmployeeID").Value)) And (CLng(aJobComponent(N_ID_OWNER_JOB)) = CLng(oRecordset.Fields("OwnerID").Value)) And (CLng(aJobComponent(N_COMPANY_ID_JOB)) = CLng(oRecordset.Fields("CompanyID").Value)) And (CLng(aJobComponent(N_ZONE_ID_JOB)) = CLng(oRecordset.Fields("ZoneID").Value)) And (CLng(aJobComponent(N_AREA_ID_JOB)) = CLng(oRecordset.Fields("AreaID").Value)) And (CLng(aJobComponent(N_PAYMENT_CENTER_ID_JOB)) = CLng(oRecordset.Fields("PaymentCenterID").Value)) And (CLng(aJobComponent(N_POSITION_ID_JOB)) = CLng(oRecordset.Fields("PositionID").Value)) And (CLng(aJobComponent(N_JOB_TYPE_ID_JOB)) = CLng(oRecordset.Fields("JobTypeID").Value)) And (CLng(aJobComponent(N_SHIFT_ID_JOB)) = CLng(oRecordset.Fields("ShiftID").Value)) And (CLng(aJobComponent(N_JOURNEY_ID_JOB)) = CLng(oRecordset.Fields("JourneyID").Value)) And (CLng(aJobComponent(N_CLASSIFICATION_ID_JOB)) = CLng(oRecordset.Fields("ClassificationID").Value)) And (CLng(aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB)) = CLng(oRecordset.Fields("GroupGradeLevelID").Value)) And (CLng(aJobComponent(N_INTEGRATION_ID_JOB)) = CLng(oRecordset.Fields("IntegrationID").Value)) And (CLng(aJobComponent(N_OCCUPATION_TYPE_ID_JOB)) = CLng(oRecordset.Fields("OccupationTypeID").Value)) And (CLng(aJobComponent(N_SERVICE_ID_JOB)) = CLng(oRecordset.Fields("ServiceID").Value)) And (CLng(aJobComponent(N_LEVEL_ID_JOB)) = CLng(oRecordset.Fields("LevelID").Value)) And (CDbl(aJobComponent(D_WORKING_HOURS_JOB)) = CDbl(oRecordset.Fields("WorkingHours").Value)) And (CLng(aJobComponent(N_JOB_DATE_JOB)) = CLng(oRecordset.Fields("JobDate").Value)) And (CLng(aJobComponent(N_END_DATE_HISTORY_JOB)) = CLng(oRecordset.Fields("EndDate").Value)) And (CLng(aJobComponent(N_STATUS_ID_JOB)) = CLng(oRecordset.Fields("StatusID").Value)) Then
										Else
											If CLng(aJobComponent(N_JOB_DATE_JOB)) < CLng(oRecordset.Fields("JobDate").Value) Then
												'Reanudación de labores anticipada
												If CLng(oRecordset.Fields("StatusID").Value) = 5 Then
													lHistoryJobDate = CLng(oRecordset.Fields("JobDate").Value)
													lHistoryEndDate = CLng(oRecordset.Fields("EndDate").Value)
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update JobsHistoryList Set JobDate=" & aJobComponent(N_JOB_DATE_JOB) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=" & aJobComponent(N_STATUS_ID_JOB) & ", EmployeeID=" & aJobComponent(N_ID_EMPLOYEE_JOB) & " Where (JobID=" & aJobComponent(N_ID_JOB) & ") And (JobDate=" & lHistoryJobDate & ") And (EndDate=" & lHistoryEndDate & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
													oRecordset.MoveNext
													If Not oRecordset.EOF Then
														aJobComponent(N_JOB_DATE_JOB) = AddDaysToSerialDate(aJobComponent(N_JOB_DATE_JOB), -1)
														lHistoryJobDate = CLng(oRecordset.Fields("JobDate").Value)
														lHistoryEndDate = CLng(oRecordset.Fields("EndDate").Value)
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update JobsHistoryList Set EndDate=" & aJobComponent(N_JOB_DATE_JOB) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & " Where (JobID=" & aJobComponent(N_ID_JOB) & ") And (JobDate=" & lHistoryJobDate & ") And (EndDate=" & lHistoryEndDate & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
													End If
												Else
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From JobsHistoryList Where (JobID=" & aJobComponent(N_ID_JOB) & ") And (JobDate>=" & aJobComponent(N_JOB_DATE_JOB) & ")  And (EndDate<=" & aJobComponent(N_END_DATE_HISTORY_JOB) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
													If lErrorNumber = 0 Then
														If Not oRecordset.EOF Then
															If (CLng(oRecordset.Fields("StatusID").Value) = 2) Or (CLng(oRecordset.Fields("StatusID").Value) = 4) Then
																lHistoryJobDate = CLng(oRecordset.Fields("JobDate").Value)
																lHistoryEndDate = CLng(oRecordset.Fields("EndDate").Value)
																lStatusID = CLng(oRecordset.Fields("StatusID").Value)
																sQuery = "Update JobsHistoryList Set JobDate=" & aJobComponent(N_JOB_DATE_JOB) & ", EndDate=" & aJobComponent(N_END_DATE_HISTORY_JOB) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", StatusID=" & aJobComponent(N_STATUS_ID_JOB) & ", EmployeeID=" & aJobComponent(N_ID_EMPLOYEE_JOB) & " Where (JobID=" & aJobComponent(N_ID_JOB) & ") And (JobDate=" & lHistoryJobDate & ") And (EndDate=" & lHistoryEndDate & ")"
															Else
																lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From JobsHistoryList Where (JobID=" & aJobComponent(N_ID_JOB) & ") And (JobDate<=" & aJobComponent(N_JOB_DATE_JOB) & ")  And (EndDate>=" & aJobComponent(N_END_DATE_HISTORY_JOB) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)	
																lErrorNumber = L_ERR_NO_RECORDS
																sErrorDescription = "La plaza no se encuentra vacante en el período especificado."
															End If
														Else
															lHistoryJobDate = CLng(oRequest("EmployeeYear").Item & oRequest("EmployeeMonth").Item & oRequest("EmployeeDay").Item)
															lHistoryEndDate = CLng(oRequest("EmployeeEndYear").Item & oRequest("EmployeeEndMonth").Item & oRequest("EmployeeEndDay").Item)
															lStatusID = 2
															sErrorDescription = "No se pudo guardar la información histórica de la plaza."
															aJobComponent(N_STATUS_ID_JOB) = 1
															sQuery = "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & lHistoryJobDate & ", " & lHistoryEndDate & ", " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")"
														End If
														If lErrorNumber = 0 Then
															aJobComponent(N_STATUS_ID_JOB) = lStatusID
															aJobComponent(N_ID_EMPLOYEE_JOB) = -1
															lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
														End If
													End If
												End If
											Else
												lHistoryJobDate = CLng(oRecordset.Fields("JobDate").Value)
												lHistoryEndDate = CLng(oRecordset.Fields("EndDate").Value)
												If aJobComponent(N_JOB_DATE_JOB) = CLng(oRecordset.Fields("JobDate").Value) Then
													If aJobComponent(N_END_DATE_HISTORY_JOB) = 0 Then aJobComponent(N_END_DATE_HISTORY_JOB) = 30000000
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update JobsHistoryList Set EmployeeID = " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & ", ZoneID=" & aJobComponent(N_ZONE_ID_JOB) & ", AreaID=" & aJobComponent(N_AREA_ID_JOB) & ", PaymentCenterID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ", JobTypeID=" & aJobComponent(N_JOB_TYPE_ID_JOB) & ", ShiftID=" & aJobComponent(N_SHIFT_ID_JOB) & ", WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & ", JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & ", ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & ", OccupationTypeID=" & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ", LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & ", JobDate=" & aJobComponent(N_JOB_DATE_JOB) & ", EndDate=" & aJobComponent(N_END_DATE_HISTORY_JOB) & ", StatusID=" & aJobComponent(N_STATUS_ID_JOB) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (JobID=" & aJobComponent(N_ID_JOB) & ") And (JobDate=" & aJobComponent(N_JOB_DATE_JOB) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	'												If lErrorNumber = 0 Then
	'													sErrorDescription = "No se pudo guardar la información histórica de la plaza."
	'													If aJobComponent(N_END_DATE_HISTORY_JOB) < lHistoryEndDate Then
	'														aJobComponent(N_ID_EMPLOYEE_JOB) = -1
	'														aJobComponent(N_STATUS_ID_JOB) = 2
	'														'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & AddDaysToSerialDate(aJobComponent(N_END_DATE_HISTORY_JOB), 1) & ", " & lHistoryEndDate & ", " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	'														lErrorNumber = ModifyJobHistoryList(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
	'													End If
	'												End If
												Else
													lHistoryEndDate = CLng(oRecordset.Fields("EndDate").Value)
													lHistoryJobDate = CLng(oRecordset.Fields("JobDate").Value)
													If lHistoryEndDate <> 30000000 Then
														sErrorDescription = "No se pudo guardar la información histórica de la plaza."
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update JobsHistoryList Set EndDate=" & AddDaysToSerialDate(aJobComponent(N_JOB_DATE_JOB), -1) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & " Where (JobID=" & aJobComponent(N_ID_JOB) & ") And (JobDate=" & lHistoryJobDate & ") And (EndDate=" & lHistoryEndDate & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
													Else
														lHistoryEndDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update JobsHistoryList Set EndDate=" & AddDaysToSerialDate(aJobComponent(N_JOB_DATE_JOB), -1) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & " Where (JobID=" & aJobComponent(N_ID_JOB) & ") And (JobDate=" & lHistoryJobDate & ") And (EndDate=" & oRecordset.Fields("EndDate").Value & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
													End If
													If lErrorNumber = 0 Then
														If aJobComponent(N_STATUS_ID_JOB) = 4 Then
															aJobComponent(N_ID_EMPLOYEE_JOB) = -1
															sErrorDescription = "No se pudo guardar la información histórica de la plaza."
															'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_JOB_DATE_JOB) & ", " & aJobComponent(N_END_DATE_HISTORY_JOB) & ", " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
															lErrorNumber = ModifyJobHistoryList(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
															If lErrorNumber = 0 Then
																sErrorDescription = "No se pudo guardar la información histórica de la plaza."
																aJobComponent(N_STATUS_ID_JOB) = 5
																If aJobComponent(N_END_DATE_HISTORY_JOB) < lHistoryEndDate Then
																	'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & AddDaysToSerialDate(aJobComponent(N_END_DATE_HISTORY_JOB), 1) & ", " & lHistoryEndDate & ", " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
																	lErrorNumber = ModifyJobHistoryList(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
																End If
															End If
														ElseIf aJobComponent(N_STATUS_ID_JOB) = 2 Then
															aJobComponent(N_ID_EMPLOYEE_JOB) = -1
															sErrorDescription = "No se pudo guardar la información histórica de la plaza."
															'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_JOB_DATE_JOB) & ", " & aJobComponent(N_END_DATE_HISTORY_JOB) & ", " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
															lErrorNumber = ModifyJobHistoryList(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
	'														If lErrorNumber = 0 Then
	'															sErrorDescription = "No se pudo guardar la información histórica de la plaza."
	'															If aJobComponent(N_END_DATE_HISTORY_JOB) < lHistoryEndDate Then
	'																aJobComponent(N_ID_EMPLOYEE_JOB) = -1
	'																aJobComponent(N_STATUS_ID_JOB) = 2
	'																'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & AddDaysToSerialDate(aJobComponent(N_END_DATE_HISTORY_JOB), 1) & ", " & lHistoryEndDate & ", " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	'																lErrorNumber = ModifyJobHistoryList(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
	'															End If
	'														End If
														Else
															'sErrorDescription = "No se pudo guardar la información histórica de la plaza."
															'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_JOB_DATE_JOB) & ", " & lHistoryEndDate & ", " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
															sErrorDescription = "No se pudo guardar la información histórica de la plaza."
															'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_JOB_DATE_JOB) & ", " & aJobComponent(N_END_DATE_HISTORY_JOB) & ", " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
															lErrorNumber = ModifyJobHistoryList(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
	'														If lErrorNumber = 0 Then
	'															sErrorDescription = "No se pudo guardar la información histórica de la plaza."
	'															If aJobComponent(N_END_DATE_HISTORY_JOB) < lHistoryEndDate Then
	'																aJobComponent(N_ID_EMPLOYEE_JOB) = -1
	'																aJobComponent(N_STATUS_ID_JOB) = 2
	'																'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & AddDaysToSerialDate(aJobComponent(N_END_DATE_HISTORY_JOB), 1) & ", " & lHistoryEndDate & ", " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	'																lErrorNumber = ModifyJobHistoryList(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
	'															End If
	'														End If
														End If
													End If
												End If
											End If
										End If
									End If
									oRecordset.Close
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	ModifyJob = lErrorNumber
	Err.Clear
End Function

Function ModifyJobHistoryList(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To insert and update a new record in the JobHistoryList table
'Inputs:  oRequest, oADODBConnection
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyJobHistoryList"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sDate
	Dim lJobDate
	Dim lHistoryJobDate
	Dim lHistoryEndDate
	Dim sQuery

	bComponentInitialized = aJobComponent(B_COMPONENT_INITIALIZED_JOB)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeJobComponent(oRequest, aJobComponent)
	End If

	If aJobComponent(N_ID_JOB) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador de la plaza a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "JobComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If lErrorNumber = 0 Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobDate, EndDate From JobsHistoryList Where (JobID=" & aJobComponent(N_ID_JOB) & ") Order By JobDate Desc", "JobComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sQuery = "Select JobDate, EndDate From JobsHistoryList Where (JobID = " & aJobComponent(N_ID_JOB) & ") And (JobDate > " & aJobComponent(N_JOB_DATE_JOB) & ") And (EndDate < " & aJobComponent(N_END_DATE_HISTORY_JOB) & ")"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)	
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sQuery = "Delete From JobsHistoryList Where (JobID = " & aJobComponent(N_ID_JOB) & ") And (JobDate > " & aJobComponent(N_JOB_DATE_JOB) & ") And (EndDate < " & aJobComponent(N_END_DATE_HISTORY_JOB) & ")"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
						If lErrorNumber = 0 Then
							sQuery = "Select JobDate, EndDate From JobsHistoryList Where (JobID = " & aJobComponent(N_ID_JOB) & ") And (JobDate < " & aJobComponent(N_END_DATE_JOB) & ") And (EndDate > " & aJobComponent(N_END_DATE_HISTORY_JOB) & ")"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)	
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									sQuery = "Update JobsHistoryList Set JobDate = " & AddDaysToSerialDate(aJobComponent(N_END_DATE_JOB),1) & " Where (JobID = " & aJobComponent(N_ID_JOB) & ") And (JobDate < " & aJobComponent(N_END_DATE_HISTORY_JOB) & ") And (EndDate > " & aJobComponent(N_END_DATE_JOB) & ")"
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, null)
								End If
							End If
							If lErrorNumber = 0 Then
								sQuery = "Select JobDate, EndDate From JobsHistoryList Where (JobID = " & aJobComponent(N_ID_JOB) & ") And (EndDate > " & aJobComponent(N_JOB_DATE_JOB) & ") And (EndDate < " & aJobComponent(N_END_DATE_HISTORY_JOB) & ")"
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If Not oRecordset.EOF Then
									sQuery = "Update JobsHistoryList Set EndDate = " & AddDaysToSerialDate(aJobComponent(N_JOB_DATE_JOB),-1) & " Where (JobID = " & aJobComponent(N_ID_JOB) & ") And (EndDate > " & aJobComponent(N_JOB_DATE_JOB) & ") And (EndDate < " & aJobComponent(N_END_DATE_HISTORY_JOB) & ")"
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, null)
								End If
							End If
							If lErrorNumber = 0 Then
								sQuery = "Select JobDate, EndDate From JobsHistoryList Where (JobID = " & aJobComponent(N_ID_JOB) & ") And (JobDate = " & aJobComponent(N_JOB_DATE_JOB) & ")"
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If oRecordset.EOF Then
										If aJobComponent(N_END_DATE_HISTORY_JOB) = 0 Then aJobComponent(N_END_DATE_HISTORY_JOB) = 30000000
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_JOB_DATE_JOB) & ", " & aJobComponent(N_END_DATE_HISTORY_JOB) & ", " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, NULL)
									Else
										sQuery = "Update JobsHistoryList Set EndDate = " & aJobComponent(N_END_DATE_HISTORY_JOB) & " Where (JobID = " & aJobComponent(N_ID_JOB) & ") And (JobDate = " & aJobComponent(N_JOB_DATE_JOB) & ")"
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, NULL)
									End If
								End If
							End If
							If lErrorNumber <> 0 Then
								sErrorDescription = "El historial de la plaza no pudo actualizarse correctamente"
							End If
						End If
					End If
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into JobsHistoryList (JobID, JobDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aJobComponent(N_ID_JOB) & ", " & aJobComponent(N_START_DATE_JOB) & ", " & aJobComponent(N_END_DATE_HISTORY_JOB) & ", " & aJobComponent(N_ID_EMPLOYEE_JOB) & ", " & aJobComponent(N_ID_OWNER_JOB) & ", " & aJobComponent(N_COMPANY_ID_JOB) & ", " & aJobComponent(N_ZONE_ID_JOB) & ", " & aJobComponent(N_AREA_ID_JOB) & ", " & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aJobComponent(N_POSITION_ID_JOB) & ", " & aJobComponent(N_JOB_TYPE_ID_JOB) & ", " & aJobComponent(N_SHIFT_ID_JOB) & ", " & aJobComponent(D_WORKING_HOURS_JOB) & ", " & aJobComponent(N_JOURNEY_ID_JOB) & ", " & aJobComponent(N_CLASSIFICATION_ID_JOB) & ", " & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aJobComponent(N_INTEGRATION_ID_JOB) & ", " & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aJobComponent(N_SERVICE_ID_JOB) & ", " & aJobComponent(N_LEVEL_ID_JOB) & ", " & aJobComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					If lErrorNumber <> 0 Then sErrorDescription = "No se pudo actualizar el historial de la plaza."
				End If
			Else
				sErrorDescription = "No se pudo obtener el historial de la plaza"
			End If
		End If
	End If
	Set oRecordset = Nothing
	ModifyJobHistoryList = lErrorNumber
	Err.Clear
End Function

Function SetActiveForJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given job
'Inputs:  oRequest, oADODBConnection
'Outputs: aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForJob"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aJobComponent(B_COMPONENT_INITIALIZED_JOB)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeJobComponent(oRequest, aJobComponent)
	End If

	If aJobComponent(N_ID_JOB) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador de la plaza a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "JobComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo modificar la información de la plaza."
		If Len(oRequest("SetActive").Item) > 0 Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Jobs Set Active=" & CInt(oRequest("SetActive").Item) & " Where (JobID=" & aJobComponent(N_ID_JOB) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Jobs Set Active=" & aJobComponent(N_ACTIVE_JOB) & " Where (JobID=" & aJobComponent(N_ID_JOB) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	SetActiveForJob = lErrorNumber
	Err.Clear
End Function

Function RemoveJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a job from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveJob"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aJobComponent(B_COMPONENT_INITIALIZED_JOB)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeJobComponent(oRequest, aJobComponent)
	End If

	If aJobComponent(N_ID_JOB) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó la plaza a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "JobComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información de la plaza."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Jobs Where (JobID=" & aJobComponent(N_ID_JOB) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la información de la plaza."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From JobsBudgetsLKP Where (JobID=" & aJobComponent(N_ID_JOB) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información de la plaza."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From JobsHistoryList Where (JobID=" & aJobComponent(N_ID_JOB) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

			sErrorDescription = "No se pudo eliminar la información de la plaza."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Employees Set JobID=-1 Where (JobID=" & aJobComponent(N_ID_JOB) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	RemoveJob = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfJob(aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific job exists in the database
'Inputs:  aJobComponent
'Outputs: aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfJob"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aJobComponent(B_COMPONENT_INITIALIZED_JOB)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeJobComponent(oRequest, aJobComponent)
	End If

	If Len(aJobComponent(S_NUMBER_JOB)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número de la plaza para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "JobComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia de la plaza en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Jobs Where (JobID<>" & aJobComponent(N_ID_JOB) & ") And (JobNumber='" & Replace(aJobComponent(S_NUMBER_JOB), "'", "") & "')", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aJobComponent(B_IS_DUPLICATED_JOB) = True
				aJobComponent(N_ID_JOB) = -1
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfJob = lErrorNumber
	Err.Clear
End Function

Function CheckJobInformationConsistency(aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aJobComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckJobInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aJobComponent(N_ID_JOB)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador de la plaza no es un valor numérico."
		bIsCorrect = False
	End If
	If Len(aJobComponent(S_NUMBER_JOB)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El número de la plaza está vacío."
		bIsCorrect = False
	End If
	If Not IsNumeric(aJobComponent(N_ID_OWNER_JOB)) Then aJobComponent(N_ID_OWNER_JOB) = -1
	If Not IsNumeric(aJobComponent(N_ID_EMPLOYEE_JOB)) Then aJobComponent(N_ID_EMPLOYEE_JOB) = -1
	If Not IsNumeric(aJobComponent(N_ZONE_ID_JOB)) Then aJobComponent(N_ZONE_ID_JOB) = -1
	If Not IsNumeric(aJobComponent(N_AREA_ID_JOB)) Then aJobComponent(N_AREA_ID_JOB) = -1
	If Not IsNumeric(aJobComponent(N_PAYMENT_CENTER_ID_JOB)) Then aJobComponent(N_PAYMENT_CENTER_ID_JOB) = -1
	If Not IsNumeric(aJobComponent(N_POSITION_ID_JOB)) Then aJobComponent(N_POSITION_ID_JOB) = -1
	If Not IsNumeric(aJobComponent(N_JOB_TYPE_ID_JOB)) Then aJobComponent(N_JOB_TYPE_ID_JOB) = 1
	If Not IsNumeric(aJobComponent(N_SHIFT_ID_JOB)) Then aJobComponent(N_SHIFT_ID_JOB) = -1
	If Not IsNumeric(aJobComponent(N_JOURNEY_ID_JOB)) Then aJobComponent(N_JOURNEY_ID_JOB) = -1
	If Not IsNumeric(aJobComponent(N_CLASSIFICATION_ID_JOB)) Then aJobComponent(N_CLASSIFICATION_ID_JOB) = -1
	If Not IsNumeric(aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB)) Then aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) = -1
	If Not IsNumeric(aJobComponent(N_INTEGRATION_ID_JOB)) Then aJobComponent(N_INTEGRATION_ID_JOB) = -1
	If Not IsNumeric(aJobComponent(N_OCCUPATION_TYPE_ID_JOB)) Then aJobComponent(N_OCCUPATION_TYPE_ID_JOB) = -1
	If Not IsNumeric(aJobComponent(N_SERVICE_ID_JOB)) Then aJobComponent(N_SERVICE_ID_JOB) = -1
	If Not IsNumeric(aJobComponent(N_LEVEL_ID_JOB)) Then aJobComponent(N_LEVEL_ID_JOB) = -1
	If Not IsNumeric(aJobComponent(D_WORKING_HOURS_JOB)) Then aJobComponent(D_WORKING_HOURS_JOB) = -1
	If Not IsNumeric(aJobComponent(N_START_DATE_JOB)) Then aJobComponent(N_START_DATE_JOB) = 0
	If Not IsNumeric(aJobComponent(N_END_DATE_JOB)) Then aJobComponent(N_END_DATE_JOB) = 0
	If Not IsNumeric(aJobComponent(N_STATUS_ID_JOB)) Then aJobComponent(N_STATUS_ID_JOB) = 0
	If Not IsNumeric(aJobComponent(N_ACTIVE_JOB)) Then aJobComponent(N_ACTIVE_JOB) = 1

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información de la plaza contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "JobComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckJobInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayConceptValuesJobsTable(oRequest, oADODBConnection, aJobComponent, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the information about the concepts
'		  asigned to a job
'Inputs:  oRequest, oADODBConnection, iEmployeeTypeID, lPositionID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayConceptValuesJobsTable"
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	If aJobComponent(N_ID_JOB) <> -1 Then
		lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
	End If

	Select Case aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)
		Case 0
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, ConceptName, ConceptAmount From Concepts, ConceptsValues Where (Concepts.ConceptID = ConceptsValues.ConceptID) And (ConceptsValues.EmployeeTypeID = " & aJobComponent(N_EMPLOYEE_TYPE_ID_JOB) & ") And (ConceptsValues.JobStatusID = " & aJobComponent(N_STATUS_ID_JOB) & " Or JobStatusID = -1) And ((ConceptsValues.WorkingHours = " & aJobComponent(D_WORKING_HOURS_JOB) & ") Or (ConceptsValues.WorkingHours=-1)) And (ConceptsValues.EndDate=30000000) And (ConceptsValues.StatusID=1) And (ConceptsValues.ConceptID In (1,2,3,12,36,38,49)) And (ConceptsValues.EconomicZoneID = " & aJobComponent(N_ECONOMIC_ZONE_ID_JOB) & " Or EconomicZoneID = 0) And ((ConceptsValues.PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ")  Or (ConceptsValues.PositionID=-1)) Order by ConceptsValues.ConceptID", "ReportsQueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Case Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptShortName, ConceptName, ConceptAmount From Concepts, ConceptsValues Where (Concepts.ConceptID = ConceptsValues.ConceptID) And (ConceptsValues.EmployeeTypeID = " & aJobComponent(N_EMPLOYEE_TYPE_ID_JOB) & ") And (ConceptsValues.EndDate=30000000) And (ConceptsValues.StatusID=1) And (ConceptsValues.ConceptID In (1,2,3,12,36,38,49)) And (ConceptsValues.EconomicZoneID = " & aJobComponent(N_ECONOMIC_ZONE_ID_JOB) & " Or EconomicZoneID = 0) And ((ConceptsValues.PositionID=" & aJobComponent(N_POSITION_ID_JOB) & ")  Or (ConceptsValues.PositionID=-1)) Order by ConceptsValues.ConceptID", "ReportsQueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End Select

	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("Concepto,Monto", ",", -1, vbBinaryCompare)
				asCellWidths = Split("200,200", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If
				asCellAlignments = Split(",RIGHT", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("ConceptName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & "$  " & FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value)*2, 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR
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
		Else
			Call DisplayErrorMessage("Búsqueda vacía", "No existen registros en el tabulador autorizado.")
		End If
	End If

	Set oRecordset = Nothing
	DisplayConceptValuesJobsTable = lErrorNumber
	Err.Clear
End Function

Function DisplayJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a job
'Inputs:  oRequest, oADODBConnection, aJobComponent
'Outputs: aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayJob"
	Dim sShortNames
	Dim sNames
	Dim lErrorNumber

	If aJobComponent(N_ID_JOB) <> -1 Then
		lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Número de la plaza:&nbsp;</B></FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(aJobComponent(S_NUMBER_JOB)) & "</B></FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. empleado (titularidad):&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aJobComponent(N_ID_OWNER_JOB) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Empresa:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "Companies", aJobComponent(N_COMPANY_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entidad:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "Zones", aJobComponent(N_ZONE_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Área:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "Areas", aJobComponent(N_AREA_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Centro de pago:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "Areas", aJobComponent(N_PAYMENT_CENTER_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Servicio:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "Services", aJobComponent(N_SERVICE_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "Positions", aJobComponent(N_POSITION_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de plaza:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "JobTypes", aJobComponent(N_JOB_TYPE_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
			Response.Write "</TR>"
'			Response.Write "<TR>"
'				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horario:&nbsp;</FONT></TD>"
'				Call GetNameFromTable(oADODBConnection, "Shifts", aJobComponent(N_SHIFT_ID_JOB), "", "", sNames, sErrorDescription)
'				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
'			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Jornada:&nbsp;</B></FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aJobComponent(D_WORKING_HOURS_JOB) & " hrs.</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Turno:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "Journeys", aJobComponent(N_JOURNEY_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de ocupación:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "JobTypes", N_JOB_TYPE_ID_JOB, "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
			Response.Write "</TR>"
			If aJobComponent(N_LEVEL_ID_JOB) <> -1 Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nivel:&nbsp;</FONT></TD>"
					Call GetNameFromTable(oADODBConnection, "Levels", aJobComponent(N_LEVEL_ID_JOB), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
			Else
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Grupo Grado Nivel:&nbsp;</FONT></TD>"
					Call GetNameFromTable(oADODBConnection, "GroupGradeLevels", aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Integración:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(aJobComponent(N_INTEGRATION_ID_JOB)) & "</B></FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Clasificación:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(aJobComponent(N_CLASSIFICATION_ID_JOB)) & "</B></FONT></TD>"
				Response.Write "</TR>"
			End If
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aJobComponent(N_START_DATE_JOB), -1, -1, -1) & "</FONT></TD>"
			Response.Write "</TR>"
			If aJobComponent(N_END_DATE_JOB) > 0 And aJobComponent(N_END_DATE_JOB) < 30000000 Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de término:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aJobComponent(N_END_DATE_JOB), -1, -1, -1) & "</FONT></TD>"
				Response.Write "</TR>"
			Else
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de término:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">a la fecha</FONT></TD>"
				Response.Write "</TR>"
			End If
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estatus:&nbsp;</FONT></TD>"
				Call GetNameFromTable(oADODBConnection, "StatusJobs", aJobComponent(N_STATUS_ID_JOB), "", "", sNames, sErrorDescription)
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Activo:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayYesNo(aJobComponent(N_ACTIVE_JOB), True) & "</FONT></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE></DIV><BR />"
	End If

	DisplayJob = lErrorNumber
	Err.Clear
End Function

Function DisplayJobFormSp(oRequest, oADODBConnection, sAction, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a job from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aJobComponent
'Outputs: aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayJobForm"
	Dim oRecordset
	Dim sShortNames
	Dim sNames
	Dim lErrorNumber

	If aJobComponent(N_ID_JOB) <> -1 Then
		lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
		aEmployeeComponent(N_JOB_ID_EMPLOYEE) = aJobComponent(N_ID_JOB)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckJobFields(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If Len(oRequest("Delete").Item) > 0 Then Response.Write "return true;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckJobFields" & vbNewLine

			Response.Write "function ShowJobFields() {" & vbNewLine
				Response.Write "var sEmployeeTypeID = document.JobFrm.EmployeeTypeID.value;" & vbNewLine
				Response.Write "if (sEmployeeTypeID != '1') {" & vbNewLine
					Response.Write "HideDisplay(document.all['ClassificationIDDiv']);" & vbNewLine
					Response.Write "HideDisplay(document.all['IntegrationIDDiv']);" & vbNewLine
					Response.Write "HideDisplay(document.all['GroupGradeLevelIDDiv']);" & vbNewLine
					Response.Write "ShowDisplay(document.all['LevelIDDiv']);" & vbNewLine
				Response.Write "} else {" & vbNewLine
					Response.Write "ShowDisplay(document.all['ClassificationIDDiv']);" & vbNewLine
					Response.Write "ShowDisplay(document.all['IntegrationIDDiv']);" & vbNewLine
					Response.Write "ShowDisplay(document.all['GroupGradeLevelIDDiv']);" & vbNewLine
					Response.Write "HideDisplay(document.all['LevelIDDiv']);" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "SearchRecord(document.JobFrm.CompanyID.value, 'AreasForCompany', 'SearchPositionsCatalogsIFrame', 'JobFrm');" & vbNewLine
			Response.Write "} // End of ShowJobFields" & vbNewLine

			Response.Write "function ShowAreaFields() {" & vbNewLine
				Response.Write "SearchRecord(document.JobFrm.AreaID.value, 'ZoneForArea', 'SearchZoneForAreaCatalogsIFrame', 'JobFrm.ZoneID');" & vbNewLine
				Response.Write "HidePopupItem('WaitDiv', parent.window.document.all['WaitDiv']);" & vbNewLine
			Response.Write "} // End of ShowAreaFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""JobFrm"" ID=""JobFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckJobFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Jobs"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobID"" ID=""JobIDHdn"" VALUE=""" & aJobComponent(N_ID_JOB) & """ />"
			If Len(oRequest("EmployeeID").Item) > 0 Then Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & oRequest("EmployeeID").Item & """ />"
			If Len(oRequest("Tab").Item) > 0 Then Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""" & oRequest("Tab").Item & """ />"
			Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchPositionsCatalogsIFrame"" FRAMEBORDER=""0"" WIDTH=""320"" HEIGHT=""0""></IFRAME><BR />"
			Response.Write "<B>Datos de la plaza</B><BR /><BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Número de la plaza:&nbsp;</B></FONT></TD>"
					If aJobComponent(N_ID_JOB) = -1 Then lErrorNumber = GetConsecutiveID(oADODBConnection, 100, aJobComponent(S_NUMBER_JOB), sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B><INPUT TYPE=""HIDDEN"" NAME=""JobNumber"" ID=""JobNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & aJobComponent(S_NUMBER_JOB) & """ CLASS=""TextFields"" />" & CleanStringForHTML(aJobComponent(S_NUMBER_JOB)) & "</B></FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""CompanyDiv"" ID=""CompanyDiv"">"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Empresa:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""CompanyID"" ID=""CompanyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Companies", "CompanyID", "CompanyShortName, CompanyName", "(ParentID>-1) And (Active=1)", "CompanyShortName", aEmployeeComponent(N_COMPANY_ID_EMPLOYEE), "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						If aJobComponent(N_ID_JOB) = -1 Then
							Response.Write "<SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowPopupItem('WaitDiv', parent.window.document.all['WaitDiv'], false); SearchRecord(this.value, 'PositionsCatalogsLKP', 'SearchPositionsCatalogsIFrame', 'JobFrm');"">"
								sErrorDescription = "No se pudo obtener la información del registro."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Positions.PositionID, Positions.PositionShortName, Positions.PositionName, Positions.EmployeeTypeID, EmployeeTypeName, CompanyName, GroupGradeLevelShortName, GroupGradeLevelName, LevelName, ClassificationID, IntegrationID, WorkingHours From Positions, EmployeeTypes, Companies, GroupGradeLevels, Levels Where (Positions.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Positions.CompanyID=Companies.CompanyID) And (Positions.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (Positions.LevelID=Levels.LevelID) And (Positions.EndDate=30000000) And (EmployeeTypes.EndDate=30000000) And (Companies.EndDate=30000000) And (GroupGradeLevels.EndDate=30000000) And (Levels.EndDate=30000000) And (PositionID>-1) Order By Positions.PositionShortName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									Do While Not oRecordset.EOF
										Response.Write "<OPTION VALUE=""" & CStr(oRecordset.Fields("PositionID").Value) & """"
											If aJobComponent(N_POSITION_ID_JOB) = CLng(oRecordset.Fields("PositionID").Value) Then Response.Write " SELECTED=""1"""
										Response.Write ">" & CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value) & " (Tabulador: " & CStr(oRecordset.Fields("EmployeeTypeName").Value) & ", Compañía: " & CStr(oRecordset.Fields("CompanyName").Value) & ", "
											If CLng(oRecordset.Fields("EmployeeTypeID").Value) = 1 Then
												Response.Write "GGN: " & CStr(oRecordset.Fields("GroupGradeLevelShortName").Value) & ", Clasificación:" & CStr(oRecordset.Fields("ClassificationID").Value) & ", Integración: " & CStr(oRecordset.Fields("IntegrationID").Value)
											Else
												Response.Write "Nivel: " & CStr(oRecordset.Fields("LevelName").Value)
											End If
										Response.Write ", Horas laboradas: " & CStr(oRecordset.Fields("WorkingHours").Value) & ")" & "</OPTION>"
										oRecordset.MoveNext
										If Err.number <> 0 Then Exit Do
									Loop
								End If
							Response.Write "</SELECT>"
						Else
							Call GetNameFromTable(oADODBConnection, "Positions", aJobComponent(N_POSITION_ID_JOB), "", "", sNames, sErrorDescription)
							Response.Write CleanStringForHTML(sNames)
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PositionID"" ID=""PositionIDHdn"" VALUE=""" & aJobComponent(N_POSITION_ID_JOB) & """ />"
						End If
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				If (CInt(aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)) <> 1) And (CInt(aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)) > -1) Then
					aJobComponent(N_CLASSIFICATION_ID_JOB) = -1
					aJobComponent(N_INTEGRATION_ID_JOB) = -1
				Else
					If aJobComponent(N_CLASSIFICATION_ID_JOB) = -1 Then aJobComponent(N_CLASSIFICATION_ID_JOB) = 0
					If aJobComponent(N_INTEGRATION_ID_JOB) = -1 Then aJobComponent(N_INTEGRATION_ID_JOB) = 0
				End If
'				If aJobComponent(N_ID_JOB) <> -1 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Servicio:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""ServiceID"" ID=""ServiceIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Services", "ServiceID", "ServiceShortName, ServiceName", "", "ServiceShortName", aJobComponent(N_SERVICE_ID_JOB), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"

					Response.Write "<TR NAME=""ClassificationIDDiv"" ID=""ClassificationIDDiv"""
						If (CInt(aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)) <> 1) And (CInt(aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)) > -1) Then Response.Write " STYLE=""display: none"""
					Response.Write ">"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clasificación:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""ClassificationID"" ID=""ClassificationIDTxt"" SIZE=""2"" MAXLENGTH=""1"" VALUE=""" & aJobComponent(N_CLASSIFICATION_ID_JOB) & """ READONLY=""1"" CLASS=""SpecialTextFields"" /></TD>"
					Response.Write "</TR>"

					Response.Write "<TR NAME=""GroupGradeLevelIDDiv"" ID=""GroupGradeLevelIDDiv"""
						If (CInt(aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)) <> 1) And (CInt(aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)) > -1) Then Response.Write " STYLE=""display: none"""
					Response.Write ">"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Grupo, grado, nivel:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							If aJobComponent(N_STATUS_ID_JOB) = 2 Then
								Response.Write "<SELECT NAME=""GroupGradeLevelID"" ID=""GroupGradeLevelIDCmb"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "GroupGradeLevels", "GroupGradeLevelID", "GroupGradeLevelName", "GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB), "GroupGradeLevelName", aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB), "Ninguno;;;-1", sErrorDescription)
								Response.Write "</SELECT>"
							Else
								Call GetNameFromTable(oADODBConnection, "GroupGradeLevels", aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB), "", "", sNames, sErrorDescription)
								Response.Write CleanStringForHTML(sNames)
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GroupGradeLevelID"" ID=""GroupGradeLevelIDHdn"" VALUE=""" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & """ />"
							End If
						Response.Write "</FONT></TD>"
'						Response.Write "<TD><SELECT NAME=""GroupGradeLevelID"" ID=""GroupGradeLevelIDCmb"" SIZE=""1"" CLASS=""Lists"">"
'							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "GroupGradeLevels", "GroupGradeLevelID", "GroupGradeLevelName", "GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB), "GroupGradeLevelName", aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB), "Ninguno;;;-1", sErrorDescription)
'						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR NAME=""IntegrationIDDiv"" ID=""IntegrationIDDiv"""
						If (CInt(aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)) <> 1) And (CInt(aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)) > -1) Then Response.Write " STYLE=""display: none"""
					Response.Write ">"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Integración:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""IntegrationID"" ID=""IntegrationIDTxt"" VALUE=""" & aJobComponent(N_INTEGRATION_ID_JOB) & """ SIZE=""2"" MAXLENGTH=""2"" READONLY CLASS=""SpecialTextFields"" /></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Turno:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""JourneyID"" ID=""JourneyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Journeys", "JourneyID", "JourneyShortName, JourneyName", "", "JourneyShortName", aJobComponent(N_JOURNEY_ID_JOB), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horarios:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""ShiftID"" ID=""ShiftIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Shifts", "ShiftID", "ShiftShortName, ShiftName", "", "ShiftShortName", aJobComponent(N_SHIFT_ID_JOB), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR NAME=""WorkingHoursDiv"" ID=""WorkingHoursDiv"">"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horas laboradas:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""WorkingHours"" ID=""WorkingHoursTxt"" VALUE=""" & aJobComponent(D_WORKING_HOURS_JOB) & """ SIZE=""4"" MAXLENGTH=""4"" READONLY CLASS=""SpecialTextFields"" /></TD>"
					Response.Write "</TR>"
					Response.Write "<TR NAME=""LevelIDDiv"" ID=""LevelIDDiv"""
						If (CInt(aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)) = 1) And (CInt(aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)) > -1) Then Response.Write " STYLE=""display: none"""
					Response.Write ">"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nivel:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""LevelID"" ID=""LevelIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Levels", "LevelID", "LevelName", "LevelID=" & aJobComponent(N_LEVEL_ID_JOB), "LevelID", aJobComponent(N_LEVEL_ID_JOB), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"

					Response.Write "<TR><TD COLSPAN=""2""><IMG SRC=""Images/DotBlue.gif"" WIDTH=""640"" HEIGHT=""1"" /></TD></TR>"

					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. del empleado que tiene la titularidad:&nbsp;</FONT></TD>"
						Response.Write "<TD>"
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""OwnerID"" ID=""OwnerIDTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""
								If aJobComponent(N_ID_OWNER_JOB) > -0 Then Response.Write aJobComponent(N_ID_OWNER_JOB)
							Response.Write """ CLASS=""TextFields"" />"
							Response.Write "<A HREF=""javascript: SearchRecord(document.JobFrm.OwnerID.value, 'EmployeeNumber', 'SearchEmployeeNumberIFrame', 'JobFrm.OwnerID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""320"" HEIGHT=""22""></IFRAME>"
						Response.Write "</TD>"
					Response.Write "</TR>"
'				End If
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de tabulador:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeShortName, EmployeeTypeName", "(EndDate=30000000) And (Active=1)", "EmployeeTypeShortName", aJobComponent(N_EMPLOYEE_TYPE_ID_JOB), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de ocupación:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""OccupationTypeID"" ID=""OccupationTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "OccupationTypes", "OccupationTypeID", "OccupationTypeName", "", "OccupationTypeID", aJobComponent(N_OCCUPATION_TYPE_ID_JOB), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estatus:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""StatusID"" ID=""StatusIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						If aJobComponent(N_ID_JOB) = -1 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusJobs", "StatusID", "StatusName", "(StatusID=2)", "StatusName", aJobComponent(N_STATUS_ID_JOB), "Ninguno;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusJobs", "StatusID", "StatusName", "(StatusID>-1) And (Active=1)", "StatusName", aJobComponent(N_STATUS_ID_JOB), "Ninguno;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"

				'Response.Write "<TR>"
				'	Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entidad:&nbsp;</FONT></TD>"
				'	Response.Write "<TD><SELECT NAME=""ZoneID"" ID=""ZoneIDCmb"" SIZE=""1"" CLASS=""Lists"">"
				'		Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneCode, ZoneName", "(ZoneID>-1) And (ParentID=-1)", "ZoneCode", aJobComponent(N_ZONE_ID_JOB), "Ninguna;;;-1", sErrorDescription)
				'	Response.Write "</SELECT></TD>"
				'Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Centro de trabajo:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						If aJobComponent(N_ID_JOB) = -1 Then
							Response.Write "<SELECT NAME=""AreaID"" ID=""AreaIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""SearchRecord(this.value, 'ZoneForArea', 'SearchZoneForAreaCatalogsIFrame', 'JobFrm.ZoneID')""></SELECT>"
						Else
							Call GetNameFromTable(oADODBConnection, "Areas", aJobComponent(N_AREA_ID_JOB), "", "", sNames, sErrorDescription)
							Response.Write CleanStringForHTML(sNames)
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaID"" ID=""AreaIDHdn"" VALUE=""" & aJobComponent(N_AREA_ID_JOB) & """ />"
						End If
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entidad:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						If aJobComponent(N_ID_JOB) = -1 Then
							Call GetNameFromTable(oADODBConnection, "Zones", aJobComponent(N_ZONE_ID_JOB), "", "", sNames, sErrorDescription)
							Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchZoneForAreaCatalogsIFrame"" FRAMEBORDER=""0"" WIDTH=""600"" HEIGHT=""20""></IFRAME>"
						Else
							Call GetNameFromTable(oADODBConnection, "Zones", aJobComponent(N_ZONE_ID_JOB), "", "", sNames, sErrorDescription)
							Response.Write CleanStringForHTML(sNames)
						End If
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ZoneID"" ID=""ZoneIDHdn"" VALUE=""" & aJobComponent(N_ZONE_ID_JOB) & """ />"
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Centro de pago:&nbsp;</FONT>"
					Response.Write "</TD>"
					Response.Write "<TD>"
						Response.Write "<SELECT NAME=""PaymentCenterID"" ID=""PaymentCenterIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(AreaID>-1) And (Active=1)", "AreaCode", aJobComponent(N_PAYMENT_CENTER_ID_JOB), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT><BR />"
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
					Response.Write "<TD>" & DisplayDateCombosUsingSerial(aJobComponent(N_START_DATE_JOB), "Start", N_START_YEAR, Year(Date()) + 1, True, False) & "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de término:&nbsp;</FONT></TD>"
					Response.Write "<TD>" & DisplayDateCombosUsingSerial(aJobComponent(N_END_DATE_JOB), "End", N_START_YEAR, Year(Date()) + 10, True, True) & "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Activo:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" ID=""ActiveRd"" VALUE=""1"""
							If aJobComponent(N_ACTIVE_JOB) = 1 Then Response.Write " CHECKED=""1"""
						Response.Write " />Sí&nbsp;&nbsp;&nbsp;"
						Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" ID=""ActiveRd"" VALUE=""0"""
							If aJobComponent(N_ACTIVE_JOB) = 0 Then Response.Write " CHECKED=""0"""
						Response.Write " />No&nbsp;&nbsp;&nbsp;"
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"
			Response.Write "<BR /><BR />"

			'Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""480"" HEIGHT=""1"" /><BR /><BR />"
			'Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Partidas presupuestales:<BR /></FONT>"
			'Response.Write "&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetID"" ID=""BudgetIDLst"" SIZE=""10"" MULTIPLE=""1"" CLASS=""Lists"">"
		    '	Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Budgets As Partida, Budgets As SubPartida, Budgets As Tipo", "Tipo.BudgetID", "'Partida: ' As Temp1, Partida.BudgetName, 'SubPartida: ' As Temp2, SubPartida.BudgetName, 'Tipo: ' As Temp3, Tipo.BudgetName", "(Partida.BudgetID=SubPartida.ParentID) And (SubPartida.BudgetID=Tipo.ParentID) And (SubPartida.ParentID>-1) And (Tipo.BudgetID>-1) And (Tipo.Active=1)", "Partida.BudgetName, SubPartida.BudgetName, Tipo.BudgetName", aJobComponent(S_BUDGETS_ID_JOB), "Ninguna;;;-1", sErrorDescription)
			'Response.Write "</SELECT><BR />"
			'Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""480"" HEIGHT=""1"" /><BR /><BR />"
			If CLng(aJobComponent(N_ID_EMPLOYEE_JOB)) = -1 Then
				If aJobComponent(N_ID_JOB) = -1 Then
					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
				ElseIf Len(oRequest("Delete").Item) > 0 Then
					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveJobWngDiv']); JobFrm.Remove.focus()"" />"
				Else
					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
				End If
				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			End If
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveJobWngDiv", "¿Está seguro que desea borrar la plaza de la base de datos?")
		Response.Write "</FORM>"
		If aJobComponent(N_ID_JOB) = -1 Then
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "ShowPopupItem('WaitDiv', parent.window.document.all['WaitDiv'], false);" & vbNewLine
				Response.Write "SearchRecord(document.JobFrm.PositionID.value, 'PositionsCatalogsLKP', 'SearchPositionsCatalogsIFrame', 'JobFrm');" & vbNewLine
				'Response.Write "ShowJobFields();" & vbNewLine
				'Response.Write "SendURLValuesToForm('ClassificationID=" & aJobComponent(N_CLASSIFICATION_ID_JOB) & "&IntegrationID=" & aJobComponent(N_INTEGRATION_ID_JOB) & "&WorkingHours=" & aJobComponent(D_WORKING_HOURS_JOB) & "&CompanyID=" & aJobComponent(N_COMPANY_ID_JOB) & "&GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & "&JourneyID=" & aJobComponent(N_JOURNEY_ID_JOB) & "&LevelID=" & aJobComponent(N_LEVEL_ID_JOB) & "', document.JobFrm);" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
	End If

	DisplayJobForm = lErrorNumber
	Err.Clear
End Function

Function DisplayJobForm(oRequest, oADODBConnection, sAction, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a job from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aJobComponent
'Outputs: aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayJobForm"
	Dim sAreasCondition
	Dim oRecordset
	Dim sShortNames
	Dim sNames
	Dim lErrorNumber
	Dim oVacancyRecordset
	Dim oOwnerRecordset
	Dim lCurrentDate

	If aJobComponent(N_ID_JOB) <> -1 Then
		lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
		aEmployeeComponent(N_JOB_ID_EMPLOYEE) = aJobComponent(N_ID_JOB)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine

			Response.Write "function TotalDays(iDay, iMonth, iYear){" & vbNewLine
				Response.Write "iMonth = (iMonth + 9) % 12;" & vbNewLine
				Response.Write "iYear = iYear - Math.floor(iMonth/10);" & vbNewLine
				Response.Write "return (365 * iYear + Math.floor(iYear/4) - Math.floor(iYear/100) + Math.floor(iYear/400) + Math.floor((iMonth * 306 + 5)/10) + iDay - 1)" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "function GetDifferenceBetweenDates(iDay1, iMonth1, iYear1, iDay2, iMonth2, iYear2){" & vbNewLine
				Response.Write "return TotalDays(iDay2, iMonth2, iYear2) - TotalDays(iDay1, iMonth1, iYear1)" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "function CheckOwnerField(oForm) {" & vbNewLine
				Response.Write "if (oForm.OwnerID.value.length == 0) {" & vbNewLine
				Response.Write "alert ('Favor de indicar el número de empleado');" & vbNewLine
				Response.Write "return false;" & vbNewLine
				Response.Write "}else{" & vbNewLine
				Response.Write "return true;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (parseInt(oForm.JobDay.value) + parseInt(oForm.JobMonth.value) + parseInt(oForm.JobYear.value) != 0) {" & vbNewLine
					Response.Write "if ((parseInt('1' + oForm.OwnerJobDay.value) - 100) * (parseInt('1' + oForm.OwnerJobMonth.value) - 100) * parseInt(oForm.OwnerJobYear.value) == 0) {" & vbNewLine
						Response.Write "alert('Favor de introducir una fecha de inicio de la titularidad válida.');" & vbNewLine
						Response.Write "oForm.OwnerJobDay.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "}" & vbNewLine

			Response.Write "function CheckJobFields(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If aJobComponent(N_ID_JOB) <> -1 Then
						Response.Write "if (parseInt(oForm.JobEndDay.value) + parseInt(oForm.JobEndMonth.value) + parseInt(oForm.JobEndYear.value) != 0) {" & vbNewLine
							Response.Write "if ((parseInt('1' + oForm.JobEndDay.value) - 100) * (parseInt('1' + oForm.JobEndMonth.value) - 100) * parseInt(oForm.JobEndYear.value) == 0) {" & vbNewLine
								Response.Write "alert('Favor de introducir una fecha de fin válida.');" & vbNewLine
								Response.Write "oForm.JobEndDay.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (parseInt(oForm.JobDay.value) + parseInt(oForm.JobMonth.value) + parseInt(oForm.JobYear.value) != 0) {" & vbNewLine
							Response.Write "if ((parseInt('1' + oForm.JobDay.value) - 100) * (parseInt('1' + oForm.JobMonth.value) - 100) * parseInt(oForm.JobYear.value) == 0) {" & vbNewLine
								Response.Write "alert('Favor de introducir una fecha de inicio válida.');" & vbNewLine
								Response.Write "oForm.JobDay.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if ((parseInt('1' + oForm.JobEndDay.value) - 100) * (parseInt('1' + oForm.JobEndMonth.value) - 100) * parseInt(oForm.JobEndYear.value) != 0) {" & vbNewLine
							Response.Write "if (((parseInt('1' + oForm.JobDay.value) - 100) + (parseInt('1' + oForm.JobMonth.value) - 100) * 100 + parseInt(oForm.JobYear.value) * 10000) > ((parseInt('1' + oForm.JobEndDay.value) - 100) + (parseInt('1' + oForm.JobEndMonth.value) - 100) * 100 + parseInt(oForm.JobEndYear.value) * 10000)) {" & vbNewLine
								Response.Write "alert('Favor de verificar la vigencia.');" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "}" & vbNewLine

						Response.Write "if (GetDifferenceBetweenDates(parseInt('1' + oForm.JobDay.value) - 100, parseInt('1' + oForm.JobMonth.value) - 100, parseInt(oForm.JobYear.value), parseInt('1' + oForm.StartDateHdn.value.substr(6,2)) - 100, parseInt('1' + oForm.StartDateHdn.value.substr(4,2)) - 100, parseInt(oForm.StartDateHdn.value.substr(0,4))) > 0) {" & vbNewLine
							Response.Write "alert ('La fecha inicial de vigencia del cambio es anterior a la vigencia de la plaza, favor de verificar.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (oForm.EndDateHdn.value != 30000000) {" & vbNewLine
							Response.Write "if (GetDifferenceBetweenDates(parseInt('1' + oForm.JobEndDay.value) - 100, parseInt('1' + oForm.JobEndMonth.value) - 100, parseInt(oForm.JobEndYear.value), parseInt('1' + oForm.EndDateHdn.value.substr(6,2)) - 100, parseInt('1' + oForm.EndDateHdn.value.substr(4,2)) - 100, parseInt(oForm.EndDateHdn.value.substr(0,4))) > 0) {" & vbNewLine
								Response.Write "alert ('La fecha final de vigencia del cambio es posterior a la vigencia de la plaza, favor de verificar.');" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "}"
						Response.Write "if (GetDifferenceBetweenDates(parseInt('1' + oForm.JobDay.value) - 100, parseInt('1' + oForm.JobMonth.value) - 100, parseInt(oForm.JobYear.value), parseInt('1' + oForm.VacancyStartDateHdn.value.substr(6,2)) - 100, parseInt('1' + oForm.VacancyStartDateHdn.value.substr(4,2)) - 100, parseInt(oForm.VacancyStartDateHdn.value.substr(0,4))) > 0) {" & vbNewLine
							Response.Write "alert ('La fecha inicial de vigencia del cambio es anterior a la vacancia de la plaza, favor de verificar.');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					Else
						Response.Write "if (parseInt(oForm.EndDay.value) + parseInt(oForm.EndMonth.value) + parseInt(oForm.EndYear.value) != 0) {" & vbNewLine
							Response.Write "if ((parseInt('1' + oForm.EndDay.value) - 100) * (parseInt('1' + oForm.EndMonth.value) - 100) * parseInt(oForm.EndYear.value) == 0) {" & vbNewLine
								Response.Write "alert('Favor de introducir una fecha de fin válida.');" & vbNewLine
								Response.Write "oForm.EndDay.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (parseInt(oForm.StartDay.value) + parseInt(oForm.StartMonth.value) + parseInt(oForm.StartYear.value) != 0) {" & vbNewLine
							Response.Write "if ((parseInt('1' + oForm.StartDay.value) - 100) * (parseInt('1' + oForm.StartMonth.value) - 100) * parseInt(oForm.StartYear.value) == 0) {" & vbNewLine
								Response.Write "alert('Favor de introducir una fecha de inicio válida.');" & vbNewLine
								Response.Write "oForm.StartDay.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if ((parseInt('1' + oForm.EndDay.value) - 100) * (parseInt('1' + oForm.EndMonth.value) - 100) * parseInt(oForm.EndYear.value) != 0) {" & vbNewLine
							Response.Write "if (((parseInt('1' + oForm.StartDay.value) - 100) + (parseInt('1' + oForm.StartMonth.value) - 100) * 100 + parseInt(oForm.StartYear.value) * 10000) > ((parseInt('1' + oForm.EndDay.value) - 100) + (parseInt('1' + oForm.EndMonth.value) - 100) * 100 + parseInt(oForm.EndYear.value) * 10000)) {" & vbNewLine
								Response.Write "alert('Favor de verificar la vigencia.');" & vbNewLine
								Response.Write "oForm.StartDay.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Response.Write "}" & vbNewLine
					End If
					If Len(oRequest("Delete").Item) > 0 Then Response.Write "return true;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckJobFields" & vbNewLine

			Response.Write "function ShowJobFields() {" & vbNewLine
				Response.Write "var sEmployeeTypeID = document.JobFrm.EmployeeTypeID.value;" & vbNewLine
				Response.Write "if (sEmployeeTypeID != '1') {" & vbNewLine
					Response.Write "HideDisplay(document.all['ClassificationIDDiv']);" & vbNewLine
					Response.Write "HideDisplay(document.all['IntegrationIDDiv']);" & vbNewLine
					Response.Write "HideDisplay(document.all['GroupGradeLevelIDDiv']);" & vbNewLine
					Response.Write "ShowDisplay(document.all['LevelIDDiv']);" & vbNewLine
				Response.Write "} else {" & vbNewLine
					Response.Write "ShowDisplay(document.all['ClassificationIDDiv']);" & vbNewLine
					Response.Write "ShowDisplay(document.all['IntegrationIDDiv']);" & vbNewLine
					Response.Write "ShowDisplay(document.all['GroupGradeLevelIDDiv']);" & vbNewLine
					Response.Write "HideDisplay(document.all['LevelIDDiv']);" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "SearchRecord(document.JobFrm.CompanyID.value, 'AreasForCompany', 'SearchPositionsCatalogsIFrame', 'JobFrm');" & vbNewLine
			Response.Write "} // End of ShowJobFields" & vbNewLine

			Response.Write "function ShowAreaFields() {" & vbNewLine
				Response.Write "SearchRecord(document.JobFrm.AreaID.value, 'ZoneForArea', 'SearchZoneForAreaCatalogsIFrame', 'JobFrm.ZoneID');" & vbNewLine
				Response.Write "HidePopupItem('WaitDiv', parent.window.document.all['WaitDiv']);" & vbNewLine
			Response.Write "} // End of ShowAreaFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		If (aJobComponent(N_STATUS_ID_JOB) = 2) Or (aJobComponent(N_ID_JOB) = -1) Then
			Response.Write "<FORM NAME=""JobFrm"" ID=""JobFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckJobFields(this)"">"
		Else
			Response.Write "<FORM NAME=""JobFrm"" ID=""JobFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckOwnerField(this);"">"
		End If
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Jobs"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobID"" ID=""JobIDHdn"" VALUE=""" & aJobComponent(N_ID_JOB) & """ />"
			If Len(oRequest("EmployeeID").Item) > 0 Then Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & oRequest("EmployeeID").Item & """ />"
			If Len(oRequest("Tab").Item) > 0 Then Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""" & oRequest("Tab").Item & """ />"
			Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchPositionsCatalogsIFrame"" FRAMEBORDER=""0"" WIDTH=""320"" HEIGHT=""0""></IFRAME><BR />"
			Response.Write "<B>Datos de la plaza</B><BR /><BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Número de la plaza:&nbsp;</B></FONT></TD>"
					If aJobComponent(N_ID_JOB) = -1 Then lErrorNumber = GetConsecutiveID(oADODBConnection, 100, aJobComponent(S_NUMBER_JOB), sErrorDescription)
					'Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B><INPUT TYPE=""TEXT"" NAME=""JobNumber"" ID=""JobNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & aJobComponent(S_NUMBER_JOB) & """ CLASS=""TextFields"" />" & CleanStringForHTML(aJobComponent(S_NUMBER_JOB)) & "</B></FONT></TD>"
					If oRequest("Action").Item = "Jobs" Then
						If Len(oRequest("Change").Item) > 0 Then
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>" & aJobComponent(S_NUMBER_JOB) & "</B></FONT></TD>"
						Else
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B><INPUT TYPE=""TEXT"" NAME=""JobNumber"" ID=""JobNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & aJobComponent(S_NUMBER_JOB) & """ CLASS=""TextFields"" /></B></FONT></TD>"
						End If
					Else
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>" & aJobComponent(S_NUMBER_JOB) & "</B></FONT></TD>"
					End If
				Response.Write "</TR>"
				Response.Write "<TR>"
						If (aJobComponent(N_ID_JOB) = -1) Or (aJobComponent(N_STATUS_ID_JOB) = 2) Then
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							'Response.Write "<SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowPopupItem('WaitDiv', parent.window.document.all['WaitDiv'], false); SearchRecord(this.value, 'PositionsCatalogsLKP', 'SearchPositionsCatalogsIFrame', 'JobFrm');"">"
							Response.Write "<SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								sErrorDescription = "No se pudo obtener la información del registro."
								lCurrentDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
								If CInt(oRequest("ReasonID").Item) = 59 Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Positions.PositionID, Positions.PositionShortName, Positions.PositionName, Positions.EmployeeTypeID, EmployeeTypeName, CompanyName, GroupGradeLevelShortName, GroupGradeLevelName, LevelName, ClassificationID, IntegrationID, WorkingHours From Positions, EmployeeTypes, Companies, GroupGradeLevels, Levels Where (Positions.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Positions.CompanyID=Companies.CompanyID) And (Positions.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (Positions.LevelID=Levels.LevelID) And (Positions.EndDate>=" & lCurrentDate & ") And (EmployeeTypes.EndDate>=" & lCurrentDate & ") And (Companies.EndDate>=" & lCurrentDate & ") And (GroupGradeLevels.EndDate>=" & lCurrentDate & ") And (Levels.EndDate>=" & lCurrentDate & ") And (PositionID>-1) And PositionID<>" & L_HONORARY_POSITION_ID & " And Positions.Depreciated = 0 Order By Positions.PositionShortName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
								Else
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Positions.PositionID, Positions.PositionShortName, Positions.PositionName, Positions.EmployeeTypeID, EmployeeTypeName, CompanyName, GroupGradeLevelShortName, GroupGradeLevelName, LevelName, ClassificationID, IntegrationID, WorkingHours From Positions, EmployeeTypes, Companies, GroupGradeLevels, Levels Where (Positions.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Positions.CompanyID=Companies.CompanyID) And (Positions.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (Positions.LevelID=Levels.LevelID) And (Positions.EndDate>=" & lCurrentDate & ") And (EmployeeTypes.EndDate>=" & lCurrentDate & ") And (Companies.EndDate>=" & lCurrentDate & ") And (GroupGradeLevels.EndDate>=" & lCurrentDate & ") And (Levels.EndDate>=" & lCurrentDate & ") And (PositionID>-1) And PositionID<>" & L_HONORARY_POSITION_ID & " Order By Positions.PositionShortName", "SearchRecord.asp", "_root", 000, sErrorDescription, oRecordset)
								End If
								If lErrorNumber = 0 Then
									Do While Not oRecordset.EOF
										Response.Write "<OPTION VALUE=""" & CStr(oRecordset.Fields("PositionID").Value) & """"
											'If aJobComponent(N_POSITION_ID_JOB) = CLng(oRecordset.Fields("PositionID").Value) Then Response.Write " SELECTED=""1"""
											If aJobComponent(N_POSITION_ID_JOB) = CLng(oRecordset.Fields("PositionID").Value) Then Response.Write " SELECTED"
										Response.Write ">" & CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value) & " (Tabulador: " & CStr(oRecordset.Fields("EmployeeTypeName").Value) & ", Compañía: " & CStr(oRecordset.Fields("CompanyName").Value) & ", "
											If CLng(oRecordset.Fields("EmployeeTypeID").Value) = 1 Then
												Response.Write "GGN: " & CStr(oRecordset.Fields("GroupGradeLevelShortName").Value) & ", Clasificación:" & CStr(oRecordset.Fields("ClassificationID").Value) & ", Integración: " & CStr(oRecordset.Fields("IntegrationID").Value)
											Else
												Response.Write "Nivel: " & CStr(oRecordset.Fields("LevelName").Value)
											End If
										Response.Write ", Horas laboradas: " & CStr(oRecordset.Fields("WorkingHours").Value) & ")" & "</OPTION>"
										oRecordset.MoveNext
										If Err.number <> 0 Then Exit Do
									Loop
								End If
							Response.Write "</SELECT>"
						Else
							Response.Write "<TR NAME=""CompanyDiv"" ID=""CompanyDiv"">"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Empresa:&nbsp;</FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
									If aJobComponent(N_STATUS_ID_JOB) = 2 Then
										Response.Write "<SELECT NAME=""CompanyID"" ID=""CompanyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
											Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Companies", "CompanyID", "CompanyShortName, CompanyName", "(ParentID>-1) And (Active=1)", "CompanyShortName", aJobComponent(N_COMPANY_ID_JOB), "Ninguna;;;-1", sErrorDescription)
										Response.Write "</SELECT>"
									Else
										Call GetNameFromTable(oADODBConnection, "Companies", aJobComponent(N_COMPANY_ID_JOB), "", "", sNames, sErrorDescription)
										Response.Write CleanStringForHTML(sNames)
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CompanyID"" ID=""CompanyIDHdn"" VALUE=""" & aJobComponent(N_COMPANY_ID_JOB) & """ />"
									End If
								Response.Write "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
								If aJobComponent(N_STATUS_ID_JOB) = 2 Then
									Response.Write "<SELECT name = ""PositionID"" ID=""PositionIDCmb SIZE=""1"" CLASS=""Lists"">"
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID", "PositionShortName, PositionName","","PositionShortName", aJobComponent(N_POSITION_ID_JOB), "Ninguna;;;-1", sErrorDescription)
									Response.Write "</SELECT>"
								Else
									Call GetNameFromTable(oADODBConnection, "Positions", aJobComponent(N_POSITION_ID_JOB), "", "", sNames, sErrorDescription)
									Response.Write CleanStringForHTML(sNames)
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PositionID"" ID=""PositionIDHdn"" VALUE=""" & aJobComponent(N_POSITION_ID_JOB) & """ />"
								End If
								Response.Write "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR NAME=""WorkingHoursDiv"" ID=""WorkingHoursDiv"">"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Jornada:&nbsp;</FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
									If aJobComponent(N_STATUS_ID_JOB) <> 2 Then
										Response.Write "<INPUT TYPE=""TEXT"" NAME=""WorkingHours"" ID=""WorkingHoursTxt"" VALUE=""" & CStr(aJobComponent(D_WORKING_HOURS_JOB)) + " horas" & """ SIZE=""4"" MAXLENGTH=""4"" CLASS=""SpecialTextFields"" />"
									Else
										Response.Write CStr(aJobComponent(D_WORKING_HOURS_JOB)) & " horas"
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""WorkingHours"" ID=""WorkingHoursHdn"" VALUE=""" & aJobComponent(D_WORKING_HOURS_JOB) & """ />"
									End If
								Response.Write "</FONT></TD>"
							Response.Write "</TR>"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JourneyTypeID From Journeys Where JourneyID = (Select JourneyID From Jobs Where JobID = " & aJobComponent(N_ID_JOB) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de Jornada:&nbsp;</FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
									Response.Write "<INPUT TYPE=""TEXT"" NAME=""JourneyType"" ID=""JourneyTypeTxt"" VALUE=""" & CStr(oRecordset.Fields("JourneyTypeID").Value) & """ SIZE=""4"" MAXLENGTH=""4"" CLASS=""SpecialTextFields"" />"
								Response.Write "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nivel:&nbsp;</FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aJobComponent(N_LEVEL_ID_JOB)
								Response.Write "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
							Response.Write "<TR NAME=""ClassificationIDDiv"" ID=""ClassificationIDDiv"""
								If (CInt(aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)) <> 1) And (CInt(aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)) > -1) Then Response.Write " STYLE=""display: none"""
							Response.Write ">"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clasificación:&nbsp;</FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
									If aJobComponent(N_STATUS_ID_JOB) <> 2 Then
										Response.Write "<INPUT TYPE=""TEXT"" NAME=""ClassificationID"" ID=""ClassificationIDTxt"" SIZE=""2"" MAXLENGTH=""1"" VALUE=""" & aJobComponent(N_CLASSIFICATION_ID_JOB) & """ READONLY=""1"" CLASS=""SpecialTextFields"" />"
									Else
										Response.Write CleanStringForHTML(aJobComponent(N_CLASSIFICATION_ID_JOB))
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ClassificationID"" ID=""ClassificationIDHdn"" VALUE=""" & aJobComponent(N_CLASSIFICATION_ID_JOB) & """ />"
									End If
								Response.Write "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR NAME=""GroupGradeLevelIDDiv"" ID=""GroupGradeLevelIDDiv"""
								If (CInt(aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)) <> 1) And (CInt(aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)) > -1) Then Response.Write " STYLE=""display: none"""
							Response.Write ">"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Grupo, grado, nivel:&nbsp;</FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
									If aJobComponent(N_STATUS_ID_JOB) = 2 Then
										Response.Write "<SELECT NAME=""GroupGradeLevelID"" ID=""GroupGradeLevelIDCmb"" SIZE=""1"" CLASS=""Lists"">"
											Response.Write GenerateListOptionsFromQuery(oADODBConnection, "GroupGradeLevels", "GroupGradeLevelID", "GroupGradeLevelName", "GroupGradeLevelID=" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB), "GroupGradeLevelName", aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB), "Ninguno;;;-1", sErrorDescription)
										Response.Write "</SELECT>"
									Else
										Call GetNameFromTable(oADODBConnection, "GroupGradeLevels", aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB), "", "", sNames, sErrorDescription)
										Response.Write CleanStringForHTML(sNames)
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GroupGradeLevelID"" ID=""GroupGradeLevelIDHdn"" VALUE=""" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & """ />"
									End If
								Response.Write "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR NAME=""IntegrationIDDiv"" ID=""IntegrationIDDiv"""
								If (CInt(aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)) <> 1) And (CInt(aJobComponent(N_EMPLOYEE_TYPE_ID_JOB)) > -1) Then Response.Write " STYLE=""display: none"""
							Response.Write ">"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Integración:&nbsp;</FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
									If aJobComponent(N_STATUS_ID_JOB) <> 2 Then
										Response.Write "<INPUT TYPE=""TEXT"" NAME=""IntegrationID"" ID=""IntegrationIDTxt"" VALUE=""" & aJobComponent(N_INTEGRATION_ID_JOB) & """ SIZE=""2"" MAXLENGTH=""2"" READONLY CLASS=""SpecialTextFields"" />"
									Else
										Response.Write CleanStringForHTML(aJobComponent(N_INTEGRATION_ID_JOB))
										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IntegrationID"" ID=""IntegrationIDHdn"" VALUE=""" & aJobComponent(N_INTEGRATION_ID_JOB) & """ />"
									End If
								Response.Write "</FONT></TD>"
							Response.Write "</TR>"
'							Response.Write "<TR>"
'								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de plaza:&nbsp;</FONT></TD>"
'								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
'									If aJobComponent(N_STATUS_ID_JOB) = 2 Then
'										Response.Write "<SELECT NAME=""JobTypeID"" ID=""JobTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
'											If lReasonID = 59 Then
'												Response.Write GenerateListOptionsFromQuery(oADODBConnection, "JobTypes", "JobTypeID", "JobTypeShortName As RecordID, JobTypeName", "JobTypeID not in(2,4)", "JobTypeID", aJobComponent(N_JOB_TYPE_ID_JOB), "Ninguno;;;-1", sErrorDescription)
'											Else
'												Response.Write GenerateListOptionsFromQuery(oADODBConnection, "JobTypes", "JobTypeID", "JobTypeShortName As RecordID, JobTypeName", "", "JobTypeID", aJobComponent(N_JOB_TYPE_ID_JOB), "Ninguno;;;-1", sErrorDescription)
'											End If
'										Response.Write "</SELECT>"
'									Else
'										Call GetNameFromTable(oADODBConnection, "JobTypes", aJobComponent(N_JOB_TYPE_ID_JOB), "", "", sNames, sErrorDescription)
'										Response.Write CleanStringForHTML(sNames)
'										Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobTypeID"" ID=""JobTypeIDHdn"" VALUE=""" & aJobComponent(N_JOB_TYPE_ID_JOB) & """ />"
'									End If
'								Response.Write "</FONT></TD>"
'							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de Puesto:&nbsp;</FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
								Call GetNameFromTable(oADODBConnection, "PositionTypes", aJobComponent(N_POSITION_TYPE_ID_JOB), "", "", sNames, sErrorDescription)
								Response.Write CleanStringForHTML(sNames)
								Response.Write "</FONT></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de Tabulador:&nbsp;</FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
								Call GetNameFromTable(oADODBConnection, "EmployeeTypes", aJobComponent(N_EMPLOYEE_TYPE_ID_JOB), "", "", sNames, sErrorDescription)
								Response.Write CleanStringForHTML(sNames)
						End If
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de ocupación:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						If (aJobComponent(N_ID_JOB) = -1) Or (aJobComponent(N_STATUS_ID_JOB) = 2) Then
							Response.Write "<SELECT NAME=""OccupationTypeID"" ID=""OccupationTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "OccupationTypes", "OccupationTypeID", "OccupationTypeShortName, OccupationTypeName", "", "OccupationTypeID", aJobComponent(N_OCCUPATION_TYPE_ID_JOB), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
						Else
							Call GetNameFromTable(oADODBConnection, "OccupationTypes", aJobComponent(N_OCCUPATION_TYPE_ID_JOB), "", "", sNames, sErrorDescription)
							Response.Write CleanStringForHTML(sNames)
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OccupationTypeID"" ID=""OccupationTypeIDHdn"" VALUE=""" & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & """ />"
						End If
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de Plaza:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						If (aJobComponent(N_ID_JOB) = -1) Or (aJobComponent(N_STATUS_ID_JOB) = 2) Then
							Response.Write "<SELECT NAME=""JobTypeID"" ID=""JobTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "JobTypes", "JobTypeID", "JobTypeShortName, JobTypeName", "", "JobTypeID", aJobComponent(N_JOB_TYPE_ID_JOB), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
						Else
							Call GetNameFromTable(oADODBConnection, "OccupationTypes", aJobComponent(N_OCCUPATION_TYPE_ID_JOB), "", "", sNames, sErrorDescription)
							Response.Write CleanStringForHTML(sNames)
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OccupationTypeID"" ID=""OccupationTypeIDHdn"" VALUE=""" & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & """ />"
						End If
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Turno:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
					If (aJobComponent(N_ID_JOB) = -1) Or (aJobComponent(N_STATUS_ID_JOB) = 2) Then
						If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
							Response.Write "<SELECT NAME=""JourneyID"" ID=""JourneyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Journeys", "JourneyID", "JourneyShortName, JourneyName", "(Active=1)", "JourneyShortName", aJobComponent(N_JOURNEY_ID_JOB), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT><BR />"
						Else
							Call GetNameFromTable(oADODBConnection, "Journeys", aJobComponent(N_JOURNEY_ID_JOB), "", "", sNames, sErrorDescription)
							Response.Write CleanStringForHTML(sNames)
						End If
					Else
						Call GetNameFromTable(oADODBConnection, "Journeys", aJobComponent(N_JOURNEY_ID_JOB), "", "", sNames, sErrorDescription)
						Response.Write CleanStringForHTML(sNames)
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OccupationTypeID"" ID=""OccupationTypeIDHdn"" VALUE=""" & aJobComponent(N_JOURNEY_ID_JOB) & """ />"
					End If
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
'				Response.Write "<TR>"
'					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horario:&nbsp;</FONT></TD>"
'					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
'						If aJobComponent(N_ID_JOB) = -1 Then
'							Response.Write "<SELECT NAME=""ShiftID"" ID=""ShiftIDCmb"" SIZE=""1"" CLASS=""Lists"">"
'								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Shifts", "ShiftID", "ShiftShortName, ShiftName", "", "ShiftShortName", aJobComponent(N_SHIFT_ID_JOB), "Ninguno;;;-1", sErrorDescription)
'							Response.Write "</SELECT>"
'						Else
'							Call GetNameFromTable(oADODBConnection, "Shifts", aJobComponent(N_SHIFT_ID_JOB), "", "", sNames, sErrorDescription)
'							Response.Write CleanStringForHTML(sNames)
'							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ShiftID"" ID=""ShiftIDHdn"" VALUE=""" & aJobComponent(N_SHIFT_ID_JOB) & """ />"
'						End If
'					Response.Write "</FONT></TD>"
'				Response.Write "</TR>"

				Response.Write "<TR><TD>&nbsp;</TD></TR>"
				Response.Write "<TR><TD COLSPAN=""2""><IMG SRC=""Images/DotBlue.gif"" WIDTH=""640"" HEIGHT=""1"" /></TD></TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estatus:&nbsp;</FONT></TD>"
						If (aJobComponent(N_ID_JOB) = -1) Or (aJobComponent(N_STATUS_ID_JOB) = 2) Then
							Response.Write "<TD><SELECT NAME=""StatusID"" ID=""StatusIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusJobs", "StatusID", "StatusName", "(StatusID=2)", "StatusName", aJobComponent(N_STATUS_ID_JOB), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Else
							Call GetNameFromTable(oADODBConnection, "StatusJobs", aJobComponent(N_STATUS_ID_JOB), "", "", sNames, sErrorDescription)
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write CleanStringForHTML(sNames)
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OccupationTypeID"" ID=""OccupationTypeIDHdn"" VALUE=""" & aJobComponent(N_STATUS_ID_JOB) & """ /></FONT></TD>"
						End If
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Centro de trabajo:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						If (aJobComponent(N_ID_JOB) = -1) Or (aJobComponent(N_STATUS_ID_JOB) = 2) Then
							Response.Write "<SELECT NAME=""AreaID"" ID=""AreaIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""SearchRecord(this.value, 'ZoneForArea', 'SearchZoneForAreaCatalogsIFrame', 'JobFrm.ZoneID')"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(AreaID>-1) And (Active=1)", "AreaCode", aJobComponent(N_PAYMENT_CENTER_ID_JOB), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
						Else
							Call GetNameFromTable(oADODBConnection, "Areas", aJobComponent(N_AREA_ID_JOB), "", "", sNames, sErrorDescription)
							Response.Write CleanStringForHTML(sNames)
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OccupationTypeID"" ID=""OccupationTypeIDHdn"" VALUE=""" & aJobComponent(N_AREA_ID_JOB) & """ /></FONT>"
						End If
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entidad:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						If (aJobComponent(N_ID_JOB) <> -1) Or (aJobComponent(N_STATUS_ID_JOB) <> 2) Then
							Call GetNameFromTable(oADODBConnection, "Zones", aJobComponent(N_ZONE_ID_JOB), "", "", sNames, sErrorDescription)
							Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchZoneForAreaCatalogsIFrame"" FRAMEBORDER=""0"" WIDTH=""600"" HEIGHT=""20""></IFRAME>"
						Else
							Call GetNameFromTable(oADODBConnection, "Zones", aJobComponent(N_ZONE_ID_JOB), "", "", sNames, sErrorDescription)
							Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchZoneForAreaCatalogsIFrame"" FRAMEBORDER=""0"" WIDTH=""600"" HEIGHT=""20""></IFRAME>"
						End If
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ZoneID"" ID=""ZoneID"" VALUE=""" & aJobComponent(N_ZONE_ID_JOB) & """ />"
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD>"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Centro de pago:&nbsp;</FONT>"
					Response.Write "</TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						If (aJobComponent(N_ID_JOB) = -1) Or (aJobComponent(N_STATUS_ID_JOB) = 2) Then
							Response.Write "<SELECT NAME=""PaymentCenterID"" ID=""PaymentCenterIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(AreaID>-1) And (Active=1)", "AreaCode", aJobComponent(N_PAYMENT_CENTER_ID_JOB), "Ninguno;;;-1", sErrorDescription)
								Else
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(AreaID=" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & ") And (Active=1)", "AreaCode", aJobComponent(N_PAYMENT_CENTER_ID_JOB), "Ninguno;;;-1", sErrorDescription)
								End If
							Response.Write "</SELECT><BR />"
						Else
							Call GetNameFromTable(oADODBConnection, "PaymentCenters", aJobComponent(N_PAYMENT_CENTER_ID_JOB), "", "", sNames, sErrorDescription)
							Response.Write CleanStringForHTML(sNames)
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OccupationTypeID"" ID=""OccupationTypeIDHdn"" VALUE=""" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & """ /></FONT></TD>"
						End If
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Servicio:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
					If (aJobComponent(N_ID_JOB) = -1) Or (aJobComponent(N_STATUS_ID_JOB) = 2) Then
						Response.Write "<SELECT NAME=""ServiceID"" ID=""ServiceIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Services", "ServiceID", "ServiceShortName, ServiceName", "(Active=1)", "ServiceShortName", aJobComponent(N_SERVICE_ID_JOB), "Ninguno;;;-1", sErrorDescription)
							Else
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Services", "ServiceID", "ServiceShortName, ServiceName", "(Active=1) And (ServiceID=" & aJobComponent(N_SERVICE_ID_JOB) & ")", "ServiceShortName", aJobComponent(N_SERVICE_ID_JOB), "Ninguno;;;-1", sErrorDescription)
							End If
						Response.Write "</SELECT><BR />"
					Else
						Call GetNameFromTable(oADODBConnection, "Services", aJobComponent(N_SERVICE_ID_JOB), "", "", sNames, sErrorDescription)
						Response.Write CleanStringForHTML(sNames)
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OccupationTypeID"" ID=""OccupationTypeIDHdn"" VALUE=""" & aJobComponent(N_SERVICE_ID_JOB) & """ /></FONT></TD>"
					End If
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select JobDate, EndDate From JobsHistoryList Where JobId =" & aJobComponent(N_ID_JOB) & " And StatusID = 2 Order By JobDate Desc", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oVacancyRecordset)
			If Not oVacancyRecordset.EOF Then
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""VacancyStartDate"" ID=""VacancyStartDateHdn"" VALUE=""" & oVacancyRecordset.Fields("JobDate").Value & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""VacancyEndDate"" ID=""VacancyEndDateHdn"" VALUE=""" & oVacancyRecordset.Fields("EndDate").Value & """ />"
			End If
			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><BR /><B>Vigencia de la plaza</B></FONT>"
			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				If aJobComponent(N_ID_JOB) <> -1 Then
					If aJobComponent(N_START_DATE_JOB) = 0 Then
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Indefinida</FONT></TD>"
						Response.Write "</TR>"
					Else
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aJobComponent(N_START_DATE_JOB), -1, -1, -1) & "</FONT></TD>"
						Response.Write "</TR>"
					End If
				Else
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(CLng(Left(GetSerialNumberForDate(""), Len("00000000"))), "JobStartDate", Year(Date())-1, Year(Date())+2, True, False) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate"" ID=""StartDateHdn"" VALUE=""" & aJobComponent(N_START_DATE_JOB) & """ />"
				If aJobComponent(N_ID_JOB) <> -1 Then
					If aJobComponent(N_END_DATE_JOB) = 30000000 Then
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de término:&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Indefinida</FONT></TD>"
						Response.Write "</TR>"
					Else
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de término:&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aJobComponent(N_END_DATE_JOB), -1, -1, -1) & "</FONT></TD>"
						Response.Write "</TR>"
					End If
				Else
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de término:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(CLng(Left(GetSerialNumberForDate(""), Len("00000000"))), "JobEndDate", Year(Date())-1, Year(Date())+2, True, True) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EndDate"" ID=""EndDateHdn"" VALUE=""" & aJobComponent(N_END_DATE_JOB) & """ />"
			Response.Write "</TABLE>"
			If (CInt(aJobComponent(N_JOB_TYPE_ID_JOB)) <> 4) And (aJobComponent(N_STATUS_ID_JOB) <> 2) And (aJobComponent(N_ID_JOB) <> -1) And (aJobComponent(N_POSITION_TYPE_ID_JOB) <> 2)Then
				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><BR /><B>Titularidad de la plaza</B></FONT>"
				Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"	
					Response.Write "<TR>"
						If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. del empleado que tiene la titularidad:&nbsp;</FONT></TD>"
							Response.Write "<TD>"
								Response.Write "<INPUT TYPE=""TEXT"" NAME=""OwnerID"" ID=""OwnerIDTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""
									If aJobComponent(N_ID_OWNER_JOB) > 0 Then Response.Write aJobComponent(N_ID_OWNER_JOB)
								Response.Write """ CLASS=""TextFields"" />"
								Response.Write "<A HREF=""javascript: SearchRecord(document.JobFrm.OwnerID.value, 'EmployeeNumber', 'SearchEmployeeNumberIFrame', 'JobFrm.OwnerID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A><IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""320"" HEIGHT=""22""></IFRAME>"
							Response.Write "</TD>"
						Else
							If aJobComponent(N_ID_OWNER_JOB) > 0 Then 
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeName, EmployeeLastName, EmployeeLastName2 From Employees Where EmployeeID = " & aJobComponent(N_ID_OWNER_JOB), "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oOwnerRecordset)
								If lErrorNumber = 0 Then
									If Not oOwnerRecordset.EOF Then
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. del empleado que tiene la titularidad:&nbsp;</FONT></TD>"
										Response.Write "<TD>"
											Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & Right("000000" & aJobComponent(N_ID_OWNER_JOB), Len("000000")) & " " & oOwnerRecordset.Fields("EmployeeName").Value & " " & oOwnerRecordset.Fields("EmployeeLastName").Value & " " & oOwnerRecordset.Fields("EmployeeLastName2").Value & "</FONT>"
										Response.Write "</TD>"
									Else
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. del empleado que tiene la titularidad:&nbsp;</FONT></TD>"
										Response.Write "<TD>"
											Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & Right("000000" & aJobComponent(N_ID_OWNER_JOB), Len("000000")) & " Empleado no registrado" & "</FONT>"
										Response.Write "</TD>"
									End If
								End If
							Else
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Esta plaza no tiene titular:&nbsp;</FONT></TD>"
							End If
						End If
					Response.Write "</TR>"
			If  (Len(oRequest("Action").Item) > 0) Then
					Response.Write "<TR>"
						Response.Write "<TD COLLSPAN=2>"
						Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ApplyOwnerBtn"" VALUE=""Aplicar Titularidad"" CLASS=""Buttons"" />"
								Response.Write "</TD>"
					Response.Write "</TR>"
				Response.Write "</TABLE><BR />"

				Response.Write "<FONT FACE=""Arial"" SIZE=""2""><BR /><B>Indique la fecha a partir de que la titularidad será aplicada:&nbsp;</B></FONT>"
				Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
						Response.Write "<TD>" & DisplayDateCombosUsingSerial(aJobComponent(N_JOB_DATE_JOB), "OwnerJob", Year(Date())-40, Year(Date())+1, True, False) & "</TD>"
					Response.Write "</TR>"
			Else
				Response.Write "</TABLE><BR />"
			End If
			End If
			If (aJobComponent(N_ID_JOB) = -1) Or (aJobComponent(N_STATUS_ID_JOB) = 2) Then
				If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 And aJobComponent(N_ID_JOB) > 0 Then
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""><BR /><B>Indique la fecha a partir de que la modificación a la plaza será vigente:&nbsp;</B></FONT>"
					Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
					Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT></TD>"
							Response.Write "<TD>" & DisplayDateCombosUsingSerial(aJobComponent(N_JOB_DATE_JOB), "Job", Year(Date())-1, Year(Date())+2, True, False) & "</TD>"
						Response.Write "</TR>"
					If (aJobComponent(N_STATUS_ID_JOB) = 2) Or (aJobComponent(N_STATUS_ID_JOB) = 4) Or (aJobComponent(N_STATUS_ID_JOB) = 5) Or (aJobComponent(N_STATUS_ID_JOB) = 7) Then
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de término:&nbsp;</FONT></TD>"
							If aJobComponent(N_END_DATE_HISTORY_JOB) = 30000000 Then
								aJobComponent(N_END_DATE_HISTORY_JOB) = AddDaysToSerialDate(aJobComponent(N_JOB_DATE_JOB), 15)
								Response.Write "<TD>" & DisplayDateCombosUsingSerial(aJobComponent(N_END_DATE_HISTORY_JOB), "JobEnd", Year(Date())-1, Year(Date())+1, True, True) & "</TD>"
							Else
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aJobComponent(N_END_DATE_HISTORY_JOB), "JobEnd", Year(Date())-1, Year(Date())+2, True, True) & "</FONT></TD>"
							End If
						Response.Write "</TR>"
					Else
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de término:&nbsp;</FONT></TD>"
							If aJobComponent(N_END_DATE_HISTORY_JOB) = 30000000 Then
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Indefinida</FONT></TD>"
							Else
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aJobComponent(N_END_DATE_HISTORY_JOB), -1, -1, -1) & "</FONT></TD>"
							End If
						Response.Write "</TR>"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobEndDate"" ID=""JobEndDateHdn"" VALUE=""" & aJobComponent(N_END_DATE_HISTORY_JOB) & """ />"
					End If
				Else
					Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				End If
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Activo:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" ID=""ActiveRd"" VALUE=""1"""
								If aJobComponent(N_ACTIVE_JOB) = 1 Then Response.Write " CHECKED=""1"""
							Response.Write " />Sí&nbsp;&nbsp;&nbsp;"
							Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" ID=""ActiveRd"" VALUE=""0"""
								If aJobComponent(N_ACTIVE_JOB) = 0 Then Response.Write " CHECKED=""0"""
							Response.Write " />No&nbsp;&nbsp;&nbsp;"
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
				Response.Write "</TABLE>"
				Response.Write "<BR /><BR />"
			End If

			'Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""480"" HEIGHT=""1"" /><BR /><BR />"
			'Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Partidas presupuestales:<BR /></FONT>"
			'Response.Write "&nbsp;&nbsp;&nbsp;<SELECT NAME=""BudgetID"" ID=""BudgetIDLst"" SIZE=""10"" MULTIPLE=""1"" CLASS=""Lists"">"
			'	Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Budgets As Partida, Budgets As SubPartida, Budgets As Tipo", "Tipo.BudgetID", "'Partida: ' As Temp1, Partida.BudgetName, 'SubPartida: ' As Temp2, SubPartida.BudgetName, 'Tipo: ' As Temp3, Tipo.BudgetName", "(Partida.BudgetID=SubPartida.ParentID) And (SubPartida.BudgetID=Tipo.ParentID) And (SubPartida.ParentID>-1) And (Tipo.BudgetID>-1) And (Tipo.Active=1)", "Partida.BudgetName, SubPartida.BudgetName, Tipo.BudgetName", aJobComponent(S_BUDGETS_ID_JOB), "Ninguna;;;-1", sErrorDescription)
			'Response.Write "</SELECT><BR />"
			'Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""480"" HEIGHT=""1"" /><BR /><BR />"
			If (aJobComponent(N_ID_JOB) = -1) Or (aJobComponent(N_STATUS_ID_JOB) = 2) Then
			If StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) <> 0 Then
					If aJobComponent(N_ID_JOB) = -1 Then
						If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS) = N_ADD_PERMISSIONS) And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_ModificacionDePlazas & ",", vbBinaryCompare) > 0) Then
							Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
							Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "'"" />"
						End If
					ElseIf Len(oRequest("Delete").Item) > 0 Then
						If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveJobWngDiv']); JobFrm.Remove.focus()"" />"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
							Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "'"" />"
						End If
					Else
						If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) And (StrComp("0", aLoginComponent(N_PROFILE_ID_LOGIN), vbBinaryCompare) = 0 Or InStr(1, "," & aLoginComponent(N_USER_PERMISSIONS2_LOGIN) & ",", "," & N_03_ModificacionDePlazas & ",", vbBinaryCompare) > 0) Then
							Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
							Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "'"" />"
						End If
					End If
			End If
			End If
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			If aJobComponent(N_ID_JOB) = -1 Then
				Response.Write "ShowPopupItem('WaitDiv', parent.window.document.all['WaitDiv'], false);" & vbNewLine
				Response.Write "SearchRecord(document.JobFrm.PositionID.value, 'PositionsCatalogsLKP', 'SearchPositionsCatalogsIFrame', 'JobFrm');" & vbNewLine
			Else
				Response.Write "SearchRecord(document.JobFrm.AreaID.value, 'ZoneForArea', 'SearchZoneForAreaCatalogsIFrame', 'JobFrm.ZoneID');" & vbNewLine
			End If
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayJobForm = lErrorNumber
	Err.Clear
End Function

Function DisplayJobAsHiddenFields(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a job using
'		  hidden form fields
'Inputs:  oRequest, oADODBConnection, aJobComponent
'Outputs: aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayJobAsHiddenFields"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobID"" ID=""JobIDHdn"" VALUE=""" & aJobComponent(N_ID_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OwnerID"" ID=""OwnerIDHdn"" VALUE=""" & aJobComponent(N_ID_OWNER_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobNumber"" ID=""JobNumberHdn"" VALUE=""" & aJobComponent(S_NUMBER_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ZoneID"" ID=""ZoneIDHdn"" VALUE=""" & aJobComponent(N_ZONE_ID_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaID"" ID=""AreaIDHdn"" VALUE=""" & aJobComponent(N_AREA_ID_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaymentCenterID"" ID=""PaymentCenterIDHdn"" VALUE=""" & aJobComponent(N_PAYMENT_CENTER_ID_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PositionID"" ID=""PositionIDHdn"" VALUE=""" & aJobComponent(N_POSITION_ID_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobTypeID"" ID=""JobTypeIDHdn"" VALUE=""" & aJobComponent(N_JOB_TYPE_ID_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ShiftID"" ID=""ShiftIDHdn"" VALUE=""" & aJobComponent(N_SHIFT_ID_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JourneyID"" ID=""JourneyIDHdn"" VALUE=""" & aJobComponent(N_JOURNEY_ID_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ClassificationID"" ID=""ClassificationIDHdn"" VALUE=""" & aJobComponent(N_CLASSIFICATION_ID_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GroupGradeLevelID"" ID=""GroupGradeLevelIDHdn"" VALUE=""" & aJobComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IntegrationID"" ID=""IntegrationIDHdn"" VALUE=""" & aJobComponent(N_INTEGRATION_ID_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OccupationTypeID"" ID=""OccupationTypeIDHdn"" VALUE=""" & aJobComponent(N_OCCUPATION_TYPE_ID_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ServiceID"" ID=""ServiceIDHdn"" VALUE=""" & aJobComponent(N_SERVICE_ID_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""LevelID"" ID=""LevelIDHdn"" VALUE=""" & aJobComponent(N_LEVEL_ID_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""WorkingHours"" ID=""WorkingHoursHdn"" VALUE=""" & aJobComponent(N_LEVEL_ID_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate"" ID=""StartDateHdn"" VALUE=""" & aJobComponent(N_START_DATE_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EndDate"" ID=""EndDateHdn"" VALUE=""" & aJobComponent(N_END_DATE_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobStartDate"" ID=""JobStartDateHdn"" VALUE=""" & aJobComponent(N_JOB_DATE_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobEndDate"" ID=""JobEndDateHdn"" VALUE=""" & aJobComponent(N_END_DATE_HISTORY_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StatusID"" ID=""StatusIDHdn"" VALUE=""" & aJobComponent(N_STATUS_ID_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Active"" ID=""ActiveHdn"" VALUE=""" & aJobComponent(N_ACTIVE_JOB) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetsID"" ID=""BudgetsIDHdn"" VALUE=""" & aJobComponent(S_BUDGETS_ID_JOB) & """ />"

	DisplayJobAsHiddenFields = Err.number
	Err.Clear
End Function

Function DisplayFreeJobsTable(oRequest, oADODBConnection, lIDColumn, bForExport, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the free jobs
'		  from the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bForExport, aJobComponent
'Outputs: aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFreeJobsTable"
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

	aJobComponent(S_QUERY_CONDITION_JOB) = " And (Jobs.StatusID In (0,5,6))"
	lErrorNumber = GetJobs(oRequest, oADODBConnection, aJobComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""700"" BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("&nbsp;,No.,Área,Entidad,Puesto,Estatus", ",", -1, vbBinaryCompare)
				asCellWidths = Split("20,120,120,120,120,100", ",", -1, vbBinaryCompare)
				asCellAlignments = Split(",,,,,", ",", -1, vbBinaryCompare)

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
					sFontBegin = ""
					sFontEnd = ""
					If CInt(oRecordset.Fields("Active").Value) = 0 Then
						sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
						sFontEnd = "</FONT>"
					End If
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("JobID").Value), oRequest("JobID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""JobID"" ID=""JobIDRd"" VALUE=""" & CStr(oRecordset.Fields("JobID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""JobID"" ID=""JobIDChk"" VALUE=""" & CStr(oRecordset.Fields("JobID").Value) & """ />"
						Case Else
							sRowContents = sRowContents & "&nbsp;"
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("JobNumber").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value)) & sBoldEnd & sFontEnd

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
			Response.Write "<FONT FACE=""Arial"" SIZE=""2"" COLOR=""#" & S_WARNING_FOR_GUI & """><B>No existen plazas vacantes.</B></FONT>"
			If ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_JOBS_PERMISSIONS) = N_JOBS_PERMISSIONS) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS) = N_ADD_PERMISSIONS) Then Response.Write "<FONT FACE=""Arial"" SIZE=""2"" COLOR=""#" & S_WARNING_FOR_GUI & """><BR /><BR /><A HREF=""Jobs.asp?New=1"">¿Desea dar de alta una plaza?</A></FONT>"
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayFreeJobsTable = lErrorNumber
	Err.Clear
End Function

Function DisplayJobsTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, bForExport, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the jobs from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, bForExport, aJobComponent
'Outputs: aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayJobsTable"
	Dim oRecordset
	Dim oJobsRecordset
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
	Dim iJobs
	Dim iTotalJobs
	Dim iCurrentJobs
	Dim iIndex
	Dim sPage
	Dim sTarget
	Dim lErrorNumber

	If bForExport Then bUseLinks = False
	sTarget = ""
	sPage = GetASPFileName("")
	If aJobComponent(B_SEND_TO_IFRAME_JOB) Then
		sTarget = " TARGET=""FormsIFrame"""
		sPage = "ShowForms.asp"
	End If
	iTotalJobs = 0
	If False Then 'aJobComponent(N_SHOW_BY_JOB) = N_SHOW_BY_AREA Then
		sErrorDescription = "No se pudo obtener la información de las plazas."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(JobsInArea) As JobsSum From Areas Where (AreaID>-1)" & Replace(aJobComponent(S_QUERY_CONDITION_JOB), "Jobs.", "Areas."), "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then iTotalJobs = CLng(oRecordset.Fields("JobsSum").Value)
	End If

	lErrorNumber = GetJobs(oRequest, oADODBConnection, aJobComponent, oRecordset, sErrorDescription)
	If (aJobComponent(N_SHOW_BY_JOB) = 0) And (oRecordset.EOF) Then
		lErrorNumber = L_ERR_NO_RECORDS
		sErrorDescription = "No existen plazas registradas en la base de datos."
	End If

	If lErrorNumber = 0 Then
		Response.Write "<TABLE WIDTH=""700"" BORDER="""
			If bForExport Then
				Response.Write "1"
			Else
				Response.Write "0"
			End If
		Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			If aJobComponent(N_SHOW_BY_JOB) = N_SHOW_BY_AREA Then
				If bUseLinks Then
					asColumnsTitles = Split("&nbsp;,No.,Entidad,Puesto,Estatus,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,100,150,250,100,80", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("&nbsp;,No.,Entidad,Puesto,Estatus", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,120,170,270,120", ",", -1, vbBinaryCompare)
				End If
				asCellAlignments = Split(",,,,,CENTER", ",", -1, vbBinaryCompare)
			ElseIf aJobComponent(N_SHOW_BY_JOB) = N_SHOW_BY_POSITION Then
				If bUseLinks Then
					asColumnsTitles = Split("&nbsp;,No.,Área,Entidad,Estatus,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,100,150,250,100,80", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("&nbsp;,No.,Área,Entidad,Estatus", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,120,170,270,120", ",", -1, vbBinaryCompare)
				End If
				asCellAlignments = Split(",,,,,CENTER", ",", -1, vbBinaryCompare)
			Else
				If bUseLinks Then
					asColumnsTitles = Split("&nbsp;,Plaza,Nivel,Adscripción,Entidad,Puesto,Estatus,Fecha inicio,Fecha Fin,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,100,100,100,100,100,100,100,100,80", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("&nbsp;,No. de plaza,Área,Entidad,Puesto,Estatus", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,120,120,120,120,100", ",", -1, vbBinaryCompare)
				End If
				asCellAlignments = Split(",CENTER,CENTER,,,,CENTER,CENTER,CENTER,,CENTER", ",", -1, vbBinaryCompare)
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

			iJobs = 1
			If Not oRecordset.EOF Then
				Do While Not oRecordset.EOF
					sFontBegin = ""
					sFontEnd = ""
					If CInt(oRecordset.Fields("Active").Value) = 0 Then
						sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
						sFontEnd = "</FONT>"
					End If
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("JobID").Value), oRequest("JobID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""JobID"" ID=""JobIDRd"" VALUE=""" & CStr(oRecordset.Fields("JobID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""JobID"" ID=""JobIDChk"" VALUE=""" & CStr(oRecordset.Fields("JobID").Value) & """ />"
						Case Else
							sRowContents = sRowContents & "&nbsp;"
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("JobNumber").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("LevelName").Value)) & sBoldEnd & sFontEnd
					If aJobComponent(N_SHOW_BY_JOB) <> N_SHOW_BY_AREA Then sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName").Value)) & sBoldEnd & sFontEnd
					If aJobComponent(N_SHOW_BY_JOB) <> N_SHOW_BY_POSITION Then sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(getDateFromSerialNumber(oRecordset.Fields("JobDate").Value)) & sBoldEnd & sFontEnd
					If CStr(oRecordset.Fields("EndDate").Value) = "30000000" Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "Indefinida" & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & getDateFromSerialNumber(CStr(oRecordset.Fields("EndDate").Value)) & sBoldEnd & sFontEnd
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ReasonName").Value)) & sBoldEnd & sFontEnd
					If bUseLinks Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							sRowContents = sRowContents & "<A HREF=""" & sPage & "?Action=Jobs&JobID=" & CStr(oRecordset.Fields("JobID").Value) & "&JobNumber=" & CStr(oRecordset.Fields("JobID").Value) & "&Tab=1&Change=1""" & sTarget & ">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If

						'If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						'	sRowContents = sRowContents & "<A HREF=""" & sPage & "?Action=Jobs&JobID=" & CStr(oRecordset.Fields("JobID").Value) & "&Tab=1&Delete=1""" & sTarget & ">"
						'		sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
						'	sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						'End If

						'If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
						'	If CInt(oRecordset.Fields("Active").Value) = 0 Then
						'		sRowContents = sRowContents & "<A HREF=""" & sPage & "?Action=Jobs&JobID=" & CStr(oRecordset.Fields("JobID").Value) & "&SetActive=1""" & sTarget & "><IMG SRC=""Images/BtnActive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Activar"" BORDER=""0"" /></A>"
						'	Else
						'		sRowContents = sRowContents & "<A HREF=""" & sPage & "?Action=Jobs&JobID=" & CStr(oRecordset.Fields("JobID").Value) & "&SetActive=0""" & sTarget & "><IMG SRC=""Images/BtnDeactive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Desactivar"" BORDER=""0"" /></A>"
						'	End If
						'End If
						sRowContents = sRowContents & "&nbsp;"

					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					iJobs = iJobs + 1
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			End If

			If False Then 'aJobComponent(N_SHOW_BY_JOB) <> 0 Then
				lErrorNumber = GetJobsNotAdded(oRequest, oADODBConnection, aJobComponent, oRecordset, sErrorDescription)
				If lErrorNumber = 0 Then
					iJobs = 1
					If Not oRecordset.EOF Then
						Do While Not oRecordset.EOF
							sRowContents = "&nbsp;"
							sRowContents = sRowContents & TABLE_SEPARATOR & "---"
							If aJobComponent(N_SHOW_BY_JOB) <> N_SHOW_BY_AREA Then sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & "---"
							If aJobComponent(N_SHOW_BY_JOB) <> N_SHOW_BY_POSITION Then sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value) & ". " & CStr(oRecordset.Fields("PositionName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & "---"
							If bUseLinks Then
								sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
								If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS) = N_ADD_PERMISSIONS Then
									sRowContents = sRowContents & "<A HREF=""" & sPage & "?Action=Jobs&JobID=-1&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & "&PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & "&Tab=1&New=1""" & sTarget & ">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;"
								End If
							End If

							iCurrentJobs = 0
							sErrorDescription = "No se pudo obtener la información de las plazas."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Count(JobID) As TotalJobsForPosition From Jobs Where (AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & ") And (PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & ")", "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oJobsRecordset)
							If lErrorNumber = 0 Then
								If Not oJobsRecordset.EOF Then
									If Not IsNull(oJobsRecordset.Fields("TotalJobsForPosition").Value) Then iCurrentJobs = CLng(oJobsRecordset.Fields("TotalJobsForPosition").Value)
								End If
								oJobsRecordset.Close
							End If

							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							For iIndex = 1 To CLng(oRecordset.Fields("JobsInArea").Value) - iCurrentJobs
								If bForExport Then
									lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
								Else
									lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
								End If
							Next
							iJobs = iJobs + CLng(oRecordset.Fields("JobsInArea").Value)
							oRecordset.MoveNext
							If Err.number <> 0 Then Exit Do
						Loop
					End If
				End If
			End If

			sRowContents = TABLE_SEPARATOR & "---"
			If aJobComponent(N_SHOW_BY_JOB) <> N_SHOW_BY_AREA Then sRowContents = sRowContents & TABLE_SEPARATOR & "---"
			sRowContents = sRowContents & TABLE_SEPARATOR & "---"
			If aJobComponent(N_SHOW_BY_JOB) <> N_SHOW_BY_POSITION Then sRowContents = sRowContents & TABLE_SEPARATOR & "---"
			sRowContents = sRowContents & TABLE_SEPARATOR & "---"
			If bUseLinks Then
				sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
				If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS) = N_ADD_PERMISSIONS Then
					sRowContents = sRowContents & "<A HREF=""" & sPage & "?Action=Jobs&JobID=-1&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & "&Tab=1&New=1""" & sTarget & ">"
						sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
					sRowContents = sRowContents & "</A>&nbsp;"
				End If
			End If
			asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
			For iJobs = iJobs To iTotalJobs
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Next
		Response.Write "</TABLE>" & vbNewLine
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	Set oJobsRecordset = Nothing
	DisplayJobsTable = lErrorNumber
	Err.Clear
End Function

Function DisplayJobsHistoryListTable(oRequest, oADODBConnection, bForExport, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To display the employees that are in process of movement
'
'Inputs:  oRequest, oADODBConnection, bForExport, aJobComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayJobsHistoryListTable"
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
	Dim sCondition
	Dim sFields
	Dim sStatusEmployeesIDs
	Dim sDate
	Dim sQuery

	sDate = CLng(Left(GetSerialNumberForDate(""), (Len("00000000"))))
	'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusName, Employees.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, JobsHistoryList.JobID, JobDate, JobsHistoryList.EndDate, ReasonName From JobsHistoryList, Employees, StatusJobs, EmployeesHistoryList, Reasons Where (JobsHistoryList.JobID=" & aJobComponent(N_ID_JOB) & ") And (JobsHistoryList.EmployeeID = Employees.EmployeeID) And (JobsHistoryList.StatusID = StatusJobs.StatusID) And (JobsHistoryList.EmployeeID = EmployeesHistoryList.EmployeeID) And (Employees.EmployeeID = EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.JobID = JobsHistoryList.JobID) And (EmployeesHistoryList.ReasonID = Reasons.ReasonID) And (JobsHistoryList.JobDate = EmployeesHistoryList.EmployeeDate) And (JobsHistoryList.EndDate = EmployeesHistoryList.EndDate) Order by JobDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StatusName, Employees.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, JobsHistoryList.JobID, JobDate, JobsHistoryList.EndDate From JobsHistoryList, Employees, StatusJobs Where (JobsHistoryList.JobID=" & aJobComponent(N_ID_JOB) & ") And (JobsHistoryList.EmployeeID = Employees.EmployeeID) And (JobsHistoryList.StatusID = StatusJobs.StatusID) And (JobsHistoryList.EndDate <> 0) Order by JobDate Desc", "EmployeeComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
			If Not bForExport Then
				Response.Write "0"
			Else
				Response.Write "1"
			End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("Fecha inicio,Fecha fin,Estatus,No. Emp.,Nombre, Movimiento", ",", -1, vbBinaryCompare)
			asCellWidths = Split(",,,,,", ",", -1, vbBinaryCompare)
			If bForExport Then
				lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			Else
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If
			End If
			asCellAlignments = Split(",,,,", ",", -1, vbBinaryCompare)
			Do While Not oRecordset.EOF
				sFontBegin = ""
				sFontEnd = ""
				sBoldBegin = ""
				sBoldEnd = ""
				If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
					sBoldBegin = "<B>"
					sBoldEnd = "</B>"
				End If
				If CLng(oRecordset.Fields("JobDate").Value) = 0 Then
					sRowContents = sFontBegin & sBoldBegin & "-" & sBoldEnd & sFontEnd
				Else
					sRowContents = sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("JobDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
				End If
				If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "Indefinida" & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value)) & sBoldEnd & sFontEnd
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & Right("000000" & CStr(oRecordset.Fields("EmployeeNumber").Value), Len("000000")) & sBoldEnd & sFontEnd
				If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value)) & sBoldEnd & sFontEnd
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value)) & " " & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value)) & sBoldEnd & sFontEnd
				End If
				sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ReasonName").Value)) & sBoldEnd & sFontEnd
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
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "Esta plaza no tiene historial de ocupación."
		End If
	End If
	
	Set oRecordset = Nothing
	DisplayJobsHistoryListTable = lErrorNumber
	Err.Clear
End Function

Function DisplayPendingJobsTable(oRequest, oADODBConnection, bForExport, sAction, lReasonID, iActive, aJobComponent, sErrorDescription)
'************************************************************
'Purpose: To display the jobs that are in process of movement
'
'Inputs:  oRequest, oADODBConnection, bForExport, sAction, lReasonID
'Outputs: aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPendingJobsTable"
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
	Dim sCondition
	Dim sFields
	Dim sStatusEmployeesIDs
	Dim sDate
	Dim sQuery

	sDate = CLng(Left(GetSerialNumberForDate(""), (Len("00000000"))))

	If lReasonID = 60 Then
		sQuery = "Select JobsHistoryList.*, Areas.AreaCode, Areas.AreaName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, PositionShortName, PositionName, PositionTypeName, JobTypeName, JourneyName, ShiftName, ServiceShortName, ServiceName, StatusName, UserName, UserLastName From Areas, JobsHistoryList, JobTypes, Journeys, Areas As PaymentCenters, Positions, PositionTypes, StatusJobs, Services, Shifts, Users Where (JobsHistoryList.AreaID=Areas.AreaID) And (JobsHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (JobsHistoryList.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (JobsHistoryList.JobTypeID=JobTypes.JobTypeID) And (JobsHistoryList.StatusID=StatusJobs.StatusID) And (JobsHistoryList.UserID=Users.UserID) And (JobsHistoryList.ServiceID=Services.ServiceID) And (JobsHistoryList.ShiftID=Shifts.ShiftID) And (JobsHistoryList.JourneyID=Journeys.JourneyID) And (JobsHistoryList.ModifyDate=" & sDate & ")" & " Order By JobsHistoryList.JobID, JobsHistoryList.JobDate"
	Else
		sQuery = "Select Jobs.*, Areas.AreaCode, Areas.AreaName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, PositionShortName, PositionName, PositionTypeName, JobTypeName, JourneyName, ShiftName, ServiceShortName, ServiceName, StatusName, UserName, UserLastName From Areas, Jobs, JobsHistoryList, JobTypes, Journeys, Areas As PaymentCenters, Positions, PositionTypes, StatusJobs, Services, Shifts, Users Where (Jobs.AreaID=Areas.AreaID) And (Jobs.PaymentCenterID=PaymentCenters.AreaID) And (Jobs.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (Jobs.JobTypeID=JobTypes.JobTypeID) And (Jobs.StatusID=StatusJobs.StatusID) And (JobsHistoryList.JobID=Jobs.JobID) And (JobsHistoryList.UserID=Users.UserID) And (Jobs.ServiceID=Services.ServiceID) And (Jobs.ShiftID=Shifts.ShiftID) And (Jobs.JourneyID=Journeys.JourneyID) And (Jobs.Active=0) Order By Jobs.StartDate"
	End If

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & sQuery & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReasonID"" ID=""ReasonIDHdn"" VALUE="&lReasonID&" />"

	If lErrorNumber = 0 Then
		sErrorDescription = "No existe información."
		If Not oRecordset.EOF Then
			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bForExport Or iActive = 1 Then
					asColumnsTitles = Split("Número de plaza,Fecha inicio,Fecha fin,Clave centro de trabajo,Centro de trabajo,Clave centro de pago,Centro de pago,Clave del servicio,Descripción del Servicio,Clave del puesto,Descripción del Puesto,Tipo de plaza,Tipo de ocupación,Turno,Jornada,Horario,Estatus,Usuario", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("Número de plaza,Fecha inicio,Fecha fin,Clave centro de trabajo,Centro de trabajo,Clave centro de pago,Centro de pago,Clave del servicio,Descripción del Servicio,Clave del puesto,Descripción del Puesto,Tipo de plaza,Tipo de ocupación,Turno,Jornada,Horario,Estatus,Usuario,Acciones", ",", -1, vbBinaryCompare)
				End If
				asCellWidths = Split(",,,,,,,,,,,,,,,,,,", ",", -1, vbBinaryCompare)
				asCellAlignments = Split("CENTER,,,CENTER,,CENTER,,CENTER,,CENTER,,CENTER,CENTER,,CENTER,,CENTER,,CENTER", ",", -1, vbBinaryCompare)
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
					sRowContents = Right("000000" & CleanStringForHTML(CStr(oRecordset.Fields("JobID").Value)), 6)
					If lReasonID = 60 Then
						If CLng(oRecordset.Fields("JobDate").Value) = 0 Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "Indefinida"
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("JobDate").Value), -1, -1, -1)
						End If
					Else
						If CLng(oRecordset.Fields("StartDate").Value) = 0 Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "Indefinida"
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
						End If
					End If
					If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "Indefinida"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ServiceShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ServiceName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("JobTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("JourneyName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("WorkingHours").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ShiftName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("UserName").Value) & " " & CStr(oRecordset.Fields("UserLastName").Value))

					If Not bForExport And (lReasonID <> 60)Then
						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "Jobs.asp" & "?Action=" & sAction & "&Remove=1&Pending=1&JobID=" & CStr(oRecordset.Fields("JobID").Value) & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Cancelar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "Jobs.asp" & "?Action=" & sAction & "&Pending=1&JobID=" & CStr(oRecordset.Fields("JobID").Value) & "&SetActive=1 "">"
								sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""10"" ALT=""Aplicar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & "&nbsp;&nbsp;<A HREF=""" & "Jobs.asp" & "?Action=" & sAction & "&Tab=1&Change=1&JobID=" & CStr(oRecordset.Fields("JobID").Value) & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""" & CStr(oRecordset.Fields("JobID").Value) & """ ID=""" & CStr(oRecordset.Fields("JobID").Value) & "Chk"" Value=""" & CStr(oRecordset.Fields("JobID").Value) & """/>"
						End If
					End If
					If lReasonID = 60 Then
						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;&nbsp;<A HREF=""" & "Jobs.asp" & "?Action=" & sAction & "&Tab=1&Change=1&JobID=" & CStr(oRecordset.Fields("JobID").Value) & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
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
			Response.Write "</TABLE></DIV>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
		End If
	End If

	Set oRecordset = Nothing
	DisplayPendingJobsTable = lErrorNumber
	Err.Clear
End Function

Function GetJobLastEmployeeMovement(oADODBConnection, iJobID, sMovement, sErrorDescription)
'************************************************************
'Purpose: To get the last movement of the last employee
'		  assigned at the job	
'Inputs:  oADODBConnection, iJobID
'Outputs: sMovement, sErrorDescription
'************************************************************
	On Error Resume Next

	Const S_FUNCTION_NAME = "GetJobLastEmployeeMovement"
	Dim lErrorNumber
	Dim lEmployeeID
	Dim sQuery

	sQuery = "Select EmployeeID From Employees Where (JobID=" & iJobID & ")"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "JobComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordSet)
	sQuery = "Select ReasonName From Reasons Where ReasonID = (Select ReasonID From EmployeesHistoryList Where EmployeeID = (Select EmployeeID From Employees Where JobID = " & iJobID & ") And EmployeeDate = (Select MAX(EmployeeDate) As EmployeeDate From EmployeesHistoryList Where EmployeeID = " & lEmployeeID & "))"

	GetJobLastEmployeeMovement = lErrorNumber
	Err.Clear
End Function
%>