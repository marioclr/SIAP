<%
Const N_ID_AREA = 0
Const N_PARENT_ID_AREA = 1
Const S_CODE_AREA = 2
Const S_SHORT_NAME_AREA = 3
Const S_NAME_AREA = 4
Const S_PATH_AREA = 5
Const S_URCTAUX_AREA = 6
Const N_COMPANY_ID_AREA = 7
Const N_TYPE_ID_AREA = 8
Const N_CONFINE_TYPE_ID_AREA = 9
Const N_LEVEL_TYPE_ID_AREA = 10
Const N_CENTER_TYPE_ID_AREA = 11
Const N_CENTER_SUBTYPE_ID_AREA = 12
Const N_ATTENTION_LEVEL_ID_AREA = 13
Const S_ADDRESS_AREA = 14
Const S_CITY_AREA = 15
Const S_ZIP_CODE_AREA = 16
Const N_ZONE_ID_AREA = 17
Const N_ECONOMIC_ZONE_ID_AREA = 18
Const N_PAYMENT_CENTER_ID_AREA = 19
Const N_GENERATING_AREA_ID_AREA = 20
Const N_BRANCH_ID_AREA = 21
Const N_SUBBRANCH_ID_AREA = 22
Const N_CASHIER_OFFICE_ID_AREA = 23
Const N_START_DATE_AREA = 24
Const N_END_DATE_AREA = 25
Const N_FINISH_DATE_AREA = 26
Const N_JOBS_AREA = 27
Const N_TOTAL_JOBS_AREA = 28
Const N_STATUS_ID_AREA = 29
Const N_ACTIVE_AREA = 30
Const N_ONLY_PAYMENT = 31
Const N_PARENT_ZONE = 32
Const N_PREVIOUS_PARENT_ID_AREA = 33
Const N_PREVIOUS_JOBS_AREA = 34
Const N_PREVIOUS_TOTAL_JOBS_AREA = 35
Const N_PREVIOUS_STATUS_ID_AREA = 36
Const S_POSITIONS_ID_AREA = 37
Const S_JOBS_FOR_POSITIONS_AREA = 38
Const B_SEND_TO_IFRAME_AREA = 39
Const S_QUERY_CONDITION_AREA = 40
Const B_CHECK_FOR_DUPLICATED_AREA = 41
Const B_IS_DUPLICATED_AREA = 42
Const B_COMPONENT_INITIALIZED_AREA = 43
Const N_LEVEL_AREA = 44
Const N_PARENT_ID2_AREA = 45
Const S_FILTER_CONDITION_AREA = 46
Const N_AREA_COMPONENT_SIZE = 46

Dim aAreaComponent()
Redim aAreaComponent(N_AREA_COMPONENT_SIZE)

Function InitializeAreaComponent(oRequest, aAreaComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Area
'         Component using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aAreaComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeAreaComponent"
	Redim Preserve aAreaComponent(N_AREA_COMPONENT_SIZE)
	Dim oItem

	If IsEmpty(aAreaComponent(N_ID_AREA)) Then
		If Len(oRequest("AreaID").Item) > 0 Then
			aAreaComponent(N_ID_AREA) = CLng(oRequest("AreaID").Item)
		Else
			aAreaComponent(N_ID_AREA) = -1
		End If
	End If

	If IsEmpty(aAreaComponent(N_PARENT_ID_AREA)) Then
		If Len(oRequest("ParentID").Item) > 0 Then
			aAreaComponent(N_PARENT_ID_AREA) = CLng(oRequest("ParentID").Item)
		Else
			aAreaComponent(N_PARENT_ID_AREA) = -1
		End If
	End If

	If IsEmpty(aAreaComponent(S_CODE_AREA)) Then
		If Len(oRequest("AreaCode").Item) > 0 Then
			aAreaComponent(S_CODE_AREA) = oRequest("AreaCode").Item
		Else
			aAreaComponent(S_CODE_AREA) = ""
		End If
	End If
	aAreaComponent(S_CODE_AREA) = Left(aAreaComponent(S_CODE_AREA), 5)

	If IsEmpty(aAreaComponent(S_SHORT_NAME_AREA)) Then
		If Len(oRequest("AreaShortName").Item) > 0 Then
			aAreaComponent(S_SHORT_NAME_AREA) = oRequest("AreaShortName").Item
		Else
			aAreaComponent(S_SHORT_NAME_AREA) = aAreaComponent(S_CODE_AREA)
		End If
	End If
	aAreaComponent(S_SHORT_NAME_AREA) = Left(aAreaComponent(S_SHORT_NAME_AREA), 10)

	If IsEmpty(aAreaComponent(S_NAME_AREA)) Then
		If Len(oRequest("AreaName").Item) > 0 Then
			aAreaComponent(S_NAME_AREA) = oRequest("AreaName").Item
		Else
			aAreaComponent(S_NAME_AREA) = ""
		End If
	End If
	aAreaComponent(S_NAME_AREA) = Left(aAreaComponent(S_NAME_AREA), 255)

	If IsEmpty(aAreaComponent(S_PATH_AREA)) Then
		If Len(oRequest("AreaPath").Item) > 0 Then
			aAreaComponent(S_PATH_AREA) = oRequest("AreaPath").Item
		Else
			aAreaComponent(S_PATH_AREA) = ",-1,"
			If aAreaComponent(N_ID_AREA) > -1 Then aAreaComponent(S_PATH_AREA) = aAreaComponent(S_PATH_AREA) & aAreaComponent(N_ID_AREA) & ","
		End If
	End If
	aAreaComponent(S_PATH_AREA) = Left(aAreaComponent(S_PATH_AREA), 255)

	If IsEmpty(aAreaComponent(S_URCTAUX_AREA)) Then
		If Len(oRequest("URCTAUX").Item) > 0 Then
			aAreaComponent(S_URCTAUX_AREA) = oRequest("URCTAUX").Item
		Else
			aAreaComponent(S_URCTAUX_AREA) = ""
		End If
	End If
	aAreaComponent(S_URCTAUX_AREA) = Left(aAreaComponent(S_URCTAUX_AREA), 10)

	If IsEmpty(aAreaComponent(N_COMPANY_ID_AREA)) Then
		If Len(oRequest("CompanyID").Item) > 0 Then
			aAreaComponent(N_COMPANY_ID_AREA) = CLng(oRequest("CompanyID").Item)
		Else
			aAreaComponent(N_COMPANY_ID_AREA) = -1
		End If
	End If

	If IsEmpty(aAreaComponent(N_TYPE_ID_AREA)) Then
		If Len(oRequest("AreaTypeID").Item) > 0 Then
			aAreaComponent(N_TYPE_ID_AREA) = CLng(oRequest("AreaTypeID").Item)
		Else
			aAreaComponent(N_TYPE_ID_AREA) = -1
		End If
	End If

	If IsEmpty(aAreaComponent(N_CONFINE_TYPE_ID_AREA)) Then
		If Len(oRequest("ConfineTypeID").Item) > 0 Then
			aAreaComponent(N_CONFINE_TYPE_ID_AREA) = CLng(oRequest("ConfineTypeID").Item)
		Else
			aAreaComponent(N_CONFINE_TYPE_ID_AREA) = -1
		End If
	End If

	If IsEmpty(aAreaComponent(N_LEVEL_TYPE_ID_AREA)) Then
		If Len(oRequest("AreaLevelTypeID").Item) > 0 Then
			aAreaComponent(N_LEVEL_TYPE_ID_AREA) = CLng(oRequest("AreaLevelTypeID").Item)
		Else
			aAreaComponent(N_LEVEL_TYPE_ID_AREA) = -1
		End If
	End If

	If IsEmpty(aAreaComponent(N_CENTER_TYPE_ID_AREA)) Then
		If Len(oRequest("CenterTypeID").Item) > 0 Then
			aAreaComponent(N_CENTER_TYPE_ID_AREA) = CLng(oRequest("CenterTypeID").Item)
		Else
			aAreaComponent(N_CENTER_TYPE_ID_AREA) = -1
		End If
	End If

	If IsEmpty(aAreaComponent(N_CENTER_SUBTYPE_ID_AREA)) Then
		If Len(oRequest("CenterSubtypeID").Item) > 0 Then
			aAreaComponent(N_CENTER_SUBTYPE_ID_AREA) = CLng(oRequest("CenterSubtypeID").Item)
		Else
			aAreaComponent(N_CENTER_SUBTYPE_ID_AREA) = -1
		End If
	End If

	If IsEmpty(aAreaComponent(N_ATTENTION_LEVEL_ID_AREA)) Then
		If Len(oRequest("AttentionLevelID").Item) > 0 Then
			aAreaComponent(N_ATTENTION_LEVEL_ID_AREA) = CLng(oRequest("AttentionLevelID").Item)
		Else
			aAreaComponent(N_ATTENTION_LEVEL_ID_AREA) = -1
		End If
	End If

	If IsEmpty(aAreaComponent(S_ADDRESS_AREA)) Then
		If Len(oRequest("AreaAddress").Item) > 0 Then
			aAreaComponent(S_ADDRESS_AREA) = oRequest("AreaAddress").Item
		Else
			aAreaComponent(S_ADDRESS_AREA) = ""
		End If
	End If
	aAreaComponent(S_ADDRESS_AREA) = Left(aAreaComponent(S_ADDRESS_AREA), 255)

	If IsEmpty(aAreaComponent(S_CITY_AREA)) Then
		If Len(oRequest("AreaCity").Item) > 0 Then
			aAreaComponent(S_CITY_AREA) = oRequest("AreaCity").Item
		Else
			aAreaComponent(S_CITY_AREA) = ""
		End If
	End If
	aAreaComponent(S_CITY_AREA) = Left(aAreaComponent(S_CITY_AREA), 100)

	If IsEmpty(aAreaComponent(S_ZIP_CODE_AREA)) Then
		If Len(oRequest("AreaZip").Item) > 0 Then
			aAreaComponent(S_ZIP_CODE_AREA) = oRequest("AreaZip").Item
		Else
			aAreaComponent(S_ZIP_CODE_AREA) = ""
		End If
	End If
	aAreaComponent(S_ZIP_CODE_AREA) = Left(aAreaComponent(S_ZIP_CODE_AREA), 10)

	If IsEmpty(aAreaComponent(N_ZONE_ID_AREA)) Then
		If Len(oRequest("ZoneID").Item) > 0 Then
			aAreaComponent(N_ZONE_ID_AREA) = CLng(oRequest("ZoneID").Item)
		Else
			aAreaComponent(N_ZONE_ID_AREA) = -1
		End If
	End If

	If IsEmpty(aAreaComponent(N_ECONOMIC_ZONE_ID_AREA)) Then
		If Len(oRequest("EconomicZoneID").Item) > 0 Then
			aAreaComponent(N_ECONOMIC_ZONE_ID_AREA) = CInt(oRequest("EconomicZoneID").Item)
		Else
			aAreaComponent(N_ECONOMIC_ZONE_ID_AREA) = 2
		End If
	End If

	If IsEmpty(aAreaComponent(N_PAYMENT_CENTER_ID_AREA)) Then
		If Len(oRequest("PaymentCenterID").Item) > 0 Then
			aAreaComponent(N_PAYMENT_CENTER_ID_AREA) = CLng(oRequest("PaymentCenterID").Item)
		Else
			aAreaComponent(N_PAYMENT_CENTER_ID_AREA) = -1
		End If
	End If

	If IsEmpty(aAreaComponent(N_GENERATING_AREA_ID_AREA)) Then
		If Len(oRequest("GeneratingAreaID").Item) > 0 Then
			aAreaComponent(N_GENERATING_AREA_ID_AREA) = CLng(oRequest("GeneratingAreaID").Item)
		Else
			aAreaComponent(N_GENERATING_AREA_ID_AREA) = -1
		End If
	End If

	If IsEmpty(aAreaComponent(N_BRANCH_ID_AREA)) Then
		If Len(oRequest("BranchID").Item) > 0 Then
			aAreaComponent(N_BRANCH_ID_AREA) = CLng(oRequest("BranchID").Item)
		Else
			aAreaComponent(N_BRANCH_ID_AREA) = -1
		End If
	End If

	If IsEmpty(aAreaComponent(N_SUBBRANCH_ID_AREA)) Then
		If Len(oRequest("SubBranchID").Item) > 0 Then
			aAreaComponent(N_SUBBRANCH_ID_AREA) = CLng(oRequest("SubBranchID").Item)
		Else
			aAreaComponent(N_SUBBRANCH_ID_AREA) = -1
		End If
	End If

	If IsEmpty(aAreaComponent(N_CASHIER_OFFICE_ID_AREA)) Then
		If Len(oRequest("CashierOfficeID").Item) > 0 Then
			aAreaComponent(N_CASHIER_OFFICE_ID_AREA) = CLng(oRequest("CashierOfficeID").Item)
		Else
			aAreaComponent(N_CASHIER_OFFICE_ID_AREA) = -1
		End If
	End If

	If IsEmpty(aAreaComponent(N_START_DATE_AREA)) Then
		If Len(oRequest("StartYear").Item) > 0 Then
			aAreaComponent(N_START_DATE_AREA) = CLng(oRequest("StartYear").Item & Right(("0" & oRequest("StartMonth").Item), Len("00")) & Right(("0" & oRequest("StartDay").Item), Len("00")))
		ElseIf Len(oRequest("StartDate").Item) > 0 Then
			aAreaComponent(N_START_DATE_AREA) = CLng(oRequest("StartDate").Item)
		Else
			aAreaComponent(N_START_DATE_AREA) = 0
		End If
	End If

	If IsEmpty(aAreaComponent(N_END_DATE_AREA)) Then
		If Len(oRequest("EndYear").Item) > 0 Then
			aAreaComponent(N_END_DATE_AREA) = CLng(oRequest("EndYear").Item & Right(("0" & oRequest("EndMonth").Item), Len("00")) & Right(("0" & oRequest("EndDay").Item), Len("00")))
		ElseIf Len(oRequest("EndDate").Item) > 0 Then
			aAreaComponent(N_END_DATE_AREA) = CLng(oRequest("EndDate").Item)
		Else
			aAreaComponent(N_END_DATE_AREA) = 30000000
		End If
	End If
	If aAreaComponent(N_END_DATE_AREA) = 0 Then aAreaComponent(N_END_DATE_AREA) = 30000000

	If IsEmpty(aAreaComponent(N_FINISH_DATE_AREA)) Then
		If Len(oRequest("FinishYear").Item) > 0 Then
			aAreaComponent(N_FINISH_DATE_AREA) = CLng(oRequest("FinishYear").Item & Right(("0" & oRequest("FinishMonth").Item), Len("00")) & Right(("0" & oRequest("FinishDay").Item), Len("00")))
		ElseIf Len(oRequest("FinishDate").Item) > 0 Then
			aAreaComponent(N_FINISH_DATE_AREA) = CLng(oRequest("FinishDate").Item)
		Else
			aAreaComponent(N_FINISH_DATE_AREA) = 30000000
		End If
	End If
	If aAreaComponent(N_FINISH_DATE_AREA) = 0 Then aAreaComponent(N_FINISH_DATE_AREA) = 30000000

	If IsEmpty(aAreaComponent(N_JOBS_AREA)) Then
		If Len(oRequest("JobsInArea").Item) > 0 Then
			aAreaComponent(N_JOBS_AREA) = CLng(oRequest("JobsInArea").Item)
		Else
			aAreaComponent(N_JOBS_AREA) = 0
		End If
	End If

	If IsEmpty(aAreaComponent(N_TOTAL_JOBS_AREA)) Then
		If Len(oRequest("TotalJobs").Item) > 0 Then
			aAreaComponent(N_TOTAL_JOBS_AREA) = CLng(oRequest("TotalJobs").Item)
		Else
			aAreaComponent(N_TOTAL_JOBS_AREA) = aAreaComponent(N_JOBS_AREA)
		End If
	End If

	If IsEmpty(aAreaComponent(N_STATUS_ID_AREA)) Then
		If Len(oRequest("StatusID").Item) > 0 Then
			aAreaComponent(N_STATUS_ID_AREA) = CLng(oRequest("StatusID").Item)
		Else
			aAreaComponent(N_STATUS_ID_AREA) = -1
		End If
	End If

	If IsEmpty(aAreaComponent(N_ACTIVE_AREA)) Then
		If Len(oRequest("Active").Item) > 0 Then
			aAreaComponent(N_ACTIVE_AREA) = CInt(oRequest("Active").Item)
		Else
			aAreaComponent(N_ACTIVE_AREA) = 1
		End If
	End If

	If IsEmpty(aAreaComponent(N_ONLY_PAYMENT)) Then
		If Len(oRequest("OnlyPayment").Item) > 0 Then
			aAreaComponent(N_ONLY_PAYMENT) = CInt(oRequest("OnlyPayment").Item)
		Else
			aAreaComponent(N_ONLY_PAYMENT) = 0
		End If
	End If

	If IsEmpty(aAreaComponent(N_PREVIOUS_PARENT_ID_AREA)) Then
		If Len(oRequest("PreviousParentID").Item) > 0 Then
			aAreaComponent(N_PREVIOUS_PARENT_ID_AREA) = CLng(oRequest("PreviousParentID").Item)
		Else
			aAreaComponent(N_PREVIOUS_PARENT_ID_AREA) = aAreaComponent(N_PARENT_ID_AREA)
		End If
	End If

	If IsEmpty(aAreaComponent(N_PREVIOUS_JOBS_AREA)) Then
		If Len(oRequest("PreviousJobsInArea").Item) > 0 Then
			aAreaComponent(N_PREVIOUS_JOBS_AREA) = CLng(oRequest("PreviousJobsInArea").Item)
		Else
			aAreaComponent(N_PREVIOUS_JOBS_AREA) = aAreaComponent(N_JOBS_AREA)
		End If
	End If

	If IsEmpty(aAreaComponent(N_PREVIOUS_TOTAL_JOBS_AREA)) Then
		If Len(oRequest("PreviousTotalJobs").Item) > 0 Then
			aAreaComponent(N_PREVIOUS_TOTAL_JOBS_AREA) = CLng(oRequest("PreviousTotalJobs").Item)
		Else
			aAreaComponent(N_PREVIOUS_TOTAL_JOBS_AREA) = aAreaComponent(N_TOTAL_JOBS_AREA)
		End If
	End If

	If IsEmpty(aAreaComponent(N_PREVIOUS_STATUS_ID_AREA)) Then
		If Len(oRequest("PreviousStatusID").Item) > 0 Then
			aAreaComponent(N_PREVIOUS_STATUS_ID_AREA) = CLng(oRequest("PreviousStatusID").Item)
		Else
			aAreaComponent(N_PREVIOUS_STATUS_ID_AREA) = aAreaComponent(N_STATUS_ID_AREA)
		End If
	End If

	If IsEmpty(aAreaComponent(S_POSITIONS_ID_AREA)) Then
		If Len(oRequest("PositionID").Item) > 0 Then
			aAreaComponent(S_POSITIONS_ID_AREA) = ""
			For Each oItem In oRequest("PositionID")
				aAreaComponent(S_POSITIONS_ID_AREA) = aAreaComponent(S_POSITIONS_ID_AREA) & CLng(oItem) & ","
			Next
			If Len(aAreaComponent(S_POSITIONS_ID_AREA)) > 0 Then aAreaComponent(S_POSITIONS_ID_AREA) = Left(aAreaComponent(S_POSITIONS_ID_AREA), (Len(aAreaComponent(S_POSITIONS_ID_AREA)) - Len(",")))
		ElseIf Len(oRequest("PositionIDs").Item) > 0 Then
			aAreaComponent(S_POSITIONS_ID_AREA) = CLng(oRequest("PositionIDs").Item)
		Else
			aAreaComponent(S_POSITIONS_ID_AREA) = ""
		End If
	End If

	If IsEmpty(aAreaComponent(S_JOBS_FOR_POSITIONS_AREA)) Then
		If Len(oRequest("JobsForPosition").Item) > 0 Then
			aAreaComponent(S_JOBS_FOR_POSITIONS_AREA) = ""
			For Each oItem In oRequest("JobsForPosition")
				aAreaComponent(S_JOBS_FOR_POSITIONS_AREA) = aAreaComponent(S_JOBS_FOR_POSITIONS_AREA) & CLng(oItem) & ","
			Next
			If Len(aAreaComponent(S_JOBS_FOR_POSITIONS_AREA)) > 0 Then aAreaComponent(S_JOBS_FOR_POSITIONS_AREA) = Left(aAreaComponent(S_JOBS_FOR_POSITIONS_AREA), (Len(aAreaComponent(S_JOBS_FOR_POSITIONS_AREA)) - Len(",")))
		ElseIf Len(oRequest("JobsForPositions").Item) > 0 Then
			aAreaComponent(S_JOBS_FOR_POSITIONS_AREA) = CLng(oRequest("JobsForPositions").Item)
		Else
			aAreaComponent(S_JOBS_FOR_POSITIONS_AREA) = ""
		End If
	End If

	aAreaComponent(B_SEND_TO_IFRAME_AREA) = False
	aAreaComponent(S_QUERY_CONDITION_AREA) = ""
	aAreaComponent(B_CHECK_FOR_DUPLICATED_AREA) = True
	aAreaComponent(B_IS_DUPLICATED_AREA) = False
	If IsEmpty(aAreaComponent(S_FILTER_CONDITION_AREA)) Then
		If Len(oRequest("FilterCondition").Item) > 0 Then
			aAreaComponent(S_FILTER_CONDITION_AREA) = oRequest("AreaCity").Item
		Else
			aAreaComponent(S_FILTER_CONDITION_AREA) = ""
		End If
	End If

	aAreaComponent(B_COMPONENT_INITIALIZED_AREA) = True
	InitializeAreaComponent = Err.number
	Err.Clear
End Function

Function AddArea(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new area into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddArea"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAreaComponent(B_COMPONENT_INITIALIZED_AREA)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAreaComponent(oRequest, aAreaComponent)
	End If

	If aAreaComponent(N_ID_AREA) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "Areas", "AreaID", "", 1, aAreaComponent(N_ID_AREA), sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		If aAreaComponent(B_CHECK_FOR_DUPLICATED_AREA) Then
			lErrorNumber = CheckExistencyOfArea(aAreaComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aAreaComponent(B_IS_DUPLICATED_AREA) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un área con la clave " & aAreaComponent(S_CODE_AREA) & " o con el nombre " & aAreaComponent(S_NAME_AREA) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AreaComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				lErrorNumber = GetAreaPath(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					If Not CheckAreaInformationConsistency(aAreaComponent, sErrorDescription) Then
						lErrorNumber = -1
					Else
						sErrorDescription = "No se pudo guardar la información del nuevo registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Areas (AreaID, ParentID, AreaCode, AreaShortName, AreaName, AreaPath, URCTAUX, CompanyID, AreaTypeID, ConfineTypeID, AreaLevelTypeID, CenterTypeID, CenterSubtypeID, AttentionLevelID, AreaAddress, AreaCity, AreaZip, ZoneID, EconomicZoneID, PaymentCenterID, GeneratingAreaID, BranchID, SubBranchID, CashierOfficeID, StartDate, EndDate, FinishDate, JobsInArea, TotalJobs, StatusID, Active, OnlyPaymentCenter) Values (" & aAreaComponent(N_ID_AREA) & ", " & aAreaComponent(N_PARENT_ID_AREA) & ", '" & Replace(aAreaComponent(S_CODE_AREA), "'", "") & "', '" & Replace(aAreaComponent(S_SHORT_NAME_AREA), "'", "") & "', '" & Replace(aAreaComponent(S_NAME_AREA), "'", "´") & "', '" & Replace(aAreaComponent(S_PATH_AREA), "'", "") & "', '" & Replace(aAreaComponent(S_URCTAUX_AREA), "'", "") & "', " & aAreaComponent(N_COMPANY_ID_AREA) & ", " & aAreaComponent(N_TYPE_ID_AREA) & ", " & aAreaComponent(N_CONFINE_TYPE_ID_AREA) & ", " & aAreaComponent(N_LEVEL_TYPE_ID_AREA) & ", " & aAreaComponent(N_CENTER_TYPE_ID_AREA) & ", " & aAreaComponent(N_CENTER_SUBTYPE_ID_AREA) & ", " & aAreaComponent(N_ATTENTION_LEVEL_ID_AREA) & ", '" & Replace(aAreaComponent(S_ADDRESS_AREA), "'", "´") & "', '" & Replace(aAreaComponent(S_CITY_AREA), "'", "´") & "', '" & Replace(aAreaComponent(S_ZIP_CODE_AREA), "'", "´") & "', " & aAreaComponent(N_ZONE_ID_AREA) & ", " & aAreaComponent(N_ECONOMIC_ZONE_ID_AREA) & ", " & aAreaComponent(N_PAYMENT_CENTER_ID_AREA) & ", " & aAreaComponent(N_GENERATING_AREA_ID_AREA) & ", " & aAreaComponent(N_BRANCH_ID_AREA) & ", " & aAreaComponent(N_SUBBRANCH_ID_AREA) & ", " & aAreaComponent(N_CASHIER_OFFICE_ID_AREA) & ", " & aAreaComponent(N_START_DATE_AREA) & ", " & aAreaComponent(N_END_DATE_AREA) & ", " & aAreaComponent(N_FINISH_DATE_AREA) & ", " & aAreaComponent(N_JOBS_AREA) & ", " & aAreaComponent(N_TOTAL_JOBS_AREA) & ", " & aAreaComponent(N_STATUS_ID_AREA) & ", " & aAreaComponent(N_ACTIVE_AREA) & ", " & aAreaComponent(N_ONLY_PAYMENT) & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudo guardar la información del nuevo registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into AreasHistoryList (AreaID, AreaDate, ParentID, AreaCode, AreaShortName, AreaName, CompanyID, AreaTypeID, ConfineTypeID, AreaLevelTypeID, CenterTypeID, CenterSubtypeID, AttentionLevelID, ZoneID, EconomicZoneID, PaymentCenterID, GeneratingAreaID, BranchID, SubBranchID, CashierOfficeID, JobsInArea, TotalJobs, StatusID, UserID) Values (" & aAreaComponent(N_ID_AREA) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aAreaComponent(N_PARENT_ID_AREA) & ", '" & Replace(aAreaComponent(S_CODE_AREA), "'", "") & "', '" & Replace(aAreaComponent(S_SHORT_NAME_AREA), "'", "") & "', '" & Replace(aAreaComponent(S_NAME_AREA), "'", "´") & "', " & aAreaComponent(N_COMPANY_ID_AREA) & ", " & aAreaComponent(N_TYPE_ID_AREA) & ", " & aAreaComponent(N_CONFINE_TYPE_ID_AREA) & ", " & aAreaComponent(N_LEVEL_TYPE_ID_AREA) & ", " & aAreaComponent(N_CENTER_TYPE_ID_AREA) & ", " & aAreaComponent(N_CENTER_SUBTYPE_ID_AREA) & ", " & aAreaComponent(N_ATTENTION_LEVEL_ID_AREA) & ", " & aAreaComponent(N_ZONE_ID_AREA) & ", " & aAreaComponent(N_ECONOMIC_ZONE_ID_AREA) & ", " & aAreaComponent(N_PAYMENT_CENTER_ID_AREA) & ", " & aAreaComponent(N_GENERATING_AREA_ID_AREA) & ", " & aAreaComponent(N_BRANCH_ID_AREA) & ", " & aAreaComponent(N_SUBBRANCH_ID_AREA) & ", " & aAreaComponent(N_CASHIER_OFFICE_ID_AREA) & ", " & aAreaComponent(N_JOBS_AREA) & ", " & aAreaComponent(N_TOTAL_JOBS_AREA) & ", " & aAreaComponent(N_STATUS_ID_AREA) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
					End If
				End If
			End If
		End If
	End If

	AddArea = lErrorNumber
	Err.Clear
End Function

Function GetArea(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about an area from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetArea"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAreaComponent(B_COMPONENT_INITIALIZED_AREA)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAreaComponent(oRequest, aAreaComponent)
	End If

	If aAreaComponent(N_ID_AREA) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AreaComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Areas Where AreaID=" & aAreaComponent(N_ID_AREA), "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AreaComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
				oRecordset.Close
			Else
				aAreaComponent(N_PARENT_ID_AREA) = CLng(oRecordset.Fields("ParentID").Value)
				aAreaComponent(S_CODE_AREA) = CStr(oRecordset.Fields("AreaCode").Value)
				aAreaComponent(S_SHORT_NAME_AREA) = CStr(oRecordset.Fields("AreaShortName").Value)
				aAreaComponent(S_NAME_AREA) = CStr(oRecordset.Fields("AreaName").Value)
				aAreaComponent(S_PATH_AREA) = CStr(oRecordset.Fields("AreaPath").Value)
				aAreaComponent(S_URCTAUX_AREA) = CStr(oRecordset.Fields("URCTAUX").Value)
				aAreaComponent(N_COMPANY_ID_AREA) = CLng(oRecordset.Fields("CompanyID").Value)
				aAreaComponent(N_TYPE_ID_AREA) = CLng(oRecordset.Fields("AreaTypeID").Value)
				aAreaComponent(N_CONFINE_TYPE_ID_AREA) = CLng(oRecordset.Fields("ConfineTypeID").Value)
				aAreaComponent(N_LEVEL_TYPE_ID_AREA) = CLng(oRecordset.Fields("AreaLevelTypeID").Value)
				aAreaComponent(N_CENTER_TYPE_ID_AREA) = CLng(oRecordset.Fields("CenterTypeID").Value)
				aAreaComponent(N_CENTER_SUBTYPE_ID_AREA) = CLng(oRecordset.Fields("CenterSubtypeID").Value)
				aAreaComponent(N_ATTENTION_LEVEL_ID_AREA) = CLng(oRecordset.Fields("AttentionLevelID").Value)
				aAreaComponent(S_ADDRESS_AREA) = CStr(oRecordset.Fields("AreaAddress").Value)
				aAreaComponent(S_CITY_AREA) = CStr(oRecordset.Fields("AreaCity").Value)
				aAreaComponent(S_ZIP_CODE_AREA) = CStr(oRecordset.Fields("AreaZip").Value)
				aAreaComponent(N_ZONE_ID_AREA) = CLng(oRecordset.Fields("ZoneID").Value)
				aAreaComponent(N_ECONOMIC_ZONE_ID_AREA) = CInt(oRecordset.Fields("EconomicZoneID").Value)
				aAreaComponent(N_PAYMENT_CENTER_ID_AREA) = CLng(oRecordset.Fields("PaymentCenterID").Value)
				aAreaComponent(N_GENERATING_AREA_ID_AREA) = CLng(oRecordset.Fields("GeneratingAreaID").Value)
				aAreaComponent(N_BRANCH_ID_AREA) = CLng(oRecordset.Fields("BranchID").Value)
				aAreaComponent(N_SUBBRANCH_ID_AREA) = CLng(oRecordset.Fields("SubBranchID").Value)
				aAreaComponent(N_CASHIER_OFFICE_ID_AREA) = CLng(oRecordset.Fields("CashierOfficeID").Value)
				aAreaComponent(N_START_DATE_AREA) = CLng(oRecordset.Fields("StartDate").Value)
				aAreaComponent(N_END_DATE_AREA) = CLng(oRecordset.Fields("EndDate").Value)
				aAreaComponent(N_FINISH_DATE_AREA) = CLng(oRecordset.Fields("FinishDate").Value)
				aAreaComponent(N_JOBS_AREA) = CLng(oRecordset.Fields("JobsInArea").Value)
				aAreaComponent(N_TOTAL_JOBS_AREA) = CLng(oRecordset.Fields("TotalJobs").Value)
				aAreaComponent(N_STATUS_ID_AREA) = CLng(oRecordset.Fields("StatusID").Value)
				aAreaComponent(N_ACTIVE_AREA) = CInt(oRecordset.Fields("Active").Value)
				aAreaComponent(N_ONLY_PAYMENT) = CInt(oRecordset.Fields("OnlyPaymentCenter").Value)

				aAreaComponent(N_PREVIOUS_PARENT_ID_AREA) = CLng(oRecordset.Fields("ParentID").Value)
				aAreaComponent(N_PREVIOUS_JOBS_AREA) = CLng(oRecordset.Fields("JobsInArea").Value)
				aAreaComponent(N_PREVIOUS_STATUS_ID_AREA) = CLng(oRecordset.Fields("StatusID").Value)
				oRecordset.Close

				aAreaComponent(N_PREVIOUS_TOTAL_JOBS_AREA) = aAreaComponent(N_TOTAL_JOBS_AREA)

				aAreaComponent(S_POSITIONS_ID_AREA) = ""
				aAreaComponent(S_JOBS_FOR_POSITIONS_AREA) = ""
				sErrorDescription = "No se pudo obtener la información del registro."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From AreasPositionsLKP Where (AreaID=" & aAreaComponent(N_ID_AREA) & ") Order By PositionID", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						aAreaComponent(S_POSITIONS_ID_AREA) = aAreaComponent(S_POSITIONS_ID_AREA) & CLng(oRecordset.Fields("PositionID").Value) & ","
						aAreaComponent(S_JOBS_FOR_POSITIONS_AREA) = aAreaComponent(S_JOBS_FOR_POSITIONS_AREA) & CLng(oRecordset.Fields("JobsInArea").Value) & ","
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
					If Len(aAreaComponent(S_POSITIONS_ID_AREA)) > 0 Then aAreaComponent(S_POSITIONS_ID_AREA) = Left(aAreaComponent(S_POSITIONS_ID_AREA), (Len(aAreaComponent(S_POSITIONS_ID_AREA)) - Len(",")))
					If Len(aAreaComponent(S_JOBS_FOR_POSITIONS_AREA)) > 0 Then aAreaComponent(S_JOBS_FOR_POSITIONS_AREA) = Left(aAreaComponent(S_JOBS_FOR_POSITIONS_AREA), (Len(aAreaComponent(S_JOBS_FOR_POSITIONS_AREA)) - Len(",")))
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetArea = lErrorNumber
	Err.Clear
End Function

Function GetAreaByCode(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about an area from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAreaByCode"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAreaComponent(B_COMPONENT_INITIALIZED_AREA)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAreaComponent(oRequest, aAreaComponent)
	End If

	If Len(aAreaComponent(S_CODE_AREA)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el código del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AreaComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Areas Where (AreaCode='" & aAreaComponent(S_CODE_AREA) & "')", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AreaComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
				oRecordset.Close
			Else
				aAreaComponent(N_PARENT_ID_AREA) = CLng(oRecordset.Fields("ParentID").Value)
				aAreaComponent(N_ID_AREA) = CLng(oRecordset.Fields("AreaID").Value)
				aAreaComponent(S_SHORT_NAME_AREA) = CStr(oRecordset.Fields("AreaShortName").Value)
				aAreaComponent(S_NAME_AREA) = CStr(oRecordset.Fields("AreaName").Value)
				aAreaComponent(S_PATH_AREA) = CStr(oRecordset.Fields("AreaPath").Value)
				aAreaComponent(S_URCTAUX_AREA) = CStr(oRecordset.Fields("URCTAUX").Value)
				aAreaComponent(N_COMPANY_ID_AREA) = CLng(oRecordset.Fields("CompanyID").Value)
				aAreaComponent(N_TYPE_ID_AREA) = CLng(oRecordset.Fields("AreaTypeID").Value)
				aAreaComponent(N_CONFINE_TYPE_ID_AREA) = CLng(oRecordset.Fields("ConfineTypeID").Value)
				aAreaComponent(N_LEVEL_TYPE_ID_AREA) = CLng(oRecordset.Fields("AreaLevelTypeID").Value)
				aAreaComponent(N_CENTER_TYPE_ID_AREA) = CLng(oRecordset.Fields("CenterTypeID").Value)
				aAreaComponent(N_CENTER_SUBTYPE_ID_AREA) = CLng(oRecordset.Fields("CenterSubtypeID").Value)
				aAreaComponent(N_ATTENTION_LEVEL_ID_AREA) = CLng(oRecordset.Fields("AttentionLevelID").Value)
				aAreaComponent(S_ADDRESS_AREA) = CStr(oRecordset.Fields("AreaAddress").Value)
				aAreaComponent(S_CITY_AREA) = CStr(oRecordset.Fields("AreaCity").Value)
				aAreaComponent(S_ZIP_CODE_AREA) = CStr(oRecordset.Fields("AreaZip").Value)
				aAreaComponent(N_ZONE_ID_AREA) = CLng(oRecordset.Fields("ZoneID").Value)
				aAreaComponent(N_ECONOMIC_ZONE_ID_AREA) = CInt(oRecordset.Fields("EconomicZoneID").Value)
				aAreaComponent(N_PAYMENT_CENTER_ID_AREA) = CLng(oRecordset.Fields("PaymentCenterID").Value)
				aAreaComponent(N_GENERATING_AREA_ID_AREA) = CLng(oRecordset.Fields("GeneratingAreaID").Value)
				aAreaComponent(N_BRANCH_ID_AREA) = CLng(oRecordset.Fields("BranchID").Value)
				aAreaComponent(N_SUBBRANCH_ID_AREA) = CLng(oRecordset.Fields("SubBranchID").Value)
				aAreaComponent(N_CASHIER_OFFICE_ID_AREA) = CLng(oRecordset.Fields("CashierOfficeID").Value)
				aAreaComponent(N_START_DATE_AREA) = CLng(oRecordset.Fields("StartDate").Value)
				aAreaComponent(N_END_DATE_AREA) = CLng(oRecordset.Fields("EndDate").Value)
				aAreaComponent(N_FINISH_DATE_AREA) = CLng(oRecordset.Fields("FinishDate").Value)
				aAreaComponent(N_JOBS_AREA) = CLng(oRecordset.Fields("JobsInArea").Value)
				aAreaComponent(N_TOTAL_JOBS_AREA) = CLng(oRecordset.Fields("TotalJobs").Value)
				aAreaComponent(N_STATUS_ID_AREA) = CLng(oRecordset.Fields("StatusID").Value)
				aAreaComponent(N_ACTIVE_AREA) = CInt(oRecordset.Fields("Active").Value)
				aAreaComponent(N_ONLY_PAYMENT) = CInt(oRecordset.Fields("OnlyPaymentCenter").Value)

				aAreaComponent(N_PREVIOUS_PARENT_ID_AREA) = CLng(oRecordset.Fields("ParentID").Value)
				aAreaComponent(N_PREVIOUS_JOBS_AREA) = CLng(oRecordset.Fields("JobsInArea").Value)
				aAreaComponent(N_PREVIOUS_STATUS_ID_AREA) = CLng(oRecordset.Fields("StatusID").Value)
				oRecordset.Close

				aAreaComponent(N_PREVIOUS_TOTAL_JOBS_AREA) = aAreaComponent(N_TOTAL_JOBS_AREA)

				aAreaComponent(S_POSITIONS_ID_AREA) = ""
				aAreaComponent(S_JOBS_FOR_POSITIONS_AREA) = ""
				sErrorDescription = "No se pudo obtener la información del registro."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From AreasPositionsLKP Where (AreaID=" & aAreaComponent(N_ID_AREA) & ") Order By PositionID", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						aAreaComponent(S_POSITIONS_ID_AREA) = aAreaComponent(S_POSITIONS_ID_AREA) & CLng(oRecordset.Fields("PositionID").Value) & ","
						aAreaComponent(S_JOBS_FOR_POSITIONS_AREA) = aAreaComponent(S_JOBS_FOR_POSITIONS_AREA) & CLng(oRecordset.Fields("JobsInArea").Value) & ","
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
					If Len(aAreaComponent(S_POSITIONS_ID_AREA)) > 0 Then aAreaComponent(S_POSITIONS_ID_AREA) = Left(aAreaComponent(S_POSITIONS_ID_AREA), (Len(aAreaComponent(S_POSITIONS_ID_AREA)) - Len(",")))
					If Len(aAreaComponent(S_JOBS_FOR_POSITIONS_AREA)) > 0 Then aAreaComponent(S_JOBS_FOR_POSITIONS_AREA) = Left(aAreaComponent(S_JOBS_FOR_POSITIONS_AREA), (Len(aAreaComponent(S_JOBS_FOR_POSITIONS_AREA)) - Len(",")))
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetAreaByCode = lErrorNumber
	Err.Clear
End Function

Function GetAreaLevel(oRequest, oADODBConnection, iAreaID, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To get the path for an area from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAreaLevel"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	If aAreaComponent(N_PARENT_ID_AREA) = -1 Then
		aAreaComponent(S_PATH_AREA) = ",-1,"
	Else
		sErrorDescription = "No se pudo obtener la ruta del área."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaLevelTypeID, ZoneID From Areas Where (AreaID=" & iAreaID & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				aAreaComponent(N_LEVEL_AREA) = -1
                aAreaComponent(N_PARENT_ZONE) = -1
			Else
				aAreaComponent(N_LEVEL_AREA) = CInt(oRecordset.Fields("AreaLevelTypeID").Value)
                aAreaComponent(N_PARENT_ZONE) = CInt(oRecordset.Fields("ZoneID").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetAreaLevel = lErrorNumber
	Err.Clear
End Function

Function GetAreaParentID(oRequest, oADODBConnection, iAreaID, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To get the path for an area from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAreaLevel"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	If aAreaComponent(N_PARENT_ID_AREA) = -1 Then
		aAreaComponent(S_PATH_AREA) = ",-1,"
	Else
		sErrorDescription = "No se pudo obtener la ruta del área."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ParentID From Areas Where (AreaID=" & iAreaID & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				aAreaComponent(N_PARENT_ID2_AREA) = -1
			Else
				aAreaComponent(N_PARENT_ID2_AREA) = CInt(oRecordset.Fields("ParentID").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetAreaLevel = lErrorNumber
	Err.Clear
End Function

Function GetAreaPath(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To get the path for an area from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAreaPath"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAreaComponent(B_COMPONENT_INITIALIZED_AREA)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAreaComponent(oRequest, aAreaComponent)
	End If

	If aAreaComponent(N_PARENT_ID_AREA) = -1 Then
		aAreaComponent(S_PATH_AREA) = ",-1,"
	Else
		sErrorDescription = "No se pudo obtener la ruta del área."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaPath From Areas Where AreaID=" & aAreaComponent(N_PARENT_ID_AREA), "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				aAreaComponent(S_PATH_AREA) = ",-1,"
			Else
				aAreaComponent(S_PATH_AREA) = CStr(oRecordset.Fields("AreaPath").Value)
			End If
			aAreaComponent(S_PATH_AREA) = aAreaComponent(S_PATH_AREA) & aAreaComponent(N_ID_AREA) & ","
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetAreaPath = lErrorNumber
	Err.Clear
End Function

Function GetAreas(oRequest, oADODBConnection, aAreaComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the areas from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAreaComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAreas"
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAreaComponent(B_COMPONENT_INITIALIZED_AREA)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAreaComponent(oRequest, aAreaComponent)
	End If

	If (Len(aAreaComponent(S_QUERY_CONDITION_AREA)) > 0) Or (StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0) Then
		sCondition = Trim(aAreaComponent(S_QUERY_CONDITION_AREA))
		If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
			sCondition = Trim(sCondition & " And (Areas.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))")
		End If
		If InStr(1, sCondition, "And ", vbTextCompare) <> 1 Then
			sCondition = "And " & sCondition
		End If
	End If
	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Areas.*, CompanyShortName, CompanyName, AreaTypeShortName, AreaTypeName, ConfineTypeShortName, ConfineTypeName, CenterTypeShortName, CenterTypeName, CenterSubtypeShortName, CenterSubtypeName, AttentionLevelShortName, AttentionLevelName, Zones01.ZoneCode, Zones01.ZoneName, Zones03.ZoneName As Poblacion, Zones02.ZoneName As Municipio, EconomicZoneCode, EconomicZoneName, PaymentCenters.AreaCode As PaymentCenterShortName, PaymentCenters.AreaName As PaymentCenterName, CashierOfficeShortName, CashierOfficeName, StatusName From Areas, Companies, AreaTypes, ConfineTypes, CenterTypes, CenterSubtypes, AttentionLevels, Zones As Zones03, Zones As Zones02, Zones As Zones01, EconomicZones, Areas As PaymentCenters, CashierOffices, StatusAreas Where (Areas.CompanyID=Companies.CompanyID) And (Areas.AreaTypeID=AreaTypes.AreaTypeID) And (Areas.ConfineTypeID=ConfineTypes.ConfineTypeID) And (Areas.CenterTypeID=CenterTypes.CenterTypeID) And (Areas.CenterSubtypeID=CenterSubtypes.CenterSubtypeID) And (Areas.AttentionLevelID=AttentionLevels.AttentionLevelID) And (Areas.ZoneID=Zones03.ZoneID) And (Zones03.ParentID=Zones02.ZoneID) And (Zones02.ParentID=Zones01.ZoneID) And (Areas.EconomicZoneID=EconomicZones.EconomicZoneID) And (Areas.PaymentCenterID=PaymentCenters.AreaID) And (Areas.CashierOfficeID=CashierOffices.CashierOfficeID) And (Areas.StatusID=StatusAreas.StatusID) And (Areas.AreaID>-1) " & sCondition & " Order By Areas.AreaShortName, Areas.AreaName", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetAreas = lErrorNumber
	Err.Clear
End Function

Function ModifyArea(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing area in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyArea"
	Dim oRecordset
	Dim sDate
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAreaComponent(B_COMPONENT_INITIALIZED_AREA)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAreaComponent(oRequest, aAreaComponent)
	End If

	If aAreaComponent(N_ID_AREA) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AreaComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aAreaComponent(B_CHECK_FOR_DUPLICATED_AREA) Then
			lErrorNumber = CheckExistencyOfArea(aAreaComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aAreaComponent(B_IS_DUPLICATED_AREA) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un área con la clave corta " & aAreaComponent(S_CODE_AREA) & " o de diez posiciones " & aAreaComponent(S_SHORT_NAME_AREA) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AreaComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				lErrorNumber = GetAreaPath(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					If Not CheckAreaInformationConsistency(aAreaComponent, sErrorDescription) Then
						lErrorNumber = -1
					Else
						sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
						sErrorDescription = "No se pudo modificar la información del registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Areas Set ParentID=" & aAreaComponent(N_PARENT_ID_AREA) & ", AreaCode='" & Replace(aAreaComponent(S_CODE_AREA), "'", "") & "', AreaShortName='" & Replace(aAreaComponent(S_SHORT_NAME_AREA), "'", "") & "', AreaName='" & Replace(aAreaComponent(S_NAME_AREA), "'", "") & "', AreaPath='" & Replace(aAreaComponent(S_PATH_AREA), "'", "") & "', URCTAUX='" & Replace(aAreaComponent(S_URCTAUX_AREA), "'", "") & "', CompanyID=" & aAreaComponent(N_COMPANY_ID_AREA) & ", AreaTypeID=" & aAreaComponent(N_TYPE_ID_AREA) & ", ConfineTypeID=" & aAreaComponent(N_CONFINE_TYPE_ID_AREA) & ", AreaLevelTypeID=" & aAreaComponent(N_LEVEL_TYPE_ID_AREA) & ", CenterTypeID=" & aAreaComponent(N_CENTER_TYPE_ID_AREA) & ", CenterSubtypeID=" & aAreaComponent(N_CENTER_SUBTYPE_ID_AREA) & ", AttentionLevelID=" & aAreaComponent(N_ATTENTION_LEVEL_ID_AREA) & ", AreaAddress='" & Replace(aAreaComponent(S_ADDRESS_AREA), "'", "´") & "', AreaCity='" & Replace(aAreaComponent(S_CITY_AREA), "'", "´") & "', AreaZip='" & Replace(aAreaComponent(S_ZIP_CODE_AREA), "'", "´") & "', ZoneID=" & aAreaComponent(N_ZONE_ID_AREA) & ", EconomicZoneID=" & aAreaComponent(N_ECONOMIC_ZONE_ID_AREA) & ", PaymentCenterID=" & aAreaComponent(N_PAYMENT_CENTER_ID_AREA) & ", GeneratingAreaID=" & aAreaComponent(N_GENERATING_AREA_ID_AREA) & ", BranchID=" & aAreaComponent(N_BRANCH_ID_AREA) & ", SubBranchID=" & aAreaComponent(N_SUBBRANCH_ID_AREA) & ", CashierOfficeID=" & aAreaComponent(N_CASHIER_OFFICE_ID_AREA) & ", StartDate=" & aAreaComponent(N_START_DATE_AREA) & ", EndDate=" & aAreaComponent(N_END_DATE_AREA) & ", FinishDate=" & aAreaComponent(N_FINISH_DATE_AREA) & ", JobsInArea=" & aAreaComponent(N_JOBS_AREA) & ", TotalJobs=" & aAreaComponent(N_TOTAL_JOBS_AREA) & ", StatusID=" & aAreaComponent(N_STATUS_ID_AREA) & ", Active=" & aAreaComponent(N_ACTIVE_AREA) & ", OnlyPaymentCenter=" & aAreaComponent(N_ONLY_PAYMENT) & " Where (AreaID=" & aAreaComponent(N_ID_AREA) & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						If (lErrorNumber = 0) And (aAreaComponent(N_JOBS_AREA) <> aAreaComponent(N_PREVIOUS_JOBS_AREA)) Then
							sErrorDescription = "No se pudo modificar la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID From AreasHistoryList Where (AreaID=" & aAreaComponent(N_ID_AREA) & ") And (AreaDate=" & sDate & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo modificar la información del registro."
								If oRecordset.EOF Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into AreasHistoryList (AreaID, AreaDate, ParentID, AreaCode, AreaShortName, AreaName, CompanyID, AreaTypeID, ConfineTypeID, AreaLevelTypeID, CenterTypeID, CenterSubtypeID, AttentionLevelID, ZoneID, EconomicZoneID, PaymentCenterID, GeneratingAreaID, BranchID, SubBranchID, CashierOfficeID, JobsInArea, TotalJobs, StatusID, UserID) Values (" & aAreaComponent(N_ID_AREA) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aAreaComponent(N_PARENT_ID_AREA) & ", '" & Replace(aAreaComponent(S_CODE_AREA), "'", "") & "', '" & Replace(aAreaComponent(S_SHORT_NAME_AREA), "'", "") & "', '" & Replace(aAreaComponent(S_NAME_AREA), "'", "´") & "', " & aAreaComponent(N_COMPANY_ID_AREA) & ", " & aAreaComponent(N_TYPE_ID_AREA) & ", " & aAreaComponent(N_CONFINE_TYPE_ID_AREA) & ", " & aAreaComponent(N_LEVEL_TYPE_ID_AREA) & ", " & aAreaComponent(N_CENTER_TYPE_ID_AREA) & ", " & aAreaComponent(N_CENTER_SUBTYPE_ID_AREA) & ", " & aAreaComponent(N_ATTENTION_LEVEL_ID_AREA) & ", " & aAreaComponent(N_ZONE_ID_AREA) & ", " & aAreaComponent(N_ECONOMIC_ZONE_ID_AREA) & ", " & aAreaComponent(N_PAYMENT_CENTER_ID_AREA) & ", " & aAreaComponent(N_GENERATING_AREA_ID_AREA) & ", " & aAreaComponent(N_BRANCH_ID_AREA) & ", " & aAreaComponent(N_SUBBRANCH_ID_AREA) & ", " & aAreaComponent(N_CASHIER_OFFICE_ID_AREA) & ", " & aAreaComponent(N_JOBS_AREA) & ", " & aAreaComponent(N_TOTAL_JOBS_AREA) & ", " & aAreaComponent(N_STATUS_ID_AREA) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								Else
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update AreasHistoryList Set ParentID=" & aAreaComponent(N_PARENT_ID_AREA) & ", AreaCode='" & Replace(aAreaComponent(S_CODE_AREA), "'", "") & "', AreaShortName='" & Replace(aAreaComponent(S_SHORT_NAME_AREA), "'", "") & "', AreaName='" & Replace(aAreaComponent(S_NAME_AREA), "'", "") & "', CompanyID=" & aAreaComponent(N_COMPANY_ID_AREA) & ", AreaTypeID=" & aAreaComponent(N_TYPE_ID_AREA) & ", ConfineTypeID=" & aAreaComponent(N_CONFINE_TYPE_ID_AREA) & ", AreaLevelTypeID=" & aAreaComponent(N_LEVEL_TYPE_ID_AREA) & ", CenterTypeID=" & aAreaComponent(N_CENTER_TYPE_ID_AREA) & ", CenterSubtypeID=" & aAreaComponent(N_CENTER_SUBTYPE_ID_AREA) & ", AttentionLevelID=" & aAreaComponent(N_ATTENTION_LEVEL_ID_AREA) & ", ZoneID=" & aAreaComponent(N_ZONE_ID_AREA) & ", EconomicZoneID=" & aAreaComponent(N_ECONOMIC_ZONE_ID_AREA) & ", PaymentCenterID=" & aAreaComponent(N_PAYMENT_CENTER_ID_AREA) & ", GeneratingAreaID=" & aAreaComponent(N_GENERATING_AREA_ID_AREA) & ", BranchID=" & aAreaComponent(N_BRANCH_ID_AREA) & ", SubBranchID=" & aAreaComponent(N_SUBBRANCH_ID_AREA) & ", CashierOfficeID=" & aAreaComponent(N_CASHIER_OFFICE_ID_AREA) & ", JobsInArea=" & aAreaComponent(N_JOBS_AREA) & ", TotalJobs=" & aAreaComponent(N_TOTAL_JOBS_AREA) & ", StatusID=" & aAreaComponent(N_STATUS_ID_AREA) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (AreaID=" & aAreaComponent(N_ID_AREA) & ") And (AreaDate=" & sDate & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								End If
								oRecordset.Close
								If lErrorNumber = 0 Then
									lErrorNumber = ModifyAreaAmount(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	ModifyArea = lErrorNumber
	Err.Clear
End Function

Function ModifyAreaAmount(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the area amount and the amount of the parents
'Inputs:  oRequest, oADODBConnection
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyAreaAmount"
	Dim oRecordset
	Dim alAreaIDs
	Dim iIndex
	Dim sDate
	Dim lErrorNumber
	Dim aTempAreaComponent()
	Dim bComponentInitialized

	bComponentInitialized = aAreaComponent(B_COMPONENT_INITIALIZED_AREA)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAreaComponent(oRequest, aAreaComponent)
	End If

	If aAreaComponent(N_ID_AREA) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AreaComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Len(aAreaComponent(S_PATH_AREA)) = 0 Then lErrorNumber = GetAreaPath(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
		sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
		alAreaIDs = Split(aAreaComponent(S_PATH_AREA), ",", -1, vbBinaryCompare)
		For iIndex = 0 To UBound(alAreaIDs)
			If Len(alAreaIDs(iIndex)) > 0 Then
				If (CLng(alAreaIDs(iIndex)) > -1) And (CLng(alAreaIDs(iIndex)) <> aAreaComponent(N_ID_AREA)) Then
					Redim aTempAreaComponent(N_AREA_COMPONENT_SIZE)
					aTempAreaComponent(N_ID_AREA) = CLng(alAreaIDs(iIndex))
					lErrorNumber = GetArea(oRequest, oADODBConnection, aTempAreaComponent, sErrorDescription)
					aTempAreaComponent(N_TOTAL_JOBS_AREA) = aTempAreaComponent(N_TOTAL_JOBS_AREA) + (aAreaComponent(N_JOBS_AREA) - aAreaComponent(N_PREVIOUS_JOBS_AREA))

					sErrorDescription = "No se pudo modificar la información del registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Areas Set TotalJobs=" & aTempAreaComponent(N_TOTAL_JOBS_AREA) & " Where (AreaID=" & aTempAreaComponent(N_ID_AREA) & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

					sErrorDescription = "No se pudo modificar la información del registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID From AreasHistoryList Where (AreaID=" & aTempAreaComponent(N_ID_AREA) & ") And (AreaDate=" & sDate & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudo modificar la información del registro."
						If oRecordset.EOF Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into AreasHistoryList (AreaID, AreaDate, ParentID, AreaCode, AreaShortName, AreaName, CompanyID, AreaTypeID, ConfineTypeID, AreaLevelTypeID, CenterTypeID, CenterSubtypeID, AttentionLevelID, ZoneID, EconomicZoneID, PaymentCenterID, GeneratingAreaID, BranchID, SubBranchID, CashierOfficeID, JobsInArea, TotalJobs, StatusID, UserID) Values (" & aAreaComponent(N_ID_AREA) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aAreaComponent(N_PARENT_ID_AREA) & ", '" & Replace(aAreaComponent(S_CODE_AREA), "'", "") & "', '" & Replace(aAreaComponent(S_SHORT_NAME_AREA), "'", "") & "', '" & Replace(aAreaComponent(S_NAME_AREA), "'", "´") & "', " & aAreaComponent(N_COMPANY_ID_AREA) & ", " & aAreaComponent(N_TYPE_ID_AREA) & ", " & aAreaComponent(N_CONFINE_TYPE_ID_AREA) & ", " & aAreaComponent(N_LEVEL_TYPE_ID_AREA) & ", " & aAreaComponent(N_CENTER_TYPE_ID_AREA) & ", " & aAreaComponent(N_CENTER_SUBTYPE_ID_AREA) & ", " & aAreaComponent(N_ATTENTION_LEVEL_ID_AREA) & ", " & aAreaComponent(N_ZONE_ID_AREA) & ", " & aAreaComponent(N_ECONOMIC_ZONE_ID_AREA) & ", " & aAreaComponent(N_PAYMENT_CENTER_ID_AREA) & ", " & aAreaComponent(N_GENERATING_AREA_ID_AREA) & ", " & aAreaComponent(N_BRANCH_ID_AREA) & ", " & aAreaComponent(N_SUBBRANCH_ID_AREA) & ", " & aAreaComponent(N_CASHIER_OFFICE_ID_AREA) & ", " & aAreaComponent(N_JOBS_AREA) & ", " & aAreaComponent(N_TOTAL_JOBS_AREA) & ", " & aAreaComponent(N_STATUS_ID_AREA) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update AreasHistoryList Set ParentID=" & aTempAreaComponent(N_PARENT_ID_AREA) & ", JobsInArea=" & aTempAreaComponent(N_JOBS_AREA) & ", TotalJobs=" & aTempAreaComponent(N_TOTAL_JOBS_AREA) & ", StatusID=" & aTempAreaComponent(N_STATUS_ID_AREA) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (AreaID=" & aTempAreaComponent(N_ID_AREA) & ") And (AreaDate=" & sDate & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
					End If
				End If
			End If
		Next
	End If

	Set oRecordset = Nothing
	ModifyAreaAmount = lErrorNumber
	Err.Clear
End Function

Function ModifyAreaPositions(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing area in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyAreaPositions"
	Dim alPositionIDs
	Dim alJobs
	Dim iIndex
	Dim oRecordset
	Dim sDate
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAreaComponent(B_COMPONENT_INITIALIZED_AREA)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAreaComponent(oRequest, aAreaComponent)
	End If

	If aAreaComponent(N_ID_AREA) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AreaComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo modificar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From AreasPositionsLKP Where (AreaID=" & aAreaComponent(N_ID_AREA) & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		If (lErrorNumber = 0) And (Len(aAreaComponent(S_POSITIONS_ID_AREA)) > 0) Then
			alPositionIDs = Split(aAreaComponent(S_POSITIONS_ID_AREA), ",", -1, vbBinaryCompare)
			alJobs = Split(aAreaComponent(S_JOBS_FOR_POSITIONS_AREA), ",", -1, vbBinaryCompare)
			For iIndex = 0 To UBound(alPositionIDs)
				sErrorDescription = "No se pudo modificar la información del registro."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into AreasPositionsLKP (AreaID, PositionID, JobsInArea) Values (" & aAreaComponent(N_ID_AREA) & ", " & alPositionIDs(iIndex) & ", " & alJobs(iIndex) & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
			Next
		End If

		sDate = Left(GetSerialNumberForDate(""), Len("00000000"))
		sErrorDescription = "No se pudo modificar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From AreasPositionsHistoryList Where (AreaID=" & aAreaComponent(N_ID_AREA) & ") And (StartDate=" & sDate & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo modificar la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update AreasPositionsHistoryList Set EndDate=" & sDate & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (AreaID=" & aAreaComponent(N_ID_AREA) & ") And (EndDate=30000000)", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		For iIndex = 0 To UBound(alPositionIDs)
			sErrorDescription = "No se pudo modificar la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into AreasPositionsHistoryList (AreaID, PositionID, StartDate, EndDate, JobsInArea, StartUserID, EndUserID) Values (" & aAreaComponent(N_ID_AREA) & ", " & alPositionIDs(iIndex) & ", " & sDate & ", 30000000, " & alJobs(iIndex) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", -1)", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
		Next
	End If

	ModifyAreaPositions = lErrorNumber
	Err.Clear
End Function

Function SetActiveForArea(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given area
'Inputs:  oRequest, oADODBConnection
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForArea"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAreaComponent(B_COMPONENT_INITIALIZED_AREA)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAreaComponent(oRequest, aAreaComponent)
	End If

	If aAreaComponent(N_ID_AREA) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AreaComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo modificar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Areas Set Active=" & CInt(oRequest("SetActive").Item) & " Where (AreaID=" & aAreaComponent(N_ID_AREA) & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	SetActiveForArea = lErrorNumber
	Err.Clear
End Function

Function RemoveArea(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To remove an area from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveArea"
    Dim oRecordset
    Dim oRecordset1
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAreaComponent(B_COMPONENT_INITIALIZED_AREA)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAreaComponent(oRequest, aAreaComponent)
	End If

	If aAreaComponent(N_ID_AREA) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el registro a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AreaComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
        lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Areas Where (ParentID=" & aAreaComponent(N_ID_AREA) & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
        If Not oRecordset.EOF Then
		    lErrorNumber = -1
		    sErrorDescription = "No se puede eliminar el registro seleccionado debido a que existen centros de trabajo/pago asignados a él."
            Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AreaComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
        Else
            lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Jobs Where (AreaID=" & aAreaComponent(N_ID_AREA) & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset1)
            If Not oRecordset1.EOF Then
		        lErrorNumber = -1
		        sErrorDescription = "No se puede eliminar el registro seleccionado debido a que existen plazas asignadas a este centro de trabajo."
                Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AreaComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
            Else
		        sErrorDescription = "No se pudo eliminar la información del registro."
		        lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Areas Where (AreaID=" & aAreaComponent(N_ID_AREA) & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		        If lErrorNumber = 0 Then
			        sErrorDescription = "No se pudo eliminar la información del registro."
			        lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From AreasHistoryList Where (AreaID=" & aAreaComponent(N_ID_AREA) & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

			        'sErrorDescription = "No se pudo eliminar la información del registro."
			        'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Jobs Where (AreaID=" & aAreaComponent(N_ID_AREA) & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

			        sErrorDescription = "No se pudo eliminar la información del registro."
			        lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From AreasPositionsHistoryList Where (AreaID=" & aAreaComponent(N_ID_AREA) & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		        End If
            End If
        End If
	End If

	Set oRecordset = Nothing
    Set oRecordset1 = Nothing
    RemoveArea = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfArea(aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific area exists in the database
'Inputs:  aAreaComponent
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfArea"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAreaComponent(B_COMPONENT_INITIALIZED_AREA)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAreaComponent(oRequest, aAreaComponent)
	End If

	If Len(aAreaComponent(S_NAME_AREA)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el nombre del registro para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AreaComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Areas Where (AreaID<>" & aAreaComponent(N_ID_AREA) & ") And (ParentID=" & aAreaComponent(N_PARENT_ID_AREA) & ") And ((AreaCode='" & Replace(aAreaComponent(S_CODE_AREA), "'", "") & "') Or (AreaName='" & Replace(aAreaComponent(S_SHORT_NAME_AREA), "'", "") & "'))", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aAreaComponent(B_IS_DUPLICATED_AREA) = True
				aAreaComponent(N_ID_AREA) = -1
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfArea = lErrorNumber
	Err.Clear
End Function

Function CheckAreaInformationConsistency(aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aAreaComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckAreaInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aAreaComponent(N_ID_AREA)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del registro no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aAreaComponent(N_PARENT_ID_AREA)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del área a la que pertenece esta área no es un valor numérico."
		bIsCorrect = False
	End If
	If Len(aAreaComponent(S_CODE_AREA)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La clave del registro está vacía."
		bIsCorrect = False
	End If
	If Len(aAreaComponent(S_SHORT_NAME_AREA)) = 0 Then aAreaComponent(S_SHORT_NAME_AREA) = aAreaComponent(S_CODE_AREA)
	If Len(aAreaComponent(S_NAME_AREA)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del registro está vacío."
		bIsCorrect = False
	End If
	If Len(aAreaComponent(S_PATH_AREA)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La ruta del área está vacía."
		bIsCorrect = False
	End If
	If Not IsNumeric(aAreaComponent(N_COMPANY_ID_AREA)) Then aAreaComponent(N_COMPANY_ID_AREA) = -1
	If Not IsNumeric(aAreaComponent(N_TYPE_ID_AREA)) Then aAreaComponent(N_TYPE_ID_AREA) = -1
	If Not IsNumeric(aAreaComponent(N_CONFINE_TYPE_ID_AREA)) Then aAreaComponent(N_CONFINE_TYPE_ID_AREA) = -1
	If Not IsNumeric(aAreaComponent(N_LEVEL_TYPE_ID_AREA)) Then aAreaComponent(N_LEVEL_TYPE_ID_AREA) = -1
	If Not IsNumeric(aAreaComponent(N_CENTER_TYPE_ID_AREA)) Then aAreaComponent(N_CENTER_TYPE_ID_AREA) = -1
	If Not IsNumeric(aAreaComponent(N_CENTER_SUBTYPE_ID_AREA)) Then aAreaComponent(N_CENTER_SUBTYPE_ID_AREA) = -1
	If Not IsNumeric(aAreaComponent(N_ATTENTION_LEVEL_ID_AREA)) Then aAreaComponent(N_ATTENTION_LEVEL_ID_AREA) = 1
	If Len(aAreaComponent(S_ADDRESS_AREA)) = 0 Then aAreaComponent(S_ADDRESS_AREA) = "."
	If Len(aAreaComponent(S_CITY_AREA)) = 0 Then aAreaComponent(S_CITY_AREA) = "."
	If Len(aAreaComponent(S_ZIP_CODE_AREA)) = 0 Then aAreaComponent(S_ZIP_CODE_AREA) = "."
	If Not IsNumeric(aAreaComponent(N_ZONE_ID_AREA)) Then aAreaComponent(N_ZONE_ID_AREA) = -1
	If Not IsNumeric(aAreaComponent(N_ECONOMIC_ZONE_ID_AREA)) Then aAreaComponent(N_ECONOMIC_ZONE_ID_AREA) = 1
	If Not IsNumeric(aAreaComponent(N_PAYMENT_CENTER_ID_AREA)) Then aAreaComponent(N_PAYMENT_CENTER_ID_AREA) = 1
	If aAreaComponent(N_PAYMENT_CENTER_ID_AREA) = -2 Then aAreaComponent(N_PAYMENT_CENTER_ID_AREA) = aAreaComponent(N_ID_AREA)
	If Not IsNumeric(aAreaComponent(N_GENERATING_AREA_ID_AREA)) Then aAreaComponent(N_GENERATING_AREA_ID_AREA) = 1
	If Not IsNumeric(aAreaComponent(N_BRANCH_ID_AREA)) Then aAreaComponent(N_BRANCH_ID_AREA) = 1
	If Not IsNumeric(aAreaComponent(N_SUBBRANCH_ID_AREA)) Then aAreaComponent(N_SUBBRANCH_ID_AREA) = 1
	If Not IsNumeric(aAreaComponent(N_CASHIER_OFFICE_ID_AREA)) Then aAreaComponent(N_CASHIER_OFFICE_ID_AREA) = 1
	If Not IsNumeric(aAreaComponent(N_START_DATE_AREA)) Then aAreaComponent(N_START_DATE_AREA) = 0
	If Not IsNumeric(aAreaComponent(N_END_DATE_AREA)) Then aAreaComponent(N_END_DATE_AREA) = 30000000
	If Not IsNumeric(aAreaComponent(N_FINISH_DATE_AREA)) Then aAreaComponent(N_FINISH_DATE_AREA) = 30000000
	If Not IsNumeric(aAreaComponent(N_JOBS_AREA)) Then aAreaComponent(N_JOBS_AREA) = 0
	If Not IsNumeric(aAreaComponent(N_TOTAL_JOBS_AREA)) Then aAreaComponent(N_JOBS_AREA) = 0
	If Not IsNumeric(aAreaComponent(N_STATUS_ID_AREA)) Then aAreaComponent(N_STATUS_ID_AREA) = -1
	If Not IsNumeric(aAreaComponent(N_ACTIVE_AREA)) Then aAreaComponent(N_ACTIVE_AREA) = 1
	If Not IsNumeric(aAreaComponent(N_PREVIOUS_PARENT_ID_AREA)) Then aAreaComponent(N_PREVIOUS_PARENT_ID_AREA) = aAreaComponent(N_PARENT_ID_AREA)
	If Not IsNumeric(aAreaComponent(N_PREVIOUS_JOBS_AREA)) Then aAreaComponent(N_PREVIOUS_JOBS_AREA) = aAreaComponent(N_JOBS_AREA)
	If Not IsNumeric(aAreaComponent(N_PREVIOUS_TOTAL_JOBS_AREA)) Then aAreaComponent(N_PREVIOUS_TOTAL_JOBS_AREA) = aAreaComponent(N_TOTAL_JOBS_AREA)
	If Not IsNumeric(aAreaComponent(N_PREVIOUS_STATUS_ID_AREA)) Then aAreaComponent(N_PREVIOUS_STATUS_ID_AREA) = aAreaComponent(STATUS_ID_AREA)

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del registro contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AreaComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckAreaInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayArea(oRequest, oADODBConnection, bForExport, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an area
'Inputs:  oRequest, oADODBConnection, bForExport, aAreaComponent
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayArea"
	Dim sNames
	Dim asPath
	Dim lErrorNumber

	Dim aZoneComponent()
	Redim aZoneComponent(N_ZONE_COMPONENT_SIZE)

	If aAreaComponent(N_ID_AREA) <> -1 Then
		lErrorNumber = GetArea(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""AreaFrm"" ID=""AreaFrm"" ACTION=""Areas.asp"" METHOD=""GET"">"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Clave:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(aAreaComponent(S_CODE_AREA)) & "</B></FONT></TD>"
				Response.Write "</TR>"
				If Not B_ISSSTE Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Código:&nbsp;</B></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(aAreaComponent(S_SHORT_NAME_AREA)) & "</B></FONT></TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Nombre:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(aAreaComponent(S_NAME_AREA)) & "</B></FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>UR-CT-AUX:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aAreaComponent(S_URCTAUX_AREA)) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Empresa:&nbsp;</B></FONT></TD>"
					Call GetNameFromTable(oADODBConnection, "Companies", aAreaComponent(N_COMPANY_ID_AREA), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Tipo de área:&nbsp;</B></FONT></TD>"
					Call GetNameFromTable(oADODBConnection, "AreaTypes", aAreaComponent(N_TYPE_ID_AREA), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Ámbito del área:&nbsp;</B></FONT></TD>"
					Call GetNameFromTable(oADODBConnection, "ConfineTypes", aAreaComponent(N_CONFINE_TYPE_ID_AREA), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Tipo de centro de trabajo:&nbsp;</B></FONT></TD>"
					Call GetNameFromTable(oADODBConnection, "CenterTypes", aAreaComponent(N_CENTER_TYPE_ID_AREA), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Subtipo de centro de trabajo:&nbsp;</B></FONT></TD>"
					Call GetNameFromTable(oADODBConnection, "CenterSubtypes", aAreaComponent(N_CENTER_SUBTYPE_ID_AREA), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Nivel de atención:&nbsp;</B></FONT></TD>"
					Call GetNameFromTable(oADODBConnection, "AttentionLevels", aAreaComponent(N_ATTENTION_LEVEL_ID_AREA), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD ALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><B>Dirección:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aAreaComponent(S_ADDRESS_AREA) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Ciudad:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aAreaComponent(S_CITY_AREA) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>C.P.:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aAreaComponent(S_ZIP_CODE_AREA) & "</FONT></TD>"
				Response.Write "</TR>"

				aZoneComponent(N_ID_ZONE) = aAreaComponent(N_ZONE_ID_AREA)
				Call GetNameFromTable(oADODBConnection, "ParentZoneIDs", aAreaComponent(N_ZONE_ID_AREA), "", "", aZoneComponent(N_PARENT_ID_ZONE), sErrorDescription)
				Call GetZonePath(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
				aZoneComponent(S_PATH_ZONE) = aZoneComponent(S_PATH_ZONE) & BuildList("-1", ",", 3)
				asPath = Split(aZoneComponent(S_PATH_ZONE), ",")
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Entidad:&nbsp;</B></FONT></TD>"
					Call GetNameFromTable(oADODBConnection, "FullZones", asPath(2), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Municipio:&nbsp;</B></FONT></TD>"
					Call GetNameFromTable(oADODBConnection, "FullZones", asPath(3), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Población:&nbsp;</B></FONT></TD>"
					Call GetNameFromTable(oADODBConnection, "FullZones", aAreaComponent(N_ZONE_ID_AREA), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"

				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Zona económica:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aAreaComponent(N_ECONOMIC_ZONE_ID_AREA) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Centro de pago:&nbsp;</B></FONT></TD>"
					Call GetNameFromTable(oADODBConnection, "Areas", aAreaComponent(N_PAYMENT_CENTER_ID_AREA), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Área generadora:&nbsp;</B></FONT></TD>"
					Call GetNameFromTable(oADODBConnection, "GeneratingAreas", aAreaComponent(N_GENERATING_AREA_ID_AREA), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				If False Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Rama:&nbsp;</B></FONT></TD>"
						Call GetNameFromTable(oADODBConnection, "Branches", aAreaComponent(N_BRANCH_ID_AREA), "", "", sNames, sErrorDescription)
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Subrama:&nbsp;</B></FONT></TD>"
						Call GetNameFromTable(oADODBConnection, "SubBranches", aAreaComponent(N_SUBBRANCH_ID_AREA), "", "", sNames, sErrorDescription)
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Pagaduría SIPE:&nbsp;</B></FONT></TD>"
					Call GetNameFromTable(oADODBConnection, "CashierOffices", aAreaComponent(N_CASHIER_OFFICE_ID_AREA), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				If (aAreaComponent(N_START_DATE_AREA) > 0) And (aAreaComponent(N_START_DATE_AREA) < 30000000) Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aAreaComponent(N_START_DATE_AREA), -1, -1, -1) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
				If (aAreaComponent(N_END_DATE_AREA) > 0) And (aAreaComponent(N_END_DATE_AREA) < 30000000) Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de término:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aAreaComponent(N_END_DATE_AREA), -1, -1, -1) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
				If (aAreaComponent(N_FINISH_DATE_AREA) > 0) And (aAreaComponent(N_FINISH_DATE_AREA) < 30000000) Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Fecha inhabilitado:&nbsp;</B></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aAreaComponent(N_FINISH_DATE_AREA), -1, -1, -1) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
				If Not B_ISSSTE Then
					If aAreaComponent(N_LEVEL_TYPE_ID_AREA) = 2 Then
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Plazas:&nbsp;</B></FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aAreaComponent(N_JOBS_AREA), 0, True, False, True) & "</FONT></TD>"
						Response.Write "</TR>"
					ElseIf aAreaComponent(N_LEVEL_TYPE_ID_AREA) = 1 Then
						Response.Write "<TR><TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"">Plazas de todos los centros de trabajo: " & FormatNumber(aAreaComponent(N_TOTAL_JOBS_AREA), 0, True, False, True) & "</FONT></TD></TR>"
					End If
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Estatus:&nbsp;</B></FONT></TD>"
						Call GetNameFromTable(oADODBConnection, "StatusAreas", aAreaComponent(N_STATUS_ID_AREA), "", "", sNames, sErrorDescription)
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Activo:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayYesNo(aAreaComponent(N_ACTIVE_AREA), True) & "</FONT></TD>"
				Response.Write "</TR>"
				If (Not bForExport) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Then
					Response.Write "<TR><TD COLSPAN=""2"" ALIGN=""RIGHT"">"
						Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Areas"" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaID"" ID=""AreaIDHdn"" VALUE=""" & aAreaComponent(N_ID_AREA) & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""1"" />"
						If aAreaComponent(B_SEND_TO_IFRAME_AREA) Then
							Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Change"" ID=""ChangeBtn"" VALUE=""Modificar"" CLASS=""Buttons"" onClick=""document.FormsIFrame.location.replace('ShowForms.asp?Action=Areas&AreaID=" & aAreaComponent(N_ID_AREA) & "&Tab=1')"" />"
						Else
							Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Change"" ID=""ChangeBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
						End If
					Response.Write "</TD></TR>"
				End If
			Response.Write "</TABLE>"
		Response.Write "</FORM>"
	End If

	DisplayArea = lErrorNumber
	Err.Clear
End Function

Function DisplayAreaCompact(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an area
'Inputs:  oRequest, oADODBConnection, aAreaComponent
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAreaCompact"
	Dim sNames
	Dim lErrorNumber

	If aAreaComponent(N_ID_AREA) <> -1 Then
		lErrorNumber = GetArea(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<FORM NAME=""AreaFrm"" ID=""AreaFrm"" ACTION=""Areas.asp"" METHOD=""GET"">"
			Response.Write "<TABLE WIDTH=""200"" BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR><TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """ COLSPAN=""5""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>"
				Response.Write "<TR BGCOLOR=""#" & S_SELECTED_BGCOLOR_MENU & """>"
					Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /></TD>"
					Response.Write "<TD COLSPAN=""3""><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(aAreaComponent(S_SHORT_NAME_AREA) & ". " & aAreaComponent(S_NAME_AREA)) & "</B></FONT></TD>"
					Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR><TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """ COLSPAN=""5""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>"
				Response.Write "<TR>"
					Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /></TD>"
					Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""3"" HEIGHT=""1"" /></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Inicio:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aAreaComponent(N_START_DATE_AREA), -1, -1, -1) & "</FONT></TD>"
					Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /></TD>"
				Response.Write "</TR>"
				If aAreaComponent(N_END_DATE_AREA) > 0 Then
					Response.Write "<TR>"
						Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /></TD>"
						Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""3"" HEIGHT=""1"" /></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Término:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aAreaComponent(N_END_DATE_AREA), -1, -1, -1) & "</FONT></TD>"
						Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /></TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR>"
					Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /></TD>"
					Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""3"" HEIGHT=""1"" /></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estatus:&nbsp;</FONT></TD>"
					Call GetNameFromTable(oADODBConnection, "StatusAreas", aAreaComponent(N_STATUS_ID_AREA), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
					Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /></TD>"
					Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""3"" HEIGHT=""1"" /></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Activo:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayYesNo(aAreaComponent(N_ACTIVE_AREA), True) & "</FONT></TD>"
					Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /></TD>"
				Response.Write "</TR>"
				If Not B_ISSSTE Then
					If aAreaComponent(N_LEVEL_TYPE_ID_AREA) = 2 Then
						Response.Write "<TR>"
							Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /></TD>"
							Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""3"" HEIGHT=""1"" /></TD>"
							Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"">Plazas:&nbsp;" & FormatNumber(aAreaComponent(N_JOBS_AREA), 0, True, False, True) & "</FONT></TD>"
							Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /></TD>"
						Response.Write "</TR>"
					ElseIf aAreaComponent(N_LEVEL_TYPE_ID_AREA) = 1 Then
						If (aAreaComponent(N_TOTAL_JOBS_AREA) - aAreaComponent(N_JOBS_AREA)) > 0 Then
							Response.Write "<TR>"
								Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /></TD>"
								Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""3"" HEIGHT=""1"" /></TD>"
								Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"">Plazas en todos los centros de trabajo:&nbsp;" & FormatNumber(aAreaComponent(N_TOTAL_JOBS_AREA) - aAreaComponent(N_JOBS_AREA), 0, True, False, True) & "</FONT></TD>"
								Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /></TD>"
							Response.Write "</TR>"
						End If
					End If
				End If
				If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
					Response.Write "<TR>"
						Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /></TD>"
						Response.Write "<TD COLSPAN=""3"" ALIGN=""RIGHT"">"
							Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""26"" ALIGN=""ABSMIDDLE"" />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaID"" ID=""AreaIDHdn"" VALUE=""" & aAreaComponent(N_ID_AREA) & """ />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""1"" />"
							If aAreaComponent(B_SEND_TO_IFRAME_AREA) Then
								Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Change"" ID=""ChangeBtn"" VALUE=""Modificar"" CLASS=""Buttons"" onClick=""document.FormsIFrame.location.replace('ShowForms.asp?Action=Areas&AreaID=" & aAreaComponent(N_ID_AREA) & "&Tab=1')"" />"
							Else
								Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Change"" ID=""ChangeBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
							End If
						Response.Write "&nbsp;</TD>"
						Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /></TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR><TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """ COLSPAN=""5""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>"
			Response.Write "</TABLE>"
		Response.Write "</FORM>"
	End If

	DisplayAreaCompact = lErrorNumber
	Err.Clear
End Function

Function DisplayAreaHistoryList(oRequest, oADODBConnection, bForExport, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To display the changes to the information for the
'		  given area
'Inputs:  oRequest, oADODBConnection, bForExport, aAreaComponent
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAreaHistoryList"
	Dim oRecordset
	Dim sNames
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	bComponentInitialized = aAreaComponent(B_COMPONENT_INITIALIZED_AREA)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAreaComponent(oRequest, aAreaComponent)
	End If

	If aAreaComponent(N_ID_AREA) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AreaComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreasHistoryList.*, UserName, UserLastName, StatusName From AreasHistoryList, Users, StatusAreas Where (AreasHistoryList.UserID=Users.UserID) And (AreasHistoryList.StatusID=StatusAreas.StatusID) And (AreaID=" & aAreaComponent(N_ID_AREA) & ") Order By AreaDate", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Response.Write "<TABLE WIDTH=""950"" BORDER="""
			If bForExport Then
				Response.Write "1"
			Else
				Response.Write "0"
			End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If B_ISSSTE Then
					asColumnsTitles = Split("Fecha,Responsable del cambio,Nombre,Área a la que pertenece,Estatus", ",", -1, vbBinaryCompare)
					asCellWidths = Split("200,200,200,200,150", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("Fecha,Responsable del cambio,Nombre,Área a la que pertenece,Plazas,Total de Plazas,Estatus", ",", -1, vbBinaryCompare)
					asCellWidths = Split("200,150,150,150,100,100,100", ",", -1, vbBinaryCompare)
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

				asCellAlignments = Split(",,,,RIGHT,RIGHT,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					If CLng(oRecordset.Fields("AreaDate").Value) = 0 Then
						sRowContents = "---"
					Else
						sRowContents = DisplayDateFromSerialNumber(CLng(oRecordset.Fields("AreaDate").Value), -1, -1, -1)
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("UserName").Value) & " " & CStr(oRecordset.Fields("UserLastName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value))
					Call GetNameFromTable(oADODBConnection, "Areas", CLng(oRecordset.Fields("ParentID").Value), "", "", sNames, sErrorDescription)
					If Len(sNames) = 0 Then sNames = "Ninguna"
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(sNames)
					If Request.Cookies("SIAP_SectionID") <> 3 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CLng(oRecordset.Fields("JobsInArea").Value), 0, True, False, True)
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CLng(oRecordset.Fields("TotalJobs").Value), 0, True, False, True)
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value))

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
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	DisplayAreaHistoryList = lErrorNumber
	Err.Clear
End Function

Function DisplayAreaPositionsHistoryList(oRequest, oADODBConnection, bForExport, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To display the changes to the information for the
'		  given area
'Inputs:  oRequest, oADODBConnection, bForExport, aAreaComponent
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAreaHistoryList"
	Dim oRecordset
	Dim sNames
	Dim lCurrentDate
	Dim lTotalJobs
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	bComponentInitialized = aAreaComponent(B_COMPONENT_INITIALIZED_AREA)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAreaComponent(oRequest, aAreaComponent)
	End If

	If aAreaComponent(N_ID_AREA) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "AreaComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreasPositionsHistoryList.*, StartUsers.UserName As StartUserName, StartUsers.UserLastName As StartUserLastName, EndUsers.UserID As EndUserID, EndUsers.UserName As EndUserName, EndUsers.UserLastName As EndUserLastName, PositionName From AreasPositionsHistoryList, Users As StartUsers, Users As EndUsers, Positions Where (AreasPositionsHistoryList.StartUserID=StartUsers.UserID) And (AreasPositionsHistoryList.EndUserID=EndUsers.UserID) And (AreasPositionsHistoryList.PositionID=Positions.PositionID) And (AreaID=" & aAreaComponent(N_ID_AREA) & ") Order By AreasPositionsHistoryList.StartDate, AreasPositionsHistoryList.EndDate, PositionName", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Response.Write "<TABLE WIDTH=""950"" BORDER="""
			If bForExport Then
				Response.Write "1"
			Else
				Response.Write "0"
			End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If B_ISSSTE Then
					asColumnsTitles = Split("Fecha alta,Fecha baja,Responsable del alta,Responsable de la baja,Puesto", ",", -1, vbBinaryCompare)
					asCellWidths = Split("200,200,200,200,150", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("Fecha alta,Fecha baja,Responsable del alta,Responsable de la baja,Puesto,Plazas", ",", -1, vbBinaryCompare)
					asCellWidths = Split("200,200,150,150,150,100", ",", -1, vbBinaryCompare)
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

				lCurrentDate = 0
				lTotalJobs = 0
				asCellAlignments = Split(",,,,,RIGHT,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					If lCurrentDate <> CLng(oRecordset.Fields("StartDate").Value) Then
						If lCurrentDate > 0 Then
							sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & "<B>Total</B>" & TABLE_SEPARATOR & FormatNumber(lTotalJobs, 0, True, False, True)
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
							lTotalJobs = 0
						End If
						sRowContents = DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
						If CLng(oRecordset.Fields("EndDate").Value) <> 30000000 Then
							sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""5"" />" & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & "<SPAN COLS=""5"" />Al día de hoy"
						End If
						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If
						lCurrentDate = CLng(oRecordset.Fields("StartDate").Value)
					End If
					sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StartUserName").Value) & " " & CStr(oRecordset.Fields("StartUserLastName").Value))
					If CLng(oRecordset.Fields("EndUserID").Value) > -1 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EndUserName").Value) & " " & CStr(oRecordset.Fields("EndUserLastName").Value))
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "---"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value))
					If Request.Cookies("SIAP_SectionID") <> 3 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CLng(oRecordset.Fields("JobsInArea").Value), 0, True, False, True)
						lTotalJobs = lTotalJobs + CLng(oRecordset.Fields("JobsInArea").Value)
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If

					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				If Request.Cookies("SIAP_SectionID") <> 3 Then
					sRowContents = TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & TABLE_SEPARATOR & "<B>Total</B>" & TABLE_SEPARATOR & FormatNumber(lTotalJobs, 0, True, False, True)
				End If
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Response.Write "</TABLE>" & vbNewLine
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	DisplayAreaPositionsHistoryList = lErrorNumber
	Err.Clear
End Function

Function DisplayAreaForm(oRequest, oADODBConnection, sAction, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an area from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aAreaComponent
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAreaForm"
	Dim asPath
	Dim lErrorNumber
	Dim sFilter

	sFilter = ""
	Dim aZoneComponent()
	Redim aZoneComponent(N_ZONE_COMPONENT_SIZE)

	If aAreaComponent(N_ID_AREA) <> -1 Then
		lErrorNumber = GetArea(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "var aCenterSubtypeID = new Array("
				Call GenerateJavaScriptArrayFromQuery(oADODBConnection, "CenterSubtypes", "CenterSubtypeID", "CenterTypeID, CenterSubtypeShortName, CenterSubtypeName", "(EndDate=30000000) And (Active=1)", "CenterTypeID, CenterSubtypeID", sErrorDescription)
			Response.Write "['-2', '-2', '', '']);" & vbNewLine

			Response.Write "function CheckAreaFields(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If Len(oRequest("Delete").Item) > 0 Then Response.Write "return true;" & vbNewLine
					Response.Write "if (oForm.AreaCode.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de introducir el código del registro.');" & vbNewLine
						Response.Write "oForm.AreaCode.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					If Not B_ISSSTE Then
						Response.Write "if (oForm.AreaShortName.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir la clave del registro.');" & vbNewLine
							Response.Write "oForm.AreaShortName.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					End If
					Response.Write "if (oForm.AreaName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de introducir el nombre del registro.');" & vbNewLine
						Response.Write "oForm.AreaName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (! CheckIntegerValue(oForm.EconomicZoneID, 'la zona económica', N_BOTH_FLAG, N_CLOSED_FLAG, 2, 3))" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "if (! CheckIntegerValue(oForm.JobsInArea, 'las plazas del área', N_NO_RANK_FLAG, N_OPEN_FLAG, 0, 0))" & vbNewLine
						Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine

                If aAreaComponent(N_LEVEL_AREA) > 0 Then
                    If aAreaComponent(N_PARENT_ZONE) <> -1 Then
                        'Response.Write "alert('Entidad: ' + oForm.ZoneID01.value + 'SoloCentroPago: ' + oForm.OnlyPayment[0].checked + ' ' + oForm.OnlyPayment[1].checked);" & vbNewLine
					    Response.Write "if (oForm.OnlyPayment[1].checked == true && oForm.ZoneID01.value != " & aAreaComponent(N_PARENT_ZONE) & ") {" & vbNewLine
						    Response.Write "alert('Favor de seleccionar una Entidad que corresponda con la del Area generadora.');" & vbNewLine
						    Response.Write "oForm.ZoneID01.focus();" & vbNewLine
						    Response.Write "return false;" & vbNewLine
					    Response.Write "}" & vbNewLine
                    End If
                End If

				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckAreaFields" & vbNewLine

			Response.Write "function UpdateSubtypes(sTypeID) {" & vbNewLine
				Response.Write "oForm = document.AreaFrm;" & vbNewLine

				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "RemoveAllItemsFromList(null, oForm.CenterSubtypeID);" & vbNewLine
					Response.Write "for (var i=0; i<aCenterSubtypeID.length; i++)" & vbNewLine
						Response.Write "if (aCenterSubtypeID[i][1] == sTypeID)" & vbNewLine
							Response.Write "AddItemToList(aCenterSubtypeID[i][2] + '. ' + aCenterSubtypeID[i][3], aCenterSubtypeID[i][0], null, oForm.CenterSubtypeID);" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of UpdateSubtypes" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine

		Response.Write "<FORM NAME=""AreaFrm"" ID=""AreaFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckAreaFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Areas"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaID"" ID=""AreaIDHdn"" VALUE=""" & aAreaComponent(N_ID_AREA) & """ />"
			If Len(oRequest("Change").Item) = 0 Then
                Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ParentID"" ID=""ParentIDHdn"" VALUE=""" & aAreaComponent(N_PARENT_ID_AREA) & """ />"
			    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaPath"" ID=""AreaPathHdn"" VALUE=""" & aAreaComponent(S_PATH_AREA) & """ />"
            End If
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TotalJobs"" ID=""TotalJobsHdn"" VALUE=""" & aAreaComponent(N_TOTAL_JOBS_AREA) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PreviousParentID"" ID=""PreviousParentIDHdn"" VALUE=""" & aAreaComponent(N_PREVIOUS_PARENT_ID_AREA) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PreviousJobsInArea"" ID=""PreviousJobsInAreaHdn"" VALUE=""" & aAreaComponent(N_PREVIOUS_JOBS_AREA) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PreviousTotalJobs"" ID=""PreviousTotalJobsHdn"" VALUE=""" & aAreaComponent(N_PREVIOUS_TOTAL_JOBS_AREA) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PreviousStatusID"" ID=""PreviousStatusIDHdn"" VALUE=""" & aAreaComponent(N_PREVIOUS_STATUS_ID_AREA) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaymentCenters"" ID=""PaymentCentersHdn"" VALUE=""" & oRequest("PaymentCenters").Item & """ />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
                If Len(oRequest("Change").Item) > 0 Then
				    If aAreaComponent(N_LEVEL_AREA) > 0 Then
				    	Response.Write "<TR>"
				    		Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Área:&nbsp;</FONT></TD>"
				    		Response.Write "<TD><SELECT NAME=""ParentID"" ID=""ParentIDCmb"" SIZE=""1"" CLASS=""Lists"">"
				    			Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaShortName, AreaName", "(Active=1) And (Areas.ParentID=-1)", "AreaShortName", aAreaComponent(N_PARENT_ID_AREA), "Ninguno;;;-1", sErrorDescription)
				    		Response.Write "</SELECT></TD>"
				    	Response.Write "</TR>"
				    End If
                End If
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clave:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""AreaCode"" ID=""AreaCodeTxt"" SIZE=""5"" MAXLENGTH=""5"" VALUE=""" & CleanStringForHTML(aAreaComponent(S_CODE_AREA)) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				If B_ISSSTE Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Código:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""AreaShortName"" ID=""AreaShortNameTxt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & CleanStringForHTML(aAreaComponent(S_SHORT_NAME_AREA)) & """ CLASS=""TextFields"" /></TD>"
					Response.Write "</TR>"
				Else
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaShortName"" ID=""AreaShortNameHdn"" VALUE=""" & CleanStringForHTML(aAreaComponent(S_SHORT_NAME_AREA)) & """ />"
				End If
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""AreaName"" ID=""AreaNameTxt"" SIZE=""60"" MAXLENGTH=""100"" VALUE=""" & CleanStringForHTML(aAreaComponent(S_NAME_AREA)) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				If aAreaComponent(N_LEVEL_AREA) > 0 Then
					If aAreaComponent(N_ZONE_ID_AREA) = -1 Then
						aZoneComponent(N_PARENT_ID_ZONE) = aAreaComponent(N_PARENT_ID_AREA)
						'Call GetNameFromTable(oADODBConnection, "ParentZoneIDs", aAreaComponent(N_ZONE_ID_AREA), "", "", aZoneComponent(N_PARENT_ID_ZONE), sErrorDescription)
						Call GetZonePath(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
						aZoneComponent(S_PATH_ZONE) = aZoneComponent(S_PATH_ZONE) & BuildList("-1", ",", 3)
						asPath = Split(aZoneComponent(S_PATH_ZONE), ",")
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entidad:&nbsp;</FONT></TD>"
							Response.Write "<TD><SELECT NAME=""ZoneID01"" ID=""ZoneID01Cmb"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.value == '-1') {document.HierarchyMenu02IFrame.location.href = 'HierarchyMenu.asp';} else {document.HierarchyMenu02IFrame.location.href = 'HierarchyMenu.asp?Action=Zones&TargetField=HierarchyMenu03IFrame&SecondTargetField=AreaFrm.ZoneID&ParentID=' + this.value + '&PathLevel=2';}"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneCode, ZoneName", "(Active=1) And (ParentID=-1) And (ZoneID>-1)", "ZoneCode", asPath(2), "Ninguna;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Municipio:&nbsp;</FONT></TD>"
							Response.Write "<TD><IFRAME SRC=""HierarchyMenu.asp?Action=Zones&TargetField=HierarchyMenu03IFrame&SecondTargetField=AreaFrm.ZoneID&ParentID=" & asPath(2) & "&ZoneID=" & asPath(3) & "&PathLevel=2"" NAME=""HierarchyMenu02IFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""26""></IFRAME></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Población:&nbsp;</FONT></TD>"
							Response.Write "<TD><IFRAME SRC=""HierarchyMenu.asp?Action=Zones&TargetField=AreaFrm.ZoneID&ParentID=" & asPath(3) & "&ZoneID=" & aAreaComponent(N_ZONE_ID_AREA) & "&PathLevel=3"" NAME=""HierarchyMenu03IFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""24""></IFRAME></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Zona económica:&nbsp;</FONT></TD>"
							Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EconomicZoneID"" ID=""EconomicZoneIDTxt"" SIZE=""1"" MAXLENGTH=""1"" VALUE=""" & aAreaComponent(N_ECONOMIC_ZONE_ID_AREA) & """ CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
					Else
						aZoneComponent(N_ID_ZONE) = aAreaComponent(N_ZONE_ID_AREA)
						Call GetNameFromTable(oADODBConnection, "ParentZoneIDs", aAreaComponent(N_ZONE_ID_AREA), "", "", aZoneComponent(N_PARENT_ID_ZONE), sErrorDescription)
						Call GetZonePath(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
						aZoneComponent(S_PATH_ZONE) = aZoneComponent(S_PATH_ZONE) & BuildList("-1", ",", 3)
						asPath = Split(aZoneComponent(S_PATH_ZONE), ",")
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entidad:&nbsp;</FONT></TD>"
							Response.Write "<TD><SELECT NAME=""ZoneID01"" ID=""ZoneID01Cmb"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.value == '-1') {document.HierarchyMenu02IFrame.location.href = 'HierarchyMenu.asp';} else {document.HierarchyMenu02IFrame.location.href = 'HierarchyMenu.asp?Action=Zones&TargetField=HierarchyMenu03IFrame&SecondTargetField=AreaFrm.ZoneID&ParentID=' + this.value + '&PathLevel=2';}"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneCode, ZoneName", "(Active=1) And (ParentID=-1) And (ZoneID>-1)", "ZoneCode", asPath(2), "Ninguna;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Municipio:&nbsp;</FONT></TD>"
							Response.Write "<TD><IFRAME SRC=""HierarchyMenu.asp?Action=Zones&TargetField=HierarchyMenu03IFrame&SecondTargetField=AreaFrm.ZoneID&ParentID=" & asPath(2) & "&ZoneID=" & asPath(3) & "&PathLevel=2"" NAME=""HierarchyMenu02IFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""26""></IFRAME></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Población:&nbsp;</FONT></TD>"
							Response.Write "<TD><IFRAME SRC=""HierarchyMenu.asp?Action=Zones&TargetField=AreaFrm.ZoneID&ParentID=" & asPath(3) & "&ZoneID=" & aAreaComponent(N_ZONE_ID_AREA) & "&PathLevel=3"" NAME=""HierarchyMenu03IFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""24""></IFRAME></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Zona económica:&nbsp;</FONT></TD>"
							Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EconomicZoneID"" ID=""EconomicZoneIDTxt"" SIZE=""1"" MAXLENGTH=""1"" VALUE=""" & aAreaComponent(N_ECONOMIC_ZONE_ID_AREA) & """ CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
					End If
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Empresa:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""CompanyID"" ID=""CompanyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Companies", "CompanyID", "CompanyShortName, CompanyName", "(CompanyID>-1) And (Active=1)", "CompanyShortName", aAreaComponent(N_COMPANY_ID_AREA), "Ninguna;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de área:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""AreaTypeID"" ID=""AreaTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "AreaTypes", "AreaTypeID", "AreaTypeShortName, AreaTypeName", "(Active=1)", "AreaTypeShortName", aAreaComponent(N_TYPE_ID_AREA), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Ámbito del área:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""ConfineTypeID"" ID=""ConfineTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "ConfineTypes", "ConfineTypeID", "ConfineTypeShortName, ConfineTypeName", "(ConfineTypeID>-1) And (Active=1)", "ConfineTypeShortName", aAreaComponent(N_CONFINE_TYPE_ID_AREA), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"
                Else
                    If aAreaComponent(N_ZONE_ID_AREA) = -1 Then
						aZoneComponent(N_PARENT_ID_ZONE) = aAreaComponent(N_PARENT_ID_AREA)
						'Call GetNameFromTable(oADODBConnection, "ParentZoneIDs", aAreaComponent(N_ZONE_ID_AREA), "", "", aZoneComponent(N_PARENT_ID_ZONE), sErrorDescription)
						Call GetZonePath(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
						aZoneComponent(S_PATH_ZONE) = aZoneComponent(S_PATH_ZONE) & BuildList("-1", ",", 3)
						asPath = Split(aZoneComponent(S_PATH_ZONE), ",")
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entidad:&nbsp;</FONT></TD>"
							Response.Write "<TD><SELECT NAME=""ZoneID"" ID=""ZoneIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneCode, ZoneName", "(Active=1) And (ParentID=-1) And (ZoneID>-1)", "ZoneCode", asPath(2), "Ninguna;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
                    Else
						aZoneComponent(N_ID_ZONE) = aAreaComponent(N_ZONE_ID_AREA)
						Call GetNameFromTable(oADODBConnection, "ParentZoneIDs", aAreaComponent(N_ZONE_ID_AREA), "", "", aZoneComponent(N_PARENT_ID_ZONE), sErrorDescription)
						Call GetZonePath(oRequest, oADODBConnection, aZoneComponent, sErrorDescription)
						aZoneComponent(S_PATH_ZONE) = aZoneComponent(S_PATH_ZONE) & BuildList("-1", ",", 3)
						asPath = Split(aZoneComponent(S_PATH_ZONE), ",")
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entidad:&nbsp;</FONT></TD>"
							Response.Write "<TD><SELECT NAME=""ZoneID"" ID=""ZoneIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneCode, ZoneName", "(Active=1) And (ParentID=-1) And (ZoneID>-1)", "ZoneCode", asPath(2), "Ninguna;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
                    End If
				End If
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">UR-CT-AUX:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""URCTAUX"" ID=""URCTAUXTxt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & CleanStringForHTML(aAreaComponent(S_URCTAUX_AREA)) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				If Not B_ISSSTE Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de nivel del área:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""AreaLevelTypeID"" ID=""AreaLevelTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "AreaLevelTypes", "AreaLevelTypeID", "AreaLevelTypeName", "(AreaLevelTypeID>-1) And (Active=1)", "AreaLevelTypeID", aAreaComponent(N_LEVEL_TYPE_ID_AREA), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"
				Else
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaLevelTypeID"" ID=""AreaLevelTypeIDHdn"" VALUE=""" & aAreaComponent(N_LEVEL_TYPE_ID_AREA) & """ />"
				End If
				If aAreaComponent(N_LEVEL_AREA) > 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de centro de trabajo:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""CenterTypeID"" ID=""CenterTypeIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""UpdateSubtypes(this.value)"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "CenterTypes", "CenterTypeID", "CenterTypeShortName, CenterTypeName", "(Active=1)", "CenterTypeShortName", aAreaComponent(N_CENTER_TYPE_ID_AREA), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Subtipo de centro de trabajo:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""CenterSubtypeID"" ID=""CenterSubtypeIDCmb"" SIZE=""1"" CLASS=""Lists""></SELECT></TD>"
					Response.Write "</TR>"
					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
						Response.Write "UpdateSubtypes(document.AreaFrm.CenterTypeID.value);" & vbNewLine
						Response.Write "SelectItemByValue('" & aAreaComponent(N_CENTER_SUBTYPE_ID_AREA) & "', false, document.AreaFrm.CenterSubtypeID);" & vbNewLine
					Response.Write "//--></SCRIPT>" & vbNewLine
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nivel de atención:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""AttentionLevelID"" ID=""AttentionLevelIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "AttentionLevels", "AttentionLevelID", "AttentionLevelShortName, AttentionLevelName", "(Active=1)", "AttentionLevelShortName", aAreaComponent(N_ATTENTION_LEVEL_ID_AREA), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"">Dirección:<BR />"
							Response.Write "<TEXTAREA NAME=""AreaAddress"" ID=""AreaAddressTxtArea"" ROWS=""5"" COLS=""60"" CLASS=""TextFields"">" & aAreaComponent(S_ADDRESS_AREA) & "</TEXTAREA>"
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Ciudad:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""AreaCity"" ID=""AreaCityTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & aAreaComponent(S_CITY_AREA) & """ CLASS=""TextFields"" /></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">C.P.:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""AreaZip"" ID=""AreaZipTxt"" SIZE=""5"" MAXLENGTH=""10"" VALUE=""" & aAreaComponent(S_ZIP_CODE_AREA) & """ CLASS=""TextFields"" /></TD>"
					Response.Write "</TR>"

					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ZoneID"" ID=""ZoneIDHdn"" VALUE=""" & aAreaComponent(N_ZONE_ID_AREA) & """ CLASS=""TextFields"" />"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Centro de pago:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""PaymentCenterID"" ID=""PaymentCenterIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write "<OPTION VALUE=""-2"">El centro de trabajo es el centro de pago</OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(Active=1) And (ParentID>0)", "AreaCode", aAreaComponent(N_PAYMENT_CENTER_ID_AREA), "Ninguna;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"
					If True Then
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BranchID"" ID=""BranchIDHdn"" VALUE=""" & aAreaComponent(N_BRANCH_ID_AREA) & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SubBranchID"" ID=""SubBranchIDHdn"" VALUE=""" & aAreaComponent(N_SUBBRANCH_ID_AREA) & """ />"
					Else
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Rama:&nbsp;</FONT></TD>"
							Response.Write "<TD><SELECT NAME=""BranchID"" ID=""BranchIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Branches", "BranchID", "BranchShortName, BranchName", "(Active=1)", "BranchShortName", aAreaComponent(N_BRANCH_ID_AREA), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Subrama:&nbsp;</FONT></TD>"
							Response.Write "<TD><SELECT NAME=""SubBranchID"" ID=""SubBranchIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SubBranches", "SubBranchID", "SubBranchShortName, SubBranchName", "(Active=1)", "SubBranchShortName", aAreaComponent(N_SUBBRANCH_ID_AREA), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
					End If
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Pagaduría SIPE:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""CashierOfficeID"" ID=""CashierOfficeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "CashierOffices", "CashierOfficeID", "CashierOfficeShortName, CashierOfficeName", "(Active=1)", "CashierOfficeShortName", aAreaComponent(N_CASHIER_OFFICE_ID_AREA), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
						Response.Write "<TD>" & DisplayDateCombosUsingSerial(aAreaComponent(N_START_DATE_AREA), "Start", N_START_YEAR, Year(Date()), True, False) & "</TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de término:&nbsp;</FONT></TD>"
						Response.Write "<TD>" & DisplayDateCombosUsingSerial(aAreaComponent(N_END_DATE_AREA), "End", N_START_YEAR, Year(Date()), True, True) & "</TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha inhabilitado:&nbsp;</FONT></TD>"
						Response.Write "<TD>" & DisplayDateCombosUsingSerial(aAreaComponent(N_FINISH_DATE_AREA), "Finish", N_START_YEAR, Year(Date()), True, True) & "</TD>"
					Response.Write "</TR>"
					If Not B_ISSSTE Then
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Plazas:&nbsp;</FONT></TD>"
							Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""JobsInArea"" ID=""JobsInAreaTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & aAreaComponent(N_JOBS_AREA) & """ CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
						If aAreaComponent(N_ID_AREA) > -1 Then Response.Write "<TR><TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"">Total de plazas de esta área y sus subáreas: " & FormatNumber(aAreaComponent(N_TOTAL_JOBS_AREA), 0, True, False, True) & "</FONT></TD></TR>"
					Else
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobsInArea"" ID=""JobsInAreaHdn"" VALUE=""" & aAreaComponent(N_JOBS_AREA) & """ />"
					End If
					If Not B_ISSSTE Then
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estatus:&nbsp;</FONT></TD>"
							Response.Write "<TD><SELECT NAME=""StatusID"" ID=""StatusIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusAreas", "StatusID", "StatusName", "(StatusID>-1) And (Active=1)", "StatusName", aAreaComponent(N_STATUS_ID_AREA), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Response.Write "</TR>"
					Else
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StatusID"" ID=""StatusIDHdn"" VALUE=""" & aAreaComponent(N_STATUS_ID_AREA) & """ />"
					End If
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Activo:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" ID=""ActiveRd"" VALUE=""1"""
								If aAreaComponent(N_ACTIVE_AREA) = 1 Then Response.Write " CHECKED=""1"""
							Response.Write " />Sí&nbsp;&nbsp;&nbsp;"
							Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" ID=""ActiveRd"" VALUE=""0"""
								If aAreaComponent(N_ACTIVE_AREA) = 0 Then Response.Write " CHECKED=""0"""
							Response.Write " />No&nbsp;&nbsp;&nbsp;"
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Aplica sólo como centro de pago:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write "<INPUT TYPE=""RADIO"" NAME=""OnlyPayment"" ID=""OnlyPaymentRd"" VALUE=""1"""
								If aAreaComponent(N_ONLY_PAYMENT) = 1 Then Response.Write " CHECKED=""1"""
							Response.Write " />Sí&nbsp;&nbsp;&nbsp;"
							Response.Write "<INPUT TYPE=""RADIO"" NAME=""OnlyPayment"" ID=""OnlyPaymentRd"" VALUE=""0"""
								If aAreaComponent(N_ONLY_PAYMENT) = 0 Then Response.Write " CHECKED=""0"""
							Response.Write " />No&nbsp;&nbsp;&nbsp;"
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
				Else
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CenterTypeID"" ID=""CenterTypeIDHdn"" VALUE=""" & aAreaComponent(N_CENTER_TYPE_ID_AREA) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CenterSubtypeID"" ID=""CenterSubtypeIDHdn"" VALUE=""" & aAreaComponent(N_CENTER_SUBTYPE_ID_AREA) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AttentionLevelID"" ID=""AttentionLevelIDHdn"" VALUE=""" & aAreaComponent(N_ATTENTION_LEVEL_ID_AREA) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaAddress"" ID=""AreaAddressHdn"" VALUE=""" & aAreaComponent(S_ADDRESS_AREA) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaCity"" ID=""AreaCityHdn"" VALUE=""" & aAreaComponent(S_CITY_AREA) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaZip"" ID=""AreaZipHdn"" VALUE=""" & aAreaComponent(S_ZIP_CODE_AREA) & """ />"
					'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ZoneID"" ID=""ZoneIDHdn"" VALUE=""" & aAreaComponent(N_ZONE_ID_AREA) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EconomicZoneID"" ID=""EconomicZoneIDHdn"" VALUE=""" & aAreaComponent(N_ECONOMIC_ZONE_ID_AREA) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaymentCenterID"" ID=""PaymentCenterIDHdn"" VALUE=""" & aAreaComponent(N_PAYMENT_CENTER_ID_AREA) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GeneratingAreaID"" ID=""GeneratingAreaIDHdn"" VALUE=""" & aAreaComponent(N_GENERATING_AREA_ID_AREA) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BranchID"" ID=""BranchIDHdn"" VALUE=""" & aAreaComponent(N_BRANCH_ID_AREA) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SubBranchID"" ID=""SubBranchIDHdn"" VALUE=""" & aAreaComponent(N_SUBBRANCH_ID_AREA) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CashierOfficeID"" ID=""CashierOfficeIDHdn"" VALUE=""" & aAreaComponent(N_CASHIER_OFFICE_ID_AREA) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate"" ID=""StartDateHdn"" VALUE=""" & aAreaComponent(N_START_DATE_AREA) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EndDate"" ID=""EndDateHdn"" VALUE=""" & aAreaComponent(N_END_DATE_AREA) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FinishDate"" ID=""FinishDateHdn"" VALUE=""" & aAreaComponent(N_FINISH_DATE_AREA) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobsInArea"" ID=""JobsInAreaHdn"" VALUE=""" & aAreaComponent(N_JOBS_AREA) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StatusID"" ID=""StatusIDHdn"" VALUE=""" & aAreaComponent(N_STATUS_ID_AREA) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Active"" ID=""ActiveHdn"" VALUE=""" & aAreaComponent(N_ACTIVE_AREA) & """ />"
				End If
			Response.Write "</TABLE><BR />"

			If aAreaComponent(N_ID_AREA) = -1 Then
				If Len(oRequest("ApplyFilter").Item) = 0 Then
					If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
				End If
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveAreaWngDiv']); AreaFrm.Remove.focus()"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			If Len(oRequest("ApplyFilter").Item) = 0 Then
				If Len(oRequest("ParentID").Item) > 0 Then
					If Len(oRequest("AreaID").Item) > 0 Then
						Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?ParentID=" & oRequest("ParentID").Item & "'"" />"
					Else
						Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?ParentID=" & aAreaComponent(N_PARENT_ID2_AREA) & "'"" />"
					End If
				Else
					Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?ParentID=-1'"" />"
				End If
			Else
				If aAreaComponent(N_ID_AREA) <> -1 Then
					sFilter = "ApplyFilter=1&ZoneID=" & CStr(oRequest("ZoneID").Item) & "&AreaCode=" & CStr(oRequest("AreaCode").Item) & "&AreaShortName=" & CStr(oRequest("AreaShortName").Item) & "&CenterTypeID=" & CStr(oRequest("CenterTypeID").Item)
					Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?" & sFilter & "'"" />"
				Else
					Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?ParentID=-1'"" />"
				End If
			End If

			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveAreaWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayAreaForm = lErrorNumber
	Err.Clear
End Function

Function DisplayAreaPositionsForm(oRequest, oADODBConnection, sAction, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an area from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aAreaComponent
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAreaPositionsForm"
	Dim oRecordset
	Dim lErrorNumber

	If aAreaComponent(N_ID_AREA) <> -1 Then
		lErrorNumber = GetArea(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckAreaFields(oForm) {" & vbNewLine
				Response.Write "var lJobs = 0;" & vbNewLine

				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "for (var i=0; i<oForm.JobsForPosition.options.length; i++)" & vbNewLine
						Response.Write "lJobs += parseInt(oForm.JobsForPosition.options[i].value);" & vbNewLine

					Response.Write "if (lJobs != parseInt(oForm.JobsInArea.value)) {" & vbNewLine
						Response.Write "alert('La cantidad de plazas especificadas (' + lJobs + ') no concuerda con las ' + oForm.JobsInArea.value + ' plazas definidas para el área.');" & vbNewLine
						Response.Write "oForm.PositionID.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine

				Response.Write "SelectAllItemsFromList(oForm.PositionID);" & vbNewLine
				Response.Write "SelectAllItemsFromList(oForm.JobsForPosition);" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckAreaFields" & vbNewLine

			Response.Write "function AddJobsToPosition() {" & vbNewLine
				Response.Write "var oForm = document.AreaFrm;" & vbNewLine

				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "if (CheckIntegerValue(oForm.TempJobs, 'las plazas para el puesto', N_MINIMUM_ONLY_FLAG, N_OPEN_FLAG, 0, 0)) {" & vbNewLine
						Response.Write "SelectListItemByValue(oForm.TempPositionID.value, false, oForm.PositionIDCmb);" & vbNewLine
						Response.Write "SelectSameItems(oForm.PositionIDCmb, document.AreaFrm.JobsForPosition);" & vbNewLine
						Response.Write "RemoveSelectedItemsFromList(null, oForm.PositionIDCmb);" & vbNewLine
						Response.Write "RemoveSelectedItemsFromList(null, oForm.JobsForPosition);" & vbNewLine

						Response.Write "AddItemToList(oForm.TempJobs.value + ': ' + GetSelectedText(oForm.TempPositionID), oForm.TempPositionID.value, null, oForm.PositionIDCmb);" & vbNewLine
						Response.Write "AddItemToList(oForm.TempJobs.value, oForm.TempJobs.value, null, oForm.JobsForPosition);" & vbNewLine
					Response.Write "}" & vbNewLine

					Response.Write "oForm.TempJobs.value = '';" & vbNewLine
					Response.Write "oForm.TempPositionID.focus();" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of AddJobsToPosition" & vbNewLine

			Response.Write "function RemovePosition() {" & vbNewLine
				Response.Write "var oForm = document.AreaFrm;" & vbNewLine

				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "RemoveSelectedItemsFromList(null, oForm.PositionIDCmb);" & vbNewLine
					Response.Write "RemoveSelectedItemsFromList(null, oForm.JobsForPosition);" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of RemovePosition" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""AreaFrm"" ID=""AreaFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckAreaFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""AreaPositions"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Tab"" ID=""TabHdn"" VALUE=""" & oRequest("Tab").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaID"" ID=""AreaIDHdn"" VALUE=""" & aAreaComponent(N_ID_AREA) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobsInArea"" ID=""JobsInAreaHdn"" VALUE=""" & aAreaComponent(N_JOBS_AREA) & """ />"

			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Esta área tiene definidas " & FormatNumber(aAreaComponent(N_JOBS_AREA), 0, True, False, True) & " plazas.</B><BR />Defina qué puestos integran dichas plazas.<BR /><BR /></FONT>"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write "Puestos:&nbsp;"
						Response.Write "<SELECT NAME=""TempPositionID"" ID=""TempPositionIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID", "PositionShortName, PositionName", "(PositionID>-1) And (Active=1)", "PositionShortName", "", "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT><BR />"

						Response.Write "Plazas:&nbsp;"
						Response.Write "<INPUT TYPE=""TEXT"" NAME=""TempJobs"" ID=""TempJobsHdn"" SIZE=""6"" MAXLENGTH=""6"" VALUE="""" CLASS=""TextFields"" />"
					Response.Write "</FONT></TD>"
					Response.Write "<TD>&nbsp;</TD>"
					Response.Write "<TD VALIGN=""TOP"">"
						Response.Write "<A HREF=""javascript: AddJobsToPosition();""><IMG SRC=""Images/BtnCrclAdd.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Agregar"" BORDER=""0"" /></A>"
						Response.Write "<BR /><BR />"
						Response.Write "<A HREF=""javascript: RemovePosition();""><IMG SRC=""Images/BtnCrclDelete.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Agregar"" BORDER=""0"" /></A>"
					Response.Write "</TD>"
					Response.Write "<TD>&nbsp;</TD>"
					Response.Write "<TD VALIGN=""TOP"">"
						sErrorDescription = "No se pudo obtener la información del registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreasPositionsLKP.*, PositionShortName, PositionName From AreasPositionsLKP, Positions Where (AreasPositionsLKP.PositionID=Positions.PositionID) And (AreaID=" & aAreaComponent(N_ID_AREA) & ") Order By PositionShortName", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Response.Write "<SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""10"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItems(this, document.AreaFrm.JobsForPosition)"">"
							If lErrorNumber = 0 Then
								Do While Not oRecordset.EOF
									Response.Write "<OPTION VALUE=""" & oRecordset.Fields("PositionID").Value & """>" & oRecordset.Fields("JobsInArea").Value & ": " & oRecordset.Fields("PositionShortName").Value & ". " & oRecordset.Fields("PositionName").Value & "</OPTION>"
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
							End If
						Response.Write "</SELECT>"
						Response.Write "<SELECT NAME=""JobsForPosition"" ID=""JobsForPositionCmb"" SIZE=""10"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""SelectSameItems(this, document.AreaFrm.PositionID)"" STYLE=""width: 0px"">"
							If lErrorNumber = 0 Then
								oRecordset.MoveFirst
								Do While Not oRecordset.EOF
									Response.Write "<OPTION VALUE=""" & oRecordset.Fields("JobsInArea").Value & """>" & oRecordset.Fields("JobsInArea").Value & "</OPTION>"
									oRecordset.MoveNext
									If Err.number <> 0 Then Exit Do
								Loop
							End If
						Response.Write "</SELECT>"
					Response.Write "</TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE><BR />"

			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?AreaID=" & aAreaComponent(N_ID_AREA) & "'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveAreaWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayAreaPositionsForm = lErrorNumber
	Err.Clear
End Function

Function DisplayAreaAsHiddenFields(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an area using
'		  hidden form fields
'Inputs:  oRequest, oADODBConnection, aAreaComponent
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAreaAsHiddenFields"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaID"" ID=""AreaIDHdn"" VALUE=""" & aAreaComponent(N_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ParentID"" ID=""ParentIDHdn"" VALUE=""" & aAreaComponent(N_PARENT_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaCode"" ID=""AreaCodeHdn"" VALUE=""" & aAreaComponent(S_CODE_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaShortName"" ID=""AreaShortNameHdn"" VALUE=""" & aAreaComponent(S_SHORT_NAME_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaName"" ID=""AreaNameHdn"" VALUE=""" & aAreaComponent(S_NAME_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaPath"" ID=""AreaPathHdn"" VALUE=""" & aAreaComponent(S_PATH_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""URCTAUX"" ID=""URCTAUXHdn"" VALUE=""" & aAreaComponent(S_URCTAUX_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CompanyID"" ID=""CompanyIDHdn"" VALUE=""" & aAreaComponent(N_COMPANY_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaTypeID"" ID=""AreaTypeIDHdn"" VALUE=""" & aAreaComponent(N_TYPE_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConfineTypeID"" ID=""ConfineTypeIDHdn"" VALUE=""" & aAreaComponent(N_CONFINE_TYPE_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaLevelTypeID"" ID=""AreaLevelTypeIDHdn"" VALUE=""" & aAreaComponent(N_LEVEL_TYPE_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CenterTypeID"" ID=""CenterTypeIDHdn"" VALUE=""" & aAreaComponent(N_CENTER_TYPE_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CenterSubtypeID"" ID=""CenterSubtypeIDHdn"" VALUE=""" & aAreaComponent(N_CENTER_SUBTYPE_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AttentionLevelID"" ID=""AttentionLevelIDHdn"" VALUE=""" & aAreaComponent(N_ATTENTION_LEVEL_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaAddress"" ID=""AreaAddressHdn"" VALUE=""" & aAreaComponent(S_ADDRESS_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaCity"" ID=""AreaCityHdn"" VALUE=""" & aAreaComponent(S_CITY_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaZip"" ID=""AreaZipHdn"" VALUE=""" & aAreaComponent(S_ZIP_CODE_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ZoneID"" ID=""ZoneIDHdn"" VALUE=""" & aAreaComponent(N_ZONE_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EconomicZoneID"" ID=""EconomicZoneIDHdn"" VALUE=""" & aAreaComponent(N_ECONOMIC_ZONE_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaymentCenterID"" ID=""PaymentCenterIDHdn"" VALUE=""" & aAreaComponent(N_PAYMENT_CENTER_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GeneratingAreaID"" ID=""GeneratingAreaIDHdn"" VALUE=""" & aAreaComponent(N_GENERATING_AREA_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BranchID"" ID=""BranchIDHdn"" VALUE=""" & aAreaComponent(N_BRANCH_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SubBranchID"" ID=""SubBranchIDHdn"" VALUE=""" & aAreaComponent(N_SUBBRANCH_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CashierOfficeID"" ID=""CashierOfficeIDHdn"" VALUE=""" & aAreaComponent(N_CASHIER_OFFICE_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate"" ID=""StartDateHdn"" VALUE=""" & aAreaComponent(N_START_DATE_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EndDate"" ID=""EndDateHdn"" VALUE=""" & aAreaComponent(N_END_DATE_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FinishDate"" ID=""FinishDateHdn"" VALUE=""" & aAreaComponent(N_FINISH_DATE_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobsInArea"" ID=""JobsInAreaHdn"" VALUE=""" & aAreaComponent(N_JOBS_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TotalJobs"" ID=""TotalJobsHdn"" VALUE=""" & aAreaComponent(N_TOTAL_JOBS_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StatusID"" ID=""StatusIDHdn"" VALUE=""" & aAreaComponent(N_STATUS_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Active"" ID=""ActiveHdn"" VALUE=""" & aAreaComponent(N_ACTIVE_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PositionIDs"" ID=""PositionIDsHdn"" VALUE=""" & aAreaComponent(S_POSITIONS_ID_AREA) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobsForPositions"" ID=""JobsForPositionsHdn"" VALUE=""" & aAreaComponent(S_JOBS_FOR_POSITIONS_AREA) & """ />"

	DisplayAreaAsHiddenFields = Err.number
	Err.Clear
End Function

Function DisplayAreaPath(oRequest, oADODBConnection, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To display the path of an area
'Inputs:  oRequest, oADODBConnection, aAreaComponent
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAreaPath"
	Dim sFullPath
	Dim sTempPath
	Dim lAreaID
	Dim bFirst
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aAreaComponent(B_COMPONENT_INITIALIZED_AREA)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeAreaComponent(oRequest, aAreaComponent)
	End If

	sFullPath = ""
	bFirst = True
	lAreaID = CLng(aAreaComponent(N_ID_AREA))
	Do While (lAreaID <> -1)
		sErrorDescription = "No se pudo obtener la ruta del área."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID, AreaName, ParentID, AreaPath From Areas Where (AreaID=" & lAreaID & ")", "AreaComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				If bFirst Then
					sFullPath = "<B>" & CStr(oRecordset.Fields("AreaName").Value) & "</B>" & sFullPath
					bFirst = False
				Else
					sTempPath = "<A "
						If InStr(1, "," & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ",", "," & CStr(oRecordset.Fields("AreaID").Value) & ",", vbBinaryCompare) > 0 Then sTempPath = sTempPath & "HREF=""" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "&AreaID=" & lAreaID & """"
					sTempPath = sTempPath & ">" & CStr(oRecordset.Fields("AreaName").Value) & "</A> > "
					sFullPath = sTempPath & sFullPath
				End If
				lAreaID = CLng(oRecordset.Fields("ParentID").Value)
			Else
				lAreaID = -1
			End If
		Else
			lAreaID = -1
		End If
	Loop
	Response.Write sFullPath

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayAreaPath = Err.number
	Err.Clear
End Function

Function DisplayAreasInSmallIcons(oRequest, oADODBConnection, bUseLinks, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the areas from
'		  the database using small icons
'Inputs:  oRequest, oADODBConnection, bUseLinks, aAreaComponent
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAreasInSmallIcons"
	Dim sNames
	Dim sColor
	Dim sBoldBegin
	Dim sBoldEnd
	Dim oRecordset
	Dim lErrorNumber

	lErrorNumber = GetAreas(oRequest, oADODBConnection, aAreaComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		Response.Write "<TABLE WIDTH=""200"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			Response.Write "<TR><TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """ COLSPAN=""7""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>"
			Response.Write "<TR BGCOLOR=""#" & S_SELECTED_BGCOLOR_MENU & """>"
				Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""20"" /></TD>"
				Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""3"" HEIGHT=""1"" /></TD>"
				Response.Write "<TD WIDTH=""12"" VALIGN=""MIDDLE"">"
					Response.Write "<A HREF=""" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "&AreaID=" & aAreaComponent(N_PARENT_ID_AREA) & """><IMG SRC=""Images/IcnArrowLevelUpBlack.gif"" WIDTH=""9"" HEIGHT=""9"" ALT=""Regresar al área anterior"" BORDER=""0"" /></A>"
				Response.Write "</TD>"
				Response.Write "<TD WIDTH=""100%"" VALIGN=""MIDDLE""><FONT FACE=""Verdana"" SIZE=""1""><B>CENTROS DE TRABAJO</B></FONT></TD>"
				Response.Write "<TD ALIGN=""RIGHT"" VALIGN=""TOP""><FONT FACE=""Verdana"" SIZE=""1"">&nbsp;</FONT></TD>"
				If bUseLinks Then Response.Write "<TD ALIGN=""RIGHT"" VALIGN=""TOP""><FONT FACE=""Verdana"" SIZE=""1"">&nbsp;</FONT></TD>"
				Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR><TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """ COLSPAN=""7""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>"
			Response.Write "<TR>"
				Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
				Response.Write "<TD COLSPAN=""4""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""5"" /></TD>"
				If bUseLinks Then Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""5"" /></TD>"
				Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
			Response.Write "</TR>"
			If Not oRecordset.EOF Then
				Do While Not oRecordset.EOF
					sColor = ""
					If CInt(oRecordset.Fields("Active").Value) = 0 Then
						sColor = " COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """"
					End If
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("AreaID").Value), oRequest("AreaID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					Response.Write "<TR>"
						Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
						Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""3"" HEIGHT=""1"" /></TD>"
						sNames = Replace(CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value) & ". " & CStr(oRecordset.Fields("AreaName").Value)), "&#34;", """")
						Response.Write "<TD WIDTH=""12"" VALIGN=""TOP""><IMG SRC=""Images/IcnArea.gif"" WIDTH=""12"" HEIGHT=""12"" ALT=""" & sNames & """ BORDER=""0"" />&nbsp;</TD>"
						If Len(sNames) > 42 Then sNames = Left(sNames, 30) & "..."
						Response.Write "<TD WIDTH=""100%"" VALIGN=""TOP""><A HREF=""" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & """><FONT FACE=""Verdana"" SIZE=""1""" & sColor & ">" & sBoldBegin & sNames & sBoldEnd & "</FONT></A></TD>"
						Response.Write "<TD ALIGN=""RIGHT"" VALIGN=""TOP""><FONT FACE=""Verdana"" SIZE=""1""" & sColor & ">" & FormatNumber(CLng(oRecordset.Fields("TotalJobs").Value), 0, True, False, True) & "</FONT></A></TD>"
						If bUseLinks Then
							Response.Write "<TD ALIGN=""CENTER"" VALIGN=""TOP"">"
								If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
									Response.Write "<A HREF=""" & GetASPFileName("") & "?Action=Areas&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & "&Tab=1&Change=1"">"
										Response.Write "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
									Response.Write "</A>&nbsp;"
								End If

								If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
									Response.Write "<A HREF=""" & GetASPFileName("") & "?Action=Areas&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & "&Tab=1&Delete=1"">"
										Response.Write "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									Response.Write "</A>&nbsp;"
								End If

								If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
									If CInt(oRecordset.Fields("Active").Value) = 0 Then
										Response.Write "<A HREF=""" & GetASPFileName("") & "?Action=Areas&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & "&SetActive=1""><IMG SRC=""Images/BtnActive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Activar"" BORDER=""0"" /></A>"
									Else
										Response.Write "<A HREF=""" & GetASPFileName("") & "?Action=Areas&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & "&SetActive=0""><IMG SRC=""Images/BtnDeactive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Desactivar"" BORDER=""0"" /></A>"
									End If
								End If
							Response.Write "</TD>"
						End If
						Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
						Response.Write "<TD COLSPAN=""4""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""5"" /></TD>"
						If bUseLinks Then Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""5"" /></TD>"
						Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
					Response.Write "</TR>"
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Else
				Response.Write "<TR>"
					Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
					Response.Write "<TD COLSPAN=""4""><FONT FACE=""Verdana"" SIZE=""1"">No existen subáreas registradas en esta área.</FONT></TD>"
					If bUseLinks Then Response.Write "<TD><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""5"" /></TD>"
					Response.Write "<TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD>"
				Response.Write "</TR>"
			End If
			Response.Write "<TR><TD BGCOLOR=""#" & S_BGCOLOR_FOR_GUI & """ COLSPAN=""7""><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""1"" /></TD></TR>"
		Response.Write "</TABLE>" & vbNewLine
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayAreasInSmallIcons = lErrorNumber
	Err.Clear
End Function

Function DisplayAreasTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, bForExport, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the areas from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aAreaComponent
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAreasTable"
	Dim iRecordCounter
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
	Dim sTarget
	Dim sFilter
	Dim iStartPage
	Dim lErrorNumber

	sTarget = ""
	sFilter = ""
	iStartPage = 1
	If Len(oRequest("StartPage").Item) > 0 Then iStartPage = CInt(oRequest("StartPage").Item)
	If aAreaComponent(B_SEND_TO_IFRAME_AREA) Then sTarget = " TARGET=""FormsIFrame"""
	If Len(oRequest("ApplyFilter").Item) > 0 Then sFilter = "&ApplyFilter=1&ZoneID=" & CStr(oRequest("ZoneID").Item) & "&AreaCode=" & CStr(oRequest("AreaCode").Item) & "&AreaShortName=" & CStr(oRequest("AreaShortName").Item) & "&CenterTypeID=" & CStr(oRequest("CenterTypeID").Item)
	lErrorNumber = GetAreas(oRequest, oADODBConnection, aAreaComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetch(oRequest, iStartPage, ROWS_CATALOG, oRecordset)
			Response.Write "<TABLE BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks And (Not bForExport) Then
					If aAreaComponent(N_LEVEL_AREA) > 0 Then
						asColumnsTitles = Split("&nbsp;,Acciones,Clave,Clave 10 pos,Denominación Centro de trabajo,Entidad,Municipio,Población,Zona económica,Empresa,Tipo de área,Ámbito del área,UR-CT-AUX,Tipo de centro de trabajo,Subtipo de centro de trabajo,Nivel de atención,Dirección,Ciudad,C.P.,Clave,Denominación Centro de pago,Pagaduría SIPE,Fecha de inicio,Fecha término,Fecha inhabilitado,Activo", ",", -1, vbBinaryCompare)
						asCellWidths = Split("20,80,100,200,100,100,200,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
						asCellAlignments = Split(",,,,,,,,,,,,,,,,,,,,,,CENTER,CENTER", ",", -1, vbBinaryCompare)
					Else
						asColumnsTitles = Split("&nbsp;,Acciones,Área,Entidad", ",", -1, vbBinaryCompare)
						asCellWidths = Split("20,80,650,100", ",", -1, vbBinaryCompare)
						asCellAlignments = Split(",CENTER,,", ",", -1, vbBinaryCompare)
					End If
				Else
					If aAreaComponent(N_LEVEL_AREA) > 0 Then
						asColumnsTitles = Split("&nbsp;,Clave,Clave 10 pos,Denominación Centro de trabajo,Entidad,Municipio,Población,Zona económica,Empresa,Tipo de área,Ámbito del área,UR-CT-AUX,Tipo de centro de trabajo,Subtipo de centro de trabajo,Nivel de atención,Dirección,Ciudad,C.P.,Clave,Denominación Centro de pago,Pagaduría SIPE,Fecha de inicio,Fecha término,Fecha inhabilitado,Activo", ",", -1, vbBinaryCompare)
						asCellWidths = Split("20,100,200,100,200,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
						asCellAlignments = Split(",,,,,,,,,,,,,,,,,,,,,,CENTER", ",", -1, vbBinaryCompare)
					Else
						'asColumnsTitles = Split("&nbsp;,Clave,Nombre", ",", -1, vbBinaryCompare)
						asColumnsTitles = Split("&nbsp;,Área,Entidad", ",", -1, vbBinaryCompare)
						asCellWidths = Split("20,430,100", ",", -1, vbBinaryCompare)
						asCellAlignments = Split(",,,", ",", -1, vbBinaryCompare)
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

				iRecordCounter = 0
				Do While Not oRecordset.EOF
					sFontBegin = ""
					sFontEnd = ""
					If CInt(oRecordset.Fields("Active").Value) = 0 Then
						sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
						sFontEnd = "</FONT>"
					End If
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("AreaID").Value), oRequest("AreaID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					If bUseLinks Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "<NOBR>&nbsp;"
						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Areas&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & "&ParentID=" & oRequest("ParentID").Item & "&Change=1&StartPage=" & oRequest("StartPage").Item & "&PaymentCenters=" & oRequest("PaymentCenters").Item & sFilter & """" & sTarget & ">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If

						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Areas&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & "&ParentID=" & oRequest("ParentID").Item & "&Delete=1&StartPage=" & oRequest("StartPage").Item & "&PaymentCenters=" & oRequest("PaymentCenters").Item & sFilter & """" & sTarget & ">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If

						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							If CInt(oRecordset.Fields("Active").Value) = 0 Then
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Areas&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & "&SetActive=1&StartPage=" & oRequest("StartPage").Item & "&PaymentCenters=" & oRequest("PaymentCenters").Item & sFilter & """><IMG SRC=""Images/BtnActive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Activar"" BORDER=""0"" /></A>"
							Else
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Areas&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & "&SetActive=0&StartPage=" & oRequest("StartPage").Item & "&PaymentCenters=" & oRequest("PaymentCenters").Item & sFilter & """><IMG SRC=""Images/BtnDeactive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Desactivar"" BORDER=""0"" /></A>"
							End If
						End If
						sRowContents = sRowContents & "&nbsp;</NOBR>"

					End If
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""AreaID"" ID=""AreaIDRd"" VALUE=""" & CStr(oRecordset.Fields("AreaID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""AreaID"" ID=""AreaIDChk"" VALUE=""" & CStr(oRecordset.Fields("AreaID").Value) & """ />"
						Case Else
							sRowContents = sRowContents & "&nbsp;"
					End Select
					'If aAreaComponent(N_LEVEL_AREA) > 0 Then
					'	sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("GeneratingAreaShortName").Value) & " " & CStr(oRecordset.Fields("GeneratingAreaName").Value)) & sBoldEnd & sFontEnd
					'End If

					If aAreaComponent(N_LEVEL_AREA) > 0 Then
                        'If CInt(oRecordset.Fields("OnlyPaymentCenter").Value) = 0 Then
                        If True Then
						    sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
							    If bForExport Then sRowContents = sRowContents & "=T("""
							    sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
							    If bForExport Then sRowContents = sRowContents & """)"
						    sRowContents = sRowContents & sBoldEnd & sFontEnd
						    sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;<A"
							    If (Not bForExport) And ((Len(oRequest("AreaID").Item) = 0) And (Len(oRequest("ParentID").Item) = 0)) And (Len(oRequest("PaymentCenters").Item) = 0) Then sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "&ParentID=" & CStr(oRecordset.Fields("AreaID").Value) & "&PaymentCenters=1" & "&ReadOnly=" & oRequest("ReadOnly").Item & """"
						    sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value)) & sBoldEnd & sFontEnd & "</A>"
                        Else
                            sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("---") & sBoldEnd & sFontEnd
                            sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("---") & sBoldEnd & sFontEnd
					    End If
                    End If
                    'If CInt(oRecordset.Fields("OnlyPaymentCenter").Value) = 0 Then
                    If True Then
					    sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
						    If (Not bForExport) And (Len(oRequest("PaymentCenters").Item) = 0) Then sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "&ParentID=" & CStr(oRecordset.Fields("AreaID").Value) & "&PaymentCenters=1" & "&ReadOnly=" & oRequest("ReadOnly").Item & """"
					    sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value)) & sBoldEnd & sFontEnd & "</A>"
                    Else
                        sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML("---") & sBoldEnd & sFontEnd
                    End If
					If aAreaComponent(N_LEVEL_AREA) > 0 Then
					    sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
						    'sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("ZoneCode").Value) & " " & CStr(oRecordset.Fields("ZoneName").Value))
						    sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName").Value))
					    sRowContents = sRowContents & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("Municipio").Value))
						sRowContents = sRowContents & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("Poblacion").Value))
						sRowContents = sRowContents & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EconomicZoneCode").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("CompanyShortName").Value) & " " & CStr(oRecordset.Fields("CompanyName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AreaTypeShortName").Value) & " " & CStr(oRecordset.Fields("AreaTypeName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ConfineTypeShortName").Value) & " " & CStr(oRecordset.Fields("ConfineTypeName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("URCTAUX").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("CenterTypeShortName").Value) & " " & CStr(oRecordset.Fields("CenterTypeName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("CenterSubtypeShortName").Value) & " " & CStr(oRecordset.Fields("CenterSubtypeName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AttentionLevelShortName").Value) & " " & CStr(oRecordset.Fields("AttentionLevelName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaAddress").Value))
						sRowContents = sRowContents & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaCity").Value))
						sRowContents = sRowContents & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaZip").Value))
						sRowContents = sRowContents & sBoldEnd & sFontEnd
						'sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value) & " " & CStr(oRecordset.Fields("PaymentCenterName").Value)) & sBoldEnd & sFontEnd

						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value))
						sRowContents = sRowContents & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterName").Value))
						sRowContents = sRowContents & sBoldEnd & sFontEnd

						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("CashierOfficeShortName").Value) & " " & CStr(oRecordset.Fields("CashierOfficeName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
							If (CLng(oRecordset.Fields("StartDate").Value) = 0) Or (CLng(oRecordset.Fields("StartDate").Value) = 30000000) Then
								sRowContents = sRowContents & "<CENTER>---</CENTER>"
							Else
								sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
							End If
						sRowContents = sRowContents & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
							If (CLng(oRecordset.Fields("EndDate").Value) = 0) Or (CLng(oRecordset.Fields("EndDate").Value) = 30000000) Then
								sRowContents = sRowContents & "<CENTER>---</CENTER>"
							Else
								sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
							End If
						sRowContents = sRowContents & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
							If (CLng(oRecordset.Fields("FinishDate").Value) = 0) Or (CLng(oRecordset.Fields("FinishDate").Value) = 30000000) Then
								sRowContents = sRowContents & "<CENTER>---</CENTER>"
							Else
								sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("FinishDate").Value), -1, -1, -1)
							End If
						sRowContents = sRowContents & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayYesNo(CInt(oRecordset.Fields("Active").Value), True) & sBoldEnd & sFontEnd
                    Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("Poblacion").Value))
						sRowContents = sRowContents & sBoldEnd & sFontEnd
					End If

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
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayAreasTable = lErrorNumber
	Err.Clear
End Function

Function DisplayAreasTableFull(oRequest, oADODBConnection, lIDColumn, bUseLinks, bForExport, aAreaComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the areas from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aAreaComponent
'Outputs: aAreaComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAreasTableFull"
	Dim iRecordCounter
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
	Dim sTarget
	Dim sFind
	Dim iStartPage
	Dim lErrorNumber
	Dim aiAreasID
    Dim aiAreasID1
	Dim iIndex

	sTarget = ""
	sFind = ""
	iStartPage = 1
	If Len(oRequest("StartPage").Item) > 0 Then iStartPage = CInt(oRequest("StartPage").Item)
	If aAreaComponent(B_SEND_TO_IFRAME_AREA) Then sTarget = " TARGET=""FormsIFrame"""
	If Len(oRequest("AreaFind").Item) > 0 Then sFind = "&AreaFind=1"


	aAreaComponent(S_QUERY_CONDITION_AREA) = " And (Areas.ParentID=-1)"
	lErrorNumber = GetAreas(oRequest, oADODBConnection, aAreaComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				aiAreasID = aiAreasID & LIST_SEPARATOR & CDbl(oRecordset.Fields("AreaID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("AreaName").Value) & SECOND_LIST_SEPARATOR
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			aiAreasID = Split(aiAreasID, SECOND_LIST_SEPARATOR)
			'For iIndex = 1 To UBound(aiAreasID)
			'	aiAreasID(iIndex) = CDbl(aiAreasID(iIndex))
			'Next
		End If
	End If

	If Not bForExport Then Call DisplayIncrementalFetch(oRequest, iStartPage, ROWS_CATALOG, oRecordset)
	Response.Write "<TABLE BORDER="""
		If bForExport Then
			Response.Write "1"
		Else
			Response.Write "0"
		End If
	Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
		If bUseLinks And (Not bForExport) Then
			If Len(oRequest("PaymentCenters").Item) = 0 Then
				asColumnsTitles = Split("Acciones,Área generadora,Municipio,Población,Entidad,Clave 10 pos.,Código,Nombre,UR-CT-AUX,Empresa,Tipo de área,Ámbito del área,Tipo de centro de trabajo,Subtipo de centro de trabajo,Nivel de atención,Dirección,Ciudad,C.P.,Zona económica,Centro de pago,Pagaduría SIPE,Fecha de inicio,Fecha término,Fecha inhabilitado", ",", -1, vbBinaryCompare)
				asCellWidths = Split("20,80,100,100,200,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				asCellAlignments = Split(",,,,,,,,,,,,,,,,,,,,,,CENTER,CENTER", ",", -1, vbBinaryCompare)
			Else
				asColumnsTitles = Split("Acciones,Área generadora,Municipio,Población,Entidad,Clave 10 pos.,Nombre,UR-CT-AUX,Empresa,Tipo de área,Ámbito del área,Tipo de centro de trabajo,Subtipo de centro de trabajo,Nivel de atención,Dirección,Ciudad,C.P.,Zona económica,Centro de pago,Pagaduría SIPE,Fecha de inicio,Fecha término,Fecha inhabilitado,Activo", ",", -1, vbBinaryCompare)
				asCellWidths = Split("20,80,100,200,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				asCellAlignments = Split(",,,,,,,,,,,,,,,,,,,,,CENTER,CENTER", ",", -1, vbBinaryCompare)
			End If
		Else
			If Len(oRequest("PaymentCenters").Item) = 0 Then
				asColumnsTitles = Split("Área generadora,Municipio,Población,Entidad,Clave 10 pos.,Código,Nombre,UR-CT-AUX,Empresa,Tipo de área,Ámbito del área,Tipo de centro de trabajo,Subtipo de centro de trabajo,Nivel de atención,Dirección,Ciudad,C.P.,Zona económica,Centro de pago,Pagaduría SIPE,Fecha de inicio,Fecha término,Fecha inhabilitado,Activo", ",", -1, vbBinaryCompare)
				asCellWidths = Split("20,100,100,200,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				asCellAlignments = Split(",,,,,,,,,,,,,,,,,,,,,,CENTER", ",", -1, vbBinaryCompare)
			Else
				asColumnsTitles = Split("Área generadora,Municipio,Población,Entidad,Clave 10 pos.,Nombre,UR-CT-AUX,Empresa,Tipo de área,Ámbito del área,Tipo de centro de trabajo,Subtipo de centro de trabajo,Nivel de atención,Dirección,Ciudad,C.P.,Zona económica,Centro de pago,Pagaduría SIPE,Fecha de inicio,Fecha término,Fecha inhabilitado,Activo", ",", -1, vbBinaryCompare)
				asCellWidths = Split("20,100,200,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				asCellAlignments = Split(",,,,,,,,,,,,,,,,,,,,,CENTER", ",", -1, vbBinaryCompare)
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

	For iIndex = 1 To UBound(aiAreasID)
        aiAreasID1 = Split(aiAreasID(iIndex), LIST_SEPARATOR)
		aAreaComponent(S_QUERY_CONDITION_AREA) = " And (Areas.ParentID=" & aiAreasID1(1) & ")"
		lErrorNumber = GetAreas(oRequest, oADODBConnection, aAreaComponent, oRecordset, sErrorDescription)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				iRecordCounter = 0
				Do While Not oRecordset.EOF
					sFontBegin = ""
					sFontEnd = ""
					If CInt(oRecordset.Fields("Active").Value) = 0 Then
						sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
						sFontEnd = "</FONT>"
					End If
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("AreaID").Value), oRequest("AreaID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					If bUseLinks Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "<NOBR>&nbsp;"
						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Areas&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & "&ParentID=" & oRequest("ParentID").Item & "&Change=1&StartPage=" & oRequest("StartPage").Item & "&PaymentCenters=" & oRequest("PaymentCenters").Item & sFind & """" & sTarget & ">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If

						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Areas&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & "&ParentID=" & oRequest("ParentID").Item & "&Delete=1&StartPage=" & oRequest("StartPage").Item & "&PaymentCenters=" & oRequest("PaymentCenters").Item & sFind & """" & sTarget & ">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If

						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							If CInt(oRecordset.Fields("Active").Value) = 0 Then
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Areas&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & "&SetActive=1&StartPage=" & oRequest("StartPage").Item & "&PaymentCenters=" & oRequest("PaymentCenters").Item & sFind & """><IMG SRC=""Images/BtnActive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Activar"" BORDER=""0"" /></A>"
							Else
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Areas&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & "&SetActive=0&StartPage=" & oRequest("StartPage").Item & "&PaymentCenters=" & oRequest("PaymentCenters").Item & sFind & """><IMG SRC=""Images/BtnDeactive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Desactivar"" BORDER=""0"" /></A>"
							End If
						End If
						sRowContents = sRowContents & "&nbsp;</NOBR>"

					End If
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""AreaID"" ID=""AreaIDRd"" VALUE=""" & CStr(oRecordset.Fields("AreaID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""AreaID"" ID=""AreaIDChk"" VALUE=""" & CStr(oRecordset.Fields("AreaID").Value) & """ />"
						Case Else
							sRowContents = sRowContents & "&nbsp;"
					End Select
					sRowContents = sRowContents & sFontBegin & sBoldBegin
						sRowContents = sRowContents & CleanStringForHTML(CStr(aiAreasID1(2)))
					sRowContents = sRowContents & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("Municipio").Value))
					sRowContents = sRowContents & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("Poblacion").Value))
					sRowContents = sRowContents & sBoldEnd & sFontEnd
                    sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName").Value)) & sBoldEnd & sFontEnd
					If Len(oRequest("PaymentCenters").Item) = 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;<A"
							If (Not bForExport) And ((Len(oRequest("AreaID").Item) = 0) And (Len(oRequest("ParentID").Item) = 0)) And (Len(oRequest("PaymentCenters").Item) = 0) Then sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "&ParentID=" & CStr(oRecordset.Fields("AreaID").Value) & "&PaymentCenters=1" & "&ReadOnly=" & oRequest("ReadOnly").Item & """"
						sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value)) & sBoldEnd & sFontEnd & "</A>"
					End If
					'If (Len(oRequest("PaymentCenters").Item) > 0) Or (Len(oRequest("AreaID").Item) > 0) Or (Len(oRequest("ParentID").Item) > 0) Then
					'If (Len(oRequest("PaymentCenters").Item) > 0) Then
					'If aAreaComponent(N_LEVEL_AREA) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
							If bForExport Then sRowContents = sRowContents & "=T("""
							sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaCode").Value))
							If bForExport Then sRowContents = sRowContents & """)"
						sRowContents = sRowContents & sBoldEnd & sFontEnd
					'End If
					sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
						'If (Not bForExport) And ((Len(oRequest("AreaID").Item) = 0) And (Len(oRequest("ParentID").Item) = 0)) And (Len(oRequest("PaymentCenters").Item) = 0) Then sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & "&ReadOnly=" & oRequest("ReadOnly").Item & """"
						If (Not bForExport) And (Len(oRequest("PaymentCenters").Item) = 0) Then sRowContents = sRowContents & " HREF=""" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "&ParentID=" & CStr(oRecordset.Fields("AreaID").Value) & "&PaymentCenters=1" & "&ReadOnly=" & oRequest("ReadOnly").Item & """"
					sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value)) & sBoldEnd & sFontEnd & "</A>"
					'If (Len(oRequest("PaymentCenters").Item) > 0) Or (Len(oRequest("AreaID").Item) > 0) Or (Len(oRequest("ParentID").Item) > 0) Then
					'If (Len(oRequest("PaymentCenters").Item) > 0) Then
					If True Then'Len(oRequest("PaymentCenters").Item) = 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("URCTAUX").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("CompanyShortName").Value) & " " & CStr(oRecordset.Fields("CompanyName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AreaTypeShortName").Value) & " " & CStr(oRecordset.Fields("AreaTypeName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ConfineTypeShortName").Value) & " " & CStr(oRecordset.Fields("ConfineTypeName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("CenterTypeShortName").Value) & " " & CStr(oRecordset.Fields("CenterTypeName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("CenterSubtypeShortName").Value) & " " & CStr(oRecordset.Fields("CenterSubtypeName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AttentionLevelShortName").Value) & " " & CStr(oRecordset.Fields("AttentionLevelName").Value)) & sBoldEnd & sFontEnd
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaAddress").Value))
					sRowContents = sRowContents & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaCity").Value))
					sRowContents = sRowContents & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("AreaZip").Value))
					sRowContents = sRowContents & sBoldEnd & sFontEnd
					If True Then'Len(oRequest("PaymentCenters").Item) = 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EconomicZoneCode").Value) & " " & CStr(oRecordset.Fields("EconomicZoneName").Value)) & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterShortName").Value) & " " & CStr(oRecordset.Fields("PaymentCenterName").Value)) & sBoldEnd & sFontEnd							
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("CashierOfficeShortName").Value) & " " & CStr(oRecordset.Fields("CashierOfficeName").Value)) & sBoldEnd & sFontEnd
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
						If (CLng(oRecordset.Fields("StartDate").Value) = 0) Or (CLng(oRecordset.Fields("StartDate").Value) = 30000000) Then
							sRowContents = sRowContents & "<CENTER>---</CENTER>"
						Else
							sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1)
						End If
					sRowContents = sRowContents & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
						If (CLng(oRecordset.Fields("EndDate").Value) = 0) Or (CLng(oRecordset.Fields("EndDate").Value) = 30000000) Then
							sRowContents = sRowContents & "<CENTER>---</CENTER>"
						Else
							sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
						End If
					sRowContents = sRowContents & sBoldEnd & sFontEnd
					If True Then'Len(oRequest("PaymentCenters").Item) = 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
							If (CLng(oRecordset.Fields("FinishDate").Value) = 0) Or (CLng(oRecordset.Fields("FinishDate").Value) = 30000000) Then
								sRowContents = sRowContents & "<CENTER>---</CENTER>"
							Else
								sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("FinishDate").Value), -1, -1, -1)
							End If
						sRowContents = sRowContents & sBoldEnd & sFontEnd
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayYesNo(CInt(oRecordset.Fields("Active").Value), True) & sBoldEnd & sFontEnd

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
			End If
		End If
	Next

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayAreasTableFull = lErrorNumber
	Err.Clear
End Function
%>