<!-- #include file="ConceptAddComponent.asp" -->
<!-- #include file="ConceptDisplayTablesComponent.asp" -->
<%
Const N_ID_CONCEPT = 0
Const N_START_DATE_CONCEPT = 1
Const N_END_DATE_CONCEPT = 2
Const S_SHORT_NAME_CONCEPT = 3
Const S_NAME_CONCEPT = 4
Const N_BUDGET_ID_CONCEPT = 5
Const N_PAYROLL_TYPE_ID_CONCEPT = 6
Const N_PERIOD_ID_CONCEPT = 7
Const N_IS_DEDUCTION_CONCEPT = 8
Const N_FOR_ALIMONY_CONCEPT = 9
Const N_ON_LEAVE_CONCEPT = 10
Const N_FOR_TAX_CONCEPT = 11
Const N_ORDER_IN_LIST_CONCEPT = 12
Const D_TAX_AMOUNT_CONCEPT = 13
Const N_TAX_CURRENCY_ID_CONCEPT = 14
Const N_TAX_QTTY_ID_CONCEPT = 15
Const D_TAX_MIN_CONCEPT = 16
Const N_TAX_MIN_QTTY_ID_CONCEPT = 17
Const D_TAX_MAX_CONCEPT = 18
Const N_TAX_MAX_QTTY_ID_CONCEPT = 19
Const D_EXEMPT_AMOUNT_CONCEPT = 20
Const N_EXEMPT_CURRENCY_ID_CONCEPT = 21
Const N_EXEMPT_QTTY_ID_CONCEPT = 22
Const D_EXEMPT_MIN_CONCEPT = 23
Const N_EXEMPT_MIN_QTTY_ID_CONCEPT = 24
Const D_EXEMPT_MAX_CONCEPT = 25
Const N_EXEMPT_MAX_QTTY_ID_CONCEPT = 26

Const N_RECORD_ID_CONCEPT = 27
Const N_COMPANY_ID_CONCEPT = 28
Const N_EMPLOYEE_TYPE_ID_CONCEPT = 29
Const N_POSITION_TYPE_ID_CONCEPT = 30
Const N_EMPLOYEE_STATUS_ID_CONCEPT = 31
Const N_JOB_STATUS_ID_CONCEPT = 32
Const N_CLASSIFICATION_ID_CONCEPT = 33
Const N_GROUP_GRADE_LEVEL_ID_CONCEPT = 34
Const N_INTEGRATION_ID_CONCEPT = 35
Const N_JOURNEY_ID_CONCEPT = 36
Const D_WORKING_HOURS_CONCEPT = 37
Const N_ADDITIONAL_SHIFT_CONCEPT = 38
Const N_LEVEL_ID_CONCEPT = 39
Const N_ECONOMIC_ZONE_ID_CONCEPT = 40
Const N_SERVICE_ID_CONCEPT = 41
Const N_ANTIQUITY_ID_CONCEPT = 42
Const N_ANTIQUITY2_ID_CONCEPT = 43
Const N_ANTIQUITY3_ID_CONCEPT = 44
Const N_ANTIQUITY4_ID_CONCEPT = 45
Const N_FOR_RISK_CONCEPT = 46
Const N_GENDER_ID_CONCEPT = 47
Const N_HAS_CHILDREN_CONCEPT = 48
Const N_SCHOOLARSHIP_ID_CONCEPT = 49
Const N_HAS_SYNDICATE_CONCEPT = 50
Const N_START_DATE_FOR_VALUE_CONCEPT = 51
Const N_END_DATE_FOR_VALUE_CONCEPT = 52
Const N_START_DATE_FOR_REGISTRATION_CONCEPT = 53
Const N_END_DATE_FOR_REGISTRATION_CONCEPT = 54
Const N_END_DATE_FOR_AUTHORIZATION_CONCEPT = 55
Const D_CONCEPT_AMOUNT_CONCEPT = 56
Const N_CURRENCY_ID_CONCEPT = 57
Const N_CONCEPT_QTTY_ID_CONCEPT = 58
Const N_CONCEPT_TYPE_ID_CONCEPT = 59
Const S_APPLIES_ID_CONCEPT = 60
Const D_CONCEPT_MIN_CONCEPT = 61
Const N_CONCEPT_MIN_QTTY_ID_CONCEPT = 62
Const D_CONCEPT_MAX_CONCEPT = 63
Const N_CONCEPT_MAX_QTTY_ID_CONCEPT = 64
Const N_POSITION_ID_CONCEPT = 65
Const S_POSITION_SHORT_NAME_CONCEPT = 66
Const N_START_USER_ID_CONCEPT = 67
Const N_END_USER_ID_CONCEPT = 68
Const N_STATUS_ID_CONCEPT = 69
Const N_CENTER_TYPE_ID = 70
Const N_IS_ACTIVE1 = 71
Const N_IS_ACTIVE2 = 72
Const N_IS_ACTIVE3 = 73
Const N_IS_ACTIVE4 = 74
Const N_IS_CREDIT = 75
Const N_IS_OTHER = 76
Const N_ACTIVE = 77
Const S_QUERY_CONDITION_CONCEPT = 78
Const B_CHECK_FOR_DUPLICATED_CONCEPT = 79
Const B_IS_DUPLICATED_CONCEPT = 80
Const B_COMPONENT_INITIALIZED_CONCEPT = 81

Const N_CONCEPT_COMPONENT_SIZE = 81

Dim aConceptComponent()
Redim aConceptComponent(N_CONCEPT_COMPONENT_SIZE)

Function InitializeConceptComponent(oRequest, aConceptComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Concept
'         Component using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aConceptComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeConceptComponent"
	Redim Preserve aConceptComponent(N_CONCEPT_COMPONENT_SIZE)
	Dim oItem

	If IsEmpty(aConceptComponent(N_ID_CONCEPT)) Then
		If Len(oRequest("ConceptID").Item) > 0 Then
			aConceptComponent(N_ID_CONCEPT) = CLng(oRequest("ConceptID").Item)
		Else
			aConceptComponent(N_ID_CONCEPT) = -1
		End If
	End If

	'If IsEmpty(aConceptComponent(N_START_DATE_CONCEPT)) Then
		If Len(oRequest("StartYear").Item) > 0 Then
			aConceptComponent(N_START_DATE_CONCEPT) = CLng(oRequest("StartYear").Item & Right(("0" & oRequest("StartMonth").Item), Len("00")) & Right(("0" & oRequest("StartDay").Item), Len("00")))
		ElseIf Len(oRequest("StartDateYear").Item) > 0 Then
			aConceptComponent(N_START_DATE_CONCEPT) = CLng(oRequest("StartDateYear").Item & Right(("0" & oRequest("StartDateMonth").Item), Len("00")) & Right(("0" & oRequest("StartDateDay").Item), Len("00")))
		ElseIf Len(oRequest("StartDate").Item) > 0 Then
			aConceptComponent(N_START_DATE_CONCEPT) = CLng(oRequest("StartDate").Item)
		Else
			aConceptComponent(N_START_DATE_CONCEPT) = Left(GetSerialNumberForDate(""), Len("00000000"))
		End If
	'End If

	'If IsEmpty(aConceptComponent(N_END_DATE_CONCEPT)) Then
		If Len(oRequest("EndYear").Item) > 0 Then
			aConceptComponent(N_END_DATE_CONCEPT) = CLng(oRequest("EndYear").Item & Right(("0" & oRequest("EndMonth").Item), Len("00")) & Right(("0" & oRequest("EndDay").Item), Len("00")))
		ElseIf Len(oRequest("EndDateYear").Item) > 0 Then
			aConceptComponent(N_END_DATE_CONCEPT) = CLng(oRequest("EndDateYear").Item & Right(("0" & oRequest("EndDateMonth").Item), Len("00")) & Right(("0" & oRequest("EndDateDay").Item), Len("00")))
		ElseIf Len(oRequest("EndDate").Item) > 0 Then
			aConceptComponent(N_END_DATE_CONCEPT) = CLng(oRequest("EndDate").Item)
		Else
			aConceptComponent(N_END_DATE_CONCEPT) = 30000000
		End If
	'End If
	If CLng(aConceptComponent(N_END_DATE_CONCEPT)) = 0 Then aConceptComponent(N_END_DATE_CONCEPT) = 30000000

	If IsEmpty(aConceptComponent(S_SHORT_NAME_CONCEPT)) Then
		If Len(oRequest("ConceptShortName").Item) > 0 Then
			aConceptComponent(S_SHORT_NAME_CONCEPT) = oRequest("ConceptShortName").Item
		Else
			aConceptComponent(S_SHORT_NAME_CONCEPT) = ""
		End If
	End If
	aConceptComponent(S_SHORT_NAME_CONCEPT) = Left(aConceptComponent(S_SHORT_NAME_CONCEPT), 5)

	If IsEmpty(aConceptComponent(S_NAME_CONCEPT)) Then
		If Len(oRequest("ConceptName").Item) > 0 Then
			aConceptComponent(S_NAME_CONCEPT) = oRequest("ConceptName").Item
		Else
			aConceptComponent(S_NAME_CONCEPT) = ""
		End If
	End If
	aConceptComponent(S_NAME_CONCEPT) = Left(aConceptComponent(S_NAME_CONCEPT), 255)

	If IsEmpty(aConceptComponent(N_BUDGET_ID_CONCEPT)) Then
		If Len(oRequest("BudgetID").Item) > 0 Then
			aConceptComponent(N_BUDGET_ID_CONCEPT) = CLng(oRequest("BudgetID").Item)
		Else
			aConceptComponent(N_BUDGET_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_PAYROLL_TYPE_ID_CONCEPT)) Then
		If Len(oRequest("PayrollTypeID").Item) > 0 Then
			aConceptComponent(N_PAYROLL_TYPE_ID_CONCEPT) = CLng(oRequest("PayrollTypeID").Item)
		Else
			aConceptComponent(N_PAYROLL_TYPE_ID_CONCEPT) = 1
		End If
	End If

	If IsEmpty(aConceptComponent(N_PERIOD_ID_CONCEPT)) Then
		If Len(oRequest("PeriodID").Item) > 0 Then
			aConceptComponent(N_PERIOD_ID_CONCEPT) = CLng(oRequest("PeriodID").Item)
		Else
			aConceptComponent(N_PERIOD_ID_CONCEPT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_IS_DEDUCTION_CONCEPT)) Then
		If Len(oRequest("IsDeduction").Item) > 0 Then
			aConceptComponent(N_IS_DEDUCTION_CONCEPT) = CInt(oRequest("IsDeduction").Item)
		Else
			aConceptComponent(N_IS_DEDUCTION_CONCEPT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_FOR_ALIMONY_CONCEPT)) Then
		If Len(oRequest("ForAlimony").Item) > 0 Then
			aConceptComponent(N_FOR_ALIMONY_CONCEPT) = CInt(oRequest("ForAlimony").Item)
		Else
			aConceptComponent(N_FOR_ALIMONY_CONCEPT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_ON_LEAVE_CONCEPT)) Then
		If Len(oRequest("OnLeave").Item) > 0 Then
			aConceptComponent(N_ON_LEAVE_CONCEPT) = CInt(oRequest("OnLeave").Item)
		Else
			aConceptComponent(N_ON_LEAVE_CONCEPT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_ORDER_IN_LIST_CONCEPT)) Then
		If Len(oRequest("OrderInList").Item) > 0 Then
			aConceptComponent(N_ORDER_IN_LIST_CONCEPT) = CInt(oRequest("OrderInList").Item)
		Else
			aConceptComponent(N_ORDER_IN_LIST_CONCEPT) = 400
		End If
	End If

	If IsEmpty(aConceptComponent(D_TAX_AMOUNT_CONCEPT)) Then
		If Len(oRequest("TaxAmount").Item) > 0 Then
			aConceptComponent(D_TAX_AMOUNT_CONCEPT) = CDbl(oRequest("TaxAmount").Item)
		Else
			aConceptComponent(D_TAX_AMOUNT_CONCEPT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_TAX_CURRENCY_ID_CONCEPT)) Then
		If Len(oRequest("TaxCurrencyID").Item) > 0 Then
			aConceptComponent(N_TAX_CURRENCY_ID_CONCEPT) = CLng(oRequest("TaxCurrencyID").Item)
		Else
			aConceptComponent(N_TAX_CURRENCY_ID_CONCEPT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_TAX_QTTY_ID_CONCEPT)) Then
		If Len(oRequest("TaxQttyID").Item) > 0 Then
			aConceptComponent(N_TAX_QTTY_ID_CONCEPT) = CInt(oRequest("TaxQttyID").Item)
		Else
			aConceptComponent(N_TAX_QTTY_ID_CONCEPT) = 1
		End If
	End If

	If IsEmpty(aConceptComponent(D_TAX_MIN_CONCEPT)) Then
		If Len(oRequest("TaxMin").Item) > 0 Then
			aConceptComponent(D_TAX_MIN_CONCEPT) = CDbl(oRequest("TaxMin").Item)
		Else
			aConceptComponent(D_TAX_MIN_CONCEPT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_TAX_MIN_QTTY_ID_CONCEPT)) Then
		If Len(oRequest("TaxMinQttyID").Item) > 0 Then
			aConceptComponent(N_TAX_MIN_QTTY_ID_CONCEPT) = CInt(oRequest("TaxMinQttyID").Item)
		Else
			aConceptComponent(N_TAX_MIN_QTTY_ID_CONCEPT) = 1
		End If
	End If

	If IsEmpty(aConceptComponent(D_TAX_MAX_CONCEPT)) Then
		If Len(oRequest("TaxMax").Item) > 0 Then
			aConceptComponent(D_TAX_MAX_CONCEPT) = CDbl(oRequest("TaxMax").Item)
		Else
			aConceptComponent(D_TAX_MAX_CONCEPT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_TAX_MAX_QTTY_ID_CONCEPT)) Then
		If Len(oRequest("TaxMaxQttyID").Item) > 0 Then
			aConceptComponent(N_TAX_MAX_QTTY_ID_CONCEPT) = CInt(oRequest("TaxMaxQttyID").Item)
		Else
			aConceptComponent(N_TAX_MAX_QTTY_ID_CONCEPT) = 1
		End If
	End If

	If IsEmpty(aConceptComponent(D_EXEMPT_AMOUNT_CONCEPT)) Then
		If Len(oRequest("ExemptAmount").Item) > 0 Then
			aConceptComponent(D_EXEMPT_AMOUNT_CONCEPT) = CDbl(oRequest("ExemptAmount").Item)
		Else
			aConceptComponent(D_EXEMPT_AMOUNT_CONCEPT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_EXEMPT_CURRENCY_ID_CONCEPT)) Then
		If Len(oRequest("ExemptCurrencyID").Item) > 0 Then
			aConceptComponent(N_EXEMPT_CURRENCY_ID_CONCEPT) = CLng(oRequest("ExemptCurrencyID").Item)
		Else
			aConceptComponent(N_EXEMPT_CURRENCY_ID_CONCEPT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_EXEMPT_QTTY_ID_CONCEPT)) Then
		If Len(oRequest("ExemptQttyID").Item) > 0 Then
			aConceptComponent(N_EXEMPT_QTTY_ID_CONCEPT) = CInt(oRequest("ExemptQttyID").Item)
		Else
			aConceptComponent(N_EXEMPT_QTTY_ID_CONCEPT) = 1
		End If
	End If

	If IsEmpty(aConceptComponent(D_EXEMPT_MIN_CONCEPT)) Then
		If Len(oRequest("ExemptMin").Item) > 0 Then
			aConceptComponent(D_EXEMPT_MIN_CONCEPT) = CDbl(oRequest("ExemptMin").Item)
		Else
			aConceptComponent(D_EXEMPT_MIN_CONCEPT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_EXEMPT_MIN_QTTY_ID_CONCEPT)) Then
		If Len(oRequest("ExemptMinQttyID").Item) > 0 Then
			aConceptComponent(N_EXEMPT_MIN_QTTY_ID_CONCEPT) = CInt(oRequest("ExemptMinQttyID").Item)
		Else
			aConceptComponent(N_EXEMPT_MIN_QTTY_ID_CONCEPT) = 1
		End If
	End If

	If IsEmpty(aConceptComponent(D_EXEMPT_MAX_CONCEPT)) Then
		If Len(oRequest("ExemptMax").Item) > 0 Then
			aConceptComponent(D_EXEMPT_MAX_CONCEPT) = CDbl(oRequest("ExemptMax").Item)
		Else
			aConceptComponent(D_EXEMPT_MAX_CONCEPT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_EXEMPT_MAX_QTTY_ID_CONCEPT)) Then
		If Len(oRequest("ExemptMaxQttyID").Item) > 0 Then
			aConceptComponent(N_EXEMPT_MAX_QTTY_ID_CONCEPT) = CInt(oRequest("ExemptMaxQttyID").Item)
		Else
			aConceptComponent(N_EXEMPT_MAX_QTTY_ID_CONCEPT) = 1
		End If
	End If

	If IsEmpty(aConceptComponent(N_RECORD_ID_CONCEPT)) Then
		If Len(oRequest("RecordID").Item) > 0 Then
			aConceptComponent(N_RECORD_ID_CONCEPT) = CLng(oRequest("RecordID").Item)
		Else
			aConceptComponent(N_RECORD_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_COMPANY_ID_CONCEPT)) Then
		If Len(oRequest("CompanyID").Item) > 0 Then
			aConceptComponent(N_COMPANY_ID_CONCEPT) = CLng(oRequest("CompanyID").Item)
		Else
			aConceptComponent(N_COMPANY_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT)) Then
		If Len(oRequest("EmployeeTypeID").Item) > 0 Then
			aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = CLng(oRequest("EmployeeTypeID").Item)
		ElseIf Len(oRequest("Tab").Item) > 0 Then
			aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = CLng(oRequest("Tab").Item)
		Else
			aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = 1
		End If
	End If

	If IsEmpty(aConceptComponent(N_POSITION_TYPE_ID_CONCEPT)) Then
		If Len(oRequest("PositionTypeID").Item) > 0 Then
			aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) = CLng(oRequest("PositionTypeID").Item)
		Else
			aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT)) Then
		If Len(oRequest("EmployeeStatusID").Item) > 0 Then
			aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) = CLng(oRequest("EmployeeStatusID").Item)
		Else
			aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_JOB_STATUS_ID_CONCEPT)) Then
		If Len(oRequest("JobStatusID").Item) > 0 Then
			aConceptComponent(N_JOB_STATUS_ID_CONCEPT) = CLng(oRequest("JobStatusID").Item)
		Else
			aConceptComponent(N_JOB_STATUS_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_CLASSIFICATION_ID_CONCEPT)) Then
		If Len(oRequest("ClassificationID").Item) > 0 Then
			aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) = CLng(oRequest("ClassificationID").Item)
		Else
			aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT)) Then
		If Len(oRequest("GroupGradeLevelID").Item) > 0 Then
			aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) = CLng(oRequest("GroupGradeLevelID").Item)
		Else
			aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_INTEGRATION_ID_CONCEPT)) Then
		If Len(oRequest("IntegrationID").Item) > 0 Then
			aConceptComponent(N_INTEGRATION_ID_CONCEPT) = CLng(oRequest("IntegrationID").Item)
		Else
			aConceptComponent(N_INTEGRATION_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_JOURNEY_ID_CONCEPT)) Then
		If Len(oRequest("JourneyID").Item) > 0 Then
			aConceptComponent(N_JOURNEY_ID_CONCEPT) = CLng(oRequest("JourneyID").Item)
		Else
			aConceptComponent(N_JOURNEY_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(D_WORKING_HOURS_CONCEPT)) Then
		If Len(oRequest("WorkingHours").Item) > 0 Then
			aConceptComponent(D_WORKING_HOURS_CONCEPT) = CDbl(oRequest("WorkingHours").Item)
		Else
			aConceptComponent(D_WORKING_HOURS_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT)) Then
		If Len(oRequest("AdditionalShift").Item) > 0 Then
			aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) = CInt(oRequest("AdditionalShift").Item)
		Else
			aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_LEVEL_ID_CONCEPT)) Then
		If Len(oRequest("LevelID").Item) > 0 Then
			aConceptComponent(N_LEVEL_ID_CONCEPT) = CLng(oRequest("LevelID").Item)
		Else
			If Len(oRequest("PositionIDTemp").Item) > 0 Then
				Dim asPositionIDTemp
				asPositionIDTemp = Split(oRequest("PositionIDTemp").Item, ",")
				aConceptComponent(N_LEVEL_ID_CONCEPT) = CInt(asPositionIDTemp(1))
			Else
				aConceptComponent(N_LEVEL_ID_CONCEPT) = -1
			End If
		End If
	End If

	If IsEmpty(aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT)) Then
		If Len(oRequest("EconomicZoneID").Item) > 0 Then
			aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) = CInt(oRequest("EconomicZoneID").Item)
		Else
			aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_SERVICE_ID_CONCEPT)) Then
		If Len(oRequest("ServiceID").Item) > 0 Then
			aConceptComponent(N_SERVICE_ID_CONCEPT) = CLng(oRequest("ServiceID").Item)
		Else
			aConceptComponent(N_SERVICE_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_ANTIQUITY_ID_CONCEPT)) Then
		If Len(oRequest("AntiquityID").Item) > 0 Then
			aConceptComponent(N_ANTIQUITY_ID_CONCEPT) = CLng(oRequest("AntiquityID").Item)
		Else
			aConceptComponent(N_ANTIQUITY_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_ANTIQUITY2_ID_CONCEPT)) Then
		If Len(oRequest("Antiquity2ID").Item) > 0 Then
			aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) = CLng(oRequest("Antiquity2ID").Item)
		Else
			aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_ANTIQUITY3_ID_CONCEPT)) Then
		If Len(oRequest("Antiquity3ID").Item) > 0 Then
			aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) = CLng(oRequest("Antiquity3ID").Item)
		Else
			aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_ANTIQUITY4_ID_CONCEPT)) Then
		If Len(oRequest("Antiquity4ID").Item) > 0 Then
			aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) = CLng(oRequest("Antiquity4ID").Item)
		Else
			aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_FOR_RISK_CONCEPT)) Then
		If Len(oRequest("ForRisk").Item) > 0 Then
			aConceptComponent(N_FOR_RISK_CONCEPT) = CInt(oRequest("ForRisk").Item)
		Else
			aConceptComponent(N_FOR_RISK_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_GENDER_ID_CONCEPT)) Then
		If Len(oRequest("GenderID").Item) > 0 Then
			aConceptComponent(N_GENDER_ID_CONCEPT) = CInt(oRequest("GenderID").Item)
		Else
			aConceptComponent(N_GENDER_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_HAS_CHILDREN_CONCEPT)) Then
		If Len(oRequest("HasChildren").Item) > 0 Then
			aConceptComponent(N_HAS_CHILDREN_CONCEPT) = CInt(oRequest("HasChildren").Item)
		Else
			aConceptComponent(N_HAS_CHILDREN_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT)) Then
		If Len(oRequest("SchoolarshipID").Item) > 0 Then
			aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) = CLng(oRequest("SchoolarshipID").Item)
		Else
			aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_HAS_SYNDICATE_CONCEPT)) Then
		If Len(oRequest("HasSyndicate").Item) > 0 Then
			aConceptComponent(N_HAS_SYNDICATE_CONCEPT) = CInt(oRequest("HasSyndicate").Item)
		Else
			aConceptComponent(N_HAS_SYNDICATE_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT)) Then
		If Len(oRequest("StartForValueYear").Item) > 0 Then
			aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) = CLng(oRequest("StartForValueYear").Item & Right(("0" & oRequest("StartForValueMonth").Item), Len("00")) & Right(("0" & oRequest("StartForValueDay").Item), Len("00")))
		ElseIf Len(oRequest("StartForValueDate").Item) > 0 Then
			aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) = CLng(oRequest("StartForValueDate").Item)
		Else
			aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
		End If
	End If

	If IsEmpty(aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT)) Then
		If Len(oRequest("EndForValueYear").Item) > 0 Then
			aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) = CLng(oRequest("EndForValueYear").Item & Right(("0" & oRequest("EndForValueMonth").Item), Len("00")) & Right(("0" & oRequest("EndForValueDay").Item), Len("00")))
		ElseIf Len(oRequest("EndDateForValue").Item) > 0 Then
			aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) = CLng(oRequest("EndDateForValue").Item)
		Else
			aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) = 30000000
		End If
	End If
	If CLng(aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT)) = 0 Then aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) = 30000000

	If IsEmpty(aConceptComponent(N_CONCEPT_QTTY_ID_CONCEPT)) Then
		If Len(oRequest("ConceptQttyID").Item) > 0 Then
			aConceptComponent(N_CONCEPT_QTTY_ID_CONCEPT) = CInt(oRequest("ConceptQttyID").Item)
		Else
			aConceptComponent(N_CONCEPT_QTTY_ID_CONCEPT) = 1
		End If
	End If

	If IsEmpty(aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT)) Then
		If Len(oRequest("ConceptAmount").Item) > 0 Then
			aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT) = CDbl(oRequest("ConceptAmount").Item)
			If CInt(Request.Cookies("SIAP_SectionID")) <> 4 Then
				aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT) = CDbl(aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT)/2)
			End If
		Else
			aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_CURRENCY_ID_CONCEPT)) Then
		If Len(oRequest("CurrencyID").Item) > 0 Then
			aConceptComponent(N_CURRENCY_ID_CONCEPT) = CLng(oRequest("CurrencyID").Item)
		Else
			aConceptComponent(N_CURRENCY_ID_CONCEPT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_CONCEPT_TYPE_ID_CONCEPT)) Then
		If Len(oRequest("ConceptTypeID").Item) > 0 Then
			aConceptComponent(N_CONCEPT_TYPE_ID_CONCEPT) = CInt(oRequest("ConceptTypeID").Item)
		Else
			aConceptComponent(N_CONCEPT_TYPE_ID_CONCEPT) = 1
		End If
	End If

	If IsEmpty(aConceptComponent(S_APPLIES_ID_CONCEPT)) Then
		If Len(oRequest("AppliesToID").Item) > 0 Then
			aConceptComponent(S_APPLIES_ID_CONCEPT) = Replace(oRequest("AppliesToID").Item, " ", "")
		Else
			aConceptComponent(S_APPLIES_ID_CONCEPT) = "-1"
		End If
	End If
	aConceptComponent(S_APPLIES_ID_CONCEPT) = Left(aConceptComponent(S_APPLIES_ID_CONCEPT), 255)

	If IsEmpty(aConceptComponent(D_CONCEPT_MIN_CONCEPT)) Then
		If Len(oRequest("ConceptMin").Item) > 0 Then
			aConceptComponent(D_CONCEPT_MIN_CONCEPT) = CDbl(oRequest("ConceptMin").Item)
		Else
			aConceptComponent(D_CONCEPT_MIN_CONCEPT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_CONCEPT_MIN_QTTY_ID_CONCEPT)) Then
		If Len(oRequest("ConceptMinQttyID").Item) > 0 Then
			aConceptComponent(N_CONCEPT_MIN_QTTY_ID_CONCEPT) = CInt(oRequest("ConceptMinQttyID").Item)
		Else
			aConceptComponent(N_CONCEPT_MIN_QTTY_ID_CONCEPT) = 1
		End If
	End If

	If IsEmpty(aConceptComponent(D_CONCEPT_MAX_CONCEPT)) Then
		If Len(oRequest("ConceptMax").Item) > 0 Then
			aConceptComponent(D_CONCEPT_MAX_CONCEPT) = CDbl(oRequest("ConceptMax").Item)
		Else
			aConceptComponent(D_CONCEPT_MAX_CONCEPT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_CONCEPT_MAX_QTTY_ID_CONCEPT)) Then
		If Len(oRequest("ConceptMaxQttyID").Item) > 0 Then
			aConceptComponent(N_CONCEPT_MAX_QTTY_ID_CONCEPT) = CInt(oRequest("ConceptMaxQttyID").Item)
		Else
			aConceptComponent(N_CONCEPT_MAX_QTTY_ID_CONCEPT) = 1
		End If
	End If

	If IsEmpty(aConceptComponent(N_POSITION_ID_CONCEPT)) Then
		If Len(oRequest("PositionID").Item) > 0 Then
			aConceptComponent(N_POSITION_ID_CONCEPT) = CLng(oRequest("PositionID").Item)
		Else
			aConceptComponent(N_POSITION_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_START_USER_ID_CONCEPT)) Then
		If Len(oRequest("StartUserID").Item) > 0 Then
			aConceptComponent(N_START_USER_ID_CONCEPT) = CLng(oRequest("StartUserID").Item)
		Else
			aConceptComponent(N_START_USER_ID_CONCEPT) = aLoginComponent(N_USER_ID_LOGIN)
		End If
	End If

	If IsEmpty(aConceptComponent(N_END_USER_ID_CONCEPT)) Then
		If Len(oRequest("EndUserID").Item) > 0 Then
			aConceptComponent(N_END_USER_ID_CONCEPT) = CLng(oRequest("EndUserID").Item)
		Else
			aConceptComponent(N_END_USER_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_STATUS_ID_CONCEPT)) Then
		If Len(oRequest("StatusID").Item) > 0 Then
			aConceptComponent(N_STATUS_ID_CONCEPT) = CLng(oRequest("StatusID").Item)
		Else
			aConceptComponent(N_STATUS_ID_CONCEPT) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_CENTER_TYPE_ID)) Then
		If Len(oRequest("CenterTypeID").Item) > 0 Then
			aConceptComponent(N_CENTER_TYPE_ID) = CLng(oRequest("CenterTypeID").Item)
		Else
			aConceptComponent(N_CENTER_TYPE_ID) = -1
		End If
	End If

	If IsEmpty(aConceptComponent(N_IS_ACTIVE1)) Then
		If Len(oRequest("IsActive1").Item) > 0 Then
			aConceptComponent(N_IS_ACTIVE1) = CLng(oRequest("IsActive1").Item)
		Else
			aConceptComponent(N_IS_ACTIVE1) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_IS_ACTIVE2)) Then
		If Len(oRequest("IsActive2").Item) > 0 Then
			aConceptComponent(N_IS_ACTIVE2) = CLng(oRequest("IsActive2").Item)
		Else
			aConceptComponent(N_IS_ACTIVE2) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_IS_ACTIVE3)) Then
		If Len(oRequest("IsActive3").Item) > 0 Then
			aConceptComponent(N_IS_ACTIVE3) = CLng(oRequest("IsActive3").Item)
		Else
			aConceptComponent(N_IS_ACTIVE3) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_IS_ACTIVE4)) Then
		If Len(oRequest("IsActive4").Item) > 0 Then
			aConceptComponent(N_IS_ACTIVE4) = CLng(oRequest("IsActive4").Item)
		Else
			aConceptComponent(N_IS_ACTIVE4) = 0
		End If
	End If

	If aConceptComponent(N_IS_DEDUCTION_CONCEPT) = 0 Then
		If aConceptComponent(D_TAX_AMOUNT_CONCEPT) = 100 Then
			aConceptComponent(N_FOR_TAX_CONCEPT) = 1
		Else
			aConceptComponent(N_FOR_TAX_CONCEPT) = 0
		End If
	Else
		If aConceptComponent(D_TAX_AMOUNT_CONCEPT) = 100 Then
			aConceptComponent(N_FOR_TAX_CONCEPT) = 0
		Else
			aConceptComponent(N_FOR_TAX_CONCEPT) = 1
		End If
	End If

	If IsEmpty(aConceptComponent(N_IS_CREDIT)) Then
		If Len(oRequest("IsCredit").Item) > 0 Then
			aConceptComponent(N_IS_CREDIT) = CLng(oRequest("IsCredit").Item)
		Else
			aConceptComponent(N_IS_CREDIT) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_IS_OTHER)) Then
		If Len(oRequest("IsOther").Item) > 0 Then
			aConceptComponent(N_IS_OTHER) = CLng(oRequest("IsOther").Item)
		Else
			aConceptComponent(N_IS_OTHER) = 0
		End If
	End If

	If IsEmpty(aConceptComponent(N_ACTIVE)) Then
		If Len(oRequest("Active").Item) > 0 Then
			aConceptComponent(N_ACTIVE) = CLng(oRequest("Active").Item)
		Else
			aConceptComponent(N_ACTIVE) = 0
		End If
	End If

	aConceptComponent(S_QUERY_CONDITION_CONCEPT) = ""
	aConceptComponent(B_CHECK_FOR_DUPLICATED_CONCEPT) = True
	aConceptComponent(B_IS_DUPLICATED_CONCEPT) = False

	aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT) = True
	InitializeConceptComponent = Err.number
	Err.Clear
End Function

Function AuthorizeConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To authorize the concept value in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AuthorizeConceptValue"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro a autorizar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConceptsValues Set StatusID=1" & ", AuthorizationDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & " Where (RecordID=" & aConceptComponent(N_RECORD_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	Set oRecordset = Nothing
	AuthorizeConceptValue = lErrorNumber
	Err.Clear
End Function

Function GetConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a concept from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetConcept"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sNames

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If aConceptComponent(N_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Concepts Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (StartDate=" & aConceptComponent(N_START_DATE_CONCEPT) & ") Order By StartDate Desc", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
				oRecordset.Close
			Else
				aConceptComponent(N_START_DATE_CONCEPT) = CLng(oRecordset.Fields("StartDate").Value)
				aConceptComponent(N_END_DATE_CONCEPT) = CLng(oRecordset.Fields("EndDate").Value)
				aConceptComponent(S_SHORT_NAME_CONCEPT) = CStr(oRecordset.Fields("ConceptShortName").Value)
				aConceptComponent(S_NAME_CONCEPT) = CStr(oRecordset.Fields("ConceptName").Value)
				aConceptComponent(N_BUDGET_ID_CONCEPT) = CLng(oRecordset.Fields("BudgetID").Value)
				aConceptComponent(N_PAYROLL_TYPE_ID_CONCEPT) = CLng(oRecordset.Fields("PayrollTypeID").Value)
				aConceptComponent(N_PERIOD_ID_CONCEPT) = CLng(oRecordset.Fields("PeriodID").Value)
				aConceptComponent(N_IS_DEDUCTION_CONCEPT) = CInt(oRecordset.Fields("IsDeduction").Value)
				aConceptComponent(N_FOR_ALIMONY_CONCEPT) = CInt(oRecordset.Fields("ForAlimony").Value)
				aConceptComponent(N_ON_LEAVE_CONCEPT) = CInt(oRecordset.Fields("OnLeave").Value)
				aConceptComponent(N_ORDER_IN_LIST_CONCEPT) = CLng(oRecordset.Fields("OrderInList").Value)
				aConceptComponent(D_TAX_AMOUNT_CONCEPT) = CDbl(oRecordset.Fields("TaxAmount").Value)
				aConceptComponent(N_TAX_CURRENCY_ID_CONCEPT) = CLng(oRecordset.Fields("TaxCurrencyID").Value)
				aConceptComponent(N_TAX_QTTY_ID_CONCEPT) = CInt(oRecordset.Fields("TaxQttyID").Value)
				aConceptComponent(D_TAX_MIN_CONCEPT) = CDbl(oRecordset.Fields("TaxMin").Value)
				aConceptComponent(N_TAX_MIN_QTTY_ID_CONCEPT) = CInt(oRecordset.Fields("TaxMinQttyID").Value)
				aConceptComponent(D_TAX_MAX_CONCEPT) = CDbl(oRecordset.Fields("TaxMax").Value)
				aConceptComponent(N_TAX_MAX_QTTY_ID_CONCEPT) = CInt(oRecordset.Fields("TaxMaxQttyID").Value)
				aConceptComponent(D_EXEMPT_AMOUNT_CONCEPT) = CDbl(oRecordset.Fields("ExemptAmount").Value)
				aConceptComponent(N_EXEMPT_CURRENCY_ID_CONCEPT) = CLng(oRecordset.Fields("ExemptCurrencyID").Value)
				aConceptComponent(N_EXEMPT_QTTY_ID_CONCEPT) = CInt(oRecordset.Fields("ExemptQttyID").Value)
				aConceptComponent(D_EXEMPT_MIN_CONCEPT) = CDbl(oRecordset.Fields("ExemptMin").Value)
				aConceptComponent(N_EXEMPT_MIN_QTTY_ID_CONCEPT) = CInt(oRecordset.Fields("ExemptMinQttyID").Value)
				aConceptComponent(D_EXEMPT_MAX_CONCEPT) = CDbl(oRecordset.Fields("ExemptMax").Value)
				aConceptComponent(N_EXEMPT_MAX_QTTY_ID_CONCEPT) = CInt(oRecordset.Fields("ExemptMaxQttyID").Value)
			End If
		End If
		oRecordset.Close
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From CreditTypes Where (CreditTypeID=" & aConceptComponent(N_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aConceptComponent(N_IS_CREDIT) = CInt(oRecordset.Fields("CreditTypeID").Value)
				aConceptComponent(N_IS_OTHER) = CInt(oRecordset.Fields("IsOther").Value)
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetConcept = lErrorNumber
	Err.Clear
End Function

Function GetConceptCrossType(oADODBConnection, aConceptComponent, sConceptCrossType, sErrorDescription)
'************************************************************
'Purpose: To get the type of crossing for the
'         record to insert
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sEmployeeConceptType, lStartDate, lEndDate, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetConceptCrossType"
	Dim oRecordset
	Dim lErrorNumber
	Dim sQuery

	sQuery = "Select * From Concepts Where (ConceptShortName='" & Replace(aConceptComponent(S_SHORT_NAME_CONCEPT), "'", "") & "')" & _
			 " And (StartDate<" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate>" & aConceptComponent(N_END_DATE_CONCEPT) & ") Order By StartDate Desc"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sConceptCrossType = "Cross"
		Else
			sQuery = "Select * From Concepts Where (ConceptShortName='" & Replace(aConceptComponent(S_SHORT_NAME_CONCEPT), "'", "") & "')" & _
					 " And (StartDate>" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate<" & aConceptComponent(N_END_DATE_CONCEPT) & ") Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sConceptCrossType = "Inner"
				Else
					sQuery = "Select * From Concepts Where (ConceptShortName='" & Replace(aConceptComponent(S_SHORT_NAME_CONCEPT), "'", "") & "')" & _
							 " And (StartDate<" & aConceptComponent(N_START_DATE_CONCEPT) & ") And ((EndDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ") And (EndDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate Desc"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sConceptCrossType = "Left"
						Else
							sQuery = "Select * From Concepts Where (ConceptShortName='" & Replace(aConceptComponent(S_SHORT_NAME_CONCEPT), "'", "") & "')" & _
									 " And (StartDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ") And ((EndDate>=" & aConceptComponent(N_END_DATE_CONCEPT) & ") And (StartDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									sConceptCrossType = "Right"
								End If
							Else
								lErrorNumber = -1
								sErrorDescription = "No se pudo verifiar si el registro se empalma con otros puestos para guardías y suplencias."
							End If
						End If
					Else
						lErrorNumber = -1
						sErrorDescription = "No se pudo verifiar si el registro se empalma con otros puestos para guardías y suplencias."
					End If
				End If
			Else
				lErrorNumber = -1
				sErrorDescription = "No se pudo verifiar si el registro se empalma con otros puestos para guardías y suplencias."
			End If
		End If
	Else
		lErrorNumber = -1
		sErrorDescription = "No se pudo verifiar si el registro se empalma con otros puestos para guardías y suplencias."
	End If

	Set oRecordset = Nothing
	GetConceptCrossType = lErrorNumber
	Err.Clear
End Function

Function GetConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a concept from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetConceptValue"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If (Len(oRequest("RecordID").Item) = 0) And (aConceptComponent(N_RECORD_ID_CONCEPT) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		'If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
		'	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From ConceptsValues Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ") And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (PositionTypeID=" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ") And (EmployeeStatusID=" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ") And (JobStatusID=" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ") And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EconomicZoneID=" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (AntiquityID=" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ") And (AdditionalShift=" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ") And (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (EndDate=30000000)", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		'Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From ConceptsValues Where (RecordID=" & aConceptComponent(N_RECORD_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		'End If
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aConceptComponent(N_RECORD_ID_CONCEPT) = CLng(oRecordset.Fields("RecordID").Value)
				aConceptComponent(N_ID_CONCEPT) = CLng(oRecordset.Fields("ConceptID").Value)
				aConceptComponent(N_COMPANY_ID_CONCEPT) = CLng(oRecordset.Fields("CompanyID").Value)
				aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = CLng(oRecordset.Fields("EmployeeTypeID").Value)
				aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) = CLng(oRecordset.Fields("PositionTypeID").Value)
				aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) = CLng(oRecordset.Fields("EmployeeStatusID").Value)
				aConceptComponent(N_JOB_STATUS_ID_CONCEPT) = CLng(oRecordset.Fields("JobStatusID").Value)
				aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) = CLng(oRecordset.Fields("ClassificationID").Value)
				aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) = CLng(oRecordset.Fields("GroupGradeLevelID").Value)
				aConceptComponent(N_INTEGRATION_ID_CONCEPT) = CLng(oRecordset.Fields("IntegrationID").Value)
				aConceptComponent(N_JOURNEY_ID_CONCEPT) = CLng(oRecordset.Fields("JourneyID").Value)
				aConceptComponent(D_WORKING_HOURS_CONCEPT) = CDbl(oRecordset.Fields("WorkingHours").Value)
				aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) = CLng(oRecordset.Fields("AdditionalShift").Value)
				aConceptComponent(N_LEVEL_ID_CONCEPT) = CLng(oRecordset.Fields("LevelID").Value)
				aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) = CInt(oRecordset.Fields("EconomicZoneID").Value)
				aConceptComponent(N_SERVICE_ID_CONCEPT) = CLng(oRecordset.Fields("ServiceID").Value)
				aConceptComponent(N_ANTIQUITY_ID_CONCEPT) = CLng(oRecordset.Fields("AntiquityID").Value)
				aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) = CLng(oRecordset.Fields("Antiquity2ID").Value)
				aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) = CLng(oRecordset.Fields("Antiquity3ID").Value)
				aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) = CLng(oRecordset.Fields("Antiquity4ID").Value)
				aConceptComponent(N_FOR_RISK_CONCEPT) = CInt(oRecordset.Fields("ForRisk").Value)
				aConceptComponent(N_GENDER_ID_CONCEPT) = CInt(oRecordset.Fields("GenderID").Value)
				aConceptComponent(N_HAS_CHILDREN_CONCEPT) = CInt(oRecordset.Fields("HasChildren").Value)
				aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) = CInt(oRecordset.Fields("SchoolarshipID").Value)
				aConceptComponent(N_HAS_SYNDICATE_CONCEPT) = CInt(oRecordset.Fields("HasSyndicate").Value)
				aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) = CLng(oRecordset.Fields("StartDate").Value)
				aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) = CLng(oRecordset.Fields("EndDate").Value)
				aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT) = CDbl(oRecordset.Fields("ConceptAmount").Value)
				aConceptComponent(N_CURRENCY_ID_CONCEPT) = CLng(oRecordset.Fields("CurrencyID").Value)
				aConceptComponent(N_CONCEPT_QTTY_ID_CONCEPT) = CInt(oRecordset.Fields("ConceptQttyID").Value)
				aConceptComponent(N_CONCEPT_TYPE_ID_CONCEPT) = CInt(oRecordset.Fields("ConceptTypeID").Value)
				aConceptComponent(S_APPLIES_ID_CONCEPT) = CStr(oRecordset.Fields("AppliesToID").Value)
				aConceptComponent(D_CONCEPT_MIN_CONCEPT) = CStr(oRecordset.Fields("ConceptMin").Value)
				aConceptComponent(N_CONCEPT_MIN_QTTY_ID_CONCEPT) = CStr(oRecordset.Fields("ConceptMinQttyID").Value)
				aConceptComponent(D_CONCEPT_MAX_CONCEPT) = CStr(oRecordset.Fields("ConceptMax").Value)
				aConceptComponent(N_CONCEPT_MAX_QTTY_ID_CONCEPT) = CStr(oRecordset.Fields("ConceptMaxQttyID").Value)
				aConceptComponent(N_POSITION_ID_CONCEPT) = CLng(oRecordset.Fields("PositionID").Value)
				aConceptComponent(N_STATUS_ID_CONCEPT) = CInt(oRecordset.Fields("StatusID").Value)
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetConceptValue = lErrorNumber
	Err.Clear
End Function

Function GetConceptValues(oRequest, oADODBConnection, aConceptComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all concepts values 
'         from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetConceptValues"
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	Call GetStartAndEndDatesFromURL("FilterStart", "FilterEnd", "OcurredDate", False, sCondition)
	sCondition = sCondition & aConceptComponent(S_QUERY_CONDITION_CONCEPT)

	sCondition  = Trim(sCondition)
	If Len(sCondition ) > 0 Then
		If InStr(1, sCondition , "And ", vbBinaryCompare) <> 1 Then sCondition  = "And " & sCondition
	End If

	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptsValues.RecordID, ConceptsValues.ConceptID, ConceptsValues.ConceptAmount, ConceptsValues.StartDate, ConceptsValues.EndDate, ConceptsValues.StatusID, Positions.PositionID, ConceptsValues.LevelID, Levels.LevelShortName, ConceptsValues.EconomicZoneID, ConceptsValues.ClassificationID, ConceptsValues.IntegrationID, ConceptsValues.GroupGradeLevelID, GroupGradeLevels.GroupGradeLevelShortName, ConceptsValues.WorkingHours, ConceptsValues.AntiquityID, ConceptsValues.Antiquity2ID, Positions.PositionShortName, Positions.PositionName, PositionTypes.PositionTypeID, PositionTypes.PositionTypeShortName, PositionTypes.PositionTypeName, ConceptsValues.EmployeeTypeID, EmployeeTypes.EmployeeTypeName From ConceptsValues, Positions, PositionTypes, GroupGradeLevels, Levels, EmployeeTypes Where (ConceptsValues.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (ConceptsValues.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (ConceptsValues.LevelID=Levels.LevelID) And (ConceptsValues.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) " & sCondition & " Order by PositionID, StartDate, LevelID, ClassificationID, IntegrationID, GroupGradeLevelID, WorkingHours, ConceptsValues.PositionTypeID, EconomicZoneID", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""Select ConceptsValues.RecordID, ConceptsValues.ConceptID, ConceptsValues.ConceptAmount, ConceptsValues.StartDate, ConceptsValues.EndDate, ConceptsValues.StatusID, Positions.PositionID, ConceptsValues.LevelID, Levels.LevelShortName, ConceptsValues.EconomicZoneID, ConceptsValues.ClassificationID, ConceptsValues.IntegrationID, ConceptsValues.GroupGradeLevelID, GroupGradeLevels.GroupGradeLevelShortName, ConceptsValues.WorkingHours, ConceptsValues.AntiquityID, ConceptsValues.Antiquity2ID, Positions.PositionShortName, Positions.PositionName, PositionTypes.PositionTypeID, PositionTypes.PositionTypeShortName, PositionTypes.PositionTypeName, EmployeeTypes.EmployeeTypeName From ConceptsValues, Positions, PositionTypes, GroupGradeLevels, Levels, EmployeeTypes Where (ConceptsValues.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (ConceptsValues.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (ConceptsValues.LevelID=Levels.LevelID) And (ConceptsValues.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) " & sCondition & " Order by PositionID, StartDate, LevelID, ClassificationID, IntegrationID, GroupGradeLevelID, WorkingHours, ConceptsValues.PositionTypeID, EconomicZoneID"" />"

	GetConceptValues = lErrorNumber
	Err.Clear
End Function

Function GetConcepts(oRequest, oADODBConnection, aConceptComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the concepts from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetConcepts"
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If (Len(aConceptComponent(S_QUERY_CONDITION_CONCEPT)) > 0) Then
		sCondition = Trim(aConceptComponent(S_QUERY_CONDITION_CONCEPT))
		If InStr(1, sCondition, "And ", vbTextCompare) <> 1 Then
			sCondition = "And " & sCondition
		End If
	End If
	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Concepts Where (ConceptID>-760211) " & sCondition & " Order By ConceptShortName, StartDate", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetConcepts = lErrorNumber
	Err.Clear
End Function

Function GetConceptValueCrossType(oADODBConnection, aConceptComponent, sConceptValuesCrossType, sErrorDescription)
'************************************************************
'Purpose: To get the type of crossing for the
'         record to insert
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sEmployeeConceptType, lStartDate, lEndDate, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetConceptValueCrossType"
	Dim oRecordset
	Dim lErrorNumber
	Dim sQuery

	sQuery = "Select * From ConceptsValues Where (RecordID<>" & aConceptComponent(N_RECORD_ID_CONCEPT) & ") And (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ")" & _
			" And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (PositionTypeID=" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ") And (EmployeeStatusID=" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ")" & _
			" And (JobStatusID=" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ")" & _
			" And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (JourneyID=" & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ")" & _
			" And (AdditionalShift=" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EconomicZoneID=" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ")" & _
			" And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (AntiquityID=" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ") And (Antiquity2ID=" & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ")" & _
			" And (Antiquity3ID=" & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ") And (Antiquity4ID=" & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ") And (ForRisk=" & aConceptComponent(N_FOR_RISK_CONCEPT) & ")" & _
			" And (GenderID=" & aConceptComponent(N_GENDER_ID_CONCEPT) & ") And (HasChildren=" & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ") And (SchoolarshipID=" & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ")" & _
			" And (HasSyndicate=" & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ") And (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ")" & _
			" And (StartDate<" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ") And (EndDate>" & aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & ") Order By StartDate Desc"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sConceptValuesCrossType = "Cross"
		Else
			sQuery = "Select * From ConceptsValues Where (RecordID<>" & aConceptComponent(N_RECORD_ID_CONCEPT) & ") And (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ")" & _
					" And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (PositionTypeID=" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ") And (EmployeeStatusID=" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ")" & _
					" And (JobStatusID=" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ")" & _
					" And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (JourneyID=" & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ")" & _
					" And (AdditionalShift=" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EconomicZoneID=" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ")" & _
					" And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (AntiquityID=" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ") And (Antiquity2ID=" & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ")" & _
					" And (Antiquity3ID=" & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ") And (Antiquity4ID=" & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ") And (ForRisk=" & aConceptComponent(N_FOR_RISK_CONCEPT) & ")" & _
					" And (GenderID=" & aConceptComponent(N_GENDER_ID_CONCEPT) & ") And (HasChildren=" & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ") And (SchoolarshipID=" & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ")" & _
					" And (HasSyndicate=" & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ") And (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ")" & _
					" And (StartDate>" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ") And (EndDate<" & aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & ") Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sConceptValuesCrossType = "Inner"
				Else
					sQuery = "Select * From ConceptsValues Where (RecordID<>" & aConceptComponent(N_RECORD_ID_CONCEPT) & ") And (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ")" & _
							" And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (PositionTypeID=" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ") And (EmployeeStatusID=" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ")" & _
							" And (JobStatusID=" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ")" & _
							" And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (JourneyID=" & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ")" & _
							" And (AdditionalShift=" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EconomicZoneID=" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ")" & _
							" And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (AntiquityID=" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ") And (Antiquity2ID=" & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ")" & _
							" And (Antiquity3ID=" & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ") And (Antiquity4ID=" & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ") And (ForRisk=" & aConceptComponent(N_FOR_RISK_CONCEPT) & ")" & _
							" And (GenderID=" & aConceptComponent(N_GENDER_ID_CONCEPT) & ") And (HasChildren=" & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ") And (SchoolarshipID=" & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ")" & _
							" And (HasSyndicate=" & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ") And (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ")" & _
							" And (StartDate<" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ") And ((EndDate<=" & aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & ") And (EndDate>=" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate Desc"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sConceptValuesCrossType = "Left"
						Else
							sQuery = "Select * From ConceptsValues Where (RecordID<>" & aConceptComponent(N_RECORD_ID_CONCEPT) & ") And (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ")" & _
									" And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (PositionTypeID=" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ") And (EmployeeStatusID=" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ")" & _
									" And (JobStatusID=" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ")" & _
									" And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (JourneyID=" & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ")" & _
									" And (AdditionalShift=" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EconomicZoneID=" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ")" & _
									" And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (AntiquityID=" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ") And (Antiquity2ID=" & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ")" & _
									" And (Antiquity3ID=" & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ") And (Antiquity4ID=" & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ") And (ForRisk=" & aConceptComponent(N_FOR_RISK_CONCEPT) & ")" & _
									" And (GenderID=" & aConceptComponent(N_GENDER_ID_CONCEPT) & ") And (HasChildren=" & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ") And (SchoolarshipID=" & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ")" & _
									" And (HasSyndicate=" & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ") And (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ")" & _
									" And (StartDate>=" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ") And ((EndDate>=" & aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & ") And (StartDate<=" & aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									sConceptValuesCrossType = "Right"
								End If
							Else
								lErrorNumber = -1
								sErrorDescription = "No se pudo verifiar si el registro se empalma con otros puestos para guardías y suplencias."
							End If
						End If
					Else
						lErrorNumber = -1
						sErrorDescription = "No se pudo verifiar si el registro se empalma con otros puestos para guardías y suplencias."
					End If
				End If
			Else
				lErrorNumber = -1
				sErrorDescription = "No se pudo verifiar si el registro se empalma con otros puestos para guardías y suplencias."
			End If
		End If
	Else
		lErrorNumber = -1
		sErrorDescription = "No se pudo verifiar si el registro se empalma con otros puestos para guardías y suplencias."
	End If

	Set oRecordset = Nothing
	GetConceptValueCrossType = lErrorNumber
	Err.Clear
End Function

Function GetPositions(oRequest, oADODBConnection, aConceptComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the concepts from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPositions"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sCondition

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	sErrorDescription = "No se pudo obtener la información de los puestos."
	If Len(aConceptComponent(S_QUERY_CONDITION_CONCEPT)) > 0 Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Positions Where " & aConceptComponent(S_QUERY_CONDITION_CONCEPT), "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Positions", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If
	'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""Select ConceptsValues.RecordID, ConceptsValues.ConceptID, ConceptsValues.ConceptAmount, ConceptsValues.StartDate, ConceptsValues.EndDate, ConceptsValues.StatusID, Positions.PositionID, ConceptsValues.LevelID, Levels.LevelShortName, ConceptsValues.EconomicZoneID, ConceptsValues.ClassificationID, ConceptsValues.IntegrationID, ConceptsValues.GroupGradeLevelID, GroupGradeLevels.GroupGradeLevelShortName, ConceptsValues.WorkingHours, ConceptsValues.AntiquityID, ConceptsValues.Antiquity2ID, Positions.PositionShortName, Positions.PositionName, PositionTypes.PositionTypeID, PositionTypes.PositionTypeShortName, PositionTypes.PositionTypeName From ConceptsValues, Positions, PositionTypes, GroupGradeLevels, Levels Where (ConceptsValues.PositionID=Positions.PositionID) And (Positions.PositionTypeID=PositionTypes.PositionTypeID) And (ConceptsValues.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (ConceptsValues.LevelID=Levels.LevelID)" & sCondition & " Order by PositionID, StartDate, LevelID, ClassificationID, IntegrationID, GroupGradeLevelID, WorkingHours, ConceptsValues.PositionTypeID, EconomicZoneID"" />"

	GetPositions = lErrorNumber
	Err.Clear
End Function

Function GetPositionDataForSpecialJourneysLKP(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the concepts from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPositionDataForSpecialJourneysLKP"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If aConceptComponent(N_POSITION_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del puesto para obtener el nivel y la jornada a la que aplica."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información de los registros."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Positions Where (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El puesto especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
				oRecordset.Close
			Else
				aConceptComponent(N_LEVEL_ID_CONCEPT) = oRecordset.Fields("LevelID").Value
				aConceptComponent(D_WORKING_HOURS_CONCEPT) = oRecordset.Fields("WorkingHours").Value
			End If
		Else
			lErrorNumber = -1
			sErrorDescription = "Error al obtener el puesto especificado."
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
		End If
	End If

	GetPositionDataForSpecialJourneysLKP = lErrorNumber
	Err.Clear
End Function

Function GetPositionSpecialJourneyCrossingType(oADODBConnection, aConceptComponent, sPositionSpecialJourneyCrossingType, sErrorDescription)
'************************************************************
'Purpose: To get the type of crossing for the
'         record to insert
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sEmployeeConceptType, lStartDate, lEndDate, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPositionSpecialJourneyCrossingType"
	Dim oRecordset
	Dim lErrorNumber
	Dim sQuery

	sQuery = "Select * From PositionsSpecialJourneysLKP Where (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ")" & _
			 " And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (CenterTypeID=" & aConceptComponent(N_CENTER_TYPE_ID) & ")" & _
			 " And (StartDate<" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate>" & aConceptComponent(N_END_DATE_CONCEPT) & ") Order By StartDate Desc"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sPositionSpecialJourneyCrossingType = "Cross"
		Else
			sQuery = "Select * From PositionsSpecialJourneysLKP Where (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ")" & _
					 " And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (CenterTypeID=" & aConceptComponent(N_CENTER_TYPE_ID) & ")" & _
					 " And (StartDate>" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate<" & aConceptComponent(N_END_DATE_CONCEPT) & ") Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sPositionSpecialJourneyCrossingType = "Inner"
				Else
					sQuery = "Select * From PositionsSpecialJourneysLKP Where (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ")" & _
							 " And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (CenterTypeID=" & aConceptComponent(N_CENTER_TYPE_ID) & ")" & _
							 " And (StartDate<" & aConceptComponent(N_START_DATE_CONCEPT) & ") And ((EndDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ") And (EndDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate Desc"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sPositionSpecialJourneyCrossingType = "Left"
						Else
							sQuery = "Select * From PositionsSpecialJourneysLKP Where (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ")" & _
									 " And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (CenterTypeID=" & aConceptComponent(N_CENTER_TYPE_ID) & ")" & _
									 " And (StartDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ") And ((EndDate>=" & aConceptComponent(N_END_DATE_CONCEPT) & ") And (StartDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									sPositionSpecialJourneyCrossingType = "Right"
								End If
							Else
								lErrorNumber = -1
								sErrorDescription = "No se pudo verifiar si el registro se empalma con otros puestos para guardías y suplencias."
							End If
						End If
					Else
						lErrorNumber = -1
						sErrorDescription = "No se pudo verifiar si el registro se empalma con otros puestos para guardías y suplencias."
					End If
				End If
			Else
				lErrorNumber = -1
				sErrorDescription = "No se pudo verifiar si el registro se empalma con otros puestos para guardías y suplencias."
			End If
		End If
	Else
		lErrorNumber = -1
		sErrorDescription = "No se pudo verifiar si el registro se empalma con otros puestos para guardías y suplencias."
	End If

	Set oRecordset = Nothing
	GetPositionSpecialJourneyCrossingType = lErrorNumber
	Err.Clear
End Function

Function GetPositionsSpecialJourneysLKP(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a concept from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPositionsSpecialJourneysLKP"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro de puesto para guardias y suplencias."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From PositionsSpecialJourneysLKP Where (RecordID=" & aConceptComponent(N_RECORD_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aConceptComponent(N_RECORD_ID_CONCEPT) = CLng(oRecordset.Fields("RecordID").Value)
				aConceptComponent(N_START_DATE_CONCEPT) = CLng(oRecordset.Fields("StartDate").Value)
				aConceptComponent(N_END_DATE_CONCEPT) = CLng(oRecordset.Fields("EndDate").Value)
				aConceptComponent(N_POSITION_ID_CONCEPT) = CLng(oRecordset.Fields("PositionID").Value)
				aConceptComponent(N_LEVEL_ID_CONCEPT) = CLng(oRecordset.Fields("LevelID").Value)
				aConceptComponent(D_WORKING_HOURS_CONCEPT) = CDbl(oRecordset.Fields("WorkingHours").Value)
				aConceptComponent(N_SERVICE_ID_CONCEPT) = CLng(oRecordset.Fields("ServiceID").Value)
				aConceptComponent(N_CENTER_TYPE_ID) = CLng(oRecordset.Fields("CenterTypeID").Value)
				aConceptComponent(N_IS_ACTIVE1) = CInt(oRecordset.Fields("IsActive1").Value)
				aConceptComponent(N_IS_ACTIVE2) = CInt(oRecordset.Fields("IsActive2").Value)
				aConceptComponent(N_IS_ACTIVE3) = CInt(oRecordset.Fields("IsActive3").Value)
				aConceptComponent(N_IS_ACTIVE4) = CInt(oRecordset.Fields("IsActive4").Value)
				aConceptComponent(N_ACTIVE) = CInt(oRecordset.Fields("Active").Value)
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetPositionsSpecialJourneysLKP = lErrorNumber
	Err.Clear
End Function

Function ConceptHasChanged(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a concept from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ConceptHasChanged"
	Dim oRecordset
	Dim oRecordset1
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim iIsCredit
	Dim iIsOther

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If aConceptComponent(N_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Concepts Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (StartDate=" & CLng(oRequest("StartDate").Item) & ") Order By StartDate Desc", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From CreditTypes Where (CreditTypeID=" & aConceptComponent(N_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset1)
				If lErrorNumber = 0 Then
					If Not oRecordset1.EOF Then
						iIsCredit = CInt(oRecordset1.Fields("CreditTypeID").Value)
						iIsOther = CInt(oRecordset1.Fields("IsOther").Value)
					Else
						iIsCredit = 0
						iIsOther = 0
					End If
					oRecordset1.Close
				End If
				If (aConceptComponent(S_NAME_CONCEPT) = CStr(oRecordset.Fields("ConceptName").Value)) And _
				(aConceptComponent(N_BUDGET_ID_CONCEPT) = CLng(oRecordset.Fields("BudgetID").Value)) And _
				(aConceptComponent(N_PAYROLL_TYPE_ID_CONCEPT) = CLng(oRecordset.Fields("PayrollTypeID").Value)) And _
				(aConceptComponent(N_PERIOD_ID_CONCEPT) = CLng(oRecordset.Fields("PeriodID").Value)) And _
				(aConceptComponent(N_IS_DEDUCTION_CONCEPT) = CInt(oRecordset.Fields("IsDeduction").Value)) And _
				(aConceptComponent(N_FOR_ALIMONY_CONCEPT) = CInt(oRecordset.Fields("ForAlimony").Value)) And _
				(aConceptComponent(N_ON_LEAVE_CONCEPT) = CInt(oRecordset.Fields("OnLeave").Value)) And _
				(aConceptComponent(N_ORDER_IN_LIST_CONCEPT) = CLng(oRecordset.Fields("OrderInList").Value)) And _
				(aConceptComponent(D_TAX_AMOUNT_CONCEPT) = CDbl(oRecordset.Fields("TaxAmount").Value)) And _
				(aConceptComponent(N_TAX_CURRENCY_ID_CONCEPT) = CLng(oRecordset.Fields("TaxCurrencyID").Value)) And _
				(aConceptComponent(N_TAX_QTTY_ID_CONCEPT) = CInt(oRecordset.Fields("TaxQttyID").Value)) And _
				(aConceptComponent(D_TAX_MIN_CONCEPT) = CDbl(oRecordset.Fields("TaxMin").Value)) And _
				(aConceptComponent(N_TAX_MIN_QTTY_ID_CONCEPT) = CInt(oRecordset.Fields("TaxMinQttyID").Value)) And _
				(aConceptComponent(D_TAX_MAX_CONCEPT) = CDbl(oRecordset.Fields("TaxMax").Value)) And _
				(aConceptComponent(N_TAX_MAX_QTTY_ID_CONCEPT) = CInt(oRecordset.Fields("TaxMaxQttyID").Value)) And _
				(aConceptComponent(D_EXEMPT_AMOUNT_CONCEPT) = CDbl(oRecordset.Fields("ExemptAmount").Value)) And _
				(aConceptComponent(N_EXEMPT_CURRENCY_ID_CONCEPT) = CLng(oRecordset.Fields("ExemptCurrencyID").Value)) And _
				(aConceptComponent(N_EXEMPT_QTTY_ID_CONCEPT) = CInt(oRecordset.Fields("ExemptQttyID").Value)) And _
				(aConceptComponent(D_EXEMPT_MIN_CONCEPT) = CDbl(oRecordset.Fields("ExemptMin").Value)) And _
				(aConceptComponent(N_EXEMPT_MIN_QTTY_ID_CONCEPT) = CInt(oRecordset.Fields("ExemptMinQttyID").Value)) And _
				(aConceptComponent(D_EXEMPT_MAX_CONCEPT) = CDbl(oRecordset.Fields("ExemptMax").Value)) And _
				(aConceptComponent(N_EXEMPT_MAX_QTTY_ID_CONCEPT) = CInt(oRecordset.Fields("ExemptMaxQttyID").Value)) And _
				(aConceptComponent(N_START_DATE_CONCEPT) = CLng(oRecordset.Fields("StartDate").Value)) And _
				(aConceptComponent(N_END_DATE_CONCEPT) = CLng(oRecordset.Fields("EndDate").Value)) And _
				(aConceptComponent(N_IS_CREDIT) = iIsCredit) And _
				(aConceptComponent(N_IS_OTHER) = iIsOther) _
				Then
					ConceptHasChanged = False
				Else
					ConceptHasChanged = True
				End If
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function ConceptValueHasChanged(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a concept from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ConceptValueHasChanged"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From ConceptsValues Where (RecordID=" & aConceptComponent(N_RECORD_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
				oRecordset.Close
			Else
				If (aConceptComponent(N_ID_CONCEPT) = CStr(oRecordset.Fields("ConceptID").Value)) And _
				(aConceptComponent(N_COMPANY_ID_CONCEPT) = CStr(oRecordset.Fields("CompanyID").Value)) And _
				(aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = CLng(oRecordset.Fields("EmployeeTypeID").Value)) And _
				(aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) = CLng(oRecordset.Fields("PositionTypeID").Value)) And _
				(aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) = CLng(oRecordset.Fields("EmployeeStatusID").Value)) And _
				(aConceptComponent(N_JOB_STATUS_ID_CONCEPT) = CInt(oRecordset.Fields("JobStatusID").Value)) And _
				(aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) = CInt(oRecordset.Fields("ClassificationID").Value)) And _
				(aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) = CInt(oRecordset.Fields("GroupGradeLevelID").Value)) And _
				(aConceptComponent(N_INTEGRATION_ID_CONCEPT) = CInt(oRecordset.Fields("IntegrationID").Value)) And _
				(aConceptComponent(N_JOURNEY_ID_CONCEPT) = CLng(oRecordset.Fields("JourneyID").Value)) And _
				(aConceptComponent(D_WORKING_HOURS_CONCEPT) = CDbl(oRecordset.Fields("WorkingHours").Value)) And _
				(aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) = CDbl(oRecordset.Fields("AdditionalShift").Value)) And _
				(aConceptComponent(N_LEVEL_ID_CONCEPT) = CLng(oRecordset.Fields("LevelID").Value)) And _
				(aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) = CInt(oRecordset.Fields("EconomicZoneID").Value)) And _
				(aConceptComponent(N_SERVICE_ID_CONCEPT) = CDbl(oRecordset.Fields("ServiceID").Value)) And _
				(aConceptComponent(N_ANTIQUITY_ID_CONCEPT) = CInt(oRecordset.Fields("AntiquityID").Value)) And _
				(aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) = CDbl(oRecordset.Fields("Antiquity2ID").Value)) And _
				(aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) = CInt(oRecordset.Fields("Antiquity3ID").Value)) And _
				(aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) = CDbl(oRecordset.Fields("Antiquity4ID").Value)) And _
				(aConceptComponent(N_FOR_RISK_CONCEPT) = CLng(oRecordset.Fields("ForRisk").Value)) And _
				(aConceptComponent(N_GENDER_ID_CONCEPT) = CInt(oRecordset.Fields("GenderID").Value)) And _
				(aConceptComponent(N_HAS_CHILDREN_CONCEPT) = CDbl(oRecordset.Fields("HasChildren").Value)) And _
				(aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) = CInt(oRecordset.Fields("SchoolarshipID").Value)) And _
				(aConceptComponent(N_HAS_SYNDICATE_CONCEPT) = CDbl(oRecordset.Fields("HasSyndicate").Value)) And _
				(aConceptComponent(N_POSITION_ID_CONCEPT) = CInt(oRecordset.Fields("PositionID").Value)) And _
				(aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) = CLng(oRecordset.Fields("StartDate").Value)) And _
				(aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) = CLng(oRecordset.Fields("EndDate").Value)) Then
					ConceptValueHasChanged = False
				Else
					ConceptValueHasChanged = True
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function ModifyConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing concept in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyConcept"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim sSpecialCondition
	Dim bHasChanged

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If aConceptComponent(N_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Not CheckConceptInformationConsistency(aConceptComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			If aConceptComponent(N_END_DATE_CONCEPT) = 0 Then aConceptComponent(N_END_DATE_CONCEPT) = 30000000
			bHasChanged = True
			If Not ConceptHasChanged(oRequest, oADODBConnection, aConceptComponent, sErrorDescription) Then
				bHasChanged = False
				sSpecialCondition = "(ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (StartDate<>" & CLng(oRequest("StartDate").Item) & ") And"
			End If
			sQuery = "Select * From Concepts Where " & sSpecialCondition & " (ConceptShortName='" & aConceptComponent(S_SHORT_NAME_CONCEPT) & "') And (StartDate<" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate>" & aConceptComponent(N_END_DATE_CONCEPT) & ") Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Concepts Set EndDate=" & AddDaysToSerialDate(aConceptComponent(N_START_DATE_CONCEPT), -1) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (ConceptID=" & oRecordset.Fields("ConceptID").Value & ") And (StartDate=" & oRecordset.Fields("StartDate").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Concepts (ConceptID, StartDate, EndDate, ConceptShortName, ConceptName, BudgetID, PayrollTypeID, PeriodID, IsDeduction, ForAlimony, OnLeave, OrderInList, TaxAmount, TaxCurrencyID, TaxQttyID, TaxMin, TaxMinQttyID, TaxMax, TaxMaxQttyID, ExemptAmount, ExemptCurrencyID, ExemptQttyID, ExemptMin, ExemptMinQttyID, ExemptMax, ExemptMaxQttyID, StartUserID, EndUserID, StatusID) Values (" & oRecordset.Fields("ConceptID").Value & ", " & AddDaysToSerialDate(aConceptComponent(N_END_DATE_CONCEPT), 1) & ", " & oRecordset.Fields("EndDate").Value & ", '" & Replace(oRecordset.Fields("ConceptShortName").Value, "'", "") & "', '" & Replace(oRecordset.Fields("ConceptName").Value, "'", "´") & "', " & oRecordset.Fields("BudgetID").Value & ", " & oRecordset.Fields("PayrollTypeID").Value & ", " & oRecordset.Fields("PeriodID").Value & ", " & oRecordset.Fields("IsDeduction").Value & ", " & oRecordset.Fields("ForAlimony").Value & ", " & oRecordset.Fields("OnLeave").Value & ", " & oRecordset.Fields("OrderInList").Value & ", " & oRecordset.Fields("TaxAmount").Value & ", " & oRecordset.Fields("TaxCurrencyID").Value & ", " & oRecordset.Fields("TaxQttyID").Value & ", " & oRecordset.Fields("TaxMin").Value & ", " & oRecordset.Fields("TaxMinQttyID").Value & ", " & oRecordset.Fields("TaxMax").Value & ", " & oRecordset.Fields("TaxMaxQttyID").Value & ", " & oRecordset.Fields("ExemptAmount").Value & ", " & oRecordset.Fields("ExemptCurrencyID").Value & ", " & oRecordset.Fields("ExemptQttyID").Value & ", " & oRecordset.Fields("ExemptMin").Value & ", " & oRecordset.Fields("ExemptMinQttyID").Value & ", " & oRecordset.Fields("ExemptMax").Value & ", " & oRecordset.Fields("ExemptMaxQttyID").Value & ", " & oRecordset.Fields("StartUserID").Value & ", " & oRecordset.Fields("EndUserID").Value & ", " & oRecordset.Fields("StatusID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
			sQuery = "Select * From Concepts Where " & sSpecialCondition & " (ConceptShortName='" & aConceptComponent(S_SHORT_NAME_CONCEPT) & "') And (StartDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ") Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Concepts Where " & sSpecialCondition & " (ConceptShortName='" & aConceptComponent(S_SHORT_NAME_CONCEPT) & "') And (StartDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				End If
			End If
			sQuery = "Select * From Concepts Where " & sSpecialCondition & " (ConceptShortName='" & aConceptComponent(S_SHORT_NAME_CONCEPT) & "') And (StartDate<" & aConceptComponent(N_START_DATE_CONCEPT) & ") And ((EndDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ") And (EndDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Concepts Set EndDate=" & AddDaysToSerialDate(aConceptComponent(N_START_DATE_CONCEPT), -1) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (ConceptID=" & oRecordset.Fields("ConceptID").Value & ") And (StartDate=" & oRecordset.Fields("StartDate").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
			sQuery = "Select * From Concepts Where " & sSpecialCondition & " (ConceptShortName='" & aConceptComponent(S_SHORT_NAME_CONCEPT) & "') And (StartDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ") And ((EndDate>=" & aConceptComponent(N_END_DATE_CONCEPT) & ") And (StartDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Concepts Set StartDate=" & AddDaysToSerialDate(aConceptComponent(N_END_DATE_CONCEPT), 1) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (ConceptID=" & oRecordset.Fields("ConceptID").Value & ") And (StartDate=" & oRecordset.Fields("StartDate").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
			sErrorDescription = "No se pudo guardar la información del nuevo registro."
			If Not bHasChanged Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Concepts Set StartDate=" & aConceptComponent(N_START_DATE_CONCEPT) & ", EndDate=" & aConceptComponent(N_END_DATE_CONCEPT) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (StartDate=" & CLng(oRequest("StartDate").Item) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Concepts (ConceptID, StartDate, EndDate, ConceptShortName, ConceptName, BudgetID, PayrollTypeID, PeriodID, IsDeduction, ForAlimony, OnLeave, OrderInList, TaxAmount, TaxCurrencyID, TaxQttyID, TaxMin, TaxMinQttyID, TaxMax, TaxMaxQttyID, ExemptAmount, ExemptCurrencyID, ExemptQttyID, ExemptMin, ExemptMinQttyID, ExemptMax, ExemptMaxQttyID, StartUserID, EndUserID, StatusID) Values (" & aConceptComponent(N_ID_CONCEPT) & ", " & aConceptComponent(N_START_DATE_CONCEPT) & ", " & aConceptComponent(N_END_DATE_CONCEPT) & ", '" & Replace(aConceptComponent(S_SHORT_NAME_CONCEPT), "'", "") & "', '" & Replace(aConceptComponent(S_NAME_CONCEPT), "'", "´") & "', " & aConceptComponent(N_BUDGET_ID_CONCEPT) & ", " & aConceptComponent(N_PAYROLL_TYPE_ID_CONCEPT) & ", " & aConceptComponent(N_PERIOD_ID_CONCEPT) & ", " & aConceptComponent(N_IS_DEDUCTION_CONCEPT) & ", " & aConceptComponent(N_FOR_ALIMONY_CONCEPT) & ", " & aConceptComponent(N_ON_LEAVE_CONCEPT) & ", " & aConceptComponent(N_ORDER_IN_LIST_CONCEPT) & ", " & aConceptComponent(D_TAX_AMOUNT_CONCEPT) & ", " & aConceptComponent(N_TAX_CURRENCY_ID_CONCEPT) & ", " & aConceptComponent(N_TAX_QTTY_ID_CONCEPT) & ", " & aConceptComponent(D_TAX_MIN_CONCEPT) & ", " & aConceptComponent(N_TAX_MIN_QTTY_ID_CONCEPT) & ", " & aConceptComponent(D_TAX_MAX_CONCEPT) & ", " & aConceptComponent(N_TAX_MAX_QTTY_ID_CONCEPT) & ", " & aConceptComponent(D_EXEMPT_AMOUNT_CONCEPT) & ", " & aConceptComponent(N_EXEMPT_CURRENCY_ID_CONCEPT) & ", " & aConceptComponent(N_EXEMPT_QTTY_ID_CONCEPT) & ", " & aConceptComponent(D_EXEMPT_MIN_CONCEPT) & ", " & aConceptComponent(N_EXEMPT_MIN_QTTY_ID_CONCEPT) & ", " & aConceptComponent(D_EXEMPT_MAX_CONCEPT) & ", " & aConceptComponent(N_EXEMPT_MAX_QTTY_ID_CONCEPT) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", 1)", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If aConceptComponent(N_IS_CREDIT) > 0 Then
				sQuery = "Select * From CreditTypes Where (CreditTypeID=" & aConceptComponent(N_ID_CONCEPT) & ")"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then										 
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update CreditTypes Set CreditTypeID=" & aConceptComponent(N_ID_CONCEPT) & ", CreditTypeShortName='" & Replace(aConceptComponent(S_SHORT_NAME_CONCEPT), "'", "") & "', CreditTypeName='" & Replace(aConceptComponent(S_NAME_CONCEPT), "'", "´") & "', IsOther=" & aConceptComponent(N_IS_OTHER) & " Where (CreditTypeID=" & aConceptComponent(N_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Else							
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into CreditTypes (CreditTypeID, CreditTypeShortName, CreditTypeName, IsOther, Active) Values (" & aConceptComponent(N_ID_CONCEPT) & ", '" & Replace(aConceptComponent(S_SHORT_NAME_CONCEPT), "'", "") & "', '" & Replace(aConceptComponent(S_NAME_CONCEPT), "'", "´") & "', " & aConceptComponent(N_IS_OTHER) & ", 0)", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From CreditTypes Where (CreditTypeID=" & aConceptComponent(N_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If	
		End If
	End If

	Set oRecordset = Nothing
	ModifyConcept = lErrorNumber
	Err.Clear
End Function

Function ModifyConceptHistory(oRequest, oADODBConnection, ConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To insert and update a new record in the Concepts table
'Inputs:  oRequest, oADODBConnection
'Outputs: ConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyJobHistoryList"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sDate
	Dim lStartDate
	Dim lHistoryStartDate
	Dim lHistoryEndDate
	Dim sQuery

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If aConceptComponent(N_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador de la plaza a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If lErrorNumber = 0 Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select StartDate, EndDate From Concepts Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") Order By StartDate Desc", "aConceptComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sQuery = "select StartDate, enddate from Concepts Where (ConceptID = " & aConceptComponent(N_ID_CONCEPT) & ") And (StartDate > " & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate < " & aConceptComponent(N_END_DATE_CONCEPT) & ")"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)	
					If lErrorNumbeer = 0 Then
						If Not oRecordset.EOF Then
							sQuery = "Delete From Concepts Where (ConceptID = " & aConceptComponent(N_ID_CONCEPT) & ") And (StartDate > " & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate < " & aConceptComponent(N_END_DATE_CONCEPT) & ")"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
						If lErrorNumber = 0 Then
							sQuery = "Select StartDate, EndDate From Concepts Where (ConceptID = " & aConceptComponent(N_ID_CONCEPT) & ") And (StartDate < " & aConceptComponent(N_END_DATE_CONCEPT) & ") And (EndDate > " & aConceptComponent(N_END_DATE_CONCEPT) & ")"
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)	
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Concepts Set EndDate=" & AddDaysToSerialDate(aConceptComponent(N_START_DATE_CONCEPT), -1) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (EndDate=30000000)", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, null)
								End If
							End If
							If lErrorNumber = 0 Then
								sQuery = "Select StartDate, EndDate From Concepts Where (ConceptID = " & aConceptComponent(N_ID_CONCEPT) & ") And (EndDate > " & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate < " & aConceptComponent(N_END_DATE_CONCEPT) & ")"
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If Not oRecordset.EOF Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Concepts Set EndDate=" & AddDaysToSerialDate(aConceptComponent(N_START_DATE_CONCEPT), -1) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (EndDate=30000000)", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, null)
								End If
							End If
							If lErrorNumber = 0 Then
								sQuery = "Select StartDate, EndDate From Concepts Where (ConceptID = " & aConceptComponent(N_ID_CONCEPT) & ") And (StartDate = " & aConceptComponent(N_START_DATE_CONCEPT) & ")"
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If oRecordset.EOF Then
										If aConceptComponent(N_END_DATE_CONCEPT) = 0 Then aConceptComponent(N_END_DATE_CONCEPT) = 30000000
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Concepts (ConceptID, StartDate, EndDate, EmployeeID, OwnerID, CompanyID, ZoneID, AreaID, PaymentCenterID, PositionID, JobTypeID, ShiftID, WorkingHours, JourneyID, ClassificationID, GroupGradeLevelID, IntegrationID, OccupationTypeID, ServiceID, LevelID, StatusID, UserID, ModifyDate) Values (" & aConceptComponent(N_ID_CONCEPT) & ", " & aConceptComponent(N_START_DATE_CONCEPT) & ", " & aConceptComponent(N_END_DATE_CONCEPT) & ", " & aConceptComponent(N_ID_EMPLOYEE_JOB) & ", " & aConceptComponent(N_ID_OWNER_JOB) & ", " & aConceptComponent(N_COMPANY_ID_JOB) & ", " & aConceptComponent(N_ZONE_ID_JOB) & ", " & aConceptComponent(N_AREA_ID_JOB) & ", " & aConceptComponent(N_PAYMENT_CENTER_ID_JOB) & ", " & aConceptComponent(N_POSITION_ID_CONCEPT) & ", " & aConceptComponent(N_JOB_TYPE_ID_JOB) & ", " & aConceptComponent(N_SHIFT_ID_JOB) & ", " & aConceptComponent(D_WORKING_HOURS_JOB) & ", " & aConceptComponent(N_JOURNEY_ID_JOB) & ", " & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ", " & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_JOB) & ", " & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ", " & aConceptComponent(N_OCCUPATION_TYPE_ID_JOB) & ", " & aConceptComponent(N_SERVICE_ID_JOB) & ", " & aConceptComponent(N_LEVEL_ID_JOB) & ", " & aConceptComponent(N_STATUS_ID_JOB) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ")", "aConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, NULL)
									Else
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Concepts Set EndDate=" & AddDaysToSerialDate(aConceptComponent(N_START_DATE_CONCEPT), -1) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (EndDate=30000000)", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, NULL)
									End If
								End If
							End If
							If lErrorNumber <> 0 Then
								sErrorDescription = "El historial de la plaza no pudo actualizarse correctamente"
							End If
						End If
					End If
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Concepts (ConceptID, StartDate, EndDate, ConceptShortName, ConceptName, BudgetID, PayrollTypeID, PeriodID, IsDeduction, ForAlimony, OnLeave, OrderInList, TaxAmount, TaxCurrencyID, TaxQttyID, TaxMin, TaxMinQttyID, TaxMax, TaxMaxQttyID, ExemptAmount, ExemptCurrencyID, ExemptQttyID, ExemptMin, ExemptMinQttyID, ExemptMax, ExemptMaxQttyID, StartUserID, EndUserID) Values (" & aConceptComponent(N_ID_CONCEPT) & ", " & aConceptComponent(N_START_DATE_CONCEPT) & ", " & aConceptComponent(N_END_DATE_CONCEPT) & ", '" & Replace(aConceptComponent(S_SHORT_NAME_CONCEPT), "'", "") & "', '" & Replace(aConceptComponent(S_NAME_CONCEPT), "'", "´") & "', " & aConceptComponent(N_BUDGET_ID_CONCEPT) & ", " & aConceptComponent(N_PAYROLL_TYPE_ID_CONCEPT) & ", " & aConceptComponent(N_PERIOD_ID_CONCEPT) & ", " & aConceptComponent(N_IS_DEDUCTION_CONCEPT) & ", " & aConceptComponent(N_FOR_ALIMONY_CONCEPT) & ", " & aConceptComponent(N_ON_LEAVE_CONCEPT) & ", " & aConceptComponent(N_ORDER_IN_LIST_CONCEPT) & ", " & aConceptComponent(D_TAX_AMOUNT_CONCEPT) & ", " & aConceptComponent(N_TAX_CURRENCY_ID_CONCEPT) & ", " & aConceptComponent(N_TAX_QTTY_ID_CONCEPT) & ", " & aConceptComponent(D_TAX_MIN_CONCEPT) & ", " & aConceptComponent(N_TAX_MIN_QTTY_ID_CONCEPT) & ", " & aConceptComponent(D_TAX_MAX_CONCEPT) & ", " & aConceptComponent(N_TAX_MAX_QTTY_ID_CONCEPT) & ", " & aConceptComponent(D_EXEMPT_AMOUNT_CONCEPT) & ", " & aConceptComponent(N_EXEMPT_CURRENCY_ID_CONCEPT) & ", " & aConceptComponent(N_EXEMPT_QTTY_ID_CONCEPT) & ", " & aConceptComponent(D_EXEMPT_MIN_CONCEPT) & ", " & aConceptComponent(N_EXEMPT_MIN_QTTY_ID_CONCEPT) & ", " & aConceptComponent(D_EXEMPT_MAX_CONCEPT) & ", " & aConceptComponent(N_EXEMPT_MAX_QTTY_ID_CONCEPT) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", 0)", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
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

Function ModifyConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new concept into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyConceptValue"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim sSpecialCondition
	Dim bHasChanged
	Dim lNewRecordID

	If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) = 0 Then aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) = 30000000
		bHasChanged = True
		If Not ConceptValueHasChanged(oRequest, oADODBConnection, aConceptComponent, sErrorDescription) Then
			bHasChanged = False
			sSpecialCondition = "(RecordID<>" & aConceptComponent(N_RECORD_ID_CONCEPT) & ") And"
		End If
		sQuery = "Select * From ConceptsValues Where " & sSpecialCondition & " (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ") And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (PositionTypeID=" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ") And (EmployeeStatusID=" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ") And (JobStatusID=" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ") And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (JourneyID=" & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (AdditionalShift=" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EconomicZoneID=" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (AntiquityID=" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ") And (Antiquity2ID=" & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ") And (Antiquity3ID=" & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ") And (Antiquity4ID=" & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ") And (ForRisk=" & aConceptComponent(N_FOR_RISK_CONCEPT) & ") And (GenderID=" & aConceptComponent(N_GENDER_ID_CONCEPT) & ") And (HasChildren=" & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ") And (SchoolarshipID=" & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ") And (HasSyndicate=" & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ") And (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (StartDate<" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ") And (EndDate>" & aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & ") Order By StartDate Desc"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConceptsValues Set EndDate=" & AddDaysToSerialDate(aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT), -1) & " Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				lErrorNumber = GetNewIDFromTable(oADODBConnection, "ConceptsValues", "RecordID", "", 1, lNewRecordID, sErrorDescription)
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into ConceptsValues (RecordID, ConceptID, CompanyID, EmployeeTypeID, PositionTypeID, EmployeeStatusID, JobStatusID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, WorkingHours, AdditionalShift, LevelID, EconomicZoneID, ServiceID, AntiquityID, Antiquity2ID, Antiquity3ID, Antiquity4ID, ForRisk, GenderID, HasChildren, SchoolarshipID, HasSyndicate, StartDate, EndDate, RegistrationStartDate, AuthorizationDate, RegistrationEndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, AppliesToID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, PositionID, StartUserID, EndUserID, StatusID) Values (" & lNewRecordID & ", " & oRecordset.Fields("ConceptID").Value & ", " & oRecordset.Fields("CompanyID").Valuea & ", " & oRecordset.Fields("EmployeeTypeID").Value & ", " & oRecordset.Fields("PositionTypeID").Value & ", " & oRecordset.Fields("EmployeeStatusID").Value & ", " & oRecordset.Fields("JobStatusID").Value & ", " & oRecordset.Fields("ClassificationID").Value & ", " & oRecordset.Fields("GroupGradeLevelID").Value & ", " & oRecordset.Fields("IntegrationID").Value & ", " & oRecordset.Fields("JourneyID").Value & ", " & oRecordset.Fields("WorkingHours").Value & ", " & oRecordset.Fields("AdditionalShift").Value & ", " & oRecordset.Fields("LevelID").Value & ", " & oRecordset.Fields("EconomicZoneID").Value & ", " & oRecordset.Fields("ServiceID").Value & ", " & oRecordset.Fields("AntiquityID").Value & ", " & oRecordset.Fields("Antiquity2ID").Value & ", " & oRecordset.Fields("Antiquity3ID").Value & ", " & oRecordset.Fields("Antiquity4ID").Value & ", " & oRecordset.Fields("ForRisk").Value & ", " & oRecordset.Fields("GenderID").Value & ", " & oRecordset.Fields("HasChildren").Value & ", " & oRecordset.Fields("SchoolarshipID").Value & ", " & oRecordset.Fields("HasSyndicate").Value & ", " & AddDaysToSerialDate(aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT), 1) & ", " & oRecordset.Fields("EndDate").Value & ", " & oRecordset.Fields("RegistrationStartDate").Value & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", 0, " & oRecordset.Fields("ConceptAmount").Value & ", " & oRecordset.Fields("CurrencyID").Value & ", " & oRecordset.Fields("ConceptQttyID").Value & ", " & oRecordset.Fields("ConceptTypeID").Value & ", '" & oRecordset.Fields("AppliesToID").Value & "', " & oRecordset.Fields("ConceptMin").Value & ", " & oRecordset.Fields("ConceptMinQttyID").Value & ", " & oRecordset.Fields("ConceptMax").Value & ", " & oRecordset.Fields("ConceptMaxQttyID").Value & ", " & oRecordset.Fields("PositionID").Value & ", " & oRecordset.Fields("StartUserID").Value & ", " & aConceptComponent(N_END_USER_ID_CONCEPT) & ", " & oRecordset.Fields("StatusID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
		sQuery = "Select * From ConceptsValues Where " & sSpecialCondition & " (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ") And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (PositionTypeID=" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ") And (EmployeeStatusID=" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ") And (JobStatusID=" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ") And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (JourneyID=" & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (AdditionalShift=" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EconomicZoneID=" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (AntiquityID=" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ") And (Antiquity2ID=" & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ") And (Antiquity3ID=" & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ") And (Antiquity4ID=" & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ") And (ForRisk=" & aConceptComponent(N_FOR_RISK_CONCEPT) & ") And (GenderID=" & aConceptComponent(N_GENDER_ID_CONCEPT) & ") And (HasChildren=" & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ") And (SchoolarshipID=" & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ") And (HasSyndicate=" & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ") And (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (StartDate>=" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ") And (EndDate<=" & aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & ") Order By StartDate Desc"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				Do While Not oRecordset.EOF
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From ConceptsValues Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			End If
		End If
		sQuery = "Select * From ConceptsValues Where " & sSpecialCondition & " (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ") And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (PositionTypeID=" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ") And (EmployeeStatusID=" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ") And (JobStatusID=" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ") And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (JourneyID=" & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (AdditionalShift=" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EconomicZoneID=" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (AntiquityID=" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ") And (Antiquity2ID=" & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ") And (Antiquity3ID=" & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ") And (Antiquity4ID=" & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ") And (ForRisk=" & aConceptComponent(N_FOR_RISK_CONCEPT) & ") And (GenderID=" & aConceptComponent(N_GENDER_ID_CONCEPT) & ") And (HasChildren=" & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ") And (SchoolarshipID=" & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ") And (HasSyndicate=" & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ") And (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (StartDate<" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ") And ((EndDate<=" & aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & ") And (EndDate>=" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate Desc"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConceptsValues Set EndDate=" & AddDaysToSerialDate(aConceptComponent(N_START_DATE_CONCEPT), -1) & " Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
		sQuery = "Select * From ConceptsValues Where " & sSpecialCondition & " (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ") And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (PositionTypeID=" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ") And (EmployeeStatusID=" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ") And (JobStatusID=" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ") And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (JourneyID=" & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (AdditionalShift=" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EconomicZoneID=" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (AntiquityID=" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ") And (Antiquity2ID=" & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ") And (Antiquity3ID=" & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ") And (Antiquity4ID=" & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ") And (ForRisk=" & aConceptComponent(N_FOR_RISK_CONCEPT) & ") And (GenderID=" & aConceptComponent(N_GENDER_ID_CONCEPT) & ") And (HasChildren=" & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ") And (SchoolarshipID=" & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ") And (HasSyndicate=" & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ") And (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (StartDate>=" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ") And ((EndDate>" & aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & ") And (StartDate<=" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConceptsValues Set StartDate=" & AddDaysToSerialDate(aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT), 1) & " Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
		sErrorDescription = "No se pudo guardar la información del nuevo registro."
		If Not bHasChanged Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConceptsValues Set StartDate=" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ", EndDate=" & aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & " Where (RecordID=" & aConceptComponent(N_RECORD_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Else
			sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
			lErrorNumber = GetNewIDFromTable(oADODBConnection, "ConceptsValues", "RecordID", "", 1, lNewRecordID, sErrorDescription)
			If lErrorNumber = 0 Then
				sQuery = "Select * From ConceptsValues Where " & sSpecialCondition & " (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ") And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (PositionTypeID=" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ") And (EmployeeStatusID=" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ") And (JobStatusID=" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ") And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (JourneyID=" & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (AdditionalShift=" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EconomicZoneID=" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (AntiquityID=" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ") And (Antiquity2ID=" & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ") And (Antiquity3ID=" & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ") And (Antiquity4ID=" & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ") And (ForRisk=" & aConceptComponent(N_FOR_RISK_CONCEPT) & ") And (GenderID=" & aConceptComponent(N_GENDER_ID_CONCEPT) & ") And (HasChildren=" & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ") And (SchoolarshipID=" & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ") And (HasSyndicate=" & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ") And (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (StartDate=" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ") And (EndDate=" & aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & ")"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConceptsValues Set ConceptAmount=" & aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT) & " Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Else
						sErrorDescription = "No se pudo guardar la información del nuevo registro."
						aConceptComponent(N_STATUS_ID_CONCEPT) = 1
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into ConceptsValues (RecordID, ConceptID, CompanyID, EmployeeTypeID, PositionTypeID, EmployeeStatusID, JobStatusID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, WorkingHours, AdditionalShift, LevelID, EconomicZoneID, ServiceID, AntiquityID, Antiquity2ID, Antiquity3ID, Antiquity4ID, ForRisk, GenderID, HasChildren, SchoolarshipID, HasSyndicate, StartDate, EndDate, RegistrationStartDate, AuthorizationDate, RegistrationEndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, AppliesToID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, PositionID, StartUserID, EndUserID, StatusID) Values (" & lNewRecordID & ", " & aConceptComponent(N_ID_CONCEPT) & ", " & aConceptComponent(N_COMPANY_ID_CONCEPT) & ", " & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ", " & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ", " & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ", " & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ", " & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ", " & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ", " & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ", " & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ", " & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ", " & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ", " & aConceptComponent(N_LEVEL_ID_CONCEPT) & ", " & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ", " & aConceptComponent(N_SERVICE_ID_CONCEPT) & ", " & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ", " & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ", " & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ", " & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ", " & aConceptComponent(N_FOR_RISK_CONCEPT) & ", " & aConceptComponent(N_GENDER_ID_CONCEPT) & ", " & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ", " & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ", " & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ", " & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ", " & aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", 0, 0, " & aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT) & ", " & aConceptComponent(N_CURRENCY_ID_CONCEPT) & ", " & aConceptComponent(N_CONCEPT_QTTY_ID_CONCEPT) & ", " & aConceptComponent(N_CONCEPT_TYPE_ID_CONCEPT) & ", '" & Replace(aConceptComponent(S_APPLIES_ID_CONCEPT), "'", "") & "', " & aConceptComponent(D_CONCEPT_MIN_CONCEPT) & ", " & aConceptComponent(N_CONCEPT_MIN_QTTY_ID_CONCEPT) & ", " & aConceptComponent(D_CONCEPT_MAX_CONCEPT) & ", " & aConceptComponent(N_CONCEPT_MAX_QTTY_ID_CONCEPT) & ", " & aConceptComponent(N_POSITION_ID_CONCEPT) & ", " & aConceptComponent(N_START_USER_ID_CONCEPT) & ", " & aConceptComponent(N_END_USER_ID_CONCEPT) & ", " & aConceptComponent(N_STATUS_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
			End If
		End If
	End If

	ModifyConceptValue = lErrorNumber
	Err.Clear
End Function

Function ModifyPositionsSpecialJourneysLKP(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the concept value in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyPositionsSpecialJourneysLKP"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim sSpecialCondition
	Dim bHasChanged
	Dim lNewRecordID

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Not CheckConceptInformationConsistency(aConceptComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			If aConceptComponent(N_END_DATE_CONCEPT) = 0 Then aConceptComponent(N_END_DATE_CONCEPT) = 30000000
			bHasChanged = True
			If Not PositionsSpecialJourneysLKPHasChanged(oRequest, oADODBConnection, aConceptComponent, sErrorDescription) Then
				bHasChanged = False
				sSpecialCondition = "(RecordID<>" & aConceptComponent(N_RECORD_ID_CONCEPT) & ") And"
			End If
			sQuery = "Select * From PositionsSpecialJourneysLKP Where " & sSpecialCondition & " (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (CenterTypeID=" & aConceptComponent(N_CENTER_TYPE_ID) & ") And (StartDate<" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate>" & aConceptComponent(N_END_DATE_CONCEPT) & ") Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PositionsSpecialJourneysLKP Set EndDate=" & AddDaysToSerialDate(aConceptComponent(N_START_DATE_CONCEPT), -1) & " Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					lErrorNumber = GetNewIDFromTable(oADODBConnection, "PositionsSpecialJourneysLKP", "RecordID", "", 1, lNewRecordID, sErrorDescription)
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PositionsSpecialJourneysLKP (RecordID, StartDate, EndDate, PositionID, LevelID, WorkingHours, ServiceID, CenterTypeID, IsActive1, IsActive2, IsActive3, IsActive4, Active) Values (" & lNewRecordID & ", " & AddDaysToSerialDate(aConceptComponent(N_END_DATE_CONCEPT), 1) & ", " & oRecordset.Fields("EndDate").Value & ", " & oRecordset.Fields("PositionID").Value & ", '" & oRecordset.Fields("LevelID").Value & ", " & oRecordset.Fields("WorkingHours").Value & ", " & oRecordset.Fields("ServiceID").Value & ", " & oRecordset.Fields("CenterTypeID").Value & ", " & oRecordset.Fields("IsActive1").Value & ", " & oRecordset.Fields("IsActive2").Value & ", " & oRecordset.Fields("IsActive3").Value & ", " & oRecordset.Fields("IsActive4").Value & ", " & oRecordset.Fields("Active").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
			sQuery = "Select * From PositionsSpecialJourneysLKP Where " & sSpecialCondition & " (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (CenterTypeID=" & aConceptComponent(N_CENTER_TYPE_ID) & ") And (StartDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ") Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					Do While Not oRecordset.EOF
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PositionsSpecialJourneysLKP Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
				End If
			End If
			sQuery = "Select * From PositionsSpecialJourneysLKP Where " & sSpecialCondition & " (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (CenterTypeID=" & aConceptComponent(N_CENTER_TYPE_ID) & ") And (StartDate<" & aConceptComponent(N_START_DATE_CONCEPT) & ") And ((EndDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ") And (EndDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PositionsSpecialJourneysLKP Set EndDate=" & AddDaysToSerialDate(aConceptComponent(N_START_DATE_CONCEPT), -1) & " Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
			sQuery = "Select * From PositionsSpecialJourneysLKP Where " & sSpecialCondition & " (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (CenterTypeID=" & aConceptComponent(N_CENTER_TYPE_ID) & ") And (StartDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ") And ((EndDate>" & aConceptComponent(N_END_DATE_CONCEPT) & ") And (StartDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PositionsSpecialJourneysLKP Set StartDate=" & AddDaysToSerialDate(aConceptComponent(N_END_DATE_CONCEPT), 1) & " Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
			sErrorDescription = "No se pudo guardar la información del nuevo registro."
			If Not bHasChanged Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PositionsSpecialJourneysLKP Set StartDate=" & aConceptComponent(N_START_DATE_CONCEPT) & ", EndDate=" & aConceptComponent(N_END_DATE_CONCEPT) & " Where (RecordID=" & aConceptComponent(N_RECORD_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Else
				sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
				lErrorNumber = GetNewIDFromTable(oADODBConnection, "PositionsSpecialJourneysLKP", "RecordID", "", 1, lNewRecordID, sErrorDescription)
				If lErrorNumber = 0 Then
					sQuery = "Select * From PositionsSpecialJourneysLKP Where " & sSpecialCondition & " (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (CenterTypeID=" & aConceptComponent(N_CENTER_TYPE_ID) & ") And (StartDate=" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate=" & aConceptComponent(N_END_DATE_CONCEPT) & ")"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PositionsSpecialJourneysLKP Set IsActive1=" & aConceptComponent(N_IS_ACTIVE1) & ", IsActive2=" & aConceptComponent(N_IS_ACTIVE1) & ", IsActive3=" & aConceptComponent(N_IS_ACTIVE3) & ", IsActive4=" & aConceptComponent(N_IS_ACTIVE4) & " Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						Else
							sErrorDescription = "No se pudo guardar la información del nuevo registro."
							aConceptComponent(N_ACTIVE) = 1
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PositionsSpecialJourneysLKP (RecordID, StartDate, EndDate, PositionID, LevelID, WorkingHours, ServiceID, CenterTypeID, IsActive1, IsActive2, IsActive3, IsActive4, Active) Values (" & lNewRecordID & ", " & aConceptComponent(N_START_DATE_CONCEPT) & ", " & aConceptComponent(N_END_DATE_CONCEPT) & ", " & aConceptComponent(N_POSITION_ID_CONCEPT) & ", " & aConceptComponent(N_LEVEL_ID_CONCEPT) & ", " & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ", " & aConceptComponent(N_SERVICE_ID_CONCEPT) & ", " & aConceptComponent(N_CENTER_TYPE_ID) & ", " & aConceptComponent(N_IS_ACTIVE1) & ", " & aConceptComponent(N_IS_ACTIVE2) & ", " & aConceptComponent(N_IS_ACTIVE3) & ", " & aConceptComponent(N_IS_ACTIVE4) & ", " & aConceptComponent(N_ACTIVE) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
					End If
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	ModifyPositionsSpecialJourneysLKP = lErrorNumber
	Err.Clear
End Function

Function RemoveConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a concept from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveConcept"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If aConceptComponent(N_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el concepto a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del concepto."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Concepts Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (StartDate=" & aConceptComponent(N_START_DATE_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		If False Then
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo eliminar la información del concepto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From ConceptsValues Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo eliminar la información del concepto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesConceptsLKP Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo eliminar la información del concepto."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PositionsConceptsLKP Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
	End If

	RemoveConcept = lErrorNumber
	Err.Clear
End Function

Function RemoveConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a concept from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveConceptValue"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el registro del concepto a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del concepto."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From ConceptsValues Where (RecordID=" & aConceptComponent(N_RECORD_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		If Len(oRequest("RecordID2").Item) > 0 Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConceptsValues Set EndDate = 30000000, RegistrationEndDate = 0, EndUserID = -1 Where (RecordID=" & CLng(oRequest("RecordID2").Item) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	RemoveConceptValue = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfConcept(aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific bank account exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfConcept"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim sConceptCrossingType

	sErrorDescription = "No se pudo revisar la existencia de la cuenta bancaria en la base de datos."
	sQuery = "Select * From Concepts Where (ConceptShortName='" & Replace(aConceptComponent(S_SHORT_NAME_CONCEPT), "'", "") & "')" & _
			 " And (((StartDate >= " &  aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate <= " &  aConceptComponent(N_END_DATE_CONCEPT) & "))" & _
			 " Or ((EndDate >= " &  aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate <= " &  aConceptComponent(N_END_DATE_CONCEPT) & "))" & _
			 " Or ((EndDate >= " &  aConceptComponent(N_START_DATE_CONCEPT) & ") And (StartDate <= " &  aConceptComponent(N_END_DATE_CONCEPT) & "))) Order By StartDate Desc"

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If aConceptComponent(N_START_DATE_CONCEPT) <> CLng(oRecordset.Fields("StartDate").Value) Then
				aConceptComponent(N_ID_CONCEPT) = CLng(oRecordset.Fields("ConceptID").Value)
				lErrorNumber = GetConceptCrossType(oADODBConnection, aConceptComponent, sConceptCrossingType, sErrorDescription)
				If lErrorNumber = 0 Then
					Select Case sConceptCrossingType
						Case "Left"
							aConceptComponent(N_STATUS_ID_CONCEPT) = 0
						Case "Right"
							aConceptComponent(N_STATUS_ID_CONCEPT) = -1
						Case "Inner"
							aConceptComponent(N_STATUS_ID_CONCEPT) = -2
						Case "Cross"
							aConceptComponent(N_STATUS_ID_CONCEPT) = -3
					End Select
					CheckExistencyOfConcept = True
				Else
					sErrorDescription = "No se puede agregar el Concepto " & Replace(aConceptComponent(S_SHORT_NAME_CONCEPT), "'", "") & " con fecha de inicio " & DisplayDateFromSerialNumber(aConceptComponent(N_START_DATE_CONCEPT), -1, -1, -1) & " debido a que existe una registrada en el periodo indicado"
					CheckExistencyOfConcept = False
				End If
			Else
				sErrorDescription = "No se puede agregar el Concepto " & Replace(aConceptComponent(S_SHORT_NAME_CONCEPT), "'", "") & " con fecha de inicio " & DisplayDateFromSerialNumber(aConceptComponent(N_START_DATE_CONCEPT), -1, -1, -1) & " debido a que existe uno registrado con la misma fecha de inicio"
				CheckExistencyOfConcept = False
			End If
		Else
			aConceptComponent(N_STATUS_ID_CONCEPT) = 0
			CheckExistencyOfConcept = True
		End If
	Else
		sErrorDescription = "No se puede agregar el Concepto " & Replace(aConceptComponent(S_SHORT_NAME_CONCEPT), "'", "") & " con fecha de inicio " & DisplayDateFromSerialNumber(aConceptComponent(N_START_DATE_CONCEPT), -1, -1, -1) & " debido a que existe una registrada en el periodo indicado"
		CheckExistencyOfConcept = False
	End If
	oRecordset.Close

	Set oRecordset = Nothing
	Err.Clear
End Function

Function CheckExistencyOfConceptValue(aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific bank account exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfConceptValue"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim sEmployeeConceptType
	Dim sConceptValuesCrossType

	sErrorDescription = "No se pudo revisar la existencia de la cuenta bancaria en la base de datos."
	sQuery = "Select * From ConceptsValues Where (RecordID<>" & aConceptComponent(N_RECORD_ID_CONCEPT) & ") And (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ")" & _
			 " And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (PositionTypeID=" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ") And (EmployeeStatusID=" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ")" & _
			 " And (JobStatusID=" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ")" & _
			 " And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (JourneyID=" & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ")" & _
			 " And (AdditionalShift=" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EconomicZoneID=" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ")" & _
			 " And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (AntiquityID=" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ") And (Antiquity2ID=" & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ")" & _
			 " And (Antiquity3ID=" & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ") And (Antiquity4ID=" & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ") And (ForRisk=" & aConceptComponent(N_FOR_RISK_CONCEPT) & ")" & _
			 " And (GenderID=" & aConceptComponent(N_GENDER_ID_CONCEPT) & ") And (HasChildren=" & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ") And (SchoolarshipID=" & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ")" & _
			 " And (HasSyndicate=" & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ") And (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ")" & _
			 " And (((StartDate >= " &  aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ") And (EndDate <= " &  aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & "))" & _
			 " Or ((EndDate >= " &  aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ") And (EndDate <= " &  aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & "))" & _
			 " Or ((EndDate >= " &  aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ") And (StartDate <= " &  aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & "))) Order By StartDate Desc"

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) <> CLng(oRecordset.Fields("StartDate").Value) Then
				lErrorNumber = GetConceptValueCrossType(oADODBConnection, aConceptComponent, sConceptValuesCrossType, sErrorDescription)
				If lErrorNumber = 0 Then
					Select Case sConceptValuesCrossType
						Case "Left"
							aConceptComponent(N_STATUS_ID_CONCEPT) = 0
						Case "Right"
							aConceptComponent(N_STATUS_ID_CONCEPT) = -1
						Case "Inner"
							aConceptComponent(N_STATUS_ID_CONCEPT) = -2
						Case "Cross"
							aConceptComponent(N_STATUS_ID_CONCEPT) = -3
					End Select
					CheckExistencyOfConceptValue = True
				Else
					sErrorDescription = "No se puede agregar el tabulador del puesto " & aConceptComponent(S_POSITION_SHORT_NAME_CONCEPT) & " con fecha de inicio " & DisplayDateFromSerialNumber(aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT), -1, -1, -1) & " debido a que existe una registrado en el periodo indicado"
					CheckExistencyOfConceptValue = False
				End If
			Else
				sErrorDescription = "No se puede agregar el tabulador del puesto " & aConceptComponent(S_POSITION_SHORT_NAME_CONCEPT) & " con fecha de inicio " & DisplayDateFromSerialNumber(aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT), -1, -1, -1) & " debido a que existe uno registrado con la misma fecha de inicio"
				CheckExistencyOfConceptValue = False
			End If
		Else
			aConceptComponent(N_STATUS_ID_CONCEPT) = 0
			CheckExistencyOfConceptValue = True
		End If
	Else
		sErrorDescription = "No se puede agregar el tabulador del puesto " & aConceptComponent(S_POSITION_SHORT_NAME_CONCEPT) & " con fecha de inicio " & DisplayDateFromSerialNumber(aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT), -1, -1, -1) & " debido a que existe uno registrado en el periodo indicado"
		CheckExistencyOfConceptValue = False
	End If
	oRecordset.Close

	Set oRecordset = Nothing
	Err.Clear
End Function

Function CheckExistencyOfPosition(aConceptComponent, sPositionShortName, sErrorDescription)
'************************************************************
'Purpose: To check if a specific position exists in the database
'Inputs:  aConceptComponent, sPositionShortName
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfPosition"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If Len(sPositionShortName) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó la clave del puesto para para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Positions Where (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ") And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (PositionShortName ='" & sPositionShortName & "')  And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (Workinghours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (((StartDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ")) Or ((EndDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ")) Or ((EndDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (StartDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ")))", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El puesto " & sPositionShortName & ", con los parámetros indicados no se encuentra en a base de datos."
			Else
				If CLng(oRecordset.Fields("EndDate").Value) < CLng(aConceptComponent(N_START_DATE_FOR_REGISTRATION_CONCEPT)) Then
					lErrorNumber = -1
					sErrorDescription = "El puesto " & sPositionShortName & " no es vigente en la fecha de inicio indicada."
				Else
					aConceptComponent(N_POSITION_ID_CONCEPT) = CLng(oRecordset.Fields("PositionID").Value)
				End If
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfPosition = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfPositionForSpecialJourneys(aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific position exists in the database
'Inputs:  aConceptComponent, sPositionShortName
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfPositionForSpecialJourneys"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If Len(aConceptComponent(S_SHORT_NAME_CONCEPT)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó la clave del puesto para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Positions Where (PositionShortName ='" & aConceptComponent(S_SHORT_NAME_CONCEPT) & "') And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (EndDate=30000000) And (Active=1)", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = -1
				sErrorDescription = "No existe un puesto con las condiciones especificadas."
			Else
				If CLng(oRecordset.Fields("EndDate").Value) < CLng(aConceptComponent(N_START_DATE_FOR_REGISTRATION_CONCEPT)) Then
					lErrorNumber = -1
					sErrorDescription = "El puesto especificado no es vigente en la fecha de inicio indicada."
				Else
					aConceptComponent(N_POSITION_ID_CONCEPT) = CLng(oRecordset.Fields("PositionID").Value)
				End If
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfPositionForSpecialJourneys = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfPositionsSpecialJourneysLKP(aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific bank account exists in the database
'Inputs:  aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfPositionsSpecialJourneysLKP"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim sEmployeeConceptType
	Dim sPositionSpecialJourneyCrossingType

	If aConceptComponent(N_POSITION_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el puesto del registro para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia de la cuenta bancaria en la base de datos."
		sQuery = "Select * From PositionsSpecialJourneysLKP Where (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ")" & _
				 " And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (CenterTypeID=" & aConceptComponent(N_CENTER_TYPE_ID) & ")" & _
				 " And (((StartDate >= " &  aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate <= " &  aConceptComponent(N_END_DATE_CONCEPT) & "))" & _
				 " Or ((EndDate >= " &  aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate <= " &  aConceptComponent(N_END_DATE_CONCEPT) & "))" & _
				 " Or ((EndDate >= " &  aConceptComponent(N_START_DATE_CONCEPT) & ") And (StartDate <= " &  aConceptComponent(N_END_DATE_CONCEPT) & "))) Order By StartDate Desc"

		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aEmployeeComponent(N_CONCEPT_CREDIT_TYPE) = 3
				lErrorNumber = GetPositionSpecialJourneyCrossingType(oADODBConnection, aConceptComponent, sPositionSpecialJourneyCrossingType, sErrorDescription)
				If lErrorNumber = 0 Then
					Select Case sPositionSpecialJourneyCrossingType
						Case "Left"
							aConceptComponent(N_ACTIVE) = 0
						Case "Right"
							aConceptComponent(N_ACTIVE) = -1
						Case "Inner"
							aConceptComponent(N_ACTIVE) = -2
						Case "Cross"
							aConceptComponent(N_ACTIVE) = -3
					End Select
					CheckExistencyOfPositionsSpecialJourneysLKP = True
				Else
					sErrorDescription = "No se puede agregar el puesto " & aConceptComponent(N_POSITION_ID_CONCEPT) & " para guardías y suplencias con fecha de inicio " & DisplayDateFromSerialNumber(aConceptComponent(N_START_DATE_CONCEPT), -1, -1, -1) & " debido a que existe una registrada en el periodo indicado"
					CheckExistencyOfPositionsSpecialJourneysLKP = False
				End If
			Else
				aConceptComponent(N_ACTIVE) = 0
				CheckExistencyOfPositionsSpecialJourneysLKP = True
			End If
		Else
			sErrorDescription = "No se puede agregar el puesto " & aConceptComponent(N_POSITION_ID_CONCEPT) & " para guardías y suplencias con fecha de inicio " & DisplayDateFromSerialNumber(aConceptComponent(N_START_DATE_CONCEPT), -1, -1, -1) & " debido a que existe una registrada en el periodo indicado"
			CheckExistencyOfPositionsSpecialJourneysLKP = False
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function CheckConceptInformationConsistency(aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aConceptComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckConceptInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aConceptComponent(N_ID_CONCEPT)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del registro no es un valor numérico."
		bIsCorrect = False
	End If
	If StrComp(oRequest("Action").Item, "Concepts", vbBinaryCompare) = 0 Then
		If Not IsNumeric(aConceptComponent(N_START_DATE_CONCEPT)) Then aConceptComponent(N_START_DATE_CONCEPT) = Left(GetSerialNumberForDate(""), Len("00000000"))
		If Not IsNumeric(aConceptComponent(N_END_DATE_CONCEPT)) Then aConceptComponent(N_END_DATE_CONCEPT) = 0
		If Len(aConceptComponent(S_SHORT_NAME_CONCEPT)) = 0 Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- La clave del registro está vacía."
			bIsCorrect = False
		End If
		If Len(aConceptComponent(S_NAME_CONCEPT)) = 0 Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del registro está vacío."
			bIsCorrect = False
		End If
		If Not IsNumeric(aConceptComponent(N_BUDGET_ID_CONCEPT)) Then aConceptComponent(N_BUDGET_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_PAYROLL_TYPE_ID_CONCEPT)) Then aConceptComponent(N_PAYROLL_TYPE_ID_CONCEPT) = 1
		If Not IsNumeric(aConceptComponent(N_PERIOD_ID_CONCEPT)) Then aConceptComponent(N_PERIOD_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_IS_DEDUCTION_CONCEPT)) Then aConceptComponent(N_IS_DEDUCTION_CONCEPT) = 0
		If Not IsNumeric(aConceptComponent(N_FOR_ALIMONY_CONCEPT)) Then aConceptComponent(N_FOR_ALIMONY_CONCEPT) = 0
		If Not IsNumeric(aConceptComponent(N_ON_LEAVE_CONCEPT)) Then aConceptComponent(N_ON_LEAVE_CONCEPT) = 0
		If Not IsNumeric(aConceptComponent(N_ORDER_IN_LIST_CONCEPT)) Then aConceptComponent(N_ORDER_IN_LIST_CONCEPT) = 400
		If Not IsNumeric(aConceptComponent(D_TAX_AMOUNT_CONCEPT)) Then aConceptComponent(D_TAX_AMOUNT_CONCEPT) = 0
		If Not IsNumeric(aConceptComponent(N_TAX_CURRENCY_ID_CONCEPT)) Then aConceptComponent(N_TAX_CURRENCY_ID_CONCEPT) = 0
		If Not IsNumeric(aConceptComponent(N_TAX_QTTY_ID_CONCEPT)) Then aConceptComponent(N_TAX_QTTY_ID_CONCEPT) = 1
		If Not IsNumeric(aConceptComponent(D_TAX_MIN_CONCEPT)) Then aConceptComponent(D_TAX_MIN_CONCEPT) = 0
		If Not IsNumeric(aConceptComponent(N_TAX_MIN_QTTY_ID_CONCEPT)) Then aConceptComponent(N_TAX_MIN_QTTY_ID_CONCEPT) = 1
		If Not IsNumeric(aConceptComponent(D_TAX_MAX_CONCEPT)) Then aConceptComponent(D_TAX_MAX_CONCEPT) = 0
		If Not IsNumeric(aConceptComponent(N_TAX_MAX_QTTY_ID_CONCEPT)) Then aConceptComponent(N_TAX_MAX_QTTY_ID_CONCEPT) = 1
		If Not IsNumeric(aConceptComponent(D_EXEMPT_AMOUNT_CONCEPT)) Then aConceptComponent(D_EXEMPT_AMOUNT_CONCEPT) = 0
		If Not IsNumeric(aConceptComponent(N_EXEMPT_CURRENCY_ID_CONCEPT)) Then aConceptComponent(N_EXEMPT_CURRENCY_ID_CONCEPT) = 0
		If Not IsNumeric(aConceptComponent(N_EXEMPT_QTTY_ID_CONCEPT)) Then aConceptComponent(N_EXEMPT_QTTY_ID_CONCEPT) = 1
		If Not IsNumeric(aConceptComponent(D_EXEMPT_MIN_CONCEPT)) Then aConceptComponent(D_EXEMPT_MIN_CONCEPT) = 0
		If Not IsNumeric(aConceptComponent(N_EXEMPT_MIN_QTTY_ID_CONCEPT)) Then aConceptComponent(N_EXEMPT_MIN_QTTY_ID_CONCEPT) = 1
		If Not IsNumeric(aConceptComponent(D_EXEMPT_MAX_CONCEPT)) Then aConceptComponent(D_EXEMPT_MAX_CONCEPT) = 0
		If Not IsNumeric(aConceptComponent(N_EXEMPT_MAX_QTTY_ID_CONCEPT)) Then aConceptComponent(N_EXEMPT_MAX_QTTY_ID_CONCEPT) = 1
	ElseIf StrComp(oRequest("Action").Item, "ConceptValues", vbBinaryCompare) = 0 Then
		If Not IsNumeric(aConceptComponent(N_RECORD_ID_CONCEPT)) Then aConceptComponent(N_RECORD_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_COMPANY_ID_CONCEPT)) Then aConceptComponent(N_COMPANY_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT)) Then aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_POSITION_TYPE_ID_CONCEPT)) Then aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT)) Then aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_JOB_STATUS_ID_CONCEPT)) Then aConceptComponent(N_JOB_STATUS_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_CLASSIFICATION_ID_CONCEPT)) Then aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT)) Then aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_INTEGRATION_ID_CONCEPT)) Then aConceptComponent(N_INTEGRATION_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_JOURNEY_ID_CONCEPT)) Then aConceptComponent(N_JOURNEY_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(D_WORKING_HOURS_CONCEPT)) Then aConceptComponent(D_WORKING_HOURS_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT)) Then aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) = 0
		If Not IsNumeric(aConceptComponent(N_LEVEL_ID_CONCEPT)) Then aConceptComponent(N_LEVEL_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT)) Then aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) = 0
		If Not IsNumeric(aConceptComponent(N_SERVICE_ID_CONCEPT)) Then aConceptComponent(N_SERVICE_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_ANTIQUITY_ID_CONCEPT)) Then aConceptComponent(N_ANTIQUITY_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_ANTIQUITY2_ID_CONCEPT)) Then aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_ANTIQUITY3_ID_CONCEPT)) Then aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_ANTIQUITY4_ID_CONCEPT)) Then aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_FOR_RISK_CONCEPT)) Then aConceptComponent(N_FOR_RISK_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_GENDER_ID_CONCEPT)) Then aConceptComponent(N_GENDER_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_HAS_CHILDREN_CONCEPT)) Then aConceptComponent(N_HAS_CHILDREN_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT)) Then aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_HAS_SYNDICATE_CONCEPT)) Then aConceptComponent(N_HAS_SYNDICATE_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT)) Then aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
		If Not IsNumeric(aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT)) Then aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) = 30000000
		If Not IsNumeric(aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT)) Then aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT) = 0
		If Not IsNumeric(aConceptComponent(N_CURRENCY_ID_CONCEPT)) Then aConceptComponent(N_CURRENCY_ID_CONCEPT) = 0
		If Not IsNumeric(aConceptComponent(N_CONCEPT_QTTY_ID_CONCEPT)) Then aConceptComponent(N_CONCEPT_QTTY_ID_CONCEPT) = 1
		If Not IsNumeric(aConceptComponent(N_CONCEPT_TYPE_ID_CONCEPT)) Then aConceptComponent(N_CONCEPT_TYPE_ID_CONCEPT) = 3
		If Len(aConceptComponent(S_APPLIES_ID_CONCEPT)) = 0 Then aConceptComponent(S_APPLIES_ID_CONCEPT) = "-1"
		If Not IsNumeric(aConceptComponent(D_CONCEPT_MIN_CONCEPT)) Then aConceptComponent(D_CONCEPT_MIN_CONCEPT) = 0
		If Not IsNumeric(aConceptComponent(N_CONCEPT_MIN_QTTY_ID_CONCEPT)) Then aConceptComponent(N_CONCEPT_MIN_QTTY_ID_CONCEPT) = 1
		If Not IsNumeric(aConceptComponent(D_CONCEPT_MAX_CONCEPT)) Then aConceptComponent(D_CONCEPT_MAX_CONCEPT) = 0
		If Not IsNumeric(aConceptComponent(N_CONCEPT_MAX_QTTY_ID_CONCEPT)) Then aConceptComponent(N_CONCEPT_MAX_QTTY_ID_CONCEPT) = 1
		If Not IsNumeric(aConceptComponent(N_START_USER_ID_CONCEPT)) Then aConceptComponent(N_START_USER_ID_CONCEPT) = aLoginComponent(N_USER_ID_LOGIN)
		If Not IsNumeric(aConceptComponent(N_END_USER_ID_CONCEPT)) Then aConceptComponent(N_END_USER_ID_CONCEPT) = -1
		If Not IsNumeric(aConceptComponent(N_STATUS_ID_CONCEPT)) Then aConceptComponent(N_STATUS_ID_CONCEPT) = -2
	End If

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del registro contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckConceptInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayConceptHistoryList(oRequest, oADODBConnection, bForExport, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To display the changes to the information for the
'		  given concept
'Inputs:  oRequest, oADODBConnection, bForExport, aConceptComponent
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayConceptHistoryList"
	Dim oRecordset
	Dim sNames
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If aConceptComponent(N_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptsHistoryList.*, UserName, UserLastName, CurrencySymbol, StatusName From ConceptsHistoryList, Users, Currencies, StatusConcepts Where (ConceptsHistoryList.UserID=Users.UserID) And (ConceptsHistoryList.CurrencyID=Currencies.CurrencyID) And (ConceptsHistoryList.StatusID=StatusConcepts.StatusID) And (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") Order By ConceptDate", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Response.Write "<TABLE WIDTH=""800"" BORDER="""
			If bForExport Then
				Response.Write "1"
			Else
				Response.Write "0"
			End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("Fecha,Responsable del cambio,Concepto al que pertenece,Monto,Moneda,Estatus", ",", -1, vbBinaryCompare)
				asCellWidths = Split("200,150,150,100,100,100", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split(",,,RIGHT,,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sRowContents = DisplayDateFromSerialNumber(CLng(oRecordset.Fields("ConceptDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("UserName").Value) & " " & CStr(oRecordset.Fields("UserLastName").Value))
					Call GetNameFromTable(oADODBConnection, "Concept", CLng(oRecordset.Fields("ParentID").Value), "", "", sNames, sErrorDescription)
					If Len(sNames) = 0 Then sNames = "Ninguna"
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(sNames)
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CLng(oRecordset.Fields("Amount").Value), 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CurrencySymbol").Value))
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
	DisplayConceptHistoryList = lErrorNumber
	Err.Clear
End Function

Function DisplayConceptForm(oRequest, oADODBConnection, sAction, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a concept from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aConceptComponent
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayConceptForm"
	Dim lErrorNumber

	If aConceptComponent(N_ID_CONCEPT) <> -1 Then
		lErrorNumber = GetConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckConceptFields(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If Len(oRequest("Delete").Item) > 0 Then Response.Write "return true;" & vbNewLine
					Response.Write "if (oForm.ConceptShortName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de introducir la clave del registro.');" & vbNewLine
						Response.Write "oForm.ConceptShortName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (oForm.ConceptName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de introducir el nombre del registro.');" & vbNewLine
						Response.Write "oForm.ConceptName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (! CheckIntegerValue(oForm.OrderInList, 'el orden para los reportes', N_MINIMUM_ONLY_FLAG, N_OPEN_FLAG, 0, 0))" & vbNewLine
						Response.Write "return false;" & vbNewLine
					If False Then
						Response.Write "oForm.TaxAmount.value = oForm.TaxAmount.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.TaxAmount, 'el monto gravable del concepto', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.TaxMin.value = oForm.TaxMin.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.TaxMin, 'el monto mínimo gravable del concepto', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.TaxMax.value = oForm.TaxMax.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.TaxMax, 'el monto máximo gravable del concepto', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.ExemptAmount.value = oForm.ExemptAmount.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.ExemptAmount, 'el monto exento del concepto', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.ExemptMin.value = oForm.ExemptMin.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.ExemptMin, 'el monto mínimo exento del concepto', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.ExemptMax.value = oForm.ExemptMax.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.ExemptMax, 'el monto máximo exento del concepto', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
					End If
					Response.Write "SetTaxValue();" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "if (((parseInt('1' + oForm.EndDay.value) - 100) + ('1' + parseInt(oForm.EndMonth.value) - 100) + parseInt(oForm.EndYear.value)) > 0 ) {" & vbNewLine
					Response.Write "if ((parseInt('1' + oForm.EndDay.value) - 100) * (parseInt('1' + oForm.EndMonth.value) - 100) * parseInt(oForm.EndYear.value) > 0 ) {" & vbNewLine
						Response.Write "if (((parseInt('1' + oForm.StartDay.value) - 100) + ((parseInt('1' + oForm.StartMonth.value) -100) * 100) + parseInt(oForm.StartYear.value) * 10000) > ((parseInt('1' + oForm.EndDay.value) - 100) + ((parseInt('1' + oForm.EndMonth.value) - 100) * 100) + parseInt(oForm.EndYear.value) * 10000)) {" & vbNewLine
							Response.Write "alert('Favor de verificar la vigencia del registro de crédito del empleado.');" & vbNewLine
							Response.Write "oForm.EndDay.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "else" & vbNewLine
					Response.Write "{" & vbNewLine
						Response.Write "alert('Favor de verificar la vigencia del movimiento');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckConceptFields" & vbNewLine
			Response.Write "function ShowAmountFields(sValue, sFieldsName) {" & vbNewLine
				Response.Write "var oForm = document.ConceptFrm;" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "if (sValue == '1')" & vbNewLine
						Response.Write "ShowDisplay(document.all[sFieldsName + 'CurrencySpn']);" & vbNewLine
					Response.Write "else" & vbNewLine
						Response.Write "HideDisplay(document.all[sFieldsName + 'CurrencySpn']);" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of ShowAmountFields" & vbNewLine
			Response.Write "function SetTaxValue() {" & vbNewLine
				Response.Write "var oForm = document.ConceptFrm;" & vbNewLine
				Response.Write "var i;" & vbNewLine
				Response.Write "var j;" & vbNewLine
					Response.Write "for (i=0;i<oForm.ForTax.length;i++){" & vbNewLine
						Response.Write "if (oForm.ForTax[i].checked)" & vbNewLine
							Response.Write "break;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "for (j=0;j<oForm.IsDeduction.length;j++){" & vbNewLine
						Response.Write "if (oForm.IsDeduction[j].checked)" & vbNewLine
							Response.Write "break;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "if (oForm.IsDeduction[j].value == 0) {" & vbNewLine
						'Response.Write "alert('IsDeduction es false');" & vbNewLine
						Response.Write "if (oForm.ForTax[i].value == 1) {" & vbNewLine
							Response.Write "oForm.TaxAmount.value=100;" & vbNewLine
							'Response.Write "alert('ForTax =' + oForm.TaxAmount.value);" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "else {" & vbNewLine
							Response.Write "oForm.TaxAmount.value=0;" & vbNewLine
							'Response.Write "alert('ForTax =' + oForm.TaxAmount.value);" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "else {" & vbNewLine
						'Response.Write "alert('IsDeduction es true');" & vbNewLine
						Response.Write "if (oForm.ForTax[i].value == 1) {" & vbNewLine
							Response.Write "oForm.TaxAmount.value=0;" & vbNewLine
							'Response.Write "alert('ForTax =' + oForm.TaxAmount.value);" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "else {" & vbNewLine
							Response.Write "oForm.TaxAmount.value=100;" & vbNewLine
							'Response.Write "alert('ForTax =' + oForm.TaxAmount.value);" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of ShowAmountFields" & vbNewLine

		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""ConceptFrm"" ID=""ConceptFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckConceptFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Concepts"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptID"" ID=""ConceptIDHdn"" VALUE=""" & aConceptComponent(N_ID_CONCEPT) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate"" ID=""StartDateHdn"" VALUE=""" & aConceptComponent(N_START_DATE_CONCEPT) & """ />"

			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaxAmount"" ID=""TaxAmountHdn"" VALUE=""" & aConceptComponent(D_TAX_AMOUNT_CONCEPT) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaxQttyID"" ID=""TaxQttyIDHdn"" VALUE=""" & aConceptComponent(N_TAX_QTTY_ID_CONCEPT) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaxCurrencyID"" ID=""TaxCurrencyIDHdn"" VALUE=""" & aConceptComponent(N_TAX_CURRENCY_ID_CONCEPT) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaxMin"" ID=""TaxMinHdn"" VALUE=""" & aConceptComponent(D_TAX_MIN_CONCEPT) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaxMinQttyID"" ID=""TaxMinQttyIDHdn"" VALUE=""" & aConceptComponent(N_TAX_MIN_QTTY_ID_CONCEPT) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaxMax"" ID=""TaxMaxHdn"" VALUE=""" & aConceptComponent(D_TAX_MAX_CONCEPT) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaxMaxQttyID"" ID=""TaxMaxQttyIDHdn"" VALUE=""" & aConceptComponent(N_TAX_MAX_QTTY_ID_CONCEPT) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ExemptAmount"" ID=""ExemptAmountHdn"" VALUE=""" & aConceptComponent(D_EXEMPT_AMOUNT_CONCEPT) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ExemptQttyID"" ID=""ExemptQttyIDHdn"" VALUE=""" & aConceptComponent(N_EXEMPT_QTTY_ID_CONCEPT) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ExemptCurrencyID"" ID=""ExemptCurrencyIDHdn"" VALUE=""" & aConceptComponent(N_EXEMPT_CURRENCY_ID_CONCEPT) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ExemptMin"" ID=""ExemptMinHdn"" VALUE=""" & aConceptComponent(D_EXEMPT_MIN_CONCEPT) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ExemptMinQttyID"" ID=""ExemptMinQttyIDHdn"" VALUE=""" & aConceptComponent(N_EXEMPT_MIN_QTTY_ID_CONCEPT) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ExemptMax"" ID=""ExemptMaxHdn"" VALUE=""" & aConceptComponent(D_EXEMPT_MAX_CONCEPT) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ExemptMaxQttyID"" ID=""ExemptMaxQttyIDHdn"" VALUE=""" & aConceptComponent(N_EXEMPT_MAX_QTTY_ID_CONCEPT) & """ />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clave:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""ConceptShortName"" ID=""ConceptShortNameTxt"" SIZE=""10"" MAXLENGTH=""5"" VALUE=""" & CleanStringForHTML(aConceptComponent(S_SHORT_NAME_CONCEPT)) & """ CLASS=""TextFields"""
					If Len(oRequest("Change").Item) > 0 Then
						Response.Write  " READONLY=""READONLY"" /></TD>"
					Else
						Response.Write " /></TD>"
					End If
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""ConceptName"" ID=""ConceptNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & CleanStringForHTML(aConceptComponent(S_NAME_CONCEPT)) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write DisplayDateCombosUsingSerial(aConceptComponent(N_START_DATE_CONCEPT), "Start", N_START_YEAR, Year(Date) + 1, True, False)
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write DisplayDateCombosUsingSerial(aConceptComponent(N_END_DATE_CONCEPT), "End", Year(Date())-2, Year(Date())+2, True, True)
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Partida:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""BudgetID"" ID=""BudgetIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE=""-1"">Ninguna</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Budgets As Partida, Budgets As SubPartida, Budgets As Tipo", "Tipo.BudgetID", "'Partida: ' As Temp1, Partida.BudgetName, '. SubPartida: ' As Temp2, SubPartida.BudgetName, '. Tipo: ' As Temp3, Tipo.BudgetShortName, Tipo.BudgetName", "(Partida.BudgetID=SubPartida.ParentID) And (SubPartida.BudgetID=Tipo.ParentID) And (SubPartida.ParentID>-1) And (Tipo.BudgetID>-1) And (Tipo.Active=1)", "Partida.BudgetName, SubPartida.BudgetName, Tipo.BudgetName", aConceptComponent(N_BUDGET_ID_CONCEPT), "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de nómina:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""PayrollTypeID"" ID=""PayrollTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PayrollTypes", "PayrollTypeID", "PayrollTypeName", "", "PayrollTypeID", aConceptComponent(N_PAYROLL_TYPE_ID_CONCEPT), "Ordinaria;;;1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Periodicidad:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""PeriodID"" ID=""PeriodIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Periods", "PeriodID", "PeriodName", "(PeriodID>-1) And (Active=1)", "PeriodID", aConceptComponent(N_PERIOD_ID_CONCEPT), "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write "¿Es una deducción?&nbsp;"
						Response.Write "<INPUT TYPE=""Radio"" NAME=""IsDeduction"" ID=""IsDeductionRd"" VALUE=""1"""
							If aConceptComponent(N_IS_DEDUCTION_CONCEPT) = 1 Then
								Response.Write " CHECKED=""1"""
							End If
						Response.Write " />Sí&nbsp;&nbsp;&nbsp;<INPUT TYPE=""Radio"" NAME=""IsDeduction"" ID=""IsDeductionRd"" VALUE=""0"""
							If aConceptComponent(N_IS_DEDUCTION_CONCEPT) = 0 Then
								Response.Write " CHECKED=""1"""
							End If
						Response.Write " />No<BR />"
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"

				Response.Write "<TR>"
					Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write "¿Aplica para la pensión alimenticia?&nbsp;"
						Response.Write "<INPUT TYPE=""Radio"" NAME=""ForAlimony"" ID=""ForAlimonyRd"" VALUE=""1"""
							If aConceptComponent(N_FOR_ALIMONY_CONCEPT) = 1 Then
								Response.Write " CHECKED=""1"""
							End If
						Response.Write " />Sí&nbsp;&nbsp;&nbsp;<INPUT TYPE=""Radio"" NAME=""ForAlimony"" ID=""ForAlimonyRd"" VALUE=""0"""
							If aConceptComponent(N_FOR_ALIMONY_CONCEPT) = 0 Then
								Response.Write " CHECKED=""1"""
							End If
						Response.Write " />No<BR />"
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write "¿Se paga si tiene licencia sin sueldo?&nbsp;"
						Response.Write "<INPUT TYPE=""Radio"" NAME=""OnLeave"" ID=""OnLeaveRd"" VALUE=""1"""
							If aConceptComponent(N_ON_LEAVE_CONCEPT) = 1 Then
								Response.Write " CHECKED=""1"""
							End If
						Response.Write " />Sí&nbsp;&nbsp;&nbsp;<INPUT TYPE=""Radio"" NAME=""OnLeave"" ID=""OnLeaveRd"" VALUE=""0"""
							If aConceptComponent(N_ON_LEAVE_CONCEPT) = 0 Then
								Response.Write " CHECKED=""1"""
							End If
						Response.Write " />No<BR />"
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write "¿Aplican impuestos?&nbsp;"
						Response.Write "<INPUT TYPE=""Radio"" NAME=""ForTax"" ID=""ForTaxRd"" VALUE=""1"" "
							If aConceptComponent(N_FOR_TAX_CONCEPT) = 1 Then
								Response.Write " CHECKED=""1"""
							End If
						Response.Write " />Sí&nbsp;&nbsp;&nbsp;<INPUT TYPE=""Radio"" NAME=""ForTax"" ID=""ForTaxRd"" VALUE=""0"" "
							If aConceptComponent(N_FOR_TAX_CONCEPT) = 0 Then
								Response.Write " CHECKED=""1"""
							End If
						Response.Write " />No<BR />"
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD COLSPAN=""2"">"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Orden para los reportes:&nbsp;</FONT>"
						Response.Write "<INPUT TYPE=""TEXT"" NAME=""OrderInList"" ID=""OrderInListTxt"" SIZE=""3"" MAXLENGTH=""3"" VALUE=""" & CleanStringForHTML(aConceptComponent(N_ORDER_IN_LIST_CONCEPT)) & """ CLASS=""TextFields"" />"
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Es crédito:&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""CHECKBOX"" NAME=""IsCredit"" ID=""IsCreditChk"" SIZE=""2"" MAXLENGTH=""2"" VALUE=""1"" onclick=""if (this.checked) {ShowDisplay(document.all['ReasonsDiv']) } else {HideDisplay(document.all['ReasonsDiv'])}; "" "
						If CInt(aConceptComponent(N_IS_CREDIT)) > 0 Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write "/></TD>"
				Response.Write "</TR>"
				If False Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Monto gravable:&nbsp;</NOBR></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""TaxAmount"" ID=""TaxAmountTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aConceptComponent(D_TAX_AMOUNT_CONCEPT), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
							Response.Write "<SELECT NAME=""TaxQttyID"" ID=""TaxQttyIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'Tax');"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID Not In (23,24,25,33,34,35))", "QttyID", aConceptComponent(N_TAX_QTTY_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
							Response.Write "<SPAN ID=""TaxCurrencySpn""><SELECT NAME=""TaxCurrencyID"" ID=""TaxCurrencyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Currencies", "CurrencyID", "CurrencyName", "(CurrencyID>-1) And (Active=1)", "CurrencyName", aConceptComponent(N_TAX_CURRENCY_ID_CONCEPT), "Ninguna;;;-1", sErrorDescription)
							Response.Write "</SELECT></SPAN>"
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Mínimo:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""TaxMin"" ID=""TaxMinTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aConceptComponent(D_TAX_MIN_CONCEPT), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
							Response.Write "<SELECT NAME=""TaxMinQttyID"" ID=""TaxMinQttyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID In (1,3,4,5,13))", "QttyID", aConceptComponent(N_TAX_MIN_QTTY_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Máximo:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""TaxMax"" ID=""TaxMaxTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aConceptComponent(D_TAX_MAX_CONCEPT), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
							Response.Write "<SELECT NAME=""TaxMaxQttyID"" ID=""TaxMaxQttyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID In (1,3,4,5,13))", "QttyID", aConceptComponent(N_TAX_MAX_QTTY_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Monto exento:&nbsp;</NOBR></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""ExemptAmount"" ID=""ExemptAmountTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aConceptComponent(D_EXEMPT_AMOUNT_CONCEPT), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
							Response.Write "<SELECT NAME=""ExemptQttyID"" ID=""ExemptQttyIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value, 'Exempt');"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID Not In (23,24,25,33,34,35))", "QttyID", aConceptComponent(N_EXEMPT_QTTY_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
							Response.Write "<SPAN ID=""ExemptCurrencySpn""><SELECT NAME=""ExemptCurrencyID"" ID=""ExemptCurrencyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Currencies", "CurrencyID", "CurrencyName", "(CurrencyID>-1) And (Active=1)", "CurrencyName", aConceptComponent(N_EXEMPT_CURRENCY_ID_CONCEPT), "Ninguna;;;-1", sErrorDescription)
							Response.Write "</SELECT></SPAN>"
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Mínimo:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""ExemptMin"" ID=""ExemptMinTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aConceptComponent(D_EXEMPT_MIN_CONCEPT), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
							Response.Write "<SELECT NAME=""ExemptMinQttyID"" ID=""ExemptMinQttyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID In (1,3,4,5,13))", "QttyID", aConceptComponent(N_EXEMPT_MIN_QTTY_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Máximo:&nbsp;</FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""ExemptMax"" ID=""ExemptMaxTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aConceptComponent(D_EXEMPT_MAX_CONCEPT), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
							Response.Write "<SELECT NAME=""ExemptMaxQttyID"" ID=""ExemptMaxQttyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID In (1,3,4,5,13))", "QttyID", aConceptComponent(N_EXEMPT_MAX_QTTY_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
				End If
			Response.Write "</TABLE>"
			Response.Write "<DIV NAME=""ReasonsDiv"" ID=""ReasonsDiv"" STYLE=""display: none"">"
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Es tercero:&nbsp;</NOBR></FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""CHECKBOX"" NAME=""IsOther"" ID=""IsOtherChk"" SIZE=""2"" MAXLENGTH=""2"" VALUE=""1"" "
							If CInt(aConceptComponent(N_IS_OTHER)) > 0 Then
								Response.Write " CHECKED=""1"""
							End If
						Response.Write "/></TD>"
					Response.Write "</TR>"
				Response.Write "</TABLE><BR />"
			Response.Write "</DIV>"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "if (document.ConceptFrm.IsCredit.checked) {ShowDisplay(document.all['ReasonsDiv']) } else {HideDisplay(document.all['ReasonsDiv'])};" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
			'Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			'	Response.Write "ShowAmountFields(document.ConceptFrm.TaxQttyID.value, 'Tax');" & vbNewLine
			'	Response.Write "ShowAmountFields(document.ConceptFrm.ExemptQttyID.value, 'Exempt');" & vbNewLine
			'Response.Write "//--></SCRIPT>" & vbNewLine
			If aConceptComponent(N_ID_CONCEPT) = -1 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveConceptWngDiv']); ConceptFrm.Remove.focus()"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "&StartDate=" & aConceptComponent(N_START_DATE_CONCEPT) & "'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveConceptWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayConceptForm = lErrorNumber
	Err.Clear
End Function

Function DisplayConceptValuesForm(oRequest, oADODBConnection, sAction, bFull, iSelectedTab, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a concept from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, bFull, iSelectedTab, aConceptComponent
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayConceptValuesForm"
	Dim sNames
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sEmployeeTypeName

	Select Case iSelectedTab
		Case 0
			sEmployeeTypeName = "Médica, paramédica y grupos afines"
		Case 1
			sEmployeeTypeName = "Funcionario"
		Case 2
			sEmployeeTypeName = "Operativo"
		Case 3
			sEmployeeTypeName = "Alta responsabilidad"
		Case 4
			sEmployeeTypeName = "Enlace"
		Case 5
			sEmployeeTypeName = "Residente"
		Case 6
			sEmployeeTypeName = "Becario"
	End Select
	If bFull Then
		If Len(oRequest("RecordID").Item) > 0 Then
			'lErrorNumber = GetConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
			If lErrorNumber = 0 Then lErrorNumber = GetConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
		End If
	Else
		If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
			bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
			If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
				Call InitializeConceptComponent(oRequest, aConceptComponent)
			End If
			aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) = -1
		Else
			lErrorNumber = GetConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
			'lErrorNumber = GetConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
		End If
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckConceptValuesFields(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If Len(oRequest("Delete").Item) > 0 Then Response.Write "return true;" & vbNewLine
					If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
						If bFull Then
							Response.Write "if (oForm.HasClassificationID.checked)" & vbNewLine
								Response.Write "if (!CheckIntegerValue(oForm.ClassificationID, 'la clasificación', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
									Response.Write "return false;" & vbNewLine
							Response.Write "if (oForm.HasIntegrationID.checked)" & vbNewLine
								Response.Write "if (!CheckIntegerValue(oForm.IntegrationID, 'la integración', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
									Response.Write "return false;" & vbNewLine
						End If
						If iSelectedTab = 0 Then
							Response.Write "if (oForm.HasWorkingHours.checked)" & vbNewLine
								Response.Write "if (!CheckFloatValue(oForm.WorkingHours, 'las horas trabajadas', N_BOTH_FLAG, N_CLOSED_FLAG, 0, 24))" & vbNewLine
									Response.Write "return false;" & vbNewLine
						End If
					End If
					'Response.Write "if (parseInt((oForm.StartForValueYear.value)*10000) + (parseInt(oForm.StartForValueMonth.value)*100) + parseInt(oForm.StartForValueDay.value) < oForm.StartDateForValueConcept.value) {" & vbNewLine
					'	Response.Write "alert('Favor de seleccionar una vigencia mayor a la actual.');" & vbNewLine
					'	Response.Write "return false;" & vbNewLine
					'Response.Write "}" & vbNewLine
					Response.Write "oForm.ConceptAmount.value = oForm.ConceptAmount.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "if (! CheckFloatValue(oForm.ConceptAmount, 'el monto del concepto', N_NO_RANK_FLAG, N_OPEN_FLAG, 0, 0))" & vbNewLine
						Response.Write "return false;" & vbNewLine

					Response.Write "if (((oForm.ConceptQttyID.value == '2') || (oForm.ConceptQttyID.value == '8')) && (GetSelectedValues(oForm.AppliesToID) == '')) {" & vbNewLine
						Response.Write "alert('Seleccione el(los) concepto(s) que se utiliza(n) para calcular el concepto');" & vbNewLine
						Response.Write "oForm.AppliesToID.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					If bFull Then
						Response.Write "oForm.ConceptMin.value = oForm.ConceptMin.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.ConceptMin, 'el monto mínimo', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.ConceptMax.value = oForm.ConceptMax.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (! CheckFloatValue(oForm.ConceptMax, 'el monto máximo', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
					End If
					Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckConceptValuesFields" & vbNewLine
			
			Response.Write "function ShowAmountFields(sValue) {" & vbNewLine
				Response.Write "var oForm = document.ConceptValuesFrm;" & vbNewLine

				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "HideDisplay(document.all['ConceptCurrencySpn']);" & vbNewLine
					Response.Write "HideDisplay(document.all['ConceptAppliesToSpn']);" & vbNewLine
					Response.Write "switch (sValue) {" & vbNewLine
						Response.Write "case '1':" & vbNewLine
							Response.Write "ShowDisplay(document.all['ConceptCurrencySpn']);" & vbNewLine
							Response.Write "break;" & vbNewLine
						Response.Write "case '2':" & vbNewLine
							Response.Write "ShowDisplay(document.all['ConceptAppliesToSpn']);" & vbNewLine
							Response.Write "break;" & vbNewLine
						Response.Write "case '8':" & vbNewLine
							Response.Write "ShowDisplay(document.all['ConceptAppliesToSpn']);" & vbNewLine
							Response.Write "break;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of ShowAmountFields" & vbNewLine

			Response.Write " function ShowConceptField(sField, bShow) {" & vbNewLine
				Response.Write "var sForm = 'document.ConceptValuesFrm.';" & vbNewLine
				Response.Write "var oField = eval(sForm + sField);" & vbNewLine
				Response.Write "if (oField) {" & vbNewLine
					Response.Write "if (bShow) {" & vbNewLine
						Response.Write "oField.value = '';" & vbNewLine
						Response.Write "ShowDisplay(document.all['Has' + sField + 'Div']);" & vbNewLine
						Response.Write "HideDisplay(document.all['NotHas' + sField + 'Div']);" & vbNewLine
					Response.Write "} else {" & vbNewLine
						Response.Write "oField.value = '-1';" & vbNewLine
						Response.Write "HideDisplay(document.all['Has' + sField + 'Div']);" & vbNewLine
						Response.Write "ShowDisplay(document.all['NotHas' + sField + 'Div']);" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of ShowConceptField" & vbNewLine

			Response.Write "function DisplayInfoMessage() {" & vbNewLine
				'Response.Write "HideDisplay(document.all['PositionIDMessageDiv'])" & vbNewLine
				'Response.Write "if (document.ConceptValuesFrm.PositionID.value == -1){" & vbNewLine
				'	Response.Write "ShowDisplay(document.all['PositionIDMessageDiv'])" & vbNewLine
				'Response.Write "}" & vbNewLine
				Response.Write "alert('Nivel ' + document.ConceptValuesFrm.LevelID.value)" & vbNewLine
			Response.Write "} // End of DisplayInfoMessage" & vbNewLine

		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<DIV NAME=""PositionIDMessageDiv"" ID=""PositionIDMessageDiv"" STYLE=""display: none"">"
				Call DisplayInstructionsMessage("Advertencia", "Si no se indica el puesto en el registro de tabuladores de pago, se añadira un registro para cada puesto activo del tabulador " & sEmployeeTypeName & ", con las condiciones que se establezcan en la pantalla de captura. Los puestos activos para este tabulador son los que estan cargados en el combo de selección de Puesto que se muestra abajo.")
		Response.Write "</DIV>"
		Response.Write "<FORM NAME=""ConceptValuesFrm"" ID=""ConceptValuesFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckConceptValuesFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""ConceptsValues"" />"
			'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SaveConceptValue"" ID=""SaveConceptValueHdn"" VALUE=""1"" />"
			If iSelectedTab <> -1 Then
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SaveConceptValue"" ID=""SaveConceptValueHdn"" VALUE=""1"" />"
			End If
			
			If bFull Then
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptID"" ID=""ConceptIDHdn"" VALUE=""" & aConceptComponent(N_ID_CONCEPT) & """ />"
			Else
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""bFull"" ID=""bFullHdn"" VALUE=""1"" />"
			End If	
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RecordID"" ID=""RecordIDHdn"" VALUE=""" & aConceptComponent(N_RECORD_ID_CONCEPT) & """ />"
			If aConceptComponent(N_STATUS_ID_CONCEPT) <> -1 Then
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StatusID"" ID=""StatusIDHdn"" VALUE=""" & aConceptComponent(N_STATUS_ID_CONCEPT) & """ />"
			Else
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StatusID"" ID=""StatusIDHdn"" VALUE=""0"" />"
			End If

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Tipo de tabulador:&nbsp;</B></FONT></TD>"
					If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 And iSelectedTab = -1 Then
						Response.Write "<TD><SELECT NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeName", "(EmployeeTypeID>=0) And (EmployeeTypeID<7) And (Active=1)", "EmployeeTypeName", aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Else
						If aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = -1 Then
							sNames = "Todos"
							'Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
						Else
							Response.Write "<TD><SELECT NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeName", "EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT), "EmployeeTypeName", aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
							'Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EmployeeTypes", "EmployeeTypeID", "EmployeeTypeName", "", "EmployeeTypeName", aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
							'Call GetNameFromTable(oADODBConnection, "EmployeeTypes", aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT), "", "", sNames, sErrorDescription)
						End If
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(sNames) & "</B></FONT></TD>"
						'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeTypeID"" ID=""EmployeeTypeIDHdn"" VALUE=""" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & """ />"
					End If	
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Compañía:&nbsp;</FONT></TD>"					
					If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
						Response.Write "<TD><SELECT NAME=""CompanyID"" ID=""CompanyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write "<OPTION VALUE=""-1"">Todas</OPTION>"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Companies", "CompanyID", "CompanyShortName, CompanyName", "(CompanyID>0) And (Active=1)", "CompanyShortName", aConceptComponent(N_COMPANY_ID_CONCEPT), "Ninguna;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Else
						If aConceptComponent(N_COMPANY_ID_CONCEPT) = -1 Then
							sNames = "Todas"
						Else
							Call GetNameFromTable(oADODBConnection, "Companies", aConceptComponent(N_COMPANY_ID_CONCEPT), "", "", sNames, sErrorDescription)
						End If
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CompanyID"" ID=""CompanyIDHdn"" VALUE=""" & aConceptComponent(N_COMPANY_ID_CONCEPT) & """ />"
					End If
				Response.Write "</TR>"
				If bFull Then
					Response.Write "<TR>"
						Call GetNameFromTable(oADODBConnection, "FullConcepts", aConceptComponent(N_ID_CONCEPT), "", "", sNames, "")
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Concepto:&nbsp;</B></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(sNames) & "</B></FONT></TD>"
					Response.Write "</TR>"
				Else
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Concepto:&nbsp;</FONT></TD>"
						If aConceptComponent(N_ID_CONCEPT) = -1 Then
							Response.Write "<TD><SELECT NAME=""ConceptID"" ID=""ConceptIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Select Case iSelectedTab
									Case 0
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID In(1,38,49))", "ConceptID", aConceptComponent(N_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
									Case 1, 2, 3, 4
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID In(1,3))", "ConceptID", aConceptComponent(N_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
									Case 5
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID In(39,89))", "ConceptID", aConceptComponent(N_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
									Case 6
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID In(14))", "ConceptID", aConceptComponent(N_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
									Case Else
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID In(1,2,14,38,49,89))", "ConceptID", aConceptComponent(N_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
								End Select
							Response.Write "</SELECT></TD>"
						Else
							Call GetNameFromTable(oADODBConnection, "Concepts", aConceptComponent(N_ID_CONCEPT), "", "", sNames, sErrorDescription)
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptID"" ID=""ConceptIDHdn"" VALUE=""" & aConceptComponent(N_ID_CONCEPT) & """ />"
						End If
					Response.Write "</TR>"
				End If
				If (iSelectedTab = 0) Or (iSelectedTab = 2) Or bFull Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de puesto:&nbsp;</FONT></TD>"
						If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
							Response.Write "<TD><SELECT NAME=""PositionTypeID"" ID=""PositionTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PositionTypes", "PositionTypeID", "PositionTypeShortName, PositionTypeName", "(PositionTypeID>0) And (Active=1)", "PositionTypeShortName", aConceptComponent(N_POSITION_TYPE_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Else
							If aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) = -1 Then
								sNames = "Todos"
							Else
								Call GetNameFromTable(oADODBConnection, "PositionTypes", aConceptComponent(N_POSITION_TYPE_ID_CONCEPT), "", "", sNames, sErrorDescription)
							End If
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PositionTypeID"" ID=""PositionTypeIDHdn"" VALUE=""" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & """ />"
						End If
					Response.Write "</TR>"
				End If
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
					If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
						If aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = -1 Then
							Response.Write "<TD><SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE="""">Todos</OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID", "PositionShortName, PositionName, 'Cia:' As Temp, CompanyID", "(PositionID>-1)", "PositionShortName", aConceptComponent(N_POSITION_ID_CONCEPT), "", sErrorDescription)
							Response.Write "</SELECT></TD>"
						ElseIf aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = 1 Then
							Response.Write "<TD><SELECT NAME=""PositionIDTemp"" ID=""PositionIDTempCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""var asTemp = this.value.split(','); document.ConceptValuesFrm.PositionID.value = asTemp[0]; document.ConceptValuesFrm.GroupGradeLevelID.value = asTemp[1]; document.ConceptValuesFrm.ClassificationID.value = asTemp[2]; document.ConceptValuesFrm.IntegrationID.value = asTemp[3];"">"
								Response.Write "<OPTION VALUE=""-1,-1,-1,-1"">Todos</OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions, GroupGradeLevels", "PositionID, Positions.GroupGradeLevelID, ClassificationID, IntegrationID", "PositionShortName, PositionName, 'GGN:' As Temp1, GroupGradeLevelShortName, 'Clas:' As Temp2, ClassificationID, 'Int:' As Temp3, IntegrationID, 'Cia:' As Temp, CompanyID", "(Positions.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (PositionID>-1) And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ")", "PositionShortName", aConceptComponent(N_POSITION_ID_CONCEPT), "", sErrorDescription)
							Response.Write "</SELECT></TD>"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PositionID"" ID=""PositionIDHdn"" VALUE=""" & aConceptComponent(N_POSITION_ID_CONCEPT) & """ />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GroupGradeLevelID"" ID=""GroupGradeLevelIDHdn"" VALUE=""" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & """ />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ClassificationID"" ID=""ClassificationIDHdn"" VALUE=""" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & """ />"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IntegrationID"" ID=""IntegrationIDHdn"" VALUE=""" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & """ />"
						Else
							Response.Write "<TD><SELECT NAME=""PositionIDTemp"" ID=""PositionIDTempCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""var asTemp = this.value.split(','); document.ConceptValuesFrm.PositionID.value = asTemp[0]; document.ConceptValuesFrm.LevelID.value = asTemp[1];"">"
								Response.Write "<OPTION VALUE=""-1,-1"">Todos</OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Positions", "PositionID, LevelID", "PositionShortName, PositionName, 'Nivel:' As Temp, LevelID, 'Cia:' As Temp, CompanyID", "(PositionID>-1) And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ")", "PositionShortName", aConceptComponent(N_POSITION_ID_CONCEPT) & "," & aConceptComponent(N_LEVEL_ID_CONCEPT), "", sErrorDescription)
							Response.Write "</SELECT></TD>"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PositionID"" ID=""PositionIDHdn"" VALUE=""" & aConceptComponent(N_POSITION_ID_CONCEPT) & """ />"
						End If
					Else
						If aConceptComponent(N_POSITION_ID_CONCEPT) = -1 Then
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Todos</FONT></TD>"
						Else
							Call GetNameFromTable(oADODBConnection, "Positions", aConceptComponent(N_POSITION_ID_CONCEPT), "", "", sNames, sErrorDescription)
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
						End If
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PositionID"" ID=""PositionIDHdn"" VALUE=""" & aConceptComponent(N_POSITION_ID_CONCEPT) & """ />"
					End If
				Response.Write "</TR>"
				If (InStr(1, ",0,2,3,4,6,", ("," & iSelectedTab & ","), vbBinaryCompare) > 0) Or bFull Then
					If aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) <> 1 Then
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nivel:&nbsp;</FONT></TD>"
							If aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) = -1 Then
								If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
									Response.Write "<TD><SELECT NAME=""LevelID"" ID=""LevelIDCmb"" SIZE=""1"" CLASS=""Lists"">"
										Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Levels", "LevelID", "LevelName", "(LevelID>-1) And (Active=1)", "LevelName", aConceptComponent(N_LEVEL_ID_CONCEPT), "", sErrorDescription)
									Response.Write "</SELECT></TD>"
								Else
									If aConceptComponent(N_LEVEL_ID_CONCEPT) = -1 Then
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Todos</FONT></TD>"
									Else
										Call GetNameFromTable(oADODBConnection, "Levels", aConceptComponent(N_LEVEL_ID_CONCEPT), "", "", sNames, sErrorDescription)
										sNames = Right(("00" & sNames), Len("000"))
										Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(Left(sNames, Len("00")) & "-" & Right(sNames, Len("0"))) & "</FONT></TD>"
									End If
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""LevelID"" ID=""LevelIDHdn"" VALUE=""" & aConceptComponent(N_LEVEL_ID_CONCEPT) & """ />"
								End If
							Else
								Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""LevelID"" ID=""LevelIDHdn"" SIZE=""3"" MAXLENGTH=""3"" VALUE=""" & aConceptComponent(N_LEVEL_ID_CONCEPT) & """ CLASS=""TextFields"" DISABLED=""1"" /></TD>"
							End If
						Response.Write "</TR>"
					Else
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""LevelID"" ID=""LevelIDHdn"" VALUE=""" & aConceptComponent(N_LEVEL_ID_CONCEPT) & """ />"
					End If
				Else
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""LevelID"" ID=""LevelIDHdn"" VALUE=""" & aConceptComponent(N_LEVEL_ID_CONCEPT) & """ />"
				End If
				If aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) <> 1 Then
					If (iSelectedTab = 1) Or bFull Then
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Grupo, grado, nivel:&nbsp;</FONT></TD>"
							If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
								Response.Write "<TD><SELECT NAME=""GroupGradeLevelID"" ID=""GroupGradeLevelIDCmb"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "GroupGradeLevels", "GroupGradeLevelID", "GroupGradeLevelName", "(GroupGradeLevelID>-1) And (Active=1)", "GroupGradeLevelName", aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT), "", sErrorDescription)
								Response.Write "</SELECT></TD>"
							Else
								If aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) = -1 Then
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Todos</FONT></TD>"
								Else
									Call GetNameFromTable(oADODBConnection, "GroupGradeLevels", aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT), "", "", sNames, sErrorDescription)
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
								End If
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GroupGradeLevelID"" ID=""GroupGradeLevelIDHdn"" VALUE=""" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & """ />"
							End If
						Response.Write "</TR>"
					End If
				End If
				If bFull Then
					'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeStatusID"" ID=""EmployeeStatusIDHdn"" VALUE=""" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & """ />"
 					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estatus del empleado:&nbsp;</FONT></TD>"
						If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
							Response.Write "<TD><SELECT NAME=""EmployeeStatusID"" ID=""EmployeeStatusIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusEmployees", "StatusID", "StatusName", "(StatusID>-1) And (Active=1)", "StatusName", aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Else
							If aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) = -1 Then
								sNames = "Todos"
							Else
								Call GetNameFromTable(oADODBConnection, "StatusEmployees", aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT), "", "", sNames, sErrorDescription)
							End If
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeStatusID"" ID=""EmployeeStatusIDHdn"" VALUE=""" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & """ />"
						End If
					Response.Write "</TR>"
 					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estatus de la plaza:&nbsp;</FONT></TD>"
						If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
							Response.Write "<TD><SELECT NAME=""JobStatusID"" ID=""JobStatusIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusJobs", "StatusID", "StatusName", "(StatusID>-1) And (Active=1)", "StatusName", aConceptComponent(N_JOB_STATUS_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Else
							If aConceptComponent(N_JOB_STATUS_ID_CONCEPT) = -1 Then
								sNames = "Todos"
							Else
								Call GetNameFromTable(oADODBConnection, "StatusJobs", aConceptComponent(N_JOB_STATUS_ID_CONCEPT), "", "", sNames, sErrorDescription)
							End If
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JobStatusID"" ID=""JobStatusIDHdn"" VALUE=""" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & """ />"
						End If
					Response.Write "</TR>"
				End If
				If aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) <> 1 Then
					If (iSelectedTab = 1) Or bFull Then
 						Response.Write "<TR>"
							If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
								Response.Write "<TD>"
									Response.Write "<INPUT TYPE=""CHECKBOX"" NAME=""HasClassificationID"" ID=""HasClassificationIDChk"""
										If aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) > -1 Then Response.Write " CHECKED=""1"""
									Response.Write " onClick=""ShowConceptField('ClassificationID', this.checked)""/>"
									Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Clasificación:&nbsp;</FONT>"
								Response.Write "</TD>"
								Response.Write "<TD><SPAN ID=""HasClassificationIDDiv"""
										If aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) = -1 Then Response.Write " STYLE=""display: none"""
									Response.Write "><INPUT TYPE=""TEXT"" NAME=""ClassificationID"" ID=""ClassificationIDTxt"" VALUE=""" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & """ SIZE=""2"" MAXLENGTH=""2"" CLASS=""TextFields"" /></SPAN><SPAN ID=""NotHasClassificationIDDiv"""
										If aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) > -1 Then Response.Write " STYLE=""display: none"""
									Response.Write "><FONT FACE=""Arial"" SIZE=""2"">N/A</FONT>"
								Response.Write "</SPAN></TD>"
							Else
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clasificación:&nbsp;</FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
									If aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) = -1 Then
										Response.Write "Todas"
									Else
										Response.Write aConceptComponent(N_CLASSIFICATION_ID_CONCEPT)
									End If
								Response.Write "</FONT></TD>"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ClassificationID"" ID=""ClassificationIDHdn"" VALUE=""" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & """ />"
							End If
						Response.Write "</TR>"
					End If
					If (iSelectedTab = 1) Or bFull Then
						Response.Write "<TR>"
							If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
								Response.Write "<TD>"
									Response.Write "<INPUT TYPE=""CHECKBOX"" NAME=""HasIntegrationID"" ID=""HasIntegrationIDChk"""
										If aConceptComponent(N_INTEGRATION_ID_CONCEPT) > -1 Then Response.Write " CHECKED=""1"""
									Response.Write " onClick=""ShowConceptField('IntegrationID', this.checked)""/>"
									Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Integración:&nbsp;</FONT>"
								Response.Write "</TD>"
								Response.Write "<TD><SPAN ID=""HasIntegrationIDDiv"""
										If aConceptComponent(N_INTEGRATION_ID_CONCEPT) = -1 Then Response.Write " STYLE=""display: none"""
									Response.Write "><INPUT TYPE=""TEXT"" NAME=""IntegrationID"" ID=""IntegrationIDTxt"" VALUE=""" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & """ SIZE=""2"" MAXLENGTH=""2"" CLASS=""TextFields"" /></SPAN><SPAN ID=""NotHasIntegrationIDDiv"""
										If aConceptComponent(N_INTEGRATION_ID_CONCEPT) > -1 Then Response.Write " STYLE=""display: none"""
									Response.Write "><FONT FACE=""Arial"" SIZE=""2"">N/A</FONT>"
								Response.Write "</SPAN></TD>"
							Else
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Integración:&nbsp;</FONT></TD>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
									If aConceptComponent(N_INTEGRATION_ID_CONCEPT) = -1 Then
										Response.Write "Todas"
									Else
										Response.Write aConceptComponent(N_INTEGRATION_ID_CONCEPT)
									End If
								Response.Write "</FONT></TD>"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IntegrationID"" ID=""IntegrationIDHdn"" VALUE=""" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & """ />"
							End If
						Response.Write "</TR>"
					End If
				End If
				If bFull Then
 					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Jornada:&nbsp;</FONT></TD>"
						If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
							Response.Write "<TD><SELECT NAME=""JourneyID"" ID=""JourneyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Journeys", "JourneyID", "JourneyName", "(Active=1)", "JourneyName", aConceptComponent(N_JOURNEY_ID_CONCEPT), "Ninguna;;;-1", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Else
							If aConceptComponent(N_JOURNEY_ID_CONCEPT) = -1 Then
								sNames = "Todas"
							Else
								Call GetNameFromTable(oADODBConnection, "Journeys", aConceptComponent(N_JOURNEY_ID_CONCEPT), "", "", sNames, sErrorDescription)
							End If
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JourneyID"" ID=""JourneyIDHdn"" VALUE=""" & aConceptComponent(N_JOURNEY_ID_CONCEPT) & """ />"
						End If
					Response.Write "</TR>"
				End If
				If (iSelectedTab = 0) Or bFull Then
					Response.Write "<TR>"
						If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
							Response.Write "<TD>"
								Response.Write "<INPUT TYPE=""CHECKBOX"" NAME=""HasWorkingHours"" ID=""HasWorkingHoursChk"""
									If aConceptComponent(D_WORKING_HOURS_CONCEPT) > -1 Then Response.Write " CHECKED=""1"""
								Response.Write " onClick=""ShowConceptField('WorkingHours', this.checked)""/>"
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Horas laboradas:&nbsp;</FONT>"
							Response.Write "</TD>"
							Response.Write "<TD><SPAN ID=""HasWorkingHoursDiv"""
									If aConceptComponent(D_WORKING_HOURS_CONCEPT) = -1 Then Response.Write " STYLE=""display: none"""
								Response.Write "><INPUT TYPE=""TEXT"" NAME=""WorkingHours"" ID=""WorkingHoursTxt"" VALUE=""" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & """ SIZE=""4"" MAXLENGTH=""4"" CLASS=""TextFields"" /></SPAN><SPAN ID=""NotHasWorkingHoursDiv"""
									If aConceptComponent(D_WORKING_HOURS_CONCEPT) > -1 Then Response.Write " STYLE=""display: none"""
								Response.Write "><FONT FACE=""Arial"" SIZE=""2"">N/A</FONT>"
							Response.Write "</SPAN></TD>"
						Else
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horas laboradas:&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
								If aConceptComponent(D_WORKING_HOURS_CONCEPT) = -1 Then
									Response.Write "Todas"
								Else
									Response.Write aConceptComponent(D_WORKING_HOURS_CONCEPT)
								End If
							Response.Write "</FONT></TD>"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""WorkingHours"" ID=""WorkingHoursHdn"" VALUE=""" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & """ />"
						End If
					Response.Write "</TR>"
				End If
				If bFull Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Turno opcional:&nbsp;</FONT></TD>"
						If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
							Response.Write "<TD><SELECT NAME=""AdditionalShift"" ID=""AdditionalShiftCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
								Response.Write "<OPTION VALUE=""0"">No</OPTION>"
								Response.Write "<OPTION VALUE=""1"">Sí</OPTION>"
							Response.Write "</SELECT></TD>"
						Else
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
								If aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) = -1 Then
									Response.Write "Todos"
								Else
									Response.Write DisplayYesNo(aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT), True)
								End If
							Response.Write "</FONT></TD>"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AdditionalShift"" ID=""AdditionalShiftHdn"" VALUE=""" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & """ />"
						End If
					Response.Write "</TR>"
				End If
				If (InStr(1, ",0,1,2,4,5,6,", ("," & iSelectedTab & ","), vbBinaryCompare) > 0) Or bFull Then
					If aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) <> 1 Then
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Zona económica:&nbsp;</FONT></TD>"
							If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
								Response.Write "<TD><SELECT NAME=""EconomicZoneID"" ID=""EconomicZoneIDCmb"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write "<OPTION VALUE=""0"">Todas</OPTION>"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "EconomicZones", "EconomicZoneID", "EconomicZoneName", "(EconomicZoneID>0) And (Active=1)", "EconomicZoneName", aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT), "", sErrorDescription)
								Response.Write "</SELECT></TD>"
							Else
								If aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) = 0 Then
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Todas</FONT></TD>"
								Else
									Call GetNameFromTable(oADODBConnection, "EconomicZones", aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT), "", "", sNames, sErrorDescription)
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
								End If
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EconomicZoneID"" ID=""EconomicZoneIDHdn"" VALUE=""" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & """ />"
							End If
						Response.Write "</TR>"
					Else
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EconomicZoneID"" ID=""EconomicZoneIDHdn"" VALUE=""" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & """ />"
					End If
				End If
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Inicio de vigencia:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write DisplayDateCombosUsingSerial(aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT), "StartForValue", N_START_YEAR, Year(Date) + 1, True, False)
					Response.Write "</FONT></TD>"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDateForValueConcept"" ID=""StartDateForValueConceptHdn"" VALUE=""" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & """ />"
				Response.Write "</TR>"
				If bFull Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Servicio:&nbsp;</FONT></TD>"
						If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
							Response.Write "<TD><SELECT NAME=""ServiceID"" ID=""ServiceIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Services", "ServiceID", "ServiceShortName, ServiceName", "(ServiceID>0) And (Active=1)", "ServiceName", aConceptComponent(N_SERVICE_ID_CONCEPT), "", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Else
							If aConceptComponent(N_SERVICE_ID_CONCEPT) = -1 Then
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Todos</FONT></TD>"
							Else
								Call GetNameFromTable(oADODBConnection, "Services", aConceptComponent(N_SERVICE_ID_CONCEPT), "", "", sNames, sErrorDescription)
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							End If
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ServiceID"" ID=""ServiceIDHdn"" VALUE=""" & aConceptComponent(N_SERVICE_ID_CONCEPT) & """ />"
						End If
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Antigüedad en el ISSSTE:&nbsp;</FONT></TD>"
						If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
							Response.Write "<TD><SELECT NAME=""AntiquityID"" ID=""AntiquityIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE=""-1"">Todas</OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Antiquities", "AntiquityID", "AntiquityName", "(AntiquityID>-1)", "StartYears", aConceptComponent(N_ANTIQUITY_ID_CONCEPT), "", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Else
							If (aConceptComponent(N_ANTIQUITY_ID_CONCEPT) = -1) Or (aConceptComponent(N_ANTIQUITY_ID_CONCEPT) = 0) Then
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cualquiera</FONT></TD>"
							Else
								Call GetNameFromTable(oADODBConnection, "Antiquities", aConceptComponent(N_ANTIQUITY_ID_CONCEPT), "", "", sNames, sErrorDescription)
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							End If
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AntiquityID"" ID=""AntiquityIDHdn"" VALUE=""" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & """ />"
						End If
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Antigüedad consecutiva:&nbsp;</FONT></TD>"
						If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
							Response.Write "<TD><SELECT NAME=""Antiquity2ID"" ID=""Antiquity2IDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE=""-1"">Todas</OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Antiquities", "AntiquityID", "AntiquityName", "(AntiquityID>-1)", "StartYears", aConceptComponent(N_ANTIQUITY2_ID_CONCEPT), "", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Else
							If aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) = -1 Then
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cualquiera</FONT></TD>"
							Else
								Call GetNameFromTable(oADODBConnection, "Antiquities", aConceptComponent(N_ANTIQUITY2_ID_CONCEPT), "", "", sNames, sErrorDescription)
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							End If
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Antiquity2ID"" ID=""Antiquity2IDHdn"" VALUE=""" & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & """ />"
						End If
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Antigüedad en el ISSSTE con plaza de base:&nbsp;</FONT></TD>"
						If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
							Response.Write "<TD><SELECT NAME=""Antiquity3ID"" ID=""Antiquity3IDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE=""-1"">Todas</OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Antiquities", "AntiquityID", "AntiquityName", "(AntiquityID>-1)", "StartYears", aConceptComponent(N_ANTIQUITY3_ID_CONCEPT), "", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Else
							If aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) = -1 Then
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cualquiera</FONT></TD>"
							Else
								Call GetNameFromTable(oADODBConnection, "Antiquities", aConceptComponent(N_ANTIQUITY3_ID_CONCEPT), "", "", sNames, sErrorDescription)
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							End If
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Antiquity3ID"" ID=""Antiquity3IDHdn"" VALUE=""" & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & """ />"
						End If
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Antigüedad federal:&nbsp;</FONT></TD>"
						If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
							Response.Write "<TD><SELECT NAME=""Antiquity4ID"" ID=""Antiquity4IDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE=""-1"">Todas</OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Antiquities", "AntiquityID", "AntiquityName", "(AntiquityID>-1)", "StartYears", aConceptComponent(N_ANTIQUITY4_ID_CONCEPT), "", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Else
							If aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) = -1 Then
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cualquiera</FONT></TD>"
							Else
								Call GetNameFromTable(oADODBConnection, "Antiquities", aConceptComponent(N_ANTIQUITY4_ID_CONCEPT), "", "", sNames, sErrorDescription)
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							End If
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Antiquity4ID"" ID=""Antiquity4IDHdn"" VALUE=""" & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & """ />"
						End If
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Riesgos profesionales:&nbsp;</FONT></TD>"
						If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
							Response.Write "<TD><SELECT NAME=""ForRisk"" ID=""ForRiskCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
								Response.Write "<OPTION VALUE=""0"">No</OPTION>"
								Response.Write "<OPTION VALUE=""1"">Sí</OPTION>"
							Response.Write "</SELECT></TD>"
						Else
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
								If aConceptComponent(N_FOR_RISK_CONCEPT) = -1 Then
									Response.Write "Todos"
								Else
									Response.Write DisplayYesNo(aConceptComponent(N_FOR_RISK_CONCEPT), True)
								End If
							Response.Write "</FONT></TD>"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ForRisk"" ID=""ForRiskHdn"" VALUE=""" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & """ />"
						End If
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Género:&nbsp;</FONT></TD>"
						If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
							Response.Write "<TD><SELECT NAME=""GenderID"" ID=""GenderIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Genders", "GenderID", "GenderName", "", "GenderID", aConceptComponent(N_GENDER_ID_CONCEPT), "", sErrorDescription)
							Response.Write "</SELECT></TD>"
						Else
							If aConceptComponent(N_GENDER_ID_CONCEPT) = -1 Then
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Todos</FONT></TD>"
							Else
								Call GetNameFromTable(oADODBConnection, "Genders", aConceptComponent(N_GENDER_ID_CONCEPT), "", "", sNames, sErrorDescription)
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
							End If
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""GenderID"" ID=""GenderIDHdn"" VALUE=""" & aConceptComponent(N_GENDER_ID_CONCEPT) & """ />"
						End If
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Hijos:&nbsp;</FONT></TD>"
						If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
							Response.Write "<TD><SELECT NAME=""HasChildren"" ID=""HasChildrenCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
								Response.Write "<OPTION VALUE=""0"">No</OPTION>"
								Response.Write "<OPTION VALUE=""1"">Sí</OPTION>"
							Response.Write "</SELECT></TD>"
						Else
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
								If aConceptComponent(N_HAS_CHILDREN_CONCEPT) = -1 Then
									Response.Write "Todos"
								Else
									Response.Write DisplayYesNo(aConceptComponent(N_HAS_CHILDREN_CONCEPT), True)
								End If
							Response.Write "</FONT></TD>"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""HasChildren"" ID=""HasChildrenHdn"" VALUE=""" & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & """ />"
						End If
					Response.Write "</TR>"
					If (aConceptComponent(N_RECORD_ID_CONCEPT) = -1) Or (aConceptComponent(N_HAS_CHILDREN_CONCEPT) <> 0) Then
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Escolaridad (hijos):&nbsp;</FONT></TD>"
							If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
								Response.Write "<TD><SELECT NAME=""SchoolarshipID"" ID=""SchoolarshipIDCmb"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write "<OPTION VALUE=""-1"">Todas</OPTION>"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Schoolarships", "SchoolarshipID", "SchoolarshipName", "(SchoolarshipID In (1,2,3))", "SchoolarshipID", aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT), "", sErrorDescription)
								Response.Write "</SELECT></TD>"
							Else
								If aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) = -1 Then
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Todas</FONT></TD>"
								Else
									Call GetNameFromTable(oADODBConnection, "Schoolarships", aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT), "", "", sNames, sErrorDescription)
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
								End If
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SchoolarshipID"" ID=""SchoolarshipIDHdn"" VALUE=""" & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & """ />"
							End If
						Response.Write "</TR>"
					Else
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SchoolarshipID"" ID=""SchoolarshipIDHdn"" VALUE=""-1"" />"
					End If
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Sindicalizado:&nbsp;</FONT></TD>"
						If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
							Response.Write "<TD><SELECT NAME=""HasSyndicate"" ID=""HasSyndicateCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write "<OPTION VALUE=""-1"">Todos</OPTION>"
								Response.Write "<OPTION VALUE=""0"">No</OPTION>"
								Response.Write "<OPTION VALUE=""1"">Sí</OPTION>"
							Response.Write "</SELECT></TD>"
						Else
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
								If aConceptComponent(N_HAS_SYNDICATE_CONCEPT) = -1 Then
									Response.Write "Todos"
								Else
									Response.Write DisplayYesNo(aConceptComponent(N_HAS_SYNDICATE_CONCEPT), True)
								End If
							Response.Write "</FONT></TD>"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""HasSyndicate"" ID=""HasSyndicateHdn"" VALUE=""" & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & """ />"
						End If
					Response.Write "</TR>"
				End If
				Response.Write "<TR>"
					Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><NOBR>Monto del concepto (mensual):&nbsp;</NOBR></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aConceptComponent(D_CONCEPT_AMOUNT_CONCEPT), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
						If bFull Then
							Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowAmountFields(this.value);"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID Not In (23,24,25,33,34,35))", "QttyID", aConceptComponent(N_CONCEPT_QTTY_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
							Response.Write "<SPAN ID=""ConceptCurrencySpn""><SELECT NAME=""ConceptCurrencyID"" ID=""ConceptCurrencyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Currencies", "CurrencyID", "CurrencyName", "(CurrencyID>-1) And (Active=1)", "CurrencyName", aConceptComponent(N_CURRENCY_ID_CONCEPT), "Ninguna;;;-1", sErrorDescription)
							Response.Write "</SELECT></SPAN><BR />"
							Response.Write "<SPAN ID=""ConceptAppliesToSpn"" STYLE=""display: none""><SELECT NAME=""AppliesToID"" ID=""AppliesToIDCmb"" SIZE=""10"" MULTIPLE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID<>" & aConceptComponent(N_ID_CONCEPT) & ")", "ConceptShortName, ConceptName", aConceptComponent(S_APPLIES_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></SPAN>"
						Else
							Response.Write "<SELECT NAME=""ConceptQttyID"" ID=""ConceptQttyIDCmb"" SIZE=""1"" DISABLED CLASS=""Lists"" onChange=""ShowAmountFields(this.value);"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID Not In (23,24,25,33,34,35))", "QttyID", aConceptComponent(N_CONCEPT_QTTY_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
							Response.Write "<SPAN ID=""ConceptCurrencySpn""><SELECT NAME=""ConceptCurrencyID"" ID=""ConceptCurrencyIDCmb"" SIZE=""1"" DISABLED CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Currencies", "CurrencyID", "CurrencyName", "(CurrencyID>-1) And (Active=1)", "CurrencyName", aConceptComponent(N_CURRENCY_ID_CONCEPT), "Ninguna;;;-1", sErrorDescription)
							Response.Write "</SELECT></SPAN><BR />"
							Response.Write "<SPAN ID=""ConceptAppliesToSpn"" STYLE=""display: none""><SELECT NAME=""AppliesToID"" ID=""AppliesToIDCmb"" SIZE=""10"" MULTIPLE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Concepts", "ConceptID", "ConceptShortName, ConceptName", "(ConceptID<>" & aConceptComponent(N_ID_CONCEPT) & ")", "ConceptShortName, ConceptName", aConceptComponent(S_APPLIES_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT></SPAN>"
						End If
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				If bFull Then
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><NOBR>Monto mínimo:&nbsp;</NOBR></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptMin"" ID=""ConceptMinTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aConceptComponent(D_CONCEPT_MIN_CONCEPT), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
							Response.Write "<SELECT NAME=""ConceptMinQttyID"" ID=""ConceptMinQttyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID In (1,3,13,23,24,25,33,34,35))", "QttyID", aConceptComponent(N_CONCEPT_MIN_QTTY_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
					Response.Write "<TR>"
						Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><NOBR>Monto máximo:&nbsp;</NOBR></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""ConceptMax"" ID=""ConceptMaxTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aConceptComponent(D_CONCEPT_MAX_CONCEPT), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
							Response.Write "<SELECT NAME=""ConceptMaxQttyID"" ID=""ConceptMaxQttyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID In (1,3,13,23,24,25,33,34,35))", "QttyID", aConceptComponent(N_CONCEPT_MAX_QTTY_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
							Response.Write "</SELECT>"
						Response.Write "</FONT></TD>"
					Response.Write "</TR>"
				End If
				If False Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""ConceptTypeID"" ID=""ConceptTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "ConceptTypes", "ConceptTypeID", "ConceptTypeName", "", "ConceptTypeID", aConceptComponent(N_CONCEPT_TYPE_ID_CONCEPT), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"
				Else
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptTypeID"" ID=""ConceptTypeIDHdn"" VALUE=""" & aConceptComponent(N_CONCEPT_TYPE_ID_CONCEPT) & """ />"
				End If

			Response.Write "</TABLE><BR />"
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "ShowAmountFields(document.ConceptValuesFrm.ConceptQttyID.value);" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine

			If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
				Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveConceptWngDiv']); ConceptValuesFrm.Remove.focus()"" />"
			Else
				Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=" & oRequest("Action").Item & "&ConceptID=" & aConceptComponent(N_ID_CONCEPT) & "'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveConceptWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			'Response.Write "DisplayInfoMessage();" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayConceptValuesForm = lErrorNumber
	Err.Clear
End Function

Function DisplayConceptAsHiddenFields(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a concept using
'		  hidden form fields
'Inputs:  oRequest, oADODBConnection, aConceptComponent
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayConceptAsHiddenFields"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptID"" ID=""ConceptIDHdn"" VALUE=""" & aConceptComponent(N_ID_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StartDate"" ID=""StartDateHdn"" VALUE=""" & aConceptComponent(N_START_DATE_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EndDate"" ID=""EndDateHdn"" VALUE=""" & aConceptComponent(N_END_DATE_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptShortName"" ID=""ConceptShortNameHdn"" VALUE=""" & aConceptComponent(S_SHORT_NAME_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConceptName"" ID=""ConceptNameHdn"" VALUE=""" & aConceptComponent(S_NAME_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetID"" ID=""BudgetIDHdn"" VALUE=""" & aConceptComponent(N_BUDGET_ID_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PayrollTypeID"" ID=""PayrollTypeIDHdn"" VALUE=""" & aConceptComponent(N_PAYROLL_TYPE_ID_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PeriodID"" ID=""PeriodIDHdn"" VALUE=""" & aConceptComponent(N_PERIOD_ID_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IsDeduction"" ID=""IsDeductionHdn"" VALUE=""" & aConceptComponent(N_IS_DEDUCTION_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ForAlimony"" ID=""ForAlimonyHdn"" VALUE=""" & aConceptComponent(N_FOR_ALIMONY_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OnLeave"" ID=""OnLeaveHdn"" VALUE=""" & aConceptComponent(N_ON_LEAVE_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OrderInList"" ID=""OrderInListHdn"" VALUE=""" & aConceptComponent(N_ORDER_IN_LIST_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaxAmount"" ID=""TaxAmountHdn"" VALUE=""" & aConceptComponent(D_TAX_AMOUNT_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaxCurrencyID"" ID=""TaxCurrencyIDHdn"" VALUE=""" & aConceptComponent(N_TAX_CURRENCY_ID_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaxQttyID"" ID=""TaxQttyIDHdn"" VALUE=""" & aConceptComponent(N_TAX_QTTY_ID_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaxMin"" ID=""TaxMinHdn"" VALUE=""" & aConceptComponent(D_TAX_MIN_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaxMinQttyID"" ID=""TaxMinQttyIDHdn"" VALUE=""" & aConceptComponent(N_TAX_MIN_QTTY_ID_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaxMax"" ID=""TaxMaxHdn"" VALUE=""" & aConceptComponent(D_TAX_MAX_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TaxMaxQttyID"" ID=""TaxMaxQttyIDHdn"" VALUE=""" & aConceptComponent(N_TAX_MAX_QTTY_ID_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ExemptAmount"" ID=""ExemptAmountHdn"" VALUE=""" & aConceptComponent(D_EXEMPT_AMOUNT_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ExemptCurrencyID"" ID=""ExemptCurrencyIDHdn"" VALUE=""" & aConceptComponent(N_EXEMPT_CURRENCY_ID_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ExemptQttyID"" ID=""ExemptQttyIDHdn"" VALUE=""" & aConceptComponent(N_EXEMPT_QTTY_ID_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ExemptMin"" ID=""ExemptMinHdn"" VALUE=""" & aConceptComponent(D_EXEMPT_MIN_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ExemptMinQttyID"" ID=""ExemptMinQttyIDHdn"" VALUE=""" & aConceptComponent(N_EXEMPT_MIN_QTTY_ID_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ExemptMax"" ID=""ExemptMaxHdn"" VALUE=""" & aConceptComponent(D_EXEMPT_MAX_CONCEPT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ExemptMaxQttyID"" ID=""ExemptMaxQttyIDHdn"" VALUE=""" & aConceptComponent(N_EXEMPT_MAX_QTTY_ID_CONCEPT) & """ />"

	DisplayConceptAsHiddenFields = Err.number
	Err.Clear
End Function

Function PositionsSpecialJourneysLKPHasChanged(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a concept from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "PositionsSpecialJourneysLKPHasChanged"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From PositionsSpecialJourneysLKP Where (RecordID=" & aConceptComponent(N_RECORD_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
				oRecordset.Close
			Else
				If (aConceptComponent(N_IS_ACTIVE1) = CInt(oRecordset.Fields("IsActive1").Value)) And _
				(aConceptComponent(N_IS_ACTIVE2) = CInt(oRecordset.Fields("IsActive2").Value)) And _
				(aConceptComponent(N_IS_ACTIVE3) = CInt(oRecordset.Fields("IsActive3").Value)) And _
				(aConceptComponent(N_IS_ACTIVE4) = CInt(oRecordset.Fields("IsActive4").Value)) And _
				(aConceptComponent(N_START_DATE_CONCEPT) = CLng(oRecordset.Fields("StartDate").Value)) And _
				(aConceptComponent(N_END_DATE_CONCEPT) = CLng(oRecordset.Fields("EndDate").Value)) Then
					PositionsSpecialJourneysLKPHasChanged = False
				Else
					PositionsSpecialJourneysLKPHasChanged = True
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function RemoveConceptsValuesFile(oRequest, oADODBConnection, sQuery, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To remove concepts values for employee type into the database
'Inputs:  oRequest, oADODBConnection, sQuery
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveConceptsValuesFile"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información de la aplicación de tabuladores de forma masiva."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				aConceptComponent(N_RECORD_ID_CONCEPT) = CLng(oRecordset.Fields("RecordID").Value)
				lErrorNumber = RemoveConceptValues(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
		End If
	End If

	Set oRecordset = Nothing
	RemoveConceptsValuesFile = lErrorNumber
	Err.Clear
End Function

Function RemoveConceptValues(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given employee's concept
'Inputs:  oRequest, oADODBConnection
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveConceptValues"
	Dim lErrorNumber
	Dim sQuery

	If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador de la incidencia y/o la fecha para agregar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT) = True
		lErrorNumber = GetConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
		sQuery = "Update ConceptsValues Set EndDate=30000000, RegistrationEndDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & _
			" Where (EndDate=" & AddDaysToSerialDate(aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT), -1) & ")" & _
			" And (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ") And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (PositionTypeID=" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ") And (EmployeeStatusID=" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ") And (JobStatusID=" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ") And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (JourneyID=" & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (AdditionalShift=" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EconomicZoneID=" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (AntiquityID=" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ") And (Antiquity2ID=" & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ") And (Antiquity3ID=" & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ") And (Antiquity4ID=" & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ") And (ForRisk=" & aConceptComponent(N_FOR_RISK_CONCEPT) & ") And (GenderID=" & aConceptComponent(N_GENDER_ID_CONCEPT) & ") And (HasChildren=" & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ") And (SchoolarshipID=" & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ") And (HasSyndicate=" & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ") And (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ")"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
		sErrorDescription = "No se pudo modificar la información del tabulador de pago."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete from ConceptsValues Where (RecordID=" & aConceptComponent(N_RECORD_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	RemoveConceptValues = lErrorNumber
	Err.Clear
End Function

Function SetActiveForConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new concept into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForConcept"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim lNewRecordID
	Dim sSpecialCondition

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If aConceptComponent(N_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del concepto para guardar su valor."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		lErrorNumber = GetConcept(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			sSpecialCondition = "(ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (StartDate<>" & aConceptComponent(N_START_DATE_CONCEPT) & ") And"
			sQuery = "Select * From Concepts Where " & sSpecialCondition & " (ConceptShortName='" & aConceptComponent(S_SHORT_NAME_CONCEPT) & "') And (StartDate<" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate>" & aConceptComponent(N_END_DATE_CONCEPT) & ") Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Concepts Set EndDate=" & AddDaysToSerialDate(aConceptComponent(N_START_DATE_CONCEPT), -1) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (ConceptID=" & oRecordset.Fields("ConceptID").Value & ") And (StartDate=" & oRecordset.Fields("StartDate").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Concepts (ConceptID, StartDate, EndDate, ConceptShortName, ConceptName, BudgetID, PayrollTypeID, PeriodID, IsDeduction, ForAlimony, OnLeave, OrderInList, TaxAmount, TaxCurrencyID, TaxQttyID, TaxMin, TaxMinQttyID, TaxMax, TaxMaxQttyID, ExemptAmount, ExemptCurrencyID, ExemptQttyID, ExemptMin, ExemptMinQttyID, ExemptMax, ExemptMaxQttyID, StartUserID, EndUserID) Values (" & oRecordset.Fields("ConceptID").Value & ", " & AddDaysToSerialDate(aConceptComponent(N_END_DATE_CONCEPT), 1) & ", " & oRecordset.Fields("EndDate").Value & ", '" & Replace(oRecordset.Fields("ConceptShortName").Value, "'", "") & "', '" & Replace(oRecordset.Fields("ConceptName").Value, "'", "´") & "', " & oRecordset.Fields("BudgetID").Value & ", " & oRecordset.Fields("PayrollTypeID").Value & ", " & oRecordset.Fields("PeriodID").Value & ", " & oRecordset.Fields("IsDeduction").Value & ", " & oRecordset.Fields("ForAlimony").Value & ", " & oRecordset.Fields("OnLeave").Value & ", " & oRecordset.Fields("OrderInList").Value & ", " & oRecordset.Fields("TaxAmount").Value & ", " & oRecordset.Fields("TaxCurrencyID").Value & ", " & oRecordset.Fields("TaxQttyID").Value & ", " & oRecordset.Fields("TaxMin").Value & ", " & oRecordset.Fields("TaxMinQttyID").Value & ", " & oRecordset.Fields("TaxMax").Value & ", " & oRecordset.Fields("TaxMaxQttyID").Value & ", " & oRecordset.Fields("ExemptAmount").Value & ", " & oRecordset.Fields("ExemptCurrencyID").Value & ", " & oRecordset.Fields("ExemptQttyID").Value & ", " & oRecordset.Fields("ExemptMin").Value & ", " & oRecordset.Fields("ExemptMinQttyID").Value & ", " & oRecordset.Fields("ExemptMax").Value & ", " & oRecordset.Fields("ExemptMaxQttyID").Value & ", " & oRecordset.Fields("StartUserID").Value & ", " & oRecordset.Fields("StatusID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
			sQuery = "Select * From Concepts Where " & sSpecialCondition & " (ConceptShortName='" & aConceptComponent(S_SHORT_NAME_CONCEPT) & "') And (StartDate>" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate<" & aConceptComponent(N_END_DATE_CONCEPT) & ") Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Concepts Where " & sSpecialCondition & " (ConceptShortName='" & aConceptComponent(S_SHORT_NAME_CONCEPT) & "') And (StartDate>" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate<" & aConceptComponent(N_END_DATE_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				End If
			End If
			sQuery = "Select * From Concepts Where " & sSpecialCondition & " (ConceptShortName='" & aConceptComponent(S_SHORT_NAME_CONCEPT) & "') And (StartDate<" & aConceptComponent(N_START_DATE_CONCEPT) & ") And ((EndDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ") And (EndDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Concepts Set EndDate=" & AddDaysToSerialDate(aConceptComponent(N_START_DATE_CONCEPT), -1) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (ConceptID=" & oRecordset.Fields("ConceptID").Value & ") And (StartDate=" & oRecordset.Fields("StartDate").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
			sQuery = "Select * From Concepts Where " & sSpecialCondition & " (ConceptShortName='" & aConceptComponent(S_SHORT_NAME_CONCEPT) & "') And (StartDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ") And ((EndDate>=" & aConceptComponent(N_END_DATE_CONCEPT) & ") And (StartDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					If CLng(aConceptComponent(N_END_DATE_CONCEPT)) = CLng(oRecordset.Fields("EndDate").Value) Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete Concepts Where (ConceptID=" & oRecordset.Fields("ConceptID").Value & ") And (StartDate=" & oRecordset.Fields("StartDate").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Concepts Set StartDate=" & AddDaysToSerialDate(aConceptComponent(N_END_DATE_CONCEPT), 1) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (ConceptID=" & oRecordset.Fields("ConceptID").Value & ") And (StartDate=" & oRecordset.Fields("StartDate").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
			End If
			If aConceptComponent(N_IS_CREDIT) > 0 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From CreditTypes Where (CreditTypeID=" & aConceptComponent(N_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into CreditTypes (CreditTypeID, CreditTypeShortName, CreditTypeName, IsOther, Active) Values (" & aConceptComponent(N_ID_CONCEPT) & ", '" & Replace(aConceptComponent(S_SHORT_NAME_CONCEPT), "'", "") & "', '" & Replace(aConceptComponent(S_NAME_CONCEPT), "'", "´") & "', " & aConceptComponent(N_IS_OTHER) & ", 1)", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			sErrorDescription = "No se pudo aplicar la información del nuevo Concepto."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Concepts Set StatusID=1 Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (StartDate=" & aConceptComponent(N_START_DATE_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	SetActiveForConcept = lErrorNumber
	Err.Clear
End Function

Function SetActiveForConceptsValues(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new concept value into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForConceptsValues"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim lNewRecordID

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del puesto para guardar su valor."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		lErrorNumber = GetConceptValue(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			sQuery = "Select * From ConceptsValues Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ") And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (PositionTypeID=" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ") And (EmployeeStatusID=" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ") And (JobStatusID=" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ") And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (JourneyID=" & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (AdditionalShift=" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EconomicZoneID=" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (AntiquityID=" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ") And (Antiquity2ID=" & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ") And (Antiquity3ID=" & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ") And (Antiquity4ID=" & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ") And (ForRisk=" & aConceptComponent(N_FOR_RISK_CONCEPT) & ") And (GenderID=" & aConceptComponent(N_GENDER_ID_CONCEPT) & ") And (HasChildren=" & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ") And (SchoolarshipID=" & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ") And (HasSyndicate=" & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ") And (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (StatusID=1) And (StartDate<" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ") And (EndDate>" & aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & ") Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConceptsValues Set EndDate=" & AddDaysToSerialDate(aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT), -1) & " Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					lErrorNumber = GetNewIDFromTable(oADODBConnection, "ConceptsValues", "RecordID", "", 1, lNewRecordID, sErrorDescription)
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into ConceptsValues (RecordID, ConceptID, CompanyID, EmployeeTypeID, PositionTypeID, EmployeeStatusID, JobStatusID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, WorkingHours, AdditionalShift, LevelID, EconomicZoneID, ServiceID, AntiquityID, Antiquity2ID, Antiquity3ID, Antiquity4ID, ForRisk, GenderID, HasChildren, SchoolarshipID, HasSyndicate, StartDate, EndDate, RegistrationStartDate, AuthorizationDate, RegistrationEndDate, ConceptAmount, CurrencyID, ConceptQttyID, ConceptTypeID, AppliesToID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, PositionID, StartUserID, EndUserID, StatusID) Values (" & lNewRecordID & ", " & oRecordset.Fields("ConceptID").Value & ", " & oRecordset.Fields("CompanyID").Valuea & ", " & oRecordset.Fields("EmployeeTypeID").Value & ", " & oRecordset.Fields("PositionTypeID").Value & ", " & oRecordset.Fields("EmployeeStatusID").Value & ", " & oRecordset.Fields("JobStatusID").Value & ", " & oRecordset.Fields("ClassificationID").Value & ", " & oRecordset.Fields("GroupGradeLevelID").Value & ", " & oRecordset.Fields("IntegrationID").Value & ", " & oRecordset.Fields("JourneyID").Value & ", " & oRecordset.Fields("WorkingHours").Value & ", " & oRecordset.Fields("AdditionalShift").Value & ", " & oRecordset.Fields("LevelID").Value & ", " & oRecordset.Fields("EconomicZoneID").Value & ", " & oRecordset.Fields("ServiceID").Value & ", " & oRecordset.Fields("AntiquityID").Value & ", " & oRecordset.Fields("Antiquity2ID").Value & ", " & oRecordset.Fields("Antiquity3ID").Value & ", " & oRecordset.Fields("Antiquity4ID").Value & ", " & oRecordset.Fields("ForRisk").Value & ", " & oRecordset.Fields("GenderID").Value & ", " & oRecordset.Fields("HasChildren").Value & ", " & oRecordset.Fields("SchoolarshipID").Value & ", " & oRecordset.Fields("HasSyndicate").Value & ", " & AddDaysToSerialDate(aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT), 1) & ", " & oRecordset.Fields("EndDate").Value & ", " & oRecordset.Fields("RegistrationStartDate").Value & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", 0, " & oRecordset.Fields("ConceptAmount").Value & ", " & oRecordset.Fields("CurrencyID").Value & ", " & oRecordset.Fields("ConceptQttyID").Value & ", " & oRecordset.Fields("ConceptTypeID").Value & ", '" & oRecordset.Fields("AppliesToID").Value & "', " & oRecordset.Fields("ConceptMin").Value & ", " & oRecordset.Fields("ConceptMinQttyID").Value & ", " & oRecordset.Fields("ConceptMax").Value & ", " & oRecordset.Fields("ConceptMaxQttyID").Value & ", " & oRecordset.Fields("PositionID").Value & ", " & oRecordset.Fields("StartUserID").Value & ", " & aConceptComponent(N_END_USER_ID_CONCEPT) & ", " & oRecordset.Fields("StatusID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
			sQuery = "Select * From ConceptsValues Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ") And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (PositionTypeID=" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ") And (EmployeeStatusID=" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ") And (JobStatusID=" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ") And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (JourneyID=" & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (AdditionalShift=" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EconomicZoneID=" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (AntiquityID=" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ") And (Antiquity2ID=" & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ") And (Antiquity3ID=" & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ") And (Antiquity4ID=" & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ") And (ForRisk=" & aConceptComponent(N_FOR_RISK_CONCEPT) & ") And (GenderID=" & aConceptComponent(N_GENDER_ID_CONCEPT) & ") And (HasChildren=" & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ") And (SchoolarshipID=" & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ") And (HasSyndicate=" & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ") And (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (StatusID=1) And (StartDate>" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ") And (EndDate<" & aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & ") Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From ConceptsValues Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ") And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (PositionTypeID=" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ") And (EmployeeStatusID=" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ") And (JobStatusID=" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ") And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (JourneyID=" & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (AdditionalShift=" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EconomicZoneID=" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (AntiquityID=" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ")And (Antiquity2ID=" & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ") And (Antiquity3ID=" & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ") And (Antiquity4ID=" & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ") And (ForRisk=" & aConceptComponent(N_FOR_RISK_CONCEPT) & ") And (GenderID=" & aConceptComponent(N_GENDER_ID_CONCEPT) & ") And (HasChildren=" & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ") And (SchoolarshipID=" & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ") And (HasSyndicate=" & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ") And (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (StartDate>" & oRecordset.Fields("StartDate").Value & ") And (EndDate<" & aConceptComponent(N_END_DATE_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				End If
			End If
			sQuery = "Select * From ConceptsValues Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ") And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (PositionTypeID=" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ") And (EmployeeStatusID=" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ") And (JobStatusID=" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ") And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (JourneyID=" & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (AdditionalShift=" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EconomicZoneID=" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (AntiquityID=" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ") And (Antiquity2ID=" & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ") And (Antiquity3ID=" & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ") And (Antiquity4ID=" & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ") And (ForRisk=" & aConceptComponent(N_FOR_RISK_CONCEPT) & ") And (GenderID=" & aConceptComponent(N_GENDER_ID_CONCEPT) & ") And (HasChildren=" & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ") And (SchoolarshipID=" & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ") And (HasSyndicate=" & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ") And (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (StatusID=1) And (StartDate<" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ") And ((EndDate<=" & aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT) & ") And (EndDate>=" & aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConceptsValues Set EndDate=" & AddDaysToSerialDate(aConceptComponent(N_START_DATE_FOR_VALUE_CONCEPT), -1) & " Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
			sQuery = "Select * From ConceptsValues Where (ConceptID=" & aConceptComponent(N_ID_CONCEPT) & ") And (CompanyID=" & aConceptComponent(N_COMPANY_ID_CONCEPT) & ") And (EmployeeTypeID=" & aConceptComponent(N_EMPLOYEE_TYPE_ID_CONCEPT) & ") And (PositionTypeID=" & aConceptComponent(N_POSITION_TYPE_ID_CONCEPT) & ") And (EmployeeStatusID=" & aConceptComponent(N_EMPLOYEE_STATUS_ID_CONCEPT) & ") And (JobStatusID=" & aConceptComponent(N_JOB_STATUS_ID_CONCEPT) & ") And (ClassificationID=" & aConceptComponent(N_CLASSIFICATION_ID_CONCEPT) & ") And (GroupGradeLevelID=" & aConceptComponent(N_GROUP_GRADE_LEVEL_ID_CONCEPT) & ") And (IntegrationID=" & aConceptComponent(N_INTEGRATION_ID_CONCEPT) & ") And (JourneyID=" & aConceptComponent(N_JOURNEY_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (AdditionalShift=" & aConceptComponent(N_ADDITIONAL_SHIFT_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (EconomicZoneID=" & aConceptComponent(N_ECONOMIC_ZONE_ID_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (AntiquityID=" & aConceptComponent(N_ANTIQUITY_ID_CONCEPT) & ") And (Antiquity2ID=" & aConceptComponent(N_ANTIQUITY2_ID_CONCEPT) & ") And (Antiquity3ID=" & aConceptComponent(N_ANTIQUITY3_ID_CONCEPT) & ") And (Antiquity4ID=" & aConceptComponent(N_ANTIQUITY4_ID_CONCEPT) & ") And (ForRisk=" & aConceptComponent(N_FOR_RISK_CONCEPT) & ") And (GenderID=" & aConceptComponent(N_GENDER_ID_CONCEPT) & ") And (HasChildren=" & aConceptComponent(N_HAS_CHILDREN_CONCEPT) & ") And (SchoolarshipID=" & aConceptComponent(N_SCHOOLARSHIP_ID_CONCEPT) & ") And (HasSyndicate=" & aConceptComponent(N_HAS_SYNDICATE_CONCEPT) & ") And (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (StatusID=1) And (StartDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ") And ((EndDate>=" & aConceptComponent(N_END_DATE_CONCEPT) & ") And (StartDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					If CLng(aConceptComponent(N_END_DATE_CONCEPT)) = CLng(oRecordset.Fields("EndDate").Value) Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete ConceptsValues Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConceptsValues Set StartDate=" & AddDaysToSerialDate(aConceptComponent(N_END_DATE_FOR_VALUE_CONCEPT), 1) & " Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
			End If
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo guardar la información del nuevo registro."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConceptsValues Set StatusID=1, AuthorizationDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", EndUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " where (RecordID=" & aConceptComponent(N_RECORD_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If
	End If

	SetActiveForConceptsValues = lErrorNumber
	Err.Clear
End Function

Function SetActiveForPositionsSpecialJourneys(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new concept value into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aConceptComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForPositionsSpecialJourneys"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery
	Dim lNewRecordID
	Dim sSpecialCondition

	bComponentInitialized = aConceptComponent(B_COMPONENT_INITIALIZED_CONCEPT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeConceptComponent(oRequest, aConceptComponent)
	End If

	If aConceptComponent(N_RECORD_ID_CONCEPT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del puesto para guardar su valor."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "ConceptComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		lErrorNumber = GetPositionsSpecialJourneysLKP(oRequest, oADODBConnection, aConceptComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			sSpecialCondition = "(RecordID<>" & aConceptComponent(N_RECORD_ID_CONCEPT) & ") And"
			sQuery = "Select * From PositionsSpecialJourneysLKP Where " & sSpecialCondition & " (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (CenterTypeID=" & aConceptComponent(N_CENTER_TYPE_ID) & ") And (StartDate<" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate>" & aConceptComponent(N_END_DATE_CONCEPT) & ") Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PositionsSpecialJourneysLKP Set EndDate=" & AddDaysToSerialDate(aConceptComponent(N_START_DATE_CONCEPT), -1) & " Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					lErrorNumber = GetNewIDFromTable(oADODBConnection, "PositionsSpecialJourneysLKP", "RecordID", "", 1, lNewRecordID, sErrorDescription)
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PositionsSpecialJourneysLKP (RecordID, StartDate, EndDate, PositionID, LevelID, WorkingHours, ServiceID, CenterTypeID, IsActive1, IsActive2, IsActive3, IsActive4, Active) Values (" & lNewRecordID & ", " & AddDaysToSerialDate(aConceptComponent(N_END_DATE_CONCEPT), 1) & ", " & oRecordset.Fields("EndDate").Value & ", " & oRecordset.Fields("PositionID").Value & ", " & oRecordset.Fields("LevelID").Value & ", " & oRecordset.Fields("WorkingHours").Value & ", " & oRecordset.Fields("ServiceID").Value & ", " & oRecordset.Fields("CenterTypeID").Value & ", " & oRecordset.Fields("IsActive1").Value & ", " & oRecordset.Fields("IsActive2").Value & ", " & oRecordset.Fields("IsActive3").Value & ", " & oRecordset.Fields("IsActive4").Value & ", " & oRecordset.Fields("Active").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
			sQuery = "Select * From PositionsSpecialJourneysLKP Where " & sSpecialCondition & " (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (CenterTypeID=" & aConceptComponent(N_CENTER_TYPE_ID) & ") And (StartDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ") And (EndDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ") Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					Do While Not oRecordset.EOF
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PositionsSpecialJourneysLKP Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
				End If
			End If
			sQuery = "Select * From PositionsSpecialJourneysLKP Where " & sSpecialCondition & " (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (CenterTypeID=" & aConceptComponent(N_CENTER_TYPE_ID) & ") And (StartDate<" & aConceptComponent(N_START_DATE_CONCEPT) & ") And ((EndDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ") And (EndDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate Desc"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PositionsSpecialJourneysLKP Set EndDate=" & AddDaysToSerialDate(aConceptComponent(N_START_DATE_CONCEPT), -1) & " Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
			sQuery = "Select * From PositionsSpecialJourneysLKP Where " & sSpecialCondition & " (PositionID=" & aConceptComponent(N_POSITION_ID_CONCEPT) & ") And (LevelID=" & aConceptComponent(N_LEVEL_ID_CONCEPT) & ") And (WorkingHours=" & aConceptComponent(D_WORKING_HOURS_CONCEPT) & ") And (ServiceID=" & aConceptComponent(N_SERVICE_ID_CONCEPT) & ") And (CenterTypeID=" & aConceptComponent(N_CENTER_TYPE_ID) & ") And (StartDate>=" & aConceptComponent(N_START_DATE_CONCEPT) & ") And ((EndDate>" & aConceptComponent(N_END_DATE_CONCEPT) & ") And (StartDate<=" & aConceptComponent(N_END_DATE_CONCEPT) & ")) And (StartDate<EndDate) Order By StartDate"
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PositionsSpecialJourneysLKP Set StartDate=" & AddDaysToSerialDate(aConceptComponent(N_END_DATE_CONCEPT), 1) & " Where (RecordID=" & oRecordset.Fields("RecordID").Value & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
			sErrorDescription = "No se pudo aplicar la información del nuevo registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PositionsSpecialJourneysLKP Set Active=1 Where (RecordID=" & aConceptComponent(N_RECORD_ID_CONCEPT) & ")", "ConceptComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	SetActiveForPositionsSpecialJourneys = lErrorNumber
	Err.Clear
End Function

Function VerifyExistenceOfCatalogInDate(oADODBConnection, sTableName, sIDField, sValueField, sAddCondition, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To verify if a record exist in determinated database table
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistenceOfCatalogInDate"
	Dim lErrorNumber
	Dim sQuery
	Dim asField
	Dim asValue
	Dim iIndex
	Dim sQueryFieldCondition

	asField = Split(sIDField, ",", -1, vbBinaryCompare)
	asValue = Split(sValueField, ",", -1, vbBinaryCompare)

	If UBound(asField) = UBound(asValue) Then
		sQueryFieldCondition = "(" & asField(0) & "='" & asValue(0) & "')"
		sQueryFieldCondition = sQueryFieldCondition & _ 
							   " And (((" & asField(1) & ">=" & asValue(1) & ") And (" & asField(1) & "<=" & asValue(2) & "))" & _
							   " Or ((" & asField(2) & ">=" & asValue(1) & ") And (" & asField(2) & "<=" & asValue(2) & "))" & _
							   " Or ((" & asField(2) & ">=" & asValue(1) & ") And (" & asField(1) & "<=" & asValue(2) & ")))"
	Else
		lErrorNumber = -1
		sErrorDescription = "No se pudo verificar si el registro esta activo"
	End If

	sQuery = "Select * From " & sTableName & " Where " & sQueryFieldCondition & sAddCondition

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "QueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	If lErrorNumber = 0 Then
		VerifyExistenceOfCatalogInDate = (Not oRecordset.EOF)
	Else
		sErrorDescription = "Error al verificar si el registro esta activo."
		VerifyExistenceOfCatalogInDate = False
	End If
	Err.Clear
End Function
%>