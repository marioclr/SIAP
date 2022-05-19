<!-- #include file="ExternalSpecialJourneyComponent.asp" -->
<%
Const N_SPECIAL_JOURNEY_ID = 0
Const N_SPECIAL_JOURNEY_EMPLOYEE_ID = 1
Const S_SPECIAL_JOURNEY_EMPLOYEE_NUMBER = 2
Const S_SPECIAL_JOURNEY_EMPLOYEE_NAME = 3
Const S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME = 4
Const S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME2 = 5
Const S_SPECIAL_JOURNEY_RFC = 6
Const S_SPECIAL_JOURNEY_CURP = 7
Const N_SPECIAL_JOURNEY_ORIGINAL_EMPLOYEE_ID = 8
Const N_SPECIAL_JOURNEY_POSITION_ID = 9
Const N_SPECIAL_JOURNEY_AREA_ID = 10
Const N_SPECIAL_JOURNEY_SERVICE_ID = 11
Const N_SPECIAL_JOURNEY_LEVEL_ID = 12
Const N_SPECIAL_JOURNEY_WORKING_HOURS = 13
Const N_SPECIAL_JOURNEY_SHIFT_ID = 14
Const N_SPECIAL_JOURNEY_RISKLEVEL_ID = 15
Const N_SPECIAL_JOURNEY_SPECIAL_JOURNEY_ID = 16
Const S_SPECIAL_JOURNEY_DOCUMENTNUMBER = 17
Const N_SPECIAL_JOURNEY_STARTDATE = 18
Const N_SPECIAL_JOURNEY_ENDDATE = 19
Const N_SPECIAL_JOURNEY_STARTHOUR = 20
Const N_SPECIAL_JOURNEY_ENDHOUR = 21
Const N_SPECIAL_JOURNEY_JOURNEY_ID = 22
Const D_SPECIAL_JOURNEY_WORKED_HOURS = 23
Const N_SPECIAL_JOURNEY_MOVEMENT_ID = 24
Const N_SPECIAL_JOURNEY_FACTOR_ID = 25
Const N_SPECIAL_JOURNEY_REASON_ID = 26
Const S_SPECIAL_JOURNEY_COMMENTS = 27
Const D_SPECIAL_JOURNEY_CONCEPT_AMOUNT = 28

Const D_SPECIAL_JOURNEY_TAX_AMOUNT = 29
Const D_SPECIAL_JOURNEY_NET_AMOUNT = 30
Const N_SPECIAL_JOURNEY_BENEFICIARY_ID1 = 31
Const D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT1 = 32
Const N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID1 = 33
Const N_SPECIAL_JOURNEY_BENEFICIARY_ID2 = 34
Const D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT2 = 35
Const N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID2 = 36
Const N_SPECIAL_JOURNEY_BENEFICIARY_ID3 = 37
Const D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT3 = 38
Const N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID3 = 39

Const N_SPECIAL_JOURNEY_ADD_USERID = 40
Const N_SPECIAL_JOURNEY_ADD_DATE = 41
Const N_SPECIAL_JOURNEY_APPLIED_DATE = 42
Const N_SPECIAL_JOURNEY_REMOVED = 43
Const N_SPECIAL_JOURNEY_REMOVE_USER_ID = 44
Const N_SPECIAL_JOURNEY_REMOVED_DATE = 45
Const N_SPECIAL_JOURNEY_APPLIED_REMOVE_DATE = 46
Const N_SPECIAL_JOURNEY_ACTIVE = 47
Const B_SPECIAL_JOURNEY_EXIST_EXTERNAL = 48
Const N_SPECIAL_JOURNEY_EXTERNAL_ID = 49
Const S_SPECIAL_JOURNEY_SPEP = 50
Const S_SPECIAL_JOURNEY_FOLIO = 51

Const B_IS_DUPLICATED_SPECIAL_JOURNEY = 52
Const B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY = 53
Const S_QUERY_CONDITION_SPECIAL_JOURNEY = 54

Const N_SPECIAL_JOURNEY_COMPONENT_SIZE = 54

Dim aSpecialJourneyComponent()
Redim aSpecialJourneyComponent(N_SPECIAL_JOURNEY_COMPONENT_SIZE)

Function InitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Special Journey Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aSpecialJourneyComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeSpecialJourneyComponent"
	Dim oItem
	Redim Preserve aSpecialJourneyComponent(N_SPECIAL_JOURNEY_COMPONENT_SIZE)

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ID)) Then
		If Len(oRequest("RecordID").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ID) = CLng(oRequest("RecordID").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ID) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID)) Then
		If Len(oRequest("EmployeeID").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) = CLng(oRequest("EmployeeID").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NUMBER)) Then
		If Len(oRequest("EmployeeNumber").Item) > 0 Then
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NUMBER) = oRequest("EmployeeNumber").Item
		Else
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NUMBER) = ""
		End If
	End If
	aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NUMBER) = Left(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NUMBER), 6)

	If IsEmpty(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NAME)) Then
		If Len(oRequest("EmployeeName").Item) > 0 Then
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NAME) = oRequest("EmployeeName").Item
		Else
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NAME) = ""
		End If
	End If
	aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NAME) = Left(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NAME), 100)

	If IsEmpty(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME)) Then
		If Len(oRequest("EmployeeLastName").Item) > 0 Then
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME) = oRequest("EmployeeLastName").Item
		Else
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME) = ""
		End If
	End If
	aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME) = Left(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME), 100)

	If IsEmpty(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME2)) Then
		If Len(oRequest("EmployeeLastName2").Item) > 0 Then
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME2) = oRequest("EmployeeLastName2").Item
		Else
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME2) = ""
		End If
	End If
	aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME2) = Left(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME2), 100)

	If IsEmpty(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_RFC)) Then
		If Len(oRequest("RFC").Item) > 0 Then
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_RFC) = Left(oRequest("RFC").Item, 13)
		Else
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_RFC) = ""
		End If
	End If
	aSpecialJourneyComponent(S_SPECIAL_JOURNEY_RFC) = Left(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_RFC), 100)

	If IsEmpty(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_CURP)) Then
		If Len(oRequest("CURP").Item) > 0 Then
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_CURP) = Left(oRequest("CURP").Item, 18)
		Else
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_CURP) = ""
		End If
	End If
	aSpecialJourneyComponent(S_SPECIAL_JOURNEY_CURP) = Left(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_CURP), 100)

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ORIGINAL_EMPLOYEE_ID)) Then
		If Len(oRequest("OriginalEmployeeID").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ORIGINAL_EMPLOYEE_ID) = CLng(oRequest("OriginalEmployeeID").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ORIGINAL_EMPLOYEE_ID) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_POSITION_ID)) Then
		If Len(oRequest("PositionID").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_POSITION_ID) = CLng(oRequest("PositionID").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_POSITION_ID) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_AREA_ID)) Then
		If Len(oRequest("AreaID").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_AREA_ID) = CLng(oRequest("AreaID").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_AREA_ID) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SERVICE_ID)) Then
		If Len(oRequest("ServiceID").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SERVICE_ID) = CLng(oRequest("ServiceID").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SERVICE_ID) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_LEVEL_ID)) Then
		If Len(oRequest("LevelID").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_LEVEL_ID) = CLng(oRequest("LevelID").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_LEVEL_ID) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_WORKING_HOURS)) Then
		If Len(oRequest("WorkingHours").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_WORKING_HOURS) = CDbl(oRequest("WorkingHours").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_WORKING_HOURS) = 0
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SHIFT_ID)) Then
		If Len(oRequest("ShiftID").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SHIFT_ID) = CLng(oRequest("ShiftID").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SHIFT_ID) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_RISKLEVEL_ID)) Then
		If Len(oRequest("RiskLevelID").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_RISKLEVEL_ID) = CLng(oRequest("RiskLevelID").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_RISKLEVEL_ID) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SPECIAL_JOURNEY_ID)) Then
		If Len(oRequest("SpecialJourneyID").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SPECIAL_JOURNEY_ID) = CLng(oRequest("SpecialJourneyID").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SPECIAL_JOURNEY_ID) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_DOCUMENTNUMBER)) Then
		If Len(oRequest("DocumentNumber").Item) > 0 Then
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_DOCUMENTNUMBER) = oRequest("DocumentNumber").Item
		Else
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_DOCUMENTNUMBER) = ""
		End If
	End If
	aSpecialJourneyComponent(S_SPECIAL_JOURNEY_DOCUMENTNUMBER) = Left(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_DOCUMENTNUMBER), 50)

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE)) Then
		If Len(oRequest("StartDateYear").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) = CInt(oRequest("StartDateYear").Item) & Right(("0" & oRequest("StartDateMonth").Item), Len("00")) & Right(("0" & oRequest("StartDateDay").Item), Len("00"))
		ElseIf Len(oRequest("StartDate").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) = CLng(oRequest("StartDate").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE)) Then
		If Len(oRequest("EndDateYear").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE) = CInt(oRequest("EndDateYear").Item) & Right(("0" & oRequest("EndDateMonth").Item), Len("00")) & Right(("0" & oRequest("EndDateDay").Item), Len("00"))
		ElseIf Len(oRequest("EndDate").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE) = CLng(oRequest("EndDate").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE) = aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE)
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTHOUR)) Then
		If Len(oRequest("StartHour").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTHOUR) = CInt(oRequest("StartHour").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTHOUR) = 0
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDHOUR)) Then
		If Len(oRequest("EndHour").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDHOUR) = CInt(oRequest("EndHour").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDHOUR) = 0
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_JOURNEY_ID)) Then
		If Len(oRequest("JourneyID").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_JOURNEY_ID) = CLng(oRequest("JourneyID").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_JOURNEY_ID) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(D_SPECIAL_JOURNEY_WORKED_HOURS)) Then
		If Len(oRequest("WorkedHours").Item) > 0 Then
			aSpecialJourneyComponent(D_SPECIAL_JOURNEY_WORKED_HOURS) = CDbl(oRequest("WorkedHours").Item)
		Else
			aSpecialJourneyComponent(D_SPECIAL_JOURNEY_WORKED_HOURS) = 0
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_MOVEMENT_ID)) Then
		If Len(oRequest("MovementID").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_MOVEMENT_ID) = CLng(oRequest("MovementID").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_MOVEMENT_ID) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_FACTOR_ID)) Then
		If Len(oRequest("FactorID").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_FACTOR_ID) = CLng(oRequest("FactorID").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_FACTOR_ID) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REASON_ID)) Then
		If Len(oRequest("ReasonID").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REASON_ID) = CLng(oRequest("ReasonID").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REASON_ID) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_COMMENTS)) Then
		If Len(oRequest("Comments").Item) > 0 Then
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_COMMENTS) = oRequest("Comments").Item
		Else
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_COMMENTS) = ""
		End If
	End If
	aSpecialJourneyComponent(S_SPECIAL_JOURNEY_COMMENTS) = Left(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_COMMENTS), 2000)

	If IsEmpty(aSpecialJourneyComponent(D_SPECIAL_JOURNEY_CONCEPT_AMOUNT)) Then
		If Len(oRequest("ConceptAmount").Item) > 0 Then
			aSpecialJourneyComponent(D_SPECIAL_JOURNEY_CONCEPT_AMOUNT) = CDbl(oRequest("ConceptAmount").Item)
		Else
			aSpecialJourneyComponent(D_SPECIAL_JOURNEY_CONCEPT_AMOUNT) = 0
		End If
	End If


	If IsEmpty(aSpecialJourneyComponent(D_SPECIAL_JOURNEY_TAX_AMOUNT)) Then
		If Len(oRequest("TaxAmount").Item) > 0 Then
			aSpecialJourneyComponent(D_SPECIAL_JOURNEY_TAX_AMOUNT) = CDbl(oRequest("TaxAmount").Item)
		Else
			aSpecialJourneyComponent(D_SPECIAL_JOURNEY_TAX_AMOUNT) = 0
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(D_SPECIAL_JOURNEY_NET_AMOUNT)) Then
		If Len(oRequest("NetAmount").Item) > 0 Then
			aSpecialJourneyComponent(D_SPECIAL_JOURNEY_NET_AMOUNT) = CDbl(oRequest("NetAmount").Item)
		Else
			aSpecialJourneyComponent(D_SPECIAL_JOURNEY_NET_AMOUNT) = 0
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_ID1)) Then
		If Len(oRequest("BeneficiaryID1").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_ID1) = CLng(oRequest("BeneficiaryID1").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_ID1) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT1)) Then
		If Len(oRequest("BeneficiaryAmount1").Item) > 0 Then
			aSpecialJourneyComponent(D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT1) = CDbl(oRequest("BeneficiaryAmount1").Item)
		Else
			aSpecialJourneyComponent(D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT1) = 0
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID1)) Then
		If Len(oRequest("BeneficiaryAreaID1").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID1) = CLng(oRequest("BeneficiaryAreaID1").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID1) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_ID2)) Then
		If Len(oRequest("BeneficiaryID2").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_ID2) = CLng(oRequest("BeneficiaryID2").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_ID2) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT2)) Then
		If Len(oRequest("BeneficiaryAmount2").Item) > 0 Then
			aSpecialJourneyComponent(D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT2) = CDbl(oRequest("BeneficiaryAmount2").Item)
		Else
			aSpecialJourneyComponent(D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT2) = 0
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID2)) Then
		If Len(oRequest("BeneficiaryAreaID2").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID2) = CLng(oRequest("BeneficiaryAreaID2").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID2) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_ID3)) Then
		If Len(oRequest("BeneficiaryID3").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_ID3) = CLng(oRequest("BeneficiaryID3").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_ID3) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT3)) Then
		If Len(oRequest("BeneficiaryAmount3").Item) > 0 Then
			aSpecialJourneyComponent(D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT3) = CDbl(oRequest("BeneficiaryAmount3").Item)
		Else
			aSpecialJourneyComponent(D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT3) = 0
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID3)) Then
		If Len(oRequest("BeneficiaryAreaID3").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID3) = CLng(oRequest("BeneficiaryAreaID3").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID3) = -1
		End If
	End If


	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_USERID)) Then
		If Len(oRequest("AddUserID").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_USERID) = CLng(oRequest("AddUserID").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_USERID) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_DATE)) Then
		If Len(oRequest("AddYear").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_DATE) = CInt(oRequest("AddYear").Item) & Right(("0" & oRequest("AddMonth").Item), Len("00")) & Right(("0" & oRequest("AddDay").Item), Len("00"))
		ElseIf Len(oRequest("AddDate").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_DATE) = CLng(oRequest("AddDate").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_DATE) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE)) Then
		If Len(oRequest("AppliedYear").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE) = CInt(oRequest("AppliedYear").Item) & Right(("0" & oRequest("AppliedMonth").Item), Len("00")) & Right(("0" & oRequest("AppliedDay").Item), Len("00"))
		ElseIf Len(oRequest("AppliedDate").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE) = CLng(oRequest("AppliedDate").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE) = 0
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED)) Then
		If Len(oRequest("Removed").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED) = CLng(oRequest("Removed").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVE_USER_ID)) Then
		If Len(oRequest("RemoveUserID").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVE_USER_ID) = CLng(oRequest("RemoveUserID").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVE_USER_ID) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED_DATE)) Then
		If Len(oRequest("RemovedYear").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED_DATE) = CInt(oRequest("RemovedYear").Item) & Right(("0" & oRequest("RemovedMonth").Item), Len("00")) & Right(("0" & oRequest("RemovedDay").Item), Len("00"))
		ElseIf Len(oRequest("RemovedDate").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED_DATE) = CLng(oRequest("RemovedDate").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED_DATE) = 0
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_REMOVE_DATE)) Then
		If Len(oRequest("AppliedRemoveYear").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_REMOVE_DATE) = CInt(oRequest("AppliedRemoveYear").Item) & Right(("0" & oRequest("AppliedRemoveMonth").Item), Len("00")) & Right(("0" & oRequest("AppliedRemoveDay").Item), Len("00"))
		ElseIf Len(oRequest("AppliedRemoveDate").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_REMOVE_DATE) = CLng(oRequest("AppliedRemoveDate").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_REMOVE_DATE) = 0
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE)) Then
		If Len(oRequest("Active").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE) = CLng(oRequest("Active").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE) = 0
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EXTERNAL_ID)) Then
		If Len(oRequest("ExternalID").Item) > 0 Then
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EXTERNAL_ID) = CLng(oRequest("ExternalID").Item)
		Else
			aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EXTERNAL_ID) = -1
		End If
	End If

	If IsEmpty(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_SPEP)) Then
		If Len(oRequest("SPEP").Item) > 0 Then
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_SPEP) = Left(oRequest("SPEP").Item, 15)
		Else
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_SPEP) = ""
		End If
	End If
	aSpecialJourneyComponent(S_SPECIAL_JOURNEY_SPEP) = Left(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_SPEP), 15)

	If IsEmpty(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_FOLIO)) Then
		If Len(oRequest("Folio").Item) > 0 Then
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_FOLIO) = CStr(oRequest("Folio").Item)
		Else
			aSpecialJourneyComponent(S_SPECIAL_JOURNEY_FOLIO) = ""
		End If
	End If
	aSpecialJourneyComponent(S_SPECIAL_JOURNEY_FOLIO) = Left(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_FOLIO), 15)


	aSpecialJourneyComponent(B_SPECIAL_JOURNEY_EXIST_EXTERNAL) = False 'Len(oRequest("Add").Item) > 0
	aSpecialJourneyComponent(B_IS_DUPLICATED_SPECIAL_JOURNEY) = False
	aSpecialJourneyComponent(B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY) = True
	aSpecialJourneyComponent(S_QUERY_CONDITION_SPECIAL_JOURNEY) = ""
	InitializeSpecialJourneyComponent = Err.number
	Err.Clear
End Function

Function AddSpecialJourney(oRequest, oADODBConnection, iSpecialJourneyType, aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new Special Journey for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddSpecialJourney"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim lDate
	Dim bIsForPeriod
	Dim iJourneyTypeID
	Dim iRedordID

	bIsForPeriod = True
	lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))

	bComponentInitialized = aSpecialJourneyComponent(B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
	End If

	If (aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) = -1) Or (aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) = -1) Or (aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador de la incidencia y/o la fecha para agregar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If lErrorNumber = 0 Then
			If aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE) < aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) Then
				lErrorNumber = -1
				sErrorDescription = "La fecha de fin (" & DisplayDateFromSerialNumber(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE), -1, -1, -1) & ") no debe de ser menor a la fecha de inicio (" & DisplayDateFromSerialNumber(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE), -1, -1, -1) & ")"
			End If
		End If
		If lErrorNumber = 0 Then
			If iSpecialJourneyType = 1 Then ' Internos
				If VerifyRequerimentsForEmployeesSpecialJourneys(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
					aSpecialJourneyComponent(B_IS_DUPLICATED_SPECIAL_JOURNEY) = False
					lErrorNumber = CheckExistencyOfSpecialJourney(aSpecialJourneyComponent, bIsForPeriod, sErrorDescription)
					If lErrorNumber = 0 Then
						If aSpecialJourneyComponent(B_IS_DUPLICATED_SPECIAL_JOURNEY) Then
							lErrorNumber = L_ERR_DUPLICATED_RECORD
							sErrorDescription = "Ya existe un registro con fecha de inicio " & DisplayDateFromSerialNumber(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE), -1, -1, -1) & " para el empleado " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID)
							Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
						Else
							If Not CheckSpecialJourneyInformationConsistency(aSpecialJourneyComponent, sErrorDescription) Then
								lErrorNumber = -1
							Else
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
									lErrorNumber = GetNewIDFromTable(oADODBConnection, "EmployeesSpecialJourneys", "RecordID", "", 1, iRedordID, sErrorDescription)
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudo guardar la información del registro."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesSpecialJourneys (RecordID, EmployeeID, OriginalEmployeeID, PositionID, AreaID, ServiceID, LevelID, WorkingHours, ShiftID, RiskLevelID, SpecialJourneyID, DocumentNumber, StartDate, EndDate, StartHour, EndHour, JourneyID, WorkedHours, MovementID, FactorID, ReasonID, Comments, ConceptAmount, TaxAmount, NetAmount, BeneficiaryID1, BeneficiaryAmount1, BeneficiaryAreaID1, BeneficiaryID2, BeneficiaryAmount2, BeneficiaryAreaID2, BeneficiaryID3, BeneficiaryAmount3, BeneficiaryAreaID3, AddUserID, AddDate, AppliedDate, Removed, RemoveUserID, RemovedDate, AppliedRemoveDate, Active) Values (" & iRedordID & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) & ", " & _
															aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ORIGINAL_EMPLOYEE_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_POSITION_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_AREA_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SERVICE_ID) & ", " & _
															aSpecialJourneyComponent(N_SPECIAL_JOURNEY_LEVEL_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_WORKING_HOURS) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SHIFT_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_RISKLEVEL_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SPECIAL_JOURNEY_ID) & ", '" & Replace(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_DOCUMENTNUMBER), "'", "´") & "', " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTHOUR) & ", " &  aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDHOUR) & ", " &  aSpecialJourneyComponent(N_SPECIAL_JOURNEY_JOURNEY_ID) & ", " & _
															aSpecialJourneyComponent(D_SPECIAL_JOURNEY_WORKED_HOURS) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_MOVEMENT_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_FACTOR_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REASON_ID) & ", '" & Replace(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_COMMENTS), "'", "´") & "', " & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_CONCEPT_AMOUNT) & ", " & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_TAX_AMOUNT) & ", " & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_NET_AMOUNT) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_ID1) & ", " & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT1) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID1) & ", " & _
															aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_ID2) & ", " & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT2) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID2) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_ID3) & ", " & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT3) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID3) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_USERID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_DATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVE_USER_ID) & ", " & _
															aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED_DATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_REMOVE_DATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE) & ")", "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									End If
								End If
							End If
						End If
					End If
				Else
					lErrorNumber = -1
				End If
			ElseIf iSpecialJourneyType = 2 Then ' Externos
				If VerifyRequerimentsForExternalSpecialJourneys(oADODBConnection, aSpecialJourneyComponent, sErrorDescription) Then
					aSpecialJourneyComponent(B_IS_DUPLICATED_SPECIAL_JOURNEY) = False
					lErrorNumber = CheckExistencyOfSpecialJourney(aSpecialJourneyComponent, bIsForPeriod, sErrorDescription)
					If lErrorNumber = 0 Then
						If aSpecialJourneyComponent(B_IS_DUPLICATED_SPECIAL_JOURNEY) Then
							lErrorNumber = L_ERR_DUPLICATED_RECORD
							sErrorDescription = "Ya existe un registro el día " & DisplayDateFromSerialNumber(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE), -1, -1, -1) & " para el empleado " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID)
							Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
						Else
							If Not CheckSpecialJourneyInformationConsistency(aSpecialJourneyComponent, sErrorDescription) Then
								lErrorNumber = -1
							Else
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
									lErrorNumber = GetNewIDFromTable(oADODBConnection, "EmployeesSpecialJourneys", "RecordID", "", 1, iRedordID, sErrorDescription)
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudo guardar la información del registro."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesSpecialJourneys (RecordID, EmployeeID, OriginalEmployeeID, PositionID, AreaID, ServiceID, LevelID, WorkingHours, ShiftID, RiskLevelID, SpecialJourneyID, DocumentNumber, StartDate, EndDate, StartHour, EndHour, JourneyID, WorkedHours, MovementID, FactorID, ReasonID, Comments, ConceptAmount, TaxAmount, NetAmount, BeneficiaryID1, BeneficiaryAmount1, BeneficiaryAreaID1, BeneficiaryID2, BeneficiaryAmount2, BeneficiaryAreaID2, BeneficiaryID3, BeneficiaryAmount3, BeneficiaryAreaID3, AddUserID, AddDate, AppliedDate, Removed, RemoveUserID, RemovedDate, AppliedRemoveDate, Active) Values (" & iRedordID & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) & ", " & _
															aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ORIGINAL_EMPLOYEE_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_POSITION_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_AREA_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SERVICE_ID) & ", " & _
															aSpecialJourneyComponent(N_SPECIAL_JOURNEY_LEVEL_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_WORKING_HOURS) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SHIFT_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_RISKLEVEL_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SPECIAL_JOURNEY_ID) & ", '" & Replace(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_DOCUMENTNUMBER), "'", "´") & "', " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTHOUR) & ", " &  aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDHOUR) & ", " &  aSpecialJourneyComponent(N_SPECIAL_JOURNEY_JOURNEY_ID) & ", " & _
															aSpecialJourneyComponent(D_SPECIAL_JOURNEY_WORKED_HOURS) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_MOVEMENT_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_FACTOR_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REASON_ID) & ", '" & Replace(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_COMMENTS), "'", "´") & "', " & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_CONCEPT_AMOUNT) & ", " & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_TAX_AMOUNT) & ", " & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_NET_AMOUNT) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_ID1) & ", " & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT1) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID1) & ", " & _
															aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_ID2) & ", " & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT2) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID2) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_ID3) & ", " & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_BENEFICIARY_AMOUNT3) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_BENEFICIARY_AREA_ID3) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_USERID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_DATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVE_USER_ID) & ", " & _
															aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED_DATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_REMOVE_DATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE) & ")", "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
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
	End If

	AddSpecialJourney = lErrorNumber
	Err.Clear
End Function

Function AddSpecialJourney_SP(oRequest, oADODBConnection, iSpecialJourneyType, aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new Special Journey for the employee into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddSpecialJourney_SP"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim lDate
	Dim bIsForPeriod
	Dim iJourneyTypeID
	Dim iRedordID

	bIsForPeriod = True
	lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))

	bComponentInitialized = aSpecialJourneyComponent(B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
	End If

	If (aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) = -1) Or (aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) = -1) Or (aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador de la incidencia y/o la fecha para agregar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If lErrorNumber = 0 Then
			If aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE) < aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) Then
				lErrorNumber = -1
				sErrorDescription = "La fecha de fin (" & DisplayDateFromSerialNumber(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE), -1, -1, -1) & ") no debe de ser menor a la fecha de inicio (" & DisplayDateFromSerialNumber(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE), -1, -1, -1) & ")"
			End If
		End If
		If lErrorNumber = 0 Then
			If iSpecialJourneyType = 1 Then ' Internos
				If VerifyRequerimentsForEmployeesSpecialJourneys(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
					aSpecialJourneyComponent(B_IS_DUPLICATED_SPECIAL_JOURNEY) = False
					lErrorNumber = CheckExistencyOfSpecialJourney(aSpecialJourneyComponent, bIsForPeriod, sErrorDescription)
					If lErrorNumber = 0 Then
						If aSpecialJourneyComponent(B_IS_DUPLICATED_SPECIAL_JOURNEY) Then
							lErrorNumber = L_ERR_DUPLICATED_RECORD
							sErrorDescription = "Ya existe un registro el día " & DisplayDateFromSerialNumber(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE), -1, -1, -1) & " para el empleado " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID)
							Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
						Else
							If Not CheckSpecialJourneyInformationConsistency(aSpecialJourneyComponent, sErrorDescription) Then
								lErrorNumber = -1
							Else
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
									lErrorNumber = GetNewIDFromTable(oADODBConnection, "EmployeesSpecialJourneys", "RecordID", "", 1, iRedordID, sErrorDescription)
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudo guardar la información del registro."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesSpecialJourneys (RecordID, EmployeeID, EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, OriginalEmployeeID, PositionID, AreaID, ServiceID, LevelID, WorkingHours, ShiftID, RiskLevelID, SpecialJourneyID, DocumentNumber, StartDate, EndDate, StartHour, EndHour, JourneyID, WorkedHours, MovementID, FactorID, ReasonID, Comments, ConceptAmount, AddUserID, AddDate, AppliedDate, Removed, RemoveUserID, RemovedDate, AppliedRemoveDate, Active) Values (" & iRedordID & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) & ", '" & _
															aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NUMBER) & "', '" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NAME) & "', '" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME) & "', '" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME2) & "', '" & Replace(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_RFC), "'", "´") & "', '" & Replace(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_CURP), "'", "´") & "', " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ORIGINAL_EMPLOYEE_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_POSITION_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_AREA_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SERVICE_ID) & ", " & _
															aSpecialJourneyComponent(N_SPECIAL_JOURNEY_LEVEL_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_WORKING_HOURS) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SHIFT_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_RISKLEVEL_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SPECIAL_JOURNEY_ID) & ", '" & Replace(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_DOCUMENTNUMBER), "'", "´") & "', " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTHOUR) & ", " &  aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDHOUR) & ", " &  aSpecialJourneyComponent(N_SPECIAL_JOURNEY_JOURNEY_ID) & ", " & _
															aSpecialJourneyComponent(D_SPECIAL_JOURNEY_WORKED_HOURS) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_MOVEMENT_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_FACTOR_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REASON_ID) & ", '" & Replace(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_COMMENTS), "'", "´") & "', " & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_CONCEPT_AMOUNT) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_USERID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_DATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVE_USER_ID) & ", " & _
															aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED_DATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_REMOVE_DATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE) & ")", "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									End If
								End If
							End If
						End If
					End If
				Else
					lErrorNumber = -1
				End If
			ElseIf iSpecialJourneyType = 2 Then ' Externos
				If VerifyRequerimentsForExternalSpecialJourneys(oADODBConnection, aSpecialJourneyComponent, sErrorDescription) Then
					aSpecialJourneyComponent(B_IS_DUPLICATED_SPECIAL_JOURNEY) = False
					lErrorNumber = CheckExistencyOfSpecialJourney(aSpecialJourneyComponent, bIsForPeriod, sErrorDescription)
					If lErrorNumber = 0 Then
						If aSpecialJourneyComponent(B_IS_DUPLICATED_SPECIAL_JOURNEY) Then
							lErrorNumber = L_ERR_DUPLICATED_RECORD
							sErrorDescription = "Ya existe un registro el día " & DisplayDateFromSerialNumber(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE), -1, -1, -1) & " para el empleado " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID)
							Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
						Else
							If Not CheckSpecialJourneyInformationConsistency(aSpecialJourneyComponent, sErrorDescription) Then
								lErrorNumber = -1
							Else
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
									lErrorNumber = GetNewIDFromTable(oADODBConnection, "EmployeesSpecialJourneys", "RecordID", "", 1, iRedordID, sErrorDescription)
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudo guardar la información del registro."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesSpecialJourneys (RecordID, EmployeeID, EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, OriginalEmployeeID, PositionID, AreaID, ServiceID, LevelID, WorkingHours, ShiftID, RiskLevelID, SpecialJourneyID, DocumentNumber, StartDate, EndDate, StartHour, EndHour, JourneyID, WorkedHours, MovementID, FactorID, ReasonID, Comments, ConceptAmount, AddUserID, AddDate, AppliedDate, Removed, RemoveUserID, RemovedDate, AppliedRemoveDate, Active) Values (" & iRedordID & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) & ", '" & _
															aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NUMBER) & "', '" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NAME) & "', '" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME) & "', '" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME2) & "', '" & Replace(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_RFC), "'", "´") & "', '" & Replace(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_CURP), "'", "´") & "', " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ORIGINAL_EMPLOYEE_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_POSITION_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_AREA_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SERVICE_ID) & ", " & _
															aSpecialJourneyComponent(N_SPECIAL_JOURNEY_LEVEL_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_WORKING_HOURS) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SHIFT_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_RISKLEVEL_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SPECIAL_JOURNEY_ID) & ", '" & Replace(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_DOCUMENTNUMBER), "'", "´") & "', " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTHOUR) & ", " &  aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDHOUR) & ", " &  aSpecialJourneyComponent(N_SPECIAL_JOURNEY_JOURNEY_ID) & ", " & _
															aSpecialJourneyComponent(D_SPECIAL_JOURNEY_WORKED_HOURS) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_MOVEMENT_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_FACTOR_ID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REASON_ID) & ", '" & Replace(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_COMMENTS), "'", "´") & "', " & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_CONCEPT_AMOUNT) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_USERID) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_DATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVE_USER_ID) & ", " & _
															aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED_DATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_REMOVE_DATE) & ", " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE) & ")", "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
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
	End If

	AddSpecialJourney_SP = lErrorNumber
	Err.Clear
End Function

Function AddSpecialJourneyFile(oRequest, oADODBConnection, sQuery, lReasonID, aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new absence for the employee into the database
'Inputs:  oRequest, oADODBConnection, sQuery, lReasonID
'Outputs: aSpecialJourneyComponent, aJobComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddSpecialJourneyFile"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información de la aplicación de incidencias masivos de los empleados."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				If Not IsEmpty(oRequest(CStr(oRecordset.Fields("EmployeeID").Value) & CStr(oRecordset.Fields("AbsenceID").Value) & CStr(oRecordset.Fields("OcurredDate").Value))) Then
					aSpecialJourneyComponent(N_EMPLOYEE_ID_ABSENCE) = CLng(oRecordset.Fields("EmployeeID").Value)
					aSpecialJourneyComponent(N_ABSENCE_ID_ABSENCE) = CLng(oRecordset.Fields("AbsenceID").Value)
					aSpecialJourneyComponent(N_OCURRED_DATE_ABSENCE) = CLng(oRecordset.Fields("OcurredDate").Value)
					aSpecialJourneyComponent(N_APPLIED_DATE_ABSENCE) = CLng(oRequest("AppliedDate").Item)
					lErrorNumber = SetActiveForEmployeeAbsence(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
				End If
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
		End If
	End If

	Set oRecordset = Nothing
	AddSpecialJourneyFile = lErrorNumber
	Err.Clear
End Function

Function GetSpecialJourney(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about an absence for the
'         employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetSpecialJourney"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aSpecialJourneyComponent(B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
	End If

	If aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ID) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador del concepto y/o la fecha para obtener la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesSpecialJourneys Where (RecordID=" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ID) & ")", "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ID) = CLng(oRecordset.Fields("RecordID").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) = CLng(oRecordset.Fields("EmployeeID").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NUMBER) = CStr(oRecordset.Fields("EmployeeNumber").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NAME) = CInt(oRecordset.Fields("EmployeeName").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME) = CLng(oRecordset.Fields("EmployeeLastName").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME2) = CInt(oRecordset.Fields("EmployeeLastName2").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_RFC) = CStr(oRecordset.Fields("RFC").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_CURP) = CLng(oRecordset.Fields("CURP").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ORIGINAL_EMPLOYEE_ID) = CLng(oRecordset.Fields("OriginalEmployeeID").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_POSITION_ID) = CInt(oRecordset.Fields("PositionID").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_AREA_ID) = CLng(oRecordset.Fields("AreaID").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SERVICE_ID) = CLng(oRecordset.Fields("ServiceID").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_LEVEL_ID) = CLng(oRecordset.Fields("LevelID").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_WORKING_HOURS) = CLng(oRecordset.Fields("WorkingHours").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SHIFT_ID) = CLng(oRecordset.Fields("ShiftID").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_RISKLEVEL_ID) = CStr(oRecordset.Fields("RiskLevelID").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SPECIAL_JOURNEY_ID) = CInt(oRecordset.Fields("SpecialJourneyID").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_DOCUMENTNUMBER) = CLng(oRecordset.Fields("DocumentNumber").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) = CInt(oRecordset.Fields("StartDate").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE) = CStr(oRecordset.Fields("EndDate").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTHOUR) = CLng(oRecordset.Fields("StartHour").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDHOUR) = CLng(oRecordset.Fields("EndHour").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_JOURNEY_ID) = CInt(oRecordset.Fields("JourneyID").Value)
				aSpecialJourneyComponent(D_SPECIAL_JOURNEY_WORKED_HOURS) = CLng(oRecordset.Fields("WorkedHours").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_MOVEMENT_ID) = CLng(oRecordset.Fields("MovementID").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_FACTOR_ID) = CLng(oRecordset.Fields("FactorID").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REASON_ID) = CLng(oRecordset.Fields("ReasonID").Value)
                aSpecialJourneyComponent(S_SPECIAL_JOURNEY_COMMENTS) = CLng(oRecordset.Fields("Comments").Value)
				aSpecialJourneyComponent(D_SPECIAL_JOURNEY_CONCEPT_AMOUNT) = CLng(oRecordset.Fields("ConceptAmount").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_USERID) = CLng(oRecordset.Fields("AddUserID").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_DATE) = CLng(oRecordset.Fields("AddDate").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE) = CLng(oRecordset.Fields("AppliedDate").Value)
                aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED) = CLng(oRecordset.Fields("Removed").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVE_USER_ID) = CLng(oRecordset.Fields("RemoveUserID").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED_DATE) = CLng(oRecordset.Fields("RemovedDate").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_REMOVE_DATE) = CLng(oRecordset.Fields("AppliedRemoveDate").Value)
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE) = CLng(oRecordset.Fields("Active").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetSpecialJourney = lErrorNumber
	Err.Clear
End Function

Function GetSpecialJourneys(oRequest, oADODBConnection, aSpecialJourneyComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the Special Journeys for
'         the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aSpecialJourneyComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetSpecialJourneys"
	Dim sTables
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aSpecialJourneyComponent(B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
	End If

	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		sTables = ", Jobs"
		sCondition = "And (Employees.JobID=Jobs.JobID) And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Jobs.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	Else
		sTables = ""
		sCondition = ""
	End If

	Call GetStartAndEndDatesFromURL("FilterStart", "FilterEnd", "OcurredDate", False, sCondition)
	
	If Len(sCondition ) > 0 Then
		If InStr(1, sCondition , "And ", vbBinaryCompare) = 0 Then sCondition  = "And " & sCondition
	End If
    If aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) > 0 Then
        sCondition = sCondition & " And (EmployeesSpecialJourneys.EmployeeID=" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) & ")"
    Else
        sCondition = sCondition & " And (EmployeesSpecialJourneys.EmployeeID=0)"
    End If
    If aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE) > 0 Then
        sCondition = sCondition & " And (EmployeesSpecialJourneys.Active=1)"
    Else
        sCondition = sCondition & " And (EmployeesSpecialJourneys.Active=0)"
    End If

    sCondition = sCondition & aSpecialJourneyComponent(S_QUERY_CONDITION_SPECIAL_JOURNEY)

	sQuery = "Select EmployeesSpecialJourneys.*, EmployeeName + ' ' + EmployeeLastName + ' ' + EmployeeLastName2 As EmployeeFullName, RFC, CURP From EmployeesSpecialJourneys, Employees"
    sQuery = sQuery & " Where (EmployeesSpecialJourneys.EmployeeID = Employees.EmployeeID) " & sCondition & " Order By EmployeesSpecialJourneys.EmployeeID, EmployeesSpecialJourneys.StartDate"
	sErrorDescription = "No se pudo obtener la información de los registros."

	If CInt(Request.Cookies("SIAP_SectionID")) <> 7 Then  ' Dif. de Desc.
		If CInt(Request.Cookies("SIAP_SubSectionID")) = 22 Then  ' Igual a Prestaciones e incidencias
			sCondition = sCondition & " And (EmployeesSpecialJourneys.PositionID=323)"
		Else ' Igual a Inf. - Emp. - Inci
			sCondition = sCondition & " And (EmployeesSpecialJourneys.PositionID<>323)"
		End If
	Else ' Igual a Desc.
		If CInt(Request.Cookies("SIAP_SubSectionID")) = 721 Then  ' Igual a Prestaciones e incidencias
			sCondition = sCondition & " And (EmployeesSpecialJourneys.PositionID=323)"
		Else
			sCondition = sCondition & " And (EmployeesSpecialJourneys.PositionID<>323)"
		End If
	End If
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & sQuery & """ />"

	GetSpecialJourneys = lErrorNumber
	Err.Clear
End Function

Function GetCrossingSpecialJourneys(oADODBConnection, aSpecialJourneyComponent, sAbsenceIDs, sAbsenceCrossType, lAbsenceID, lStartDate, lEndDate, lVacationPeriod, lDays, sErrorDescription)
'************************************************************
'Purpose: To get the type of crossing absence for the
'         absence to insert
'Inputs:  oRequest, oADODBConnection, aSpecialJourneyComponent
'Outputs: sAbsenceIDs, lStartDate, lEndDate, lVacationPeriod, lDays, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetCrossingSpecialJourneys"
	Dim oRecordset
	Dim lErrorNumber
	Dim sQuery

	If (InStr(1, sAbsenceIDs, "-1", vbBinaryCompare) > 0) And (aSpecialJourneyComponent(N_ABSENCE_ID_ABSENCE) = 34) Then
		sAbsenceIDs = "35, 37"
		sQuery = "Select * from EmployeesSpecialJourneys Where (EmployeeID = " & aSpecialJourneyComponent(N_ID_EMPLOYEE) & ") And (AbsenceID IN (" & sAbsenceIDs & "))" & _
				 " And (((OcurredDate <= " &  aSpecialJourneyComponent(N_OCURRED_DATE_ABSENCE) & ") And (EndDate >= " &  aSpecialJourneyComponent(N_OCURRED_DATE_ABSENCE) & "))" & _
				 " And ((OcurredDate <= " &  aSpecialJourneyComponent(N_END_DATE_ABSENCE) & ") And (EndDate >= " &  aSpecialJourneyComponent(N_END_DATE_ABSENCE) & ")" & _
				 " Or (OcurredDate <= " &  aSpecialJourneyComponent(N_END_DATE_ABSENCE) & ") And (EndDate <= " &  aSpecialJourneyComponent(N_END_DATE_ABSENCE) & ")))"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sAbsenceCrossType = "Inner"
				lAbsenceID = CInt(oRecordset.Fields("AbsenceID").Value)
				lStartDate = CLng(oRecordset.Fields("OcurredDate").Value)
				lEndDate = CLng(oRecordset.Fields("EndDate").Value)
				lVacationPeriod = CLng(oRecordset.Fields("VacationPeriod").Value)
				lDays = CInt(oRecordset.Fields("AbsenceHours").Value)
			Else
				sQuery = "Select * from EmployeesSpecialJourneys Where (EmployeeID = " & aSpecialJourneyComponent(N_ID_EMPLOYEE) & ") And (AbsenceID IN (" & sAbsenceIDs & "))" & _
						 " And (((OcurredDate > " & aSpecialJourneyComponent(N_OCURRED_DATE_ABSENCE) & ") And (EndDate > " &  aSpecialJourneyComponent(N_OCURRED_DATE_ABSENCE) & "))" & _
						 " And ((OcurredDate <= " &  aSpecialJourneyComponent(N_END_DATE_ABSENCE) & ") And (EndDate > " &  aSpecialJourneyComponent(N_END_DATE_ABSENCE) & ")))"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						sAbsenceCrossType = "Left"
						lAbsenceID = CInt(oRecordset.Fields("AbsenceID").Value)
						lStartDate = CLng(oRecordset.Fields("OcurredDate").Value)
						lEndDate = CLng(oRecordset.Fields("EndDate").Value)
						lVacationPeriod = CLng(oRecordset.Fields("VacationPeriod").Value)
						lDays = CInt(oRecordset.Fields("AbsenceHours").Value)
					Else
						sQuery = "Select * from EmployeesSpecialJourneys Where (EmployeeID = " & aSpecialJourneyComponent(N_ID_EMPLOYEE) & ") And (AbsenceID IN (" & sAbsenceIDs & "))" & _
								 " And (((OcurredDate < " & aSpecialJourneyComponent(N_OCURRED_DATE_ABSENCE) & ") And (EndDate >= " &  aSpecialJourneyComponent(N_OCURRED_DATE_ABSENCE) & "))" & _
								 " And ((OcurredDate < " &  aSpecialJourneyComponent(N_END_DATE_ABSENCE) & ") And (EndDate < " &  aSpecialJourneyComponent(N_END_DATE_ABSENCE) & ")))"
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
		If InStr(1, ",50,51,52,53,54,55,56,", "," & aSpecialJourneyComponent(N_ABSENCE_ID_ABSENCE) & ",", vbBinaryCompare) > 0 Then
			sAbsenceIDs = "50,51,52,53,54,55,56"
		Else
			sAbsenceIDs = "41,42,43,44,45,46,47,48,49,57,58"
		End If
		sQuery = "Select * From EmployeesSpecialJourneys Where (EmployeeID = " & aSpecialJourneyComponent(N_ID_EMPLOYEE) & ")" & _
				 " And (AbsenceID IN (" & sAbsenceIDs & "))" & _
				 " And (OcurredDate>=" & aSpecialJourneyComponent(N_OCURRED_DATE_ABSENCE) & ")" & _
				 " And (EndDate>=OcurredDate)"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sAbsenceCrossType = "Cross"
				lAbsenceID = CInt(oRecordset.Fields("AbsenceID").Value)
				lStartDate = CLng(oRecordset.Fields("OcurredDate").Value)
				lEndDate = CLng(oRecordset.Fields("EndDate").Value)
			Else
				sQuery = "Select * From EmployeesSpecialJourneys Where (EmployeeID = " & aSpecialJourneyComponent(N_ID_EMPLOYEE) & ")" & _
						 " And (AbsenceID IN (" & sAbsenceIDs & "))" & _
						 " And (OcurredDate<" & aSpecialJourneyComponent(N_OCURRED_DATE_ABSENCE) & ")" & _
						 " And (EndDate>=" &  aSpecialJourneyComponent(N_OCURRED_DATE_ABSENCE) & ")"
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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
	GetCrossingSpecialJourneys = lErrorNumber
	Err.Clear
End Function

Function ModifySpecialJourney(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing Special Journey for the employee in
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifySpecialJourney"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aSpecialJourneyComponent(B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
	End If

	If (aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) = -1) Or (aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) = -1) Or (aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador del concepto y/o la fecha para modificar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If lErrorNumber = 0 Then
			If aSpecialJourneyComponent(B_IS_DUPLICATED_ABSENCE) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un registro para el " & DisplayDateFromSerialNumber(aSpecialJourneyComponent(N_OCURRED_DATE_ABSENCE), -1, -1, -1) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				If Not CheckAbsenceInformationConsistency(aSpecialJourneyComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					sErrorDescription = "No se pudo modificar la información del registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesSpecialJourneys Set DocumentNumber='" & Replace(aSpecialJourneyComponent(S_DOCUMENT_NUMBER_ABSENCE), "'", "´") & "', AbsenceHours=" & aSpecialJourneyComponent(N_HOURS_ABSENCE) & ", JustificationID=" & aSpecialJourneyComponent(N_JUSTIFICATION_ID_ABSENCE) & ", AppliesForPunctuality=" & aSpecialJourneyComponent(N_APPLIES_FOR_PUNCTUALITY_ABSENCE) & ", Removed=" & aSpecialJourneyComponent(N_REMOVED_ABSENCE) & ", RemoveUserID=" & aSpecialJourneyComponent(N_REMOVE_USER_ID_ABSENCE) & ", RemovedDate=" & aSpecialJourneyComponent(N_REMOVED_DATE_ABSENCE) & ", AppliedRemoveDate=" & aSpecialJourneyComponent(N_APPLIED_REMOVE_DATE_ABSENCE) & ", Active=" & aSpecialJourneyComponent(N_ACTIVE_ABSENCE) & " Where (EmployeeID=" & aSpecialJourneyComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & aSpecialJourneyComponent(N_ABSENCE_ID_ABSENCE) & ") And (OcurredDate=" & aSpecialJourneyComponent(N_OCURRED_DATE_ABSENCE) & ")", "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
		End If
	End If

	ModifySpecialJourney = lErrorNumber
	Err.Clear
End Function

Function CancelSpecialJourney(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To remove an absence for the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CancelSpecialJourney"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aSpecialJourneyComponent(B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
	End If

	If (aSpecialJourneyComponent(N_EMPLOYEE_ID_ABSENCE) = -1) Or (aSpecialJourneyComponent(N_ABSENCE_ID_ABSENCE) = -1) Or (aSpecialJourneyComponent(N_OCURRED_DATE_ABSENCE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador del concepto y/o la fecha para eliminar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If IsEmpty(iAbsenceID) Then iAbsenceID = aSpecialJourneyComponent(N_ABSENCE_ID_ABSENCE)
		sErrorDescription = "No se pudo cancelar la incidencia del día " + aSpecialJourneyComponent(N_OCURRED_DATE_ABSENCE)
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesSpecialJourneys Set Removed=1, DocumentNumber='" & aSpecialJourneyComponent(S_DOCUMENT_NUMBER_ABSENCE) & "', RemoveUserID=" & aSpecialJourneyComponent(N_REMOVE_USER_ID_ABSENCE) & ", RemovedDate=" & aSpecialJourneyComponent(N_REMOVED_DATE_ABSENCE) & ", AppliedRemoveDate=" & aSpecialJourneyComponent(N_APPLIED_REMOVE_DATE_ABSENCE) & ", Active=" & aSpecialJourneyComponent(N_ACTIVE_ABSENCE) & " Where (EmployeeID=" & aSpecialJourneyComponent(N_EMPLOYEE_ID_ABSENCE) & ") And (AbsenceID=" & aSpecialJourneyComponent(N_ABSENCE_ID_ABSENCE) & ") And (OcurredDate=" & aSpecialJourneyComponent(N_OCURRED_DATE_ABSENCE) & ")", "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	CancelSpecialJourney = lErrorNumber
	Err.Clear
End Function

Function RemoveSpecialJourney(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To remove an absence for the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveSpecialJourney"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aSpecialJourneyComponent(B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
	End If

	If (aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ID) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador para eliminar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesSpecialJourneys Where (RecordID=" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ID) & ")", "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	RemoveSpecialJourney = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfExternalEmployee(aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific absence for the employee
'         exists in the database
'Inputs:  aSpecialJourneyComponent
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfExternalEmployee"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aSpecialJourneyComponent(B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeInitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
	End If

	If Len(aEmployeeComponent(S_RFC_EMPLOYEE)) < 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el RFC del empleado para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."

		sQuery = "Select * From ExternalSpecialJourneys Where (RFC='" & aEmployeeComponent(S_RFC_EMPLOYEE) & "')"
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) = CLng(oRecordset.Fields("ExternalID").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NUMBER) = CStr(oRecordset.Fields("EmployeeNumber").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NAME) = CStr(oRecordset.Fields("EmployeeName").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME) = CStr(oRecordset.Fields("EmployeeLastName").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME2) = CStr(oRecordset.Fields("EmployeeLastName2").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_RFC) = CStr(oRecordset.Fields("RFC").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_CURP) = CStr(oRecordset.Fields("CURP").Value)
				aSpecialJourneyComponent(S_SPECIAL_JOURNEY_SPEP) = CStr(oRecordset.Fields("SPEPID").Value)
				aSpecialJourneyComponent(B_SPECIAL_JOURNEY_EXIST_EXTERNAL) = True
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfExternalEmployee = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfSpecialJourney(aSpecialJourneyComponent, bIsForPeriod, sErrorDescription)
'************************************************************
'Purpose: To check if a specific absence for the employee
'         exists in the database
'Inputs:  aSpecialJourneyComponent
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfSpecialJourney"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aSpecialJourneyComponent(B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
	End If

	If (aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) = -1) Or (aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) = -1) Or (aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del empleado para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."

		sQuery = "Select * From EmployeesSpecialJourneys Where (EmployeeID=" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) & ")"
		If bIsForPeriod Then
			sQuery = sQuery & _
					 " And (((StartDate >= " &  aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) & ") And (StartDate <= " &  aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE) & "))" & _
					 " Or ((EndDate >= " &  aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) & ") And (EndDate <= " &  aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE) & "))" & _
					 " Or ((EndDate >= " &  aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) & ") And (StartDate <= " &  aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE) & ")))"
		Else
			sQuery = sQuery & " And (StartDate=" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE) & ")"
		End If

		lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aSpecialJourneyComponent(B_IS_DUPLICATED_SPECIAL_JOURNEY) = True
			Else
				aSpecialJourneyComponent(B_IS_DUPLICATED_SPECIAL_JOURNEY) = False
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfSpecialJourney = lErrorNumber
	Err.Clear
End Function

Function CheckSpecialJourneyInformationConsistency(aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aSpecialJourneyComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckSpecialJourneyInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del empleado no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La fecha de inicio del registro no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La fecha de fin del registro no es un valor numérico."
		bIsCorrect = False
	End If
	If Not IsNumeric(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_DATE)) Then aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_DATE) = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	If Len(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_DOCUMENTNUMBER)) = 0 Then aSpecialJourneyComponent(S_SPECIAL_JOURNEY_DOCUMENTNUMBER) = " "
	'If Not IsNumeric(aSpecialJourneyComponent(N_HOURS_ABSENCE)) Then aSpecialJourneyComponent(N_HOURS_ABSENCE) = 0
	If Not IsNumeric(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_USERID)) Then aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ADD_USERID) = aLoginComponent(N_SPECIAL_JOURNEY_ADD_USERID)
	If Not IsNumeric(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE)) Then aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE) = 0
	If Not IsNumeric(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED)) Then aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED) = 0
	If Not IsNumeric(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVE_USER_ID)) Then aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVE_USER_ID) = -1
	If Not IsNumeric(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED_DATE)) Then aSpecialJourneyComponent(N_SPECIAL_JOURNEY_REMOVED_DATE) = 0
	If Not IsNumeric(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_REMOVE_DATE)) Then aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_REMOVE_DATE) = 0

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del registro contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckSpecialJourneyInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayAnotherJourneyForm(oRequest, oADODBConnection, sASPFileName, sAction, iLeftWidth, iSpecialJourneyType, sAltDescription, sDescription, sErrorDescription)
'************************************************************
'Purpose: To display the information about a registration of child for the
'         employee from the database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sASPFileName, sAction, sURL, aEmployeeComponent
'Outputs: aEmployeeComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayAnotherEmployeeForm"
	Dim sNames
	Dim lErrorNumber

    Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "function CheckAnotherSpecialJourneyFields(oForm) {" & vbNewLine
            Response.Write "if (oForm) {" & vbNewLine
                If iSpecialJourneyType = 1 Then
                    Response.Write "if (oForm.EmployeeID.value.length == 0) {" & vbNewLine
                        Response.Write "alert('Favor de introducir un valor para el campo \'Número del empleado\'.');" & vbNewLine
                        Response.Write "oForm.EmployeeID.focus();" & vbNewLine
                        Response.Write "return false;" & vbNewLine
                    Response.Write "}" & vbNewLine
                Else
                    Response.Write "if (oForm.RFC.value.length == 0) {" & vbNewLine
                        Response.Write "alert('Favor de introducir un valor para el campo \'RFC\'.');" & vbNewLine
                        Response.Write "oForm.RFC.focus();" & vbNewLine
                        Response.Write "return false;" & vbNewLine
                    Response.Write "}" & vbNewLine
                End If
            Response.Write "}" & vbNewLine
		Response.Write "} // End of CheckAnotherSpecialJourneyFields" & vbNewLine
    Response.Write "//--></SCRIPT>" & vbNewLine

	Response.Write "<FORM NAME=""AnotherConceptFrm"" ID=""AnotherConceptFrm"" ACTION=""" & sASPFileName & """ METHOD=""GET"" onSubmit=""return CheckAnotherSpecialJourneyFields(this)"">"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""" & iLeftWidth & """>&nbsp;</TD>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><IMG SRC=""Images/MnLeftArrows.gif"" WIDTH=""32"" HEIGHT=""32"" ALT=""" & sAltDescription & """ BORDER=""0"" /><BR /></TD>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""350""><FONT FACE=""Arial"" SIZE=""2""><B>Otro empleado</B><BR /></FONT>"
				Response.Write "<DIV CLASS=""MenuOverflow""><FONT FACE=""Arial"" SIZE=""2"">" & sDescription & "</FONT></DIV></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""" & iLeftWidth & """>&nbsp;</TD>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;</FONT></TD>"
                If iSpecialJourneyType = 1 Then
				    Response.Write "<TD VALIGN=""TOP"" WIDTH=""350""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;Número del empleado:&nbsp;</FONT><INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" /></TD>"
                Else
                    Response.Write "<TD VALIGN=""TOP"" WIDTH=""350""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;RFC del externo:&nbsp;</FONT><INPUT TYPE=""TEXT"" NAME=""RFC"" ID=""RFCTxt"" SIZE=""13"" MAXLENGTH=""13"" CLASS=""TextFields"" /></TD>"
                End If
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & sAction & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SpecialJourneyType"" ID=""SpecialJourneyTypeHdn"" VALUE=""" & iSpecialJourneyType & """ />"
				If InStr(1, sAction, "ServiceSheet", vbBinaryCompare) > 0 Then
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & lReasonID & """ />"
				End If
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""" & iLeftWidth & """>&nbsp;</TD>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;</FONT></TD>"
				Response.Write "<TD VALIGN=""TOP"" WIDTH=""350""><INPUT TYPE=""SUBMIT"" NAME=""EmployeeConcept"" ID=""EmployeeConceptBtn"" VALUE=""Buscar empleado"" CLASS=""Buttons"" ALT=""" & sAltDescription & """/></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	DisplayAnotherEmployeeForm = lErrorNumber
	Err.Clear
End Function

Function DisplayExternalSpecialJourneyForm(oRequest, oADODBConnection, aSpecialJourneyComponent, iSpecialJourneyType, sErrorDescription)
'************************************************************
'Purpose: To display the information about an absence for the
'         employee from the database using a HTML Form
'Inputs:  oRequest, oADODBConnection, aSpecialJourneyComponent
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayExternalSpecialJourneyForm"
	Dim sNames
	Dim aRelatedAbsences
	Dim iIndex
	Dim oRecordset
	Dim lErrorNumber
	Dim sAbsenceIDs
	Dim sCaseOptions

	If lErrorNumber = 0 Then
		If Len(aEmployeeComponent(S_RFC_EMPLOYEE)) > 0 Then
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR><TD VALIGN=""TOP"">"
		End If
			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				If Len(aEmployeeComponent(S_RFC_EMPLOYEE)) > 0 Then
					Response.Write "var bBlock = true;" & vbNewLine
					Response.Write "function DoBlock() {" & vbNewLine
						Response.Write "if (bBlock) {" & vbNewLine
							Response.Write "document.SpecialJourneyFrm.DocumentNumber.focus();" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "} // End of DoBlock" & vbNewLine
				End If
				Response.Write "function CheckSpecialJourneyFields(oForm) {" & vbNewLine
					Response.Write "if (oForm) {" & vbNewLine
						If Len(aEmployeeComponent(S_RFC_EMPLOYEE)) = 0 Then
							Response.Write "if (oForm.RFC.value.length == 0) {" & vbNewLine
								Response.Write "alert('Favor de introducir un valor para el campo \'RFC\'.');" & vbNewLine
								Response.Write "oForm.RFC.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "if (oForm.RFC.value.length < 13) {" & vbNewLine
								Response.Write "alert('El RFC debe ser de 13 posiciones.');" & vbNewLine
								Response.Write "oForm.RFC.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						Else
							If CInt(Request.Cookies("SIAP_SubSectionID")) = 424 Then
								Response.Write "if (oForm.OriginalEmployeeID.value == '') {" & vbNewLine
									Response.Write "alert('Favor de especificar y validar el número del empleado a suplir.');" & vbNewLine
									Response.Write "oForm.OriginalEmployeeID.focus();" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine
								Response.Write "if ((oForm.CheckOriginalEmployeeID.value == '') || (oForm.CheckOriginalEmployeeID.value == '-1')) {" & vbNewLine
									Response.Write "alert('Favor de validar el número del empleado a suplir.');" & vbNewLine
									Response.Write "oForm.OriginalEmployeeID.focus();" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine
							End If
							Response.Write "if (oForm.DocumentNumber.value.length == 0) {" & vbNewLine
								Response.Write "alert('Favor de introducir un valor para el campo \'Folio de autorización\'.');" & vbNewLine
								Response.Write "oForm.DocumentNumber.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "oForm.WorkedHours.value = oForm.WorkedHours.value.replace(/,/gi, '');" & vbNewLine
							Response.Write "if (!CheckFloatValue(oForm.WorkedHours, 'el campo \'Días/horas reportadas\'', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "oForm.ConceptAmount.value = oForm.ConceptAmount.value.replace(/,/gi, '');" & vbNewLine
							Response.Write "if (!CheckFloatValue(oForm.ConceptAmount, 'el campo \'Percepción\'', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "if ((oForm.TempStartDate.value != BuildDateString(document.SpecialJourneyFrm.StartDateYear.value, document.SpecialJourneyFrm.StartDateMonth.value, document.SpecialJourneyFrm.StartDateDay.value)) || (oForm.TempEndDate.value != BuildDateString(document.SpecialJourneyFrm.EndDateYear.value, document.SpecialJourneyFrm.EndDateMonth.value, document.SpecialJourneyFrm.EndDateDay.value))) {" & vbNewLine
								Response.Write "alert('Favor de validar que el empleado no tenga registros en las fechas indicadas.');" & vbNewLine
								Response.Write "oForm.StartDateDay.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "if (oForm.ReportedHours.value == '') {" & vbNewLine
								Response.Write "alert('Favor de validar los días/horas registradas para el empleado.');" & vbNewLine
								Response.Write "oForm.WorkedHours.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "if (oForm.ReportedHours.value == '-3') {" & vbNewLine
								Response.Write "alert('No existe Presupuesto para pagar la Guardia/Suplencia.');" & vbNewLine
								Response.Write "oForm.WorkedHours.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "if (oForm.ReportedHours.value == '-2') {" & vbNewLine
								Response.Write "alert('Ya existen otros registros en las fechas indicadas.');" & vbNewLine
								Response.Write "oForm.WorkedHours.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "} else {" & vbNewLine
								Response.Write "if (oForm.ReportedHours.value != '1') {" & vbNewLine
									Response.Write "alert('Las horas registradas para el empleado en la quincena exceden el número de horas establecidas como máximo.');" & vbNewLine
									Response.Write "oForm.WorkedHours.focus();" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine
							Response.Write "}" & vbNewLine
						End If
					Response.Write "}" & vbNewLine
					'Response.Write "return ValidateSpecialJourneyFields(window.document.SpecialJourneyFrm);" & vbNewLine
					Response.Write "return true;" & vbNewLine
				Response.Write "} // End of CheckSpecialJourneyFields" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine

		If Len(aEmployeeComponent(S_RFC_EMPLOYEE)) = 0 Then
			Response.Write "<FORM NAME=""SpecialJourneyFrm"" ID=""SpecialJourneyFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""GET"" onSubmit=""return CheckSpecialJourneyFields(this)"">"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""SpecialJourney"" />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SpecialJourneyType"" ID=""SpecialJourneyTypeHdn"" VALUE=""" & iSpecialJourneyType & """ />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SpecialJourneyID"" ID=""SpecialJourneyIDHdn"" VALUE=""" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SPECIAL_JOURNEY_ID) & """ />"
				Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">RFC del Externo:&nbsp;</FONT></TD>"
						Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""RFC"" ID=""RFCTxt"" VALUE=""" & aEmployeeComponent(S_RFC_EMPLOYEE) & """ SIZE=""13"" MAXLENGTH=""13"" CLASS=""TextFields"" /></TD>"
					Response.Write "</TR>"
				Response.Write "</TABLE>"
				Response.Write "<BR /><BR />"
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then
					Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""SpecialJourneyMovement"" ID=""SpecialJourneyMovementBtn"" VALUE=""Buscar externo"" CLASS=""Buttons"" />"
				End If
			Response.Write "</FORM>"
		Else
			lErrorNumber = CheckExistencyOfExternalEmployee(aSpecialJourneyComponent, sErrorDescription)
			If lErrorNumber = L_ERR_NO_RECORDS Then
			    Call DisplayErrorMessage("Advertencia", "Este RFC no se encuentra registrado, capture correctamente los datos debido a que serán registrados cuando guarde el registro.")
			End If
			If True Then
				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		            Response.Write "function CheckAnotherSpecialJourneyFields(oForm) {" & vbNewLine
						Response.Write "if (oForm) {" & vbNewLine
							If iSpecialJourneyType = 1 Then
								Response.Write "if (oForm.EmployeeID.value.length == 0) {" & vbNewLine
									Response.Write "alert('Favor de introducir un valor para el campo \'Número del empleado\'.');" & vbNewLine
									Response.Write "oForm.EmployeeID.focus();" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine
							Else
								Response.Write "if (oForm.RFC.value.length == 0) {" & vbNewLine
									Response.Write "alert('Favor de introducir un valor para el campo \'RFC\'.');" & vbNewLine
									Response.Write "oForm.RFC.focus();" & vbNewLine
									Response.Write "return false;" & vbNewLine
								Response.Write "}" & vbNewLine
							End If
						Response.Write "}" & vbNewLine
					Response.Write "} // End of CheckAnotherSpecialJourneyFields" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
				If CInt(Request.Cookies("SIAP_SectionID")) <> 1 Then
					Response.Write "<FORM NAME=""AnotherSpecialJourneyFrm"" ID=""AnotherSpecialJourneyFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""GET"" onSubmit="" return CheckAnotherSpecialJourneyFields(this)"">"
						Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
							Response.Write "<TR>"
								Response.Write "<TD VALIGN=""TOP"" WIDTH=""500"">&nbsp;</TD>"
								Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><IMG SRC=""Images/MnLeftArrows.gif"" WIDTH=""32"" HEIGHT=""32"" ALT=""Guardias"" BORDER=""0"" /><BR /></TD>"
								Response.Write "<TD VALIGN=""TOP"" WIDTH=""290""><FONT FACE=""Arial"" SIZE=""2""><B>Otro empleado</B><BR /></FONT>"
								Response.Write "<DIV CLASS=""MenuOverflow""><FONT FACE=""Arial"" SIZE=""2"">Registre guardias a un empleado diferente.</FONT></DIV></TD>"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD VALIGN=""TOP"" WIDTH=""500"">&nbsp;</TD>"
								Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;</FONT></TD>"
								If iSpecialJourneyType = 1 Then
									Response.Write "<TD VALIGN=""TOP"" WIDTH=""350""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;Número del empleado:&nbsp;</FONT><INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" /></TD>"
								Else
									Response.Write "<TD VALIGN=""TOP"" WIDTH=""350""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;RFC del externo:&nbsp;</FONT><INPUT TYPE=""TEXT"" NAME=""RFC"" ID=""RFCTxt"" SIZE=""13"" MAXLENGTH=""13"" CLASS=""TextFields"" /></TD>"
								End If
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""SpecialJourney"" />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SpecialJourneyType"" ID=""SpecialJourneyTypeHdn"" VALUE=""" & iSpecialJourneyType & """ />"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SpecialJourneyID"" ID=""SpecialJourneyIDHdn"" VALUE=""" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SPECIAL_JOURNEY_ID) & """ />"
							Response.Write "</TR>"
							Response.Write "<TR>"
								Response.Write "<TD VALIGN=""TOP"" WIDTH=""500"">&nbsp;</TD>"
								Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;</FONT></TD>"
								Response.Write "<TD VALIGN=""TOP"" WIDTH=""290""><INPUT TYPE=""SUBMIT"" NAME=""SpecialJourneyMovement"" ID=""SpecialJourneyMovementBtn"" VALUE=""Buscar empleado"" CLASS=""Buttons"" /></TD>"
							Response.Write "</TR>"
						Response.Write "</TABLE>"
					Response.Write "</FORM>"
				End If

				Response.Write "<FORM NAME=""SpecialJourneyFrm"" ID=""SpecialJourneyFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""GET"" onSubmit=""return CheckSpecialJourneyFields(this)"">"
					Response.Write "<BR /><BR />"
					If Len(oRequest("SpecialJourneyChange").Item) > 0 Then
						lErrorNumber = GetSpecialJourney(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SpecialJourneyChange"" ID=""SpecialJourneyChangeHdn"" VALUE=""" & oRequest("SpecialJourneyChange").Item & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RemoveUserID"" ID=""RemoveUserIDHdn"" VALUE=""" & aLoginComponent(N_USER_ID_LOGIN) & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RemovedDate"" ID=""RemovedDateHdn"" VALUE=""" & Left(GetSerialNumberForDate(""), Len("00000000")) & """ />"
						aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE) = 0
					End If
					If aSpecialJourneyComponent(B_SPECIAL_JOURNEY_EXIST_EXTERNAL) Then
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ExternalExist"" ID=""ExternalExistHdn"" VALUE=""1"" />"
					End If
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""SpecialJourney"" />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SpecialJourneyType"" ID=""SpecialJourneyTypeHdn"" VALUE=""" & iSpecialJourneyType & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RecordID"" ID=""RecordIDHdn"" VALUE=""-1"" />"
					If CInt(Request.Cookies("SIAP_SubSectionID")) <> 424 Then
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OriginalEmployeeID"" ID=""OriginalEmployeeIDHdn"" VALUE=""-1"" />"
					End If
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeNumber"" ID=""EmployeeNumberHdn"" VALUE=""800000"" />"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Información general</B></FONT>"
					Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
					Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del externo:&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NUMBER) & "</FONT><INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NUMBER) & """ /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							If (Len(oRequest("Success").Item) > 0) Then
								Response.Write "&nbsp;&nbsp;&nbsp;<TD VALIGN=""TOP"" ALIGN=""LEFT"" WIDTH=""60%"" ROWSPAN=""9"">"
								If CInt(oRequest("Success").Item) = 1 Then
									Call DisplayErrorMessage("Confirmación", "La operación con la guardia fué ejecutada exitosamente.")
								Else
									Call DisplayErrorMessage("Error al realizar la operación con la guardia.")
								End If
								Response.Write "</TD>"
							End If
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre(s):&nbsp;</FONT></TD>"
							'Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeName"" ID=""EmployeeNameTxt"" VALUE=""" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NAME) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_NAME) & " " & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME) & " " & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_EMPLOYEE_LASTNAME2)) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">RFC:&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_RFC)) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">CURP:&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_CURP)) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">SPEP ID:&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aSpecialJourneyComponent(S_SPECIAL_JOURNEY_SPEP)) & "</FONT></TD>"
						Response.Write "</TR>"
					Response.Write "</TABLE>"
                    Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
                    Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
						If Len(oRequest("SpecialJourneyChange").Item) > 0 Then
							Response.Write "<TR NAME=""AreaIDDiv"" ID=""AreaIDDiv"">"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Adscripción:&nbsp;</NOBR></FONT></TD>"
								Response.Write "<TD><SELECT NAME=""AreaID"" ID=""AreaIDCmb"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(ParentID=-1) Or (ParentID=-2)", "AreaCode", "", "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
								Response.Write "</SELECT>&nbsp;"
								Response.Write "</TD>"
							Response.Write "</TR>"
						Else
							Response.Write "<TR NAME=""AreaIDDiv"" ID=""AreaIDDiv"">"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Adscripción:&nbsp;</NOBR></FONT></TD>"
								Response.Write "<TD><SELECT NAME=""AreaID"" ID=""AreaIDCmb"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(ParentID=-1) Or (ParentID=-2)", "AreaCode", "", "Seleccione un puesto;;;-1", sErrorDescription)
								Response.Write "</SELECT>&nbsp;"
								Response.Write "</TD>"
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!-- SelectItemByValue('-1', false, document.SpecialJourneyFrm.AreaID) //--></SCRIPT>"
							Response.Write "</TR>"
						End If
						If CInt(Request.Cookies("SIAP_SubSectionID")) = 424 Then
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. del empleado a suplir:&nbsp;</FONT></TD>"
								Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""OriginalEmployeeID"" ID=""OriginalEmployeeIDTxt"" VALUE=""" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ORIGINAL_EMPLOYEE_ID) & """ SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" onChange=""document.SpecialJourneyFrm.CheckOriginalEmployeeID.value=''; document.SpecialJourneyFrm.ReportedHours.value='';"" />"
								Response.Write "<A HREF=""javascript: if (parseInt(document.SpecialJourneyFrm.EmployeeID.value) == parseInt(document.SpecialJourneyFrm.OriginalEmployeeID.value)) {alert('El empleado suplido no puede ser el mismo que el empleado suplente.'); document.SpecialJourneyFrm.OriginalEmployeeID.focus();} else {SearchRecord(document.SpecialJourneyFrm.OriginalEmployeeID.value, 'EmployeesGyS&RecordType=2&Original=1&EmployeeID=' + document.SpecialJourneyFrm.EmployeeID.value + '&RecordDate=' + document.SpecialJourneyFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'SpecialJourneyFrm');}""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado a suplir"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CheckOriginalEmployeeID"" ID=""CheckOriginalEmployeeIDHdn"" VALUE=""""  />"
							Response.Write "</TD></TR>"
							Response.Write "<TR NAME=""PositionIDDiv"" ID=""PositionIDDiv""><INPUT TYPE=""HIDDEN"" NAME=""PositionID"" ID=""PositionIDHdn"" VALUE=""-1"" /></TR>"
							Response.Write "<TR NAME=""ServiceIDDiv"" ID=""ServiceIDDiv""><INPUT TYPE=""HIDDEN"" NAME=""ServiceID"" ID=""ServiceIDHdn"" VALUE=""-1"" /></TR>"
							Response.Write "<TR NAME=""LevelIDDiv"" ID=""LevelIDDiv""><INPUT TYPE=""HIDDEN"" NAME=""LevelID"" ID=""LevelIDHdn"" VALUE=""-1"" /></TR>"
							Response.Write "<TR NAME=""WorkingHoursDiv"" ID=""WorkingHoursDiv""><INPUT TYPE=""HIDDEN"" NAME=""WorkingHours"" ID=""WorkingHoursHdn"" VALUE=""-1"" /></TR>"
						Else
							Response.Write "<TR NAME=""PositionIDDiv"" ID=""PositionIDDiv"">"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Puesto:&nbsp;</NOBR></FONT></TD>"
								Response.Write "<TD><SELECT NAME=""PositionID"" ID=""PositionIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""SearchRecord('P', 'PositionsGyS&PositionID=' + document.SpecialJourneyFrm.PositionID.value + '&AreaID=-1&ServiceID=-1&LevelID=-1&WorkingHours=-1&RecordType=1&RecordDate=' + document.SpecialJourneyFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'SpecialJourneyFrm')"">"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PositionsSpecialJourneysLKP, Positions", "Distinct Positions.PositionID", "PositionShortName, PositionName, 'Horas laboradas:' As Temp, PositionsSpecialJourneysLKP.WorkingHours", "(PositionsSpecialJourneysLKP.PositionID=Positions.PositionID) And (Positions.EndDate=30000000) And (Positions.Active=1)", "PositionShortName, PositionsSpecialJourneysLKP.WorkingHours", "", "", sErrorDescription)
								Response.Write "</SELECT>&nbsp;"
								Response.Write "</TD>"
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!-- SelectItemByValue('-1', false, document.SpecialJourneyFrm.PositionID) //--></SCRIPT>"
							Response.Write "</TR>"
							Response.Write "<TR NAME=""ServiceIDDiv"" ID=""ServiceIDDiv"">"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Servicio:&nbsp;</NOBR></FONT></TD>"
								Response.Write "<TD><SELECT NAME=""ServiceID"" ID=""ServiceIDCmb"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Services", "ServiceID", "ServiceShortName, ServiceName", "(ServiceID=-1) And (ServiceID=-2)", "ServiceShortName", "", "Seleccione un puesto;;;-1", sErrorDescription)
								Response.Write "</SELECT>&nbsp;"
								Response.Write "</TD>"
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!-- SelectItemByValue('-1', false, document.SpecialJourneyFrm.ServiceID) //--></SCRIPT>"
							Response.Write "</TR>"
							Response.Write "<TR NAME=""LevelIDDiv"" ID=""LevelIDDiv"">"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Nivel/subnivel:&nbsp;</NOBR></FONT></TD>"
								Response.Write "<TD><SELECT NAME=""LevelID"" ID=""LevelIDCmb"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Levels", "LevelID", "LevelShortName", "(LevelID=-1) And (LevelID=-2)", "LevelShortName", "", "Seleccione un puesto;;;-1", sErrorDescription)
								Response.Write "</SELECT>&nbsp;"
								Response.Write "</TD>"
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!-- SelectItemByValue('-1', false, document.SpecialJourneyFrm.LevelID) //--></SCRIPT>"
							Response.Write "</TR>"
							Response.Write "<TR NAME=""WorkingHoursDiv"" ID=""WorkingHoursDiv"">"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Horas laboradas:&nbsp;</NOBR></FONT></TD>"
								Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""WorkingHours"" ID=""WorkingHoursCmb"" SIZE=""1"" SIZE=""1"" CLASS=""Lists""  onChange=""SearchRecord('W', 'PositionsGyS&PositionID=' + document.SpecialJourneyFrm.PositionID.value + '&AreaID=' + document.SpecialJourneyFrm.AreaID.value + '&ServiceID=' + document.SpecialJourneyFrm.ServiceID.value + '&LevelID=' + document.SpecialJourneyFrm.LevelID.value + '&WorkingHours=' + document.SpecialJourneyFrm.WorkingHours.value + '&RecordType=1&RecordDate=' + document.SpecialJourneyFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'SpecialJourneyFrm')"">"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Levels", "LevelID", "LevelShortName", "(LevelID=-1) And (LevelID=-2)", "LevelShortName", "", "Seleccione un puesto;;;-1", sErrorDescription)
								Response.Write "</SELECT>&nbsp;"
								Response.Write "</TD>"
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!-- SelectItemByValue('-1', false, document.SpecialJourneyFrm.WorkingHours) //--></SCRIPT>"
							Response.Write "</TR>"
						End If
						Response.Write "<TR NAME=""SpecialJourneyFrm_ShiftIDDiv"" ID=""SpecialJourneyFrm_ShiftIDDiv""><INPUT TYPE=""HIDDEN"" NAME=""ShiftID"" ID=""ShiftIDHdn"" VALUE=""-1"" /></TR>"
						'Response.Write "<TR NAME=""SpecialJourneyFrm_RiskLevelIDDiv"" ID=""SpecialJourneyFrm_RiskLevelIDDiv""><INPUT TYPE=""HIDDEN"" NAME=""RiskLevelID"" ID=""RiskLevelIDHdn"" VALUE=""-1"" /></TR>"
						Response.Write "<TR NAME=""SpecialJourneyFrm_SpecialJourneyIDDiv"" ID=""SpecialJourneyFrm_SpecialJourneyIDDiv""><INPUT TYPE=""HIDDEN"" NAME=""SpecialJourneyID"" ID=""SpecialJourneyIDHdn"" VALUE=""-1"" /></TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Folio de autorización:&nbsp;</FONT></TD>"
							Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""DocumentNumber"" ID=""DocumentNumberTxt"" VALUE=""" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_DOCUMENTNUMBER) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
				        Response.Write "<TR>"
					        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
					        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE), "StartDate", Year(Date())-1, Year(Date()), True, False) & "</FONT></TD>"
				        Response.Write "</TR>"
				        Response.Write "<TR>"
					        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin:&nbsp;</FONT></TD>"
					        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE), "EndDate", Year(Date())-1, Year(Date())+2, True, False) & "</FONT></TD>"
				        Response.Write "</TR>"

					    If Len(oRequest("SpecialJourneyChange").Item) > 0 Then
						    Response.Write "<TR NAME=""JourneyIDDiv"" ID=""JourneyIDDiv"">"
							    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Turno:&nbsp;</NOBR></FONT></TD>"
							    Response.Write "<TD><SELECT NAME=""JourneyID"" ID=""JourneyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
								    Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SpecialJourneys", "JourneyID", "JourneyShortName, JourneyName", "(RecordTypeID In (-1," & CInt(Request.Cookies("SIAP_SubSectionID")) & ")) And (Active=1)", "JourneyShortName", "", "", sErrorDescription)
							    Response.Write "</SELECT>&nbsp;"
							    Response.Write "</TD>"
						    Response.Write "</TR>"
					    Else
							If CInt(Request.Cookies("SIAP_SubSectionID")) = 424 Then 'Suplencias
								Response.Write "<TR NAME=""JourneyIDDiv"" ID=""JourneyIDDiv"">"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Turno:&nbsp;</NOBR></FONT></TD>"
									Response.Write "<TD><SELECT NAME=""JourneyID"" ID=""JourneyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SpecialJourneys", "JourneyID", "JourneyShortName, JourneyName", "(RecordTypeID In (-1," & CInt(Request.Cookies("SIAP_SubSectionID")) & ")) And (Active=1)", "JourneyShortName", "", "", sErrorDescription)
									Response.Write "</SELECT>&nbsp;"
									Response.Write "</TD>"
								Response.Write "</TR>"
							End If
					    End If
						If CInt(Request.Cookies("SIAP_SubSectionID")) = 424 Then 'Suplencias
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Días / horas reportadas:&nbsp;</FONT></TD>"
								'Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""WorkedHours"" ID=""WorkedHoursTxt"" VALUE=""" & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_WORKED_HOURS) & """ SIZE=""5"" MAXLENGTH=""5"" CLASS=""TextFields"" onFocus=""DoBlock();"" onChange=""document.SpecialJourneyFrm.ReportedHours.value='';"" />"
								Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""WorkedHours"" ID=""WorkedHoursTxt"" VALUE=""" & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_WORKED_HOURS) & """ SIZE=""5"" MAXLENGTH=""5"" CLASS=""TextFields"" onChange=""document.SpecialJourneyFrm.ReportedHours.value='';"" />"
								Response.Write "<A HREF=""javascript: SearchRecord(document.SpecialJourneyFrm.EmployeeID.value, 'RecordsForGyS&TheRecordID=' + document.SpecialJourneyFrm.RecordID.value + '&RecordType=2&OriginalEmployeeID=' + document.SpecialJourneyFrm.OriginalEmployeeID.value + '&JourneyID=' + document.SpecialJourneyFrm.JourneyID.value + '&PositionID=' + document.SpecialJourneyFrm.PositionID.value + '&AreaID=' + document.SpecialJourneyFrm.AreaID.value + '&RiskLevelID=' + document.SpecialJourneyFrm.RiskLevelID.value + '&MovementID=' + document.SpecialJourneyFrm.MovementID.value + '&StartDate=' + BuildDateString(document.SpecialJourneyFrm.StartDateYear.value, document.SpecialJourneyFrm.StartDateMonth.value, document.SpecialJourneyFrm.StartDateDay.value) + '&EndDate=' + BuildDateString(document.SpecialJourneyFrm.EndDateYear.value, document.SpecialJourneyFrm.EndDateMonth.value, document.SpecialJourneyFrm.EndDateDay.value) + '&WorkingHours=' + document.SpecialJourneyFrm.WorkingHours.value + '&WorkedHours=' + document.SpecialJourneyFrm.WorkedHours.value + '&PayrollDate=' + document.SpecialJourneyFrm.AppliedDate.value, 'SearchEmployeeJourneysIFrame', 'SpecialJourneyFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar las horas reportadas por quincena para el empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportedHours"" ID=""ReportedHoursHdn"" VALUE=""""  />"
							Response.Write "</TD></TR>"
						Else
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Días / horas reportadas:&nbsp;</FONT></TD>"
								Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""WorkedHours"" ID=""WorkedHoursTxt"" VALUE=""" & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_WORKED_HOURS) & """ SIZE=""5"" MAXLENGTH=""5"" CLASS=""TextFields"" onChange=""document.SpecialJourneyFrm.ReportedHours.value='';"" />"
								Response.Write "<A HREF=""javascript: SearchRecord(document.SpecialJourneyFrm.EmployeeID.value, 'RecordsForGyS&TheRecordID=' + document.SpecialJourneyFrm.RecordID.value + '&RecordType=1&OriginalEmployeeID=' + document.SpecialJourneyFrm.OriginalEmployeeID.value + '&PositionID=' + document.SpecialJourneyFrm.PositionID.value + '&AreaID=' + document.SpecialJourneyFrm.AreaID.value + '&RiskLevelID=' + document.SpecialJourneyFrm.RiskLevelID.value + '&MovementID=' + document.SpecialJourneyFrm.MovementID.value + '&StartDate=' + BuildDateString(document.SpecialJourneyFrm.StartDateYear.value, document.SpecialJourneyFrm.StartDateMonth.value, document.SpecialJourneyFrm.StartDateDay.value) + '&EndDate=' + BuildDateString(document.SpecialJourneyFrm.EndDateYear.value, document.SpecialJourneyFrm.EndDateMonth.value, document.SpecialJourneyFrm.EndDateDay.value) + '&WorkingHours=' + document.SpecialJourneyFrm.WorkingHours.value + '&WorkedHours=' + document.SpecialJourneyFrm.WorkedHours.value + '&PayrollDate=' + document.SpecialJourneyFrm.AppliedDate.value, 'SearchEmployeeJourneysIFrame', 'SpecialJourneyFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar las horas reportadas por quincena para el empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportedHours"" ID=""ReportedHoursHdn"" VALUE=""""  />"
							Response.Write "</TD></TR>"
						End If
					    If Len(oRequest("SpecialJourneyChange").Item) > 0 Then
						    Response.Write "<TR NAME=""MovementIDDiv"" ID=""MovementIDDiv"">"
							    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Movimiento:&nbsp;</NOBR></FONT></TD>"
							    Response.Write "<TD><SELECT NAME=""MovementID"" ID=""MovementIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""document.SpecialJourneyFrm.ReportedHours.value=''; "">"
								    Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SpecialJourneysMovements", "MovementID", "MovementShortName, MovementName", "(RecordTypeID In (-1," & CInt(Request.Cookies("SIAP_SubSectionID")) & ")) And (Active=1)", "MovementShortName", "", "", sErrorDescription)
							    Response.Write "</SELECT>&nbsp;"
							    Response.Write "</TD>"
						    Response.Write "</TR>"
					    Else
						    Response.Write "<TR NAME=""MovementIDDiv"" ID=""MovementIDDiv"">"
							    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Movimiento:&nbsp;</NOBR></FONT></TD>"
							    Response.Write "<TD><SELECT NAME=""MovementID"" ID=""MovementIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""document.SpecialJourneyFrm.ReportedHours.value=''; "">"
                                    Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SpecialJourneysMovements", "MovementID", "MovementShortName, MovementName", "(RecordTypeID In (-1," & CInt(Request.Cookies("SIAP_SubSectionID")) & ")) And (Active=1)", "MovementShortName", "", "", sErrorDescription)
							    Response.Write "</SELECT>&nbsp;"
							    Response.Write "</TD>"
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!-- SelectItemByValue('-1', false, document.SpecialJourneyFrm.MovementID) //--></SCRIPT>"
						    Response.Write "</TR>"
					    End If
					    If Len(oRequest("SpecialJourneyChange").Item) > 0 Then
						    Response.Write "<TR NAME=""ReasonIDDiv"" ID=""ReasonIDDiv"">"
							    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Motivo:&nbsp;</NOBR></FONT></TD>"
							    Response.Write "<TD><SELECT NAME=""ReasonID"" ID=""ReasonIDCmb"" SIZE=""1"" CLASS=""Lists"">"
									If CInt(Request.Cookies("SIAP_SubSectionID")) = 424 Then
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SpecialJourneysReasons", "ReasonID", "ReasonShortName, ReasonName", "(RecordTypeID In (-1," & CInt(Request.Cookies("SIAP_SubSectionID")) & ")) And (Active=1)", "ReasonShortName", "", "Ninguno;;;-1", sErrorDescription)
									Else
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SpecialJourneysReasons", "ReasonID", "ReasonShortName, ReasonName", "(RecordTypeID In (-1," & CInt(Request.Cookies("SIAP_SubSectionID")) & ")) And (Active=1)", "ReasonShortName", "", "Ninguno;;;-1", sErrorDescription)
									End If
							    Response.Write "</SELECT>&nbsp;"
							    Response.Write "</TD>"
						    Response.Write "</TR>"
					    Else
						    Response.Write "<TR NAME=""ReasonIDDiv"" ID=""ReasonIDDiv"">"
							    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Motivo:&nbsp;</NOBR></FONT></TD>"
							    Response.Write "<TD><SELECT NAME=""ReasonID"" ID=""ReasonIDCmb"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SpecialJourneysReasons", "ReasonID", "ReasonShortName, ReasonName", "(RecordTypeID In (-1," & CInt(Request.Cookies("SIAP_SubSectionID")) & ")) And (Active=1)", "ReasonShortName", "", "Ninguno;;;-1", sErrorDescription)
							    Response.Write "</SELECT>&nbsp;"
							    Response.Write "</TD>"
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!-- SelectItemByValue('-1', false, document.SpecialJourneyFrm.ReasonID) //--></SCRIPT>"
						    Response.Write "</TR>"
					    End If
						Response.Write "<TR NAME=""RiskLevelIDDiv"" ID=""RiskLevelIDDiv"">"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Riesgos Profesionales:&nbsp;</NOBR></FONT></TD>"
							Response.Write "<TD><SELECT NAME=""RiskLevelID"" ID=""RiskLevelIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""document.SpecialJourneyFrm.ReportedHours.value='';"">"
								Response.Write "<OPTION VALUE=""0"">NA</OPTION>"
								Response.Write "<OPTION VALUE=""1"">&</OPTION>"
								Response.Write "<OPTION VALUE=""2"">*</OPTION>"
							Response.Write "</SELECT>&nbsp;"
							Response.Write "</TD>"
						Response.Write "</TR>"
                        Response.Write "<TR>" & vbNewLine
                            Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2""><BR />Comentarios:<BR /></FONT>"
                            Response.Write "<TEXTAREA NAME=""Comments"" ID=""CommentsTxtArea"" ROWS=""5"" COLS=""50"" MAXLENGTH=""2000"" CLASS=""TextFields""></TEXTAREA>"
                            Response.Write "</TD>"
                        Response.Write "</TR>" & vbNewLine
                        Response.Write "<TR NAME=""SpecialJourneyFrm_ConceptAmountDiv"" ID=""SpecialJourneyFrm_ConceptAmountDiv"">"
                            Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><NOBR>Percepción:&nbsp;</NOBR></FONT></TD>"
                            Response.Write "<TD VALIGN=""TOP""><INPUT TYPE=""TEXT"" NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""0.00"" CLASS=""TextFields"" onFocus=""DoBlock();"" /></TD>"
                        Response.Write "</TR>" & vbNewLine
					    If Len(oRequest("AbsenceChange").Item) > 0 Then
						    Response.Write "<TR><TD><BR /></TD></TR>"
						    Response.Write "<TR NAME=""AppliedRemoveDateDiv"" ID=""AppliedRemoveDateDiv"">"
							    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Quincena de aplicación de la justificación:&nbsp;</NOBR></FONT></TD>"
							    Response.Write "<TD><SELECT NAME=""AppliedRemoveDate"" ID=""AppliedRemoveDate"" SIZE=""1"" CLASS=""Lists"">"
								    Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "((IsClosed<>1) And (IsActive_2=1) And (PayrollTypeID=1)) Or (PayrollID=" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE) & ")", "PayrollID Desc", aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
							    Response.Write "</SELECT>&nbsp;"
							    Response.Write "</TD>"
						    Response.Write "</TR>"
					    Else
						    Response.Write "<TR NAME=""AppliedDateDiv"" ID=""AppliedDateDiv"">"
							    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Quincena de aplicación:&nbsp;</NOBR></FONT></TD>"
							    Response.Write "<TD><SELECT NAME=""AppliedDate"" ID=""AppliedDate"" SIZE=""1"" CLASS=""Lists"">"
								    If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
									    Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_5=1) And (PayrollTypeID=5)", "PayrollID Desc", aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
								    Else
									    Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_5=1) And (PayrollTypeID=5)", "PayrollID Desc", aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
								    End If
							    Response.Write "</SELECT>&nbsp;"
							    Response.Write "</TD>"
								Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!-- SelectItemByValue('-1', false, document.SpecialJourneyFrm.AppliedDate) //--></SCRIPT>"
						    Response.Write "</TR>"
					    End If
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TempStartDate"" ID=""TempStartDateHdn"" VALUE="""" />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TempEndDate"" ID=""TempEndDateHdn"" VALUE="""" />"
					Response.Write "</TABLE>"
					Response.Write "</BR>"
					'Response.Write "<DIV NAME=""ReasonsDiv"" ID=""ReasonsDiv"" STYLE=""display: none"">"
					'	Response.Write "<TEXTAREA NAME=""Reasons"" ID=""ReasonsTxtArea"" ROWS=""5"" COLS=""50"" MAXLENGTH=""2000"" VALUE="""" CLASS=""TextFields"">" & aSpecialJourneyComponent(S_REASONS_ABSENCE) & "</TEXTAREA>"
					'Response.Write "</DIV>"
                    Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
                        Response.Write "<TR>"
                            Response.Write "<TD>"
                                If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
                            Response.Write "</TD>"
                        Response.Write "</TR>"
                    Response.Write "</TABLE>"
                Response.Write "</FORM>"
				Response.Write "</TD><TD>&nbsp;&nbsp;&nbsp;</TD><TD VALIGN=""TOP"">"
						Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""250""></IFRAME><BR />"
						Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeJourneysIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""150""></IFRAME>"
				Response.Write "</TD></TR></TABLE>"

				Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
					Response.Write "document.SpecialJourneyFrm.CheckEmployeeID.value = '" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & "';" & vbNewLine
					Response.Write "document.SpecialJourneyFrm.OriginalEmployeeID.value = '" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(8) & "';" & vbNewLine
					Response.Write "document.SpecialJourneyFrm.ReportedHours.value = '1';" & vbNewLine
					Response.Write "SearchRecord(document.SpecialJourneyFrm.EmployeeNumber.value, 'EmployeesGyS&RecordType=1&AreaID=-1&RecordDate=' + document.SpecialJourneyFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'SpecialJourneyFrm');" & vbNewLine
					Response.Write "SearchRecord('1', 'PositionsGyS&PositionID=' + document.SpecialJourneyFrm.PositionID.value + '&AreaID=' + document.SpecialJourneyFrm.AreaID.value + '&ServiceID=' + document.SpecialJourneyFrm.ServiceID.value + '&LevelID=' + document.SpecialJourneyFrm.LevelID.value + '&WorkingHours=' + document.SpecialJourneyFrm.WorkingHours.value + '&RecordType=1&RecordDate=' + document.SpecialJourneyFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'SpecialJourneyFrm');" & vbNewLine
					'Response.Write "document.SpecialJourneyFrm.RFC.focus();" & vbNewLine
				Response.Write "//--></SCRIPT>" & vbNewLine
            End If
        End If
	End If

	DisplayExternalSpecialJourneyForm = lErrorNumber
	Err.Clear
End Function

Function DisplayInternalSpecialJourneyForm(oRequest, oADODBConnection, aSpecialJourneyComponent, iSpecialJourneyType, sErrorDescription)
'************************************************************
'Purpose: To display the information about an absence for the
'         employee from the database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, sAbsenceIDs, sExtraURL, aSpecialJourneyComponent
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayInternalSpecialJourneyForm"
	Dim sNames
	Dim aRelatedAbsences
	Dim iIndex
	Dim oRecordset
	Dim lErrorNumber
	Dim sAbsenceIDs
	Dim sCaseOptions
	Dim sIDs

	If lErrorNumber = 0 Then
		If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR><TD VALIGN=""TOP"">"
		End If
        Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			If aEmployeeComponent(N_ID_EMPLOYEE) <> -1 Then
				Response.Write "var bBlock = true;" & vbNewLine
				Response.Write "function DoBlock() {" & vbNewLine
					Response.Write "if (bBlock) {" & vbNewLine
						Response.Write "document.SpecialJourneyFrm.DocumentNumber.focus();" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "} // End of DoBlock" & vbNewLine
			End If
			Response.Write "function CheckSpecialJourneyFields(oForm) {" & vbNewLine
                Response.Write "if (oForm) {" & vbNewLine
                    If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
						Response.Write "if (oForm.EmployeeID.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir un valor para el campo \'No. del empleado\'.');" & vbNewLine
							Response.Write "oForm.EmployeeID.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
                    Else
						If CInt(Request.Cookies("SIAP_SubSectionID")) = 424 Then
							Response.Write "if (oForm.OriginalEmployeeID.value == '') {" & vbNewLine
								Response.Write "alert('Favor de especificar y validar el número del empleado a suplir.');" & vbNewLine
								Response.Write "oForm.OriginalEmployeeID.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
							Response.Write "if ((oForm.CheckOriginalEmployeeID.value == '') || (oForm.CheckOriginalEmployeeID.value == '-1')) {" & vbNewLine
								Response.Write "alert('Favor de validar el número del empleado a suplir.');" & vbNewLine
								Response.Write "oForm.OriginalEmployeeID.focus();" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "}" & vbNewLine
						End If
						Response.Write "if (oForm.DocumentNumber.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir un valor para el campo \'Folio de autorización\'.');" & vbNewLine
							Response.Write "oForm.DocumentNumber.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "oForm.WorkedHours.value = oForm.WorkedHours.value.replace(/,/gi, '');" & vbNewLine
						Response.Write "if (!CheckFloatValue(oForm.WorkedHours, 'el campo \'Días/horas reportadas\'', N_MINIMUM_ONLY_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "oForm.ConceptAmount.value = oForm.ConceptAmount.value.replace(/,/gi, '');" & vbNewLine
						Response.Write "if (!CheckFloatValue(oForm.ConceptAmount, 'el campo \'Percepción\'', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
                    End If
                Response.Write "}" & vbNewLine
                'Response.Write "return ValidateSpecialJourneyFields(window.document.SpecialJourneyFrm);" & vbNewLine
                Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckSpecialJourneyFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine

		If aEmployeeComponent(N_ID_EMPLOYEE) = -1 Then
			Response.Write "<FORM NAME=""SpecialJourneyFrm"" ID=""SpecialJourneyFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""GET"" onSubmit=""return CheckSpecialJourneyFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""SpecialJourney"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SpecialJourneyType"" ID=""SpecialJourneyTypeHdn"" VALUE=""" & iSpecialJourneyType & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SpecialJourneyID"" ID=""SpecialJourneyIDHdn"" VALUE=""" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SPECIAL_JOURNEY_ID) & """ />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">ID del Empleado:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"
			Response.Write "<BR />"
			If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then
				Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""SpecialJourneyMovement"" ID=""SpecialJourneyMovementBtn"" VALUE=""Buscar empleado"" CLASS=""Buttons"" />"
			End If
			Response.Write "</FORM>"
		Else
			lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
			If lErrorNumber = 0 Then
				lErrorNumber = GetEmployee(oRequest, oADODBConnection, aEmployeeComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
						Response.Write "function CheckAnotherSpecialJourneyFields(oForm) {" & vbNewLine
							Response.Write "if (oForm) {" & vbNewLine
								If iSpecialJourneyType = 1 Then
									Response.Write "if (oForm.EmployeeID.value.length == 0) {" & vbNewLine
										Response.Write "alert('Favor de introducir un valor para el campo \'Número del empleado\'.');" & vbNewLine
										Response.Write "oForm.EmployeeID.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
								Else
									Response.Write "if (oForm.RFC.value.length == 0) {" & vbNewLine
										Response.Write "alert('Favor de introducir un valor para el campo \'RFC\'.');" & vbNewLine
										Response.Write "oForm.RFC.focus();" & vbNewLine
										Response.Write "return false;" & vbNewLine
									Response.Write "}" & vbNewLine
								End If
							Response.Write "}" & vbNewLine
						Response.Write "} // End of CheckAnotherSpecialJourneyFields" & vbNewLine
					Response.Write "//--></SCRIPT>" & vbNewLine
					If CInt(Request.Cookies("SIAP_SectionID")) <> 1 Then
						Response.Write "<FORM NAME=""AnotherSpecialJourneyFrm"" ID=""AnotherSpecialJourneyFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""GET"" onSubmit=""return CheckAnotherSpecialJourneyFields(this)"">"
							Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
								Response.Write "<TR>"
									Response.Write "<TD VALIGN=""TOP"" WIDTH=""500"">&nbsp;</TD>"
									Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><IMG SRC=""Images/MnLeftArrows.gif"" WIDTH=""32"" HEIGHT=""32"" ALT=""Guardias"" BORDER=""0"" /><BR /></TD>"
									Response.Write "<TD VALIGN=""TOP"" WIDTH=""290""><FONT FACE=""Arial"" SIZE=""2""><B>Otro empleado</B><BR /></FONT>"
									Response.Write "<DIV CLASS=""MenuOverflow""><FONT FACE=""Arial"" SIZE=""2"">Registre guardias a un empleado diferente.</FONT></DIV></TD>"
								Response.Write "</TR>"
								Response.Write "<TR>"
									Response.Write "<TD VALIGN=""TOP"" WIDTH=""500"">&nbsp;</TD>"
									Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;</FONT></TD>"
									Response.Write "<TD VALIGN=""TOP"" WIDTH=""290""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;Número del empleado:&nbsp;</FONT><INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" /></TD>"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""SpecialJourney"" />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SpecialJourneyType"" ID=""SpecialJourneyTypeHdn"" VALUE=""" & iSpecialJourneyType & """ />"
									Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SpecialJourneyID"" ID=""SpecialJourneyIDHdn"" VALUE=""" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SPECIAL_JOURNEY_ID) & """ />"
								Response.Write "</TR>"
								Response.Write "<TR>"
									Response.Write "<TD VALIGN=""TOP"" WIDTH=""500"">&nbsp;</TD>"
									Response.Write "<TD VALIGN=""TOP"" WIDTH=""32""><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;</FONT></TD>"
									Response.Write "<TD VALIGN=""TOP"" WIDTH=""290""><INPUT TYPE=""SUBMIT"" NAME=""SpecialJourneyMovement"" ID=""SpecialJourneyMovementBtn"" VALUE=""Buscar empleado"" CLASS=""Buttons"" /></TD>"
								Response.Write "</TR>"
							Response.Write "</TABLE>"
						Response.Write "</FORM>"
					End If

					Response.Write "<FORM NAME=""SpecialJourneyFrm"" ID=""SpecialJourneyFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""GET"" onSubmit=""return CheckSpecialJourneyFields(this)"">"
					Response.Write "<BR /><BR />"
					If Len(oRequest("SpecialJourneyChange").Item) > 0 Then
						lErrorNumber = GetSpecialJourney(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SpecialJourneyChange"" ID=""SpecialJourneyChangeHdn"" VALUE=""" & oRequest("SpecialJourneyChange").Item & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RemoveUserID"" ID=""RemoveUserIDHdn"" VALUE=""" & aLoginComponent(N_USER_ID_LOGIN) & """ />"
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RemovedDate"" ID=""RemovedDateHdn"" VALUE=""" & Left(GetSerialNumberForDate(""), Len("00000000")) & """ />"
						aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE) = 0
					End If
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""SpecialJourney"" />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SpecialJourneyType"" ID=""SpecialJourneyTypeHdn"" VALUE=""" & iSpecialJourneyType & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SpecialJourneyID"" ID=""SpecialJourneyIDHdn"" VALUE=""" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SPECIAL_JOURNEY_ID) & """ />"
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RecordID"" ID=""RecordIDHdn"" VALUE=""-1"" />"
					If CInt(Request.Cookies("SIAP_SubSectionID")) <> 424 Then ' Si es Guardia
						Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OriginalEmployeeID"" ID=""OriginalEmployeeIDHdn"" VALUE=""-1"" />"
					End If
					'Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RiskLevelID"" ID=""RiskLevelIDHdn"" VALUE=""-1"" />"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Información general</B></FONT>"
					Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
					Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & "</FONT><INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aEmployeeComponent(S_NUMBER_EMPLOYEE) & """ /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre:&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aEmployeeComponent(S_NAME_EMPLOYEE) & " " & aEmployeeComponent(S_LAST_NAME_EMPLOYEE) & " " & aEmployeeComponent(S_LAST_NAME2_EMPLOYEE)) & "</FONT></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">RFC:&nbsp;</FONT></TD>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aEmployeeComponent(S_RFC_EMPLOYEE)) & "</FONT></TD>"
						Response.Write "</TR>"
						lErrorNumber = GetJob(oRequest, oADODBConnection, aJobComponent, sErrorDescription)
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Adscripción&nbsp;</FONT></TD>"
							Call GetNameFromTable(oADODBConnection, "Areas", aJobComponent(N_AREA_ID_JOB), "", "", sNames, sErrorDescription)
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT><INPUT TYPE=""HIDDEN"" NAME=""AreaID"" ID=""AreaIDHdn"" VALUE=""" & aJobComponent(N_AREA_ID_JOB) & """ /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Servicio:&nbsp;</FONT></TD>"
							Call GetNameFromTable(oADODBConnection, "Services", aJobComponent(N_SERVICE_ID_JOB), "", "", sNames, sErrorDescription)
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT><INPUT TYPE=""HIDDEN"" NAME=""ServiceID"" ID=""ServiceIDHdn"" VALUE=""" & aJobComponent(N_SERVICE_ID_JOB) & """ /></TD>"
						Response.Write "</TR>"
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Puesto:&nbsp;</FONT></TD>"
								sIDs = aJobComponent(N_POSITION_ID_JOB)
								Call GetNameFromTable(oADODBConnection, "Positions", sIDs, "", "", sNames, sErrorDescription)
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT><INPUT TYPE=""HIDDEN"" NAME=""PositionID"" ID=""PositionIDHdn"" VALUE=""" & aJobComponent(N_POSITION_ID_JOB) & """ /></TD>"
						Response.Write "</TR>"
					    Response.Write "<TR>"
						    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horario:&nbsp;</FONT></TD>"
						    Call GetNameFromTable(oADODBConnection, "Shifts", aEmployeeComponent(N_SHIFT_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
						    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & sNames & "</FONT><INPUT TYPE=""HIDDEN"" NAME=""ShiftID"" ID=""ShiftIDHdn"" VALUE=""" & aEmployeeComponent(N_SHIFT_ID_EMPLOYEE) & """ /></TD>"
					    Response.Write "</TR>"
			            Response.Write "<TR>"
				            Call GetNameFromTable(oADODBConnection, "Levels", aEmployeeComponent(N_LEVEL_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
				            Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nivel:&nbsp;</FONT></TD>"
				            Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT><INPUT TYPE=""HIDDEN"" NAME=""LevelID"" ID=""LevelIDHdn"" VALUE=""" & aEmployeeComponent(N_LEVEL_ID_EMPLOYEE) & """ /></TD>"
			            Response.Write "</TR>"
					    Response.Write "<TR>"
						    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Horas laboradas:&nbsp;</FONT></TD>"
						    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & "</FONT><INPUT TYPE=""HIDDEN"" NAME=""WorkingHours"" ID=""WorkingHoursHdn"" VALUE=""" & aEmployeeComponent(D_WORKING_HOURS_EMPLOYEE) & """ /></TD>"
					    Response.Write "</TR>"
						If CInt(Request.Cookies("SIAP_SubSectionID")) = 424 Then ' Si es Suplencia
							If CLng(Replace(oRequest("EmployeeID").Item, "'", "")) < 800000 Then ' Interno
								Response.Write "<TR>"
									Call GetNameFromTable(oADODBConnection, "Journeys", aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE), "", "", sNames, sErrorDescription)
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Turno:&nbsp;</FONT></TD>"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT><INPUT TYPE=""HIDDEN"" NAME=""JourneyID"" ID=""JourneyIDHdn"" VALUE=""" & aEmployeeComponent(N_JOURNEY_ID_EMPLOYEE) & """ /></TD>"
								Response.Write "</TR>"
							End If
						End If
                        Response.Write "</TABLE>"

                        Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
                        Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
						If CInt(Request.Cookies("SIAP_SubSectionID")) = 424 Then ' Si es Suplencia
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">No. del empleado a suplir:&nbsp;</FONT></TD>"
								Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""OriginalEmployeeID"" ID=""OriginalEmployeeIDTxt"" VALUE=""" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ORIGINAL_EMPLOYEE_ID) & """ SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" onChange=""document.SpecialJourneyFrm.CheckOriginalEmployeeID.value=''; document.SpecialJourneyFrm.ReportedHours.value='';"" />"
								Response.Write "<A HREF=""javascript: if (parseInt(document.SpecialJourneyFrm.EmployeeID.value) == parseInt(document.SpecialJourneyFrm.OriginalEmployeeID.value)) {alert('El empleado suplido no puede ser el mismo que el empleado suplente.'); document.SpecialJourneyFrm.OriginalEmployeeID.focus();} else {SearchRecord(document.SpecialJourneyFrm.OriginalEmployeeID.value, 'EmployeesGyS&RecordType=2&Original=1&EmployeeID=' + document.SpecialJourneyFrm.EmployeeID.value + '&RecordDate=' + document.SpecialJourneyFrm.AppliedDate.value, 'SearchEmployeeNumberIFrame', 'SpecialJourneyFrm');}""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado a suplir"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CheckOriginalEmployeeID"" ID=""CheckOriginalEmployeeIDHdn"" VALUE=""""  />"
							Response.Write "</TD></TR>"
						End If
						Response.Write "<TR>"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Folio de autorización:&nbsp;</FONT></TD>"
							Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""DocumentNumber"" ID=""DocumentNumberTxt"" VALUE=""" & aSpecialJourneyComponent(S_SPECIAL_JOURNEY_DOCUMENTNUMBER) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
						Response.Write "</TR>"
				        Response.Write "<TR>"
					        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
					        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_STARTDATE), "StartDate", Year(Date())-1, Year(Date()), True, False) & "</FONT></TD>"
				        Response.Write "</TR>"
				        Response.Write "<TR>"
					        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de fin:&nbsp;</FONT></TD>"
					        Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ENDDATE), "EndDate", Year(Date())-1, Year(Date())+2, True, False) & "</FONT></TD>"
				        Response.Write "</TR>"

						If CInt(Request.Cookies("SIAP_SubSectionID")) <> 424 Then
							If Len(oRequest("SpecialJourneyChange").Item) > 0 Then
								Response.Write "<TR NAME=""JourneyIDDiv"" ID=""JourneyIDDiv"">"
									Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Turno:&nbsp;</NOBR></FONT></TD>"
									Response.Write "<TD><SELECT NAME=""JourneyID"" ID=""JourneyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
										Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SpecialJourneys", "JourneyID", "JourneyShortName, JourneyName", "(RecordTypeID In (-1," & CInt(Request.Cookies("SIAP_SubSectionID")) & ")) And (Active=1)", "JourneyShortName", "", "", sErrorDescription)
									Response.Write "</SELECT>&nbsp;"
									Response.Write "</TD>"
								Response.Write "</TR>"
							Else
								If CInt(Request.Cookies("SIAP_SubSectionID")) = 424 Then ' Si es Suplencia
									If CLng(Replace(oRequest("EmployeeID").Item, "'", "")) >= 800000 Then ' Interno
										Response.Write "<TR NAME=""JourneyIDDiv"" ID=""JourneyIDDiv"">"
											Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Turno:&nbsp;</NOBR></FONT></TD>"
											Response.Write "<TD><SELECT NAME=""JourneyID"" ID=""JourneyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
												Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SpecialJourneys", "JourneyID", "JourneyShortName, JourneyName", "(RecordTypeID In (-1," & CInt(Request.Cookies("SIAP_SubSectionID")) & ")) And (Active=1)", "JourneyShortName", "", "", sErrorDescription)
											Response.Write "</SELECT>&nbsp;"
											Response.Write "</TD>"
										Response.Write "</TR>"
									End If
								End If
							End If
						End If
						If CInt(Request.Cookies("SIAP_SubSectionID")) = 424 Then ' Si es Suplencia
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Días / horas reportadas:&nbsp;</FONT></TD>"
								'Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""WorkedHours"" ID=""WorkedHoursTxt"" VALUE=""" & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_WORKED_HOURS) & """ SIZE=""5"" MAXLENGTH=""5"" CLASS=""TextFields"" onFocus=""DoBlock();"" onChange=""document.SpecialJourneyFrm.ReportedHours.value='';"" />"
								Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""WorkedHours"" ID=""WorkedHoursTxt"" VALUE=""" & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_WORKED_HOURS) & """ SIZE=""5"" MAXLENGTH=""5"" CLASS=""TextFields"" onChange=""document.SpecialJourneyFrm.ReportedHours.value='';"" />"
								Response.Write "<A HREF=""javascript: SearchRecord(document.SpecialJourneyFrm.EmployeeID.value, 'RecordsForGyS&TheRecordID=' + document.SpecialJourneyFrm.RecordID.value + '&RecordType=2&OriginalEmployeeID=' + document.SpecialJourneyFrm.OriginalEmployeeID.value + '&JourneyID=' + document.SpecialJourneyFrm.JourneyID.value + '&PositionID=' + document.SpecialJourneyFrm.PositionID.value + '&AreaID=' + document.SpecialJourneyFrm.AreaID.value + '&RiskLevelID=' + document.SpecialJourneyFrm.RiskLevelID.value + '&MovementID=' + document.SpecialJourneyFrm.MovementID.value + '&StartDate=' + BuildDateString(document.SpecialJourneyFrm.StartDateYear.value, document.SpecialJourneyFrm.StartDateMonth.value, document.SpecialJourneyFrm.StartDateDay.value) + '&EndDate=' + BuildDateString(document.SpecialJourneyFrm.EndDateYear.value, document.SpecialJourneyFrm.EndDateMonth.value, document.SpecialJourneyFrm.EndDateDay.value) + '&WorkingHours=' + document.SpecialJourneyFrm.WorkingHours.value + '&WorkedHours=' + document.SpecialJourneyFrm.WorkedHours.value + '&PayrollDate=' + document.SpecialJourneyFrm.AppliedDate.value, 'SearchEmployeeJourneysIFrame', 'SpecialJourneyFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar las horas reportadas por quincena para el empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportedHours"" ID=""ReportedHoursHdn"" VALUE=""""  />"
							Response.Write "</TD></TR>"
						Else
							Response.Write "<TR>"
								Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Días / horas reportadas:&nbsp;</FONT></TD>"
								Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""WorkedHours"" ID=""WorkedHoursTxt"" VALUE=""" & aSpecialJourneyComponent(D_SPECIAL_JOURNEY_WORKED_HOURS) & """ SIZE=""5"" MAXLENGTH=""5"" CLASS=""TextFields"" onChange=""document.SpecialJourneyFrm.ReportedHours.value='';"" />"
								Response.Write "<A HREF=""javascript: SearchRecord(document.SpecialJourneyFrm.EmployeeID.value, 'RecordsForGyS&TheRecordID=' + document.SpecialJourneyFrm.RecordID.value + '&RecordType=1&OriginalEmployeeID=' + document.SpecialJourneyFrm.OriginalEmployeeID.value + '&PositionID=' + document.SpecialJourneyFrm.PositionID.value + '&AreaID=' + document.SpecialJourneyFrm.AreaID.value + '&RiskLevelID=' + document.SpecialJourneyFrm.RiskLevelID.value + '&MovementID=' + document.SpecialJourneyFrm.MovementID.value + '&StartDate=' + BuildDateString(document.SpecialJourneyFrm.StartDateYear.value, document.SpecialJourneyFrm.StartDateMonth.value, document.SpecialJourneyFrm.StartDateDay.value) + '&EndDate=' + BuildDateString(document.SpecialJourneyFrm.EndDateYear.value, document.SpecialJourneyFrm.EndDateMonth.value, document.SpecialJourneyFrm.EndDateDay.value) + '&WorkingHours=' + document.SpecialJourneyFrm.WorkingHours.value + '&WorkedHours=' + document.SpecialJourneyFrm.WorkedHours.value + '&PayrollDate=' + document.SpecialJourneyFrm.AppliedDate.value, 'SearchEmployeeJourneysIFrame', 'SpecialJourneyFrm')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar las horas reportadas por quincena para el empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReportedHours"" ID=""ReportedHoursHdn"" VALUE=""""  />"
							Response.Write "</TD></TR>"
						End If
					    If Len(oRequest("SpecialJourneyChange").Item) > 0 Then
						    Response.Write "<TR NAME=""MovementIDDiv"" ID=""MovementIDDiv"">"
							    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Movimiento:&nbsp;</NOBR></FONT></TD>"
							    Response.Write "<TD><SELECT NAME=""MovementID"" ID=""MovementIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""document.SpecialJourneyFrm.ReportedHours.value='';"" >"
								    Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SpecialJourneysMovements", "MovementID", "MovementShortName, MovementName", "(RecordTypeID In (-1," & CInt(Request.Cookies("SIAP_SubSectionID")) & ")) And (Active=1)", "MovementShortName", "", "", sErrorDescription)
							    Response.Write "</SELECT>&nbsp;"
							    Response.Write "</TD>"
						    Response.Write "</TR>"
					    Else
						    Response.Write "<TR NAME=""MovementIDDiv"" ID=""MovementIDDiv"">"
							    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Movimiento:&nbsp;</NOBR></FONT></TD>"
							    Response.Write "<TD><SELECT NAME=""MovementID"" ID=""MovementIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""document.SpecialJourneyFrm.ReportedHours.value='';"" >"
                                    Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SpecialJourneysMovements", "MovementID", "MovementShortName, MovementName", "(RecordTypeID In (-1," & CInt(Request.Cookies("SIAP_SubSectionID")) & ")) And (Active=1)", "MovementShortName", "", "", sErrorDescription)
							    Response.Write "</SELECT>&nbsp;"
							    Response.Write "</TD>"
						    Response.Write "</TR>"
					    End If
					    If Len(oRequest("SpecialJourneyChange").Item) > 0 Then
						    Response.Write "<TR NAME=""ReasonIDDiv"" ID=""ReasonIDDiv"">"
							    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Motivo:&nbsp;</NOBR></FONT></TD>"
							    Response.Write "<TD><SELECT NAME=""ReasonID"" ID=""ReasonIDCmb"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SpecialJourneysReasons", "ReasonID", "ReasonShortName, ReasonName", "(RecordTypeID In (-1," & CInt(Request.Cookies("SIAP_SubSectionID")) & ")) And (Active=1)", "ReasonShortName", "", "Ninguno;;;-1", sErrorDescription)
							    Response.Write "</SELECT>&nbsp;"
							    Response.Write "</TD>"
						    Response.Write "</TR>"
					    Else
						    Response.Write "<TR NAME=""ReasonIDDiv"" ID=""ReasonIDDiv"">"
							    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Motivo:&nbsp;</NOBR></FONT></TD>"
							    Response.Write "<TD><SELECT NAME=""ReasonID"" ID=""ReasonIDCmb"" SIZE=""1"" CLASS=""Lists"">"
									Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SpecialJourneysReasons", "ReasonID", "ReasonShortName, ReasonName", "(RecordTypeID In (-1," & CInt(Request.Cookies("SIAP_SubSectionID")) & ")) And (Active=1)", "ReasonShortName", "", "Ninguno;;;-1", sErrorDescription)
							    Response.Write "</SELECT>&nbsp;"
							    Response.Write "</TD>"
						    Response.Write "</TR>"
					    End If
						Response.Write "<TR NAME=""RiskLevelIDDiv"" ID=""RiskLevelIDDiv"">"
							Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Riesgos Profesionales:&nbsp;</NOBR></FONT></TD>"
							Response.Write "<TD><SELECT NAME=""RiskLevelID"" ID=""RiskLevelIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""document.SpecialJourneyFrm.ReportedHours.value='';"">"
								Response.Write "<OPTION VALUE=""0"">NA</OPTION>"
								Response.Write "<OPTION VALUE=""1"">&</OPTION>"
								Response.Write "<OPTION VALUE=""2"">*</OPTION>"
							Response.Write "</SELECT>&nbsp;"
							Response.Write "</TD>"
						Response.Write "</TR>"
                            Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2""><BR />Comentarios:<BR /></FONT>"
                            Response.Write "<TEXTAREA NAME=""Comments"" ID=""CommentsTxtArea"" ROWS=""5"" COLS=""50"" MAXLENGTH=""2000"" CLASS=""TextFields""></TEXTAREA>"
                            Response.Write "</TD>"
                        Response.Write "</TR>" & vbNewLine
                        Response.Write "<TR NAME=""SpecialJourneyFrm_ConceptAmountDiv"" ID=""SpecialJourneyFrm_ConceptAmountDiv"">"
                            Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2""><NOBR>Percepción:&nbsp;</NOBR></FONT></TD>"
                            Response.Write "<TD VALIGN=""TOP""><INPUT TYPE=""TEXT"" NAME=""ConceptAmount"" ID=""ConceptAmountTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""0.00"" CLASS=""TextFields"" onFocus=""DoBlock();"" /></TD>"
                        Response.Write "</TR>" & vbNewLine
					    If Len(oRequest("AbsenceChange").Item) > 0 Then
						    Response.Write "<TR><TD><BR /></TD></TR>"
						    Response.Write "<TR NAME=""AppliedRemoveDateDiv"" ID=""AppliedRemoveDateDiv"">"
							    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Quincena de aplicación de la justificación:&nbsp;</NOBR></FONT></TD>"
							    Response.Write "<TD><SELECT NAME=""AppliedRemoveDate"" ID=""AppliedRemoveDate"" SIZE=""1"" CLASS=""Lists"">"
								    Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "((IsClosed<>1) And (IsActive_2=1) And (PayrollTypeID=5)) Or (PayrollID=" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE) & ")", "PayrollID Desc", aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
							    Response.Write "</SELECT>&nbsp;"
							    Response.Write "</TD>"
						    Response.Write "</TR>"
					    Else
						    Response.Write "<TR NAME=""AppliedDateDiv"" ID=""AppliedDateDiv"">"
							    Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><NOBR>Quincena de aplicación:&nbsp;</NOBR></FONT></TD>"
							    Response.Write "<TD><SELECT NAME=""AppliedDate"" ID=""AppliedDate"" SIZE=""1"" CLASS=""Lists"">"
								    If CInt(Request.Cookies("SIAP_SectionID")) = 7 Then
									    Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_5=1) And (PayrollTypeID=5)", "PayrollID Desc", aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
								    Else
									    Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(IsClosed<>1) And (IsActive_5=1) And (PayrollTypeID=5)", "PayrollID Desc", aSpecialJourneyComponent(N_SPECIAL_JOURNEY_APPLIED_DATE), "No existen nóminas abiertas para el registro de movimientos;;;-1", sErrorDescription)
								    End If
							    Response.Write "</SELECT>&nbsp;"
							    Response.Write "</TD>"
						    Response.Write "</TR>"
					    End If
					    Response.Write "</TABLE>"
					    Response.Write "<BR />"
					    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TempStartDate"" ID=""TempStartDateHdn"" VALUE="""" />"
					    Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""TempEndDate"" ID=""TempEndDateHdn"" VALUE="""" />"
					    Response.Write "<DIV NAME=""ReasonsDiv"" ID=""ReasonsDiv"" STYLE=""display: none"">"
						    Response.Write "<TEXTAREA NAME=""Reasons"" ID=""ReasonsTxtArea"" ROWS=""5"" COLS=""50"" MAXLENGTH=""2000"" VALUE="""" CLASS=""TextFields"">" & aSpecialJourneyComponent(S_REASONS_ABSENCE) & "</TEXTAREA>"
					    Response.Write "</DIV>"
                        Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
                            Response.Write "<TR>"
                                Response.Write "<TD>"
                                    If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
                                Response.Write "</TD>"
                            Response.Write "</TR>"
                        Response.Write "</TABLE>"

                    Response.Write "</FORM>"
					Response.Write "</TD><TD>&nbsp;&nbsp;&nbsp;</TD><TD VALIGN=""TOP"">"
						If CInt(Request.Cookies("SIAP_SubSectionID")) = 424 Then ' Es Guardia
							Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""250""></IFRAME><BR />"
						End If
						Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeJourneysIFrame"" FRAMEBORDER=""0"" WIDTH=""300"" HEIGHT=""150""></IFRAME>"
					Response.Write "</TD></TR></TABLE>"
				End If
            End If
        End If
	End If

	DisplayInternalSpecialJourneyForm = lErrorNumber
	Err.Clear
End Function

Function DisplaySpecialJourneyAsHiddenFields(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an absence for the
'         employee using hidden form fields
'Inputs:  oRequest, oADODBConnection, aSpecialJourneyComponent
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplaySpecialJourneyAsHiddenFields"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aSpecialJourneyComponent(N_EMPLOYEE_ID_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AbsenceID"" ID=""AbsenceIDHdn"" VALUE=""" & aSpecialJourneyComponent(N_ABSENCE_ID_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OcurredDate"" ID=""OcurredDateHdn"" VALUE=""" & aSpecialJourneyComponent(N_OCURRED_DATE_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EndDate"" ID=""EndDateHdn"" VALUE=""" & aSpecialJourneyComponent(N_END_DATE_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RegistrationDate"" ID=""RegistrationDateHdn"" VALUE=""" & aSpecialJourneyComponent(N_REGISTRATION_DATE_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DocumentNumber"" ID=""DocumentNumberHdn"" VALUE=""" & aSpecialJourneyComponent(S_DOCUMENT_NUMBER_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AbsenceHours"" ID=""AbsenceHoursHdn"" VALUE=""" & aSpecialJourneyComponent(N_HOURS_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""JustificationID"" ID=""JustificationIDHdn"" VALUE=""" & aSpecialJourneyComponent(N_JUSTIFICATION_ID_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AppliesForPunctuality"" ID=""AppliesForPunctualityHdn"" VALUE=""" & aSpecialJourneyComponent(N_APPLIES_FOR_PUNCTUALITY_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Reasons"" ID=""ReasonsHdn"" VALUE=""" & aSpecialJourneyComponent(S_REASONS_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AddUserID"" ID=""AddUserIDHdn"" VALUE=""" & aSpecialJourneyComponent(N_ADD_USER_ID_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AppliedDate"" ID=""AppliedDateHdn"" VALUE=""" & aSpecialJourneyComponent(N_APPLIED_DATE_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Removed"" ID=""RemovedHdn"" VALUE=""" & aSpecialJourneyComponent(N_REMOVED_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RemoveUserID"" ID=""RemoveUserIDHdn"" VALUE=""" & aSpecialJourneyComponent(N_REMOVE_USER_ID_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RemovedDate"" ID=""RemovedDateHdn"" VALUE=""" & aSpecialJourneyComponent(N_REMOVED_DATE_ABSENCE) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AppliedRemoveDate"" ID=""AppliedRemoveDateHdn"" VALUE=""" & aSpecialJourneyComponent(N_APPLIED_REMOVE_DATE_ABSENCE) & """ />"

	DisplaySpecialJourneyAsHiddenFields = Err.number
	Err.Clear
End Function

Function DisplayExternalSpecialJourneyTable(oRequest, oADODBConnection, bForExport, lStartPage, aSpecialJourneyComponent, iSpecialJourneyType, sErrorDescription)
'************************************************************
'Purpose: To display the absences for the given absence for
'		  the employee from the database in a table
'Inputs:  oRequest, oADODBConnection, bForExport, aSpecialJourneyComponent
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayExternalSpecialJourneyTable"
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

	oStartDate = Now()
	lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	lErrorNumber = GetExternalSpecialJourneys(oRequest, oADODBConnection, aSpecialJourneyComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetchForSections(oRequest, lStartPage, 10, aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE), oRecordset)
			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If (Not bForExport) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Or (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
					asColumnsTitles = Split("Acciones,Empleado,Folio de autorización,F. de inicio,F. de termino,Días,Nómina en que se aplica,Registró", ",", -1, vbBinaryCompare)
					asCellWidths = Split("200,100,200,200,200,100,200,200,200", ",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,,,,,CENTER,CENTER,,,,", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("Empleado,Folio de autorización,F. de inicio,F. de termino,Días,Nómina en que se aplica,Registró", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,200,200,200,100,200,200,200", ",", -1, vbBinaryCompare)
					asCellAlignments = Split(",,,,CENTER,CENTER,,,,", ",", -1, vbBinaryCompare)
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
					If (StrComp(CStr(oRecordset.Fields("RecordID").Value), oRequest("RecordID").Item, vbBinaryCompare) = 0) And (StrComp(CStr(oRecordset.Fields("StartDate").Value), oRequest("StartDate").Item, vbBinaryCompare) = 0) Then
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
						If CInt(oRecordset.Fields("Active").Value) = 1 Then
							sRowContents = sRowContents & "<A HREF=""" & "SpecialJourney.asp" & "?Action=SpecialJourney&SetDeActive=1&RecordID=" & CStr(oRecordset.Fields("RecordID").Value) & "&SpecialJourneyType=" & iSpecialJourneyType & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnDeactive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Desactivar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						Else
							sRowContents = sRowContents & "<A HREF=""" & "SpecialJourney.asp" & "?Action=SpecialJourney&SetActive=1&RecordID=" & CStr(oRecordset.Fields("RecordID").Value) & "&SpecialJourneyType=" & iSpecialJourneyType & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
							sRowContents = sRowContents & "<A HREF=""" & "SpecialJourney.asp" & "?Action=SpecialJourney&Remove=1&RecordID=" & CStr(oRecordset.Fields("RecordID").Value) & "&SpecialJourneyType=" & iSpecialJourneyType & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
                        sRowContents = sRowContents & sBoldEnd & sFontEnd
                    End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeID").Value)) & sBoldEnd & sFontEnd
                    sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("WorkedHours").Value)) & sBoldEnd & sFontEnd
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
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("UserFullName").Value)) & sBoldEnd & sFontEnd
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
			If aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) = -1 Then
				sErrorDescription = "Introduzca un número de empleado para consultar sus "
			Else
				sErrorDescription = "No existen registros en proceso de aplicación"
			End If
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayExternalSpecialJourneyTable = lErrorNumber
	Err.Clear
End Function

Function DisplayInternalSpecialJourneyTable(oRequest, oADODBConnection, bForExport, lStartPage, aSpecialJourneyComponent, iSpecialJourneyType, sErrorDescription)
'************************************************************
'Purpose: To display the absences for the given absence for
'		  the employee from the database in a table
'Inputs:  oRequest, oADODBConnection, bForExport, aSpecialJourneyComponent
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayInternalSpecialJourneyTable"
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

	oStartDate = Now()
	lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	lErrorNumber = GetSpecialJourneys(oRequest, oADODBConnection, aSpecialJourneyComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetchForSections(oRequest, CInt(oRequest("StartPage").Item), 10, aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE), oRecordset)
			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If (Not bForExport) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Or (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
					asColumnsTitles = Split("Acciones,No. Empleado,Nombre,RFC, CURP,Folio de autorización,F. de inicio,F. de termino,Días,Nómina en que se aplica,Registró", ",", -1, vbBinaryCompare)
					asCellWidths = Split("200,100,500,200,200,200,200,200,100,200,200,200", ",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,,,,,CENTER,CENTER,,,,", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("No. Empleado, Nombre,RFC, CURP,Folio de autorización,F. de inicio,F. de termino,Días,Nómina en que se aplica,Registró", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,500,200,200,200,200,200,100,200,200,200", ",", -1, vbBinaryCompare)
					asCellAlignments = Split(",,,,CENTER,CENTER,,,,", ",", -1, vbBinaryCompare)
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
					If (StrComp(CStr(oRecordset.Fields("RecordID").Value), oRequest("RecordID").Item, vbBinaryCompare) = 0) And (StrComp(CStr(oRecordset.Fields("StartDate").Value), oRequest("StartDate").Item, vbBinaryCompare) = 0) Then
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
						If CInt(oRecordset.Fields("Active").Value) = 1 Then
							sRowContents = sRowContents & "<A HREF=""" & "SpecialJourney.asp" & "?Action=SpecialJourney&SetDeActive=1&RecordID=" & CStr(oRecordset.Fields("RecordID").Value) & "&SpecialJourneyType=" & iSpecialJourneyType & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnDeactive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Desactivar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						Else
							sRowContents = sRowContents & "<A HREF=""" & "SpecialJourney.asp" & "?Action=SpecialJourney&SetActive=1&RecordID=" & CStr(oRecordset.Fields("RecordID").Value) & "&SpecialJourneyType=" & iSpecialJourneyType & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Aplicar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
							sRowContents = sRowContents & "<A HREF=""" & "SpecialJourney.asp" & "?Action=SpecialJourney&Remove=1&RecordID=" & CStr(oRecordset.Fields("RecordID").Value) & "&SpecialJourneyType=" & iSpecialJourneyType & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;"
						End If
                        sRowContents = sRowContents & sBoldEnd & sFontEnd
                    End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeID").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeFullName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)) & sBoldEnd & sFontEnd
                    sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("DocumentNumber").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("WorkedHours").Value)) & sBoldEnd & sFontEnd
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
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AddUserID").Value)) & sBoldEnd & sFontEnd
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
			If aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) = -1 Then
				sErrorDescription = "Introduzca un número de empleado para consultar sus "
			Else
				sErrorDescription = "No existen registros en proceso de aplicación"
			End If
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayInternalSpecialJourneyTable = lErrorNumber
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
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesSpecialJourneys.OcurredDate, EmployeesSpecialJourneys.RegistrationDate, EmployeesSpecialJourneys.AppliedDate, EmployeesSpecialJourneys.AbsenceID, Absences.AbsenceShortName, Absences.AbsenceName, COUNT(*) As Registros, SUM(AbsenceHours) As Dias From EmployeesSpecialJourneys, Absences Where (EmployeesSpecialJourneys.AbsenceID=Absences.AbsenceID) And (EmployeesSpecialJourneys.Active=0) And (Absences.AbsenceID<100) Group By EmployeesSpecialJourneys.OcurredDate, EmployeesSpecialJourneys.RegistrationDate, EmployeesSpecialJourneys.AppliedDate, EmployeesSpecialJourneys.AbsenceID, Absences.AbsenceShortName, Absences.AbsenceName Order by EmployeesSpecialJourneys.OcurredDate, Absences.AbsenceShortName", "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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

Function GetExternalSpecialJourneys(oRequest, oADODBConnection, aSpecialJourneyComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the absences for
'         the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aSpecialJourneyComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetExternalSpecialJourneys"
	Dim sTables
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	bComponentInitialized = aSpecialJourneyComponent(B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
	End If

	If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
		sTables = ", Jobs"
		sCondition = "And (Employees.JobID=Jobs.JobID) And ((Employees.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")) Or (Jobs.AreaID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ")))"
	Else
		sTables = ""
		sCondition = ""
	End If
	Call GetStartAndEndDatesFromURL("FilterStart", "FilterEnd", "OcurredDate", False, sCondition)

	If Len(sCondition ) > 0 Then
		If InStr(1, sCondition , "And ", vbBinaryCompare) = 0 Then sCondition  = "And " & sCondition
	End If

	sCondition  = "And (EmployeesSpecialJourneys.SpecialJourneyID = " & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_SPECIAL_JOURNEY_ID) & ")"

    If aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) > 0 Then
        sCondition = sCondition & " And (EmployeesSpecialJourneys.EmployeeID=" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_EMPLOYEE_ID) & ")"
    Else
        sCondition = sCondition & " And (EmployeesSpecialJourneys.EmployeeID=0)"
    End If
    If aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ACTIVE) > 0 Then
        sCondition = sCondition & " And (EmployeesSpecialJourneys.Active=1)"
    Else
        sCondition = sCondition & " And (EmployeesSpecialJourneys.Active=0)"
    End If
    sCondition = sCondition & aSpecialJourneyComponent(S_QUERY_CONDITION_SPECIAL_JOURNEY)

	'If CInt(aEmployeeComponent(N_ID_EMPLOYEE)) > 0 Then
	'	sQuery = sQuery & " And (BankAccounts.EmployeeID=" & aEmployeeComponent(N_ID_EMPLOYEE) & ")"
	'Else
	'	If iActive Then
	'		sQuery = sQuery & " And (BankAccounts.EmployeeID=0)"
	'	End If
	'End If

	sQuery = "Select EmployeesSpecialJourneys.*, EmployeeName + ' ' + EmployeeLastName + ' ' + EmployeeLastName2 As EmployeeFullName, RFC, CURP, SPEPID, UserName + ' ' + UserLastName As UserFullName" & _
			 " From EmployeesSpecialJourneys, ExternalSpecialJourneys, Users"
    sQuery = sQuery & " Where (EmployeesSpecialJourneys.EmployeeID = ExternalSpecialJourneys.ExternalID) " & _
			" And (EmployeesSpecialJourneys.AddUserID = Users.UserID) " & sCondition
	sQuery = sQuery & " Order By EmployeeID, StartDate"

	sErrorDescription = "No se pudo obtener la información de los registros."

	If CInt(Request.Cookies("SIAP_SectionID")) <> 7 Then  ' Dif. de Desc.
		If CInt(Request.Cookies("SIAP_SubSectionID")) = 22 Then  ' Igual a Prestaciones e incidencias
			sCondition = sCondition & " And (EmployeesSpecialJourneys.PositionID=323)"
		Else ' Igual a Inf. - Emp. - Inci
			sCondition = sCondition & " And (EmployeesSpecialJourneys.PositionID<>323)"
		End If
	Else ' Igual a Desc.
		If CInt(Request.Cookies("SIAP_SubSectionID")) = 721 Then  ' Igual a Prestaciones e incidencias
			sCondition = sCondition & " And (EmployeesSpecialJourneys.PositionID=323)"
		Else
			sCondition = sCondition & " And (EmployeesSpecialJourneys.PositionID<>323)"
		End If
	End If
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & sQuery & """ />"

	GetExternalSpecialJourneys = lErrorNumber
	Err.Clear
End Function

Function SetActiveForEmployeeSpecialJourney(oRequest, oADODBConnection, aSpecialJourneyComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given employee's concept
'Inputs:  oRequest, oADODBConnection
'Outputs: aSpecialJourneyComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForEmployeeSpecialJourney"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aSpecialJourneyComponent(B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
	End If

	If (aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ID) = -1) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del empleado y/o el identificador de la incidencia y/o la fecha para agregar la información del registro."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo modificar la información del concepto."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesSpecialJourneys Set Active=1 Where (RecordID=" & aSpecialJourneyComponent(N_SPECIAL_JOURNEY_ID) & ")", "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, Null)
	End If

	SetActiveForEmployeeSpecialJourney = lErrorNumber
	Err.Clear
End Function

Function VerifyRequerimentsForEmployeesSpecialJourneys(oADODBConnection, aEmployeeComponent, sErrorDescription)
'************************************************************
'Purpose: To verify employee status requirements to register absences
'Inputs:  oADODBConnection, lReasonID, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyRequerimentsForEmployeesSpecialJourneys"
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

	bComponentInitialized = aSpecialJourneyComponent(B_COMPONENT_INITIALIZED_SPECIAL_JOURNEY)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeSpecialJourneyComponent(oRequest, aSpecialJourneyComponent)
	End If
	'VerifyRequerimentsForEmployeesSpecialJourneys = True
	If (aEmployeeComponent(N_ID_EMPLOYEE) = -1) Then
		sErrorDescription = "No se especificó el identificador del empleado para agregar incidencias."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 0, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
		VerifyRequerimentsForEmployeesSpecialJourneys = False
    'If True Then
        VerifyRequerimentsForEmployeesSpecialJourneys = True
	Else
		'aEmployeeComponent(N_ID_EMPLOYEE) = aSpecialJourneyComponent(N_EMPLOYEE_ID_ABSENCE)
		lErrorNumber = CheckExistencyOfEmployeeID(aEmployeeComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			If VerifyUserPermissionOnEmployee(oADODBConnection, aEmployeeComponent, sErrorDescription) Then
				sQuery = "Select Employees.*, Areas.AreaID, Areas.AreaCode, Areas.AreaName, Positions.PositionID, Positions.PositionShortName, Positions.PositionName, ServiceShortName, ServiceName, ShiftShortName, ShiftName, LevelShortName, PositionsSpecialJourneysLKP.RecordID From Employees, Jobs, Areas, Positions, Services, Shifts, Levels, PositionsSpecialJourneysLKP Where (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Jobs.PositionID=Positions.PositionID) And (Employees.ServiceID=Services.ServiceID) And (Employees.ShiftID=Shifts.ShiftID) And (Employees.LevelID=Levels.LevelID) And (PositionsSpecialJourneysLKP.PositionID=Positions.PositionID) And (PositionsSpecialJourneysLKP.LevelID=Employees.LevelID) And (PositionsSpecialJourneysLKP.WorkingHours=Employees.WorkingHours) And (PositionsSpecialJourneysLKP.ServiceID=Employees.ServiceID) And (PositionsSpecialJourneysLKP.CenterTypeID=Areas.CenterTypeID) And (PositionsSpecialJourneysLKP.IsActive1=1) And (Employees.EmployeeNumber='" & aEmployeeComponent(N_ID_EMPLOYEE) & "') And (PositionsSpecialJourneysLKP.StartDate<=" & CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) & ") And (PositionsSpecialJourneysLKP.EndDate>=" & CLng(Left(GetSerialNumberForDate(""), Len("00000000"))) & ") And (PositionsSpecialJourneysLKP.Active=1)"
				sErrorDescription = "No se pudieron obtener los datos del empleado para validar que cumpla con la matríz para guardias."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						iJobID = CLng(oRecordset.Fields("JobID").Value)
						iEmployeeTypeID = CInt(oRecordset.Fields("EmployeeTypeID").Value)
						iStatusEmployeeID = CInt(oRecordset.Fields("StatusID").Value)
						iPositionTypeID = CInt(oRecordset.Fields("PositionTypeID").Value)
						iShiftID = CInt(oRecordset.Fields("ShiftID").Value)
						iServiceID = CInt(oRecordset.Fields("ServiceID").Value)
						oRecordset.Close
                        VerifyRequerimentsForEmployeesSpecialJourneys = True
					Else
						sErrorDescription = "El empleado no existe o las características de su puesto no permiten registrar guardias."
						VerifyRequerimentsForEmployeesSpecialJourneys = False
					End If
				Else
					sErrorDescription = "Error al verificar el status del empleado para registrar la incidencia"
					VerifyRequerimentsForEmployeesSpecialJourneys = False
				End If
			Else
				VerifyRequerimentsForEmployeesSpecialJourneys = False
			End If
		Else
			VerifyRequerimentsForEmployeesSpecialJourneys = False
		End If
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function Display423SearchForm(oRequest, oADODBConnection, iSectionID, sErrorDescription)
'************************************************************
'Purpose: To display the search form for the EmployeesKardex5
'Inputs:  oRequest, oADODBConnection, iSectionID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "Display423SearchForm"

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "function AddEmployeeIDToSearchList() {" & vbNewLine
			Response.Write "var oForm = document.SearchFrm;" & vbNewLine
			Response.Write "if (oForm.EmployeeID.value != '') {" & vbNewLine
				Response.Write "oForm.EmployeeID.value = '000000' + oForm.EmployeeID.value;" & vbNewLine
				Response.Write "AddItemToList(oForm.EmployeeID.value.substr(oForm.EmployeeID.value.length - 6), oForm.EmployeeID.value.substr(oForm.EmployeeID.value.length - 6), null, oForm.EmployeeIDs)" & vbNewLine
				Response.Write "SelectAllItemsFromList(oForm.EmployeeIDs);" & vbNewLine
				Response.Write "oForm.EmployeeID.value = '';" & vbNewLine
				Response.Write "oForm.EmployeeID.focus();" & vbNewLine
			Response.Write "}" & vbNewLine
		Response.Write "} // End of AddDisasterIDToSearchList" & vbNewLine

		Response.Write "function Show423Fields(sValue) {" & vbNewLine
			Response.Write "if (sValue == '1') {" & vbNewLine
				Response.Write "ShowDisplay(document.all['EmployeeNumberDiv']);" & vbNewLine
				Response.Write "HideDisplay(document.all['EmployeeNameDiv']);" & vbNewLine
				Response.Write "HideDisplay(document.all['EmployeeLastNameDiv']);" & vbNewLine
				Response.Write "HideDisplay(document.all['EmployeeLastName2Div']);" & vbNewLine
				Response.Write "HideDisplay(document.all['EmployeeRFCDiv']);" & vbNewLine
			Response.Write "} else {" & vbNewLine
				Response.Write "HideDisplay(document.all['EmployeeNumberDiv']);" & vbNewLine
				Response.Write "ShowDisplay(document.all['EmployeeNameDiv']);" & vbNewLine
				Response.Write "ShowDisplay(document.all['EmployeeLastNameDiv']);" & vbNewLine
				Response.Write "ShowDisplay(document.all['EmployeeLastName2Div']);" & vbNewLine
				Response.Write "ShowDisplay(document.all['EmployeeRFCDiv']);" & vbNewLine
			Response.Write "}" & vbNewLine
		Response.Write "} // End of Show423Fields" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine
	Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Indique los criterios que se utilizarán en la búsqueda de registros:</B></FONT><BR />"
	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Main_ISSSTE.asp"" METHOD=""GET"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & iSectionID & """ />"
		Response.Write "<TABLE BORDER=""0"" CELLPADING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR NAME=""EmployeeNumberDiv"" ID=""EmployeeNumberDiv"">"
				Response.Write "<TD VALIGN=""TOP"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Número de empleado:<BR /></FONT>"
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
					Response.Write "&nbsp;&nbsp;<INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" SIZE=""10"" MAXLENGTH=""6"" VALUE=""" & oRequest("EmployeeID").Item & """ CLASS=""TextFields"" />"
					Response.Write "&nbsp;&nbsp;<A HREF=""javascript: AddEmployeeIDToSearchList();""><IMG SRC=""Images/BtnCrclAdd.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Agregar"" BORDER=""0"" /></A>&nbsp;&nbsp;<BR />"
				Response.Write "</TD>"
					Response.Write "<TD VALIGN=""TOP""><BR />"
						Response.Write "<SELECT NAME=""EmployeeIDs"" ID=""EmployeeIDsCmb"" SIZE=""6"" MULTIPLE=""1"" CLASS=""Lists"" STYLE=""width: 100px;"">"
							If Len(oRequest("EmployeeIDs").Item) > 0 Then
								For Each oItem In oRequest("EmployeeIDs")
									Response.Write "<OPTION VALUE=""" & oItem & """ SELECTED=""1"">" & oItem & "</OPTION>"
								Next
							End If
						Response.Write "</SELECT>"
						Response.Write "&nbsp;<A HREF=""javascript: RemoveSelectedItemsFromList(null, document.ReportFrm.EmployeeIDs); SelectAllItemsFromList(document.ReportFrm.EmployeeIDs);""><IMG SRC=""Images/BtnCrclDelete.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Quitar"" BORDER=""0""></A><BR />"
					Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR NAME=""EmployeeNameDiv"" ID=""EmployeeNameDiv"" STYLE=""display: none"">"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre del empleado:</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeName"" ID=""EmployeeNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("EmployeeName").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR NAME=""EmployeeLastNameDiv"" ID=""EmployeeLastNameDiv"" STYLE=""display: none"">"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido paterno:</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName"" ID=""EmployeeLastNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("EmployeeLastName").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR NAME=""EmployeeLastName2Div"" ID=""EmployeeLastName2Div"" STYLE=""display: none"">"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Apellido materno:</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeLastName2"" ID=""EmployeeLastName2Txt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("EmployeeLastName2").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR NAME=""EmployeeRFCDiv"" ID=""EmployeeRFCDiv"" STYLE=""display: none"">"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">RFC:</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeRFC"" ID=""EmployeeRFCTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & oRequest("EmployeeRFC").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			If CInt(iSectionID) = 424 Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado a suplir:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""OriginalEmployeeID"" ID=""OriginalEmployeeIDTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & oRequest("OriginalEmployeeID").Item & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
			End If
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Adscripción:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""AreaID"" ID=""AreaIDLst"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.value == '-1') {document.HierarchyMenuIFrame.location.href = 'HierarchyMenu.asp';} else {document.HierarchyMenuIFrame.location.href = 'HierarchyMenu.asp?Action=SubAreas&TargetField=' + this.form.name + '.SubAreaID&AreaID=' + this.value;}"">"
					Response.Write "<OPTION VALUE="""">Todas</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(ParentID>-1) And (AreaID>-1) And (EndDate=30000000) And (Active=1) And (CenterTypeID In (Select Distinct CenterTypeID From PositionsSpecialJourneysLKP))", "AreaCode, AreaName", oRequest("AreaID").Item, "", sErrorDescription)
				Response.Write "</SELECT><BR />"
			Response.Write "</TR>"
			If CInt(iSectionID) <> 425 Then
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Turno:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""JourneyID"" ID=""JourneyIDLst"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SpecialJourneys", "JourneyID", "JourneyShortName, JourneyName", "(RecordTypeID In (-1," & CInt(Request.Cookies("SIAP_SubSectionID")) & ")) And (Active=1)", "JourneyShortName", oRequest("JourneyID").Item, "", sErrorDescription)
					Response.Write "</SELECT><BR />"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Movimiento:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""MovementID"" ID=""MovementIDLst"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "SpecialJourneysMovements", "MovementID", "MovementShortName, MovementName", "(RecordTypeID In (-1," & CInt(Request.Cookies("SIAP_SubSectionID")) & ")) And (Active=1)", "MovementShortName", oRequest("MovementID").Item, "", sErrorDescription)
					Response.Write "</SELECT><BR />"
				Response.Write "</TR>"
			End If
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha de inicio:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartStartYear").Item), CInt(oRequest("StartStartMonth").Item), CInt(oRequest("StartStartDay").Item), "StartStartYear", "StartStartMonth", "StartStartDay", 2009, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndStartYear").Item), CInt(oRequest("EndStartMonth").Item), CInt(oRequest("EndStartDay").Item), "EndStartYear", "EndStartMonth", "EndStartDay", 2009, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Fecha de término:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartEndYear").Item), CInt(oRequest("StartEndMonth").Item), CInt(oRequest("StartEndDay").Item), "StartEndYear", "StartEndMonth", "StartEndDay", 2009, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndEndYear").Item), CInt(oRequest("EndEndMonth").Item), CInt(oRequest("EndEndDay").Item), "EndEndYear", "EndEndMonth", "EndEndDay", 2009, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Tipo de personal:&nbsp;</FONT></TD>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
					'Response.Write "<INPUT TYPE=""Radio"" NAME=""Internal"" ID=""InternalRd"" VALUE="""""
					'	If Len(oRequest("Internal").Item) = 0 Then
					'		Response.Write " CHECKED=""1"""
					'	End If
					'Response.Write " />Ambos&nbsp;&nbsp;&nbsp;"
					Response.Write "<INPUT TYPE=""Radio"" NAME=""Internal"" ID=""InternalRd"" VALUE=""1"""
						If StrComp(oRequest("Internal").Item, "0", vbBinaryCompare) <> 0 Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " onClick=""Show423Fields(this.value);"" />Interno&nbsp;&nbsp;&nbsp;<INPUT TYPE=""Radio"" NAME=""Internal"" ID=""InternalRd"" VALUE=""0"""
						If StrComp(oRequest("Internal").Item, "0", vbBinaryCompare) = 0 Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " onClick=""Show423Fields(this.value);"" />Externo<BR /><BR />"
				Response.Write "</FONT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2"" ALIGN=""RIGHT""><BR /><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Registros"" CLASS=""Buttons"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	Display423SearchForm = Err.number
	Err.Clear
End Function

Function VerifySpecialJourneyBudgetAmount(oADODBConnection, lAppliedDate, iAreaID, lEmployeeTypeID, lNewAmount)
'************************************************************
'Purpose: To verify employee status requirements to register absences
'Inputs:  oADODBConnection, lAppliedDate, iAreaID, lEmployeeTypeID, 
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifySpecialJourneyBudgetAmount"
	Dim lErrorNumber
	Dim oRecordset
	Dim sQuery
	Dim lEmployeesSpJourneysAmount
	Dim lBudgetSpJourneysAmount
	Dim sQueryEmployeesSpJourneysAmount
	Dim sQueryBudgetSpJourneysAmount
	Dim sUR
	Dim iYear
	Dim iMonth

	Call GetNameFromTable(oADODBConnection, "AreasURCTAUX", iAreaID, "", "", sUR, "")
	sUR = CStr(Mid(sUR,1,3))
	iYear = CLng(Left(lAppliedDate, Len("1976")))
	iMonth = CLng(Mid(lAppliedDate, Len("19760"), Len("02")))

	sQueryEmployeesSpJourneysAmount = "Select SUM(ConceptAmount) TotalEmployeesSpJourneysAmount From EmployeesSpecialJourneys Where (AppliedDate = " & lAppliedDate & ") And AreaID IN (" & _
										"Select AreaID from Areas Where Substr(URCTAUX, 1, 3) = '" & sUR & "'" & _
									  ")"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQueryEmployeesSpJourneysAmount, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			lEmployeesSpJourneysAmount = CDbl(oRecordset.Fields("TotalEmployeesSpJourneysAmount").Value)
		End If
	End If

	sQueryBudgetSpJourneysAmount = "Select SUM(OriginalAmount) TotalBudgetSpJourneysAmount From Budgets_Short Where (ZoneID= " & sUR & ")" & _
								   " And (BudgetYear= " & iYear & ") And (BudgetMonth=" & iMonth & ")"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQueryBudgetSpJourneysAmount, "EmployeeSpecialJourneyComponent.asp", S_FUNCTION_NAME, 0, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			lBudgetSpJourneysAmount = CDbl(oRecordset.Fields("TotalBudgetSpJourneysAmount").Value)
		End If
	End If

	If lBudgetSpJourneysAmount >= lEmployeesSpJourneysAmount + lNewAmount Then
		VerifySpecialJourneyBudgetAmount = True
	Else
		VerifySpecialJourneyBudgetAmount = False
	End If

	Err.Clear
End Function
%>