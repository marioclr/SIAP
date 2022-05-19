<%
Const N_ID_BUDGET = 0
Const N_PARENT_ID_BUDGET = 1
Const S_SHORT_NAME_BUDGET = 2
Const S_NAME_BUDGET = 3
Const S_PATH_BUDGET = 4
Const N_BUDGET_TYPE_ID_BUDGET = 5
Const N_ACTIVE_BUDGET = 6

Const D_AMOUNT_BUDGET = 7
Const N_QTTY_ID_BUDGET = 8

Const N_AREA_ID_BUDGET = 9
Const N_PROGRAM_DUTY_ID_BUDGET = 10
Const N_FUND_ID_BUDGET = 11
Const N_DUTY_ID_BUDGET = 12
Const N_ACTIVE_DUTY_ID_BUDGET = 13
Const N_SPECIFIC_DUTY_ID_BUDGET = 14
Const N_PROGRAM_ID_BUDGET = 15
Const N_REGION_ID_BUDGET = 16
Const N_UR_BUDGET = 17
Const N_CT_BUDGET = 18
Const N_AUX_BUDGET = 19
Const N_LOCATION_ID_BUDGET = 20
Const N_BUDGET_ID1_BUDGET = 21
Const N_BUDGET_ID2_BUDGET = 22
Const N_BUDGET_ID3_BUDGET = 23
Const N_CONFINE_TYPE_ID_BUDGET = 24
Const N_ACTIVITY_ID1_BUDGET = 25
Const N_ACTIVITY_ID2_BUDGET = 26
Const N_PROCESS_ID_BUDGET = 27
Const N_YEAR_BUDGET = 28
Const N_MONTH_BUDGET = 29
Const AD_ORIGINAL_AMOUNT_BUDGET = 30
Const AD_MODIFIED_AMOUNT_BUDGET = 31

Const S_QUERY_CONDITION_BUDGET = 32
Const B_CHECK_FOR_DUPLICATED_BUDGET = 33
Const B_IS_DUPLICATED_BUDGET = 34
Const B_COMPONENT_INITIALIZED_BUDGET = 35

Const N_BUDGET_COMPONENT_SIZE = 35

Dim aBudgetComponent()
Redim aBudgetComponent(N_BUDGET_COMPONENT_SIZE)

Function InitializeBudgetComponent(oRequest, aBudgetComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Budget
'         Component using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aBudgetComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializeBudgetComponent"
	Redim Preserve aBudgetComponent(N_BUDGET_COMPONENT_SIZE)
	Dim iIndex

	If IsEmpty(aBudgetComponent(N_ID_BUDGET)) Then
		If Len(oRequest("BudgetID").Item) > 0 Then
			aBudgetComponent(N_ID_BUDGET) = CLng(oRequest("BudgetID").Item)
		ElseIf Len(oRequest("MoneyID").Item) > 0 Then
			aBudgetComponent(N_ID_BUDGET) = CLng(oRequest("MoneyID").Item)
		ElseIf Len(oRequest("ProgramID").Item) > 0 Then
			aBudgetComponent(N_ID_BUDGET) = CLng(oRequest("ProgramID").Item)
		Else
			aBudgetComponent(N_ID_BUDGET) = -1
		End If
	End If

	If IsEmpty(aBudgetComponent(N_PARENT_ID_BUDGET)) Then
		If Len(oRequest("ParentID").Item) > 0 Then
			aBudgetComponent(N_PARENT_ID_BUDGET) = CLng(oRequest("ParentID").Item)
		Else
			aBudgetComponent(N_PARENT_ID_BUDGET) = -1
		End If
	End If

	If IsEmpty(aBudgetComponent(S_SHORT_NAME_BUDGET)) Then
		If Len(oRequest("BudgetShortName").Item) > 0 Then
			aBudgetComponent(S_SHORT_NAME_BUDGET) = oRequest("BudgetShortName").Item
		ElseIf Len(oRequest("MoneyShortName").Item) > 0 Then
			aBudgetComponent(S_SHORT_NAME_BUDGET) = oRequest("MoneyShortName").Item
		ElseIf Len(oRequest("ProgramShortName").Item) > 0 Then
			aBudgetComponent(S_SHORT_NAME_BUDGET) = oRequest("ProgramShortName").Item
		Else
			aBudgetComponent(S_SHORT_NAME_BUDGET) = ""
		End If
	End If
	aBudgetComponent(S_SHORT_NAME_BUDGET) = Left(aBudgetComponent(S_SHORT_NAME_BUDGET), 10)

	If IsEmpty(aBudgetComponent(S_NAME_BUDGET)) Then
		If Len(oRequest("BudgetName").Item) > 0 Then
			aBudgetComponent(S_NAME_BUDGET) = oRequest("BudgetName").Item
		ElseIf Len(oRequest("MoneyName").Item) > 0 Then
			aBudgetComponent(S_NAME_BUDGET) = oRequest("MoneyName").Item
		ElseIf Len(oRequest("ProgramName").Item) > 0 Then
			aBudgetComponent(S_NAME_BUDGET) = oRequest("ProgramName").Item
		Else
			aBudgetComponent(S_NAME_BUDGET) = ""
		End If
	End If
	aBudgetComponent(S_NAME_BUDGET) = Left(aBudgetComponent(S_NAME_BUDGET), 255)

	If IsEmpty(aBudgetComponent(S_PATH_BUDGET)) Then
		If Len(oRequest("BudgetPath").Item) > 0 Then
			aBudgetComponent(S_PATH_BUDGET) = oRequest("BudgetPath").Item
		ElseIf Len(oRequest("MoneyPath").Item) > 0 Then
			aBudgetComponent(S_PATH_BUDGET) = oRequest("MoneyPath").Item
		ElseIf Len(oRequest("ProgramPath").Item) > 0 Then
			aBudgetComponent(S_PATH_BUDGET) = oRequest("ProgramPath").Item
		Else
			aBudgetComponent(S_PATH_BUDGET) = ",-1,"
			If aBudgetComponent(N_ID_BUDGET) > -1 Then aBudgetComponent(S_PATH_BUDGET) = aBudgetComponent(S_PATH_BUDGET) & aBudgetComponent(N_ID_BUDGET) & ","
		End If
	End If
	aBudgetComponent(S_PATH_BUDGET) = Left(aBudgetComponent(S_PATH_BUDGET), 255)

	If IsEmpty(aBudgetComponent(N_BUDGET_TYPE_ID_BUDGET)) Then
		aBudgetComponent(N_BUDGET_TYPE_ID_BUDGET) = -1
	End If

	If IsEmpty(aBudgetComponent(N_ACTIVE_BUDGET)) Then
		If Len(oRequest("Active").Item) > 0 Then
			aBudgetComponent(N_ACTIVE_BUDGET) = CInt(oRequest("Active").Item)
		Else
			aBudgetComponent(N_ACTIVE_BUDGET) = 1
		End If
	End If

	If IsEmpty(aBudgetComponent(D_AMOUNT_BUDGET)) Then
		If Len(oRequest("BudgetAmount").Item) > 0 Then
			aBudgetComponent(D_AMOUNT_BUDGET) = CDbl(oRequest("BudgetAmount").Item)
		Else
			aBudgetComponent(D_AMOUNT_BUDGET) = 0
		End If
	End If

	If IsEmpty(aBudgetComponent(N_QTTY_ID_BUDGET)) Then
		If Len(oRequest("QttyID").Item) > 0 Then
			aBudgetComponent(N_QTTY_ID_BUDGET) = CInt(oRequest("QttyID").Item)
		Else
			aBudgetComponent(N_QTTY_ID_BUDGET) = 2
		End If
	End If

	If Len(oRequest("ModifyMoneys").Item) = 0 Then
		If IsEmpty(aBudgetComponent(N_AREA_ID_BUDGET)) Then
			If Len(oRequest("AreaID").Item) > 0 Then
				aBudgetComponent(N_AREA_ID_BUDGET) = CLng(oRequest("AreaID").Item)
			Else
				aBudgetComponent(N_AREA_ID_BUDGET) = -1
			End If
		End If

		If IsEmpty(aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET)) Then
			If Len(oRequest("ProgramDutyID").Item) > 0 Then
				aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET) = CLng(oRequest("ProgramDutyID").Item)
			Else
				aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET) = -1
			End If
		End If

		If IsEmpty(aBudgetComponent(N_FUND_ID_BUDGET)) Then
			If Len(oRequest("FundID").Item) > 0 Then
				aBudgetComponent(N_FUND_ID_BUDGET) = CLng(oRequest("FundID").Item)
			Else
				aBudgetComponent(N_FUND_ID_BUDGET) = -1
			End If
		End If

		If IsEmpty(aBudgetComponent(N_DUTY_ID_BUDGET)) Then
			If Len(oRequest("DutyID").Item) > 0 Then
				aBudgetComponent(N_DUTY_ID_BUDGET) = CLng(oRequest("DutyID").Item)
			Else
				aBudgetComponent(N_DUTY_ID_BUDGET) = -1
			End If
		End If

		If IsEmpty(aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET)) Then
			If Len(oRequest("ActiveDutyID").Item) > 0 Then
				aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET) = CLng(oRequest("ActiveDutyID").Item)
			Else
				aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET) = -1
			End If
		End If

		If IsEmpty(aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET)) Then
			If Len(oRequest("SpecificDutyID").Item) > 0 Then
				aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET) = CLng(oRequest("SpecificDutyID").Item)
			Else
				aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET) = -1
			End If
		End If

		If IsEmpty(aBudgetComponent(N_PROGRAM_ID_BUDGET)) Then
			If Len(oRequest("ProgramID").Item) > 0 Then
				aBudgetComponent(N_PROGRAM_ID_BUDGET) = CLng(oRequest("ProgramID").Item)
			Else
				aBudgetComponent(N_PROGRAM_ID_BUDGET) = -1
			End If
		End If

		If IsEmpty(aBudgetComponent(N_REGION_ID_BUDGET)) Then
			If Len(oRequest("RegionID").Item) > 0 Then
				aBudgetComponent(N_REGION_ID_BUDGET) = CLng(oRequest("RegionID").Item)
			Else
				aBudgetComponent(N_REGION_ID_BUDGET) = -1
			End If
		End If

		If IsEmpty(aBudgetComponent(N_UR_BUDGET)) Then
			If Len(oRequest("BudgetUR").Item) > 0 Then
				aBudgetComponent(N_UR_BUDGET) = CLng(oRequest("BudgetUR").Item)
			Else
				aBudgetComponent(N_UR_BUDGET) = 0
			End If
		End If

		If IsEmpty(aBudgetComponent(N_CT_BUDGET)) Then
			If Len(oRequest("BudgetCT").Item) > 0 Then
				aBudgetComponent(N_CT_BUDGET) = CLng(oRequest("BudgetCT").Item)
			Else
				aBudgetComponent(N_CT_BUDGET) = 0
			End If
		End If

		If IsEmpty(aBudgetComponent(N_AUX_BUDGET)) Then
			If Len(oRequest("BudgetAUX").Item) > 0 Then
				aBudgetComponent(N_AUX_BUDGET) = CLng(oRequest("BudgetAUX").Item)
			Else
				aBudgetComponent(N_AUX_BUDGET) = 0
			End If
		End If

		If IsEmpty(aBudgetComponent(N_LOCATION_ID_BUDGET)) Then
			If Len(oRequest("LocationID").Item) > 0 Then
				aBudgetComponent(N_LOCATION_ID_BUDGET) = CLng(oRequest("LocationID").Item)
			Else
				aBudgetComponent(N_LOCATION_ID_BUDGET) = -1
			End If
		End If

		If IsEmpty(aBudgetComponent(N_BUDGET_ID1_BUDGET)) Then
			If Len(oRequest("BudgetID1").Item) > 0 Then
				aBudgetComponent(N_BUDGET_ID1_BUDGET) = CLng(oRequest("BudgetID1").Item)
			Else
				aBudgetComponent(N_BUDGET_ID1_BUDGET) = -1
			End If
		End If

		If IsEmpty(aBudgetComponent(N_BUDGET_ID2_BUDGET)) Then
			If Len(oRequest("BudgetID2").Item) > 0 Then
				aBudgetComponent(N_BUDGET_ID2_BUDGET) = CLng(oRequest("BudgetID2").Item)
			Else
				aBudgetComponent(N_BUDGET_ID2_BUDGET) = -1
			End If
		End If

		If IsEmpty(aBudgetComponent(N_BUDGET_ID3_BUDGET)) Then
			If Len(oRequest("BudgetID3").Item) > 0 Then
				aBudgetComponent(N_BUDGET_ID3_BUDGET) = CLng(oRequest("BudgetID3").Item)
			Else
				aBudgetComponent(N_BUDGET_ID3_BUDGET) = -1
				
			End If
		End If

		If IsEmpty(aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET)) Then
			If Len(oRequest("ConfineTypeID").Item) > 0 Then
				aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET) = CLng(oRequest("ConfineTypeID").Item)
			Else
				aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET) = 1
			End If
		End If

		If IsEmpty(aBudgetComponent(N_ACTIVITY_ID1_BUDGET)) Then
			If Len(oRequest("ActivityID1").Item) > 0 Then
				aBudgetComponent(N_ACTIVITY_ID1_BUDGET) = CLng(oRequest("ActivityID1").Item)
			Else
				aBudgetComponent(N_ACTIVITY_ID1_BUDGET) = -1
			End If
		End If

		If IsEmpty(aBudgetComponent(N_ACTIVITY_ID2_BUDGET)) Then
			If Len(oRequest("ActivityID2").Item) > 0 Then
				aBudgetComponent(N_ACTIVITY_ID2_BUDGET) = CLng(oRequest("ActivityID2").Item)
			Else
				aBudgetComponent(N_ACTIVITY_ID2_BUDGET) = -1
			End If
		End If

		If IsEmpty(aBudgetComponent(N_PROCESS_ID_BUDGET)) Then
			If Len(oRequest("ProcessID").Item) > 0 Then
				aBudgetComponent(N_PROCESS_ID_BUDGET) = CLng(oRequest("ProcessID").Item)
			Else
				aBudgetComponent(N_PROCESS_ID_BUDGET) = -1
			End If
		End If

		If IsEmpty(aBudgetComponent(N_YEAR_BUDGET)) Then
			If Len(oRequest("BudgetYear").Item) > 0 Then
				aBudgetComponent(N_YEAR_BUDGET) = CInt(oRequest("BudgetYear").Item)
			Else
				aBudgetComponent(N_YEAR_BUDGET) = Year(Date())
			End If
		End If

		If IsEmpty(aBudgetComponent(N_MONTH_BUDGET)) Then
			If Len(oRequest("BudgetMonth").Item) > 0 Then
				aBudgetComponent(N_MONTH_BUDGET) = CInt(oRequest("BudgetMonth").Item)
			Else
				aBudgetComponent(N_MONTH_BUDGET) = 1
			End If
		End If

		If IsEmpty(aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET)) Then
			If Len(oRequest("OriginalAmount_01").Item) > 0 Then
				For iIndex = 1 To 12
					aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) = aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) & "," & oRequest("OriginalAmount_" & Right(("0" & iIndex), Len("00"))).Item
				Next
			Else
				aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) = ",0,0,0,0,0,0,0,0,0,0,0,0"
			End If
		End If
		aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) = Split(aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET), ",")

		If IsEmpty(aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET)) Then
			If Len(oRequest("ModifiedAmount_01").Item) > 0 Then
				For iIndex = 1 To 12
					aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) = aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) & "," & oRequest("ModifiedAmount_" & Right(("0" & iIndex), Len("00"))).Item
				Next
			Else
				aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) = ",0,0,0,0,0,0,0,0,0,0,0,0"
			End If
		End If
		aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) = Split(aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET), ",")
	Else
		If IsEmpty(aBudgetComponent(N_AREA_ID_BUDGET)) Then
			If Len(oRequest("AreaID").Item) > 0 Then
				aBudgetComponent(N_AREA_ID_BUDGET) = Split(oRequest("AreaID").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_AREA_ID_BUDGET) = Split("-1,-1", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET)) Then
			If Len(oRequest("ProgramDutyID").Item) > 0 Then
				aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET) = Split(oRequest("ProgramDutyID").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET) = Split("-1,-1", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_FUND_ID_BUDGET)) Then
			If Len(oRequest("FundID").Item) > 0 Then
				aBudgetComponent(N_FUND_ID_BUDGET) = Split(oRequest("FundID").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_FUND_ID_BUDGET) = Split("-1,-1", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_DUTY_ID_BUDGET)) Then
			If Len(oRequest("DutyID").Item) > 0 Then
				aBudgetComponent(N_DUTY_ID_BUDGET) = Split(oRequest("DutyID").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_DUTY_ID_BUDGET) = Split("-1,-1", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET)) Then
			If Len(oRequest("ActiveDutyID").Item) > 0 Then
				aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET) = Split(oRequest("ActiveDutyID").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET) = Split("-1,-1", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET)) Then
			If Len(oRequest("SpecificDutyID").Item) > 0 Then
				aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET) = Split(oRequest("SpecificDutyID").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET) = Split("-1,-1", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_PROGRAM_ID_BUDGET)) Then
			If Len(oRequest("ProgramID").Item) > 0 Then
				aBudgetComponent(N_PROGRAM_ID_BUDGET) = Split(oRequest("ProgramID").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_PROGRAM_ID_BUDGET) = Split("-1,-1", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_REGION_ID_BUDGET)) Then
			If Len(oRequest("RegionID").Item) > 0 Then
				aBudgetComponent(N_REGION_ID_BUDGET) = Split(oRequest("RegionID").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_REGION_ID_BUDGET) = Split("-1,-1", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_UR_BUDGET)) Then
			If Len(oRequest("BudgetUR").Item) > 0 Then
				aBudgetComponent(N_UR_BUDGET) = Split(oRequest("BudgetUR").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_UR_BUDGET) = Split("0,0", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_CT_BUDGET)) Then
			If Len(oRequest("BudgetCT").Item) > 0 Then
				aBudgetComponent(N_CT_BUDGET) = Split(oRequest("BudgetCT").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_CT_BUDGET) = Split("0,0", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_AUX_BUDGET)) Then
			If Len(oRequest("BudgetAUX").Item) > 0 Then
				aBudgetComponent(N_AUX_BUDGET) = Split(oRequest("BudgetAUX").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_AUX_BUDGET) = Split("0,0", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_LOCATION_ID_BUDGET)) Then
			If Len(oRequest("LocationID").Item) > 0 Then
				aBudgetComponent(N_LOCATION_ID_BUDGET) = Split(oRequest("LocationID").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_LOCATION_ID_BUDGET) = Split("-1,-1", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_BUDGET_ID1_BUDGET)) Then
			If Len(oRequest("BudgetID1").Item) > 0 Then
				aBudgetComponent(N_BUDGET_ID1_BUDGET) = Split(oRequest("BudgetID1").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_BUDGET_ID1_BUDGET) = Split("-1,-1", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_BUDGET_ID2_BUDGET)) Then
			If Len(oRequest("BudgetID2").Item) > 0 Then
				aBudgetComponent(N_BUDGET_ID2_BUDGET) = Split(oRequest("BudgetID2").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_BUDGET_ID2_BUDGET) = Split("-1,-1", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_BUDGET_ID3_BUDGET)) Then
			If Len(oRequest("BudgetID3").Item) > 0 Then
				aBudgetComponent(N_BUDGET_ID3_BUDGET) = Split(oRequest("BudgetID3").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_BUDGET_ID3_BUDGET) = Split("-1,-1", ",")
				
			End If
		End If

		If IsEmpty(aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET)) Then
			If Len(oRequest("ConfineTypeID").Item) > 0 Then
				aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET) = Split(oRequest("ConfineTypeID").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET) = Split("1,1", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_ACTIVITY_ID1_BUDGET)) Then
			If Len(oRequest("ActivityID1").Item) > 0 Then
				aBudgetComponent(N_ACTIVITY_ID1_BUDGET) = Split(oRequest("ActivityID1").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_ACTIVITY_ID1_BUDGET) = Split("-1,-1", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_ACTIVITY_ID2_BUDGET)) Then
			If Len(oRequest("ActivityID2").Item) > 0 Then
				aBudgetComponent(N_ACTIVITY_ID2_BUDGET) = Split(oRequest("ActivityID2").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_ACTIVITY_ID2_BUDGET) = Split("-1,-1", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_PROCESS_ID_BUDGET)) Then
			If Len(oRequest("ProcessID").Item) > 0 Then
				aBudgetComponent(N_PROCESS_ID_BUDGET) = Split(oRequest("ProcessID").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_PROCESS_ID_BUDGET) = Split("-1,-1", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_YEAR_BUDGET)) Then
			If Len(oRequest("BudgetYear").Item) > 0 Then
				aBudgetComponent(N_YEAR_BUDGET) = Split(oRequest("BudgetYear").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_YEAR_BUDGET) = Split(Year(Date()) & "," & Year(Date()), ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(N_MONTH_BUDGET)) Then
			If Len(oRequest("BudgetMonth").Item) > 0 Then
				aBudgetComponent(N_MONTH_BUDGET) = Split(oRequest("BudgetMonth").Item, LIST_SEPARATOR)
			Else
				aBudgetComponent(N_MONTH_BUDGET) = Split("1,1", ",")
			End If
		End If

		If IsEmpty(aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET)) Then
			If Len(oRequest("OriginalAmount_01").Item) > 0 Then
				For iIndex = 1 To 12
					aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) = aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) & "," & oRequest("OriginalAmount_" & Right(("0" & iIndex), Len("00"))).Item
				Next
			Else
				aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) = ",0;;;0,0;;;0,0;;;0,0;;;0,0;;;0,0;;;0,0;;;0,0;;;0,0;;;0,0;;;0,0;;;0,0;;;0"
			End If
		End If
		aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) = Split(aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET), ",")
		For iIndex = 1 To 12
			aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET)(iIndex) = Split(aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET)(iIndex), LIST_SEPARATOR)
		Next

		If IsEmpty(aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET)) Then
			If Len(oRequest("ModifiedAmount_01").Item) > 0 Then
				For iIndex = 1 To 12
					aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) = aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) & "," & oRequest("ModifiedAmount_" & Right(("0" & iIndex), Len("00"))).Item
				Next
			Else
				aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) = ",0;;;0,0;;;0,0;;;0,0;;;0,0;;;0,0;;;0,0;;;0,0;;;0,0;;;0,0;;;0,0;;;0,0;;;0"
			End If
		End If
		aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) = Split(aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET), ",")
		For iIndex = 1 To 12
			aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET)(iIndex) = Split(aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET)(iIndex), LIST_SEPARATOR)
		Next
	End If

	aBudgetComponent(S_QUERY_CONDITION_BUDGET) = ""
	aBudgetComponent(B_CHECK_FOR_DUPLICATED_BUDGET) = True
	aBudgetComponent(B_IS_DUPLICATED_BUDGET) = False

	aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET) = True
	InitializeBudgetComponent = Err.number
	Err.Clear
End Function

Function AddBudget(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new budget into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddBudget"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	If aBudgetComponent(N_ID_BUDGET) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "Budgets", "BudgetID", "", 1, aBudgetComponent(N_ID_BUDGET), sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		If aBudgetComponent(B_CHECK_FOR_DUPLICATED_BUDGET) Then
			lErrorNumber = CheckExistencyOfBudget(aBudgetComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aBudgetComponent(B_IS_DUPLICATED_BUDGET) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un presupuesto con la clave " & aBudgetComponent(S_SHORT_NAME_BUDGET) & " o con el nombre " & aBudgetComponent(S_NAME_BUDGET) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				lErrorNumber = GetBudgetPath(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					If Not CheckBudgetInformationConsistency(aBudgetComponent, sErrorDescription) Then
						lErrorNumber = -1
					Else
						sErrorDescription = "No se pudo guardar la información del nuevo registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Budgets (BudgetID, ParentID, BudgetShortName, BudgetName, BudgetPath, BudgetTypeID, Active) Values (" & aBudgetComponent(N_ID_BUDGET) & ", " & aBudgetComponent(N_PARENT_ID_BUDGET) & ", '" & Replace(aBudgetComponent(S_SHORT_NAME_BUDGET), "'", "") & "', '" & Replace(aBudgetComponent(S_NAME_BUDGET), "'", "´") & "', '" & Replace(aBudgetComponent(S_PATH_BUDGET), "'", "") & "', " & aBudgetComponent(N_BUDGET_TYPE_ID_BUDGET) & ", " & aBudgetComponent(N_ACTIVE_BUDGET) & ")", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
			End If
		End If
	End If

	AddBudget = lErrorNumber
	Err.Clear
End Function

Function AddMoney(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new money record into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddMoney"
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim iIndex

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	If lErrorNumber = 0 Then
		If aBudgetComponent(B_CHECK_FOR_DUPLICATED_BUDGET) Then
			lErrorNumber = CheckExistencyOfMoney(aBudgetComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aBudgetComponent(B_IS_DUPLICATED_BUDGET) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un registro con los campos seleccionados cuyo monto es $" & FormatNumber(aBudgetComponent(D_AMOUNT_BUDGET), 2, True, False, True) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				aBudgetComponent(S_SHORT_NAME_BUDGET) = "."
				aBudgetComponent(S_NAME_BUDGET) = "."
				aBudgetComponent(S_PATH_BUDGET) = ","
				If Not CheckBudgetInformationConsistency(aBudgetComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					For iIndex = 1 To 12
						sErrorDescription = "No se pudo guardar la información del nuevo registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into BudgetsMoney (AreaID, ProgramDutyID, FundID, DutyID, ActiveDutyID, SpecificDutyID, ProgramID, RegionID, BudgetUR, BudgetCT, BudgetAUX, LocationID, BudgetID1, BudgetID2, BudgetID3, ConfineTypeID, ActivityID1, ActivityID2, ProcessID, BudgetYear, BudgetMonth, OriginalAmount, ModifiedAmount, AddDate, ModifyDate, AddUserID, ModifyUserID) Values (" & aBudgetComponent(N_AREA_ID_BUDGET) & ", " & aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET) & ", " & aBudgetComponent(N_FUND_ID_BUDGET) & ", " & aBudgetComponent(N_DUTY_ID_BUDGET) & ", " & aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET) & ", " & aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET) & ", " & aBudgetComponent(N_PROGRAM_ID_BUDGET) & ", " & aBudgetComponent(N_REGION_ID_BUDGET) & ", " & aBudgetComponent(N_UR_BUDGET) & ", " & aBudgetComponent(N_CT_BUDGET) & ", " & aBudgetComponent(N_AUX_BUDGET) & ", " & aBudgetComponent(N_LOCATION_ID_BUDGET) & ", " & aBudgetComponent(N_BUDGET_ID1_BUDGET) & ", " & aBudgetComponent(N_BUDGET_ID2_BUDGET) & ", " & aBudgetComponent(N_BUDGET_ID3_BUDGET) & ", " & aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET) & ", " & aBudgetComponent(N_ACTIVITY_ID1_BUDGET) & ", " & aBudgetComponent(N_ACTIVITY_ID2_BUDGET) & ", " & aBudgetComponent(N_PROCESS_ID_BUDGET) & ", " & aBudgetComponent(N_YEAR_BUDGET) & ", " & iIndex & ", " & aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET)(iIndex) & ", " & aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET)(iIndex) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", 0, " & aLoginComponent(N_USER_ID_LOGIN) & ", -1)", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Next
				End If
			End If
		End If
	End If

	AddMoney = lErrorNumber
	Err.Clear
End Function

Function AddProgram(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new program into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddProgram"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	If aBudgetComponent(N_ID_BUDGET) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "BudgetsAndPrograms", "ProgramID", "", 1, aBudgetComponent(N_ID_BUDGET), sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		If aBudgetComponent(B_CHECK_FOR_DUPLICATED_BUDGET) Then
			lErrorNumber = CheckExistencyOfProgram(aBudgetComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aBudgetComponent(B_IS_DUPLICATED_BUDGET) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un registro en la estructura programática con la clave " & aBudgetComponent(S_SHORT_NAME_BUDGET) & " o con el nombre " & aBudgetComponent(S_NAME_BUDGET) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				lErrorNumber = GetProgramPath(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					If Not CheckBudgetInformationConsistency(aBudgetComponent, sErrorDescription) Then
						lErrorNumber = -1
					Else
						sErrorDescription = "No se pudo guardar la información del nuevo registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into BudgetsAndPrograms (ProgramID, ParentID, ProgramShortName, ProgramName, ProgramPath, ProgramTypeID, BudgetAmount, QttyID) Values (" & aBudgetComponent(N_ID_BUDGET) & ", " & aBudgetComponent(N_PARENT_ID_BUDGET) & ", '" & Replace(aBudgetComponent(S_SHORT_NAME_BUDGET), "'", "") & "', '" & Replace(aBudgetComponent(S_NAME_BUDGET), "'", "´") & "', '" & Replace(aBudgetComponent(S_PATH_BUDGET), "'", "") & "', " & aBudgetComponent(N_BUDGET_TYPE_ID_BUDGET) & ", " & aBudgetComponent(D_AMOUNT_BUDGET) & ", " & aBudgetComponent(N_QTTY_ID_BUDGET) & ")", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
			End If
		End If
	End If

	AddProgram = lErrorNumber
	Err.Clear
End Function

Function GetBudget(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a budget from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetBudget"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	If aBudgetComponent(N_ID_BUDGET) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Budgets Where BudgetID=" & aBudgetComponent(N_ID_BUDGET), "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
				oRecordset.Close
			Else
				aBudgetComponent(N_PARENT_ID_BUDGET) = CLng(oRecordset.Fields("ParentID").Value)
				aBudgetComponent(S_SHORT_NAME_BUDGET) = CStr(oRecordset.Fields("BudgetShortName").Value)
				aBudgetComponent(S_NAME_BUDGET) = CStr(oRecordset.Fields("BudgetName").Value)
				aBudgetComponent(S_PATH_BUDGET) = CStr(oRecordset.Fields("BudgetPath").Value)
				aBudgetComponent(N_BUDGET_TYPE_ID_BUDGET) = CLng(oRecordset.Fields("BudgetTypeID").Value)
				aBudgetComponent(N_ACTIVE_BUDGET) = CInt(oRecordset.Fields("Active").Value)
				oRecordset.Close
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetBudget = lErrorNumber
	Err.Clear
End Function

Function GetMoney(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about the money records from
'         the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetMoney"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	sErrorDescription = "No se pudo obtener la información del registro."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select BudgetMonth, OriginalAmount, ModifiedAmount From BudgetsMoney Where (AreaID=" & aBudgetComponent(N_AREA_ID_BUDGET) & ") And (ProgramDutyID=" & aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET) & ") And (FundID=" & aBudgetComponent(N_FUND_ID_BUDGET) & ") And (DutyID=" & aBudgetComponent(N_DUTY_ID_BUDGET) & ") And (ActiveDutyID=" & aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET) & ") And (SpecificDutyID=" & aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET) & ") And (ProgramID=" & aBudgetComponent(N_PROGRAM_ID_BUDGET) & ") And (RegionID=" & aBudgetComponent(N_REGION_ID_BUDGET) & ") And (BudgetUR=" & aBudgetComponent(N_UR_BUDGET) & ") And (BudgetCT=" & aBudgetComponent(N_CT_BUDGET) & ") And (BudgetAUX=" & aBudgetComponent(N_AUX_BUDGET) & ") And (LocationID=" & aBudgetComponent(N_LOCATION_ID_BUDGET) & ") And (BudgetID1=" & aBudgetComponent(N_BUDGET_ID1_BUDGET) & ") And (BudgetID2=" & aBudgetComponent(N_BUDGET_ID2_BUDGET) & ") And (BudgetID3=" & aBudgetComponent(N_BUDGET_ID3_BUDGET) & ") And (ConfineTypeID=" & aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET) & ") And (ActivityID1=" & aBudgetComponent(N_ACTIVITY_ID1_BUDGET) & ") And (ActivityID2=" & aBudgetComponent(N_ACTIVITY_ID2_BUDGET) & ") And (ProcessID=" & aBudgetComponent(N_PROCESS_ID_BUDGET) & ") And (BudgetYear=" & aBudgetComponent(N_YEAR_BUDGET) & ") Order By BudgetMonth", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "El registro especificado no se encuentra en el sistema."
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			oRecordset.Close
		Else
			aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) = ",0,0,0,0,0,0,0,0,0,0,0,0"
			aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) = ",0,0,0,0,0,0,0,0,0,0,0,0"
			aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) = Split(aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET), ",")
			aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) = Split(aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET), ",")
			Do While Not oRecordset.EOF
				aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET)(CInt(oRecordset.Fields("BudgetMonth").Value)) = CDbl(oRecordset.Fields("OriginalAmount").Value)
				aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET)(CInt(oRecordset.Fields("BudgetMonth").Value)) = CDbl(oRecordset.Fields("ModifiedAmount").Value)
				oRecordset.MoveNext
			Loop
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetMoney = lErrorNumber
	Err.Clear
End Function

Function GetProgram(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a program from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetProgram"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	If aBudgetComponent(N_ID_BUDGET) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From BudgetsAndPrograms Where ProgramID=" & aBudgetComponent(N_ID_BUDGET), "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El registro especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
				oRecordset.Close
			Else
				aBudgetComponent(N_PARENT_ID_BUDGET) = CLng(oRecordset.Fields("ParentID").Value)
				aBudgetComponent(S_SHORT_NAME_BUDGET) = CStr(oRecordset.Fields("ProgramShortName").Value)
				aBudgetComponent(S_NAME_BUDGET) = CStr(oRecordset.Fields("ProgramName").Value)
				aBudgetComponent(S_PATH_BUDGET) = CStr(oRecordset.Fields("ProgramPath").Value)
				aBudgetComponent(N_BUDGET_TYPE_ID_BUDGET) = CLng(oRecordset.Fields("ProgramTypeID").Value)
				aBudgetComponent(D_AMOUNT_BUDGET) = CDbl(oRecordset.Fields("BudgetAmount").Value)
				aBudgetComponent(N_QTTY_ID_BUDGET) = CInt(oRecordset.Fields("QttyID").Value)
				oRecordset.Close
			End If
		End If
	End If

	Set oRecordset = Nothing
	GetProgram = lErrorNumber
	Err.Clear
End Function

Function GetBudgetPath(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To get the path for a budget from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetBudgetPath"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	If aBudgetComponent(N_PARENT_ID_BUDGET) = -1 Then
		aBudgetComponent(S_PATH_BUDGET) = ",-1," & aBudgetComponent(N_ID_BUDGET) & ","
		aBudgetComponent(N_BUDGET_TYPE_ID_BUDGET) = 7
	Else
		sErrorDescription = "No se pudo obtener la ruta del presupuesto."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select BudgetPath, BudgetTypeID From Budgets Where BudgetID=" & aBudgetComponent(N_PARENT_ID_BUDGET), "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				aBudgetComponent(S_PATH_BUDGET) = ",-1," & aBudgetComponent(N_ID_BUDGET) & ","
				aBudgetComponent(N_BUDGET_TYPE_ID_BUDGET) = 7
			Else
				aBudgetComponent(S_PATH_BUDGET) = CStr(oRecordset.Fields("BudgetPath").Value) & aBudgetComponent(N_ID_BUDGET) & ","
				aBudgetComponent(N_BUDGET_TYPE_ID_BUDGET) = CInt(oRecordset.Fields("BudgetTypeID").Value) + 1
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetBudgetPath = lErrorNumber
	Err.Clear
End Function

Function GetProgramParent(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To get the parent for a program from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetProgramParent"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	sErrorDescription = "No se pudo obtener la ruta del presupuesto."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ParentID From BudgetsAndPrograms Where ProgramID=" & aBudgetComponent(N_ID_BUDGET), "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			aBudgetComponent(N_PARENT_ID_BUDGET) = -1
		Else
			aBudgetComponent(N_PARENT_ID_BUDGET) = CLng(oRecordset.Fields("ParentID").Value)
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GetProgramParent = lErrorNumber
	Err.Clear
End Function

Function GetProgramPath(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To get the path for a program from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetProgramPath"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	If aBudgetComponent(N_PARENT_ID_BUDGET) = -1 Then
		If aBudgetComponent(N_ID_BUDGET) = -1 Then
			aBudgetComponent(S_PATH_BUDGET) = ",-1,"
		Else
			aBudgetComponent(S_PATH_BUDGET) = ",-1," & aBudgetComponent(N_ID_BUDGET) & ","
		End If
		aBudgetComponent(N_BUDGET_TYPE_ID_BUDGET) = 0
	Else
		sErrorDescription = "No se pudo obtener la ruta del presupuesto."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ProgramPath, ProgramTypeID From BudgetsAndPrograms Where ProgramID=" & aBudgetComponent(N_PARENT_ID_BUDGET), "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				aBudgetComponent(S_PATH_BUDGET) = ",-1," & aBudgetComponent(N_ID_BUDGET) & ","
				aBudgetComponent(N_BUDGET_TYPE_ID_BUDGET) = 0
			Else
				aBudgetComponent(S_PATH_BUDGET) = CStr(oRecordset.Fields("ProgramPath").Value) & aBudgetComponent(N_ID_BUDGET) & ","
				aBudgetComponent(N_BUDGET_TYPE_ID_BUDGET) = CInt(oRecordset.Fields("ProgramTypeID").Value) + 1
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetProgramPath = lErrorNumber
	Err.Clear
End Function

Function GetBudgets(oRequest, oADODBConnection, aBudgetComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the budgets from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetBudgets"
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	If (Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0) Then 'Or (aLoginComponent(N_PERMISSION_BUDGET_ID_LOGIN) <> -1) Then
		sCondition = Trim(aBudgetComponent(S_QUERY_CONDITION_BUDGET))
		'If aLoginComponent(N_PERMISSION_BUDGET_ID_LOGIN) <> -1 Then
		'	sCondition = Trim(sCondition & " And ((BudgetID=" & aLoginComponent(N_PERMISSION_BUDGET_ID_LOGIN) & ") Or (BudgetPath Like '" & S_WILD_CHAR & "," & aLoginComponent(N_PERMISSION_BUDGET_ID_LOGIN) & "," & S_WILD_CHAR & "')) ")
		'End If
		If InStr(1, sCondition, "And ", vbTextCompare) <> 1 Then
			sCondition = "And " & sCondition
		End If
	End If
	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Budgets.*, BudgetTypeName From Budgets, BudgetTypes Where (Budgets.BudgetTypeID=BudgetTypes.BudgetTypeID) And (BudgetID>-1) " & sCondition & " Order By BudgetShortName, BudgetName", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetBudgets = lErrorNumber
	Err.Clear
End Function

Function GetMoneys(oRequest, oADODBConnection, aBudgetComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the money records
'         from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetMoneys"
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
		sCondition = Trim(aBudgetComponent(S_QUERY_CONDITION_BUDGET))
		If InStr(1, sCondition, "And ", vbTextCompare) <> 1 Then
			sCondition = " And " & sCondition
		End If
		sErrorDescription = "No se pudo obtener la información de los registros."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID, BudgetsMoney.ProgramDutyID, ProgramDutyName, BudgetsMoney.FundID, FundName, BudgetsMoney.DutyID, DutyName, BudgetsMoney.ActiveDutyID, ActiveDutyName, BudgetsMoney.SpecificDutyID, SpecificDutyName, BudgetsMoney.ProgramID, ProgramName, BudgetsMoney.RegionID, Zones1.ZoneCode As ZoneCode1, Zones1.ZoneName As ZoneName1, BudgetUR, BudgetCT, BudgetAUX, BudgetsMoney.LocationID, Zones2.ZoneCode As ZoneCode2, Zones2.ZoneName As ZoneName2, BudgetsMoney.BudgetID1, Budgets1.BudgetShortName As BudgetShortName1, BudgetsMoney.BudgetID2, Budgets2.BudgetShortName As BudgetShortName2, BudgetsMoney.BudgetID3, Budgets3.BudgetShortName As BudgetShortName3, BudgetsMoney.ConfineTypeID, ConfineTypeShortName, BudgetsMoney.ActivityID1, BudgetsActivities1.ActivityName As ActivityName1, BudgetsMoney.ActivityID2, BudgetsActivities2.ActivityName As ActivityName2, BudgetsMoney.ProcessID, ProcessName, BudgetYear, Sum(OriginalAmount) As TotalAmount From BudgetsMoney, BudgetsProgramDuties, BudgetsFunds, BudgetsDuties, BudgetsActiveDuties, BudgetsSpecificDuties, BudgetsPrograms, Zones As Zones1, Zones As Zones2, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3, BudgetsActivities1, BudgetsActivities2, BudgetsProcesses, BudgetsConfineTypes Where (BudgetsMoney.ProgramDutyID=BudgetsProgramDuties.ProgramDutyID) And (BudgetsMoney.FundID=BudgetsFunds.FundID) And (BudgetsMoney.DutyID=BudgetsDuties.DutyID) And (BudgetsMoney.ActiveDutyID=BudgetsActiveDuties.ActiveDutyID) And (BudgetsMoney.SpecificDutyID=BudgetsSpecificDuties.SpecificDutyID) And (BudgetsMoney.ProgramID=BudgetsPrograms.ProgramID) And (BudgetsMoney.RegionID=Zones1.ZoneID) And (BudgetsMoney.LocationID=Zones2.ZoneID) And (BudgetsMoney.BudgetID1=Budgets1.BudgetID) And (BudgetsMoney.BudgetID2=Budgets2.BudgetID) And (BudgetsMoney.BudgetID3=Budgets3.BudgetID) And (BudgetsMoney.ConfineTypeID=BudgetsConfineTypes.ConfineTypeID) And (BudgetsMoney.ActivityID1=BudgetsActivities1.ActivityID) And (BudgetsMoney.ActivityID2=BudgetsActivities2.ActivityID) And (BudgetsMoney.ProcessID=BudgetsProcesses.ProcessID) " & sCondition & " Group By AreaID, BudgetsMoney.ProgramDutyID, ProgramDutyName, BudgetsMoney.FundID, FundName, BudgetsMoney.DutyID, DutyName, BudgetsMoney.ActiveDutyID, ActiveDutyName, BudgetsMoney.SpecificDutyID, SpecificDutyName, BudgetsMoney.ProgramID, ProgramName, BudgetsMoney.RegionID, Zones1.ZoneCode, Zones1.ZoneName, BudgetUR, BudgetCT, BudgetAUX, BudgetsMoney.LocationID, Zones2.ZoneCode, Zones2.ZoneName, BudgetsMoney.BudgetID1, Budgets1.BudgetShortName, BudgetsMoney.BudgetID2, Budgets2.BudgetShortName, BudgetsMoney.BudgetID3, Budgets3.BudgetShortName, BudgetsMoney.ConfineTypeID, ConfineTypeShortName, BudgetsMoney.ActivityID1, BudgetsActivities1.ActivityName, BudgetsMoney.ActivityID2, BudgetsActivities2.ActivityName, BudgetsMoney.ProcessID, ProcessName, BudgetYear Order By AreaID, BudgetsMoney.ProgramDutyID, ProgramDutyName, BudgetsMoney.FundID, FundName, BudgetsMoney.DutyID, DutyName, BudgetsMoney.ActiveDutyID, ActiveDutyName, BudgetsMoney.SpecificDutyID, SpecificDutyName, BudgetsMoney.ProgramID, ProgramName, BudgetsMoney.RegionID, Zones1.ZoneCode, Zones1.ZoneName, BudgetUR, BudgetCT, BudgetAUX, BudgetsMoney.LocationID, Zones2.ZoneCode, Zones2.ZoneName, BudgetsMoney.BudgetID1, Budgets1.BudgetShortName, BudgetsMoney.BudgetID2, Budgets2.BudgetShortName, BudgetsMoney.BudgetID3, Budgets3.BudgetShortName, BudgetsMoney.ConfineTypeID, ConfineTypeShortName, BudgetsMoney.ActivityID1, BudgetsActivities1.ActivityName, BudgetsMoney.ActivityID2, BudgetsActivities2.ActivityName, BudgetsMoney.ProcessID, ProcessName, BudgetYear", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Else
		lErrorNumber = -1
		sErrorDescription = "Favor de especificar los criterios para realizar la búsqueda en los registros del presupuesto."
	End If

	GetMoneys = lErrorNumber
	Err.Clear
End Function

Function GetPrograms(oRequest, oADODBConnection, aBudgetComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the programs from the
'         database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPrograms"
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	If (Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0) Then 'Or (aLoginComponent(N_PERMISSION_BUDGET_ID_LOGIN) <> -1) Then
		sCondition = Trim(aBudgetComponent(S_QUERY_CONDITION_BUDGET))
		'If aLoginComponent(N_PERMISSION_BUDGET_ID_LOGIN) <> -1 Then
		'	sCondition = Trim(sCondition & " And ((ProgramID=" & aLoginComponent(N_PERMISSION_BUDGET_ID_LOGIN) & ") Or (ProgramPath Like '" & S_WILD_CHAR & "," & aLoginComponent(N_PERMISSION_BUDGET_ID_LOGIN) & "," & S_WILD_CHAR & "')) ")
		'End If
		If InStr(1, sCondition, "And ", vbTextCompare) <> 1 Then
			sCondition = "And " & sCondition
		End If
	End If
	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select BudgetsAndPrograms.*, BudgetTypeName From BudgetsAndPrograms, BudgetTypes2 Where (BudgetsAndPrograms.ProgramTypeID=BudgetTypes2.BudgetTypeID) And (ProgramID>-1) " & sCondition & " Order By ProgramShortName, ProgramName", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetPrograms = lErrorNumber
	Err.Clear
End Function

Function ModifyBudget(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing budget in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyBudget"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	If aBudgetComponent(N_ID_BUDGET) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aBudgetComponent(B_CHECK_FOR_DUPLICATED_BUDGET) Then
			lErrorNumber = CheckExistencyOfBudget(aBudgetComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aBudgetComponent(B_IS_DUPLICATED_BUDGET) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un presupuesto con la clave " & aBudgetComponent(S_SHORT_NAME_BUDGET) & " o el nombre " & aBudgetComponent(S_NAME_BUDGET) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				lErrorNumber = GetBudgetPath(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					If Not CheckBudgetInformationConsistency(aBudgetComponent, sErrorDescription) Then
						lErrorNumber = -1
					Else
						sErrorDescription = "No se pudo modificar la información del registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Budgets Set BudgetShortName='" & Replace(aBudgetComponent(S_SHORT_NAME_BUDGET), "'", "") & "', BudgetName='" & Replace(aBudgetComponent(S_NAME_BUDGET), "'", "") & "', BudgetPath='" & Replace(aBudgetComponent(S_PATH_BUDGET), "'", "") & "', BudgetTypeID=" & aBudgetComponent(N_BUDGET_TYPE_ID_BUDGET) & ", Active=" & aBudgetComponent(N_ACTIVE_BUDGET) & " Where (BudgetID=" & aBudgetComponent(N_ID_BUDGET) & ")", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	ModifyBudget = lErrorNumber
	Err.Clear
End Function

Function ModifyMoney(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing money record in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyMoney"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim iIndex

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	aBudgetComponent(S_SHORT_NAME_BUDGET) = "."
	aBudgetComponent(S_NAME_BUDGET) = "."
	aBudgetComponent(S_PATH_BUDGET) = ","
	If Not CheckBudgetInformationConsistency(aBudgetComponent, sErrorDescription) Then
		lErrorNumber = -1
	Else
		For iIndex = 1 To 12
			sErrorDescription = "No se pudo modificar la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update BudgetsMoney Set ModifiedAmount=" & aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET)(iIndex) & ", ModifyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", ModifyUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (AreaID=" & aBudgetComponent(N_AREA_ID_BUDGET) & ") And (ProgramDutyID=" & aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET) & ") And (FundID=" & aBudgetComponent(N_FUND_ID_BUDGET) & ") And (DutyID=" & aBudgetComponent(N_DUTY_ID_BUDGET) & ") And (ActiveDutyID=" & aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET) & ") And (SpecificDutyID=" & aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET) & ") And (ProgramID=" & aBudgetComponent(N_PROGRAM_ID_BUDGET) & ") And (RegionID=" & aBudgetComponent(N_REGION_ID_BUDGET) & ") And (BudgetUR=" & aBudgetComponent(N_UR_BUDGET) & ") And (BudgetCT=" & aBudgetComponent(N_CT_BUDGET) & ") And (BudgetAUX=" & aBudgetComponent(N_AUX_BUDGET) & ") And (LocationID=" & aBudgetComponent(N_LOCATION_ID_BUDGET) & ") And (BudgetID1=" & aBudgetComponent(N_BUDGET_ID1_BUDGET) & ") And (BudgetID2=" & aBudgetComponent(N_BUDGET_ID2_BUDGET) & ") And (BudgetID3=" & aBudgetComponent(N_BUDGET_ID3_BUDGET) & ") And (ConfineTypeID=" & aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET) & ") And (ActivityID1=" & aBudgetComponent(N_ACTIVITY_ID1_BUDGET) & ") And (ActivityID2=" & aBudgetComponent(N_ACTIVITY_ID2_BUDGET) & ") And (ProcessID=" & aBudgetComponent(N_PROCESS_ID_BUDGET) & ") And (BudgetYear=" & aBudgetComponent(N_YEAR_BUDGET) & ") And (BudgetMonth=" & iIndex & ")", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Next
	End If

	Set oRecordset = Nothing
	ModifyMoney = lErrorNumber
	Err.Clear
End Function

Function ModifyProgram(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing program in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyProgram"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	If aBudgetComponent(N_ID_BUDGET) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aBudgetComponent(B_CHECK_FOR_DUPLICATED_BUDGET) Then
			lErrorNumber = CheckExistencyOfProgram(aBudgetComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aBudgetComponent(B_IS_DUPLICATED_BUDGET) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un registro en la estructura programática con la clave " & aBudgetComponent(S_SHORT_NAME_BUDGET) & " o el nombre " & aBudgetComponent(S_NAME_BUDGET) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				lErrorNumber = GetProgramPath(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
				If lErrorNumber = 0 Then
					If Not CheckBudgetInformationConsistency(aBudgetComponent, sErrorDescription) Then
						lErrorNumber = -1
					Else
						sErrorDescription = "No se pudo modificar la información del registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update BudgetsAndPrograms Set ProgramShortName='" & Replace(aBudgetComponent(S_SHORT_NAME_BUDGET), "'", "") & "', ProgramName='" & Replace(aBudgetComponent(S_NAME_BUDGET), "'", "") & "', ProgramPath='" & Replace(aBudgetComponent(S_PATH_BUDGET), "'", "") & "', ProgramTypeID=" & aBudgetComponent(N_BUDGET_TYPE_ID_BUDGET) & ", BudgetAmount=" & aBudgetComponent(D_AMOUNT_BUDGET) & ", QttyID=" & aBudgetComponent(N_QTTY_ID_BUDGET) & " Where (ProgramID=" & aBudgetComponent(N_ID_BUDGET) & ")", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	ModifyProgram = lErrorNumber
	Err.Clear
End Function

Function SetActiveForBudget(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To set the Active field for the given budget
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForBudget"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	If aBudgetComponent(N_ID_BUDGET) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del registro a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo modificar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Budgets Set Active=" & CInt(oRequest("SetActive").Item) & " Where (BudgetID=" & aBudgetComponent(N_ID_BUDGET) & ")", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	SetActiveForBudget = lErrorNumber
	Err.Clear
End Function

Function RemoveBudget(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a budget from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveBudget"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	If aBudgetComponent(N_ID_BUDGET) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el registro a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Budgets Where (BudgetID=" & aBudgetComponent(N_ID_BUDGET) & ")", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la información del registro."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From JobsBudgetsLKP Where (BudgetID=" & aBudgetComponent(N_ID_BUDGET) & ")", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	RemoveBudget = lErrorNumber
	Err.Clear
End Function

Function RemoveMoney(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To remove the money records from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveMoney"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	sErrorDescription = "No se pudo eliminar la información del registro."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From BudgetsMoney Where (AreaID=" & aBudgetComponent(N_AREA_ID_BUDGET) & ") And (ProgramDutyID=" & aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET) & ") And (FundID=" & aBudgetComponent(N_FUND_ID_BUDGET) & ") And (DutyID=" & aBudgetComponent(N_DUTY_ID_BUDGET) & ") And (ActiveDutyID=" & aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET) & ") And (SpecificDutyID=" & aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET) & ") And (ProgramID=" & aBudgetComponent(N_PROGRAM_ID_BUDGET) & ") And (RegionID=" & aBudgetComponent(N_REGION_ID_BUDGET) & ") And (BudgetUR=" & aBudgetComponent(N_UR_BUDGET) & ") And (BudgetCT=" & aBudgetComponent(N_CT_BUDGET) & ") And (BudgetAUX=" & aBudgetComponent(N_AUX_BUDGET) & ") And (LocationID=" & aBudgetComponent(N_LOCATION_ID_BUDGET) & ") And (BudgetID1=" & aBudgetComponent(N_BUDGET_ID1_BUDGET) & ") And (BudgetID2=" & aBudgetComponent(N_BUDGET_ID2_BUDGET) & ") And (BudgetID3=" & aBudgetComponent(N_BUDGET_ID3_BUDGET) & ") And (ConfineTypeID=" & aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET) & ") And (ActivityID1=" & aBudgetComponent(N_ACTIVITY_ID1_BUDGET) & ") And (ActivityID2=" & aBudgetComponent(N_ACTIVITY_ID2_BUDGET) & ") And (ProcessID=" & aBudgetComponent(N_PROCESS_ID_BUDGET) & ") And (BudgetYear=" & aBudgetComponent(N_YEAR_BUDGET) & ")", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

	RemoveMoney = lErrorNumber
	Err.Clear
End Function

Function RemoveProgram(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a program from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveProgram"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	If aBudgetComponent(N_ID_BUDGET) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el registro a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del registro."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From BudgetsAndPrograms Where (ProgramID=" & aBudgetComponent(N_ID_BUDGET) & ")", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	RemoveBudget = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfBudget(aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific budget exists in the database
'Inputs:  aBudgetComponent
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfBudget"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	If Len(aBudgetComponent(S_NAME_BUDGET)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el nombre del registro para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Budgets Where (BudgetID<>" & aBudgetComponent(N_ID_BUDGET) & ") And (ParentID=" & aBudgetComponent(N_PARENT_ID_BUDGET) & ") And ((BudgetShortName='" & Replace(aBudgetComponent(S_SHORT_NAME_BUDGET), "'", "") & "') Or (BudgetName='" & Replace(aBudgetComponent(S_NAME_BUDGET), "'", "") & "'))", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aBudgetComponent(B_IS_DUPLICATED_BUDGET) = True
				aBudgetComponent(N_ID_BUDGET) = -1
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfBudget = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfMoney(aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific money record exists in the
'         database
'Inputs:  aBudgetComponent
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfMoney"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ModifiedAmount From BudgetsMoney Where (AreaID=" & aBudgetComponent(N_AREA_ID_BUDGET) & ") And (ProgramDutyID=" & aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET) & ") And (FundID=" & aBudgetComponent(N_FUND_ID_BUDGET) & ") And (DutyID=" & aBudgetComponent(N_DUTY_ID_BUDGET) & ") And (ActiveDutyID=" & aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET) & ") And (SpecificDutyID=" & aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET) & ") And (ProgramID=" & aBudgetComponent(N_PROGRAM_ID_BUDGET) & ") And (RegionID=" & aBudgetComponent(N_REGION_ID_BUDGET) & ") And (BudgetUR=" & aBudgetComponent(N_UR_BUDGET) & ") And (BudgetCT=" & aBudgetComponent(N_CT_BUDGET) & ") And (BudgetAUX=" & aBudgetComponent(N_AUX_BUDGET) & ") And (LocationID=" & aBudgetComponent(N_LOCATION_ID_BUDGET) & ") And (BudgetID1=" & aBudgetComponent(N_BUDGET_ID1_BUDGET) & ") And (BudgetID2=" & aBudgetComponent(N_BUDGET_ID2_BUDGET) & ") And (BudgetID3=" & aBudgetComponent(N_BUDGET_ID3_BUDGET) & ") And (ConfineTypeID=" & aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET) & ") And (ActivityID1=" & aBudgetComponent(N_ACTIVITY_ID1_BUDGET) & ") And (ActivityID2=" & aBudgetComponent(N_ACTIVITY_ID2_BUDGET) & ") And (ProcessID=" & aBudgetComponent(N_PROCESS_ID_BUDGET) & ") And (BudgetYear=" & aBudgetComponent(N_YEAR_BUDGET) & ") And (BudgetMonth=" & aBudgetComponent(N_MONTH_BUDGET) & ")", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			aBudgetComponent(B_IS_DUPLICATED_BUDGET) = True
			aBudgetComponent(D_AMOUNT_BUDGET) = CDbl(oRecordset.Fields("ModifiedAmount").Value)
		End If
	End If
	oRecordset.Close

	Set oRecordset = Nothing
	CheckExistencyOfMoney = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfProgram(aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific program exists in the database
'Inputs:  aBudgetComponent
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfProgram"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	If Len(aBudgetComponent(S_NAME_BUDGET)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el nombre del registro para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del registro en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From BudgetsAndPrograms Where (ProgramID<>" & aBudgetComponent(N_ID_BUDGET) & ") And (ParentID=" & aBudgetComponent(N_PARENT_ID_BUDGET) & ") And ((ProgramShortName='" & Replace(aBudgetComponent(S_SHORT_NAME_BUDGET), "'", "") & "') Or (ProgramName='" & Replace(aBudgetComponent(S_NAME_BUDGET), "'", "") & "'))", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aBudgetComponent(B_IS_DUPLICATED_BUDGET) = True
				aBudgetComponent(N_ID_BUDGET) = -1
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfProgram = lErrorNumber
	Err.Clear
End Function

Function CheckBudgetInformationConsistency(aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aBudgetComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckBudgetInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Len(oRequest("ModifyMoneys").Item) = 0 Then
		If Not IsNumeric(aBudgetComponent(N_ID_BUDGET)) Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del registro no es un valor numérico."
			bIsCorrect = False
		End If
		If Not IsNumeric(aBudgetComponent(N_PARENT_ID_BUDGET)) Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del presupuesto al que pertenece este presupuesto no es un valor numérico."
			bIsCorrect = False
		End If
		If Len(aBudgetComponent(S_SHORT_NAME_BUDGET)) = 0 Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- La clave del registro está vacía."
			bIsCorrect = False
		End If
		If Len(aBudgetComponent(S_NAME_BUDGET)) = 0 Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre del registro está vacío."
			bIsCorrect = False
		End If
		If Not IsNumeric(aBudgetComponent(N_BUDGET_TYPE_ID_BUDGET)) Then aBudgetComponent(N_BUDGET_TYPE_ID_BUDGET) = -1
		If Len(aBudgetComponent(S_PATH_BUDGET)) = 0 Then
			sErrorDescription = sErrorDescription & "<BR />&nbsp;- La ruta del presupuesto está vacía."
			bIsCorrect = False
		End If
		If Not IsNumeric(aBudgetComponent(N_ACTIVE_BUDGET)) Then aBudgetComponent(N_ACTIVE_BUDGET) = 1
		If Not IsNumeric(aBudgetComponent(D_AMOUNT_BUDGET)) Then aBudgetComponent(D_AMOUNT_BUDGET) = 0
		If Not IsNumeric(aBudgetComponent(N_QTTY_ID_BUDGET)) Then aBudgetComponent(N_QTTY_ID_BUDGET) = 2
	End If

	If Not IsNumeric(aBudgetComponent(N_AREA_ID_BUDGET)) Then aBudgetComponent(N_AREA_ID_BUDGET) = -1
	If Not IsNumeric(aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET)) Then aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET) = -1
	If Not IsNumeric(aBudgetComponent(N_FUND_ID_BUDGET)) Then aBudgetComponent(N_FUND_ID_BUDGET) = -1
	If Not IsNumeric(aBudgetComponent(N_DUTY_ID_BUDGET)) Then aBudgetComponent(N_DUTY_ID_BUDGET) = -1
	If Not IsNumeric(aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET)) Then aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET) = -1
	If Not IsNumeric(aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET)) Then aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET) = -1
	If Not IsNumeric(aBudgetComponent(N_PROGRAM_ID_BUDGET)) Then aBudgetComponent(N_PROGRAM_ID_BUDGET) = -1
	If Not IsNumeric(aBudgetComponent(N_REGION_ID_BUDGET)) Then aBudgetComponent(N_REGION_ID_BUDGET) = -1
	If Not IsNumeric(aBudgetComponent(N_UR_BUDGET)) Then aBudgetComponent(N_UR_BUDGET) = 0
	If Not IsNumeric(aBudgetComponent(N_CT_BUDGET)) Then aBudgetComponent(N_CT_BUDGET) = 0
	If Not IsNumeric(aBudgetComponent(N_AUX_BUDGET)) Then aBudgetComponent(N_AUX_BUDGET) = 0
	If Not IsNumeric(aBudgetComponent(N_LOCATION_ID_BUDGET)) Then aBudgetComponent(N_LOCATION_ID_BUDGET) = -1
	If Not IsNumeric(aBudgetComponent(N_BUDGET_ID1_BUDGET)) Then aBudgetComponent(N_BUDGET_ID1_BUDGET) = -1
	If Not IsNumeric(aBudgetComponent(N_BUDGET_ID2_BUDGET)) Then aBudgetComponent(N_BUDGET_ID2_BUDGET) = -1
	If Not IsNumeric(aBudgetComponent(N_BUDGET_ID3_BUDGET)) Then aBudgetComponent(N_BUDGET_ID3_BUDGET) = -1
	If Not IsNumeric(aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET)) Then aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET) = 1
	If Not IsNumeric(aBudgetComponent(N_ACTIVITY_ID1_BUDGET)) Then aBudgetComponent(N_ACTIVITY_ID1_BUDGET) = -1
	If Not IsNumeric(aBudgetComponent(N_ACTIVITY_ID2_BUDGET)) Then aBudgetComponent(N_ACTIVITY_ID2_BUDGET) = -1
	If Not IsNumeric(aBudgetComponent(N_PROCESS_ID_BUDGET)) Then aBudgetComponent(N_PROCESS_ID_BUDGET) = -1
	If Not IsNumeric(aBudgetComponent(N_YEAR_BUDGET)) Then aBudgetComponent(N_YEAR_BUDGET) = Year(Date())
	If Not IsNumeric(aBudgetComponent(N_MONTH_BUDGET)) Then aBudgetComponent(N_MONTH_BUDGET) = 1

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del registro contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "BudgetComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckBudgetInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayBudgetForm(oRequest, oADODBConnection, sAction, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a budget from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aBudgetComponent
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayBudgetForm"
	Dim alPath
	Dim lErrorNumber

	If (aBudgetComponent(N_ID_BUDGET) <> -1) And (Len(oRequest("View").Item) = 0) Then
		lErrorNumber = GetBudget(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckBudgetFields(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If Len(oRequest("Delete").Item) > 0 Then Response.Write "return true;" & vbNewLine
					Response.Write "if (oForm.BudgetShortName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de introducir la clave del registro.');" & vbNewLine
						Response.Write "oForm.BudgetShortName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (oForm.BudgetName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de introducir el nombre del registro.');" & vbNewLine
						Response.Write "oForm.BudgetName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckBudgetFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""BudgetFrm"" ID=""BudgetFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckBudgetFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Section"" ID=""SectionHdn"" VALUE=""Budget"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetID"" ID=""BudgetIDHdn"" VALUE=""" & aBudgetComponent(N_ID_BUDGET) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ParentID"" ID=""ParentIDHdn"" VALUE=""" & aBudgetComponent(N_PARENT_ID_BUDGET) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetPath"" ID=""BudgetPathHdn"" VALUE=""" & aBudgetComponent(S_PATH_BUDGET) & """ />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clave:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""BudgetShortName"" ID=""BudgetShortNameTxt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & CleanStringForHTML(aBudgetComponent(S_SHORT_NAME_BUDGET)) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Call GetNameFromTable(oADODBConnection, "BudgetPath", aBudgetComponent(N_PARENT_ID_BUDGET), "", "", alPath, "")
						alPath = Split(alPath, ",")
						Select Case UBound(alPath)
							Case 2
								Response.Write "Partida:"
							Case 3
								Response.Write "Subartida:"
							Case 4
								Response.Write "Tipo de pago:"
							Case Else
								Response.Write "Nombre:"
						End Select
					Response.Write "&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""BudgetName"" ID=""BudgetNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & CleanStringForHTML(aBudgetComponent(S_NAME_BUDGET)) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Activo:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" ID=""ActiveRd"" VALUE=""1"""
							If aBudgetComponent(N_ACTIVE_BUDGET) = 1 Then Response.Write " CHECKED=""1"""
						Response.Write " />Sí&nbsp;&nbsp;&nbsp;"
						Response.Write "<INPUT TYPE=""RADIO"" NAME=""Active"" ID=""ActiveRd"" VALUE=""0"""
							If aBudgetComponent(N_ACTIVE_BUDGET) = 0 Then Response.Write " CHECKED=""0"""
						Response.Write " />No&nbsp;&nbsp;&nbsp;"
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE><BR />"

			If (aBudgetComponent(N_ID_BUDGET) = -1) Or (Len(oRequest("View").Item) > 0) Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveBudgetWngDiv']); BudgetFrm.Remove.focus()"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Section=Budget&BudgetID=" & aBudgetComponent(N_ID_BUDGET) & "'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveBudgetWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayBudgetForm = lErrorNumber
	Err.Clear
End Function

Function DisplayMoneyForm(oRequest, oADODBConnection, sAction, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a money record
'		  from the database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aBudgetComponent
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayMoneyForm"
	Dim lErrorNumber
	Dim dTotal
	Dim iIndex

	If (Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0) And (Len(oRequest("View").Item) > 0) Then
		lErrorNumber = GetMoney(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""DisplayOptions2IFrame"" FRAMEBORDER=""0"" WIDTH=""320"" HEIGHT=""0""></IFRAME>"
		Response.Write "<FORM NAME=""MoneyFrm"" ID=""MoneyFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckMoneyFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Section"" ID=""SectionHdn"" VALUE=""Money"" />"

			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>REGISTRO DEL PRESUPUESTO</B><BR /><BR /></FONT>"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Área:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aBudgetComponent(N_AREA_ID_BUDGET)) & "</FONT>"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaID"" ID=""AreaIDHdn"" VALUE=""" & aBudgetComponent(N_AREA_ID_BUDGET) & """ />"
						Else
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""AreaID"" ID=""AreaIDTxt"" SIZE=""5"" MAXLENGTH=""5"" VALUE="""
								If aBudgetComponent(N_AREA_ID_BUDGET) > 0 Then
									Response.Write aBudgetComponent(N_AREA_ID_BUDGET)
								Else
									Response.Write "200"
								End If
							Response.Write """ CLASS=""TextFields"" />"
						End If
					Response.Write "</TD>"
					'Response.Write "<TD COLSPAN=""2""><SELECT NAME=""AreaID"" ID=""AreaIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					'	If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
					'		Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(AreaID=" & aBudgetComponent(N_AREA_ID_BUDGET) & ")", "AreaCode", aBudgetComponent(N_AREA_ID_BUDGET), "Ninguna;;;-1", sErrorDescription)
					'	Else
					'		Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(AreaID>-1) And (ParentID=-1)", "AreaCode", aBudgetComponent(N_AREA_ID_BUDGET), "Ninguna;;;-1", sErrorDescription)
					'	End If
					'Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Programa&nbsp;presupuestario:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""ProgramDutyID"" ID=""ProgramDutyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsProgramDuties", "ProgramDutyID", "ProgramDutyShortName, ProgramDutyName", "(ProgramDutyID=" & aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET) & ")", "ProgramDutyShortName", aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsProgramDuties", "ProgramDutyID", "ProgramDutyShortName, ProgramDutyName", "(ProgramDutyID>-1) And (Active=1)", "ProgramDutyShortName", aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fondo:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""FundID"" ID=""FundIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsFunds", "FundID", "FundShortName, FundName", "(FundID=" & aBudgetComponent(N_FUND_ID_BUDGET) & ")", "FundShortName", aBudgetComponent(N_FUND_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsFunds", "FundID", "FundShortName, FundName", "(FundID>-1) And (Active=1)", "FundShortName", aBudgetComponent(N_FUND_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Función:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""DutyID"" ID=""DutyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsDuties", "DutyID", "DutyShortName, DutyName", "(DutyID=" & aBudgetComponent(N_DUTY_ID_BUDGET) & ")", "DutyShortName", aBudgetComponent(N_DUTY_ID_BUDGET), "Ninguna;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsDuties", "DutyID", "DutyShortName, DutyName", "(DutyID>-1) And (Active=1)", "DutyShortName", aBudgetComponent(N_DUTY_ID_BUDGET), "Ninguna;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Subfunción activa:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""ActiveDutyID"" ID=""ActiveDutyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsActiveDuties", "ActiveDutyID", "ActiveDutyShortName, ActiveDutyName", "(ActiveDutyID=" & aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET) & ")", "ActiveDutyShortName", aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET), "Ninguna;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsActiveDuties", "ActiveDutyID", "ActiveDutyShortName, ActiveDutyName", "(ActiveDutyID>-1) And (Active=1)", "ActiveDutyShortName", aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET), "Ninguna;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Subfunción específica:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""SpecificDutyID"" ID=""SpecificDutyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsSpecificDuties", "SpecificDutyID", "SpecificDutyShortName, SpecificDutyName", "(SpecificDutyID=" & aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET) & ")", "SpecificDutyShortName", aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET), "Ninguna;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsSpecificDuties", "SpecificDutyID", "SpecificDutyShortName, SpecificDutyName", "(SpecificDutyID>-1) And (Active=1)", "SpecificDutyShortName", aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET), "Ninguna;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Programa:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""ProgramID"" ID=""ProgramIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsPrograms", "ProgramID", "ProgramShortName, ProgramName", "(ProgramID=" & aBudgetComponent(N_PROGRAM_ID_BUDGET) & ")", "ProgramShortName", aBudgetComponent(N_PROGRAM_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsPrograms", "ProgramID", "ProgramShortName, ProgramName", "(ProgramID>-1) And (Active=1)", "ProgramShortName", aBudgetComponent(N_PROGRAM_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Región:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""RegionID"" ID=""RegionIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.value != '') {SearchRecord(this.value, 'Zones_Level2', 'DisplayOptions2IFrame', 'MoneyFrm');}"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneCode, ZoneName", "(ZoneID=" & aBudgetComponent(N_REGION_ID_BUDGET) & ")", "ZoneCode", aBudgetComponent(N_REGION_ID_BUDGET), "Ninguna;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneCode, ZoneName", "(ZoneID>-1) And (ParentID=-1) And (Active=1)", "ZoneCode", aBudgetComponent(N_REGION_ID_BUDGET), "Ninguna;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">UR:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(Right(("000" & aBudgetComponent(N_UR_BUDGET)), Len("000"))) & "</FONT>"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetUR"" ID=""BudgetURHdn"" VALUE=""" & aBudgetComponent(N_UR_BUDGET) & """ />"
						Else
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""BudgetUR"" ID=""BudgetURTxt"" SIZE=""3"" MAXLENGTH=""3"" VALUE=""" & CleanStringForHTML(Right(("000" & aBudgetComponent(N_UR_BUDGET)), Len("000"))) & """ CLASS=""TextFields"" />"
						End If
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">CT:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(Right(("000" & aBudgetComponent(N_CT_BUDGET)), Len("000"))) & "</FONT>"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetCT"" ID=""BudgetCTHdn"" VALUE=""" & aBudgetComponent(N_CT_BUDGET) & """ />"
						Else
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""BudgetCT"" ID=""BudgetCTTxt"" SIZE=""3"" MAXLENGTH=""3"" VALUE=""" & CleanStringForHTML(Right(("000" & aBudgetComponent(N_CT_BUDGET)), Len("000"))) & """ CLASS=""TextFields"" />"
						End If
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">AUX:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(Right(("00" & aBudgetComponent(N_AUX_BUDGET)), Len("00"))) & "</FONT>"
							Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetAUX"" ID=""BudgetAUXHdn"" VALUE=""" & aBudgetComponent(N_AUX_BUDGET) & """ />"
						Else
							Response.Write "<INPUT TYPE=""TEXT"" NAME=""BudgetAUX"" ID=""BudgetAUXTxt"" SIZE=""2"" MAXLENGTH=""2"" VALUE=""" & CleanStringForHTML(Right(("00" & aBudgetComponent(N_AUX_BUDGET)), Len("00"))) & """ CLASS=""TextFields"" />"
						End If
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Municipio:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""LocationID"" ID=""LocationIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneCode, ZoneName", "(ZoneID=" & aBudgetComponent(N_LOCATION_ID_BUDGET) & ")", "ZoneCode", aBudgetComponent(N_LOCATION_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneCode, ZoneName", "(ZoneID>-1) And (ParentID=" & aBudgetComponent(N_REGION_ID_BUDGET) & ") And (Active=1)", "ZoneCode", aBudgetComponent(N_LOCATION_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Partida:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""BudgetID1"" ID=""BudgetID1Cmb"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.value != '') {SearchRecord(this.value, 'Budget_Level2', 'DisplayOptions2IFrame', 'MoneyFrm');}"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Budgets", "BudgetID", "BudgetShortName", "(BudgetID=" & aBudgetComponent(N_BUDGET_ID1_BUDGET) & ")", "BudgetShortName", aBudgetComponent(N_BUDGET_ID1_BUDGET), "Ninguna;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Budgets", "BudgetID", "BudgetShortName", "(BudgetID>-1) And (ParentID=-1) And (Active=1)", "BudgetShortName", aBudgetComponent(N_BUDGET_ID1_BUDGET), "Ninguna;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Subpartida:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""BudgetID2"" ID=""BudgetID2Cmb"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.value != '') {SearchRecord(this.value, 'Budget_Level3', 'DisplayOptions2IFrame', 'MoneyFrm');}"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Budgets", "BudgetID", "BudgetShortName", "(BudgetID=" & aBudgetComponent(N_BUDGET_ID2_BUDGET) & ")", "BudgetShortName", aBudgetComponent(N_BUDGET_ID2_BUDGET), "Ninguna;;;-1", sErrorDescription)
						ElseIf aBudgetComponent(N_BUDGET_ID1_BUDGET) > -1 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Budgets", "BudgetID", "BudgetShortName", "(BudgetID>-1) And (ParentID=" & aBudgetComponent(N_BUDGET_ID1_BUDGET) & ") And (Active=1)", "BudgetShortName", aBudgetComponent(N_BUDGET_ID2_BUDGET), "Ninguna;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Budgets", "BudgetID", "BudgetShortName", "(BudgetID>-1) And (ParentID=1103) And (Active=1)", "BudgetShortName", aBudgetComponent(N_BUDGET_ID2_BUDGET), "Ninguna;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de pago:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""BudgetID3"" ID=""BudgetID3Cmb"" SIZE=""1"" CLASS=""Lists"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Budgets", "BudgetID", "BudgetShortName, BudgetName", "(BudgetID=" & aBudgetComponent(N_BUDGET_ID3_BUDGET) & ")", "BudgetShortName", aBudgetComponent(N_BUDGET_ID3_BUDGET), "Ninguno;;;-1", sErrorDescription)
						ElseIf aBudgetComponent(N_BUDGET_ID2_BUDGET) > -1 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Budgets", "BudgetID", "BudgetShortName, BudgetName", "(BudgetID>-1) And (ParentID=" & aBudgetComponent(N_BUDGET_ID2_BUDGET) & ") And (Active=1)", "BudgetShortName", aBudgetComponent(N_BUDGET_ID3_BUDGET), "Ninguno;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Budgets", "BudgetID", "BudgetShortName, BudgetName", "(BudgetID>-1) And (ParentID=10001) And (Active=1)", "BudgetShortName", aBudgetComponent(N_BUDGET_ID3_BUDGET), "Ninguno;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Ámbito:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""ConfineTypeID"" ID=""ConfineTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsConfineTypes", "ConfineTypeID", "ConfineTypeShortName, ConfineTypeName", "(ConfineTypeID=" & aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET) & ")", "ConfineTypeShortName", aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsConfineTypes", "ConfineTypeID", "ConfineTypeShortName, ConfineTypeName", "(ConfineTypeID>-1)", "ConfineTypeShortName", aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Actividad institucional:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""ActivityID1"" ID=""ActivityID1Cmb"" SIZE=""1"" CLASS=""Lists"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsActivities1", "ActivityID", "ActivityShortName, ActivityName", "(ActivityID=" & aBudgetComponent(N_ACTIVITY_ID1_BUDGET) & ")", "ActivityShortName", aBudgetComponent(N_ACTIVITY_ID1_BUDGET), "Ninguna;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsActivities1", "ActivityID", "ActivityShortName, ActivityName", "(ActivityID>-1) And (Active=1)", "ActivityShortName", aBudgetComponent(N_ACTIVITY_ID1_BUDGET), "Ninguna;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Actividad presupuestaria:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""ActivityID2"" ID=""ActivityID2Cmb"" SIZE=""1"" CLASS=""Lists"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsActivities2", "ActivityID", "ActivityShortName, ActivityName", "(ActivityID=" & aBudgetComponent(N_ACTIVITY_ID2_BUDGET) & ")", "ActivityShortName", aBudgetComponent(N_ACTIVITY_ID2_BUDGET), "Ninguna;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsActivities2", "ActivityID", "ActivityShortName, ActivityName", "(ActivityID>-1) And (Active=1)", "ActivityShortName", aBudgetComponent(N_ACTIVITY_ID2_BUDGET), "Ninguna;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Proceso:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""ProcessID"" ID=""ProcessIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsProcesses", "ProcessID", "ProcessShortName, ProcessName", "(ProcessID=" & aBudgetComponent(N_PROCESS_ID_BUDGET) & ")", "ProcessShortName", aBudgetComponent(N_PROCESS_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsProcesses", "ProcessID", "ProcessShortName, ProcessName", "(ProcessID>-1) And (Active=1)", "ProcessShortName", aBudgetComponent(N_PROCESS_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Año:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""BudgetYear"" ID=""BudgetYearCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE=""" & aBudgetComponent(N_YEAR_BUDGET) & """>" & aBudgetComponent(N_YEAR_BUDGET) & "</OPTION>"
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>Mes</B></FONT></TD>"
					Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>Monto original</B></FONT></TD>"
					Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2""><B>Monto modificado</B></FONT></TD>"
				Response.Write "</TR>"
				dTotal = 0
				For iIndex = 1 To 12
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(asMonthNames_es(iIndex)) & "</FONT></TD>"
						Response.Write "<TD>"
							If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
								Response.Write "<FONT FACE=""Arial"" SIZE=""2"">" & FormatNumber(aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET)(iIndex), 2, True, False, True) & "&nbsp;&nbsp;&nbsp;</FONT>"
								Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OriginalAmount_" & Right(("0" & iIndex), Len("00")) & """ ID=""OriginalAmount_" & Right(("0" & iIndex), Len("00")) & "Hdn"" VALUE=""" & aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET)(iIndex) & """ />"
							Else
								Response.Write "<INPUT TYPE=""TEXT"" NAME=""OriginalAmount_" & Right(("0" & iIndex), Len("00")) & """ ID=""OriginalAmount_" & Right(("0" & iIndex), Len("00")) & "Txt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET)(iIndex), 2, True, False, True) & """ CLASS=""TextFields"" />&nbsp;"
							End If
						Response.Write "</TD>"
						Response.Write "<TD>"
							If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
								Response.Write "&nbsp;<INPUT TYPE=""TEXT"" NAME=""ModifiedAmount_" & Right(("0" & iIndex), Len("00")) & """ ID=""ModifiedAmount_" & Right(("0" & iIndex), Len("00")) & "Txt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & FormatNumber(aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET)(iIndex), 2, True, False, True) & """ CLASS=""TextFields"" />"
							Else
								Response.Write "<CENTER><FONT FACE=""Arial"" SIZE=""2"">---</FONT></CENTER>"
							End If
						Response.Write "</TD>"
					Response.Write "</TR>"
					dTotal = dTotal + aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET)(iIndex)
				Next
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Total</B></FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2""><B>" & FormatNumber(dTotal, 2, True, False, True) & "</B></FONT></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE><BR />"

			If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) = 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveMoneyWngDiv']); MoneyFrm.Remove.focus()"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Section=Money'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveMoneyWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "var dOriginalTotal = " & dTotal & ";" & vbNewLine
			Response.Write "function CheckMoneyFields(oForm) {" & vbNewLine
				Response.Write "var dTemp = 0;" & vbNewLine
				Response.Write "var dTempTotal = 0;" & vbNewLine

				Response.Write "if (oForm) {" & vbNewLine
					If Len(oRequest("Delete").Item) > 0 Then Response.Write "return true;" & vbNewLine
					Response.Write "if (! CheckIntegerValue(oForm.BudgetUR, 'el número de la unidad responsable', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "if (! CheckIntegerValue(oForm.BudgetCT, 'el número del centro de trabajo', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "if (! CheckIntegerValue(oForm.BudgetAUX, 'el número auxiliar', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
						Response.Write "return false;" & vbNewLine

					If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) = 0 Then
						For iIndex = 1 To 12
							Response.Write "oForm.OriginalAmount_" & Right(("0" & iIndex), Len("00")) & ".value = oForm.OriginalAmount_" & Right(("0" & iIndex), Len("00")) & ".value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "dTemp = parseFloat(oForm.OriginalAmount_" & Right(("0" & iIndex), Len("00")) & ".value);" & vbNewLine
							Response.Write "if (! CheckFloatValue(oForm.OriginalAmount_" & Right(("0" & iIndex), Len("00")) & ", 'el monto original de " & asMonthNames_es(iIndex) & "', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
								Response.Write "return false;" & vbNewLine
						Next
					Else
						For iIndex = 1 To 12
							Response.Write "oForm.ModifiedAmount_" & Right(("0" & iIndex), Len("00")) & ".value = oForm.ModifiedAmount_" & Right(("0" & iIndex), Len("00")) & ".value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
							Response.Write "dTemp = parseFloat(oForm.ModifiedAmount_" & Right(("0" & iIndex), Len("00")) & ".value);" & vbNewLine
							Response.Write "if (! CheckFloatValue(oForm.ModifiedAmount_" & Right(("0" & iIndex), Len("00")) & ", 'el monto modificado de " & asMonthNames_es(iIndex) & "', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
								Response.Write "return false;" & vbNewLine
							Response.Write "dTempTotal += dTemp;" & vbNewLine
						Next
						Response.Write "dTempTotal = parseInt(dTempTotal * 100) / 100;" & vbNewLine
						Response.Write "if (dTempTotal != dOriginalTotal) {" & vbNewLine
							Response.Write "alert('El total de los presupuestos modificados es de ' + JSFormatNumber(dTempTotal, 2) + ' y no corresponde con el presupuesto original de " & FormatNumber(dTotal, 2, True, False, True) & ".');" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					End If
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckMoneyFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
	End If

	DisplayMoneyForm = lErrorNumber
	Err.Clear
End Function

Function DisplayMoneysForm(oRequest, oADODBConnection, sAction, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about two money records
'		  from the database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aBudgetComponent
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayMoneysForm"
	Dim lErrorNumber
	Dim iIndex

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "var dOriginalGlobalTotal=0;" & vbNewLine
		Response.Write "var dGlobalTotal=0;" & vbNewLine
		Response.Write "var oFrame01 = document.Search01IFrame;" & vbNewLine
		Response.Write "var oFrame02 = document.Search02IFrame;" & vbNewLine
		Response.Write "var bReady = false;" & vbNewLine

		Response.Write "function CheckBudgetMoneyFromIFrame() {" & vbNewLine
			Response.Write "var dTemp = 0;" & vbNewLine
			Response.Write "var dTempTotal1 = 0;" & vbNewLine
			Response.Write "var dTempTotal2 = 0;" & vbNewLine
			For iIndex = 1 To 12
				Response.Write "oFrame01.MoneyFrm.ModifiedAmount_" & Right(("0" & iIndex), Len("00")) & ".value = oFrame01.MoneyFrm.ModifiedAmount_" & Right(("0" & iIndex), Len("00")) & ".value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
				Response.Write "dTemp = parseFloat(oFrame01.MoneyFrm.ModifiedAmount_" & Right(("0" & iIndex), Len("00")) & ".value);" & vbNewLine
				Response.Write "if (! CheckFloatValue(oFrame01.MoneyFrm.ModifiedAmount_" & Right(("0" & iIndex), Len("00")) & ", 'el monto modificado de " & asMonthNames_es(iIndex) & "', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "dTempTotal1 += dTemp;" & vbNewLine

				Response.Write "oFrame02.MoneyFrm.ModifiedAmount_" & Right(("0" & iIndex), Len("00")) & ".value = oFrame02.MoneyFrm.ModifiedAmount_" & Right(("0" & iIndex), Len("00")) & ".value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
				Response.Write "dTemp = parseFloat(oFrame02.MoneyFrm.ModifiedAmount_" & Right(("0" & iIndex), Len("00")) & ".value);" & vbNewLine
				Response.Write "if (! CheckFloatValue(oFrame02.MoneyFrm.ModifiedAmount_" & Right(("0" & iIndex), Len("00")) & ", 'el monto modificado de " & asMonthNames_es(iIndex) & "', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "dTempTotal2 += dTemp;" & vbNewLine
			Next

			Response.Write "dTempTotal1 = parseInt(dTempTotal1 * 100) / 100;" & vbNewLine
			Response.Write "dTempTotal2 = parseInt(dTempTotal2 * 100) / 100;" & vbNewLine
			Response.Write "if ((oFrame01.dOriginalTotal + oFrame02.dOriginalTotal) != (dTempTotal1 + dTempTotal2)) {" & vbNewLine
				Response.Write "alert('El total de los presupuestos modificados es de ' + JSFormatNumber((dTempTotal1 + dTempTotal2), 2) + ' y no corresponde con el presupuesto original de ' + JSFormatNumber((oFrame01.dOriginalTotal + oFrame02.dOriginalTotal), 2) + '.');" & vbNewLine
				Response.Write "return false;" & vbNewLine
			Response.Write "}" & vbNewLine

			Response.Write "document.PrintFrm.ModifiedAmount_01.value = oFrame01.MoneyFrm.ModifiedAmount_01.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_01.value;" & vbNewLine
			Response.Write "document.PrintFrm.ModifiedAmount_02.value = oFrame01.MoneyFrm.ModifiedAmount_02.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_02.value;" & vbNewLine
			Response.Write "document.PrintFrm.ModifiedAmount_03.value = oFrame01.MoneyFrm.ModifiedAmount_03.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_03.value;" & vbNewLine
			Response.Write "document.PrintFrm.ModifiedAmount_04.value = oFrame01.MoneyFrm.ModifiedAmount_04.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_04.value;" & vbNewLine
			Response.Write "document.PrintFrm.ModifiedAmount_05.value = oFrame01.MoneyFrm.ModifiedAmount_05.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_05.value;" & vbNewLine
			Response.Write "document.PrintFrm.ModifiedAmount_06.value = oFrame01.MoneyFrm.ModifiedAmount_06.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_06.value;" & vbNewLine
			Response.Write "document.PrintFrm.ModifiedAmount_07.value = oFrame01.MoneyFrm.ModifiedAmount_07.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_07.value;" & vbNewLine
			Response.Write "document.PrintFrm.ModifiedAmount_08.value = oFrame01.MoneyFrm.ModifiedAmount_08.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_08.value;" & vbNewLine
			Response.Write "document.PrintFrm.ModifiedAmount_09.value = oFrame01.MoneyFrm.ModifiedAmount_09.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_09.value;" & vbNewLine
			Response.Write "document.PrintFrm.ModifiedAmount_10.value = oFrame01.MoneyFrm.ModifiedAmount_10.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_10.value;" & vbNewLine
			Response.Write "document.PrintFrm.ModifiedAmount_11.value = oFrame01.MoneyFrm.ModifiedAmount_11.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_11.value;" & vbNewLine
			Response.Write "document.PrintFrm.ModifiedAmount_12.value = oFrame01.MoneyFrm.ModifiedAmount_12.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_12.value;" & vbNewLine
		Response.Write "} // End of CheckBudgetMoneyFromIFrame" & vbNewLine

		Response.Write "function GetBudgetMoneyFromIFrame() {" & vbNewLine
				Response.Write "if (! bReady) {" & vbNewLine
					Response.Write "document.PrintFrm.AreaID.value = oFrame01.MoneyFrm.AreaID.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.AreaID.value;" & vbNewLine
					Response.Write "document.PrintFrm.ProgramDutyID.value = oFrame01.MoneyFrm.ProgramDutyID.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ProgramDutyID.value;" & vbNewLine
					Response.Write "document.PrintFrm.FundID.value = oFrame01.MoneyFrm.FundID.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.FundID.value;" & vbNewLine
					Response.Write "document.PrintFrm.DutyID.value = oFrame01.MoneyFrm.DutyID.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.DutyID.value;" & vbNewLine
					Response.Write "document.PrintFrm.ActiveDutyID.value = oFrame01.MoneyFrm.ActiveDutyID.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ActiveDutyID.value;" & vbNewLine
					Response.Write "document.PrintFrm.SpecificDutyID.value = oFrame01.MoneyFrm.SpecificDutyID.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.SpecificDutyID.value;" & vbNewLine
					Response.Write "document.PrintFrm.ProgramID.value = oFrame01.MoneyFrm.ProgramID.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ProgramID.value;" & vbNewLine
					Response.Write "document.PrintFrm.RegionID.value = oFrame01.MoneyFrm.RegionID.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.RegionID.value;" & vbNewLine
					Response.Write "document.PrintFrm.BudgetUR.value = oFrame01.MoneyFrm.BudgetUR.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.BudgetUR.value;" & vbNewLine
					Response.Write "document.PrintFrm.BudgetCT.value = oFrame01.MoneyFrm.BudgetCT.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.BudgetCT.value;" & vbNewLine
					Response.Write "document.PrintFrm.BudgetAUX.value = oFrame01.MoneyFrm.BudgetAUX.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.BudgetAUX.value;" & vbNewLine
					Response.Write "document.PrintFrm.LocationID.value = oFrame01.MoneyFrm.LocationID.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.LocationID.value;" & vbNewLine
					Response.Write "document.PrintFrm.BudgetID1.value = oFrame01.MoneyFrm.BudgetID1.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.BudgetID1.value;" & vbNewLine
					Response.Write "document.PrintFrm.BudgetID2.value = oFrame01.MoneyFrm.BudgetID2.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.BudgetID2.value;" & vbNewLine
					Response.Write "document.PrintFrm.BudgetID3.value = oFrame01.MoneyFrm.BudgetID3.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.BudgetID3.value;" & vbNewLine
					Response.Write "document.PrintFrm.ConfineTypeID.value = oFrame01.MoneyFrm.ConfineTypeID.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ConfineTypeID.value;" & vbNewLine
					Response.Write "document.PrintFrm.ActivityID1.value = oFrame01.MoneyFrm.ActivityID1.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ActivityID1.value;" & vbNewLine
					Response.Write "document.PrintFrm.ActivityID2.value = oFrame01.MoneyFrm.ActivityID2.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ActivityID2.value;" & vbNewLine
					Response.Write "document.PrintFrm.ProcessID.value = oFrame01.MoneyFrm.ProcessID.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ProcessID.value;" & vbNewLine
					Response.Write "document.PrintFrm.BudgetYear.value = oFrame01.MoneyFrm.BudgetYear.value + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.BudgetYear.value;" & vbNewLine
					Response.Write "document.PrintFrm.OriginalAmount_01.value = oFrame01.MoneyFrm.ModifiedAmount_01.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '') + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_01.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "document.PrintFrm.OriginalAmount_02.value = oFrame01.MoneyFrm.ModifiedAmount_02.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '') + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_02.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "document.PrintFrm.OriginalAmount_03.value = oFrame01.MoneyFrm.ModifiedAmount_03.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '') + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_03.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "document.PrintFrm.OriginalAmount_04.value = oFrame01.MoneyFrm.ModifiedAmount_04.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '') + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_04.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "document.PrintFrm.OriginalAmount_05.value = oFrame01.MoneyFrm.ModifiedAmount_05.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '') + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_05.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "document.PrintFrm.OriginalAmount_06.value = oFrame01.MoneyFrm.ModifiedAmount_06.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '') + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_06.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "document.PrintFrm.OriginalAmount_07.value = oFrame01.MoneyFrm.ModifiedAmount_07.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '') + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_07.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "document.PrintFrm.OriginalAmount_08.value = oFrame01.MoneyFrm.ModifiedAmount_08.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '') + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_08.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "document.PrintFrm.OriginalAmount_09.value = oFrame01.MoneyFrm.ModifiedAmount_09.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '') + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_09.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "document.PrintFrm.OriginalAmount_10.value = oFrame01.MoneyFrm.ModifiedAmount_10.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '') + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_10.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "document.PrintFrm.OriginalAmount_11.value = oFrame01.MoneyFrm.ModifiedAmount_11.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '') + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_11.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "document.PrintFrm.OriginalAmount_12.value = oFrame01.MoneyFrm.ModifiedAmount_12.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '') + '" & LIST_SEPARATOR & "' + oFrame02.MoneyFrm.ModifiedAmount_12.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
					Response.Write "bReady = true;" & vbNewLine
			Response.Write "}" & vbNewLine
		Response.Write "} // End of GetBudgetMoneyFromIFrame" & vbNewLine

		Response.Write "function CheckIFrames() {" & vbNewLine
			Response.Write "if ((oFrame01.bReady) && (oFrame02.bReady)) {" & vbNewLine
				Response.Write "ShowDisplay(document.all['ModifyMoneysDiv']);" & vbNewLine
				Response.Write "GetBudgetMoneyFromIFrame();" & vbNewLine
			Response.Write "} else {" & vbNewLine
				Response.Write "bReady = false;" & vbNewLine
				Response.Write "HideDisplay(document.all['ModifyMoneysDiv']);" & vbNewLine
			Response.Write "}" & vbNewLine
			Response.Write "window.setTimeout('CheckIFrames()', 1000);" & vbNewLine
		Response.Write "} // End of GetBudgetMoneyFromIFrame" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine

	Response.Write "<FORM NAME=""PrintFrm"" ID=""PrintFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckBudgetMoneyFromIFrame();"">" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AreaID"" ID=""AreaIDHdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProgramDutyID"" ID=""ProgramDutyIDHdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FundID"" ID=""FundIDHdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""DutyID"" ID=""DutyIDHdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ActiveDutyID"" ID=""ActiveDutyIDHdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SpecificDutyID"" ID=""SpecificDutyIDHdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProgramID"" ID=""ProgramIDHdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RegionID"" ID=""RegionIDHdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetUR"" ID=""BudgetURHdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetCT"" ID=""BudgetCTHdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetAUX"" ID=""BudgetAUXHdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""LocationID"" ID=""LocationIDHdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetID1"" ID=""BudgetID1Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetID2"" ID=""BudgetID2Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetID3"" ID=""BudgetID3Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ConfineTypeID"" ID=""ConfineTypeIDHdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ActivityID1"" ID=""ActivityID1Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ActivityID2"" ID=""ActivityID2Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProcessID"" ID=""ProcessIDHdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetYear"" ID=""BudgetYearHdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OriginalAmount_01"" ID=""OriginalAmount_01Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ModifiedAmount_01"" ID=""ModifiedAmount_01Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OriginalAmount_02"" ID=""OriginalAmount_02Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ModifiedAmount_02"" ID=""ModifiedAmount_02Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OriginalAmount_03"" ID=""OriginalAmount_03Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ModifiedAmount_03"" ID=""ModifiedAmount_03Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OriginalAmount_04"" ID=""OriginalAmount_04Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ModifiedAmount_04"" ID=""ModifiedAmount_04Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OriginalAmount_05"" ID=""OriginalAmount_05Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ModifiedAmount_05"" ID=""ModifiedAmount_05Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OriginalAmount_06"" ID=""OriginalAmount_06Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ModifiedAmount_06"" ID=""ModifiedAmount_06Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OriginalAmount_07"" ID=""OriginalAmount_07Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ModifiedAmount_07"" ID=""ModifiedAmount_07Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OriginalAmount_08"" ID=""OriginalAmount_08Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ModifiedAmount_08"" ID=""ModifiedAmount_08Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OriginalAmount_09"" ID=""OriginalAmount_09Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ModifiedAmount_09"" ID=""ModifiedAmount_09Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OriginalAmount_10"" ID=""OriginalAmount_10Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ModifiedAmount_10"" ID=""ModifiedAmount_10Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OriginalAmount_11"" ID=""OriginalAmount_11Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ModifiedAmount_11"" ID=""ModifiedAmount_11Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""OriginalAmount_12"" ID=""OriginalAmount_12Hdn"" />" & vbNewLine
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ModifiedAmount_12"" ID=""ModifiedAmount_12Hdn"" />" & vbNewLine

		Response.Write "<DIV NAME=""ModifyMoneysDiv"" ID=""ModifyMoneysDiv"" STYLE=""display: none"">" & vbNewLine
			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ModifyMoneys"" ID=""ModifyMoneysBtn"" VALUE=""Modificar Presupuestos"" CLASS=""Buttons"" />" & vbNewLine
			If FileExists(Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_56_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & Left(GetSerialNumberForDate(""), Len("00000000")) & ".txt"), sErrorDescription) Then
				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />" & vbNewLine
				Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Imprimir Adecuaciones"" CLASS=""Buttons"" onClick=""OpenNewWindow('Export.asp?Action=ModifiedMoneys&Excel=1&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes')"" />" & vbNewLine
				Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />" & vbNewLine
				Response.Write "<INPUT TYPE=""BUTTON"" VALUE=""Borrar Bitácora de Adecuaciones"" CLASS=""Buttons"" onClick=""OpenNewWindow('Remove.asp?Action=Rep_56&UserID=" & aLoginComponent(N_USER_ID_LOGIN) & "&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "', '', 'RemoveWnd', 320, 240, 'no', 'yes')"" />" & vbNewLine
			End If
		Response.Write "</DIV>" & vbNewLine
	Response.Write "</FORM>" & vbNewLine

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "window.setTimeout('CheckIFrames()', 1000);" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine

	DisplayMoneysForm = Err.number
	Err.Clear
End Function

Function DisplayMoneySearchForm(oRequest, oADODBConnection, sAction, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To display the search form for money records
'Inputs:  oRequest, oADODBConnection, aBudgetComponent
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayMoneySearchForm"
	Dim lErrorNumber
	Dim iIndex

	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckMoneySearchFields(oForm) {" & vbNewLine
				Response.Write "var iCounter = 0;" & vbNewLine

				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "if (oForm.AreaID.value != '')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (oForm.FundID.value != '-1')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (oForm.DutyID.value != '-1')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (oForm.ActiveDutyID.value != '-1')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (oForm.SpecificDutyID.value != '-1')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (oForm.ProgramID.value != '-1')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (oForm.RegionID.value != '-1')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (oForm.BudgetUR.value != '')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (oForm.BudgetCT.value != '')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (oForm.BudgetAUX.value != '')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (oForm.LocationID.value != '-1')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (oForm.BudgetID1.value != '-1')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (oForm.BudgetID2.value != '-1')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (oForm.BudgetID3.value != '-1')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (oForm.ActivityID1.value != '-1')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (oForm.ActivityID2.value != '-1')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (oForm.ProcessID.value != '-1')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (oForm.BudgetYear.value != '-1')" & vbNewLine
						Response.Write "iCounter++;" & vbNewLine
					Response.Write "if (iCounter < 6) {" & vbNewLine
						Response.Write "alert('Favor de seleccionar al menos 6 criterios para la búsqueda.');" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine

					Response.Write "return true;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckMoneyFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""DisplayOptionsIFrame"" FRAMEBORDER=""0"" WIDTH=""320"" HEIGHT=""0""></IFRAME>"
		Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckMoneySearchFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Section"" ID=""SectionHdn"" VALUE=""Money"" />"

			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>BÚSQUEDA DE REGISTROS DEL PRESUPUESTO</B><BR /><BR /></FONT>"
			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Indique los criterios para filtrar los registros del presupuesto:<BR /></FONT><BR />"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Área:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2"">"
						Response.Write "<INPUT TYPE=""TEXT"" NAME=""AreaID"" ID=""AreaIDTxt"" SIZE=""5"" MAXLENGTH=""5"" VALUE="""
								If aBudgetComponent(N_AREA_ID_BUDGET) > 0 Then
									Response.Write aBudgetComponent(N_AREA_ID_BUDGET)
								End If
							Response.Write """ CLASS=""TextFields"" />"
					Response.Write "</TD>"
					'Response.Write "<TD COLSPAN=""2""><SELECT NAME=""AreaID"" ID=""AreaIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					'	If Len(aBudgetComponent(S_QUERY_CONDITION_BUDGET)) > 0 Then
					'		Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(AreaID=" & aBudgetComponent(N_AREA_ID_BUDGET) & ")", "AreaCode", aBudgetComponent(N_AREA_ID_BUDGET), "Ninguna;;;-1", sErrorDescription)
					'	Else
					'		Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Areas", "AreaID", "AreaCode, AreaName", "(AreaID>-1) And (ParentID=-1)", "AreaCode", aBudgetComponent(N_AREA_ID_BUDGET), "Ninguna;;;-1", sErrorDescription)
					'	End If
					'Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Programa&nbsp;presupuestario:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""ProgramDutyID"" ID=""ProgramDutyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsProgramDuties", "ProgramDutyID", "ProgramDutyShortName, ProgramDutyName", "(ProgramDutyID>-1) And (Active=1)", "ProgramDutyShortName", aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fondo:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""FundID"" ID=""FundIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsFunds", "FundID", "FundShortName, FundName", "(FundID>-1) And (Active=1)", "FundShortName", aBudgetComponent(N_FUND_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Función:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""DutyID"" ID=""DutyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsDuties", "DutyID", "DutyShortName, DutyName", "(DutyID>-1) And (Active=1)", "DutyShortName", aBudgetComponent(N_DUTY_ID_BUDGET), "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Subfunción activa:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""ActiveDutyID"" ID=""ActiveDutyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsActiveDuties", "ActiveDutyID", "ActiveDutyShortName, ActiveDutyName", "(ActiveDutyID>-1) And (Active=1)", "ActiveDutyShortName", aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET), "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Subfunción específica:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""SpecificDutyID"" ID=""SpecificDutyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsSpecificDuties", "SpecificDutyID", "SpecificDutyShortName, SpecificDutyName", "(SpecificDutyID>-1) And (Active=1)", "SpecificDutyShortName", aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET), "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Programa:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""ProgramID"" ID=""ProgramIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsPrograms", "ProgramID", "ProgramShortName, ProgramName", "(ProgramID>-1) And (Active=1)", "ProgramShortName", aBudgetComponent(N_PROGRAM_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Región:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""RegionID"" ID=""RegionIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.value != '') {SearchRecord(this.value, 'Zones_Level2', 'DisplayOptionsIFrame', 'SearchFrm');}"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneCode, ZoneName", "(ZoneID>-1) And (ParentID=-1) And (Active=1)", "ZoneCode", aBudgetComponent(N_REGION_ID_BUDGET), "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">UR:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2"">"
						Response.Write "<INPUT TYPE=""TEXT"" NAME=""BudgetUR"" ID=""BudgetURTxt"" SIZE=""3"" MAXLENGTH=""3"" VALUE=""" & CleanStringForHTML(oRequest("BudgetUR").Item) & """ CLASS=""TextFields"" />"
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">CT:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2"">"
						Response.Write "<INPUT TYPE=""TEXT"" NAME=""BudgetCT"" ID=""BudgetCTTxt"" SIZE=""3"" MAXLENGTH=""3"" VALUE=""" & CleanStringForHTML(oRequest("BudgetCT").Item) & """ CLASS=""TextFields"" />"
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">AUX:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2"">"
						Response.Write "<INPUT TYPE=""TEXT"" NAME=""BudgetAUX"" ID=""BudgetAUXTxt"" SIZE=""2"" MAXLENGTH=""2"" VALUE=""" & CleanStringForHTML(oRequest("BudgetAUX").Item) & """ CLASS=""TextFields"" />"
					Response.Write "</TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Municipio:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""LocationID"" ID=""LocationIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						If aBudgetComponent(N_REGION_ID_BUDGET) > -1 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Zones", "ZoneID", "ZoneCode, ZoneName", "(ZoneID>-1) And (ParentID=" & aBudgetComponent(N_REGION_ID_BUDGET) & ") And (Active=1)", "ZoneCode", aBudgetComponent(N_LOCATION_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Partida:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""BudgetID1"" ID=""BudgetID1Cmb"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.value != '') {SearchRecord(this.value, 'Budget_Level2', 'DisplayOptionsIFrame', 'SearchFrm');}"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Budgets", "BudgetID", "BudgetName", "(BudgetID>-1) And (ParentID=-1) And (Active=1)", "BudgetName", aBudgetComponent(N_BUDGET_ID1_BUDGET), "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Subpartida:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""BudgetID2"" ID=""BudgetID2Cmb"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.value != '') {SearchRecord(this.value, 'Budget_Level3', 'DisplayOptionsIFrame', 'SearchFrm');}"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						If aBudgetComponent(N_BUDGET_ID1_BUDGET) > -1 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Budgets", "BudgetID", "BudgetShortName", "(BudgetID>-1) And (ParentID=" & aBudgetComponent(N_BUDGET_ID1_BUDGET) & ") And (Active=1)", "BudgetShortName", aBudgetComponent(N_BUDGET_ID2_BUDGET), "Ninguna;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de pago:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""BudgetID3"" ID=""BudgetID3Cmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						If aBudgetComponent(N_BUDGET_ID2_BUDGET) > -1 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Budgets", "BudgetID", "BudgetShortName", "(BudgetID>-1) And (ParentID=" & aBudgetComponent(N_BUDGET_ID2_BUDGET) & ") And (Active=1)", "BudgetShortName", aBudgetComponent(N_BUDGET_ID2_BUDGET), "Ninguna;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Ámbito:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""ConfineTypeID"" ID=""ConfineTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsConfineTypes", "ConfineTypeID", "ConfineTypeShortName, ConfineTypeName", "(ConfineTypeID>-1)", "ConfineTypeShortName", aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Actividad institucional:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""ActivityID1"" ID=""ActivityID1Cmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsActivities1", "ActivityID", "ActivityShortName, ActivityName", "(ActivityID>-1) And (Active=1)", "ActivityShortName", aBudgetComponent(N_ACTIVITY_ID1_BUDGET), "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Actividad presupuestaria:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""ActivityID2"" ID=""ActivityID2Cmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsActivities2", "ActivityID", "ActivityShortName, ActivityName", "(ActivityID>-1) And (Active=1)", "ActivityShortName", aBudgetComponent(N_ACTIVITY_ID2_BUDGET), "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Proceso:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""ProcessID"" ID=""ProcessIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE="""">Todos</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BudgetsProcesses", "ProcessID", "ProcessShortName, ProcessName", "(ProcessID>-1) And (Active=1)", "ProcessShortName", aBudgetComponent(N_PROCESS_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Año:&nbsp;</FONT></TD>"
					Response.Write "<TD COLSPAN=""2""><SELECT NAME=""BudgetYear"" ID=""BudgetYearCmb"" SIZE=""1"" CLASS=""Lists"">"
						For iIndex = 2010 To Year(Date())
							Response.Write "<OPTION VALUE=""" & iIndex & """>" & iIndex & "</OPTION>"
						Next
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE><BR />"

			Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar"" CLASS=""Buttons"" />"
			'Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			'Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=5'"" />"
			'Response.Write "<BR /><BR />"
		Response.Write "</FORM>"
	End If

	DisplayMoneySearchForm = lErrorNumber
	Err.Clear
End Function

Function DisplayProgramForm(oRequest, oADODBConnection, sAction, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a program from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aBudgetComponent
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayProgramForm"
	Dim lErrorNumber

	If (aBudgetComponent(N_ID_BUDGET) <> -1) And (Len(oRequest("View").Item) = 0) Then
		lErrorNumber = GetProgram(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckProgramFields(oForm) {" & vbNewLine
				Response.Write "if (oForm) {" & vbNewLine
					If Len(oRequest("Delete").Item) > 0 Then Response.Write "return true;" & vbNewLine
					Response.Write "if (oForm.ProgramShortName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de introducir la clave del registro.');" & vbNewLine
						Response.Write "oForm.ProgramShortName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (oForm.ProgramName.value.length == 0) {" & vbNewLine
						Response.Write "alert('Favor de introducir el nombre del registro.');" & vbNewLine
						Response.Write "oForm.ProgramName.focus();" & vbNewLine
						Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
					Response.Write "if (! CheckFloatValue(oForm.BudgetAmount, 'el monto', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
						Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckProgramFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""ProgramFrm"" ID=""ProgramFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckProgramFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Section"" ID=""SectionHdn"" VALUE=""Program"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProgramID"" ID=""ProgramIDHdn"" VALUE=""" & aBudgetComponent(N_ID_BUDGET) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ParentID"" ID=""ParentIDHdn"" VALUE=""" & aBudgetComponent(N_PARENT_ID_BUDGET) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProgramPath"" ID=""ProgramPathHdn"" VALUE=""" & aBudgetComponent(S_PATH_BUDGET) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ProgramTypeID"" ID=""ProgramTypeIDHdn"" VALUE=""" & aBudgetComponent(N_BUDGET_TYPE_ID_BUDGET) & """ />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Clave:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""ProgramShortName"" ID=""ProgramShortNameTxt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & CleanStringForHTML(aBudgetComponent(S_SHORT_NAME_BUDGET)) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""ProgramName"" ID=""ProgramNameTxt"" SIZE=""30"" MAXLENGTH=""100"" VALUE=""" & CleanStringForHTML(aBudgetComponent(S_NAME_BUDGET)) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Monto:&nbsp;</FONT></TD>"
					Response.Write "<TD>"
						Response.Write "<INPUT TYPE=""TEXT"" NAME=""BudgetAmount"" ID=""BudgetAmountTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & CleanStringForHTML(aBudgetComponent(D_AMOUNT_BUDGET)) & """ CLASS=""TextFields"" />&nbsp;"
						Response.Write "<SELECT NAME=""QttyID"" ID=""QttyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "QttyValues", "QttyID", "QttyName", "(QttyID In (1,2))", "QttyID", aBudgetComponent(N_QTTY_ID_BUDGET), "Ninguno;;;-1", sErrorDescription)
						Response.Write "</SELECT>"
					Response.Write "</TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE><BR />"

			If (aBudgetComponent(N_ID_BUDGET) = -1) Or (Len(oRequest("View").Item) > 0) Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemoveProgramWngDiv']); ProgramFrm.Remove.focus()"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Section=Program&ProgramID=" & aBudgetComponent(N_ID_BUDGET) & "'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemoveProgramWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayProgramForm = lErrorNumber
	Err.Clear
End Function

Function DisplayBudgetAsHiddenFields(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a budget using
'		  hidden form fields
'Inputs:  oRequest, oADODBConnection, aBudgetComponent
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayBudgetAsHiddenFields"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetID"" ID=""BudgetIDHdn"" VALUE=""" & aBudgetComponent(N_ID_BUDGET) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ParentID"" ID=""ParentIDHdn"" VALUE=""" & aBudgetComponent(N_PARENT_ID_BUDGET) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetShortName"" ID=""BudgetShortNameHdn"" VALUE=""" & aBudgetComponent(S_SHORT_NAME_BUDGET) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetName"" ID=""BudgetNameHdn"" VALUE=""" & aBudgetComponent(S_NAME_BUDGET) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""BudgetPath"" ID=""BudgetPathHdn"" VALUE=""" & aBudgetComponent(S_PATH_BUDGET) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Active"" ID=""ActiveHdn"" VALUE=""" & aBudgetComponent(N_ACTIVE_BUDGET) & """ />"

	DisplayBudgetAsHiddenFields = Err.number
	Err.Clear
End Function

Function GetMoneyAsURL(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To build a String with the money record information
'		  as URL
'Inputs:  oRequest, oADODBConnection, aBudgetComponent
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetMoneyAsURL"

	GetMoneyAsURL = "AreaID=" & aBudgetComponent(N_AREA_ID_BUDGET) & "&ProgramDutyID=" & aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET) & "&FundID=" & aBudgetComponent(N_FUND_ID_BUDGET) & "&DutyID=" & aBudgetComponent(N_DUTY_ID_BUDGET) & "&ActiveDutyID=" & aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET) & "&SpecificDutyID=" & aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET) & "&ProgramID=" & aBudgetComponent(N_PROGRAM_ID_BUDGET) & "&RegionID=" & aBudgetComponent(N_REGION_ID_BUDGET) & "&BudgetUR=" & aBudgetComponent(N_UR_BUDGET) & "&BudgetCT=" & aBudgetComponent(N_CT_BUDGET) & "&BudgetAUX=" & aBudgetComponent(N_AUX_BUDGET) & "&LocationID=" & aBudgetComponent(N_LOCATION_ID_BUDGET) & "&BudgetID1=" & aBudgetComponent(N_BUDGET_ID1_BUDGET) & "&BudgetID2=" & aBudgetComponent(N_BUDGET_ID2_BUDGET) & "&BudgetID3=" & aBudgetComponent(N_BUDGET_ID3_BUDGET) & "&ConfineTypeID=" & aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET) & "&ActivityID1=" & aBudgetComponent(N_ACTIVITY_ID1_BUDGET) & "&ActivityID2=" & aBudgetComponent(N_ACTIVITY_ID2_BUDGET) & "&ProcessID=" & aBudgetComponent(N_PROCESS_ID_BUDGET) & "&BudgetYear=" & aBudgetComponent(N_YEAR_BUDGET)

	Err.Clear
End Function

Function GetMoneyAsCondition(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To build a String with the money record information
'		  as a query condition
'Inputs:  oRequest, oADODBConnection, aBudgetComponent
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetMoneyAsURL"

	GetMoneyAsCondition = "(AreaID=" & aBudgetComponent(N_AREA_ID_BUDGET) & ") And (ProgramDutyID=" & aBudgetComponent(N_PROGRAM_DUTY_ID_BUDGET) & ") And (FundID=" & aBudgetComponent(N_FUND_ID_BUDGET) & ") And (DutyID=" & aBudgetComponent(N_DUTY_ID_BUDGET) & ") And (ActiveDutyID=" & aBudgetComponent(N_ACTIVE_DUTY_ID_BUDGET) & ") And (SpecificDutyID=" & aBudgetComponent(N_SPECIFIC_DUTY_ID_BUDGET) & ") And (ProgramID=" & aBudgetComponent(N_PROGRAM_ID_BUDGET) & ") And (RegionID=" & aBudgetComponent(N_REGION_ID_BUDGET) & ") And (BudgetUR=" & aBudgetComponent(N_UR_BUDGET) & ") And (BudgetCT=" & aBudgetComponent(N_CT_BUDGET) & ") And (BudgetAUX=" & aBudgetComponent(N_AUX_BUDGET) & ") And (LocationID=" & aBudgetComponent(N_LOCATION_ID_BUDGET) & ") And (BudgetID1=" & aBudgetComponent(N_BUDGET_ID1_BUDGET) & ") And (BudgetID2=" & aBudgetComponent(N_BUDGET_ID2_BUDGET) & ") And (BudgetID3=" & aBudgetComponent(N_BUDGET_ID3_BUDGET) & ") And (ConfineTypeID=" & aBudgetComponent(N_CONFINE_TYPE_ID_BUDGET) & ") And (ActivityID1=" & aBudgetComponent(N_ACTIVITY_ID1_BUDGET) & ") And (ActivityID2=" & aBudgetComponent(N_ACTIVITY_ID2_BUDGET) & ") And (ProcessID=" & aBudgetComponent(N_PROCESS_ID_BUDGET) & ") And (BudgetYear=" & aBudgetComponent(N_YEAR_BUDGET) & ")"

	Err.Clear
End Function

Function DisplayBudgetPath(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To display the path of a budget
'Inputs:  oRequest, oADODBConnection, aBudgetComponent
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayBudgetPath"
	Dim sFullPath
	Dim sTempPath
	Dim lBudgetID
	Dim bFirst
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	sFullPath = ""
	bFirst = True
	lBudgetID = CLng(aBudgetComponent(N_ID_BUDGET))
	Do While (lBudgetID <> -1)
		sErrorDescription = "No se pudo obtener la ruta del presupuesto."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select BudgetID, BudgetShortName, BudgetName, ParentID, BudgetPath From Budgets Where (BudgetID=" & lBudgetID & ")", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				If bFirst Then
					If StrComp(CStr(oRecordset.Fields("BudgetShortName").Value), CStr(oRecordset.Fields("BudgetName").Value), vbBinaryCompare) = 0 Then
						sFullPath = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("BudgetName").Value)) & "</B>" & sFullPath
					Else
						sFullPath = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName").Value) & ". " & CStr(oRecordset.Fields("BudgetName").Value)) & "</B>" & sFullPath
					End If
					bFirst = False
				Else
					sTempPath = "<A "
						'If (aLoginComponent(N_PERMISSION_BUDGET_ID_LOGIN) = CLng(oRecordset.Fields("BudgetID").Value)) Or (InStr(1, CStr(oRecordset.Fields("BudgetPath").Value), ("," & aLoginComponent(N_PERMISSION_BUDGET_ID_LOGIN) & ","), vbBinaryCompare) > 0) Then sTempPath = sTempPath & "HREF=""" & GetASPFileName("") & "?Section=Budget&BudgetID=" & lBudgetID & """"
						sTempPath = sTempPath & "HREF=""" & GetASPFileName("") & "?Section=Budget&BudgetID=" & lBudgetID & "&View=1"""
					If StrComp(CStr(oRecordset.Fields("BudgetShortName").Value), CStr(oRecordset.Fields("BudgetName").Value), vbBinaryCompare) = 0 Then
						sTempPath = sTempPath & ">" & CleanStringForHTML(CStr(oRecordset.Fields("BudgetName").Value)) & "</A> > "
					Else
						sTempPath = sTempPath & ">" & CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName").Value) & ". " & CStr(oRecordset.Fields("BudgetName").Value)) & "</A> > "
					End If
					sFullPath = sTempPath & sFullPath
				End If
				lBudgetID = CLng(oRecordset.Fields("ParentID").Value)
			Else
				lBudgetID = -1
			End If
		Else
			lBudgetID = -1
		End If
	Loop
	Response.Write sFullPath

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayBudgetPath = Err.number
	Err.Clear
End Function

Function DisplayProgramPath(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To display the path of a program
'Inputs:  oRequest, oADODBConnection, aBudgetComponent
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayProgramPath"
	Dim sFullPath
	Dim sTempPath
	Dim lProgramID
	Dim bFirst
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializeBudgetComponent(oRequest, aBudgetComponent)
	End If

	sFullPath = ""
	bFirst = True
	lProgramID = CLng(aBudgetComponent(N_ID_BUDGET))
	Do While (lProgramID <> -1)
		sErrorDescription = "No se pudo obtener la ruta del presupuesto."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ProgramID, ProgramShortName, ProgramName, ParentID, ProgramPath From BudgetsAndPrograms Where (ProgramID=" & lProgramID & ")", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				If bFirst Then
					If StrComp(CStr(oRecordset.Fields("ProgramShortName").Value), CStr(oRecordset.Fields("ProgramName").Value), vbBinaryCompare) = 0 Then
						sFullPath = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("ProgramName").Value)) & "</B>" & sFullPath
					Else
						sFullPath = "<B>" & CleanStringForHTML(CStr(oRecordset.Fields("ProgramShortName").Value) & ". " & CStr(oRecordset.Fields("ProgramName").Value)) & "</B>" & sFullPath
					End If
					bFirst = False
				Else
					sTempPath = "<A "
						'If (aLoginComponent(N_PERMISSION_BUDGET_ID_LOGIN) = CLng(oRecordset.Fields("ProgramID").Value)) Or (InStr(1, CStr(oRecordset.Fields("ProgramPath").Value), ("," & aLoginComponent(N_PERMISSION_BUDGET_ID_LOGIN) & ","), vbBinaryCompare) > 0) Then sTempPath = sTempPath & "HREF=""" & GetASPFileName("") & "?Section=Program&ProgramID=" & lProgramID & """"
						sTempPath = sTempPath & "HREF=""" & GetASPFileName("") & "?Section=Program&ProgramID=" & lProgramID & "&View=1"""
					If StrComp(CStr(oRecordset.Fields("ProgramShortName").Value), CStr(oRecordset.Fields("ProgramName").Value), vbBinaryCompare) = 0 Then
						sTempPath = sTempPath & ">" & CleanStringForHTML(CStr(oRecordset.Fields("ProgramName").Value)) & "</A> > "
					Else
						sTempPath = sTempPath & ">" & CleanStringForHTML(CStr(oRecordset.Fields("ProgramShortName").Value) & ". " & CStr(oRecordset.Fields("ProgramName").Value)) & "</A> > "
					End If
					sFullPath = sTempPath & sFullPath
				End If
				lProgramID = CLng(oRecordset.Fields("ParentID").Value)
			Else
				lProgramID = -1
			End If
		Else
			lProgramID = -1
		End If
	Loop
	Response.Write sFullPath

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayProgramPath = Err.number
	Err.Clear
End Function

Function DisplayBudgetTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the budget from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aBudgetComponent
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayBudgetTable"
	Dim oRecordset
	Dim sNames
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

	lErrorNumber = GetBudgets(oRequest, oADODBConnection, aBudgetComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sNames = CleanStringForHTML(CStr(oRecordset.Fields("BudgetTypeName").Value))
			Response.Write "<TABLE WIDTH=""450"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks Then
					asColumnsTitles = Split("&nbsp;," & sNames & ",Nombre,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,100,250,80", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("&nbsp;," & sNames & ",Nombre", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,100,330", ",", -1, vbBinaryCompare)
				End If
				asCellAlignments = Split(",,,CENTER", ",", -1, vbBinaryCompare)
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
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
					If StrComp(CStr(oRecordset.Fields("BudgetID").Value), oRequest("BudgetID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""BudgetID"" ID=""BudgetIDRd"" VALUE=""" & CStr(oRecordset.Fields("BudgetID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""BudgetID"" ID=""BudgetIDChk"" VALUE=""" & CStr(oRecordset.Fields("BudgetID").Value) & """ />"
						Case Else
							sRowContents = sRowContents & "&nbsp;"
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Section=Budget&BudgetID=" & CStr(oRecordset.Fields("BudgetID").Value) & "&View=1"">" & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName").Value)) & sBoldEnd & sFontEnd & "</A>"
					sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Section=Budget&BudgetID=" & CStr(oRecordset.Fields("BudgetID").Value) & "&View=1"">" & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("BudgetName").Value)) & sBoldEnd & sFontEnd & "</A>"
					If bUseLinks Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Section=Budget&BudgetID=" & CStr(oRecordset.Fields("BudgetID").Value) & "&Change=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If

						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Section=Budget&BudgetID=" & CStr(oRecordset.Fields("BudgetID").Value) & "&Delete=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If

						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							If CInt(oRecordset.Fields("Active").Value) = 0 Then
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Section=Budget&BudgetID=" & CStr(oRecordset.Fields("BudgetID").Value) & "&SetActive=1""><IMG SRC=""Images/BtnActive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Activar"" BORDER=""0"" /></A>"
							Else
								sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Section=Budget&BudgetID=" & CStr(oRecordset.Fields("BudgetID").Value) & "&SetActive=0""><IMG SRC=""Images/BtnDeactive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Desactivar"" BORDER=""0"" /></A>"
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
			sErrorDescription = "No existen registros en la base de datos."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayBudgetTable = lErrorNumber
	Err.Clear
End Function

Function DisplayFullBudgetTable(oRequest, oADODBConnection, bUseLinks, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the budget from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, bUseLinks, bForExport
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFullBudgetTable"
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim sQuery
	Dim sCondition

	sErrorDescription = "No se pudieron obtener las partidas presupuestarias."
	sCondition = ""
	If InStr(oRequest.Item,"ParentID") > 0 Then sCondition = " And (Budgets2.ParentID = " & oRequest("ParentID").Item & ") "
	sQuery = "Select Budgets1.BudgetShortName As BudgetShortName1, Budgets2.BudgetShortName As BudgetShortName2, Budgets3.BudgetID As BudgetID3, Budgets3.BudgetShortName As BudgetShortName3, Budgets3.BudgetName As BudgetName3 From Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3 Where (Budgets3.ParentID=Budgets2.BudgetID) And (Budgets2.ParentID=Budgets1.BudgetID) And (Budgets3.BudgetTypeID=9) " & sCondition & " Order By Budgets1.BudgetShortName, Budgets2.BudgetShortName, Budgets3.BudgetShortName"
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""700"" BORDER="""
				If bForExport Then Response.Write "1"
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks And (Not bForExport) Then
					asColumnsTitles = Split("Partida,Subpartida,<SPAN COLS=""2"" />Tipo de pago,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,100,100,300,100", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("Partida,Subpartida,<SPAN COLS=""2"" />Tipo de pago", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,100,100,400", ",", -1, vbBinaryCompare)
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

				asCellAlignments = Split(",,,,CENTER", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					If bForExport Then
						sRowContents = "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName1").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName2").Value)) & """)"
						sRowContents = sRowContents & TABLE_SEPARATOR & "=T(""" & CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName3").Value)) & """)"
					Else
						sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName1").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName2").Value))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName3").Value))
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BudgetName3").Value))
					If bUseLinks And (Not bForExport) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							sRowContents = sRowContents & "<A HREF=""Budgets.asp?Section=Budget&BudgetID=" & CStr(oRecordset.Fields("BudgetID3").Value) & "&Change=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If

						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & "<A HREF=""Budgets.asp?Section=Budget&BudgetID=" & CStr(oRecordset.Fields("BudgetID3").Value) & "&Delete=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If

						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							If CInt(oRecordset.Fields("Active").Value) = 0 Then
								sRowContents = sRowContents & "<A HREF=""Budgets.asp?Section=Budget&BudgetID=" & CStr(oRecordset.Fields("BudgetID3").Value) & "&SetActive=1""><IMG SRC=""Images/BtnActive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Activar"" BORDER=""0"" /></A>"
							Else
								sRowContents = sRowContents & "<A HREF=""Budgets.asp?Section=Budget&BudgetID=" & CStr(oRecordset.Fields("BudgetID3").Value) & "&SetActive=0""><IMG SRC=""Images/BtnDeactive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Desactivar"" BORDER=""0"" /></A>"
							End If
						End If
						sRowContents = sRowContents & "&nbsp;"
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
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayFullBudgetTable = lErrorNumber
	Err.Clear
End Function

Function DisplayFullProgramTable(oRequest, oADODBConnection, sProgramYear, bUseLinks, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the program from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, ProgramYear, bUseLinks, bForExport
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayFullProgramTable"
	Dim oRecordset
	Dim asProgram
	Dim iIndex
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sErrorDescription = "No se pudieron obtener las partidas presupuestarias."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From BudgetsAndPrograms Where (ProgramPath Like '" & S_WILD_CHAR & "," & sProgramYear & "," & S_WILD_CHAR & "') Order By ProgramPath", "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		asProgram = Split(",,,,,,,", ",")
		If Not oRecordset.EOF Then
			Response.Write "<TABLE WIDTH=""700"" BORDER="""
				If bForExport Then Response.Write "1"
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks And (Not bForExport) Then
					asColumnsTitles = Split("Seguro,Grupo Funcional,Función,Subfunción,Unidad Responsable o Área,Actividad Institucional,Programa Presupuestario,Monto,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("Seguro,Grupo Funcional,Función,Subfunción,Unidad Responsable o Área,Actividad Institucional,Programa Presupuestario,Monto", ",", -1, vbBinaryCompare)
					asCellWidths = Split("110,110,110,110,120,120,120", ",", -1, vbBinaryCompare)
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

				asCellAlignments = Split(",,,,CENTER", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					If StrComp(CStr(oRecordset.Fields("ProgramShortName").Value), CStr(oRecordset.Fields("ProgramName").Value), vbBinaryCompare) = 0 Then
						asProgram(CInt(oRecordset.Fields("ProgramTypeID").Value)) = CStr(oRecordset.Fields("ProgramShortName").Value)
					Else
						asProgram(CInt(oRecordset.Fields("ProgramTypeID").Value)) = CStr(oRecordset.Fields("ProgramShortName").Value) & ". " & CStr(oRecordset.Fields("ProgramName").Value)
					End If
					For iIndex = CInt(oRecordset.Fields("ProgramTypeID").Value) + 1 To UBound(asProgram)
						asProgram(iIndex) = ""
					Next
					sRowContents = ""
					If bForExport Then
						For iIndex = 1 To UBound(asProgram)
							sRowContents = sRowContents & "=T(""" & CleanStringForHTML(asProgram(iIndex)) & """)" & TABLE_SEPARATOR
						Next
					Else
						For iIndex = 1 To UBound(asProgram)
							sRowContents = sRowContents & CleanStringForHTML(asProgram(iIndex)) & TABLE_SEPARATOR
						Next
					End If
					If CInt(oRecordset.Fields("QttyID").Value) = 1 Then
						sRowContents = sRowContents & "$" & FormatNumber(CDbl(oRecordset.Fields("BudgetAmount").Value), 2, True, False, True)
					Else
						sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("BudgetAmount").Value), 2, True, False, True) & "%"
					End If
					If bUseLinks And (Not bForExport) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							sRowContents = sRowContents & "<A HREF=""Budgets.asp?Section=Program&ProgramID=" & CStr(oRecordset.Fields("ProgramID").Value) & "&Change=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If

						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & "<A HREF=""Budgets.asp?Section=Program&ProgramID=" & CStr(oRecordset.Fields("ProgramID").Value) & "&Delete=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If
						sRowContents = sRowContents & "&nbsp;"
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
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayFullProgramTable = lErrorNumber
	Err.Clear
End Function

Function DisplayMoneyTable(oRequest, oADODBConnection, bUseLinks, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the budget  from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, bUseLinks, aBudgetComponent
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayMoneyTable"
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sBoldBegin
	Dim sBoldEnd
	Dim sURL
	Dim lErrorNumber

	lErrorNumber = GetMoneys(oRequest, oADODBConnection, aBudgetComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sNames = CleanStringForHTML(CStr(oRecordset.Fields("BudgetTypeName").Value))
			Response.Write "<TABLE WIDTH=""450"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks Then
					asColumnsTitles = Split("Área,Prog. Presup.,Fondo,Función,Subfunción activa,Subfunción específica,Programa,Región,UR,CT,AUX,Municipio,Partida,Subpartida,Tipo de pago,Ámbito,Act inst,Act pres,Proceso,Año,Monto,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("Área,Prog. Presup.,Fondo,Función,Subfunción activa,Subfunción específica,Programa,Región,UR,CT,AUX,Municipio,Partida,Subpartida,Tipo de pago,Ámbito,Act inst,Act pres,Proceso,Año,Monto", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				End If
				asCellAlignments = Split(",,,,,,,,,,,,,,,,,,,RIGHT,CENTER", ",", -1, vbBinaryCompare)
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If False Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					sRowContents = sRowContents & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("AreaID").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ProgramDutyName").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("FundName").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("DutyName").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ActiveDutyName").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("SpecificDutyName").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ProgramName").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ZoneCode1").Value) & ". " & CStr(oRecordset.Fields("ZoneName1").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("BudgetUR").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("BudgetCT").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("BudgetAUX").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ZoneCode2").Value) & ". " & CStr(oRecordset.Fields("ZoneName2").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName1").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName2").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName3").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ConfineTypeShortName").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ActivityName1").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ActivityName2").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ProcessName").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("BudgetYear").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & FormatNumber(CDbl(oRecordset.Fields("TotalAmount").Value), 2, True, False, True) & sBoldEnd
				
					If bUseLinks Then
						sURL = "&AreaID=" & CStr(oRecordset.Fields("AreaID").Value) & "&ProgramDutyID=" & CStr(oRecordset.Fields("ProgramDutyID").Value) & "&FundID=" & CStr(oRecordset.Fields("FundID").Value) & "&DutyID=" & CStr(oRecordset.Fields("DutyID").Value) & "&ActiveDutyID=" & CStr(oRecordset.Fields("ActiveDutyID").Value) & "&SpecificDutyID=" & CStr(oRecordset.Fields("SpecificDutyID").Value) & "&ProgramID=" & CStr(oRecordset.Fields("ProgramID").Value) & "&RegionID=" & CStr(oRecordset.Fields("RegionID").Value) & "&BudgetUR=" & CStr(oRecordset.Fields("BudgetUR").Value) & "&BudgetCT=" & CStr(oRecordset.Fields("BudgetCT").Value) & "&BudgetAUX=" & CStr(oRecordset.Fields("BudgetAUX").Value) & "&LocationID=" & CStr(oRecordset.Fields("LocationID").Value) & "&BudgetID1=" & CStr(oRecordset.Fields("BudgetID1").Value) & "&BudgetID2=" & CStr(oRecordset.Fields("BudgetID2").Value) & "&BudgetID3=" & CStr(oRecordset.Fields("BudgetID3").Value) & "&ConfineTypeID=" & CStr(oRecordset.Fields("ConfineTypeID").Value) & "&ActivityID1=" & CStr(oRecordset.Fields("ActivityID1").Value) & "&ActivityID2=" & CStr(oRecordset.Fields("ActivityID2").Value) & "&ProcessID=" & CStr(oRecordset.Fields("ProcessID").Value) & "&BudgetYear=" & CStr(oRecordset.Fields("BudgetYear").Value)
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Section=Money" & sURL & "&Change=1&View=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If

						If False And B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Section=Money" & sURL & "&Delete=1&View=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
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
			Call DisplayErrorMessage("Mensaje del sistema", "No existen registros en la base de datos.")
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayMoneyTable = lErrorNumber
	Err.Clear
End Function

Function DisplayProgramTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, aBudgetComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the program from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aBudgetComponent
'Outputs: aBudgetComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayProgramTable"
	Dim oRecordset
	Dim sNames
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

	lErrorNumber = GetPrograms(oRequest, oADODBConnection, aBudgetComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sNames = CleanStringForHTML(CStr(oRecordset.Fields("BudgetTypeName").Value))
			Response.Write "<TABLE WIDTH=""450"" BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks Then
					asColumnsTitles = Split("&nbsp;," & sNames & ",Nombre,Monto,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,100,150,100,80", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("&nbsp;," & sNames & ",Nombre,Monto", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,100,230,100", ",", -1, vbBinaryCompare)
				End If
				asCellAlignments = Split(",,,RIGHT,CENTER", ",", -1, vbBinaryCompare)
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				Do While Not oRecordset.EOF
					sFontBegin = ""
					sFontEnd = ""
'					If CInt(oRecordset.Fields("Active").Value) = 0 Then
'						sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
'						sFontEnd = "</FONT>"
'					End If
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("ProgramID").Value), oRequest("ProgramID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""ProgramID"" ID=""ProgramIDRd"" VALUE=""" & CStr(oRecordset.Fields("ProgramID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""ProgramID"" ID=""ProgramIDChk"" VALUE=""" & CStr(oRecordset.Fields("ProgramID").Value) & """ />"
						Case Else
							sRowContents = sRowContents & "&nbsp;"
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Section=Program&ProgramID=" & CStr(oRecordset.Fields("ProgramID").Value) & "&View=1"">" & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ProgramShortName").Value)) & sBoldEnd & sFontEnd & "</A>"
					sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""" & GetASPFileName("") & "?Section=Program&ProgramID=" & CStr(oRecordset.Fields("ProgramID").Value) & "&View=1"">" & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("ProgramName").Value)) & sBoldEnd & sFontEnd & "</A>"
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
						If CInt(oRecordset.Fields("QttyID").Value) = 1 Then
							sRowContents = sRowContents & "$" & FormatNumber(CDbl(oRecordset.Fields("BudgetAmount").Value), 2, True, False, True)
						Else
							sRowContents = sRowContents & FormatNumber(CDbl(oRecordset.Fields("BudgetAmount").Value), 2, True, False, True) & "%"
						End If
					sRowContents = sRowContents & sBoldEnd & sFontEnd
					If bUseLinks Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "&nbsp;"
						If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
							sRowContents = sRowContents & "<A HREF=""Budget.asp?Section=Program&ProgramID=" & CStr(oRecordset.Fields("ProgramID").Value) & "&Change=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
						End If

						If B_DELETE And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
							sRowContents = sRowContents & "<A HREF=""Budget.asp?Section=Program&ProgramID=" & CStr(oRecordset.Fields("ProgramID").Value) & "&Delete=1"">"
								sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
							sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
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
			sErrorDescription = "No existen registros en la base de datos."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayProgramTable = lErrorNumber
	Err.Clear
End Function

Function PrintModifiedMoneys(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the modified
'		  money records
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "PrintModifiedMoneys"
	Dim oRecordset
	Dim asContents
	Dim adTotals
	Dim iIndex
	Dim jIndex
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	asContents = GetFileContents(Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_56_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & Left(GetSerialNumberForDate(""), Len("00000000")) & ".txt"), sErrorDescription)
	If Len(asContents) > 0 Then
		Response.Write "<TABLE BORDER=""1"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
			asColumnsTitles = Split("<SPAN COLS=""15"" />CLAVE PRESUPUESTARIA " & Year(Date()), ",", -1, vbBinaryCompare)
			lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			asColumnsTitles = Split("GF,FN,SP,PG,AI,PP,R,UR,CT,AUX,MUN,FD,PARTIDA,SUB PTDA,TIPO PAGO", ",", -1, vbBinaryCompare)
			lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			asColumnsTitles = Split("<SPAN COLS=""3"" />&nbsp;,<SPAN COLS=""13"" />CALENDARIO PRESUPUESTAL (PERSOS Y CENTAVOS)", ",", -1, vbBinaryCompare)
			lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			asColumnsTitles = Split("<SPAN COLS=""3"" />&nbsp;,ENERO,FEBRERO,MARZO,ABRIL,MAYO,JUNIO,JULIO,AGOSTO,SEPTIEMBRE,OCTUBRE,NOVIEMBRE,DICIEMBRE,TOTAL", ",", -1, vbBinaryCompare)
			lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
			asRowContents = Split("", TABLE_SEPARATOR, -1, vbBinaryCompare)
			lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)

			adTotals = Split(",", ",")
			adTotals(0) = Split(",0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
			adTotals(1) = Split(",0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
			For iIndex = 0 To UBound(adTotals)
				For jIndex = 1 To UBound(adTotals(iIndex))
					adTotals(iIndex)(jIndex) = 0
				Next
			Next
			asContents = Split(asContents, vbNewLine)
			For iIndex = 0 To UBound(asContents) Step 2
				If Len(asContents(iIndex)) > 0 Then
					sErrorDescription = "No se pudo obtener la información del registro."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID, ProgramDutyShortName, FundShortName, DutyShortName, ActiveDutyShortName, SpecificDutyShortName, ProgramShortName, Zones1.ZoneCode As ZoneCode1, BudgetUR, BudgetCT, BudgetAUX, Zones2.ZoneCode As ZoneCode2, Budgets1.BudgetShortName As BudgetShortName1, Budgets2.BudgetShortName As BudgetShortName2, Budgets3.BudgetShortName As BudgetShortName3, ConfineTypeShortName, BudgetsActivities1.ActivityShortName As ActivityShortName1, BudgetsActivities2.ActivityShortName As ActivityShortName2, ProcessShortName From BudgetsMoney, BudgetsProgramDuties, BudgetsFunds, BudgetsDuties, BudgetsActiveDuties, BudgetsSpecificDuties, BudgetsPrograms, Zones As Zones1, Zones As Zones2, Budgets As Budgets1, Budgets As Budgets2, Budgets As Budgets3, BudgetsActivities1, BudgetsActivities2, BudgetsProcesses, BudgetsConfineTypes Where (BudgetsMoney.ProgramDutyID=BudgetsProgramDuties.ProgramDutyID) And (BudgetsMoney.FundID=BudgetsFunds.FundID) And (BudgetsMoney.DutyID=BudgetsDuties.DutyID) And (BudgetsMoney.ActiveDutyID=BudgetsActiveDuties.ActiveDutyID) And (BudgetsMoney.SpecificDutyID=BudgetsSpecificDuties.SpecificDutyID) And (BudgetsMoney.ProgramID=BudgetsPrograms.ProgramID) And (BudgetsMoney.RegionID=Zones1.ZoneID) And (BudgetsMoney.LocationID=Zones2.ZoneID) And (BudgetsMoney.BudgetID1=Budgets1.BudgetID) And (BudgetsMoney.BudgetID2=Budgets2.BudgetID) And (BudgetsMoney.BudgetID3=Budgets3.BudgetID) And (BudgetsMoney.ConfineTypeID=BudgetsConfineTypes.ConfineTypeID) And (BudgetsMoney.ActivityID1=BudgetsActivities1.ActivityID) And (BudgetsMoney.ActivityID2=BudgetsActivities2.ActivityID) And (BudgetsMoney.ProcessID=BudgetsProcesses.ProcessID) And " & Replace(asContents(iIndex), "(", "(BudgetsMoney."), "BudgetComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If Not oRecordset.EOF Then
							sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("ProcessShortName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("DutyShortName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ActiveDutyShortName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ProgramShortName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ActivityShortName1").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ProgramDutyShortName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ZoneCode1").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BudgetUR").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BudgetCT").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BudgetAUX").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ZoneCode2").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("FundShortName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName1").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName2").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BudgetShortName3").Value))
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)

							sRowContents = "<SPAN COLS=""3"" />&nbsp;" & TABLE_SEPARATOR & asContents(iIndex + 1)
							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							If (iIndex Mod 4) = 0 Then
								For jIndex = 1 To UBound(adTotals(0))
									adTotals(0)(jIndex) = adTotals(0)(jIndex) + CDbl(Replace(Replace(asRowContents(jIndex), "<B>", ""), "</B>", ""))
								Next
							Else
								For jIndex = 1 To UBound(adTotals(1))
									adTotals(1)(jIndex) = adTotals(1)(jIndex) + CDbl(Replace(Replace(asRowContents(jIndex), "<B>", ""), "</B>", ""))
								Next
							End If

							asRowContents = Split("", TABLE_SEPARATOR, -1, vbBinaryCompare)
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						End If
					End If
					If (lErrorNumber <> 0) And (Err.number <> 0) Then Exit For
				End If
			Next
			For iIndex = 0 To UBound(adTotals)
				If iIndex = 0 Then
					sRowContents = "<SPAN COLS=""3"" />Total Reducciones"
				Else
					sRowContents = "<SPAN COLS=""3"" />Total Ampliaciones"
				End If
				For jIndex = 1 To UBound(adTotals(0))
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(adTotals(iIndex)(jIndex), 2, True, False, True)
				Next
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)

				asRowContents = Split("", TABLE_SEPARATOR, -1, vbBinaryCompare)
				lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
			Next
		Response.Write "</TABLE>" & vbNewLine
	Else
		lErrorNumber = L_ERR_NO_RECORDS
		sErrorDescription = "No existen adecuaciones registradas en el presupuesto."
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	PrintModifiedMoneys = lErrorNumber
	Err.Clear
End Function
%>