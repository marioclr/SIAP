<%
Function GetBudgetURLValues(oRequest, sSection, bAction, sCondition)
'************************************************************
'Purpose: To initialize the global variables using the URL
'Inputs:  oRequest
'Outputs: sSection, bAction, sCondition
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetBudgetURLValues"

	sSection = oRequest("Section").Item
	If Len(oRequest("ModifyMoneys").Item) > 0 Then sSection = "Money"
	If Len(sSection) = 0 Then sSection = "Budget"
	bAction = (Len(oRequest("Add").Item) > 0) Or (Len(oRequest("Modify").Item) > 0) Or (Len(oRequest("ModifyMoneys").Item) > 0) Or (Len(oRequest("Remove").Item) > 0) Or (Len(oRequest("SetActive").Item) > 0)

	sCondition = ""
	If StrComp(sSection, "Money", vbBinaryCompare) = 0 Then
		If Len(oRequest("AreaID").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.AreaID=" & oRequest("AreaID").Item & ")"
		If Len(oRequest("ProgramDutyID").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.ProgramDutyID=" & oRequest("ProgramDutyID").Item & ")"
		If Len(oRequest("FundID").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.FundID=" & oRequest("FundID").Item & ")"
		If Len(oRequest("DutyID").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.DutyID=" & oRequest("DutyID").Item & ")"
		If Len(oRequest("ActiveDutyID").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.ActiveDutyID=" & oRequest("ActiveDutyID").Item & ")"
		If Len(oRequest("SpecificDutyID").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.SpecificDutyID=" & oRequest("SpecificDutyID").Item & ")"
		If Len(oRequest("ProgramID").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.ProgramID=" & oRequest("ProgramID").Item & ")"
		If Len(oRequest("RegionID").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.RegionID=" & oRequest("RegionID").Item & ")"
		If Len(oRequest("BudgetUR").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.BudgetUR=" & oRequest("BudgetUR").Item & ")"
		If Len(oRequest("BudgetCT").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.BudgetCT=" & oRequest("BudgetCT").Item & ")"
		If Len(oRequest("BudgetAUX").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.BudgetAUX=" & oRequest("BudgetAUX").Item & ")"
		If Len(oRequest("LocationID").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.LocationID=" & oRequest("LocationID").Item & ")"
		If Len(oRequest("BudgetID1").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.BudgetID1=" & oRequest("BudgetID1").Item & ")"
		If Len(oRequest("BudgetID2").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.BudgetID2=" & oRequest("BudgetID2").Item & ")"
		If Len(oRequest("BudgetID3").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.BudgetID3=" & oRequest("BudgetID3").Item & ")"
		If Len(oRequest("ConfineTypeID").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.ConfineTypeID=" & oRequest("ConfineTypeID").Item & ")"
		If Len(oRequest("ActivityID1").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.ActivityID1=" & oRequest("ActivityID1").Item & ")"
		If Len(oRequest("ActivityID2").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.ActivityID2=" & oRequest("ActivityID2").Item & ")"
		If Len(oRequest("ProcessID").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.ProcessID=" & oRequest("ProcessID").Item & ")"
		If Len(oRequest("BudgetYear").Item) > 0 Then sCondition = sCondition & " And (BudgetsMoney.BudgetYear=" & oRequest("BudgetYear").Item & ")"
	End If

	GetBudgetURLValues = Err.number
	Err.Clear
End Function

Function DoBudgetAction(oRequest, oADODBConnection, sSection, sAction, sErrorDescription)
'************************************************************
'Purpose: To add, change or delete the information of the
'         specified component
'Inputs:  oRequest, oADODBConnection, sSection, sAction
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DoBudgetAction"
	Dim oRecordset
	Dim sNames
	Dim dTotal
	Dim aTempBudgetComponent()
	Dim iIndex
	Dim lErrorNumber

	Select Case sSection
		Case "Budget"
			If Len(oRequest("Add").Item) > 0 Then
				lErrorNumber = AddBudget(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
			ElseIf Len(oRequest("Modify").Item) > 0 Then
				lErrorNumber = ModifyBudget(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
			ElseIf Len(oRequest("Remove").Item) > 0 Then
				If aBudgetComponent(N_ID_BUDGET) > -1 Then
					lErrorNumber = GetBudget(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
				End If
				lErrorNumber = RemoveBudget(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
			ElseIf Len(oRequest("SetActive").Item) > 0 Then
				If aBudgetComponent(N_ID_BUDGET) > -1 Then
					lErrorNumber = GetBudget(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
				End If
				lErrorNumber = SetActiveForBudget(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
			End If
			If lErrorNumber = 0 Then Response.Redirect "Budget.asp?BudgetID=" & aBudgetComponent(N_PARENT_ID_BUDGET) & "&View=1"
		Case "Money"
			If Len(oRequest("Add").Item) > 0 Then
				lErrorNumber = AddMoney(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
			ElseIf Len(oRequest("Modify").Item) > 0 Then
				lErrorNumber = ModifyMoney(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
			ElseIf Len(oRequest("ModifyMoneys").Item) > 0 Then
				Redim aTempBudgetComponent(N_BUDGET_COMPONENT_SIZE)
				For iIndex = N_AREA_ID_BUDGET To N_MONTH_BUDGET
					aTempBudgetComponent(iIndex) = CLng(aBudgetComponent(iIndex)(0))
				Next
				aTempBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) = ""
				aTempBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) = ""
				For iIndex = 1 To 12
					aTempBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) = aTempBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) & LIST_SEPARATOR & aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET)(iIndex)(0)
					aTempBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) = aTempBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) & LIST_SEPARATOR & aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET)(iIndex)(0)
				Next
				aTempBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) = Split(aTempBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET), LIST_SEPARATOR)
				aTempBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) = Split(aTempBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET), LIST_SEPARATOR)
				aTempBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET) = True
				lErrorNumber = ModifyMoney(oRequest, oADODBConnection, aTempBudgetComponent, sErrorDescription)
				lErrorNumber = AppendTextToFile(Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_56_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & Left(GetSerialNumberForDate(""), Len("00000000")) & ".txt"), GetMoneyAsCondition(oRequest, oADODBConnection, aTempBudgetComponent, sErrorDescription), sErrorDescription)
				sNames = ""
				dTotal = 0
				For iIndex = 1 To 12
					sNames = sNames & FormatNumber(Abs(aTempBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET)(iIndex) - aTempBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET)(iIndex)), 2, True, False, True) & TABLE_SEPARATOR
					dTotal = dTotal + (aTempBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET)(iIndex) - aTempBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET)(iIndex))
				Next
				sNames = sNames & "<B>" & FormatNumber(Abs(dTotal), 2, True, False, True) & "</B>"
				lErrorNumber = AppendTextToFile(Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_56_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & Left(GetSerialNumberForDate(""), Len("00000000")) & ".txt"), sNames, sErrorDescription)

				Redim aTempBudgetComponent(N_BUDGET_COMPONENT_SIZE)
				For iIndex = N_AREA_ID_BUDGET To N_MONTH_BUDGET
					aTempBudgetComponent(iIndex) = CLng(aBudgetComponent(iIndex)(1))
				Next
				aTempBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) = ""
				aTempBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) = ""
				For iIndex = 1 To 12
					aTempBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) = aTempBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) & LIST_SEPARATOR & aBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET)(iIndex)(1)
					aTempBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) = aTempBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) & LIST_SEPARATOR & aBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET)(iIndex)(1)
				Next
				aTempBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET) = Split(aTempBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET), LIST_SEPARATOR)
				aTempBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET) = Split(aTempBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET), LIST_SEPARATOR)
				aTempBudgetComponent(B_COMPONENT_INITIALIZED_BUDGET) = True
				lErrorNumber = ModifyMoney(oRequest, oADODBConnection, aTempBudgetComponent, sErrorDescription)
				lErrorNumber = AppendTextToFile(Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_56_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & Left(GetSerialNumberForDate(""), Len("00000000")) & ".txt"), GetMoneyAsCondition(oRequest, oADODBConnection, aTempBudgetComponent, sErrorDescription), sErrorDescription)
				sNames = ""
				dTotal = 0
				For iIndex = 1 To 12
					sNames = sNames & FormatNumber(Abs(aTempBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET)(iIndex) - aTempBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET)(iIndex)), 2, True, False, True) & TABLE_SEPARATOR
					dTotal = dTotal + (aTempBudgetComponent(AD_ORIGINAL_AMOUNT_BUDGET)(iIndex) - aTempBudgetComponent(AD_MODIFIED_AMOUNT_BUDGET)(iIndex))
				Next
				sNames = sNames & "<B>" & FormatNumber(Abs(dTotal), 2, True, False, True) & "</B>"
				lErrorNumber = AppendTextToFile(Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_56_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & Left(GetSerialNumberForDate(""), Len("00000000")) & ".txt"), sNames, sErrorDescription)
			ElseIf Len(oRequest("Remove").Item) > 0 Then
				If aBudgetComponent(N_ID_BUDGET) > -1 Then
					lErrorNumber = GetMoney(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
				End If
				lErrorNumber = RemoveMoney(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
			End If
		Case "Program"
			If Len(oRequest("Add").Item) > 0 Then
				lErrorNumber = AddProgram(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
			ElseIf Len(oRequest("Modify").Item) > 0 Then
				lErrorNumber = ModifyProgram(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
			ElseIf Len(oRequest("Remove").Item) > 0 Then
				If aBudgetComponent(N_ID_BUDGET) > -1 Then
					lErrorNumber = GetProgram(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
				End If
				lErrorNumber = RemoveProgram(oRequest, oADODBConnection, aBudgetComponent, sErrorDescription)
			End If
			If lErrorNumber = 0 Then Response.Redirect "Budget.asp?Section=Program&ProgramID=" & aBudgetComponent(N_PARENT_ID_BUDGET) & "&View=1"
	End Select

	Set oRecordset = Nothing
	DoBudgetAction = lErrorNumber
	Err.Clear
End Function
%>