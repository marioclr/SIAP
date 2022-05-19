<!-- #include file="PayrollComponentConstants.asp" -->
<%
Function DoCalculations(aPayrollComponent, bRetroactive, bAdjustment, sErrorDescription)
'************************************************************
'Purpose: To calculate the payroll and save it into the database
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DoCalculations"
	Const ROWS_PER_FILE = 10000
	Const CONCEPTS_FOR_FACTOR = "1,3,13,14,38,49,89"
	Dim lPayID
	Dim alAntiquities
	Dim aiDays
	Dim asEmployeesQueries
	Dim iCounter
	Dim iCounter2
	Dim iIndex
	Dim jIndex
	Dim kIndex
	Dim sPeriods
	Dim sFilePath
	Dim asFileContents
	Dim asSpecialConcepts
	Dim lStartDate
	Dim lEndDate
	Dim lTempStartDate
	Dim lTempEndDate
	Dim bCurrent
	Dim bTemp
	Dim bMinMaxApplied
	Dim sQueryBegin
	Dim sQueryEnd
	Dim sCondition
	Dim sConceptCondition
	Dim sTable
	Dim lCurrentID
	Dim lCurrentID2
	Dim sCurrentID
	Dim iCurrentZoneID
	Dim iCurrentZoneTypeID
	Dim adDSM
	Dim bMonthlyTaxes
	Dim lEmployeeTypeID
	Dim adTaxes
	Dim adAllowances
	Dim adTaxInvertions
	Dim adTotal
	Dim dAmount
	Dim dTaxAmount
	Dim dAmount_55
	Dim dAmount_88
	Dim sEmployeesFor44
	Dim sEmployeeIDs
	Dim dTemp
	Dim sTemp
	Dim bTruncate
	Dim sTruncate
	Dim oRecordset
	Dim lErrorNumber

	bTruncate = False
	sTruncate = ""
	sFilePath = Server.MapPath("Export\Payroll_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & aPayrollComponent(N_ID_PAYROLL))
	aPayrollComponent(N_FOR_DATE_PAYROLL) = CLng(aPayrollComponent(N_FOR_DATE_PAYROLL))
	sTable = "Payroll_" & aPayrollComponent(N_ID_PAYROLL)
	If aPayrollComponent(N_TYPE_ID_PAYROLL) = 3 Then sTable = "Payroll_" & aPayrollComponent(N_FOR_DATE_PAYROLL)

	lTempEndDate = Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))
	Select Case Mid(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("00000"), Len("00"))
		Case "01"
			lTempEndDate = (CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1) & "1231"
		Case "02", "04", "06", "08", "09"
			lTempEndDate = lTempEndDate & "0" & (CInt(Mid(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("00000"), Len("00"))) - 1) & "31"
		Case "11"
			lTempEndDate = lTempEndDate & "1031"
		Case "03"
			lTempEndDate = lTempEndDate & "0228"
		Case "05", "07", "10"
			lTempEndDate = lTempEndDate & "0" & (CInt(Mid(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("00000"), Len("00"))) - 1) & "30"
		Case "12"
			lTempEndDate = lTempEndDate & "1130"
	End Select
	lTempEndDate = CLng(lTempEndDate)

	lPayID = CLng(aPayrollComponent(N_FOR_DATE_PAYROLL))
	If bRetroactive Then
		sFilePath = sFilePath & "_R" & aPayrollComponent(N_FOR_DATE_PAYROLL)
		lPayID = CLng(aPayrollComponent(N_FOR_DATE_PAYROLL))
	End If
	sPeriods = ""
	sPeriods = GetPeriodsForPayroll(aPayrollComponent(N_ID_PAYROLL), aPayrollComponent(N_FOR_DATE_PAYROLL), -1)
	bMonthlyTaxes = (CInt(Right(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("00"))) >= 28) Or (CInt(Right(aPayrollComponent(N_ID_PAYROLL), Len("0000"))) = 106)

	If Not bAdjustment Then
Call DisplayTimeStamp("START: LEVEL 2. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		Call BuildCondition(sCondition, "")
		If bRetroactive And (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) Then sCondition = " And (EmployeesHistoryList.EmployeeID=EmployeesRevisions.EmployeeID) And (EmployeesRevisions.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesRevisions.StartPayrollID=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")" & sCondition

Call DisplayTimeStamp("START: LEVEL 2, INSERT RECORDS, EmployeesChanges")
		If Not bTimeout Then
			sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (-" & lPayID & "," & lPayID & "))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write "<!-- Query: Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (-" & lPayID & "," & lPayID & ")) -->" & vbNewLine
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
				lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollID, '" & lPayID & "' As PayrollDate, Max(EmployeeDate) As EmployeeDate1, " & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & " As FirstDate, " & aPayrollComponent(N_FOR_DATE_PAYROLL) & " As LastDate, 0 As Concepts40 From EmployeesHistoryList, StatusEmployees, Reasons Where (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesHistoryList.EndDate>=" & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & ") And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.EmployeeTypeID>-1) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) Group By EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
				Response.Write "<!-- Query: Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollID, '" & lPayID & "' As PayrollDate, Max(EmployeeDate) As EmployeeDate1, " & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & " As FirstDate, " & aPayrollComponent(N_FOR_DATE_PAYROLL) & " As LastDate, 0 As Concepts40 From EmployeesHistoryList, StatusEmployees, Reasons Where (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesHistoryList.EndDate>=" & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & ") And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.EmployeeTypeID>-1) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) Group By EmployeeID -->" & vbNewLine
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID Not In (Select Employees.EmployeeID From BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Jobs, Companies, Areas, Positions, EmployeeTypes, PositionTypes, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones, ZoneTypes Where (EmployeesChangesLKP.EmployeeID=BankAccounts.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.JobID=Jobs.JobID) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (BankAccounts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (BankAccounts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Companies.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Areas.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Areas.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Zones.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Zones.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Positions.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Positions.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Levels.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Levels.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (GroupGradeLevels.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (GroupGradeLevels.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (PaymentCenters.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (PaymentCenters.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Response.Write "<!-- Query: Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID Not In (Select Employees.EmployeeID From BankAccounts, Employees, EmployeesChangesLKP, EmployeesHistoryList, Jobs, Companies, Areas, Positions, EmployeeTypes, PositionTypes, Levels, GroupGradeLevels, Areas As PaymentCenters, Zones, ZoneTypes Where (EmployeesChangesLKP.EmployeeID=BankAccounts.EmployeeID) And (Employees.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.JobID=Jobs.JobID) And (EmployeesHistoryList.CompanyID=Companies.CompanyID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EmployeesHistoryList.PositionID=Positions.PositionID) And (EmployeesHistoryList.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryList.PositionTypeID=PositionTypes.PositionTypeID) And (EmployeesHistoryList.LevelID=Levels.LevelID) And (EmployeesHistoryList.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryList.PaymentCenterID=PaymentCenters.AreaID) And (Zones.ZoneTypeID=ZoneTypes.ZoneTypeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (BankAccounts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (BankAccounts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (BankAccounts.Active=1) And (Companies.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Companies.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Areas.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Areas.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Zones.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Zones.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Positions.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Positions.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Levels.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Levels.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (GroupGradeLevels.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (GroupGradeLevels.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (PaymentCenters.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (PaymentCenters.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & "))) -->" & vbNewLine
				End If
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP, Absences Where (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And (Absences.AbsenceTypeID2=0) And (EmployeesAbsencesLKP.OcurredDate<=" & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & ") And (EmployeesAbsencesLKP.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesAbsencesLKP.Active=1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Response.Write "<!-- Query: Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP, Absences Where (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And (Absences.AbsenceTypeID2=0) And (EmployeesAbsencesLKP.OcurredDate<=" & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & ") And (EmployeesAbsencesLKP.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesAbsencesLKP.Active=1))) -->" & vbNewLine
				End If
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
					lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollID, -PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40 From EmployeesChangesLKP Where (PayrollDate=" & lPayID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
					Response.Write "<!-- Query: Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollID, -PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40 From EmployeesChangesLKP Where (PayrollDate=" & lPayID & ") -->" & vbNewLine
				End If
			End If
		End If

		Call BuildCondition("", sQueryBegin)
		If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
		If Not bTimeout Then
			sErrorDescription = "No se pudieron obtener los días de inactividad de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_Factors", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write "<!-- Query: Delete From Payroll_Factors -->" & vbNewLine
		End If

Call DisplayTimeStamp("START: LEVEL 2. Insert Into Payroll_Factors. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		If Not bTimeout Then
			sErrorDescription = "No se pudieron obtener los días de inactividad de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_Factors (EmployeeID, PayrollFactor) Select EmployeesHistoryList.EmployeeID, 15 As PayrollFactor From EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons" & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) " & sCondition, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write "<!-- Query: Insert Into Payroll_Factors (EmployeeID, PayrollFactor) Select EmployeesHistoryList.EmployeeID, 15 As PayrollFactor From EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons" & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) " & sCondition & " -->" & vbNewLine
		End If

Call DisplayTimeStamp("START: LEVEL 2. Update Antiquities 1 y 2. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		If Not bTimeout Then
			sErrorDescription = "No se pudieron obtener las antigüedades de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Antiquities Where (AntiquityID>-1)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				alAntiquities = ""
				Do While Not oRecordset.EOF
					alAntiquities = alAntiquities & CStr(oRecordset.Fields("AntiquityID").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("StartYears").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EndYears").Value) & LIST_SEPARATOR
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				alAntiquities = Left(alAntiquities, (Len(alAntiquities) - Len(LIST_SEPARATOR)))
				oRecordset.Close
			End If
			alAntiquities = Split(alAntiquities, LIST_SEPARATOR)
			For iIndex = 0 To UBound(alAntiquities)
				alAntiquities(iIndex) = Split(alAntiquities(iIndex), SECOND_LIST_SEPARATOR)
				alAntiquities(iIndex)(0) = CInt(alAntiquities(iIndex)(0))
				alAntiquities(iIndex)(1) = CInt(alAntiquities(iIndex)(1))
				alAntiquities(iIndex)(2) = CInt(alAntiquities(iIndex)(2))
			Next
		End If

		iCounter = 0
		If Not bTimeout Then
			aiDays = Split("0,0", ",")
			For iIndex = 0 To UBound(aiDays)
				aiDays(iIndex) = 0
			Next
			sErrorDescription = "No se pudieron obtener las antigüedades de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate, EmployeesHistoryList.JobID, StatusEmployees.Active, Reasons.ActiveEmployeeID From EmployeesHistoryList, EmployeesChangesLKP, StatusEmployees, Reasons " & sQueryBegin & " Where (EmployeesHistoryList.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (StatusEmployees.Active=1) And (ActiveEmployeeID=1) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") " & sCondition & " Order By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeDate Desc", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					sEmployeesFor44 = ""
					lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
					Do While Not oRecordset.EOF
						If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
							For iIndex = 0 To UBound(alAntiquities)
								If ((aiDays(1) / 365) >= alAntiquities(iIndex)(1)) And ((aiDays(1) / 365) < alAntiquities(iIndex)(2)) Then
									If alAntiquities(iIndex)(0) >= 8 Then
										If ((aiDays(1) >= 9125) And (aiDays(1) <= 9139)) Or ((aiDays(1) >= 10950) And (aiDays(1) <= 10964)) Then sEmployeesFor44 = sEmployeesFor44 & lCurrentID & ","
									End If
									'sErrorDescription = "No se pudo actualizar la antigüedad del empleado."
									'lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update Employees Set AntiquityID=" & alAntiquities(iIndex)(0) & ", Antiquity2ID=-" & aiDays(1) & ", Antiquity3ID=-" & aiDays(0) & " Where (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
									lErrorNumber = AppendTextToFile(sFilePath & "_PayrollAntiquity_" & Int(iCounter / ROWS_PER_FILE) & ".txt", lCurrentID & "," & alAntiquities(iIndex)(0) & "," & aiDays(1) & "," & aiDays(0), sErrorDescription)
									iCounter = iCounter + 1
									Exit For
								End If
							Next
							For iIndex = 0 To UBound(aiDays)
								aiDays(iIndex) = 0
							Next
							lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
						End If
						If CLng(oRecordset.Fields("EndDate").Value) > lTempEndDate Then
							aiDays(0) = aiDays(0) + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)), GetDateFromSerialNumber(lTempEndDate))) + 1
						Else
							aiDays(0) = aiDays(0) + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1
						End If
						If CLng(oRecordset.Fields("EndDate").Value) > aPayrollComponent(N_FOR_DATE_PAYROLL) Then
							aiDays(1) = aiDays(1) + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)), GetDateFromSerialNumber(aPayrollComponent(N_FOR_DATE_PAYROLL)))) + 1
						Else
							aiDays(1) = aiDays(1) + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1
						End If
						oRecordset.MoveNext
						'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
					oRecordset.Close
					For iIndex = 0 To UBound(alAntiquities)
						If ((aiDays(1) / 365) >= alAntiquities(iIndex)(1)) And ((aiDays(1) / 365) < alAntiquities(iIndex)(2)) Then
							If alAntiquities(iIndex)(0) >= 8 Then
								If ((aiDays(1) >= 9125) And (aiDays(1) <= 9139)) Or ((aiDays(1) >= 10950) And (aiDays(1) <= 10964)) Then sEmployeesFor44 = sEmployeesFor44 & lCurrentID & ","
							End If
							'sErrorDescription = "No se pudo actualizar la antigüedad del empleado."
							'lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update Employees Set AntiquityID=" & alAntiquities(iIndex)(0) & ", Antiquity2ID=-" & aiDays(1) & ", Antiquity3ID=-" & aiDays(0) & " Where (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
							lErrorNumber = AppendTextToFile(sFilePath & "_PayrollAntiquity_" & Int(iCounter / ROWS_PER_FILE) & ".txt", lCurrentID & "," & alAntiquities(iIndex)(0) & "," & aiDays(1) & "," & aiDays(0), sErrorDescription)
							iCounter = iCounter + 1
							Exit For
						End If
					Next
				End If
			End If
		End If

		If Not bTimeout Then
			If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 2. RUN FROM FILES, Update Employees.AntiquityID " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
				sTemp = "Update Employees Set AntiquityID=<ANTIQUITY_ID />, Antiquity2ID=-<ANTIQUITY2_ID />, Antiquity3ID=-<ANTIQUITY3_ID /> Where (EmployeeID="
				sQueryEnd = ")"
				For jIndex = 0 To iCounter Step ROWS_PER_FILE
					asFileContents = GetFileContents(sFilePath & "_PayrollAntiquity_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
					If Len(asFileContents) > 0 Then
						asFileContents = Split(asFileContents, vbNewLine)
						For iIndex = 0 To UBound(asFileContents)
							If Len(asFileContents(iIndex)) > 0 Then
								asEmployeesQueries = Split(asFileContents(iIndex), ",")
								sErrorDescription = "No se pudo actualizar la antigüedad del empleado."
								lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, Replace(Replace(Replace(sTemp, "<ANTIQUITY_ID />", asEmployeesQueries(1)), "<ANTIQUITY2_ID />", asEmployeesQueries(2)), "<ANTIQUITY3_ID />", asEmployeesQueries(3)) & asEmployeesQueries(0) & sQueryEnd, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
							End If
							'If lErrorNumber <> 0 Then Exit For
							If bTimeout Then Exit For
						Next
					End If
					Call DeleteFile(sFilePath & "_PayrollAntiquity_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
				Next
			End If
		End If

		iCounter = 0
		aiDays = Split("0,0", ",")
		For iIndex = 0 To UBound(aiDays)
			aiDays(iIndex) = 0
		Next
		If Not bTimeout Then
			sErrorDescription = "No se pudieron obtener las antigüedades de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesAntiquitiesLKP.EmployeeID, AntiquityYears, AntiquityMonths, AntiquityDays From EmployeesAntiquitiesLKP, EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons " & sQueryBegin & " Where (EmployeesAntiquitiesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) " & sCondition & " Order By EmployeesAntiquitiesLKP.EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
					Do While Not oRecordset.EOF
						If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
							'sErrorDescription = "No se pudo actualizar la antigüedad del empleado."
							'lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update Employees Set Antiquity2ID=Antiquity2ID-" & aiDays(1) & " Where (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
							lErrorNumber = AppendTextToFile(sFilePath & "_PayrollAntiquity2_" & Int(iCounter / ROWS_PER_FILE) & ".txt", lCurrentID & "," & aiDays(1), sErrorDescription)
							iCounter = iCounter + 1
							For iIndex = 0 To UBound(aiDays)
								aiDays(iIndex) = 0
							Next
							lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
						End If
						aiDays(1) = aiDays(1) + (CInt(oRecordset.Fields("AntiquityYears").Value) * 365) + Int(CInt(oRecordset.Fields("AntiquityMonths").Value) * 30.4) + CInt(oRecordset.Fields("AntiquityDays").Value)
						oRecordset.MoveNext
						'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
					oRecordset.Close
					'sErrorDescription = "No se pudo actualizar la antigüedad del empleado."
					'lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update Employees Set Antiquity2ID=Antiquity2ID-" & aiDays(1) & " Where (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
					lErrorNumber = AppendTextToFile(sFilePath & "_PayrollAntiquity2_" & Int(iCounter / ROWS_PER_FILE) & ".txt", lCurrentID & "," & aiDays(1), sErrorDescription)
					iCounter = iCounter + 1
				End If
			End If
		End If

		If Not bTimeout Then
			If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 2. RUN FROM FILES, Update Employees.Antiquity2ID- " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
				sTemp = "Update Employees Set Antiquity2ID=Antiquity2ID-<ANTIQUITY2_ID />, Antiquity3ID=Antiquity3ID-<ANTIQUITY2_ID /> Where (EmployeeID="
				sQueryEnd = ")"
				For jIndex = 0 To iCounter Step ROWS_PER_FILE
					asFileContents = GetFileContents(sFilePath & "_PayrollAntiquity2_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
					If Len(asFileContents) > 0 Then
						asFileContents = Split(asFileContents, vbNewLine)
						For iIndex = 0 To UBound(asFileContents)
							If Len(asFileContents(iIndex)) > 0 Then
								asEmployeesQueries = Split(asFileContents(iIndex), ",")
								sErrorDescription = "No se pudo actualizar la antigüedad del empleado."
								lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, Replace(sTemp, "<ANTIQUITY2_ID />", asEmployeesQueries(1)) & asEmployeesQueries(0) & sQueryEnd, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
							End If
							'If lErrorNumber <> 0 Then Exit For
							If bTimeout Then Exit For
						Next
					End If
					Call DeleteFile(sFilePath & "_PayrollAntiquity2_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
				Next
			End If
		End If

		iCounter = 0
		aiDays = Split("0,0", ",")
		For iIndex = 0 To UBound(aiDays)
			aiDays(iIndex) = 0
		Next
		If Not bTimeout Then
			sErrorDescription = "No se pudo obtener la información de los registros."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesAbsencesLKP.EmployeeID, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate From EmployeesAbsencesLKP, EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons " & sQueryBegin & " Where (EmployeesAbsencesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesAbsencesLKP.AbsenceID In (10,95)) And (EmployeesAbsencesLKP.OcurredDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") " & sCondition & " Order By EmployeesAbsencesLKP.EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
					Do While Not oRecordset.EOF
						If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
							'sErrorDescription = "No se pudo actualizar la antigüedad del empleado."
							'lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update Employees Set Antiquity2ID=Antiquity2ID+" & aiDays(1) & ", Antiquity3ID=Antiquity3ID+" & aiDays(0) & " Where (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
							lErrorNumber = AppendTextToFile(sFilePath & "_PayrollAntiquity3_" & Int(iCounter / ROWS_PER_FILE) & ".txt", lCurrentID & "," & aiDays(1) & "," & aiDays(0), sErrorDescription)
							iCounter = iCounter + 1
							For iIndex = 0 To UBound(aiDays)
								aiDays(iIndex) = 0
							Next
							lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
						End If

						If CLng(oRecordset.Fields("EndDate").Value) > lTempEndDate Then
							aiDays(0) = aiDays(0) + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value)), lTempEndDate)) + 1
						Else
							aiDays(0) = aiDays(0) + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1
						End If
						If CLng(oRecordset.Fields("EndDate").Value) > aPayrollComponent(N_FOR_DATE_PAYROLL) Then
							aiDays(1) = aiDays(1) + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value)), aPayrollComponent(N_FOR_DATE_PAYROLL))) + 1
						Else
							aiDays(1) = aiDays(1) + Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1
						End If
						oRecordset.MoveNext
						'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
					oRecordset.Close
					'sErrorDescription = "No se pudo actualizar la antigüedad del empleado."
					'lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update Employees Set Antiquity2ID=Antiquity2ID+" & aiDays(1) & ", Antiquity3ID=Antiquity3ID+" & aiDays(0) & " Where (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
					lErrorNumber = AppendTextToFile(sFilePath & "_PayrollAntiquity3_" & Int(iCounter / ROWS_PER_FILE) & ".txt", lCurrentID & "," & aiDays(1) & "," & aiDays(0), sErrorDescription)
					iCounter = iCounter + 1
				End If
			End If
		End If

		If Not bTimeout Then
			If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 2. RUN FROM FILES, Update Employees.Antiquity2ID+ " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
				sTemp = "Update Employees Set Antiquity2ID=Antiquity2ID+<ANTIQUITY2_ID />, Antiquity3ID=Antiquity3ID+<ANTIQUITY3_ID /> Where (EmployeeID="
				sQueryEnd = ")"
				For jIndex = 0 To iCounter Step ROWS_PER_FILE
					asFileContents = GetFileContents(sFilePath & "_PayrollAntiquity3_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
					If Len(asFileContents) > 0 Then
						asFileContents = Split(asFileContents, vbNewLine)
						For iIndex = 0 To UBound(asFileContents)
							If Len(asFileContents(iIndex)) > 0 Then
								asEmployeesQueries = Split(asFileContents(iIndex), ",")
								sErrorDescription = "No se pudo actualizar la antigüedad del empleado."
								lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, Replace(Replace(sTemp, "<ANTIQUITY2_ID />", asEmployeesQueries(1)), "<ANTIQUITY3_ID />", asEmployeesQueries(2)) & asEmployeesQueries(0) & sQueryEnd, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
							End If
							'If lErrorNumber <> 0 Then Exit For
							If bTimeout Then Exit For
						Next
					End If
					Call DeleteFile(sFilePath & "_PayrollAntiquity3_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
				Next
			End If
		End If

		iCounter = 0
		If Not bTimeout Then
			sErrorDescription = "No se pudieron obtener las antigüedades de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, Antiquity2ID, Antiquity3ID From Employees Where Antiquity2ID<0", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				aiDays = 0
				Do While Not oRecordset.EOF
					For iIndex = 0 To UBound(alAntiquities)
						If ((Abs(CLng(oRecordset.Fields("Antiquity3ID").Value)) / 365) >= alAntiquities(iIndex)(1)) And ((Abs(CLng(oRecordset.Fields("Antiquity3ID").Value)) / 365) < alAntiquities(iIndex)(2)) Then
							aiDays = alAntiquities(iIndex)(0)
						End If
						If ((Abs(CLng(oRecordset.Fields("Antiquity2ID").Value)) / 365) >= alAntiquities(iIndex)(1)) And ((Abs(CLng(oRecordset.Fields("Antiquity2ID").Value)) / 365) < alAntiquities(iIndex)(2)) Then
							'sErrorDescription = "No se pudo actualizar la antigüedad del empleado."
							'lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update Employees Set Antiquity2ID=" & alAntiquities(iIndex)(0) & " Where (EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
							lErrorNumber = AppendTextToFile(sFilePath & "_PayrollAntiquity4_" & Int(iCounter / ROWS_PER_FILE) & ".txt", CStr(oRecordset.Fields("EmployeeID").Value) & "," & alAntiquities(iIndex)(0) & "," & aiDays, sErrorDescription)
							iCounter = iCounter + 1
							Exit For
						End If
					Next
					oRecordset.MoveNext
				Loop
				oRecordset.Close
			End If
		End If

		If Not bTimeout Then
			If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 2. RUN FROM FILES, Update Employees.Antiquity2ID " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
				sTemp = "Update Employees Set Antiquity2ID=<ANTIQUITY2_ID />, Antiquity3ID=<ANTIQUITY3_ID /> Where (EmployeeID="
				sQueryEnd = ")"
				For jIndex = 0 To iCounter Step ROWS_PER_FILE
					asFileContents = GetFileContents(sFilePath & "_PayrollAntiquity4_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
					If Len(asFileContents) > 0 Then
						asFileContents = Split(asFileContents, vbNewLine)
						For iIndex = 0 To UBound(asFileContents)
							If Len(asFileContents(iIndex)) > 0 Then
								asEmployeesQueries = Split(asFileContents(iIndex), ",")
								sErrorDescription = "No se pudo actualizar la antigüedad del empleado."
								lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, Replace(Replace(sTemp, "<ANTIQUITY2_ID />", asEmployeesQueries(1)), "<ANTIQUITY3_ID />", asEmployeesQueries(2)) & asEmployeesQueries(0) & sQueryEnd, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
							End If
							'If lErrorNumber <> 0 Then Exit For
							If bTimeout Then Exit For
						Next
					End If
					Call DeleteFile(sFilePath & "_PayrollAntiquity4_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
				Next
			End If
		End If

Call DisplayTimeStamp("START: LEVEL 2. Update Payroll_Antiquities. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		adTotal = Split("0,0,0,0", ",")
		For iIndex = 0 To UBound(adTotal)
			adTotal(iIndex) = Split("0,0,0", ",")
			adTotal(iIndex)(0) = 0
			adTotal(iIndex)(1) = 0
			adTotal(iIndex)(2) = 0
		Next
		lCurrentID = -2
		lCurrentID2 = -2
		iCounter = 0
		bTemp = 0
		bCurrent = 0
		If Not bTimeout Then
			sErrorDescription = "No se pudieron obtener las antigüedades de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_Antiquities", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		End If
		If Not bTimeout Then
			Call BuildCondition("", sQueryBegin)
			If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"

			sErrorDescription = "No se pudieron obtener las antigüedades de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeTypeID, EmployeesHistoryList.PositionTypeID, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate, EmployeesHistoryList.ReasonID From EmployeesHistoryList, EmployeesChangesLKP, StatusEmployees, Reasons " & sQueryBegin & " Where (EmployeesHistoryList.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=" & lPayID & ") And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) " & sCondition & " Order By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeTypeID, EmployeesHistoryList.EmployeeDate", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			'((EmployeeTypeID In (3,4,5)) Or (PositionTypeID In (1,2,5))) And 
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					bMinMaxApplied = True
					lStartDate = CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")))
					Do While Not oRecordset.EOF
						If (lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value)) Or (lCurrentID2 <> CLng(oRecordset.Fields("EmployeeTypeID").Value)) Then
							If lCurrentID <> -2 Then
								For iIndex = 1 To UBound(adTotal)
									adTotal(iIndex)(1) = adTotal(iIndex)(1) + Int(adTotal(iIndex)(2) / 30.4)
									adTotal(iIndex)(2) = Int(adTotal(iIndex)(2) - (Int(adTotal(iIndex)(2) / 30.4) * 30.4))
									adTotal(iIndex)(0) = adTotal(iIndex)(0) + Int(adTotal(iIndex)(1) / 12)
									adTotal(iIndex)(1) = adTotal(iIndex)(1) Mod 12
								Next
								lErrorNumber = AppendTextToFile(sFilePath & "_EmployeesAntiquities_" & Int(iCounter / ROWS_PER_FILE) & ".txt", lCurrentID & ", " & lCurrentID2 & ", " & adTotal(1)(0) & ", " & adTotal(1)(1) & ", " & adTotal(1)(2) & ", " & adTotal(2)(0) & ", " & adTotal(2)(1) & ", " & adTotal(2)(2) & ", " & adTotal(3)(0) & ", " & adTotal(3)(1) & ", " & adTotal(3)(2) & ", " & bTemp & ", " & bCurrent, sErrorDescription)
								iCounter = iCounter + 1
							End If
							If lCurrentID2 <> CLng(oRecordset.Fields("EmployeeTypeID").Value) Then
								adTotal(2)(0) = 0
								adTotal(2)(1) = 0
								adTotal(2)(2) = 0
								bTemp = 0
								bCurrent = 0
								bMinMaxApplied = True
								lTempEndDate = CLng(oRecordset.Fields("EndDate").Value)
								lCurrentID2 = CLng(oRecordset.Fields("EmployeeTypeID").Value)
							End If
							If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
								For iIndex = 1 To UBound(adTotal)
									adTotal(iIndex)(0) = 0
									adTotal(iIndex)(1) = 0
									adTotal(iIndex)(2) = 0
								Next
								bTemp = 0
								bCurrent = 0
								bMinMaxApplied = True
								lTempEndDate = CLng(oRecordset.Fields("EndDate").Value)
								lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
								lCurrentID2 = CLng(oRecordset.Fields("EmployeeTypeID").Value)
							End If
						End If
						If bMinMaxApplied And (CLng(oRecordset.Fields("EndDate").Value) < 30000000) Then
							If (lTempEndDate < AddDaysToSerialDate(CLng(oRecordset.Fields("EmployeeDate").Value), -1)) And (CLng(oRecordset.Fields("ReasonID").Value) = 17) Then
								adTotal(2)(0) = 0
								adTotal(2)(1) = 0
								adTotal(2)(2) = 0
'								bMinMaxApplied = False
							End If
						End If
						If bMinMaxApplied Then
							If CLng(oRecordset.Fields("EndDate").Value) = 30000000 Then
								Call GetAntiquityFromSerialDates(CLng(oRecordset.Fields("EmployeeDate").Value), aPayrollComponent(N_FOR_DATE_PAYROLL), adTotal(0)(0), adTotal(0)(1), adTotal(0)(2))
							Else
								Call GetAntiquityFromSerialDates(CLng(oRecordset.Fields("EmployeeDate").Value), CLng(oRecordset.Fields("EndDate").Value), adTotal(0)(0), adTotal(0)(1), adTotal(0)(2))
							End If
							adTotal(1)(0) = adTotal(1)(0) + adTotal(0)(0)
							adTotal(1)(1) = adTotal(1)(1) + adTotal(0)(1)
							adTotal(1)(2) = adTotal(1)(2) + adTotal(0)(2)
							If ((InStr(1, ",3,4,5,", CStr(oRecordset.Fields("EmployeeTypeID").Value), vbBinaryCompare) > 0) Or (InStr(1, ",1,2,5,", CStr(oRecordset.Fields("PositionTypeID").Value), vbBinaryCompare) > 0)) Then
								adTotal(2)(0) = adTotal(2)(0) + adTotal(0)(0)
								adTotal(2)(1) = adTotal(2)(1) + adTotal(0)(1)
								adTotal(2)(2) = adTotal(2)(2) + adTotal(0)(2)
							End If
							If (adTotal(0)(0) > 1) Or ((adTotal(0)(1) >= 6) And (adTotal(0)(2) > 0)) Then bTemp = 1
							If (CLng(oRecordset.Fields("EmployeeDate").Value) <= aPayrollComponent(N_FOR_DATE_PAYROLL)) And (CLng(oRecordset.Fields("EndDate").Value) >= aPayrollComponent(N_FOR_DATE_PAYROLL)) Then bCurrent = 1

							adTotal(0)(0) = 0
							adTotal(0)(1) = 0
							adTotal(0)(2) = 0

							If (CLng(oRecordset.Fields("EmployeeDate").Value) > CLng(lStartDate & "0000")) And ((CLng(oRecordset.Fields("EndDate").Value) < CLng(lStartDate & "9999")) Or (CLng(oRecordset.Fields("EndDate").Value) = 30000000)) Then
							ElseIf (CLng(oRecordset.Fields("EmployeeDate").Value) < CLng(lStartDate & "0000")) And (CLng(oRecordset.Fields("EndDate").Value) > CLng(lStartDate & "0000")) And (CLng(oRecordset.Fields("EndDate").Value) < CLng(lStartDate & "9999")) Then
								Call GetAntiquityFromSerialDates(CLng(lStartDate & "0101"), CLng(oRecordset.Fields("EndDate").Value), adTotal(0)(0), adTotal(0)(1), adTotal(0)(2))
							ElseIf (CLng(oRecordset.Fields("EmployeeDate").Value) > CLng(lStartDate & "0000")) And (CLng(oRecordset.Fields("EmployeeDate").Value) < CLng(lStartDate & "9999")) And (CLng(oRecordset.Fields("EndDate").Value) > CLng(lStartDate & "9999")) Then
								Call GetAntiquityFromSerialDates(CLng(oRecordset.Fields("EmployeeDate").Value), aPayrollComponent(N_FOR_DATE_PAYROLL), adTotal(0)(0), adTotal(0)(1), adTotal(0)(2))
							ElseIf (CLng(oRecordset.Fields("EmployeeDate").Value) < CLng(lStartDate & "0000")) And (CLng(oRecordset.Fields("EndDate").Value) > CLng(lStartDate & "9999")) Then
								Call GetAntiquityFromSerialDates(CLng(lStartDate & "0101"), aPayrollComponent(N_FOR_DATE_PAYROLL), adTotal(0)(0), adTotal(0)(1), adTotal(0)(2))
							End If
							adTotal(3)(0) = adTotal(3)(0) + adTotal(0)(0)
							adTotal(3)(1) = adTotal(3)(1) + adTotal(0)(1)
							adTotal(3)(2) = adTotal(3)(2) + adTotal(0)(2)
							lTempEndDate = CLng(oRecordset.Fields("EndDate").Value)
						End If
						oRecordset.MoveNext
						'If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
					For iIndex = 1 To UBound(adTotal)
						adTotal(iIndex)(1) = adTotal(iIndex)(1) + Int(adTotal(iIndex)(2) / 30.4)
						adTotal(iIndex)(2) = Int(adTotal(iIndex)(2) - (Int(adTotal(iIndex)(2) / 30.4) * 30.4))
						adTotal(iIndex)(0) = adTotal(iIndex)(0) + Int(adTotal(iIndex)(1) / 12)
						adTotal(iIndex)(1) = adTotal(iIndex)(1) Mod 12
					Next
					lErrorNumber = AppendTextToFile(sFilePath & "_EmployeesAntiquities_" & Int(iCounter / ROWS_PER_FILE) & ".txt", lCurrentID & ", " & lCurrentID2 & ", " & adTotal(1)(0) & ", " & adTotal(1)(1) & ", " & adTotal(1)(2) & ", " & adTotal(2)(0) & ", " & adTotal(2)(1) & ", " & adTotal(2)(2) & ", " & adTotal(3)(0) & ", " & adTotal(3)(1) & ", " & adTotal(3)(2) & ", 15, " & bCurrent, sErrorDescription)
					iCounter = iCounter + 1
				End If

				If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 2. RUN FROM FILES, Update Antiquities 3 y 4, " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
					sQueryBegin = "Insert Into Payroll_Antiquities (EmployeeID, EmployeeTypeID, Years1, Months1, Days1, Years2, Months2, Days2, Years3, Months3, Days3, bSixMonth, bIsCurrent) Values ("
					sQueryEnd = ")"
					For jIndex = 0 To iCounter Step ROWS_PER_FILE
						asFileContents = GetFileContents(sFilePath & "_EmployeesAntiquities_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
						If Len(asFileContents) > 0 Then
							asFileContents = Split(asFileContents, vbNewLine)
							For iIndex = 0 To UBound(asFileContents)
								If Len(asFileContents(iIndex)) > 0 Then
									sErrorDescription = "No se pudo agregar la antigüedad del empleado a la tabla temporal."
									lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, sQueryBegin & asFileContents(iIndex) & sQueryEnd, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
								End If
								'If lErrorNumber <> 0 Then Exit For
								If bTimeout Then Exit For
							Next
						End If
						Call DeleteFile(sFilePath & "_EmployeesAntiquities_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
					Next
				End If

Call DisplayTimeStamp("START: LEVEL 2. Update Payroll_Factors. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
				If lErrorNumber = 0 Then
					lStartDate = GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL))
					lEndDate = aPayrollComponent(N_FOR_DATE_PAYROLL)
					If Not bTimeout Then
						sErrorDescription = "No se pudieron obtener los días de inactividad de los empleados."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payroll_Factors.EmployeeID, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate From EmployeesAbsencesLKP, EmployeesChangesLKP, Payroll_Factors, Absences Where (EmployeesAbsencesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesAbsencesLKP.EmployeeID=Payroll_Factors.EmployeeID) And (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And (Absences.AbsenceTypeID2=0) And (((EmployeesAbsencesLKP.EndDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ") And (EmployeesAbsencesLKP.EndDate>=" & lStartDate & "))) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (EmployeesAbsencesLKP.Active=1) Order By Payroll_Factors.EmployeeID, EmployeesAbsencesLKP.OcurredDate", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								lCurrentID = -2
								Do While Not oRecordset.EOF
									If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
										If lCurrentID > -2 Then
											sErrorDescription = "No se pudo actualizar los días laborados del empleado."
											lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update Payroll_Factors Set PayrollFactor=PayrollFactor-" & dAmount & " Where (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)

											sErrorDescription = "No se pudo actualizar los días laborados del empleado."
											lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update EmployeesChangesLKP Set FirstDate=" & lTempStartDate & ", LastDate=" & lTempEndDate & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (-" & lPayID & "," & lPayID & ")) And (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
										End If
										lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
										dAmount = 0
										lTempStartDate = lStartDate
										lTempEndDate = lEndDate
									End If
									lTempStartDate = CLng(oRecordset.Fields("OcurredDate").Value)
									lTempEndDate = CLng(oRecordset.Fields("EndDate").Value)
									If lTempStartDate < lStartDate Then lTempStartDate = lStartDate
									If lTempEndDate > lEndDate Then lTempEndDate = lEndDate
									If lTempStartDate <= lTempEndDate Then
										dAmount = dAmount + (Abs(DateDiff("d", GetDateFromSerialNumber(lTempStartDate), GetDateFromSerialNumber(lTempEndDate))) + 1)
									End If
									oRecordset.MoveNext
									'If Err.number <> 0 Then Exit Do
									If bTimeout Then Exit Do
								Loop
								If Not bTimeout Then
									If dAmount > 0 Then
										sErrorDescription = "No se pudo actualizar los días laborados del empleado."
										lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update Payroll_Factors Set PayrollFactor=PayrollFactor-" & dAmount & " Where (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)

										sErrorDescription = "No se pudo actualizar los días laborados del empleado."
										lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update EmployeesChangesLKP Set FirstDate=" & lTempStartDate & ", LastDate=" & lTempEndDate & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (-" & lPayID & "," & lPayID & ")) And (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
									End If
								End If
							End If
						End If
					End If

					If Not bTimeout Then
						dTemp = (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDate), GetDateFromSerialNumber(lEndDate))) + 1)
						sErrorDescription = "No se pudieron obtener los días de inactividad de los empleados."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payroll_Factors.EmployeeID, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate From EmployeesHistoryList, EmployeesChangesLKP, StatusEmployees, Reasons, Payroll_Factors Where (EmployeesHistoryList.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeID=Payroll_Factors.EmployeeID) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID=1) And (((EmployeesHistoryList.EndDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EmployeeDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate<=" & lEndDate & ") And (EmployeesHistoryList.EndDate>=" & lStartDate & "))) Order By Payroll_Factors.EmployeeID, EmployeesHistoryList.EmployeeDate", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							If Not oRecordset.EOF Then
								lCurrentID = -2
								Do While Not oRecordset.EOF
									If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
										If lCurrentID > -2 Then
											If dAmount <> dTemp Then
												sErrorDescription = "No se pudo actualizar los días laborados del empleado."
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_Factors Set PayrollFactor=PayrollFactor-" & (dTemp - dAmount) & " Where (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

												sErrorDescription = "No se pudo actualizar los días laborados del empleado."
												lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update EmployeesChangesLKP Set FirstDate=" & lTempStartDate & ", LastDate=" & lTempEndDate & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (-" & lPayID & "," & lPayID & ")) And (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
											End If
										End If
										lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
										dAmount = 0
									End If
									lTempStartDate = CLng(oRecordset.Fields("EmployeeDate").Value)
									lTempEndDate = CLng(oRecordset.Fields("EndDate").Value)
									If lTempStartDate < lStartDate Then lTempStartDate = lStartDate
									If lTempEndDate > lEndDate Then lTempEndDate = lEndDate
									If lTempStartDate <= lTempEndDate Then
										dAmount = dAmount + (Abs(DateDiff("d", GetDateFromSerialNumber(lTempStartDate), GetDateFromSerialNumber(lTempEndDate))) + 1)
									End If
									oRecordset.MoveNext
									'If Err.number <> 0 Then Exit Do
									If bTimeout Then Exit Do
								Loop
								If Not bTimeout Then
									If dAmount <> dTemp Then
										sErrorDescription = "No se pudo actualizar los días laborados del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_Factors Set PayrollFactor=PayrollFactor-" & (dTemp - dAmount) & " Where (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

										sErrorDescription = "No se pudo actualizar los días laborados del empleado."
										lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update EmployeesChangesLKP Set FirstDate=" & lTempStartDate & ", LastDate=" & lTempEndDate & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (-" & lPayID & "," & lPayID & ")) And (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
									End If
								End If
							End If
						End If
					End If

					If Not bTimeout Then
						sErrorDescription = "No se pudo actualizar los días laborados del empleado."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_Factors Set PayrollFactor=0 Where (PayrollFactor<0)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					End If

					If Not bTimeout Then
						sErrorDescription = "No se pudo actualizar los días laborados del empleado."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_Factors Set PayrollFactor=PayrollFactor/15", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					End If

					If Not bTimeout Then
						sErrorDescription = "No se pudo actualizar los días laborados del empleado."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_Factors Set PayrollFactor=1 Where (PayrollFactor>1)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					End If
				End If
			End If
		End If
	End If

Call DisplayTimeStamp("START: LEVEL 2. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
	If Not bTimeout Then
		adDSM = "0"
		sErrorDescription = "No se pudieron obtener los días de salario mínimo y los días de salario burocrático."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CurrencyValue From CurrenciesHistoryList Where (CurrencyDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (CurrencyID In (1,2,3,4,5)) Order By CurrencyID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				adDSM = adDSM & ";" & CDbl(oRecordset.Fields("CurrencyValue").Value)
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			adDSM = Split(adDSM, ";")
			For iIndex = 1 To UBound(adDSM)
				adDSM(iIndex) = CDbl(adDSM(iIndex))
			Next
		End If
	End If

	If Not bAdjustment Then
Call DisplayTimeStamp("START: LEVEL 2. Delete Payroll. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		If Not bTimeout Then
			Call BuildCondition(sCondition, sQueryBegin)
			If bRetroactive And (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) Then sCondition = " And (EmployeesHistoryList.EmployeeID=EmployeesRevisions.EmployeeID) And (EmployeesRevisions.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesRevisions.StartPayrollID=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")" & sCondition
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudieron eliminar los conceptos de pagos de la nómina."
				If (Len(sCondition) = 0) And (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				Else
					If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & "))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudieron eliminar los conceptos de pagos de la nómina."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (RecordID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & "))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
			End If
			sConceptCondition = ""
			If (aPayrollComponent(N_TYPE_ID_PAYROLL) = 3) And (Len(oRequest("PayrollConceptID").Item) > 0) Then
				sConceptCondition = " And (Concepts.ConceptID In (" & oRequest("PayrollConceptID").Item & "))"
			End If
		End If

Call DisplayTimeStamp("START: LEVEL 2, CREATE FILES EmployeesConceptsLKP, QttyID=1. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		If Not bTimeout Then
			iCounter = 0
			sQueryBegin = ""
			If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = ", EmployeesRevisions"
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & aPayrollComponent(N_FOR_DATE_PAYROLL) & "' As RecordDate, '1' As RecordID, EmployeesConceptsLKP.EmployeeID, Concepts.ConceptID, '1' As PayrollTypeID, ConceptAmount, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From EmployeesConceptsLKP, Concepts, EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Jobs, Areas, Zones " & sQueryBegin & " Where (EmployeesConceptsLKP.ConceptID=Concepts.ConceptID) And (EmployeesConceptsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.JobID=Jobs.JobID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EmployeesConceptsLKP.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesConceptsLKP.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") " & Replace(sCondition, "(Positions.", "(EmployeesHistoryList.") & sConceptCondition & " And (ConceptQttyID=1) And (Concepts.PeriodID In (" & sPeriods & ")) And (EmployeesConceptsLKP.Active=1)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If

Call DisplayTimeStamp("START: LEVEL 2, CREATE FILES, QttyID=1. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		If Not bTimeout Then
			lErrorNumber = CreateConceptsFile(oRequest, oADODBConnection, 1, aPayrollComponent(N_FOR_DATE_PAYROLL), sErrorDescription)
			If lErrorNumber = 0 Then
				iCounter = 0
				asFileContents = GetFileContents(PAYROLL_FILE1_PATH, sErrorDescription)
				If Len(asFileContents) > 0 Then
					asFileContents = Split(asFileContents, vbNewLine)
					For iIndex = 0 To UBound(asFileContents)
						If Len(asFileContents(iIndex)) > 0 Then
							asEmployeesQueries = Split(asFileContents(iIndex), LIST_SEPARATOR)
							If InStr(1, "," & sPeriods & ",", "," & asEmployeesQueries(1) & ",", vbbinaryCompare) > 0 Then
								sQueryBegin = ""
								asEmployeesQueries(2) = asEmployeesQueries(2) & sCondition
								If InStr(1, asEmployeesQueries(2), "(Jobs.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Jobs"
								If InStr(1, asEmployeesQueries(2), "(Zones.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Zones"
								If InStr(1, asEmployeesQueries(2), "(Areas.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Areas"
								If (InStr(1, asEmployeesQueries(2), "=Employees.", vbBinaryCompare) > 0) Or (InStr(1, asEmployeesQueries(2), "(Employees.", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", Employees"
								If InStr(1, asEmployeesQueries(2), "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
								If InStr(1, asEmployeesQueries(2), "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
								If InStr(1, asEmployeesQueries(2), "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
								If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, asEmployeesQueries(2), "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
								sErrorDescription = "No se pudieron obtener los empleados para registrar sus conceptos de pago en la nómina."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) " & asEmployeesQueries(2) & " Order By EmployeesHistoryList.EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									Do While Not oRecordset.EOF
										lErrorNumber = AppendTextToFile(sFilePath & "_Payroll1_" & Int(iCounter / ROWS_PER_FILE) & ".txt", asEmployeesQueries(0) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EmployeeID").Value), sErrorDescription)
										iCounter = iCounter + 1
										oRecordset.MoveNext
										'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
										If bTimeout Then Exit Do
									Loop
									oRecordset.Close
								End If
							End If
						End If
						'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
					Next
				End If

				If Not bTimeout Then
					If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 2, RUN FROM FILES, QttyID=1, " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
						sQueryBegin = "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ", 1, "
						sQueryEnd = ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")"
						For jIndex = 0 To iCounter Step ROWS_PER_FILE
							asFileContents = GetFileContents(sFilePath & "_Payroll1_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
							If Len(asFileContents) > 0 Then
								asFileContents = Split(asFileContents, vbNewLine)
								For iIndex = 0 To UBound(asFileContents)
									If Len(asFileContents(iIndex)) > 0 Then
										asEmployeesQueries = Split(asFileContents(iIndex), SECOND_LIST_SEPARATOR)
										sErrorDescription = "No se pudo agregar el concepto de pago y su monto a la nómina del empleado."
										lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, sQueryBegin & asEmployeesQueries(2) & ", " & asEmployeesQueries(0) & ", 1, " & asEmployeesQueries(1) & sQueryEnd, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
									End If
									'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
									If bTimeout Then Exit For
								Next
							End If
							Call DeleteFile(sFilePath & "_Payroll1_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
						Next
					End If
				End If
			End If
		End If

Call DisplayTimeStamp("START: LEVEL 2, CREATE FILES EmployeesConceptsLKP, QttyID=3. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		If Not bTimeout Then
			iCounter = 0
			sQueryBegin = ""
			If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = ", EmployeesRevisions"
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesConceptsLKP.EmployeeID, Concepts.ConceptID, ConceptAmount, CurrencyValue, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, Areas.EconomicZoneID, Zones.ZoneTypeID From EmployeesConceptsLKP, Concepts, EmployeesHistoryList, EmployeesChangesLKP, StatusEmployees, Reasons, Jobs, Areas, Zones, CurrenciesHistoryList " & sQueryBegin & " Where (EmployeesConceptsLKP.ConceptID=Concepts.ConceptID) And (EmployeesConceptsLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.JobID=Jobs.JobID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=CurrenciesHistoryList.CurrencyID) And (CurrenciesHistoryList.CurrencyDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesConceptsLKP.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesConceptsLKP.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesConceptsLKP.Active=1) " & Replace(sCondition, "(Positions.", "(EmployeesHistoryList.") & sConceptCondition & " And (ConceptQttyID=3) And (Concepts.PeriodID In (" & sPeriods & ")) Order By EmployeesConceptsLKP.EmployeeID, Concepts.OrderInList, Concepts.ConceptID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				Do While Not oRecordset.EOF
					dAmount = (CDbl(oRecordset.Fields("CurrencyValue").Value) * CDbl(oRecordset.Fields("ConceptAmount").Value))
					If CDbl(oRecordset.Fields("ConceptMin").Value) > 0 Then
						If CInt(oRecordset.Fields("ConceptMinQttyID").Value) = 3 Then
							If dAmount < (CDbl(oRecordset.Fields("ConceptMin").Value) * adDSM(CInt(oRecordset.Fields("ZoneTypeID").Value))) Then dAmount = CDbl(oRecordset.Fields("ConceptMin").Value) * adDSM(CInt(oRecordset.Fields("ZoneTypeID").Value))
						ElseIf CInt(oRecordset.Fields("ConceptMinQttyID").Value) = 13 Then
							If dAmount < (CDbl(oRecordset.Fields("ConceptMin").Value) * adDSM(CInt(oRecordset.Fields("EconomicZoneID").Value) + 3)) Then dAmount = CDbl(oRecordset.Fields("ConceptMin").Value) * adDSM(CInt(oRecordset.Fields("EconomicZoneID").Value) + 3)
						ElseIf CInt(oRecordset.Fields("ConceptMinQttyID").Value) = 23 Then
							If dAmount < (CDbl(oRecordset.Fields("ConceptMin").Value) * adDSM(CInt(oRecordset.Fields("ZoneTypeID").Value))) Then dAmount = CDbl(oRecordset.Fields("ConceptMin").Value) * adDSM(CInt(oRecordset.Fields("ZoneTypeID").Value))
						ElseIf CInt(oRecordset.Fields("ConceptMinQttyID").Value) = 24 Then
							If dAmount < (CDbl(oRecordset.Fields("ConceptMin").Value) * adDSM(1)) Then dAmount = CDbl(oRecordset.Fields("ConceptMin").Value) * adDSM(1)
						ElseIf CInt(oRecordset.Fields("ConceptMinQttyID").Value) = 25 Then
							If dAmount < (CDbl(oRecordset.Fields("ConceptMin").Value) * adDSM(1)) Then dAmount = CDbl(oRecordset.Fields("ConceptMin").Value) * adDSM(1)
						ElseIf CInt(oRecordset.Fields("ConceptMinQttyID").Value) = 33 Then
							If dAmount < (CDbl(oRecordset.Fields("ConceptMin").Value) * adDSM(CInt(oRecordset.Fields("EconomicZoneID").Value) + 3)) Then dAmount = CDbl(oRecordset.Fields("ConceptMin").Value) * adDSM(CInt(oRecordset.Fields("EconomicZoneID").Value) + 3)
						ElseIf CInt(oRecordset.Fields("ConceptMinQttyID").Value) = 34 Then
							If dAmount < (CDbl(oRecordset.Fields("ConceptMin").Value) * adDSM(4)) Then dAmount = CDbl(oRecordset.Fields("ConceptMin").Value) * adDSM(4)
						ElseIf CInt(oRecordset.Fields("ConceptMinQttyID").Value) = 35 Then
							If dAmount < (CDbl(oRecordset.Fields("ConceptMin").Value) * adDSM(4)) Then dAmount = CDbl(oRecordset.Fields("ConceptMin").Value) * adDSM(4)
						Else
							If dAmount < CDbl(oRecordset.Fields("ConceptMin").Value) Then dAmount = CDbl(oRecordset.Fields("ConceptMin").Value)
						End If
					End If
					If CDbl(oRecordset.Fields("ConceptMax").Value) > 0 Then
						If CInt(oRecordset.Fields("ConceptMaxQttyID").Value) = 3 Then
							If dAmount > (CDbl(oRecordset.Fields("ConceptMax").Value) * adDSM(CInt(oRecordset.Fields("ZoneTypeID").Value))) Then dAmount = CDbl(oRecordset.Fields("ConceptMax").Value) * adDSM(CInt(oRecordset.Fields("ZoneTypeID").Value))
						ElseIf CInt(oRecordset.Fields("ConceptMaxQttyID").Value) = 13 Then
							If dAmount > (CDbl(oRecordset.Fields("ConceptMax").Value) * adDSM(CInt(oRecordset.Fields("EconomicZoneID").Value) + 3)) Then dAmount = CDbl(oRecordset.Fields("ConceptMax").Value) * adDSM(CInt(oRecordset.Fields("EconomicZoneID").Value) + 3)
						ElseIf CInt(oRecordset.Fields("ConceptMaxQttyID").Value) = 23 Then
							If dAmount > (CDbl(oRecordset.Fields("ConceptMax").Value) * adDSM(CInt(oRecordset.Fields("ZoneTypeID").Value))) Then dAmount = CDbl(oRecordset.Fields("ConceptMax").Value) * adDSM(CInt(oRecordset.Fields("ZoneTypeID").Value))
						ElseIf CInt(oRecordset.Fields("ConceptMaxQttyID").Value) = 24 Then
							If dAmount > (CDbl(oRecordset.Fields("ConceptMax").Value) * adDSM(1)) Then dAmount = CDbl(oRecordset.Fields("ConceptMax").Value) * adDSM(1)
						ElseIf CInt(oRecordset.Fields("ConceptMaxQttyID").Value) = 25 Then
							If dAmount > (CDbl(oRecordset.Fields("ConceptMax").Value) * adDSM(1)) Then dAmount = CDbl(oRecordset.Fields("ConceptMax").Value) * adDSM(1)
						ElseIf CInt(oRecordset.Fields("ConceptMaxQttyID").Value) = 33 Then
							If dAmount > (CDbl(oRecordset.Fields("ConceptMax").Value) * adDSM(CInt(oRecordset.Fields("EconomicZoneID").Value) + 3)) Then dAmount = CDbl(oRecordset.Fields("ConceptMax").Value) * adDSM(CInt(oRecordset.Fields("EconomicZoneID").Value) + 3)
						ElseIf CInt(oRecordset.Fields("ConceptMaxQttyID").Value) = 34 Then
							If dAmount > (CDbl(oRecordset.Fields("ConceptMax").Value) * adDSM(4)) Then dAmount = CDbl(oRecordset.Fields("ConceptMax").Value) * adDSM(4)
						ElseIf CInt(oRecordset.Fields("ConceptMaxQttyID").Value) = 35 Then
							If dAmount > (CDbl(oRecordset.Fields("ConceptMax").Value) * adDSM(4)) Then dAmount = CDbl(oRecordset.Fields("ConceptMax").Value) * adDSM(4)
						Else
							If dAmount > CDbl(oRecordset.Fields("ConceptMax").Value) Then dAmount = CDbl(oRecordset.Fields("ConceptMax").Value)
						End If
					End If
					lErrorNumber = AppendTextToFile(sFilePath & "_Payroll3E_" & Int(iCounter / ROWS_PER_FILE) & ".txt", CStr(oRecordset.Fields("EmployeeID").Value) & ", " & CStr(oRecordset.Fields("ConceptID").Value) & ", 1, " & FormatNumber(dAmount, 2, True, False, False), sErrorDescription)
					iCounter = iCounter + 1
					oRecordset.MoveNext
					'If lErrorNumber <> 0 Then Exit Do
				Loop
				oRecordset.Close
			End If
		End If

		If Not bTimeout Then
			sQueryBegin = ""
			If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = ", EmployeesRevisions"
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesConceptsLKP.EmployeeID, Concepts.ConceptID, ConceptAmount, CurrencyValue, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, Areas.EconomicZoneID, Zones.ZoneTypeID From EmployeesConceptsLKP, Concepts, EmployeesHistoryList, EmployeesChangesLKP, StatusEmployees, Reasons, Jobs, Areas, Zones, ZoneTypes, CurrenciesHistoryList " & sQueryBegin & " Where (EmployeesConceptsLKP.ConceptID=Concepts.ConceptID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesConceptsLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.JobID=Jobs.JobID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Areas.EconomicZoneID=ZoneTypes.ZoneTypeID) And (ZoneTypes.ZoneTypeID2=CurrenciesHistoryList.CurrencyID) And (CurrenciesHistoryList.CurrencyDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesConceptsLKP.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesConceptsLKP.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesConceptsLKP.Active=1) " & Replace(sCondition, "(Positions.", "(EmployeesHistoryList.") & sConceptCondition & " And (Concepts.PeriodID In (" & sPeriods & ")) And (ConceptQttyID=13) Order By EmployeesConceptsLKP.EmployeeID, Concepts.OrderInList, Concepts.ConceptID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				Do While Not oRecordset.EOF
					lErrorNumber = AppendTextToFile(sFilePath & "_Payroll3E_" & Int(iCounter / ROWS_PER_FILE) & ".txt", CStr(oRecordset.Fields("EmployeeID").Value) & ", " & CStr(oRecordset.Fields("ConceptID").Value) & ", 1, " & FormatNumber((CDbl(oRecordset.Fields("CurrencyValue").Value) * CDbl(oRecordset.Fields("ConceptAmount").Value)), 2, True, False, False), sErrorDescription)
					iCounter = iCounter + 1
					oRecordset.MoveNext
					'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
			End If
		End If

		If Not bTimeout Then
			If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 2, RUN FROM FILES, QttyID=3 EmployeesConceptsLKP, " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
				sQueryBegin = "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ", 1, "
				sQueryEnd = ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")"
				For jIndex = 0 To iCounter Step ROWS_PER_FILE
					asFileContents = GetFileContents(sFilePath & "_Payroll3E_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
					If Len(asFileContents) > 0 Then
						asFileContents = Split(asFileContents, vbNewLine)
						For iIndex = 0 To UBound(asFileContents)
							If Len(asFileContents(iIndex)) > 0 Then
								sErrorDescription = "No se pudo agregar el concepto de pago y su monto a la nómina del empleado."
								lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, sQueryBegin & asFileContents(iIndex) & sQueryEnd, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
							End If
							'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
							If bTimeout Then Exit For
						Next
					End If
					Call DeleteFile(sFilePath & "_Payroll3E_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
				Next
			End If
		End If

Call DisplayTimeStamp("START: LEVEL 2, CREATE FILES, QttyID=3. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		If Not bTimeout Then
			lErrorNumber = CreateConceptsFile(oRequest, oADODBConnection, 3, aPayrollComponent(N_FOR_DATE_PAYROLL), sErrorDescription)
			If lErrorNumber = 0 Then
				iCounter = 0
				asFileContents = GetFileContents(PAYROLL_FILE3_PATH, sErrorDescription)
				If Len(asFileContents) > 0 Then
					asFileContents = Split(asFileContents, vbNewLine)
					For iIndex = 0 To UBound(asFileContents)
						If Len(asFileContents(iIndex)) > 0 Then
							asEmployeesQueries = Split(asFileContents(iIndex), LIST_SEPARATOR)
							If InStr(1, "," & sPeriods & ",", "," & asEmployeesQueries(2) & ",", vbbinaryCompare) > 0 Then
								sQueryBegin = ""
								asEmployeesQueries(8) = asEmployeesQueries(8) & sCondition
								If (InStr(1, asEmployeesQueries(8), "=Employees.", vbBinaryCompare) > 0) Or (InStr(1, asEmployeesQueries(8), "(Employees.", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", Employees"
								If InStr(1, asEmployeesQueries(8), "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
								If InStr(1, asEmployeesQueries(8), "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
								If InStr(1, asEmployeesQueries(8), "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
								If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, asEmployeesQueries(8), "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
								sErrorDescription = "No se pudieron obtener los empleados para registrar sus conceptos de pago en la nómina."
								If CInt(asEmployeesQueries(3)) = 3 Then
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, Zones.ZoneTypeID, CurrencyValue From EmployeesHistoryList, EmployeesChangesLKP, StatusEmployees, Reasons, Jobs, Areas, Zones, CurrenciesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.JobID=Jobs.JobID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Zones.ZoneTypeID=CurrenciesHistoryList.CurrencyID) And (CurrenciesHistoryList.CurrencyDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") " & asEmployeesQueries(8) & " Order By EmployeesHistoryList.EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								Else
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, ZoneTypes.ZoneTypeID2, CurrencyValue From EmployeesHistoryList, EmployeesChangesLKP, StatusEmployees, Reasons, Jobs, Areas, Zones, ZoneTypes, CurrenciesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.JobID=Jobs.JobID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (Areas.EconomicZoneID=ZoneTypes.ZoneTypeID) And (ZoneTypes.ZoneTypeID2=CurrenciesHistoryList.CurrencyID) And (CurrenciesHistoryList.CurrencyDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") " & asEmployeesQueries(8) & " Order By EmployeesHistoryList.EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								End If
								If lErrorNumber = 0 Then
									Do While Not oRecordset.EOF
										lErrorNumber = AppendTextToFile(sFilePath & "_Payroll3_" & Int(iCounter / ROWS_PER_FILE) & ".txt", asEmployeesQueries(0) & SECOND_LIST_SEPARATOR & FormatNumber((CDbl(oRecordset.Fields("CurrencyValue").Value) * CDbl(asEmployeesQueries(1))), 2, True, False, False) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EmployeeID").Value), sErrorDescription)
										iCounter = iCounter + 1
										oRecordset.MoveNext
										'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
									Loop
									oRecordset.Close
								End If
							End If
						End If
						'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
						If bTimeout Then Exit For
					Next
				End If

				If Not bTimeout Then
					If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 2, RUN FROM FILES, QttyID=3, " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
						sQueryBegin = "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ", 1, "
						sQueryEnd = ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")"
						For jIndex = 0 To iCounter Step ROWS_PER_FILE
							asFileContents = GetFileContents(sFilePath & "_Payroll3_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
							If Len(asFileContents) > 0 Then
								asFileContents = Split(asFileContents, vbNewLine)
								For iIndex = 0 To UBound(asFileContents)
									If Len(asFileContents(iIndex)) > 0 Then
										asEmployeesQueries = Split(asFileContents(iIndex), SECOND_LIST_SEPARATOR)
										sErrorDescription = "No se pudo agregar el concepto de pago y su monto a la nómina del empleado."
										lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, sQueryBegin & asEmployeesQueries(2) & ", " & asEmployeesQueries(0) & ", 1, " & asEmployeesQueries(1) & sQueryEnd, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
									End If
									'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
									If bTimeout Then Exit For
								Next
							End If
							Call DeleteFile(sFilePath & "_Payroll3_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
						Next
					End If
				End If
			End If
		End If

		If Not bTimeout Then
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=4) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID In (1))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=5) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=7) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PositionTypeID <> 1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=8) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PositionTypeID <> 2)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=16) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=24) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=44) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=45) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=93) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PositionTypeID <> 1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=120) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID <> 1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If

		If Not bTimeout Then
			sErrorDescription = "No se pudo limpiar la tabla temporal."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudieron obtener los montos de las nóminas de los empleados."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, RecordID, Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID, ConceptID, PayrollTypeID, ConceptAmount*PayrollFactor As ConceptAmount1, ConceptTaxes, ConceptRetention, UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", Payroll_Factors Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=Payroll_Factors.EmployeeID) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (" & CONCEPTS_FOR_FACTOR & "))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo limpiar la tabla temporal."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID In (Select EmployeeID From Payroll_Factors)) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (" & CONCEPTS_FOR_FACTOR & "))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudieron obtener los montos de las nóminas de los empleados."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo limpiar la tabla temporal."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
		End If

		If Not bTimeout Then
			If bTruncate Then
Call DisplayTimeStamp("START: LEVEL 2, TRUNCATE DECIMALS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
				If False Then
					sErrorDescription = "No se pudieron truncar los decimales de los montos."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=Round(ConceptAmount, 2) Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Else
					sErrorDescription = "No se pudo limpiar la tabla temporal."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PayrollInt", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudieron truncar los decimales de los montos."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=(ConceptAmount+0.005)*100 Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudieron truncar los decimales de los montos."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PayrollInt (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo limpiar la tabla de montos."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudieron truncar los decimales de los montos."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID From PayrollInt", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudieron truncar los decimales de los montos."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=(ConceptAmount/100) Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If

Call DisplayTimeStamp("START: LEVEL 2, CREATE FILES EmployeesConceptsLKP, QttyID=2. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		If Not bTimeout Then
			iCounter = 0
			sQueryBegin = ""
			If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = ", EmployeesRevisions"
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesConceptsLKP.EmployeeID, Concepts.ConceptID, ConceptAmount, AppliesToID, ConceptMin, ConceptMinQttyID, ConceptMax, ConceptMaxQttyID, Areas.EconomicZoneID, Zones.ZoneTypeID From EmployeesConceptsLKP, Concepts, EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Jobs, Areas, Zones " & sQueryBegin & " Where (EmployeesConceptsLKP.ConceptID=Concepts.ConceptID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesConceptsLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesHistoryList.JobID=Jobs.JobID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EmployeesConceptsLKP.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesConceptsLKP.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") " & Replace(sCondition, "(Positions.", "(EmployeesHistoryList.") & sConceptCondition & " And (ConceptQttyID=2) And (AppliesToID Is Not Null) And (Concepts.ConceptID Not In (70)) And (Concepts.PeriodID In (" & sPeriods & ")) And (EmployeesConceptsLKP.Active=1) Order By EmployeesConceptsLKP.EmployeeID, Concepts.OrderInList, Concepts.ConceptID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				Do While Not oRecordset.EOF
					lErrorNumber = AppendTextToFile(sFilePath & "_Payroll2Ea_" & Int(iCounter / ROWS_PER_FILE) & ".txt", CStr(oRecordset.Fields("EmployeeID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptAmount").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("AppliesToID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMin").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMinQttyID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMax").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMaxQttyID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("EconomicZoneID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ZoneTypeID").Value), sErrorDescription)
					iCounter = iCounter + 1
					oRecordset.MoveNext
					'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
			End If
		End If

		If Not bTimeout Then
			If (lErrorNumber = 0) And (iCounter > 0) Then
				sQueryBegin = "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ", 1, "
				sQueryEnd = ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")"
Call DisplayTimeStamp("START: LEVEL 2, GET CONCEPTS AMOUNTS EmployeesConceptsLKP, QttyID=2, " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
				For jIndex = 0 To iCounter Step ROWS_PER_FILE
					asFileContents = GetFileContents(sFilePath & "_Payroll2Ea_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
					If Len(asFileContents) > 0 Then
						asFileContents = Split(asFileContents, vbNewLine)
						For iIndex = 0 To UBound(asFileContents)
							If Len(asFileContents(iIndex)) > 0 Then
								asEmployeesQueries = Split(asFileContents(iIndex), LIST_SEPARATOR)
								sErrorDescription = "No se pudieron obtener los empleados para registrar sus conceptos de pago en la nómina."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(ConceptAmount) As TotalAmount, IsDeduction From " & sTable & ", Concepts Where (" & sTable & ".ConceptID=Concepts.ConceptID) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID=" & asEmployeesQueries(0) & ") And (Concepts.ConceptID In (" & asEmployeesQueries(3) & ")) And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") Group By IsDeduction", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									dAmount = 0
									dTaxAmount = CDbl(asEmployeesQueries(2)) / 100
									bMinMaxApplied = False
									If Not oRecordset.EOF Then
										Do While Not oRecordset.EOF
											If CInt(oRecordset.Fields("IsDeduction").Value) = 0 Then
												dAmount = dAmount + CDbl(oRecordset.Fields("TotalAmount").Value)
											Else
												dAmount = dAmount - CDbl(oRecordset.Fields("TotalAmount").Value)
											End If
											oRecordset.MoveNext
										Loop
										oRecordset.Close

										If CDbl(asEmployeesQueries(4)) > 0 Then
											If CInt(asEmployeesQueries(5)) = 3 Then
												If (dAmount * dTaxAmount) < (CDbl(asEmployeesQueries(4)) * adDSM(CInt(asEmployeesQueries(9)))) Then
													dAmount = CDbl(asEmployeesQueries(4)) * adDSM(CInt(asEmployeesQueries(9)))
													bMinMaxApplied = True
												End If
											ElseIf CInt(asEmployeesQueries(5)) = 13 Then
												If (dAmount * dTaxAmount) < (CDbl(asEmployeesQueries(4)) * adDSM(CInt(asEmployeesQueries(8)) + 3)) Then
													dAmount = CDbl(asEmployeesQueries(4)) * adDSM(CInt(asEmployeesQueries(8)) + 3)
													bMinMaxApplied = True
												End If
											ElseIf CInt(asEmployeesQueries(5)) = 23 Then
												If dAmount < (CDbl(asEmployeesQueries(4)) * adDSM(CInt(asEmployeesQueries(9)))) Then
													dAmount = CDbl(asEmployeesQueries(4)) * adDSM(CInt(asEmployeesQueries(9)))
												End If
											ElseIf CInt(asEmployeesQueries(5)) = 24 Then
												If (dAmount * dTaxAmount) < (CDbl(asEmployeesQueries(4)) * adDSM(1)) Then
													dAmount = CDbl(asEmployeesQueries(4)) * adDSM(1)
													bMinMaxApplied = True
												End If
											ElseIf CInt(asEmployeesQueries(5)) = 25 Then
												If dAmount < (CDbl(asEmployeesQueries(4)) * adDSM(1)) Then
													dAmount = CDbl(asEmployeesQueries(4)) * adDSM(1)
												End If
											ElseIf CInt(asEmployeesQueries(5)) = 33 Then
												If dAmount < (CDbl(asEmployeesQueries(4)) * adDSM(CInt(asEmployeesQueries(8)) + 3)) Then
													dAmount = CDbl(asEmployeesQueries(4)) * adDSM(CInt(asEmployeesQueries(8)) + 3)
												End If
											ElseIf CInt(asEmployeesQueries(5)) = 34 Then
												If (dAmount * dTaxAmount) < (CDbl(asEmployeesQueries(4)) * adDSM(4)) Then
													dAmount = CDbl(asEmployeesQueries(4)) * adDSM(4)
													bMinMaxApplied = True
												End If
											ElseIf CInt(asEmployeesQueries(5)) = 35 Then
												If dAmount < (CDbl(asEmployeesQueries(4)) * adDSM(4)) Then
													dAmount = CDbl(asEmployeesQueries(4)) * adDSM(4)
												End If
											Else
												dAmount = dAmount * dTaxAmount
												bMinMaxApplied = True
												If dAmount < CDbl(asEmployeesQueries(4)) Then dAmount = CDbl(asEmployeesQueries(4))
											End If
										End If
										If CDbl(asEmployeesQueries(6)) > 0 Then
											If CInt(asEmployeesQueries(7)) = 3 Then
												If (dAmount * dTaxAmount) > (CDbl(asEmployeesQueries(6)) * adDSM(CInt(asEmployeesQueries(9)))) Then
													dAmount = CDbl(asEmployeesQueries(6)) * adDSM(CInt(asEmployeesQueries(9)))
													bMinMaxApplied = True
												End If
											ElseIf CInt(asEmployeesQueries(7)) = 13 Then
												If (dAmount * dTaxAmount) > (CDbl(asEmployeesQueries(6)) * adDSM(CInt(asEmployeesQueries(8)) + 3)) Then
													dAmount = CDbl(asEmployeesQueries(6)) * adDSM(CInt(asEmployeesQueries(8)) + 3)
													bMinMaxApplied = True
												End If
											ElseIf CInt(asEmployeesQueries(7)) = 23 Then
												If dAmount > (CDbl(asEmployeesQueries(6)) * adDSM(CInt(asEmployeesQueries(9)))) Then
													dAmount = CDbl(asEmployeesQueries(6)) * adDSM(CInt(asEmployeesQueries(9)))
												End If
											ElseIf CInt(asEmployeesQueries(7)) = 24 Then
												If (dAmount * dTaxAmount) > (CDbl(asEmployeesQueries(6)) * adDSM(1)) Then
													dAmount = CDbl(asEmployeesQueries(6)) * adDSM(1)
													bMinMaxApplied = True
												End If
											ElseIf CInt(asEmployeesQueries(7)) = 25 Then
												If dAmount > (CDbl(asEmployeesQueries(6)) * adDSM(1)) Then
													dAmount = CDbl(asEmployeesQueries(6)) * adDSM(1)
												End If
											ElseIf CInt(asEmployeesQueries(7)) = 33 Then
												If dAmount > (CDbl(asEmployeesQueries(6)) * adDSM(CInt(asEmployeesQueries(8)) + 3)) Then
													dAmount = CDbl(asEmployeesQueries(6)) * adDSM(CInt(asEmployeesQueries(8)) + 3)
												End If
											ElseIf CInt(asEmployeesQueries(7)) = 34 Then
												If (dAmount * dTaxAmount) > (CDbl(asEmployeesQueries(6)) * adDSM(4)) Then
													dAmount = CDbl(asEmployeesQueries(6)) * adDSM(4)
													bMinMaxApplied = True
												End If
											ElseIf CInt(asEmployeesQueries(7)) = 35 Then
												If dAmount > (CDbl(asEmployeesQueries(6)) * adDSM(4)) Then
													dAmount = CDbl(asEmployeesQueries(6)) * adDSM(4)
												End If
											Else
												dAmount = dAmount * dTaxAmount
												bMinMaxApplied = True
												If dAmount > CDbl(asEmployeesQueries(6)) Then dAmount = CDbl(asEmployeesQueries(6))
											End If
										End If
									End If
									If Not bMinMaxApplied Then dAmount = dAmount * dTaxAmount
									sErrorDescription = "No se pudo agregar el concepto de pago y su monto a la nómina del empleado."
									lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, sQueryBegin & asEmployeesQueries(0) & ", " & asEmployeesQueries(1) & ", 1, " & FormatNumber(dAmount, 2, True, False, False) & sQueryEnd, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
								End If
							End If
							'If lErrorNumber <> 0 Then Exit For
							If bTimeout Then Exit For
						Next
					End If
					Call DeleteFile(sFilePath & "_Payroll2Ea_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
				Next
			End If
		End If

		If Not bTimeout Then
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=4) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID In (1))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=5) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=7) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PositionTypeID <> 1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=8) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PositionTypeID <> 2)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=16) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=24) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=44) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=45) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=93) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PositionTypeID <> 1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=120) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID <> 1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If

Call DisplayTimeStamp("START: LEVEL 2, CREATE FILES, QttyID=2. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		If Not bTimeout Then
			lErrorNumber = CreateConceptsFile(oRequest, oADODBConnection, 2, aPayrollComponent(N_FOR_DATE_PAYROLL), sErrorDescription)
			If lErrorNumber = 0 Then
				iCounter = 0
				asFileContents = GetFileContents(PAYROLL_FILE2_PATH, sErrorDescription)
				If Len(asFileContents) > 0 Then
					asFileContents = Split(asFileContents, vbNewLine)
					For iIndex = 0 To UBound(asFileContents)
						If Len(asFileContents(iIndex)) > 0 Then
							asEmployeesQueries = Split(asFileContents(iIndex), LIST_SEPARATOR)
							If InStr(1, "," & sPeriods & ",", "," & asEmployeesQueries(2) & ",", vbbinaryCompare) > 0 Then
								sQueryBegin = ""
								asEmployeesQueries(8) = asEmployeesQueries(8) & sCondition
								If (InStr(1, asEmployeesQueries(8), "=Employees.", vbBinaryCompare) > 0) Or (InStr(1, asEmployeesQueries(8), "(Employees.", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", Employees"
								If InStr(1, asEmployeesQueries(8), "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
								If InStr(1, asEmployeesQueries(8), "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
								If InStr(1, asEmployeesQueries(8), "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
								If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, asEmployeesQueries(8), "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
								sErrorDescription = "No se pudieron obtener los empleados para registrar sus conceptos de pago en la nómina."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, Areas.EconomicZoneID, Zones.ZoneTypeID, Sum(ConceptAmount) As TotalAmount, IsDeduction From " & sTable & ", Concepts, EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Jobs, Areas, Zones " & sQueryBegin & " Where (" & sTable & ".ConceptID=Concepts.ConceptID) And (EmployeesHistoryList.EmployeeID=" & sTable & ".EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.JobID=Jobs.JobID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (" & sTable & ".RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (" & sTable & ".ConceptID In (" & asEmployeesQueries(3) & ")) And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") " & asEmployeesQueries(8) & " Group By EmployeesHistoryList.EmployeeID, Areas.EconomicZoneID, Zones.ZoneTypeID, IsDeduction Order By EmployeesHistoryList.EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										sCurrentID = CStr(oRecordset.Fields("EmployeeID").Value)
										iCurrentZoneID = CInt(oRecordset.Fields("EconomicZoneID").Value)
										iCurrentZoneTypeID = CInt(oRecordset.Fields("ZoneTypeID").Value)
										dAmount = 0
										dTaxAmount = CDbl(asEmployeesQueries(1)) / 100
										bMinMaxApplied = False
										If Not oRecordset.EOF Then
											Do While Not oRecordset.EOF
												If StrComp(sCurrentID, CStr(oRecordset.Fields("EmployeeID").Value), vbBinaryCompare) <> 0 Then
													If CDbl(asEmployeesQueries(4)) > 0 Then
														If CInt(asEmployeesQueries(5)) = 3 Then
															If (dAmount * dTaxAmount) < (CDbl(asEmployeesQueries(4)) * adDSM(iCurrentZoneTypeID)) Then
																dAmount = CDbl(asEmployeesQueries(4)) * adDSM(iCurrentZoneTypeID)
																bMinMaxApplied = True
															End If
														ElseIf CInt(asEmployeesQueries(5)) = 13 Then
															If (dAmount * dTaxAmount) < (CDbl(asEmployeesQueries(4)) * adDSM(iCurrentZoneID + 3)) Then
																dAmount = CDbl(asEmployeesQueries(4)) * adDSM(iCurrentZoneID + 3)
																bMinMaxApplied = True
															End If
														ElseIf CInt(asEmployeesQueries(5)) = 23 Then
															If dAmount < (CDbl(asEmployeesQueries(4)) * adDSM(iCurrentZoneTypeID)) Then
																dAmount = CDbl(asEmployeesQueries(4)) * adDSM(iCurrentZoneTypeID)
															End If
														ElseIf CInt(asEmployeesQueries(5)) = 24 Then
															If (dAmount * dTaxAmount) < (CDbl(asEmployeesQueries(4)) * adDSM(1)) Then
																dAmount = CDbl(asEmployeesQueries(4)) * adDSM(1)
																bMinMaxApplied = True
															End If
														ElseIf CInt(asEmployeesQueries(5)) = 25 Then
															If dAmount < (CDbl(asEmployeesQueries(4)) * adDSM(1)) Then
																dAmount = CDbl(asEmployeesQueries(4)) * adDSM(1)
															End If
														ElseIf CInt(asEmployeesQueries(5)) = 33 Then
															If dAmount < (CDbl(asEmployeesQueries(4)) * adDSM(iCurrentZoneID + 3)) Then
																dAmount = CDbl(asEmployeesQueries(4)) * adDSM(iCurrentZoneID + 3)
															End If
														ElseIf CInt(asEmployeesQueries(5)) = 34 Then
															If (dAmount * dTaxAmount) < (CDbl(asEmployeesQueries(4)) * adDSM(4)) Then
																dAmount = CDbl(asEmployeesQueries(4)) * adDSM(4)
																bMinMaxApplied = True
															End If
														ElseIf CInt(asEmployeesQueries(5)) = 35 Then
															If dAmount < (CDbl(asEmployeesQueries(4)) * adDSM(4)) Then
																dAmount = CDbl(asEmployeesQueries(4)) * adDSM(4)
															End If
														Else
															dAmount = dAmount * dTaxAmount
															bMinMaxApplied = True
															If dAmount < CDbl(asEmployeesQueries(4)) Then dAmount = CDbl(asEmployeesQueries(4))
														End If
													End If
													If CDbl(asEmployeesQueries(6)) > 0 Then
														If CInt(asEmployeesQueries(7)) = 3 Then
															If (dAmount * dTaxAmount) > (CDbl(asEmployeesQueries(6)) * adDSM(iCurrentZoneTypeID)) Then
																dAmount = CDbl(asEmployeesQueries(6)) * adDSM(iCurrentZoneTypeID)
																bMinMaxApplied = True
															End If
														ElseIf CInt(asEmployeesQueries(7)) = 13 Then
															If (dAmount * dTaxAmount) > (CDbl(asEmployeesQueries(6)) * adDSM(iCurrentZoneID + 3)) Then
																dAmount = CDbl(asEmployeesQueries(6)) * adDSM(iCurrentZoneID + 3)
																bMinMaxApplied = True
															End If
														ElseIf CInt(asEmployeesQueries(7)) = 23 Then
															If dAmount > (CDbl(asEmployeesQueries(6)) * adDSM(iCurrentZoneTypeID)) Then
																dAmount = CDbl(asEmployeesQueries(6)) * adDSM(iCurrentZoneTypeID)
															End If
														ElseIf CInt(asEmployeesQueries(7)) = 24 Then
															If (dAmount * dTaxAmount) > (CDbl(asEmployeesQueries(6)) * adDSM(1)) Then
																dAmount = CDbl(asEmployeesQueries(6)) * adDSM(1)
																bMinMaxApplied = True
															End If
														ElseIf CInt(asEmployeesQueries(7)) = 25 Then
															If dAmount > (CDbl(asEmployeesQueries(6)) * adDSM(1)) Then
																dAmount = CDbl(asEmployeesQueries(6)) * adDSM(1)
															End If
														ElseIf CInt(asEmployeesQueries(7)) = 33 Then
															If dAmount > (CDbl(asEmployeesQueries(6)) * adDSM(iCurrentZoneID + 3)) Then
																dAmount = CDbl(asEmployeesQueries(6)) * adDSM(iCurrentZoneID + 3)
															End If
														ElseIf CInt(asEmployeesQueries(7)) = 34 Then
															If (dAmount * dTaxAmount) > (CDbl(asEmployeesQueries(6)) * adDSM(4)) Then
																dAmount = CDbl(asEmployeesQueries(6)) * adDSM(4)
																bMinMaxApplied = True
															End If
														ElseIf CInt(asEmployeesQueries(7)) = 35 Then
															If dAmount > (CDbl(asEmployeesQueries(6)) * adDSM(4)) Then
																dAmount = CDbl(asEmployeesQueries(6)) * adDSM(4)
															End If
														Else
															dAmount = dAmount * dTaxAmount
															bMinMaxApplied = True
															If dAmount > CDbl(asEmployeesQueries(6)) Then dAmount = CDbl(asEmployeesQueries(6))
														End If
													End If
													If Not bMinMaxApplied Then dAmount = dAmount * dTaxAmount
													sErrorDescription = "No se pudo agregar el concepto de pago y su monto a la nómina del empleado."
													lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ", 1, " & sCurrentID & ", " & asEmployeesQueries(0) & ", 1, " & FormatNumber(dAmount, 2, True, False, False) & ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
													sCurrentID = CStr(oRecordset.Fields("EmployeeID").Value)
													iCurrentZoneID = CInt(oRecordset.Fields("EconomicZoneID").Value)
													iCurrentZoneTypeID = CInt(oRecordset.Fields("ZoneTypeID").Value)
													dAmount = 0
													bMinMaxApplied = False
												End If
												If CInt(oRecordset.Fields("IsDeduction").Value) = 0 Then
													dAmount = dAmount + CDbl(oRecordset.Fields("TotalAmount").Value)
												Else
													dAmount = dAmount - CDbl(oRecordset.Fields("TotalAmount").Value)
												End If
												oRecordset.MoveNext
												'If lErrorNumber <> 0 Then Exit Do
												If bTimeout Then Exit Do
											Loop
											oRecordset.Close

											If CDbl(asEmployeesQueries(4)) > 0 Then
												If CInt(asEmployeesQueries(5)) = 3 Then
													If (dAmount * dTaxAmount) < (CDbl(asEmployeesQueries(4)) * adDSM(iCurrentZoneTypeID)) Then
														dAmount = CDbl(asEmployeesQueries(4)) * adDSM(iCurrentZoneTypeID)
														bMinMaxApplied = True
													End If
												ElseIf CInt(asEmployeesQueries(5)) = 13 Then
													If (dAmount * dTaxAmount) < (CDbl(asEmployeesQueries(4)) * adDSM(iCurrentZoneID + 3)) Then
														dAmount = CDbl(asEmployeesQueries(4)) * adDSM(iCurrentZoneID + 3)
														bMinMaxApplied = True
													End If
												ElseIf CInt(asEmployeesQueries(5)) = 23 Then
													If dAmount < (CDbl(asEmployeesQueries(4)) * adDSM(iCurrentZoneTypeID)) Then
														dAmount = CDbl(asEmployeesQueries(4)) * adDSM(iCurrentZoneTypeID)
													End If
												ElseIf CInt(asEmployeesQueries(5)) = 24 Then
													If (dAmount * dTaxAmount) < (CDbl(asEmployeesQueries(4)) * adDSM(1)) Then
														dAmount = CDbl(asEmployeesQueries(4)) * adDSM(1)
														bMinMaxApplied = True
													End If
												ElseIf CInt(asEmployeesQueries(5)) = 25 Then
													If dAmount < (CDbl(asEmployeesQueries(4)) * adDSM(1)) Then
														dAmount = CDbl(asEmployeesQueries(4)) * adDSM(1)
													End If
												ElseIf CInt(asEmployeesQueries(5)) = 33 Then
													If dAmount < (CDbl(asEmployeesQueries(4)) * adDSM(iCurrentZoneID + 3)) Then
														dAmount = CDbl(asEmployeesQueries(4)) * adDSM(iCurrentZoneID + 3)
													End If
												ElseIf CInt(asEmployeesQueries(5)) = 34 Then
													If (dAmount * dTaxAmount) < (CDbl(asEmployeesQueries(4)) * adDSM(4)) Then
														dAmount = CDbl(asEmployeesQueries(4)) * adDSM(4)
														bMinMaxApplied = True
													End If
												ElseIf CInt(asEmployeesQueries(5)) = 35 Then
													If dAmount < (CDbl(asEmployeesQueries(4)) * adDSM(4)) Then
														dAmount = CDbl(asEmployeesQueries(4)) * adDSM(4)
													End If
												Else
													dAmount = dAmount * dTaxAmount
													bMinMaxApplied = True
													If dAmount < CDbl(asEmployeesQueries(4)) Then dAmount = CDbl(asEmployeesQueries(4))
												End If
											End If
											If CDbl(asEmployeesQueries(6)) > 0 Then
												If CInt(asEmployeesQueries(7)) = 3 Then
													If (dAmount * dTaxAmount) > (CDbl(asEmployeesQueries(6)) * adDSM(iCurrentZoneTypeID)) Then
														dAmount = CDbl(asEmployeesQueries(6)) * adDSM(iCurrentZoneTypeID)
														bMinMaxApplied = True
													End If
												ElseIf CInt(asEmployeesQueries(7)) = 13 Then
													If (dAmount * dTaxAmount) > (CDbl(asEmployeesQueries(6)) * adDSM(iCurrentZoneID + 3)) Then
														dAmount = CDbl(asEmployeesQueries(6)) * adDSM(iCurrentZoneID + 3)
														bMinMaxApplied = True
													End If
												ElseIf CInt(asEmployeesQueries(7)) = 23 Then
													If dAmount > (CDbl(asEmployeesQueries(6)) * adDSM(iCurrentZoneTypeID)) Then
														dAmount = CDbl(asEmployeesQueries(6)) * adDSM(iCurrentZoneTypeID)
													End If
												ElseIf CInt(asEmployeesQueries(7)) = 24 Then
													If (dAmount * dTaxAmount) > (CDbl(asEmployeesQueries(6)) * adDSM(1)) Then
														dAmount = CDbl(asEmployeesQueries(6)) * adDSM(1)
														bMinMaxApplied = True
													End If
												ElseIf CInt(asEmployeesQueries(7)) = 25 Then
													If dAmount > (CDbl(asEmployeesQueries(6)) * adDSM(1)) Then
														dAmount = CDbl(asEmployeesQueries(6)) * adDSM(1)
													End If
												ElseIf CInt(asEmployeesQueries(7)) = 33 Then
													If dAmount > (CDbl(asEmployeesQueries(6)) * adDSM(iCurrentZoneID + 3)) Then
														dAmount = CDbl(asEmployeesQueries(6)) * adDSM(iCurrentZoneID + 3)
													End If
												ElseIf CInt(asEmployeesQueries(7)) = 34 Then
													If (dAmount * dTaxAmount) > (CDbl(asEmployeesQueries(6)) * adDSM(4)) Then
														dAmount = CDbl(asEmployeesQueries(6)) * adDSM(4)
														bMinMaxApplied = True
													End If
												ElseIf CInt(asEmployeesQueries(7)) = 35 Then
													If dAmount > (CDbl(asEmployeesQueries(6)) * adDSM(4)) Then
														dAmount = CDbl(asEmployeesQueries(6)) * adDSM(4)
													End If
												Else
													dAmount = dAmount * dTaxAmount
													bMinMaxApplied = True
													If dAmount > CDbl(asEmployeesQueries(6)) Then dAmount = CDbl(asEmployeesQueries(6))
												End If
											End If
										End If
										If Not bMinMaxApplied Then dAmount = dAmount * dTaxAmount
										sErrorDescription = "No se pudo agregar el concepto de pago y su monto a la nómina del empleado."
										lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ", 1, " & sCurrentID & ", " & asEmployeesQueries(0) & ", 1, " & FormatNumber(dAmount, 2, True, False, False) & ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
									End If
								End If
							End If
						End If
						'If lErrorNumber <> 0 Then Exit For
						If bTimeout Then Exit For
					Next
				End If

				If Not bTimeout Then
					sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=4) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID In (1))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=5) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=7) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PositionTypeID <> 1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=8) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PositionTypeID <> 2)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=16) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=24) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=44) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=45) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=93) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PositionTypeID <> 1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=120) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID <> 1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

					If InStr(1, ",0715,1215,", "," & Right(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("MMDD")) & ",", vbBinaryCompare) > 0 Then
						sErrorDescription = "No se pudieron eliminar los conceptos de pagos de la nómina."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=33) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID Not In (Select EmployeeID From Employees Where (StartDate Like '%0101')))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					End If
					If InStr(1, ",0115,0615,", "," & Right(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("MMDD")) & ",", vbBinaryCompare) > 0 Then
						sErrorDescription = "No se pudieron eliminar los conceptos de pagos de la nómina."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=33) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID Not In (Select EmployeeID From Employees Where (StartDate Like '%0701')))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					End If
					If InStr(1, ",0415,1015,", "," & Right(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("MMDD")) & ",", vbBinaryCompare) > 0 Then
						sErrorDescription = "No se pudieron eliminar los conceptos de pagos de la nómina."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=33) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID Not In (Select EmployeeID From Employees Where (StartDate Like '%0101') Or (StartDate Like '%0701')))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					End If
					If InStr(1, ",1215,", "," & Right(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("MMDD")) & ",", vbBinaryCompare) > 0 Then
						sErrorDescription = "No se pudieron eliminar los conceptos de pagos de la nómina."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=107) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID Not In (Select EmployeeID From Employees Where (StartDate Like '%0101')))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					End If
					If InStr(1, ",0615,", "," & Right(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("MMDD")) & ",", vbBinaryCompare) > 0 Then
						sErrorDescription = "No se pudieron eliminar los conceptos de pagos de la nómina."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=107) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID Not In (Select EmployeeID From Employees Where (StartDate Like '%0701')))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					End If
				End If
			End If
		End If

		If Not bAdjustment Then
			lErrorNumber = CalculateQttyID_8_9(oRequest, oADODBConnection, True, bRetroactive, sErrorDescription)
		End If

Call DisplayTimeStamp("START: LEVEL 2, CREATE FILES, QttyID=15. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		If Not bTimeout Then
			lErrorNumber = CreateConceptsFile(oRequest, oADODBConnection, 15, aPayrollComponent(N_FOR_DATE_PAYROLL), sErrorDescription)
			If lErrorNumber = 0 Then
				iCounter = 0
				asFileContents = GetFileContents(PAYROLL_FILE15_PATH, sErrorDescription)
				If Len(asFileContents) > 0 Then
					asFileContents = Split(asFileContents, vbNewLine)
					For iIndex = 0 To UBound(asFileContents)
						If Len(asFileContents(iIndex)) > 0 Then
							asEmployeesQueries = Split(asFileContents(iIndex), LIST_SEPARATOR)
							If InStr(1, "," & sPeriods & ",", "," & asEmployeesQueries(2) & ",", vbbinaryCompare) > 0 Then
								sQueryBegin = ""
								asEmployeesQueries(7) = asEmployeesQueries(7) & sCondition
								If InStr(1, asEmployeesQueries(7), "(Jobs.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Jobs"
								If InStr(1, asEmployeesQueries(7), "(Zones.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Zones"
								If InStr(1, asEmployeesQueries(7), "(Areas.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Areas"
								If (InStr(1, asEmployeesQueries(7), "=Employees.", vbBinaryCompare) > 0) Or (InStr(1, asEmployeesQueries(7), "(Employees.", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", Employees"
								If InStr(1, asEmployeesQueries(7), "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
								If InStr(1, asEmployeesQueries(7), "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
								If InStr(1, asEmployeesQueries(7), "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
								If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, asEmployeesQueries(7), "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
								sErrorDescription = "No se pudieron obtener los empleados para registrar sus conceptos de pago en la nómina."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, Count(ChildID) As ChildrenCount From EmployeesChangesLKP, EmployeesHistoryList, EmployeesChildrenLKP, StatusEmployees, Reasons " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.EmployeeID=EmployeesChildrenLKP.EmployeeID) And (ChildEndDate=0) " & asEmployeesQueries(7) & " Group By EmployeesHistoryList.EmployeeID Order By EmployeesHistoryList.EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									Do While Not oRecordset.EOF
										lErrorNumber = AppendTextToFile(sFilePath & "_Payroll15_" & Int(iCounter / ROWS_PER_FILE) & ".txt", asEmployeesQueries(0) & SECOND_LIST_SEPARATOR & FormatNumber((CDbl(asEmployeesQueries(1)) * CInt(oRecordset.Fields("ChildrenCount").Value)), 2, True, False, False) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EmployeeID").Value), sErrorDescription)
										iCounter = iCounter + 1
										oRecordset.MoveNext
										'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
									Loop
									oRecordset.Close
								End If
							End If
						End If
						'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
						If bTimeout Then Exit For
					Next
				End If

				If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 2, RUN FROM FILES, QttyID=15, " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
					sQueryBegin = "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ", 1, "
					sQueryEnd = ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")"
					sCurrentID = ""
					For jIndex = 0 To iCounter Step ROWS_PER_FILE
						asFileContents = GetFileContents(sFilePath & "_Payroll15_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
						If Len(asFileContents) > 0 Then
							asFileContents = Split(asFileContents, vbNewLine)
							For iIndex = 0 To UBound(asFileContents)
								If Len(asFileContents(iIndex)) > 0 Then
									asEmployeesQueries = Split(asFileContents(iIndex), SECOND_LIST_SEPARATOR)
									sErrorDescription = "No se pudo agregar el concepto de pago y su monto a la nómina del empleado."
									If StrComp(sCurrentID, (asEmployeesQueries(2) & "," & asEmployeesQueries(0)), vbBinaryCompare) <> 0 Then
										lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, sQueryBegin & asEmployeesQueries(2) & ", " & asEmployeesQueries(0) & ", 1, " & asEmployeesQueries(1) & sQueryEnd, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
										sCurrentID = asEmployeesQueries(2) & "," & asEmployeesQueries(0)
									Else
										sErrorDescription = "No se pudo actualizar el concepto de pago en la nómina del empleado."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=ConceptAmount+" & asEmployeesQueries(1) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID=" & asEmployeesQueries(2) & ") And (ConceptID=" & asEmployeesQueries(0) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									End If
								End If
								'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
								If bTimeout Then Exit For
							Next
						End If
						Call DeleteFile(sFilePath & "_Payroll15_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
					Next
				End If
			End If
		End If

Call DisplayTimeStamp("START: LEVEL 2, CREATE FILES, QttyID=20. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		If Not bTimeout Then
			sQueryBegin = ""
			If InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Zones"
			If InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Areas"
			If InStr(1, sCondition, "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
			If InStr(1, sCondition, "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
			If InStr(1, sCondition, "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
			If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
			sErrorDescription = "No se pudieron obtener los créditos vigentes de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & aPayrollComponent(N_FOR_DATE_PAYROLL) & "' As RecordDate, '1' As RecordID, Credits.EmployeeID, Credits.CreditTypeID As ConceptID, '1' As PayrollTypeID, Credits.PaymentAmount As ConceptAmount, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Credits, Concepts, EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons" & sQueryBegin & " Where (Credits.CreditTypeID=Concepts.ConceptID) And (Credits.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (Credits.CreditTypeID>0) And (Credits.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Credits.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Credits.QttyID=1) And ((Credits.PaymentsCounter<Credits.PaymentsNumber) Or (Credits.PaymentsNumber<1)) And (Credits.Active=1) And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") " & sCondition & Replace(sConceptCondition, "Concepts.ConceptID", "Credits.CreditTypeID") & " Order By EmployeesHistoryList.EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If

Call DisplayTimeStamp("START: LEVEL 2, CREATE FILES, QttyID=21. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		If Not bTimeout Then
			sQueryBegin = ""
			If InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Zones"
			If InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Areas"
			If InStr(1, sCondition, "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
			If InStr(1, sCondition, "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
			If InStr(1, sCondition, "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
			If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
			sErrorDescription = "No se pudieron obtener los créditos vigentes de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Credits.EmployeeID, Credits.CreditTypeID, Credits.PaymentAmount, Credits.AppliesToID From Credits, Concepts, EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons " & sQueryBegin & " Where (Credits.CreditTypeID=Concepts.ConceptID) And (Credits.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (Credits.CreditTypeID>0) And (Credits.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Credits.QttyID=2) And ((Credits.PaymentsCounter<Credits.PaymentsNumber) Or (Credits.PaymentsNumber<1)) And (Credits.Active=1) And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Credits.AppliesToID Is Not Null) " & sCondition & Replace(sConceptCondition, "Concepts.ConceptID", "Credits.CreditTypeID") & " Order By Credits.EmployeeID, Concepts.OrderInList, Concepts.ConceptID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			iCounter = 0
			If lErrorNumber = 0 Then
				Do While Not oRecordset.EOF
					lErrorNumber = AppendTextToFile(sFilePath & "_Payroll21a_" & Int(iCounter / ROWS_PER_FILE) & ".txt", CStr(oRecordset.Fields("EmployeeID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("CreditTypeID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("PaymentAmount").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("AppliesToID").Value), sErrorDescription)
					iCounter = iCounter + 1
					oRecordset.MoveNext
					'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
				Loop
				oRecordset.Close
			End If
		End If

		If Not bTimeout Then
			If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 2, RUN FROM FILES, QttyID=21a, " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
				iCounter2 = 0
				For jIndex = 0 To iCounter Step ROWS_PER_FILE
					asFileContents = GetFileContents(sFilePath & "_Payroll21a_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
					If Len(asFileContents) > 0 Then
						asFileContents = Split(asFileContents, vbNewLine)
						For iIndex = 0 To UBound(asFileContents)
							If Len(asFileContents(iIndex)) > 0 Then
								asEmployeesQueries = Split(asFileContents(iIndex), LIST_SEPARATOR)
								sErrorDescription = "No se pudieron obtener los conceptos de pagos y sus montos."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(ConceptAmount) As TotalAmount, IsDeduction From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", Concepts Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Concepts.ConceptID) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID=" & asEmployeesQueries(0) & ") And (Concepts.ConceptID In (" & asEmployeesQueries(3) & ")) And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") Group By IsDeduction", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									dAmount = 0
									Do While Not oRecordset.EOF
										If CInt(oRecordset.Fields("IsDeduction").Value) = 0 Then
											dAmount = dAmount + CDbl(oRecordset.Fields("TotalAmount").Value)
										Else
											dAmount = dAmount - CDbl(oRecordset.Fields("TotalAmount").Value)
										End If
										oRecordset.MoveNext
										'If lErrorNumber <> 0 Then Exit Do
									Loop
									oRecordset.Close
									lErrorNumber = AppendTextToFile(sFilePath & "_Payroll21b_" & Int(iCounter2 / ROWS_PER_FILE) & ".txt", asEmployeesQueries(0) & "," & asEmployeesQueries(1) & LIST_SEPARATOR & FormatNumber((CDbl(asEmployeesQueries(2)) * dAmount / 100), 2, True, False, False), sErrorDescription)
									iCounter2 = iCounter2 + 1
								End If
							End If
							'If lErrorNumber <> 0 Then Exit For
							If bTimeout Then Exit For
						Next
					End If
					Call DeleteFile(sFilePath & "_Payroll21a_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
				Next
			End If
		End If

		If Not bTimeout Then
			If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 2, RUN FROM FILES, QttyID=21b, " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
				sQueryBegin = "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ", 1, "
				sQueryEnd = ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")"
				For jIndex = 0 To iCounter2 Step ROWS_PER_FILE
					asFileContents = GetFileContents(sFilePath & "_Payroll21b_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
					If Len(asFileContents) > 0 Then
						asFileContents = Split(asFileContents, vbNewLine)
						For iIndex = 0 To UBound(asFileContents)
							If Len(asFileContents(iIndex)) > 0 Then
								asEmployeesQueries = Split(asFileContents(iIndex), LIST_SEPARATOR)
								sErrorDescription = "No se pudo agregar el concepto de pago y su monto a la nómina del empleado."
								lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, sQueryBegin & asEmployeesQueries(0) & ", 1, " & asEmployeesQueries(1) & sQueryEnd, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
							End If
							'If lErrorNumber <> 0 Then Exit For
							If bTimeout Then Exit For
						Next
					End If
					Call DeleteFile(sFilePath & "_Payroll21b_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
				Next
			End If
		End If
	End If

Call DisplayTimeStamp("START: LEVEL 2, RECLAMOS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
	sEmployeeIDs = "-2"
	If bAdjustment Then
		If Not bTimeout Then
			sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct EmployeeID From EmployeesAdjustmentsLKP Where (EmployeeID Not In (Select Distinct EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & lPayID & "))) And (MissingDate=" & lPayID & ") And (PayrollDate In (0," & aPayrollComponent(N_ID_PAYROLL) & ")) And (Active=1) Order By EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				Do While Not oRecordset.EOF
					sEmployeeIDs = sEmployeeIDs & "," & CStr(oRecordset.Fields("EmployeeID").Value)
					oRecordset.MoveNext
				Loop
			End If
			oRecordset.Close
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesChangesLKP Where (EmployeeID In (" & sEmployeeIDs & ")) And (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate=" & lPayID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
				lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollID, '" & lPayID & "' As PayrollDate, Max(EmployeeDate) As EmployeeDate1, 0 As FirstDate, 0 As LastDate, 0 As Concepts40 From EmployeesHistoryList Where (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (" & sEmployeeIDs & ")) Group By EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
			End If
			sEmployeeIDs = "-2"
			Call BuildCondition("", sQueryBegin)
			If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
			sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct EmployeeID From EmployeesAdjustmentsLKP Where (EmployeeID Not In (Select Distinct EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & lPayID & "))) And (MissingDate=" & lPayID & ") And (PayrollDate In (0," & aPayrollComponent(N_ID_PAYROLL) & ")) And (Active=1) And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons" & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) " & sCondition & ")) Order By EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				Do While Not oRecordset.EOF
					sEmployeeIDs = sEmployeeIDs & "," & CStr(oRecordset.Fields("EmployeeID").Value)
					oRecordset.MoveNext
				Loop
			End If
			oRecordset.Close
		End If
		sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID In (" & sEmployeeIDs & "))"
	End If

	If Not bTimeout Then
		sQueryBegin = ""
		If InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Zones"
		If InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Areas"
		If InStr(1, sCondition, "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
		If InStr(1, sCondition, "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
		If InStr(1, sCondition, "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
		If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
		sErrorDescription = "No se pudieron obtener los reclamos vigentes de los empleados."
		If bAdjustment Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, ConceptID, ConceptAmount, MissingDate From EmployeesAdjustmentsLKP, EmployeesChangesLKP, EmployeesHistoryList" & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesAdjustmentsLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesAdjustmentsLKP.MissingDate In (0," & lPayID & ")) And (EmployeesAdjustmentsLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesAdjustmentsLKP.Active=1) " & sCondition & Replace(sConceptCondition, "Concepts.", "EmployeesAdjustmentsLKP."), "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, ConceptID, ConceptAmount, MissingDate From EmployeesAdjustmentsLKP, EmployeesChangesLKP, EmployeesHistoryList" & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesAdjustmentsLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesAdjustmentsLKP.PayrollDate In (0," & lPayID & ")) And (EmployeesAdjustmentsLKP.Active=1) " & sCondition & Replace(sConceptCondition, "Concepts.", "EmployeesAdjustmentsLKP."), "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		End If
		iCounter = 0
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				lErrorNumber = AppendTextToFile(sFilePath & "_Payroll30_" & Int(iCounter / ROWS_PER_FILE) & ".txt", CStr(oRecordset.Fields("MissingDate").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("EmployeeID").Value) & ", " & CStr(oRecordset.Fields("ConceptID").Value) & ", 1, " & FormatNumber(CStr(oRecordset.Fields("ConceptAmount").Value), 2, True, False, False), sErrorDescription)
				iCounter = iCounter + 1
				oRecordset.MoveNext
				'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
		End If
	End If

	If Not bTimeout Then
		If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 2, RUN FROM FILES, RECLAMOS, " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
			sQueryBegin = "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values ("
			sQueryEnd = ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")"
			For jIndex = 0 To iCounter Step ROWS_PER_FILE
				asFileContents = GetFileContents(sFilePath & "_Payroll30_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
				If Len(asFileContents) > 0 Then
					asFileContents = Split(asFileContents, vbNewLine)
					For iIndex = 0 To UBound(asFileContents)
						If Len(asFileContents(iIndex)) > 0 Then
							asEmployeesQueries = Split(asFileContents(iIndex), LIST_SEPARATOR)
							sErrorDescription = "No se pudo agregar el concepto de pago y su monto a la nómina del empleado."
							lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, sQueryBegin & asEmployeesQueries(0) & ", -" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ", " & asEmployeesQueries(1) & sQueryEnd, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
						End If
						'If lErrorNumber <> 0 Then Exit For
						If bTimeout Then Exit For
					Next
				End If
				Call DeleteFile(sFilePath & "_Payroll30_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
			Next
		End If
	End If

	If Not bTimeout Then
		sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=4) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID In (1))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=5) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=7) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PositionTypeID <> 1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=8) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PositionTypeID <> 2)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=16) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=24) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=44) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=45) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=93) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.PositionTypeID <> 1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=120) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID <> 1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	If Not bAdjustment Then
Call DisplayTimeStamp("START: LEVEL 2, ESPECIALES. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		If Not bTimeout Then
			asSpecialConcepts = Split("18,21,29,30,34,40,41,42,43,44,50,52,70,71,72,90,92,96,151", ",") '15,18,24,26,31,37,38,39,40,46,49,50,69,70,71,B3,B7,C5,RQ
			iCounter = 0
			For kIndex = 0 To UBound(asSpecialConcepts)
				If (Len(oRequest("PayrollConceptID").Item) = 0) Or (InStr(1, "," & Replace(oRequest("PayrollConceptID").Item, " ", "") & ",", "," & asSpecialConcepts(kIndex) & ",", vbBinaryCompare) > 0) Then
					Select Case CInt(asSpecialConcepts(kIndex))
						Case 18, 96 '15. Remuneraciones por guardias | C5. Guardias PROVAC
							sQueryBegin = ""
							If InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Zones"
							If InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Areas"
							If InStr(1, sCondition, "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
							If InStr(1, sCondition, "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
							If InStr(1, sCondition, "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
							If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"

							sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
							Select Case CInt(asSpecialConcepts(kIndex))
								Case 18 '15. Remuneraciones por guardias
									sTemp = "423"
								Case 96 'C5. Guardias PROVAC
									sTemp = "426"
							End Select
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesSpecialJourneys.EmployeeID, EmployeesSpecialJourneys.EmployeeID As OriginalEmployeeID, EmployeesSpecialJourneys.SpecialJourneyID, EmployeesSpecialJourneys.StartDate, EmployeesSpecialJourneys.EndDate, EmployeesHistoryList.WorkingHours, EmployeesSpecialJourneys.WorkedHours From EmployeesChangesLKP, EmployeesHistoryList, EmployeesSpecialJourneys Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesSpecialJourneys.EmployeeID) And (EmployeesSpecialJourneys.EmployeeID<800000) And (SpecialJourneyID In (" & sTemp & ")) And (AppliedDate=" & lPayID & ") And (EmployeesSpecialJourneys.EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & ")) Order By EmployeesSpecialJourneys.EmployeeID, StartDate", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Select EmployeesSpecialJourneys.EmployeeID, EmployeesSpecialJourneys.EmployeeID As OriginalEmployeeID, EmployeesSpecialJourneys.SpecialJourneyID, EmployeesSpecialJourneys.StartDate, EmployeesSpecialJourneys.EndDate, EmployeesHistoryList.WorkingHours, EmployeesSpecialJourneys.WorkedHours From EmployeesChangesLKP, EmployeesHistoryList, EmployeesSpecialJourneys Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesSpecialJourneys.EmployeeID) And (EmployeesSpecialJourneys.EmployeeID<800000) And (SpecialJourneyID In (" & sTemp & ")) And (AppliedDate=" & lPayID & ") And (EmployeesSpecialJourneys.EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & ")) Order By EmployeesSpecialJourneys.EmployeeID, StartDate -->" & vbNewLine
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									asFileContents = ""
									Do While Not oRecordset.EOF
										dAmount = CDbl(oRecordset.Fields("WorkedHours").Value) * (Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1)
										asFileContents = asFileContents & CStr(oRecordset.Fields("EmployeeID").Value) & "," & CStr(oRecordset.Fields("OriginalEmployeeID").Value) & "," & CStr(oRecordset.Fields("StartDate").Value) & "," & dAmount & "," & CStr(oRecordset.Fields("WorkingHours").Value) & ";"
										oRecordset.MoveNext
										'If Err.number <> 0 Then Exit Do
									Loop
									oRecordset.Close
									If Len(asFileContents) > 0 Then asFileContents = Left(asFileContents, (Len(asFileContents) - Len(";")))
									asFileContents = Split(asFileContents, ";")
									lCurrentID = -2
									dAmount = 0
									For iIndex = 0 To UBound(asFileContents)
										asFileContents(iIndex) = Split(asFileContents(iIndex), ",")
										If lCurrentID <> CLng(asFileContents(iIndex)(0)) Then
											If lCurrentID <> -2 Then
												sErrorDescription = "No se pudieron agregar los montos de la nómina de los empleados."
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & lPayID & ", 1, " & lCurrentID & ", " & asSpecialConcepts(kIndex) & ", 1, " & dAmount & ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
											End If
											dAmount = 0
											lCurrentID = CLng(asFileContents(iIndex)(0))
										End If
										sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
										If aPayrollComponent(N_ID_PAYROLL) > 29999999 Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & Left(asFileContents(iIndex)(2), Len("YYYY")) & " Where (EmployeeID=" & asFileContents(iIndex)(1) & ") And (RecordID=" & GetPayrollEndDate(asFileContents(iIndex)(2)) & ") And (ConceptID In (1,4,38,130)) Group By ConceptID Order By ConceptID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										ElseIf aPayrollComponent(N_ID_PAYROLL) <> lPayID Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & Left(lPayID, Len("YYYY")) & " Where (EmployeeID=" & asFileContents(iIndex)(1) & ") And (RecordID=" & lPayID & ") And (ConceptID In (1,4,38,130)) Group By ConceptID Order By ConceptID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										Else
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID=" & asFileContents(iIndex)(1) & ") And (RecordDate=" & lPayID & ") And (ConceptID In (1,4,38,130)) Group By ConceptID Order By ConceptID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										End If
										If lErrorNumber = 0 Then
											dTemp = 0
											Do While Not oRecordset.EOF
												Select Case CLng(oRecordset.Fields("ConceptID").Value)
													Case 1
														Select Case 3
															Case 1
																dTemp = dTemp + (CDbl(oRecordset.Fields("TotalAmount").Value) * 1.1)
															Case 2
																dTemp = dTemp + (CDbl(oRecordset.Fields("TotalAmount").Value) * 1.2)
															Case Else
																dTemp = dTemp + CDbl(oRecordset.Fields("TotalAmount").Value)
														End Select
													Case 4, 38, 130
														dTemp = dTemp + CDbl(oRecordset.Fields("TotalAmount").Value)
												End Select
												oRecordset.MoveNext
												'If Err.number <> 0 Then Exit Do
											Loop
											oRecordset.Close
										End If
										dTemp = (dTemp / 15) * 2
										dAmount = dAmount + CDbl(FormatNumber(((dTemp / asFileContents(iIndex)(4)) * asFileContents(iIndex)(3)), 2, True, False, False))
										If bTimeout Then Exit For
									Next
									sErrorDescription = "No se pudieron agregar los montos de la nómina de los empleados."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & lPayID & ", 1, " & lCurrentID & ", " & asSpecialConcepts(kIndex) & ", 1, " & dAmount & ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								End If
							End If
						Case 21 '18. Prima de vacaciones
							sQueryBegin = ""
							If InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Zones"
							If InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Areas"
							If InStr(1, sCondition, "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
							If InStr(1, sCondition, "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
							If InStr(1, sCondition, "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
							If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"

							lTempEndDate = 0 '515
							sErrorDescription = "No se pudo obtener la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PeriodDate From Concepts, Periods Where (Concepts.PeriodID=Periods.PeriodID) And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID=21)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									adTotal = Split(oRecordset.Fields("PeriodDate").Value, ",")
									For iIndex = 0 To UBound(adTotal)
										If (CInt(Right(lPayID, Len("0000"))) = CInt(adTotal(iIndex))) Then
											lTempEndDate = CLng(adTotal(iIndex))
											Exit For
										End If
									Next
								End If
							End If
							If (CInt(Right(lPayID, Len("0000"))) = lTempEndDate) Or (lTempEndDate = -1) Then
								sErrorDescription = "No se pudo limpiar la tabla temporal."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll -->" & vbNewLine
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudieron obtener los montos de las nóminas de los empleados."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & aPayrollComponent(N_FOR_DATE_PAYROLL) & "' As RecordDate, '1' As RecordID, EmployeesHistoryList.EmployeeID, '21' As ConceptID, '1' As PayrollTypeID, (Sum(ConceptAmount) / 15) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '-1' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Employees " & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (ConceptID In (1,7,8,89)) " & sCondition & " Group By EmployeesHistoryList.EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & aPayrollComponent(N_FOR_DATE_PAYROLL) & "' As RecordDate, '1' As RecordID, EmployeesHistoryList.EmployeeID, '21' As ConceptID, '1' As PayrollTypeID, (Sum(ConceptAmount) / 15) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '-1' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Employees " & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (ConceptID In (1,7,8,89)) " & sCondition & " Group By EmployeesHistoryList.EmployeeID -->" & vbNewLine
								End If
								lTempStartDate = (CInt(Left(lPayID, Len("YYYY"))) * 10000) + lTempStartDate
								lTempStartDate = Left(GetSerialNumberForDate(DateAdd("d", -1, DateAdd("m", -6, GetDateFromSerialNumber(lTempStartDate)))), Len("00000000"))
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo limpiar la tabla temporal."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (Select EmployeeID From Employees Where (StartDate>" & lTempStartDate & ")))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (Select EmployeeID From Employees Where (StartDate>" & lTempStartDate & "))) -->" & vbNewLine
								End If
								sErrorDescription = "No se pudieron obtener las incidencias de los empleados."
'								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesAbsencesLKP.EmployeeID, OcurredDate, EndDate From EmployeesAbsencesLKP, Payroll Where (EmployeesAbsencesLKP.EmployeeID=Payroll.EmployeeID) And (AbsenceID In (89)) And (OcurredDate<" & Left(lPayID, Len("0000")) & "0000) And (EndDate>" & Left(lPayID, Len("0000")) & lTempEndDate & ") Order By EmployeesAbsencesLKP.EmployeeID, OcurredDate", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
'								Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Select EmployeesAbsencesLKP.EmployeeID, OcurredDate, EndDate From EmployeesAbsencesLKP, Payroll Where (EmployeesAbsencesLKP.EmployeeID=Payroll.EmployeeID) And (AbsenceID In (89)) And (OcurredDate<" & Left(lPayID, Len("0000")) & "0000) And (EndDate>" & Left(lPayID, Len("0000")) & lTempEndDate & ") Order By EmployeeID, OcurredDate -->" & vbNewLine
								lStartDate = Left(lPayID, Len("0000")) & "0000"
								lEndDate = Left(lPayID, Len("0000")) & lTempEndDate
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesAbsencesLKP.EmployeeID, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate From EmployeesAbsencesLKP, Payroll, EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Employees " & sQueryBegin & " Where (EmployeesAbsencesLKP.EmployeeID=Payroll.EmployeeID) And (Payroll.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Payroll.RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesAbsencesLKP.AbsenceID In (89)) And (((EmployeesAbsencesLKP.EndDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ") And (EmployeesAbsencesLKP.EndDate>=" & lStartDate & "))) " & sCondition & " Order By EmployeesAbsencesLKP.EmployeeID, EmployeesAbsencesLKP.OcurredDate", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Select EmployeesAbsencesLKP.EmployeeID, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate From EmployeesAbsencesLKP, Payroll, EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Employees " & sQueryBegin & " Where (EmployeesAbsencesLKP.EmployeeID=Payroll.EmployeeID) And (Payroll.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Payroll.RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesAbsencesLKP.AbsenceID In (89)) And (((EmployeesAbsencesLKP.EndDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ") And (EmployeesAbsencesLKP.EndDate>=" & lStartDate & "))) " & sCondition & " Order By EmployeesAbsencesLKP.EmployeeID, EmployeesAbsencesLKP.OcurredDate -->" & vbNewLine
								If lErrorNumber = 0 Then
									If Not oRecordset.EOF Then
										lTempStartDate = (CInt(Left(lPayID, Len("YYYY"))) * 10000) + 101
										dAmount = 0
										lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
										Do While Not oRecordset.EOF
											If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
												sErrorDescription = "No se pudieron obtener las incidencias de los empleados."
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll Set ConceptTaxes=" & Int(dAmount / 15) & " Where (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
												dAmount = 0
												lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
											End If
											lStartDate = CLng(oRecordset.Fields("OcurredDate").Value)
											lEndDate = CLng(oRecordset.Fields("EndDate").Value)
											If lStartDate < lTempStartDate Then lStartDate = lTempStartDate
											If lEndDate > lPayID Then lEndDate = lPayID
											If lStartDate <= lEndDate Then
												dAmount = dAmount + (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDate), GetDateFromSerialNumber(lEndDate))) + 1)
											End If
											oRecordset.MoveNext
											'If Err.number <> 0 Then Exit Do
										Loop
										oRecordset.Close
										sErrorDescription = "No se pudieron obtener las incidencias de los empleados."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll Set ConceptTaxes=" & Int(dAmount / 15) & " Where (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									End If
								End If
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo limpiar la tabla temporal."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (ConceptTaxes>=10)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (ConceptTaxes>=10) -->" & vbNewLine
								End If
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudieron obtener los montos de la nómina de los empleados."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, '1' As RecordID, EmployeeID, '21' As ConceptID, PayrollTypeID, ((ConceptAmount * (10 - ConceptTaxes)) / 2) As ConceptAmount1, '0' As ConceptTaxes, ConceptRetention, UserID From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, '1' As RecordID, EmployeeID, '21' As ConceptID, PayrollTypeID, ((ConceptAmount * (10 - ConceptTaxes)) / 2) As ConceptAmount1, '0' As ConceptTaxes, ConceptRetention, UserID From Payroll -->" & vbNewLine
								End If

								sErrorDescription = "No se pudo limpiar la tabla temporal."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll -->" & vbNewLine
							End If
						Case 29 '24. Estímulo adicional por antigüedad
							If CInt(Right(lPayID, Len("0000"))) = 531 Then
								sErrorDescription = "No se pudo limpiar la tabla temporal."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll -->" & vbNewLine
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudieron obtener los montos de las nóminas de los empleados."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "' As RecordDate, '0' As RecordID, EmployeesHistoryList.EmployeeID, ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '-1' As UserID From Payroll_" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Employees " & sQueryBegin & " Where (Payroll_" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Payroll_" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & ".RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (ConceptID In (1,4,89)) And (RecordID>=" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "1001) And (RecordID<=" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "9999) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, ConceptID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "' As RecordDate, '0' As RecordID, EmployeesHistoryList.EmployeeID, ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '-1' As UserID From Payroll_" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Employees " & sQueryBegin & " Where (Payroll_" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Payroll_" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & ".RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (ConceptID In (1,4,89)) And (RecordID>=" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "1001) And (RecordID<=" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "9999) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, ConceptID -->" & vbNewLine
								End If
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudieron obtener los montos de las nóminas de los empleados."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "' As RecordDate, '0' As RecordID, EmployeesHistoryList.EmployeeID, ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '-1' As UserID From Payroll_" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Employees " & sQueryBegin & " Where (Payroll_" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Payroll_" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & ".RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (ConceptID In (1,4,89)) And (RecordID>=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0000) And (RecordID<=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0331) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, ConceptID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "' As RecordDate, '0' As RecordID, EmployeesHistoryList.EmployeeID, ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '-1' As UserID From Payroll_" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Employees " & sQueryBegin & " Where (Payroll_" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Payroll_" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & ".RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (ConceptID In (1,4,89)) And (RecordID>=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0000) And (RecordID<=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0331) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, ConceptID -->" & vbNewLine
								End If
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo limpiar la tabla temporal."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (Select EmployeeID From Employees Where (PositionTypeID<>1) Or (StartDate>" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "0930)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (Select EmployeeID From Employees Where (PositionTypeID<>1) Or (StartDate>" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "0930))) -->" & vbNewLine
								End If
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo limpiar la tabla temporal."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID Not In (Select EmployeeID From EmployeesHistoryList Where (PositionTypeID=1) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And ((EmployeeDate=" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "1001) And (EndDate=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0331)) Or ((EmployeeDate=" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "1001) And (EndDate>" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0331)) Or ((EmployeeDate<" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "1001) And (EndDate=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0331)) Or ((EmployeeDate<" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "1001) And (EndDate>" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0331))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID Not In (Select EmployeeID From EmployeesHistoryList Where (PositionTypeID=1) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And ((EmployeeDate=" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "1001) And (EndDate=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0331)) Or ((EmployeeDate=" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "1001) And (EndDate>" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0331)) Or ((EmployeeDate<" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "1001) And (EndDate=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0331)) Or ((EmployeeDate<" & CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1 & "1001) And (EndDate>" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0331)))) -->" & vbNewLine
								End If
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo limpiar la tabla temporal."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (-2))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (-2)) -->" & vbNewLine
								End If
							End If
							If CInt(Right(lPayID, Len("0000"))) = 1231 Then
								sErrorDescription = "No se pudo limpiar la tabla temporal."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll -->" & vbNewLine
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudieron obtener los montos de las nóminas de los empleados."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "' As RecordDate, '0' As RecordID, EmployeesHistoryList.EmployeeID, ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '-1' As UserID From Payroll_" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Employees " & sQueryBegin & " Where (Payroll_" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Payroll_" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & ".RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (ConceptID In (1,4,89)) And (RecordID>=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0401) And (RecordID<=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0930) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, ConceptID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "' As RecordDate, '0' As RecordID, EmployeesHistoryList.EmployeeID, ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '-1' As UserID From Payroll_" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Employees " & sQueryBegin & " Where (Payroll_" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Payroll_" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & ".RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (ConceptID In (1,4,89)) And (RecordID>=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0401) And (RecordID<=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0930) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, ConceptID -->" & vbNewLine
								End If
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo limpiar la tabla temporal."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (Select EmployeeID From Employees Where (PositionTypeID<>1) Or (StartDate>" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0330)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (Select EmployeeID From Employees Where (PositionTypeID<>1) Or (StartDate>" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0330))) -->" & vbNewLine
								End If
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo limpiar la tabla temporal."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID Not In (Select EmployeeID From EmployeesHistoryList Where (PositionTypeID=1) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And ((EmployeeDate=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0401) And (EndDate=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0930)) Or ((EmployeeDate=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0401) And (EndDate>" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0930)) Or ((EmployeeDate<" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0401) And (EndDate=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0930)) Or ((EmployeeDate<" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0401) And (EndDate>" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0930))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID Not In (Select EmployeeID From EmployeesHistoryList Where (PositionTypeID=1) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And ((EmployeeDate=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0401) And (EndDate=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0930)) Or ((EmployeeDate=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0401) And (EndDate>" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0930)) Or ((EmployeeDate<" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0401) And (EndDate=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0930)) Or ((EmployeeDate<" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0401) And (EndDate>" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & "0930)))) -->" & vbNewLine
								End If
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo limpiar la tabla temporal."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (-2))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (-2)) -->" & vbNewLine
								End If
							End If
							iCounter2 = 0
							If (CInt(Right(lPayID, Len("0000"))) = 531) Or (CInt(Right(lPayID, Len("0000"))) = 1231) Then
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudieron obtener los montos de las nóminas de los empleados."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll Group By EmployeeID, ConceptID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Select EmployeeID, ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll Group By EmployeeID, ConceptID -->" & vbNewLine
									If lErrorNumber = 0 Then
										lCurrentID = -2
										adTotal = Split("0,0", ",")
										Do While Not oRecordset.EOF
											If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
												If lCurrentID <> -2 Then
													dAmount = (adTotal(0) + (adTotal(1) * 3 / 6.5)) / 15
													lErrorNumber = AppendTextToFile(sFilePath & "_Payroll_29.txt", lCurrentID & LIST_SEPARATOR & FormatNumber(dAmount, 2, True, False, False), sErrorDescription)
													iCounter2 = iCounter2 + 1
												End If
												lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
												adTotal(0) = 0
												adTotal(1) = 0
												dAmount = 0
											End If
											Select Case CLng(oRecordset.Fields("ConceptID").Value)
												Case 1
													adTotal(0) = adTotal(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
													adTotal(1) = adTotal(1) + CDbl(oRecordset.Fields("TotalAmount").Value)
												Case Else
													adTotal(0) = adTotal(0) + CDbl(oRecordset.Fields("TotalAmount").Value)
											End Select
											oRecordset.MoveNext
											'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
										Loop
										oRecordset.Close
										dAmount = (adTotal(0) + (adTotal(1) * 3 / 6.5)) / 15
										lErrorNumber = AppendTextToFile(sFilePath & "_Payroll_29.txt", lCurrentID & LIST_SEPARATOR & FormatNumber(dAmount, 2, True, False, False), sErrorDescription)
										iCounter2 = iCounter2 + 1
									End If
									If iCounter2 > 0 Then
										adTotal = Split(",2,3,;10.5;;;,4,;13;;;,5,;15.5;;;,6,;18;;;,7,8,9,;20.5", ";;;")
										For iIndex = 0 To UBound(adTotal)
											adTotal(iIndex) = Split(adTotal(iIndex), ";")
											adTotal(iIndex)(1) = CDbl(adTotal(iIndex)(1))
										Next
										asFileContents = GetFileContents(sFilePath & "_Payroll_29.txt", sErrorDescription)
										If Len(asFileContents) > 0 Then
											asFileContents = Split(asFileContents, vbNewLine)
											For iIndex = 0 To UBound(asFileContents)
												If Len(asFileContents(iIndex)) > 0 Then
													asEmployeesQueries = Split(asFileContents(iIndex), LIST_SEPARATOR)
													sErrorDescription = "No se pudieron obtener las incidencias de los empleados."
													If CInt(Right(lPayID, Len("0000"))) = 531 Then
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AbsenceID, AntiquityID, OcurredDate, EndDate, AbsenceHours From EmployeesAbsencesLKP, Employees Where (EmployeesAbsencesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=" & asEmployeesQueries(0) & ") And (AbsenceID In (29,30,82,83,89)) And (OcurredDate<" & Left(lPayID, Len("0000")) & "0331) And (EndDate>" & CInt(Left(lPayID, Len("0000"))) - 1 & "1001) Order By AbsenceID, OcurredDate", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
														Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Select AbsenceID, AntiquityID, OcurredDate, EndDate, AbsenceHours From EmployeesAbsencesLKP, Employees Where (EmployeesAbsencesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=" & asEmployeesQueries(0) & ") And (AbsenceID In (29,30,82,83,89)) And (OcurredDate<" & Left(lPayID, Len("0000")) & "0331) And (EndDate>" & CInt(Left(lPayID, Len("0000"))) - 1 & "1001) Order By AbsenceID, OcurredDate -->" & vbNewLine
													ElseIf CInt(Right(lPayID, Len("0000"))) = 1231 Then
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AbsenceID, AntiquityID, OcurredDate, EndDate, AbsenceHours From EmployeesAbsencesLKP, Employees Where (EmployeesAbsencesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=" & asEmployeesQueries(0) & ") And (AbsenceID In (29,30,82,83,89)) And (OcurredDate<" & Left(lPayID, Len("0000")) & "0930) And (EndDate>" & Left(lPayID, Len("0000")) & "0401) Order By AbsenceID, OcurredDate", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
														Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Select AbsenceID, AntiquityID, OcurredDate, EndDate, AbsenceHours From EmployeesAbsencesLKP, Employees Where (EmployeesAbsencesLKP.EmployeeID=Employees.EmployeeID) And (Employees.EmployeeID=" & asEmployeesQueries(0) & ") And (AbsenceID In (29,30,82,83,89)) And (OcurredDate<" & Left(lPayID, Len("0000")) & "0930) And (EndDate>" & Left(lPayID, Len("0000")) & "0401) Order By AbsenceID, OcurredDate -->" & vbNewLine
													End If
													If lErrorNumber = 0 Then
														dAmount = 0
														For jIndex = 0 To UBound(adTotal)
															If InStr(1, adTotal(jIndex)(0), "," & CInt(oRecordset.Fields("AntiquityID").Value) & ",", vbBinaryCompare) > 0 Then Exit For
														Next
														Do While Not oRecordset.EOF
															lStartDate = CLng(oRecordset.Fields("OcurredDate").Value)
															lEndDate = CLng(oRecordset.Fields("EndDate").Value)
															If CInt(Right(lPayID, Len("0000"))) = 531 Then
																If lStartDate < CLng(Left(lPayID, Len("0000")) & "1001") Then lStartDate = CLng(Left(lPayID, Len("0000")) & "1001")
																If lEndDate > CLng(Left(lPayID, Len("0000")) & "0331") Then lEndDate = CLng(Left(lPayID, Len("0000")) & "0331")
															ElseIf CInt(Right(lPayID, Len("0000"))) = 1231 Then
																If lStartDate < CLng(Left(lPayID, Len("0000")) & "0401") Then lStartDate = CLng(Left(lPayID, Len("0000")) & "0401")
																If lEndDate > CLng(Left(lPayID, Len("0000")) & "0930") Then lEndDate = CLng(Left(lPayID, Len("0000")) & "0930")
															End If
															If lStartDate <= lEndDate Then
																Select Case CInt(oRecordset.Fields("AbsenceID").Value)
																	Case 29, 82
																		If (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDate), GetDateFromSerialNumber(lEndDate))) + 1) > 3 Then dAmount = dAmount + (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDate), GetDateFromSerialNumber(lEndDate))) + 1) - 3
																	Case 30, 83
																		dAmount = dAmount + (Abs(DateDiff("d", GetDateFromSerialNumber(lStartDate), GetDateFromSerialNumber(lEndDate))) + 1)
																	Case 89
																		dAmount = dAmount + Int((Abs(DateDiff("d", GetDateFromSerialNumber(lStartDate), GetDateFromSerialNumber(lEndDate))) + 1) / 15)
																End Select
															End If
															oRecordset.MoveNext
															'If Err.number <> 0 Then Exit Do
														Loop
														oRecordset.Close
														dAmount = CDbl(asEmployeesQueries(1)) * (adTotal(jIndex)(1) - dAmount)
														If dAmount < 0 Then dAmount = 0
														lErrorNumber = AppendTextToFile(sFilePath & "_Payroll_Special_" & Int(iCounter2 / ROWS_PER_FILE) & ".txt", asEmployeesQueries(0) & ", " & asSpecialConcepts(kIndex) & ", 1, " & FormatNumber(dAmount, 2, True, False, False), sErrorDescription)
														iCounter = iCounter + 1
													End If
												End If
												'If lErrorNumber <> 0 Then Exit For
											Next
										End If
										Call DeleteFile(sFilePath & "_Payroll_29.txt", "")
									End If
								End If
							End If
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo limpiar la tabla temporal."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll -->" & vbNewLine
							End If
						Case 30 '26. Gratificación de fin de año (aguinaldo)

						Case 34 '31. Remuneraciones por suplencias
							sQueryBegin = ""
							If InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Zones"
							If InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Areas"
							If InStr(1, sCondition, "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
							If InStr(1, sCondition, "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
							If InStr(1, sCondition, "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
							If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"

							sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesSpecialJourneys.EmployeeID, EmployeesSpecialJourneys.OriginalEmployeeID, EmployeesSpecialJourneys.SpecialJourneyID, EmployeesSpecialJourneys.StartDate, EmployeesSpecialJourneys.EndDate, EmployeesHistoryList.WorkingHours, EmployeesHistoryList1.WorkingHours As WorkedHours, SpecialJourneyFactor1, SpecialJourneyFactor2 From EmployeesChangesLKP, EmployeesHistoryList, EmployeesChangesLKP As EmployeesChangesLKP1, EmployeesHistoryList As EmployeesHistoryList1, EmployeesSpecialJourneys, Journeys Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesSpecialJourneys.EmployeeID) And (EmployeesChangesLKP1.EmployeeID=EmployeesHistoryList1.EmployeeID) And (EmployeesChangesLKP1.EmployeeDate=EmployeesHistoryList1.EmployeeDate) And (EmployeesChangesLKP1.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP1.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP1.EmployeeID=EmployeesSpecialJourneys.OriginalEmployeeID) And (EmployeesHistoryList1.JourneyID=Journeys.JourneyID) And (EmployeesSpecialJourneys.EmployeeID<800000) And (SpecialJourneyID In (424)) And (AppliedDate=" & lPayID & ") And (EmployeesSpecialJourneys.EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & ")) Order By EmployeesSpecialJourneys.EmployeeID, EmployeesSpecialJourneys.OriginalEmployeeID, EmployeesSpecialJourneys.StartDate", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Select EmployeesSpecialJourneys.EmployeeID, EmployeesSpecialJourneys.OriginalEmployeeID, EmployeesSpecialJourneys.SpecialJourneyID, EmployeesSpecialJourneys.StartDate, EmployeesSpecialJourneys.EndDate, EmployeesHistoryList.WorkingHours, EmployeesHistoryList1.WorkingHours As WorkedHours, SpecialJourneyFactor1, SpecialJourneyFactor2 From EmployeesChangesLKP, EmployeesHistoryList, EmployeesChangesLKP As EmployeesChangesLKP1, EmployeesHistoryList As EmployeesHistoryList1, EmployeesSpecialJourneys, Journeys Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesChangesLKP.EmployeeID=EmployeesSpecialJourneys.EmployeeID) And (EmployeesChangesLKP1.EmployeeID=EmployeesHistoryList1.EmployeeID) And (EmployeesChangesLKP1.EmployeeDate=EmployeesHistoryList1.EmployeeDate) And (EmployeesChangesLKP1.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP1.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP1.EmployeeID=EmployeesSpecialJourneys.OriginalEmployeeID) And (EmployeesHistoryList1.JourneyID=Journeys.JourneyID) And (EmployeesSpecialJourneys.EmployeeID<800000) And (SpecialJourneyID In (424)) And (AppliedDate=" & lPayID & ") And (EmployeesSpecialJourneys.EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & ")) Order By EmployeesSpecialJourneys.EmployeeID, EmployeesSpecialJourneys.OriginalEmployeeID, EmployeesSpecialJourneys.StartDate -->" & vbNewLine
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									asFileContents = ""
									Do While Not oRecordset.EOF
										dAmount = CDbl(oRecordset.Fields("WorkedHours").Value) * (Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1)
										asFileContents = asFileContents & CStr(oRecordset.Fields("EmployeeID").Value) & "," & CStr(oRecordset.Fields("OriginalEmployeeID").Value) & "," & CStr(oRecordset.Fields("StartDate").Value) & "," & dAmount & "," & CStr(oRecordset.Fields("WorkingHours").Value) & "," & (Abs(DateDiff("d", GetDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)), GetDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)))) + 1) & "," & CDbl(oRecordset.Fields("SpecialJourneyFactor1").Value) & "," & CDbl(oRecordset.Fields("SpecialJourneyFactor2").Value) & ";"
										oRecordset.MoveNext
										'If Err.number <> 0 Then Exit Do
									Loop
									oRecordset.Close
									If Len(asFileContents) > 0 Then asFileContents = Left(asFileContents, (Len(asFileContents) - Len(";")))
									asFileContents = Split(asFileContents, ";")
									dAmount = 0
									lCurrentID = -2
									iCounter = 0
									lCurrentID2 = -2
									For iIndex = 0 To UBound(asFileContents)
										asFileContents(iIndex) = Split(asFileContents(iIndex), ",")
										If lCurrentID <> CLng(asFileContents(iIndex)(0)) Then
											If lCurrentID <> -2 Then
												dTemp = dTemp * (iCounter + (Int(iCounter / CDbl(adTotal(1))) * 2))
												dAmount = dAmount + dTemp
												sErrorDescription = "No se pudieron agregar los montos de la nómina de los empleados."
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & lPayID & ", 1, " & lCurrentID & ", " & asSpecialConcepts(kIndex) & ", 1, " & dAmount & ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
											End If
											dAmount = 0
											lCurrentID = CLng(asFileContents(iIndex)(0))
											iCounter = 0
											lCurrentID2 = -2
										End If
										If lCurrentID2 <> CLng(asFileContents(iIndex)(1)) Then
											If lCurrentID2 <> -2 Then
												dTemp = dTemp * (iCounter + (Int(iCounter * CDbl(adTotal(0))) * 2))
												dAmount = dAmount + dTemp
											End If

											iCounter = 0
											lCurrentID2 = CLng(asFileContents(iIndex)(1))
											dTemp = 0
											adTotal = Split((asFileContents(iIndex)(6) & ";" & asFileContents(iIndex)(7)), ";")

											sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
											If aPayrollComponent(N_ID_PAYROLL) > 29999999 Then
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & Left(asFileContents(iIndex)(2), Len("YYYY")) & " Where (EmployeeID=" & asFileContents(iIndex)(1) & ") And (RecordID=" & GetPayrollEndDate(asFileContents(iIndex)(2)) & ") And (ConceptID In (1,4,38,130)) Group By ConceptID Order By ConceptID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
											ElseIf aPayrollComponent(N_ID_PAYROLL) <> lPayID Then
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & Left(lPayID, Len("YYYY")) & " Where (EmployeeID=" & asFileContents(iIndex)(1) & ") And (RecordID=" & lPayID & ") And (ConceptID In (1,4,38,130)) Group By ConceptID Order By ConceptID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
											Else
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, Sum(ConceptAmount) As TotalAmount From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID=" & asFileContents(iIndex)(1) & ") And (RecordDate=" & lPayID & ") And (ConceptID In (1,4,38,130)) Group By ConceptID Order By ConceptID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
											End If
											If lErrorNumber = 0 Then
												Do While Not oRecordset.EOF
													Select Case CLng(oRecordset.Fields("ConceptID").Value)
														Case 1
															Select Case 3
																Case 1
																	dTemp = dTemp + (CDbl(oRecordset.Fields("TotalAmount").Value) * 1.1)
																Case 2
																	dTemp = dTemp + (CDbl(oRecordset.Fields("TotalAmount").Value) * 1.2)
																Case Else
																	dTemp = dTemp + CDbl(oRecordset.Fields("TotalAmount").Value)
															End Select
														Case 4, 38, 130
															dTemp = dTemp + CDbl(oRecordset.Fields("TotalAmount").Value)
													End Select
													oRecordset.MoveNext
													'If Err.number <> 0 Then Exit Do
												Loop
												oRecordset.Close
												dTemp = (dTemp / 15) * CDbl(adTotal(0))
											End If
										End If
										iCounter = iCounter + CInt(asFileContents(iIndex)(5))
										If bTimeout Then Exit For
									Next
									dTemp = dTemp * (iCounter + (Int(iCounter / CDbl(adTotal(1))) * 2))
									dAmount = dAmount + dTemp
									sErrorDescription = "No se pudieron agregar los montos de la nómina de los empleados."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & lPayID & ", 1, " & lCurrentID & ", " & asSpecialConcepts(kIndex) & ", 1, " & dAmount & ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								End If
							End If
						Case 40, 41, 42, 43 '37. Estímulo de asistencia|38. Estímulo de puntualidad|39. Estímulo de desempeño
							Select Case asSpecialConcepts(kIndex)
								Case 40
									dTaxAmount = 0
								Case 41
									dTaxAmount = 40
								Case 42
									dTaxAmount = 80
								Case 43
									dTaxAmount = 120
							End Select
							If ((asSpecialConcepts(kIndex) <> 43) And (InStr(1, ",0131,0228,0229,0331,0430,0531,0630,0731,0831,0930,1031,1130,1231,", Right(lPayID, Len("0000")), vbBinaryCompare) > 0)) Or ((asSpecialConcepts(kIndex) = 43) And (InStr(1, ",0131,0430,0731,1031,", Right(lPayID, Len("0000")), vbBinaryCompare) > 0)) Then
								sErrorDescription = "No se pudo limpiar la tabla temporal."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll -->" & vbNewLine
								If lErrorNumber = 0 Then
									sQueryBegin = ""
									If InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Zones"
									If InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Areas"
									If InStr(1, sCondition, "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
									If InStr(1, sCondition, "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
									If InStr(1, sCondition, "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
									If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
									sErrorDescription = "No se pudieron obtener los montos de la nómina de los empleados."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, EmployeesHistoryList.EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '100' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Employees " & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (ConceptID In (1,4,5,6,7)) And (Employees.StartDate<" & Left(GetSerialNumberForDate(DateAdd("m", -6, GetDateFromSerialNumber(lPayID))), Len("00000000")) & ")" & sCondition & " Group By EmployeesHistoryList.EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, EmployeesHistoryList.EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '100' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Employees " & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (ConceptID In (1,4,5,6,7)) And (Employees.StartDate<" & Left(GetSerialNumberForDate(DateAdd("m", -6, GetDateFromSerialNumber(lPayID))), Len("00000000")) & ")" & sCondition & " Group By EmployeesHistoryList.EmployeeID -->" & vbNewLine
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudieron obtener los montos de la nómina de los empleados."
										If StrComp(Right(lPayID, Len("0000")), "0131", vbBinaryCompare) = 0 Then
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '2' As RecordID, EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, Sum(Payroll_" & CInt(Left(lPayID, Len("0000"))) - 1 & ".ConceptAmount) As ConceptAmount1, '100' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & CInt(Left(lPayID, Len("0000"))) - 1 & " Where (EmployeeID In (Select EmployeeID From Payroll)) And (RecordID>=" & CInt(Left(lPayID, Len("0000"))) - 1 & "1200) And (RecordID<=" & CInt(Left(lPayID, Len("0000"))) - 1 & "1299) And (ConceptID In (1,4,5,6,7)) Group By EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
											Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '2' As RecordID, EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, Sum(Payroll_" & CInt(Left(lPayID, Len("0000"))) - 1 & ".ConceptAmount) As ConceptAmount1, '100' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & CInt(Left(lPayID, Len("0000"))) - 1 & " Where (EmployeeID In (Select EmployeeID From Payroll)) And (RecordID>=" & CInt(Left(lPayID, Len("0000"))) - 1 & "1200) And (RecordID<=" & CInt(Left(lPayID, Len("0000"))) - 1 & "1299) And (ConceptID In (1,4,5,6,7)) Group By EmployeeID -->" & vbNewLine

											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '2' As RecordID, EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, Sum(Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptAmount) As ConceptAmount1, '100' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID In (Select EmployeeID From Payroll)) And (EmployeeID Not In (Select EmployeeID From Payroll Where (RecordID=2))) And (RecordID>=" & CInt(Left(lPayID, Len("0000"))) - 1 & "1200) And (RecordID<=" & CInt(Left(lPayID, Len("0000"))) - 1 & "1299) And (ConceptID In (1,4,5,6,7)) Group By EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
											Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '2' As RecordID, EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, Sum(Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptAmount) As ConceptAmount1, '100' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID In (Select EmployeeID From Payroll)) And (EmployeeID Not In (Select EmployeeID From Payroll Where (RecordID=2))) And (RecordID>=" & CInt(Left(lPayID, Len("0000"))) - 1 & "1200) And (RecordID<=" & CInt(Left(lPayID, Len("0000"))) - 1 & "1299) And (ConceptID In (1,4,5,6,7)) Group By EmployeeID -->" & vbNewLine
										Else
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '2' As RecordID, EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, Sum(Payroll_" & Left(lPayID, Len("0000")) & ".ConceptAmount) As ConceptAmount1, '100' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & Left(lPayID, Len("0000")) & " Where (EmployeeID In (Select EmployeeID From Payroll)) And (RecordID>=" & Left(CLng(lPayID) - 100, Len("000000")) & "00) And (RecordID<=" & Left(CLng(lPayID) - 100, Len("000000")) & "99) And (ConceptID In (1,4,5,6,7)) Group By EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
											Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '2' As RecordID, EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, Sum(Payroll_" & Left(lPayID, Len("0000")) & ".ConceptAmount) As ConceptAmount1, '100' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & Left(lPayID, Len("0000")) & " Where (EmployeeID In (Select EmployeeID From Payroll)) And (RecordID>=" & Left(CLng(lPayID) - 100, Len("000000")) & "00) And (RecordID<=" & Left(CLng(lPayID) - 100, Len("000000")) & "99) And (ConceptID In (1,4,5,6,7)) Group By EmployeeID -->" & vbNewLine

											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '2' As RecordID, EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, Sum(Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptAmount) As ConceptAmount1, '100' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID In (Select EmployeeID From Payroll)) And (EmployeeID Not In (Select EmployeeID From Payroll Where (RecordID=2))) And (RecordID>=" & Left(CLng(lPayID) - 100, Len("000000")) & "00) And (RecordID<=" & Left(CLng(lPayID) - 100, Len("000000")) & "99) And (ConceptID In (1,4,5,6,7)) Group By EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
											Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '2' As RecordID, EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, Sum(Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptAmount) As ConceptAmount1, '100' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID In (Select EmployeeID From Payroll)) And (EmployeeID Not In (Select EmployeeID From Payroll Where (RecordID=2))) And (RecordID>=" & Left(CLng(lPayID) - 100, Len("000000")) & "00) And (RecordID<=" & Left(CLng(lPayID) - 100, Len("000000")) & "99) And (ConceptID In (1,4,5,6,7)) Group By EmployeeID -->" & vbNewLine
										End If
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudo limpiar la tabla temporal."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChangesLKP Set Concepts40=" & (dTaxAmount + 1) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (" & lPayID & ",-" & lPayID & ")) And (Concepts40<=" & dTaxAmount & ") And (EmployeeID In (Select EmployeeID From Payroll Where (RecordID=1))) And (EmployeeID Not In (Select EmployeeID From Payroll Where (RecordID=2)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

											sErrorDescription = "No se pudo limpiar la tabla temporal."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (RecordID=1)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
											Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (RecordID=1) -->" & vbNewLine
											If lErrorNumber = 0 Then
												sErrorDescription = "No se pudo limpiar la tabla temporal."
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChangesLKP Set Concepts40=" & (dTaxAmount + 2) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (" & lPayID & ",-" & lPayID & ")) And (Concepts40<=" & dTaxAmount & ") And (EmployeeID In (Select EmployeeID From Payroll Where (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesHistoryList, EmployeesChangesLKP, StatusEmployees, Reasons Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And ((EmployeesHistoryList.Active=0) Or (StatusEmployees.Active=0) Or (Reasons.ActiveEmployeeID=2)) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ")))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

												sErrorDescription = "No se pudo limpiar la tabla temporal." 'Inactivos
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesHistoryList, EmployeesChangesLKP, StatusEmployees, Reasons Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And ((EmployeesHistoryList.Active=0) Or (StatusEmployees.Active=0) Or (Reasons.ActiveEmployeeID=2)) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ")))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
												Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesHistoryList, EmployeesChangesLKP, StatusEmployees, Reasons Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And ((EmployeesHistoryList.Active=0) Or (StatusEmployees.Active=0) Or (Reasons.ActiveEmployeeID=2)) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & "))) -->" & vbNewLine
												If lErrorNumber = 0 Then
													sErrorDescription = "No se pudo limpiar la tabla temporal."
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChangesLKP Set Concepts40=" & (dTaxAmount + 3) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (" & lPayID & ",-" & lPayID & ")) And (Concepts40<=" & dTaxAmount & ") And (EmployeeID In (Select EmployeeID From Payroll Where (EmployeeID Not In (Select EmployeeID From Payroll_Antiquities Where ((Years2>0) Or (Months2>8) Or ((Months2=8) And (Days2>0))) And (bIsCurrent=1)))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

													sErrorDescription = "No se pudo limpiar la tabla temporal." 'Antigüedad
													lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID Not In (Select EmployeeID From Payroll_Antiquities Where ((Years2>0) Or (Months2>8) Or ((Months2=8) And (Days2>0))) And (bIsCurrent=1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
													Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID Not In (Select EmployeeID From Payroll_Antiquities Where ((Years2>0) Or (Months2>8) Or ((Months2=8) And (Days2>0))) And (bIsCurrent=1))) -->" & vbNewLine
												End If
												If lErrorNumber = 0 Then
													Select Case Right(lPayID, Len("0000"))
														Case "0131"
															lStartDate = (CInt(Left(lPayID, Len("0000"))) - 1) & "1201"
															lEndDate = (CInt(Left(lPayID, Len("0000"))) - 1) & "1231"
														Case "0228", "0229"
															lStartDate = Left(lPayID, Len("0000")) & "0101"
															lEndDate = Left(lPayID, Len("0000")) & "0131"
														Case "0331"
															lStartDate = Left(lPayID, Len("0000")) & "0201"
															lEndDate = Left(lPayID, Len("0000")) & "0228"
														Case "0430"
															lStartDate = Left(lPayID, Len("0000")) & "0301"
															lEndDate = Left(lPayID, Len("0000")) & "0331"
														Case "0531"
															lStartDate = Left(lPayID, Len("0000")) & "0401"
															lEndDate = Left(lPayID, Len("0000")) & "0430"
														Case "0630"
															lStartDate = Left(lPayID, Len("0000")) & "0501"
															lEndDate = Left(lPayID, Len("0000")) & "0531"
														Case "0731"
															lStartDate = Left(lPayID, Len("0000")) & "0601"
															lEndDate = Left(lPayID, Len("0000")) & "0630"
														Case "0831"
															lStartDate = Left(lPayID, Len("0000")) & "0701"
															lEndDate = Left(lPayID, Len("0000")) & "0731"
														Case "0930"
															lStartDate = Left(lPayID, Len("0000")) & "0801"
															lEndDate = Left(lPayID, Len("0000")) & "0831"
														Case "1031"
															lStartDate = Left(lPayID, Len("0000")) & "0901"
															lEndDate = Left(lPayID, Len("0000")) & "0930"
														Case "1130"
															lStartDate = Left(lPayID, Len("0000")) & "1001"
															lEndDate = Left(lPayID, Len("0000")) & "1031"
														Case "1231"
															lStartDate = Left(lPayID, Len("0000")) & "1101"
															lEndDate = Left(lPayID, Len("0000")) & "1130"
													End Select
													lStartDate = CLng(lStartDate)
													lEndDate = CLng(lEndDate)
													If lErrorNumber = 0 Then
														sErrorDescription = "No se pudo limpiar la tabla temporal."
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChangesLKP Set Concepts40=" & (dTaxAmount + 4) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (" & lPayID & ",-" & lPayID & ")) And (Concepts40<=" & dTaxAmount & ") And (EmployeeID In (Select EmployeeID From Payroll Where (EmployeeID In (Select EmployeeID From Payments Where (PaymentDate>=" & lStartDate & ") And (PaymentDate<=" & lEndDate & ") And (StatusID Not In (-2,-1,1,2,3))))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

														sErrorDescription = "No se pudieron obtener los cheques cancelados de los empleados." 'Cheques cancelados
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (Select EmployeeID From Payments Where (PaymentDate>=" & lStartDate & ") And (PaymentDate<=" & lEndDate & ") And (StatusID Not In (-2,-1,1,2,3))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
														Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (Select EmployeeID From Payments Where (PaymentDate>=" & lStartDate & ") And (PaymentDate<=" & lEndDate & ") And (StatusID Not In (-2,-1,1,2,3)))) -->" & vbNewLine
													End If
													If lErrorNumber = 0 Then
														sErrorDescription = "No se pudo limpiar la tabla temporal."
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChangesLKP Set Concepts40=" & (dTaxAmount + 5) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (" & lPayID & ",-" & lPayID & ")) And (Concepts40<=" & dTaxAmount & ") And (EmployeeID In (Select EmployeeID From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesHistoryList Where (PositionTypeID<>1) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeeDate<=" & lEndDate & ") And (EndDate>=" & lStartDate & ")))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

														sErrorDescription = "No se pudo limpiar la tabla temporal." 'Empleados que no son de base
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesHistoryList Where (PositionTypeID<>1) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeeDate<=" & lEndDate & ") And (EndDate>=" & lStartDate & ")))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
														Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesHistoryList Where (PositionTypeID<>1) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeeDate<=" & lEndDate & ") And (EndDate>=" & lStartDate & "))) -->" & vbNewLine
													End If
													If lErrorNumber = 0 Then
														sErrorDescription = "No se pudo limpiar la tabla temporal."
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChangesLKP Set Concepts40=" & (dTaxAmount + 6) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (" & lPayID & ",-" & lPayID & ")) And (Concepts40<=" & dTaxAmount & ") And (EmployeeID In (Select EmployeeID From Payroll Where (EmployeeID Not In (Select EmployeeID From EmployeesAbsencesLKP Where (AbsenceID In (50,51,52,53)) And (((EndDate>=" & lStartDate & ") And (EndDate<=" & lEndDate & ")) Or ((OcurredDate>=" & lStartDate & ") And (OcurredDate<=" & lEndDate & ")) Or ((OcurredDate>=" & lStartDate & ") And (EndDate<=" & lEndDate & ")) Or ((OcurredDate<=" & lEndDate & ") And (EndDate>=" & lStartDate & ")))))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

														sErrorDescription = "No se pudo limpiar la tabla temporal." 'Empleados que no tengan registrado el 0901,0902,0903,0904
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID Not In (Select EmployeeID From EmployeesAbsencesLKP Where (AbsenceID In (50,51,52,53)) And (((EndDate>=" & lStartDate & ") And (EndDate<=" & lEndDate & ")) Or ((OcurredDate>=" & lStartDate & ") And (OcurredDate<=" & lEndDate & ")) Or ((OcurredDate>=" & lStartDate & ") And (EndDate<=" & lEndDate & ")) Or ((OcurredDate<=" & lEndDate & ") And (EndDate>=" & lStartDate & ")))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
														Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID Not In (Select EmployeeID From EmployeesAbsencesLKP Where (AbsenceID In (50,51,52,53)) And (((EndDate>=" & lStartDate & ") And (EndDate<=" & lEndDate & ")) Or ((OcurredDate>=" & lStartDate & ") And (OcurredDate<=" & lEndDate & ")) Or ((OcurredDate>=" & lStartDate & ") And (EndDate<=" & lEndDate & ")) Or ((OcurredDate<=" & lEndDate & ") And (EndDate>=" & lStartDate & "))))) -->" & vbNewLine
													End If
													If lErrorNumber = 0 Then
														sErrorDescription = "No se pudo limpiar la tabla temporal."
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChangesLKP Set Concepts40=" & (dTaxAmount + 7) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (" & lPayID & ",-" & lPayID & ")) And (Concepts40<=" & dTaxAmount & ") And (EmployeeID In (Select EmployeeID From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesHistoryList Where (EmployeesHistoryList.StatusID In (1)) And (((EmployeesHistoryList.EndDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EmployeeDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate<=" & lEndDate & ") And (EmployeesHistoryList.EndDate>=" & lStartDate & ")))))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

														sErrorDescription = "No se pudo limpiar la tabla temporal." 'Interinato
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesHistoryList Where (EmployeesHistoryList.StatusID In (1)) And (((EmployeesHistoryList.EndDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EmployeeDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate<=" & lEndDate & ") And (EmployeesHistoryList.EndDate>=" & lStartDate & ")))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
														Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesHistoryList Where (EmployeesHistoryList.StatusID In (1)) And (((EmployeesHistoryList.EndDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EmployeeDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate<=" & lEndDate & ") And (EmployeesHistoryList.EndDate>=" & lStartDate & "))))) -->" & vbNewLine
													End If
													If lErrorNumber = 0 Then
														sErrorDescription = "No se pudo limpiar la tabla temporal."
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChangesLKP Set Concepts40=" & (dTaxAmount + 8) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (" & lPayID & ",-" & lPayID & ")) And (Concepts40<=" & dTaxAmount & ") And (EmployeeID In (Select EmployeeID From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesHistoryList Where (EmployeesHistoryList.StatusID In (58,78)) And (((EmployeesHistoryList.EndDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EmployeeDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate<=" & lEndDate & ") And (EmployeesHistoryList.EndDate>=" & lStartDate & ")))))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

														sErrorDescription = "No se pudo limpiar la tabla temporal." 'Licencias prejubilatorias y becas
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesHistoryList Where (EmployeesHistoryList.StatusID In (58,78)) And (((EmployeesHistoryList.EndDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EmployeeDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate<=" & lEndDate & ") And (EmployeesHistoryList.EndDate>=" & lStartDate & ")))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
														Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesHistoryList Where (EmployeesHistoryList.StatusID In (58,78)) And (((EmployeesHistoryList.EndDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EmployeeDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate>=" & lStartDate & ") And (EmployeesHistoryList.EndDate<=" & lEndDate & ")) Or ((EmployeesHistoryList.EmployeeDate<=" & lEndDate & ") And (EmployeesHistoryList.EndDate>=" & lStartDate & "))))) -->" & vbNewLine
													End If
													If lErrorNumber = 0 Then
														sErrorDescription = "No se pudo limpiar la tabla temporal." 'Faltas y retardos
														Select Case CInt(asSpecialConcepts(kIndex))
															Case 40
																sErrorDescription = "No se pudo limpiar la tabla temporal."
																lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChangesLKP Set Concepts40=" & (dTaxAmount + 9) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (" & lPayID & ",-" & lPayID & ")) And (Concepts40<=" & dTaxAmount & ") And (EmployeeID In (Select EmployeeID From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP, Absences Where (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And ((Absences.AbsenceID In (10,15,16,23,24,27,32,33,41,42,43,44,45,46,47,48,49,54,55,56)) Or (Absences.AbsenceTypeID2=0)) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (EmployeesAbsencesLKP.Active=1) And (((EmployeesAbsencesLKP.EndDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ") And (EmployeesAbsencesLKP.EndDate>=" & lStartDate & ")))))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

																lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP, Absences Where (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And ((Absences.AbsenceID In (10,15,16,23,24,27,32,33,41,42,43,44,45,46,47,48,49,54,55,56)) Or (Absences.AbsenceTypeID2=0)) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (EmployeesAbsencesLKP.Active=1) And (((EmployeesAbsencesLKP.EndDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ") And (EmployeesAbsencesLKP.EndDate>=" & lStartDate & ")))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
																Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP, Absences Where (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And ((Absences.AbsenceID In (10,15,16,23,24,27,32,33,41,42,43,44,45,46,47,48,49,54,55,56)) Or (Absences.AbsenceTypeID2=0)) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (EmployeesAbsencesLKP.Active=1) And (((EmployeesAbsencesLKP.EndDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ") And (EmployeesAbsencesLKP.EndDate>=" & lStartDate & "))))) -->" & vbNewLine

																sErrorDescription = "No se pudo limpiar la tabla temporal."
																lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChangesLKP Set Concepts40=" & (dTaxAmount + 10) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (" & lPayID & ",-" & lPayID & ")) And (Concepts40<=" & dTaxAmount & ") And (EmployeeID In (Select EmployeeID From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP Where (AbsenceID In (79)) And (AppliedDate>=" & Left(lPayID, Len("000000")) & "00) And (AppliedDate<=" & Left(lPayID, Len("000000")) & "99) And (JustificationID=-1) And (Removed=0) And (Active=1)))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

																lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP Where (AbsenceID In (79)) And (AppliedDate>=" & Left(lPayID, Len("000000")) & "00) And (AppliedDate<=" & Left(lPayID, Len("000000")) & "99) And (JustificationID=-1) And (Removed=0) And (Active=1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
																Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP Where (AbsenceID In (79)) And (AppliedDate In (0," & lPayID & ") And (JustificationID=-1) And (Removed=0) And (Active=1)))) -->" & vbNewLine
															Case 41
																sErrorDescription = "No se pudo limpiar la tabla temporal."
																lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChangesLKP Set Concepts40=" & (dTaxAmount + 9) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (" & lPayID & ",-" & lPayID & ")) And (Concepts40<=" & dTaxAmount & ") And (EmployeeID In (Select EmployeeID From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP, Absences Where (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And ((Absences.AbsenceID In (1,2,3,8,9,10,15,16,18,19,21,23,24,27,32,33,41,42,43,44,45,46,47,48,49,54,55,56)) Or (Absences.AbsenceTypeID2=0)) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (EmployeesAbsencesLKP.Active=1) And (((EmployeesAbsencesLKP.EndDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ") And (EmployeesAbsencesLKP.EndDate>=" & lStartDate & ")))))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

																lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP, Absences Where (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And ((Absences.AbsenceID In (1,2,3,8,9,10,15,16,18,19,21,23,24,27,32,33,41,42,43,44,45,46,47,48,49,54,55,56)) Or (Absences.AbsenceTypeID2=0)) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (EmployeesAbsencesLKP.Active=1) And (((EmployeesAbsencesLKP.EndDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ") And (EmployeesAbsencesLKP.EndDate>=" & lStartDate & ")))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
																Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP, Absences Where (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And ((Absences.AbsenceID In (1,2,3,8,9,10,15,16,18,19,21,23,24,27,32,33,41,42,43,44,45,46,47,48,49,54,55,56)) Or (Absences.AbsenceTypeID2=0)) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (EmployeesAbsencesLKP.Active=1) And (((EmployeesAbsencesLKP.EndDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ") And (EmployeesAbsencesLKP.EndDate>=" & lStartDate & "))))) -->" & vbNewLine

																sErrorDescription = "No se pudo limpiar la tabla temporal."
																lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChangesLKP Set Concepts40=" & (dTaxAmount + 10) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (" & lPayID & ",-" & lPayID & ")) And (Concepts40<=" & dTaxAmount & ") And (EmployeeID In (Select EmployeeID From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP Where (AbsenceID In (79)) And (AppliedDate>=" & Left(lPayID, Len("000000")) & "00) And (AppliedDate<=" & Left(lPayID, Len("000000")) & "99) And (JustificationID=-1) And (Removed=0) And (Active=1)))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

																lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP Where (AbsenceID In (79)) And (AppliedDate>=" & Left(lPayID, Len("000000")) & "00) And (AppliedDate<=" & Left(lPayID, Len("000000")) & "99) And (JustificationID=-1) And (Removed=0) And (Active=1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
																Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP Where (AbsenceID In (79)) And (AppliedDate In (0," & lPayID & ") And (JustificationID=-1) And (Removed=0) And (Active=1)))) -->" & vbNewLine
															Case 42, 43
																sErrorDescription = "No se pudo limpiar la tabla temporal."
																lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChangesLKP Set Concepts40=" & (dTaxAmount + 9) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (" & lPayID & ",-" & lPayID & ")) And (Concepts40<=" & dTaxAmount & ") And (EmployeeID In (Select EmployeeID From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP, Absences Where (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And ((Absences.AbsenceID In (1,2,3,8,9,10,15,16,18,19,21,23,24,27,32,33,41,42,43,44,45,46,47,48,49,54,55,56)) Or (Absences.AbsenceTypeID2=0)) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (EmployeesAbsencesLKP.Active=1) And (((EmployeesAbsencesLKP.EndDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ") And (EmployeesAbsencesLKP.EndDate>=" & lStartDate & ")))))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

																lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP, Absences Where (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And ((Absences.AbsenceID In (1,2,3,8,9,10,15,16,18,19,21,23,24,27,32,33,41,42,43,44,45,46,47,48,49,54,55,56)) Or (Absences.AbsenceTypeID2=0)) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (EmployeesAbsencesLKP.Active=1) And (((EmployeesAbsencesLKP.EndDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ") And (EmployeesAbsencesLKP.EndDate>=" & lStartDate & ")))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
																Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP, Absences Where (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And ((Absences.AbsenceID In (1,2,3,8,9,10,15,16,18,19,21,23,24,27,32,33,41,42,43,44,45,46,47,48,49,54,55,56)) Or (Absences.AbsenceTypeID2=0)) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (EmployeesAbsencesLKP.Active=1) And (((EmployeesAbsencesLKP.EndDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate>=" & lStartDate & ") And (EmployeesAbsencesLKP.EndDate<=" & lEndDate & ")) Or ((EmployeesAbsencesLKP.OcurredDate<=" & lEndDate & ") And (EmployeesAbsencesLKP.EndDate>=" & lStartDate & "))))) -->" & vbNewLine

																sErrorDescription = "No se pudo limpiar la tabla temporal."
																lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChangesLKP Set Concepts40=" & (dTaxAmount + 10) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (" & lPayID & ",-" & lPayID & ")) And (Concepts40<=" & dTaxAmount & ") And (EmployeeID In (Select EmployeeID From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP Where (AbsenceID In (40,79)) And (AppliedDate>=" & Left(lPayID, Len("000000")) & "00) And (AppliedDate<=" & Left(lPayID, Len("000000")) & "99) And (JustificationID=-1) And (Removed=0) And (Active=1)))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

																lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP Where (AbsenceID In (40,79)) And (AppliedDate>=" & Left(lPayID, Len("000000")) & "00) And (AppliedDate<=" & Left(lPayID, Len("000000")) & "99) And (JustificationID=-1) And (Removed=0) And (Active=1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
																Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesAbsencesLKP Where (AbsenceID In (40,79)) And (AppliedDate In (0," & lPayID & ") And (JustificationID=-1) And (Removed=0) And (Active=1)))) -->" & vbNewLine
														End Select
														If lErrorNumber = 0 Then
															sCurrentID = "-2"
															sErrorDescription = "No se pudieron obtener las incidencias de los empleados."
															lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesAbsencesLKP.EmployeeID, AbsenceID, OcurredDate, EndDate From EmployeesAbsencesLKP, Payroll Where (EmployeesAbsencesLKP.EmployeeID=Payroll.EmployeeID) And (OcurredDate<=" & lEndDate & ") And (EndDate>=" & lStartDate & ") And (AbsenceID In (29,30,31,34,82,83,84,87)) Order By EmployeesAbsencesLKP.EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
															Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Select EmployeesAbsencesLKP.EmployeeID, AbsenceID, OcurredDate, EndDate From EmployeesAbsencesLKP, Payroll Where (EmployeesAbsencesLKP.EmployeeID=Payroll.EmployeeID) And (OcurredDate<=" & lEndDate & ") And (EndDate>=" & lStartDate & ") And (AbsenceID In (29,30,31,34,82,83,84,87)) Order By EmployeesAbsencesLKP.EmployeeID -->" & vbNewLine
															If lErrorNumber = 0 Then
																dAmount = 0
																dTemp = 0
																lCurrentID = -2
																Do While Not oRecordset.EOF
																	If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
																		If lCurrentID <> -2 Then
																			If (dAmount > 3) Or (dTemp > 3) Then lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into EmployeesList Values (" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, "Error al insertar empleado en exclusió de estimulo")
																		End If
																		dAmount = 0
																		dTemp = 0
																		lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
																	End If
																	lTempStartDate = CLng(oRecordset.Fields("OcurredDate").Value)
																	lTempEndDate = CLng(oRecordset.Fields("EndDate").Value)
																	If lTempStartDate < lStartDate Then lTempStartDate = lStartDate
																	If lTempEndDate > lEndDate Then lTempEndDate = lEndDate
																	Select Case CLng(oRecordset.Fields("AbsenceID").Value)
																		Case 29, 30, 82, 83 '0840, 0841
																			If lTempStartDate <= lTempEndDate Then
																				dAmount = dAmount + (Abs(DateDiff("d", GetDateFromSerialNumber(lTempStartDate), GetDateFromSerialNumber(lTempEndDate))) + 1)
																			End If
																		Case 31, 34, 84, 87 '0847, 0855
																			If lTempStartDate <= lTempEndDate Then
																				dTemp = dTemp + (Abs(DateDiff("d", GetDateFromSerialNumber(lTempStartDate), GetDateFromSerialNumber(lTempEndDate))) + 1)
																			End If
																	End Select
																	oRecordset.MoveNext
																	'If Err.number <> 0 Then Exit Do
																Loop
																oRecordset.Close
																If (dAmount > 3) Or (dTemp > 3) Then lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into EmployeesList Values (" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, "Error al insertar empleado en exclusió de estimulo")
															End If
															sErrorDescription = "No se pudo limpiar la tabla temporal."
															'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChangesLKP Set Concepts40=" & (dTaxAmount + 11) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate=" & lPayID & ") And (Concepts40<=" & dTaxAmount & ") And (EmployeeID In (" & sCurrentID & "))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
															 lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChangesLKP Set Concepts40=" & (dTaxAmount + 11) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate=" & lPayID & ") And (Concepts40<=" & dTaxAmount & ") And (EmployeeID In (Select EmployeeID From EmployeesList))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

															sErrorDescription = "No se pudo limpiar la tabla temporal." 'Faltas y retardos
															'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (" & sCurrentID & "))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
															 lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesList))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
															Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (Select EmployeeID From EmployeesList)) -->" & vbNewLine
															If lErrorNumber = 0 Then
																Select Case CInt(asSpecialConcepts(kIndex))
																	Case 40, 41
																		sErrorDescription = "No se pudieron obtener los montos de la nómina de los empleados."
																		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, Payroll.EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, (ConceptAmount*1.33/30) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll, Employees Where (Payroll.EmployeeID=Employees.EmployeeID) And (Employees.Antiquity3ID<=2)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
																		Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, Payroll.EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, (ConceptAmount*1.33/30) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll, Employees Where (Payroll.EmployeeID=Employees.EmployeeID) And (Antiquity3ID<=2) -->" & vbNewLine

																		sErrorDescription = "No se pudieron obtener los montos de la nómina de los empleados."
																		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, Payroll.EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, (ConceptAmount*1.5/30) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll, Employees Where (Payroll.EmployeeID=Employees.EmployeeID) And (Employees.Antiquity3ID>=3) And (Employees.Antiquity3ID<=5)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
																		Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, Payroll.EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, (ConceptAmount*1.5/30) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll, Employees Where (Payroll.EmployeeID=Employees.EmployeeID) And (Employees.Antiquity3ID>=3) And (Employees.Antiquity3ID<=5) -->" & vbNewLine

																		sErrorDescription = "No se pudieron obtener los montos de la nómina de los empleados."
																		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, Payroll.EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, (ConceptAmount*1.66/30) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll, Employees Where (Payroll.EmployeeID=Employees.EmployeeID) And (Employees.Antiquity3ID>=6)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
																		Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, Payroll.EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, (ConceptAmount*1.66/30) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll, Employees Where (Payroll.EmployeeID=Employees.EmployeeID) And (Employees.Antiquity3ID>=6) -->" & vbNewLine
																	Case 42
																		sErrorDescription = "No se pudieron obtener los montos de la nómina de los empleados."
																		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, Payroll.EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, (ConceptAmount*2/30) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll, Employees Where (Payroll.EmployeeID=Employees.EmployeeID)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
																		Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, Payroll.EmployeeID, '" & asSpecialConcepts(kIndex) & "' As ConceptID, '1' As PayrollTypeID, (ConceptAmount*2/30) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll, Employees Where (Payroll.EmployeeID=Employees.EmployeeID) -->" & vbNewLine
																	Case 43
																		Select Case Right(lPayID, Len("0000"))
																			Case "0131"
																				lStartDate = (CLng(Left(lPayID, Len("0000"))) - 1) & "11"
																			Case "0430"
																				lStartDate = CLng(Left(lPayID, Len("0000"))) & "02"
																			Case "0731"
																				lStartDate = CLng(Left(lPayID, Len("0000"))) & "05"
																			Case "1031"
																				lStartDate = CLng(Left(lPayID, Len("0000"))) & "08"
																		End Select
																		sErrorDescription = "No se pudieron obtener los montos de la nómina de los empleados."
																		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '0' As RecordID, EmployeeID, '43' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '1' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & Left(lStartDate, Len("0000")) & " Where (EmployeeID In (Select EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=40) And (ConceptAmount>0))) And (EmployeeID In (Select EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=41) And (ConceptAmount>0))) And (EmployeeID In (Select EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=42) And (ConceptAmount>0))) And (EmployeeID In (Select EmployeeID From Payroll_" & Left(lStartDate, Len("0000")) & " Where (ConceptID=40) And (ConceptAmount>0) And (RecordID>=" & lStartDate & "00) And (RecordID<=" & lStartDate & "99))) And (EmployeeID In (Select EmployeeID From Payroll_" & Left(lStartDate, Len("0000")) & " Where (ConceptID=41) And (ConceptAmount>0) And (RecordID>=" & lStartDate & "00) And (RecordID<=" & lStartDate & "99))) And (EmployeeID In (Select EmployeeID From Payroll_" & Left(lStartDate, Len("0000")) & " Where (ConceptID=42) And (ConceptAmount>0) And (RecordID>=" & lStartDate & "00) And (RecordID<=" & lStartDate & "99))) And (RecordID>=" & lStartDate & "00) And (RecordID<=" & lStartDate & "99) And (ConceptID In (1,4,5,6,7)) Group By EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
																		Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '0' As RecordID, EmployeeID, '43' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '1' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & Left(lStartDate, Len("0000")) & " Where (EmployeeID In (Select EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=40) And (ConceptAmount>0))) And (EmployeeID In (Select EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=41) And (ConceptAmount>0))) And (EmployeeID In (Select EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=42) And (ConceptAmount>0))) And (EmployeeID In (Select EmployeeID From Payroll_" & Left(lStartDate, Len("0000")) & " Where (ConceptID=40) And (ConceptAmount>0) And (RecordID>=" & lStartDate & "00) And (RecordID<=" & lStartDate & "99))) And (EmployeeID In (Select EmployeeID From Payroll_" & Left(lStartDate, Len("0000")) & " Where (ConceptID=41) And (ConceptAmount>0) And (RecordID>=" & lStartDate & "00) And (RecordID<=" & lStartDate & "99))) And (EmployeeID In (Select EmployeeID From Payroll_" & Left(lStartDate, Len("0000")) & " Where (ConceptID=42) And (ConceptAmount>0) And (RecordID>=" & lStartDate & "00) And (RecordID<=" & lStartDate & "99))) And (RecordID>=" & lStartDate & "00) And (RecordID<=" & lStartDate & "99) And (ConceptID In (1,4,5,6,7)) Group By EmployeeID -->" & vbNewLine

																		Select Case Right(lPayID, Len("0000"))
																			Case "0131"
																				lStartDate = (CLng(Left(lPayID, Len("0000"))) - 1) & "12"
																			Case "0430"
																				lStartDate = CLng(Left(lPayID, Len("0000"))) & "03"
																			Case "0731"
																				lStartDate = CLng(Left(lPayID, Len("0000"))) & "06"
																			Case "1031"
																				lStartDate = CLng(Left(lPayID, Len("0000"))) & "09"
																		End Select
																		sErrorDescription = "No se pudieron obtener los montos de la nómina de los empleados."
																		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, EmployeeID, '43' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '10' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & Left(lStartDate, Len("0000")) & " Where (EmployeeID In (Select EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=40) And (ConceptAmount>0))) And (EmployeeID In (Select EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=41) And (ConceptAmount>0))) And (EmployeeID In (Select EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=42) And (ConceptAmount>0)))And (EmployeeID In (Select EmployeeID From Payroll_" & Left(lStartDate, Len("0000")) & " Where (ConceptID=40) And (ConceptAmount>0) And (RecordID>=" & lStartDate & "00) And (RecordID<=" & lStartDate & "99))) And (EmployeeID In (Select EmployeeID From Payroll_" & Left(lStartDate, Len("0000")) & " Where (ConceptID=41) And (ConceptAmount>0) And (RecordID>=" & lStartDate & "00) And (RecordID<=" & lStartDate & "99))) And (EmployeeID In (Select EmployeeID From Payroll_" & Left(lStartDate, Len("0000")) & " Where (ConceptID=42) And (ConceptAmount>0) And (RecordID>=" & lStartDate & "00) And (RecordID<=" & lStartDate & "99))) And (RecordID>=" & lStartDate & "00) And (RecordID<=" & lStartDate & "99) And (ConceptID In (1,4,5,6,7)) Group By EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
																		Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, EmployeeID, '43' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '10' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & Left(lStartDate, Len("0000")) & " Where (EmployeeID In (Select EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=40) And (ConceptAmount>0))) And (EmployeeID In (Select EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=41) And (ConceptAmount>0))) And (EmployeeID In (Select EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=42) And (ConceptAmount>0)))And (EmployeeID In (Select EmployeeID From Payroll_" & Left(lStartDate, Len("0000")) & " Where (ConceptID=40) And (ConceptAmount>0) And (RecordID>=" & lStartDate & "00) And (RecordID<=" & lStartDate & "99))) And (EmployeeID In (Select EmployeeID From Payroll_" & Left(lStartDate, Len("0000")) & " Where (ConceptID=41) And (ConceptAmount>0) And (RecordID>=" & lStartDate & "00) And (RecordID<=" & lStartDate & "99))) And (EmployeeID In (Select EmployeeID From Payroll_" & Left(lStartDate, Len("0000")) & " Where (ConceptID=42) And (ConceptAmount>0) And (RecordID>=" & lStartDate & "00) And (RecordID<=" & lStartDate & "99))) And (RecordID>=" & lStartDate & "00) And (RecordID<=" & lStartDate & "99) And (ConceptID In (1,4,5,6,7)) Group By EmployeeID -->" & vbNewLine

																		sErrorDescription = "No se pudieron obtener los montos de la nómina de los empleados."
																		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, Payroll.EmployeeID, '43' As ConceptID, '1' As PayrollTypeID, (Sum(ConceptAmount)*3/30) As ConceptAmount1, Sum(ConceptTaxes) As ConceptTaxes1, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll, Employees Where (Payroll.EmployeeID=Employees.EmployeeID) Group By Payroll.EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
																		Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, Payroll.EmployeeID, '43' As ConceptID, '1' As PayrollTypeID, (Sum(ConceptAmount)*3/30) As ConceptAmount1, Sum(ConceptTaxes) As ConceptTaxes1, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll, Employees Where (Payroll.EmployeeID=Employees.EmployeeID) Group By Payroll.EmployeeID -->" & vbNewLine
																End Select
																lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Truncate Table EmployeesList", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
															End If
														End If
													End If
												End If
											End If
										End If
									End If
								End If
								sErrorDescription = "No se pudo limpiar la tabla temporal."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll -->" & vbNewLine
							End If
						Case 44 '41. Premio antigüedad 25 y 30 años (mes de sueldo)
							If Len(sEmployeesFor44) > 0 Then
								sQueryBegin = ""
								If InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Zones"
								If InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Areas"
								If InStr(1, sCondition, "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
								If InStr(1, sCondition, "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
								If InStr(1, sCondition, "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
								If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"

								sEmployeesFor44 = Left(sEmployeesFor44, (Len(sEmployeesFor44) - Len(",")))
								sErrorDescription = "No se pudieron obtener los sueldos mensuales de los empleados."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, EmployeeID, '44' As ConceptID, '1' As PayrollTypeID, (Sum(ConceptAmount) * 2) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID In (1,4,5,7,8)) And (EmployeeID In (" & sEmployeesFor44 & ")) Group by EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudieron eliminar los conceptos de pagos de la nómina."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & lPayID & ") And (ConceptID=44) And (EmployeeID In (Select EmployeeID From Payroll_" & CInt(Left(lPayID, Len("YYYY"))) - 1 & " Where (ConceptID=44)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								End If
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudieron eliminar los conceptos de pagos de la nómina."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & lPayID & ") And (ConceptID=44) And (EmployeeID In (Select EmployeeID From Payroll_" & Left(lPayID, Len("YYYY")) & " Where (ConceptID=44)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								End If
							End If

						Case 50 '49. Premio trabajador del mes
							sQueryBegin = ""
							If InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Zones"
							If InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Areas"
							If InStr(1, sCondition, "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
							If InStr(1, sCondition, "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
							If InStr(1, sCondition, "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
							If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"

							sErrorDescription = "No se pudieron obtener los estímulos de los empleados."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(ConceptAmount) As MaxAmount, EconomicZoneID From ConceptsValues Where (ConceptID=1) And (StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (LevelID=20) Group by EconomicZoneID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									adTotal = Split("0,0,0,0", ",")
									For iIndex = 0 To UBound(adTotal)
										adTotal(iIndex) = 0
									Next
									Do While Not oRecordset.EOF
										adTotal(CInt(oRecordset.Fields("EconomicZoneID").Value)) = CDbl(FormatNumber((CDbl(oRecordset.Fields("MaxAmount").Value) * 2 * 0.30), 2, True, False, False))
										oRecordset.MoveNext
										'If Err.number <> 0 Then Exit Do
									Loop
									oRecordset.Close
									sErrorDescription = "No se pudieron obtener los estímulos de los empleados."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, Areas.EconomicZoneID From EmployeesAbsencesLKP, EmployeesChangesLKP, EmployeesHistoryList, Areas, StatusEmployees, Reasons " & Replace(sQueryBegin, ", Areas", "") & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.EmployeeID=EmployeesAbsencesLKP.EmployeeID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (EmployeesAbsencesLKP.AbsenceID=39) And (AppliedDate In (0," & lPayID & ")) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (EmployeesAbsencesLKP.Active=1) " & sCondition & " Order By EmployeesHistoryList.EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										Do While Not oRecordset.EOF
											lErrorNumber = AppendTextToFile(sFilePath & "_Payroll_Special_" & Int(iCounter / ROWS_PER_FILE) & ".txt", CStr(oRecordset.Fields("EmployeeID").Value) & ", " & asSpecialConcepts(kIndex) & ", 1, " & adTotal(CInt(oRecordset.Fields("EconomicZoneID").Value)), sErrorDescription)
											iCounter = iCounter + 1
											oRecordset.MoveNext
											'If Err.number <> 0 Then Exit Do
										Loop
										oRecordset.Close
									End If
								End If
							End If
						Case 52, 71 '50. Inasistencias | 70. Retardos
							sQueryBegin = ""
							If InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Zones"
							If InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Areas"
							If InStr(1, sCondition, "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
							If InStr(1, sCondition, "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
							If InStr(1, sCondition, "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
							If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
							sErrorDescription = "No se pudieron obtener los retardos de los empleados."
							If CInt(asSpecialConcepts(kIndex)) = 52 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.JourneyID, EmployeesHistoryList.WorkingHours, Absences.AbsenceID, Absences.ConceptsIDs, Shifts.JourneyTypeID, JourneyTypes.JourneyFactor, Shifts.WorkingHours As ShiftsWorkingHours, Count(AbsenceHours) As TotalHours From EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, EmployeesAbsencesLKP, Absences, Journeys, Shifts, JourneyTypes " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.EmployeeID=EmployeesAbsencesLKP.EmployeeID) And (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (Shifts.JourneyTypeID=JourneyTypes.JourneyTypeID) And (Journeys.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Journeys.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Shifts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Shifts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesAbsencesLKP.AbsenceID In (3,10,11,16,18,19,20,24,25,26,28,92,93,94)) And (AppliedDate In (0," & lPayID & ")) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (JourneyTypes.JourneyFactor>0) And (EmployeesAbsencesLKP.Active=1) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.JourneyID, EmployeesHistoryList.WorkingHours, Absences.AbsenceID, Absences.ConceptsIDs, Shifts.JourneyTypeID, JourneyTypes.JourneyFactor, Shifts.WorkingHours Order By EmployeesHistoryList.EmployeeID, Absences.ConceptsIDs", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.JourneyID, EmployeesHistoryList.WorkingHours, Absences.AbsenceID, Absences.ConceptsIDs, Shifts.JourneyTypeID, JourneyTypes.JourneyFactor, Shifts.WorkingHours As ShiftsWorkingHours, Count(AbsenceHours) As TotalHours From EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, EmployeesAbsencesLKP, Absences, Journeys, Shifts, JourneyTypes " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.EmployeeID=EmployeesAbsencesLKP.EmployeeID) And (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (Shifts.JourneyTypeID=JourneyTypes.JourneyTypeID) And (Journeys.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Journeys.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Shifts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Shifts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesAbsencesLKP.AbsenceID In (3,10,11,16,18,19,20,24,25,26,28,92,93,94)) And (AppliedDate In (0," & lPayID & ")) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (JourneyTypes.JourneyFactor>0) And (EmployeesAbsencesLKP.Active=1) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.JourneyID, EmployeesHistoryList.WorkingHours, Absences.AbsenceID, Absences.ConceptsIDs, Shifts.JourneyTypeID, JourneyTypes.JourneyFactor, Shifts.WorkingHours Order By EmployeesHistoryList.EmployeeID, Absences.ConceptsIDs -->" & vbNewLine
							Else
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.JourneyID, EmployeesHistoryList.WorkingHours, Absences.AbsenceID, Absences.ConceptsIDs, Shifts.JourneyTypeID, JourneyTypes.JourneyFactor, Shifts.WorkingHours As ShiftsWorkingHours, Count(AbsenceHours) As TotalHours From EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, EmployeesAbsencesLKP, Absences, Journeys, Shifts, JourneyTypes " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.EmployeeID=EmployeesAbsencesLKP.EmployeeID) And (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (Shifts.JourneyTypeID=JourneyTypes.JourneyTypeID) And (Journeys.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Journeys.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Shifts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Shifts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesAbsencesLKP.AbsenceID In (1,2,4,5,21,23,27)) And (AppliedDate In (0," & lPayID & ")) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (JourneyTypes.JourneyFactor>0) And (EmployeesAbsencesLKP.Active=1) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.JourneyID, EmployeesHistoryList.WorkingHours, Absences.AbsenceID, Absences.ConceptsIDs, Shifts.JourneyTypeID, JourneyTypes.JourneyFactor, Shifts.WorkingHours Order By EmployeesHistoryList.EmployeeID, Absences.ConceptsIDs", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.JourneyID, EmployeesHistoryList.WorkingHours, Absences.AbsenceID, Absences.ConceptsIDs, Shifts.JourneyTypeID, JourneyTypes.JourneyFactor, Shifts.WorkingHours As ShiftsWorkingHours, Count(AbsenceHours) As TotalHours From EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, EmployeesAbsencesLKP, Absences, Journeys, Shifts, JourneyTypes " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.EmployeeID=EmployeesAbsencesLKP.EmployeeID) And (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (Shifts.JourneyTypeID=JourneyTypes.JourneyTypeID) And (Journeys.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Journeys.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Shifts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Shifts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesAbsencesLKP.AbsenceID In (1,2,4,5,21,23,27)) And (AppliedDate In (0," & lPayID & ")) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (JourneyTypes.JourneyFactor>0) And (EmployeesAbsencesLKP.Active=1) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.JourneyID, EmployeesHistoryList.WorkingHours, Absences.AbsenceID, Absences.ConceptsIDs, Shifts.JourneyTypeID, JourneyTypes.JourneyFactor, Shifts.WorkingHours Order By EmployeesHistoryList.EmployeeID, Absences.ConceptsIDs -->" & vbNewLine
							End If
							iCounter2 = 0
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									lCurrentID = -2
									sCurrentID = ""
									adTotal = Split("0,0", ",")
									Do While Not oRecordset.EOF
										If StrComp((lCurrentID & ";" & sCurrentID), (CStr(oRecordset.Fields("EmployeeID").Value) & ";" & CStr(oRecordset.Fields("ConceptsIDs").Value)), vbBinaryCompare) <> 0 Then
											If lCurrentID <> -2 Then
												lErrorNumber = AppendTextToFile(sFilePath & "_Payroll_71.txt", lCurrentID & LIST_SEPARATOR & sCurrentID & LIST_SEPARATOR & adTotal(0) & LIST_SEPARATOR & "1", sErrorDescription)
												iCounter2 = iCounter2 + 1
												If adTotal(1) > 0 Then
													lErrorNumber = AppendTextToFile(sFilePath & "_Payroll_71.txt", lCurrentID & LIST_SEPARATOR & sCurrentID & LIST_SEPARATOR & adTotal(1) & LIST_SEPARATOR & "1.4", sErrorDescription)
													iCounter2 = iCounter2 + 1
												End If
											End If
											lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
											sCurrentID = CStr(oRecordset.Fields("ConceptsIDs").Value)
											adTotal = Split("0,0", ",")
										End If
										Select Case CInt(oRecordset.Fields("AbsenceID").Value)
											Case 1
												Select Case CInt(oRecordset.Fields("JourneyTypeID").Value)
													Case 1
														adTotal(0) = adTotal(0) + (Int(CLng(oRecordset.Fields("TotalHours").Value) / 3) / CDbl(oRecordset.Fields("JourneyFactor").Value) / 4)
													Case 2, 3
														adTotal(0) = adTotal(0) + (Int(CLng(oRecordset.Fields("TotalHours").Value) / 2) / CDbl(oRecordset.Fields("JourneyFactor").Value) / 4)
													Case 4
														adTotal(0) = adTotal(0) + (CLng(oRecordset.Fields("TotalHours").Value) / CDbl(oRecordset.Fields("JourneyFactor").Value) / 4)
												End Select
											Case 2, 23, 27
												adTotal(0) = adTotal(0) + (CLng(oRecordset.Fields("TotalHours").Value) / CDbl(oRecordset.Fields("JourneyFactor").Value) / 4)
											Case 4
												adTotal(0) = adTotal(0) + (CLng(oRecordset.Fields("TotalHours").Value) / CDbl(oRecordset.Fields("JourneyFactor").Value) / 2)
											Case 3, 11, 18, 19, 20, 24, 25, 26, 28, 93, 94
												adTotal(0) = adTotal(0) + (CLng(oRecordset.Fields("TotalHours").Value) / CDbl(oRecordset.Fields("JourneyFactor").Value))
											Case 5
												adTotal(0) = adTotal(0) + (CLng(oRecordset.Fields("TotalHours").Value) / CDbl(oRecordset.Fields("JourneyFactor").Value) / 6)
											Case 10
												If CLng(oRecordset.Fields("JourneyTypeID").Value) <> 1 Then
													adTotal(0) = adTotal(0) + (CLng(oRecordset.Fields("TotalHours").Value) / CDbl(oRecordset.Fields("JourneyFactor").Value))
												Else
													adTotal(1) = adTotal(1) + (CLng(oRecordset.Fields("TotalHours").Value) / CDbl(oRecordset.Fields("JourneyFactor").Value))
												End If
											Case 16
												adTotal(0) = adTotal(0) + (CLng(oRecordset.Fields("TotalHours").Value) / CDbl(oRecordset.Fields("JourneyFactor").Value) / 3)
											Case 21
												adTotal(0) = adTotal(0) + (CLng(oRecordset.Fields("TotalHours").Value) / (CDbl(oRecordset.Fields("JourneyFactor").Value) * CDbl(oRecordset.Fields("ShiftsWorkingHours").Value) * 2))
											Case 92
												adTotal(0) = adTotal(0) + (CLng(oRecordset.Fields("TotalHours").Value) / CDbl(oRecordset.Fields("JourneyFactor").Value) * 2 / 5)
										End Select
										oRecordset.MoveNext
										'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
									Loop
									oRecordset.Close
									lErrorNumber = AppendTextToFile(sFilePath & "_Payroll_71.txt", lCurrentID & LIST_SEPARATOR & sCurrentID & LIST_SEPARATOR & adTotal(0) & LIST_SEPARATOR & "1", sErrorDescription)
									iCounter2 = iCounter2 + 1
									If adTotal(1) > 0 Then
										lErrorNumber = AppendTextToFile(sFilePath & "_Payroll_71.txt", lCurrentID & LIST_SEPARATOR & sCurrentID & LIST_SEPARATOR & adTotal(1) & LIST_SEPARATOR & "1.4", sErrorDescription)
										iCounter2 = iCounter2 + 1
									End If
								End If

								If iCounter2 > 0 Then
									asFileContents = GetFileContents(sFilePath & "_Payroll_71.txt", sErrorDescription)
									If Len(asFileContents) > 0 Then
										asFileContents = Split(asFileContents, vbNewLine)
										dAmount = 0
										lCurrentID = -2
										For iIndex = 0 To UBound(asFileContents)
											If Len(asFileContents(iIndex)) > 0 Then
												asEmployeesQueries = Split(asFileContents(iIndex), LIST_SEPARATOR)
												If lCurrentID <> CLng(asEmployeesQueries(0)) Then
													If lCurrentID <> -2 Then
														dAmount = dAmount * CDbl(asEmployeesQueries(3))
														If (dAmount + dTaxAmount) > (dTemp * 0.3) Then dAmount = (dTemp * 0.3) - dTaxAmount
														lErrorNumber = AppendTextToFile(sFilePath & "_Payroll_Special52.txt", lCurrentID & ", " & asSpecialConcepts(kIndex) & ", 1, " & FormatNumber(dAmount, 2, True, False, False), sErrorDescription)
													End If
													dAmount = 0
													dTaxAmount = 0
													lCurrentID = CLng(asEmployeesQueries(0))
												End If
												sErrorDescription = "No se pudieron obtener los montos registrados en la nómina de los empleados."
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptID, ConceptAmount From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID=" & asEmployeesQueries(0) & ") And (ConceptID In (" & asEmployeesQueries(1) & ",52,71))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
												Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Select ConceptID, ConceptAmount From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID=" & asEmployeesQueries(0) & ") And (ConceptID In (" & asEmployeesQueries(1) & ",52,71)) -->" & vbNewLine
												If lErrorNumber = 0 Then
													If CInt(asSpecialConcepts(kIndex)) = 52 Then
														Do While Not oRecordset.EOF
															If CInt(oRecordset.Fields("ConceptID").Value) = 1 Then
																dTemp = CDbl(oRecordset.Fields("ConceptAmount").Value)
																dAmount = dAmount + (CDbl(oRecordset.Fields("ConceptAmount").Value) * CDbl(asEmployeesQueries(2)) * CDbl(asEmployeesQueries(3)))
															ElseIf CInt(oRecordset.Fields("ConceptID").Value) = 4 Then
																dAmount = dAmount + (CDbl(oRecordset.Fields("ConceptAmount").Value) * CDbl(asEmployeesQueries(2)) * CDbl(asEmployeesQueries(3)))
															ElseIf (CInt(oRecordset.Fields("ConceptID").Value) = 52) Or (CInt(oRecordset.Fields("ConceptID").Value) = 71) Then
'																dTaxAmount = dTaxAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
															Else
																dAmount = dAmount + (CDbl(oRecordset.Fields("ConceptAmount").Value) * CDbl(asEmployeesQueries(2)))
															End If
															oRecordset.MoveNext
															'If Err.number <> 0 Then Exit Do
														Loop
													Else
														Do While Not oRecordset.EOF
															If (CInt(oRecordset.Fields("ConceptID").Value) = 52) Or (CInt(oRecordset.Fields("ConceptID").Value) = 71) Then
'																dTaxAmount = dTaxAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
															Else
																dAmount = dAmount + (CDbl(oRecordset.Fields("ConceptAmount").Value) * CDbl(asEmployeesQueries(2)))
																If CInt(oRecordset.Fields("ConceptID").Value) = 1 Then dTemp = CDbl(oRecordset.Fields("ConceptAmount").Value)
															End If
															oRecordset.MoveNext
															'If Err.number <> 0 Then Exit Do
														Loop
													End If
												End If
											End If
											'If lErrorNumber <> 0 Then Exit For
											If bTimeout Then Exit For
										Next
										If (dAmount + dTaxAmount) > (dTemp * 0.3) Then dAmount = (dTemp * 0.3) - dTaxAmount
										lErrorNumber = AppendTextToFile(sFilePath & "_Payroll_Special52.txt", lCurrentID & ", " & asSpecialConcepts(kIndex) & ", 1, " & FormatNumber(dAmount, 2, True, False, False), sErrorDescription)
									End If
									Call DeleteFile(sFilePath & "_Payroll_71.txt", "")
								End If
							End If

							If Not bTimeout Then
Call DisplayTimeStamp("START: LEVEL 2, RUN FROM FILES, ESPECIALES 51 y 70. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
								sQueryBegin = "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ", 1, "
								sQueryEnd = ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")"
								asFileContents = GetFileContents(sFilePath & "_Payroll_Special52.txt", sErrorDescription)
								If Len(asFileContents) > 0 Then
									asFileContents = Split(asFileContents, vbNewLine)
									For iIndex = 0 To UBound(asFileContents)
										If Len(asFileContents(iIndex)) > 0 Then
											sErrorDescription = "No se pudo agregar el concepto de pago y su monto a la nómina del empleado."
											lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, sQueryBegin & asFileContents(iIndex) & sQueryEnd, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
										End If
										'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
										If bTimeout Then Exit For
									Next
									Call DeleteFile(sFilePath & "_Payroll_Special52.txt", "")
								End If
							End If
						Case 72 '71. Deducción por cobro de sueldos indebidos

						Case 90 'B3. Ajuste residentes
							sQueryBegin = ""
							If InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Zones"
							If InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Areas"
							If InStr(1, sCondition, "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
							If InStr(1, sCondition, "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
							If InStr(1, sCondition, "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
							If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
							sErrorDescription = "No se pudieron obtener los montos de la nómina de los empleados."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID, ConceptAmount From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons " & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And ((EmployeeTypeID=5) Or (PositionTypeID=5)) And (ConceptID In (1,4,89)) " & sCondition & " Order By EmployeesHistoryList.EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Select EmployeesHistoryList.EmployeeID, Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID, ConceptAmount From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons " & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And ((EmployeeTypeID=5) Or (PositionTypeID=5)) And (ConceptID In (1,4,89)) " & sCondition & " Order By EmployeesHistoryList.EmployeeID -->" & vbNewLine
							If lErrorNumber = 0 Then
								lCurrentID = -2
								adTotal = Split("0,0,0", ",")
								If Not oRecordset.EOF Then
									Do While Not oRecordset.EOF
										If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
											If lCurrentID <> -2 Then
												'If adTotal(1) > 0 Then
													dAmount = ((adTotal(0) + adTotal(2)) * 0.2) - adTotal(1)
													If dAmount < 0 Then dAmount = 0
													lErrorNumber = AppendTextToFile(sFilePath & "_Payroll_Special_" & Int(iCounter / ROWS_PER_FILE) & ".txt", lCurrentID & ", " & asSpecialConcepts(kIndex) & ", 1, " & FormatNumber(dAmount, 2, True, False, False), sErrorDescription)
													iCounter = iCounter + 1
												'End If
											End If
											lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
											adTotal(0) = 0
											adTotal(1) = 0
											adTotal(2) = 0
											dAmount = 0
										End If
										Select Case CLng(oRecordset.Fields("ConceptID").Value)
											Case 1
												adTotal(0) = CDbl(oRecordset.Fields("ConceptAmount").Value)
											Case 4
												adTotal(1) = CDbl(oRecordset.Fields("ConceptAmount").Value)
											Case 89
												adTotal(2) = CDbl(oRecordset.Fields("ConceptAmount").Value)
										End Select
										oRecordset.MoveNext
										'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
									Loop
									oRecordset.Close
									'If adTotal(1) > 0 Then
										dAmount = ((adTotal(0) + adTotal(2)) * 0.2) - adTotal(1)
										If dAmount < 0 Then dAmount = 0
										lErrorNumber = AppendTextToFile(sFilePath & "_Payroll_Special_" & Int(iCounter / ROWS_PER_FILE) & ".txt", lCurrentID & ", " & asSpecialConcepts(kIndex) & ", 1, " & FormatNumber(dAmount, 2, True, False, False), sErrorDescription)
										iCounter = iCounter + 1
									'End If
								End If
							End If
						Case 92 'B7. Ajuste a calendario
							If StrComp(Right(lPayID, Len("0000")), "1231", vbBinaryCompare) > 0 Then
								sErrorDescription = "No se pudo limpiar la tabla temporal."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll -->" & vbNewLine
								If lErrorNumber = 0 Then
									sQueryBegin = ""
									If InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Zones"
									If InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Areas"
									If InStr(1, sCondition, "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
									If InStr(1, sCondition, "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
									If InStr(1, sCondition, "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
									If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
									sErrorDescription = "No se pudieron obtener los montos de la nómina de los empleados."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, EmployeesHistoryList.EmployeeID, '92' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons " & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (1,3,7,8,38,49,89)) " & sCondition & " Group By EmployeesHistoryList.EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
									Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, EmployeesHistoryList.EmployeeID, '92' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons " & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (1,3,7,8,38,49,89)) " & sCondition & " Group By EmployeesHistoryList.EmployeeID -->" & vbNewLine
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudieron obtener los montos de la nómina de los empleados."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '2' As RecordID, EmployeeID, '92' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & Left(lPayID, Len("0000")) & " Where (EmployeeID In (Select EmployeeID From Payroll) And (RecordDate<" & lPayID & ") And (ConceptID In (1,3,7,8,38,49,89)) Group By EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
										Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '2' As RecordID, EmployeeID, '92' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & Left(lPayID, Len("0000")) & " Where (EmployeeID In (Select EmployeeID From Payroll) And (RecordDate<" & lPayID & ") And (ConceptID In (1,3,7,8,38,49,89)) Group By EmployeeID -->" & vbNewLine
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudieron obtener los montos de la nómina de los empleados."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '0' As RecordID, EmployeeID, '92' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll Group By EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
											Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '0' As RecordID, EmployeeID, '92' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll Group By EmployeeID -->" & vbNewLine
											If lErrorNumber = 0 Then
												sErrorDescription = "No se pudo limpiar la tabla temporal."
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (RecordID>0)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
												Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (RecordID>0) -->" & vbNewLine
												If lErrorNumber = 0 Then
													sErrorDescription = "No se pudieron obtener los montos de la nómina de los empleados."
													If (CInt(Left(lPayID, Len("0000"))) Mod 4) = 0 Then
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll Set ConceptAmount=ConceptAmount*6/15", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
														Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Update Payroll Set ConceptAmount=ConceptAmount*6/15 -->" & vbNewLine
													Else
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll Set ConceptAmount=ConceptAmount*5/15", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
														Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Update Payroll Set ConceptAmount=ConceptAmount*5/15 -->" & vbNewLine
													End If
													If lErrorNumber = 0 Then
														sErrorDescription = "No se pudieron obtener las incidencias de los empleados."
														lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesAbsencesLKP.EmployeeID, OcurredDate, EndDate, AbsenceHours From EmployeesAbsencesLKP, Payroll, Absences Where (EmployeesAbsencesLKP.EmployeeID=Payroll.EmployeeID) And (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And ((Absences.AbsenceTypeID2=0) Or (Absences.AbsenceID In (10))) And (EmployeesAbsencesLKP.OcurredDate<" & Left(lPayID, Len("0000")) & "9999) And (EmployeesAbsencesLKP.EndDate>" & Left(lPayID, Len("0000")) & "0000) Order By EmployeesAbsencesLKP.EmployeeID, EmployeesAbsencesLKP.OcurredDate", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
														Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Select EmployeesAbsencesLKP.EmployeeID, OcurredDate, EndDate, AbsenceHours From EmployeesAbsencesLKP, Payroll, Absences Where (EmployeesAbsencesLKP.EmployeeID=Payroll.EmployeeID) And (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And ((Absences.AbsenceTypeID2=0) Or (Absences.AbsenceID In (10))) And (EmployeesAbsencesLKP.OcurredDate<" & Left(lPayID, Len("0000")) & "9999) And (EmployeesAbsencesLKP.EndDate>" & Left(lPayID, Len("0000")) & "0000) Order By EmployeesAbsencesLKP.EmployeeID, EmployeesAbsencesLKP.OcurredDate -->" & vbNewLine
														If lErrorNumber = 0 Then
															dAmount = 0
															lCurrentID = -2
															iCounter2 = 0
															Do While Not oRecordset.EOF
																If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
																	If lCurrentID <> -2 Then
																		lErrorNumber = AppendTextToFile(sFilePath & "_Payroll_92.txt", lCurrentID & ", " & FormatNumber(dAmount, 2, True, False, False), sErrorDescription)
																		iCounter2 = iCounter2 + 1
																	End If
																	dAmount = 0
																	lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
																End If
																lTempStartDate = CLng(oRecordset.Fields("OcurredDate").Value)
																lTempEndDate = CLng(oRecordset.Fields("EndDate").Value)
																If lTempStartDate < CLng(Left(lPayID, Len("0000")) & "0101") Then lTempStartDate = CLng(Left(lPayID, Len("0000")) & "0101")
																If lTempEndDate > CLng(Left(lPayID, Len("0000")) & "1231") Then lTempEndDate = CLng(Left(lPayID, Len("0000")) & "1231")
																If lTempStartDate <= lTempEndDate Then
																	dAmount = dAmount + (Abs(DateDiff("d", GetDateFromSerialNumber(lTempStartDate), GetDateFromSerialNumber(lTempEndDate))) + 1)
																End If
																oRecordset.MoveNext
																'If Err.number <> 0 Then Exit Do
															Loop
															oRecordset.Close
															lErrorNumber = AppendTextToFile(sFilePath & "_Payroll_92.txt", lCurrentID & LIST_SEPARATOR & dAmount, sErrorDescription)
															iCounter2 = iCounter2 + 1
														End If
														If iCounter2 > 0 Then
															asFileContents = GetFileContents(sFilePath & "_Payroll_92.txt", sErrorDescription)
															If Len(asFileContents) > 0 Then
																asFileContents = Split(asFileContents, vbNewLine)
																For iIndex = 0 To UBound(asFileContents)
																	If Len(asFileContents(iIndex)) > 0 Then
																		asEmployeesQueries = Split(asFileContents(iIndex), LIST_SEPARATOR)
																		sErrorDescription = "No se pudieron obtener los montos registrados en la nómina de los empleados."
																		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll Set ConceptAmount=ConceptAmount*" & (360 - asEmployeesQueries(1)) & "/360 Where (EmployeeID=" & asEmployeesQueries(0) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
																		Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Update Payroll Set ConceptAmount=ConceptAmount*" & (360 - asEmployeesQueries(1)) & "/360 Where (EmployeeID=" & asEmployeesQueries(0) & ") -->" & vbNewLine
																	End If
																	'If lErrorNumber <> 0 Then Exit For
																Next
															End If
															Call DeleteFile(sFilePath & "_Payroll_92.txt", "")
														End If
														If lErrorNumber = 0 Then
															sErrorDescription = "No se pudieron agregar los montos de la nómina de los empleados."
															lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (ConceptAmount<0)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

															sErrorDescription = "No se pudieron agregar los montos de la nómina de los empleados."
															lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, EmployeeID, '92' As ConceptID, '1' As PayrollTypeID, ConceptAmount, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
															Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, EmployeeID, '92' As ConceptID, '1' As PayrollTypeID, ConceptAmount, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll -->" & vbNewLine
															If lErrorNumber = 0 Then
																sErrorDescription = "No se pudo limpiar la tabla temporal."
																lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
																Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll -->" & vbNewLine
															End If
														End If
													End If
												End If
											End If
										End If
									End If
								End If
							End If
						Case 151 'RQ. Rezago Quirúrgico
							sErrorDescription = "No se pudieron agregar los montos de la nómina de los empleados."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & lPayID & "' As RecordDate, '1' As RecordID, EmployeeID, '151' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From EmployeesSpecialJourneys Where (EmployeeID<800000) And (SpecialJourneyID=425) And (AppliedDate=" & lPayID & ") Group By EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End Select
				End If
				If bTimeout Then Exit For
			Next
		End If

		If Not bTimeout Then
			If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 2, RUN FROM FILES, ESPECIALES, " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
				sQueryBegin = "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ", 1, "
				sQueryEnd = ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")"
				For jIndex = 0 To iCounter Step ROWS_PER_FILE
					asFileContents = GetFileContents(sFilePath & "_Payroll_Special_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
					If Len(asFileContents) > 0 Then
						asFileContents = Split(asFileContents, vbNewLine)
						For iIndex = 0 To UBound(asFileContents)
							If Len(asFileContents(iIndex)) > 0 Then
								sErrorDescription = "No se pudo agregar el concepto de pago y su monto a la nómina del empleado."
								lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, sQueryBegin & asFileContents(iIndex) & sQueryEnd, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
							End If
							'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
							If bTimeout Then Exit For
						Next
					End If
					Call DeleteFile(sFilePath & "_Payroll_Special_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
				Next
			End If
		End If

		If Not bTimeout Then
			sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=44) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If

		If Not bTimeout Then
			If bTruncate Then
Call DisplayTimeStamp("START: LEVEL 2, TRUNCATE DECIMALS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
				If False Then
					sErrorDescription = "No se pudieron truncar los decimales de los montos."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=Round(ConceptAmount, 2) Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Else
					sErrorDescription = "No se pudo limpiar la tabla temporal."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PayrollInt", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudieron truncar los decimales de los montos."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=(ConceptAmount+0.005)*100 Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudieron truncar los decimales de los montos."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PayrollInt (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo limpiar la tabla de montos."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudieron truncar los decimales de los montos."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID From PayrollInt", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudieron truncar los decimales de los montos."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=(ConceptAmount/100) Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	End If

	sCondition = ""
	Call GetConditionFromURL(oRequest, sCondition, -1, -1)
	sCondition = Replace(sCondition, "(Employees.", "(EmployeesHistoryList.")
	sCondition = Replace(sCondition, "(Companies.", "(EmployeesHistoryList.")
	sCondition = Replace(sCondition, "(EmployeeTypes.", "(EmployeesHistoryList.")
	sCondition = Replace(sCondition, "(PositionTypes.", "(EmployeesHistoryList.")
	sCondition = Replace(sCondition, "(Journeys.", "(EmployeesHistoryList.")
	sCondition = Replace(sCondition, "(Shifts.", "(EmployeesHistoryList.")
	sCondition = Replace(sCondition, "(Levels.", "(EmployeesHistoryList.")
	sCondition = Replace(sCondition, "(PaymentCenters.", "(EmployeesHistoryList.")
	sCondition = Replace(sCondition, "(Positions.", "(EmployeesHistoryList.")
	'sCondition = Replace(sCondition, "(Areas.", "(EmployeesHistoryList.")
	sCondition = Replace(sCondition, "(Jobs.", "(EmployeesHistoryList.")
	If InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0 Then sCondition = " And (Areas.ZoneID=Zones.ZoneID)" & sCondition
	If InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0 Then sCondition = " And (EmployeesHistoryList.AreaID=Areas.AreaID)" & sCondition
	If InStr(1, sCondition, "(Jobs.", vbBinaryCompare) > 0 Then sCondition = " And (EmployeesHistoryList.JobID=Jobs.JobID)" & sCondition
	If bRetroactive And (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) Then sCondition = " And (EmployeesHistoryList.EmployeeID=EmployeesRevisions.EmployeeID) And (EmployeesRevisions.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesRevisions.StartPayrollID=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")" & sCondition

Call DisplayTimeStamp("START: LEVEL 2, TAXES. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
	If Not bTimeout Then
		sErrorDescription = "No se pudieron obtener las tablas del ISR inverso."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select InferiorLimit, SuperiorLimit, FixedAmount, PercentageForExcess From TaxLimits Where (StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (PeriodID=4)", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			adTaxes = ""
			Do While Not oRecordset.EOF
				adTaxes = adTaxes & CStr(oRecordset.Fields("InferiorLimit").Value) & "," & CStr(oRecordset.Fields("SuperiorLimit").Value) & "," & CStr(oRecordset.Fields("FixedAmount").Value) & "," & CStr(oRecordset.Fields("PercentageForExcess").Value) & ";"
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			adTaxes = adTaxes & "0,1E+69,0,1"
		End If
		adTaxes = Split(adTaxes, ";")
		For iIndex = 0 To UBound(adTaxes)
			adTaxes(iIndex) = Split(adTaxes(iIndex), ",")
			For jIndex = 0 To UBound(adTaxes(iIndex))
				adTaxes(iIndex)(jIndex) = CDbl(adTaxes(iIndex)(jIndex))
			Next
		Next
	End If

	If Not bTimeout Then
'		If bMonthlyTaxes Then
			sErrorDescription = "No se pudieron obtener las tablas del ISR inverso."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select InferiorLimit, SuperiorLimit, AllowanceAmount From EmploymentAllowances Where (StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (PeriodID=4)", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				adAllowances = ""
				Do While Not oRecordset.EOF
					adAllowances = adAllowances & CStr(oRecordset.Fields("InferiorLimit").Value) & "," & CStr(oRecordset.Fields("SuperiorLimit").Value) & "," & CStr(oRecordset.Fields("AllowanceAmount").Value) & ";"
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				adAllowances = adAllowances & "0,1E+69,0"
			End If
			adAllowances = Split(adAllowances, ";")
			For iIndex = 0 To UBound(adAllowances)
				adAllowances(iIndex) = Split(adAllowances(iIndex), ",")
				For jIndex = 0 To UBound(adAllowances(iIndex))
					adAllowances(iIndex)(jIndex) = CDbl(adAllowances(iIndex)(jIndex))
				Next
			Next
'		End If
	End If

	If Not bTimeout Then
		sErrorDescription = "No se pudieron obtener las tablas del ISR inverso."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select InferiorLimit, SuperiorLimit, InvertedTax, InvertedRate From TaxInvertions Where (StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (PeriodID=4)", "ReportsQueries1000Lib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			adTaxInvertions = ""
			Do While Not oRecordset.EOF
				adTaxInvertions = adTaxInvertions & CStr(oRecordset.Fields("InferiorLimit").Value) & "," & CStr(oRecordset.Fields("SuperiorLimit").Value) & "," & CStr(oRecordset.Fields("InvertedTax").Value) & "," & CStr(oRecordset.Fields("InvertedRate").Value) & ";"
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			adTaxInvertions = adTaxInvertions & "0,1E+69,0,1"
		End If
		adTaxInvertions = Split(adTaxInvertions, ";")
		For iIndex = 0 To UBound(adTaxInvertions)
			adTaxInvertions(iIndex) = Split(adTaxInvertions(iIndex), ",")
			For jIndex = 0 To UBound(adTaxInvertions(iIndex))
				adTaxInvertions(iIndex)(jIndex) = CDbl(adTaxInvertions(iIndex)(jIndex))
			Next
		Next
	End If

	If Not bTimeout Then
Call DisplayTimeStamp("START: LEVEL 2, CREATE FILES, TAXES. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		adDSM = ""
		sErrorDescription = "No se pudieron obtener los días de salario mínimo y los días de salario burocrático."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CurrencyValue From CurrenciesHistoryList Where (CurrencyDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (CurrencyID In (1,2,3,4,5)) Order By CurrencyID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				adDSM = adDSM & ";" & CDbl(oRecordset.Fields("CurrencyValue").Value)
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			adDSM = Split(adDSM, ";")
			For iIndex = 1 To UBound(adDSM)
				adDSM(iIndex) = CDbl(adDSM(iIndex))
			Next
		End If
	End If

	If (Not bTimeout) And bMonthlyTaxes And (CInt(Right(aPayrollComponent(N_ID_PAYROLL), Len("0000"))) <> 106) Then
		lErrorNumber = CalculateQttyID_8_9(oRequest, oADODBConnection, False, bRetroactive, sErrorDescription)
	End If

	If Not bTimeout Then
		sCondition = Replace(sCondition, "Employees.", "EmployeesHistoryList.")
		sQueryBegin = ""
		If InStr(1, sCondition, "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
		If InStr(1, sCondition, "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
		If InStr(1, sCondition, "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
		If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
		If bMonthlyTaxes Then
			sErrorDescription = "No se pudieron obtener los montos totales de las percepciones gravables de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordID As RecordDate1, -ConceptRetention As RecordID1, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, 0 As ConceptRetention1, -55 As UserID From Payroll_" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("YYYY")) & " Where (RecordDate<>" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (RecordID>=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("YYYYMM")) & "00) And (RecordID<" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Zones, Areas" & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) " & sCondition & "))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write "<!-- Query (55): Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordID As RecordDate1, -ConceptRetention As RecordID1, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, 0 As ConceptRetention1, -55 As UserID From Payroll_" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("YYYY")) & " Where (RecordDate<>" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (RecordID>=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("YYYYMM")) & "00) And (RecordID<" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Zones, Areas" & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) " & sCondition & "))" & vbNewLine

			sErrorDescription = "No se pudieron obtener los montos totales de las percepciones gravables de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeTypeID, Zones.ZoneTypeID, Concepts.ConceptID, Concepts.IsDeduction, Sum(ConceptAmount) As TotalAmount, Sum(ConceptTaxes) As TotalTaxes From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Concepts, Zones, Areas" & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Concepts.ConceptID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (((Concepts.IsDeduction=0) And (Concepts.TaxAmount=100)) Or (Concepts.ConceptID In (55,120)) Or ((Concepts.IsDeduction=1) And (Concepts.TaxAmount=0))) And (Concepts.ConceptID>0) And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (((Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordDate>=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("YYYYMM")) & "00) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordDate<=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("YYYYMM")) & "99)) Or (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordID=-" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeTypeID, Zones.ZoneTypeID, Concepts.ConceptID, Concepts.IsDeduction Order By EmployeesHistoryList.EmployeeID, Concepts.ConceptID, Concepts.IsDeduction", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write "<!-- Query (55): Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeTypeID, Zones.ZoneTypeID, Concepts.ConceptID, Concepts.IsDeduction, Sum(ConceptAmount) As TotalAmount, Sum(ConceptTaxes) As TotalTaxes From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Concepts, Zones, Areas" & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Concepts.ConceptID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (((Concepts.IsDeduction=0) And (Concepts.TaxAmount=100)) Or (Concepts.ConceptID In (55,120)) Or ((Concepts.IsDeduction=1) And (Concepts.TaxAmount=0))) And (Concepts.ConceptID>0) And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (((Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordDate>=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("YYYYMM")) & "00) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordDate<=" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("YYYYMM")) & "99)) Or (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordID=-" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeTypeID, Zones.ZoneTypeID, Concepts.ConceptID, Concepts.IsDeduction Order By EmployeesHistoryList.EmployeeID, Concepts.ConceptID, Concepts.IsDeduction -->" & vbNewLine
		Else
			sErrorDescription = "No se pudieron obtener los montos totales de las percepciones gravables de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeTypeID, Zones.ZoneTypeID, Concepts.ConceptID, Concepts.IsDeduction, Sum(ConceptAmount) As TotalAmount, Sum(ConceptTaxes) As TotalTaxes From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Concepts, Zones, Areas" & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Concepts.ConceptID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (((Concepts.IsDeduction=0) And (Concepts.TaxAmount=100)) Or (Concepts.ConceptID=120) Or ((Concepts.IsDeduction=1) And (Concepts.TaxAmount=0))) And (Concepts.ConceptID>0) And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And ((Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") Or (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordID=-" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeTypeID, Zones.ZoneTypeID, Concepts.ConceptID, Concepts.IsDeduction Order By EmployeesHistoryList.EmployeeID, Concepts.ConceptID, Concepts.IsDeduction", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write "<!-- Query (55): Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeTypeID, Zones.ZoneTypeID, Concepts.ConceptID, Concepts.IsDeduction, Sum(ConceptAmount) As TotalAmount, Sum(ConceptTaxes) As TotalTaxes From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Concepts, Zones, Areas" & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Concepts.ConceptID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (((Concepts.IsDeduction=0) And (Concepts.TaxAmount=100)) Or (Concepts.ConceptID=120) Or ((Concepts.IsDeduction=1) And (Concepts.TaxAmount=0))) And (Concepts.ConceptID>0) And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And ((Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") Or (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordID=-" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")) " & sCondition & " Group By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeTypeID, Zones.ZoneTypeID, Concepts.ConceptID, Concepts.IsDeduction Order By EmployeesHistoryList.EmployeeID, Concepts.ConceptID, Concepts.IsDeduction -->" & vbNewLine
		End If
		iCounter = 0
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sCurrentID = CStr(oRecordset.Fields("EmployeeID").Value)
				lEmployeeTypeID = CLng(oRecordset.Fields("EmployeeTypeID").Value)
				dAmount = 0
				dAmount_55 = 0
				dAmount_88 = 0
				Do While Not oRecordset.EOF
					If StrComp(sCurrentID, CStr(oRecordset.Fields("EmployeeID").Value), vbBinaryCompare) <> 0 Then
						If bMonthlyTaxes Then
							dTaxAmount = dAmount
						Else
							dTaxAmount = dAmount * 2
						End If
						dTemp = dTaxAmount
						For iIndex = 0 To UBound(adTaxes)
							If (adTaxes(iIndex)(0) <= dTaxAmount) And (adTaxes(iIndex)(1) >= dTaxAmount) Then
								dTaxAmount = ((dTaxAmount - adTaxes(iIndex)(0)) * adTaxes(iIndex)(3) / 100) + adTaxes(iIndex)(2)
								Exit For
							End If
						Next
						If bMonthlyTaxes And (lEmployeeTypeID < 7) And (CInt(Right(aPayrollComponent(N_ID_PAYROLL), Len("0000"))) <> 106) Then
							For iIndex = 0 To UBound(adAllowances)
								If (adAllowances(iIndex)(0) <= dTemp) And (adAllowances(iIndex)(1) >= dTemp) Then
									'46. Crédito al salario
									lErrorNumber = AppendTextToFile(sFilePath & "_Payroll53_" & Int(iCounter / ROWS_PER_FILE) & ".txt", sCurrentID & ", 48, 1, " & FormatNumber(adAllowances(iIndex)(2), 2, True, False, False), sErrorDescription)
									iCounter = iCounter + 1
									Exit For
								End If
							Next
						End If
						If bMonthlyTaxes Then
							dTaxAmount = dTaxAmount - dAmount_55
						Else
							dTaxAmount = dTaxAmount / 2
						End If
						'53. Impuesto sobre la renta
						lErrorNumber = AppendTextToFile(sFilePath & "_Payroll53_" & Int(iCounter / ROWS_PER_FILE) & ".txt", sCurrentID & ", 55, 1, " & FormatNumber(dTaxAmount, 2, True, False, False), sErrorDescription)
						iCounter = iCounter + 1

						If dAmount_88 > 0 Then
							'AN. Aportación neta patronal del Seguro de Separación Individualizado
							'lErrorNumber = AppendTextToFile(sFilePath & "_Payroll53_" & Int(iCounter / ROWS_PER_FILE) & ".txt", sCurrentID & ", 88, 1, " & FormatNumber(dAmount_88, 2, True, False, False), sErrorDescription)
							'iCounter = iCounter + 1

							dTemp = dAmount - dTaxAmount - dAmount_55 + dAmount_88
							If Not bMonthlyTaxes Then dTemp = dTemp * 2
							For iIndex = 0 To UBound(adTaxInvertions)
								If (adTaxInvertions(iIndex)(0) <= dTemp) And (adTaxInvertions(iIndex)(1) >= dTemp) Then
									dTemp = (dTemp - adTaxInvertions(iIndex)(2)) / adTaxInvertions(iIndex)(3)
									Exit For
								End If
							Next
							If Not bMonthlyTaxes Then
								dTemp = dTemp / 2
								dTemp = dTemp - dAmount
							Else
								dTemp = dTemp - dAmount
								dTemp = dTemp / 2
							End If
							'SS. Aportación patronal bruta (Cuota para el Seguro de Separación Individualizado)
							lErrorNumber = AppendTextToFile(sFilePath & "_Payroll53_" & Int(iCounter / ROWS_PER_FILE) & ".txt", sCurrentID & ", 122, 1, " & FormatNumber(dTemp, 2, True, False, False), sErrorDescription)
							iCounter = iCounter + 1

							If lEmployeeTypeID < 7 Then
								dTemp = dTemp - dAmount_88
								'IS. ISR patronal del Seguro de Separación Individualizado
								lErrorNumber = AppendTextToFile(sFilePath & "_Payroll53_" & Int(iCounter / ROWS_PER_FILE) & ".txt", sCurrentID & ", 110, 1, " & FormatNumber(dTemp, 2, True, False, False), sErrorDescription)
								iCounter = iCounter + 1
							End If
						End If
						sCurrentID = CStr(oRecordset.Fields("EmployeeID").Value)
						lEmployeeTypeID = CLng(oRecordset.Fields("EmployeeTypeID").Value)
						dAmount = 0
						dAmount_55 = 0
						dAmount_88 = 0
					End If
					If InStr(1, ",9,10,11,16,17,", ("," & CStr(oRecordset.Fields("ConceptID").Value) & ","), vbBinaryCompare) > 0 Then
						dAmount = dAmount + CDbl(oRecordset.Fields("TotalTaxes").Value) '5 por semana
					ElseIf InStr(1, ",20,21,", ("," & CStr(oRecordset.Fields("ConceptID").Value) & ","), vbBinaryCompare) > 0 Then
						dTemp = CDbl(oRecordset.Fields("TotalAmount").Value) - (adDSM(CInt(oRecordset.Fields("ZoneTypeID").Value)) * 7.5)
						If dTemp < 0 Then dTemp = 0
						dAmount = dAmount + dTemp
					ElseIf CInt(oRecordset.Fields("IsDeduction").Value) = 0 Then
						dAmount = dAmount + CDbl(oRecordset.Fields("TotalAmount").Value)
					ElseIf CLng(oRecordset.Fields("ConceptID").Value) = 55 Then
						dAmount_55 = CDbl(oRecordset.Fields("TotalAmount").Value)
					ElseIf CLng(oRecordset.Fields("ConceptID").Value) = 88 Then
					ElseIf CLng(oRecordset.Fields("ConceptID").Value) = 120 Then
						dAmount_88 = CDbl(oRecordset.Fields("TotalAmount").Value)
					Else
						dAmount = dAmount - CDbl(oRecordset.Fields("TotalAmount").Value)
					End If
					oRecordset.MoveNext
					'If lErrorNumber <> 0 Then Exit Do
				Loop
				oRecordset.Close
				If bMonthlyTaxes Then
					dTaxAmount = dAmount
				Else
					dTaxAmount = dAmount * 2
				End If
				dTemp = dTaxAmount
				For iIndex = 0 To UBound(adTaxes)
					If (adTaxes(iIndex)(0) <= dTaxAmount) And (adTaxes(iIndex)(1) >= dTaxAmount) Then
						dTaxAmount = ((dTaxAmount - adTaxes(iIndex)(0)) * adTaxes(iIndex)(3) / 100) + adTaxes(iIndex)(2)
						Exit For
					End If
				Next
				If bMonthlyTaxes And (lEmployeeTypeID < 7) And (CInt(Right(aPayrollComponent(N_ID_PAYROLL), Len("0000"))) <> 106) Then
					For iIndex = 0 To UBound(adAllowances)
						If (adAllowances(iIndex)(0) <= dTemp) And (adAllowances(iIndex)(1) >= dTemp) Then
							Response.Write "<!-- (46) " & dTemp & " -->" & vbNewLine
							'46. Crédito al salario
							lErrorNumber = AppendTextToFile(sFilePath & "_Payroll53_" & Int(iCounter / ROWS_PER_FILE) & ".txt", sCurrentID & ", 48, 1, " & FormatNumber(adAllowances(iIndex)(2), 2, True, False, False), sErrorDescription)
							iCounter = iCounter + 1
							Exit For
						End If
					Next
				End If
				If bMonthlyTaxes Then
					dTaxAmount = dTaxAmount - dAmount_55
				Else
					dTaxAmount = dTaxAmount / 2
				End If
				'53. Impuesto sobre la renta
				lErrorNumber = AppendTextToFile(sFilePath & "_Payroll53_" & Int(iCounter / ROWS_PER_FILE) & ".txt", sCurrentID & ", 55, 1, " & FormatNumber(dTaxAmount, 2, True, False, False), sErrorDescription)
				iCounter = iCounter + 1
				If dAmount_88 > 0 Then
					'AN. Aportación neta patronal del Seguro de Separación Individualizado
					'lErrorNumber = AppendTextToFile(sFilePath & "_Payroll53_" & Int(iCounter / ROWS_PER_FILE) & ".txt", sCurrentID & ", 88, 1, " & FormatNumber(dAmount_88, 2, True, False, False), sErrorDescription)
					'iCounter = iCounter + 1

					dTemp = dAmount - dTaxAmount - dAmount_55 + dAmount_88
					If Not bMonthlyTaxes Then dTemp = dTemp * 2
					For iIndex = 0 To UBound(adTaxInvertions)
						If (adTaxInvertions(iIndex)(0) <= dTemp) And (adTaxInvertions(iIndex)(1) >= dTemp) Then
							dTemp = (dTemp - adTaxInvertions(iIndex)(2)) / adTaxInvertions(iIndex)(3)
							Exit For
						End If
					Next
					If Not bMonthlyTaxes Then
						dTemp = dTemp / 2
						dTemp = dTemp - dAmount
					Else
						dTemp = dTemp - dAmount
						dTemp = dTemp / 2
					End If
					'SS. Aportación patronal bruta (Cuota para el Seguro de Separación Individualizado)
					lErrorNumber = AppendTextToFile(sFilePath & "_Payroll53_" & Int(iCounter / ROWS_PER_FILE) & ".txt", sCurrentID & ", 122, 1, " & FormatNumber(dTemp, 2, True, False, False), sErrorDescription)
					iCounter = iCounter + 1

					If lEmployeeTypeID < 7 Then
						dTemp = dTemp - dAmount_88
						'IS. ISR patronal del Seguro de Separación Individualizado
						lErrorNumber = AppendTextToFile(sFilePath & "_Payroll53_" & Int(iCounter / ROWS_PER_FILE) & ".txt", sCurrentID & ", 110, 1, " & FormatNumber(dTemp, 2, True, False, False), sErrorDescription)
						iCounter = iCounter + 1
					End If
				End If
			End If
		End If
	End If

	If Not bTimeout Then
		If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 2, RUN FROM FILES, TAXES, " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
			sQueryBegin = "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ", 1, "
			sQueryEnd = ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")"
			For jIndex = 0 To iCounter Step ROWS_PER_FILE
				asFileContents = GetFileContents(sFilePath & "_Payroll53_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
				If Len(asFileContents) > 0 Then
					asFileContents = Split(asFileContents, vbNewLine)
					For iIndex = 0 To UBound(asFileContents)
						If Len(asFileContents(iIndex)) > 0 Then
							sErrorDescription = "No se pudo agregar el concepto de pago y su monto a la nómina del empleado."
							lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, sQueryBegin & asFileContents(iIndex) & sQueryEnd, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
						End If
						'If lErrorNumber <> 0 Then Exit For
						If bTimeout Then Exit For
					Next
				End If
				Call DeleteFile(sFilePath & "_Payroll53_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
			Next
		End If
		sErrorDescription = "No se pudieron obtener los montos totales de las percepciones gravables de los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where UserID=-55", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If

	If Not bTimeout Then
		sErrorDescription = "No se pudieron eliminar los montos de SI y SS que no aplican."
		'AN,IS,SI,SS|AN,SI
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (88,110,120,122)) And (EmployeeID Not In (Select EmployeeID From EmployeesConceptsLKP Where (ConceptID In (88,120)) And (StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptAmount>0) And (Active=1)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

		sErrorDescription = "No se pudieron eliminar los montos de SI y SS que no aplican."
		'AN,IS,SI,SS|AN,SI
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (88,110,120,122)) And (EmployeeID Not In (Select EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (88,120)) And (ConceptAmount>0.004)))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If

	If Not bTimeout Then
Call DisplayTimeStamp("START: LEVEL 2, ConceptID=70 (69. Pensión alimenticia). " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		sQueryBegin = ""
		If InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Zones"
		If InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Areas"
		If InStr(1, sCondition, "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
		If InStr(1, sCondition, "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
		If InStr(1, sCondition, "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
		If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
		If aPayrollComponent(N_TYPE_ID_PAYROLL) <> 3 Then
			sErrorDescription = "No se pudieron obtener los montos de las pensiones alimenticias."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & aPayrollComponent(N_FOR_DATE_PAYROLL) & "' As RecordDate, EmployeesBeneficiariesLKP.EmployeeID As RecordID, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, '124' As ConceptID, '1' As PayrollTypeID, ConceptAmount, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From EmployeesBeneficiariesLKP, AlimonyTypes, EmployeesChangesLKP, EmployeesHistoryList" & sQueryBegin & " Where (EmployeesBeneficiariesLKP.AlimonyTypeID=AlimonyTypes.AlimonyTypeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (AlimonyTypes.ConceptQttyID=1) And (EmployeesBeneficiariesLKP.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesBeneficiariesLKP.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")" & sCondition, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If

		If lErrorNumber = 0 Then
			asEmployeesQueries = ""
			sErrorDescription = "No se pudo limpiar la tabla temporal."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AlimonyTypeID, AppliesToID From AlimonyTypes Where (ConceptQttyID=2) And (AppliesToID Is Not Null)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				Do While Not oRecordset.EOF
					asEmployeesQueries = asEmployeesQueries & CStr(oRecordset.Fields("AlimonyTypeID").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("AppliesToID").Value) & LIST_SEPARATOR
					oRecordset.MoveNext
				Loop
				oRecordset.Close
			End If
			asEmployeesQueries = Split(asEmployeesQueries, LIST_SEPARATOR)
			For iIndex = 0 To UBound(asEmployeesQueries) - 1
				asEmployeesQueries(iIndex) = Split(asEmployeesQueries(iIndex), SECOND_LIST_SEPARATOR)
				sErrorDescription = "No se pudo limpiar la tabla temporal."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudieron obtener los montos de las pensiones alimenticias."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '0' As RecordDate, EmployeesBeneficiariesLKP.EmployeeID As RecordID, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, '124' As ConceptID, '1' As PayrollTypeID, Sum(Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptAmount * EmployeesBeneficiariesLKP.ConceptAmount / 100) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", Concepts, EmployeesBeneficiariesLKP, AlimonyTypes, EmployeesChangesLKP, EmployeesHistoryList" & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesBeneficiariesLKP.EmployeeID) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Concepts.ConceptID) And (EmployeesBeneficiariesLKP.AlimonyTypeID=AlimonyTypes.AlimonyTypeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And ((Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") Or (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordID=-" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")) And (AlimonyTypes.AlimonyTypeID=" & asEmployeesQueries(iIndex)(0) & ") And (EmployeesBeneficiariesLKP.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesBeneficiariesLKP.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.IsDeduction=0) And (Concepts.ConceptID In (" & asEmployeesQueries(iIndex)(1) & "))" & sCondition & " Group By EmployeesBeneficiariesLKP.EmployeeID, EmployeesBeneficiariesLKP.BeneficiaryNumber", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudieron obtener los montos de las pensiones alimenticias."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '1' As RecordDate, EmployeesBeneficiariesLKP.EmployeeID As RecordID, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, '124' As ConceptID, '1' As PayrollTypeID, -Sum(Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptAmount * EmployeesBeneficiariesLKP.ConceptAmount / 100) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", Concepts, EmployeesBeneficiariesLKP, AlimonyTypes, EmployeesChangesLKP, EmployeesHistoryList" & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesBeneficiariesLKP.EmployeeID) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Concepts.ConceptID) And (EmployeesBeneficiariesLKP.AlimonyTypeID=AlimonyTypes.AlimonyTypeID) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And ((Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") Or (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordID=-" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")) And (AlimonyTypes.AlimonyTypeID=" & asEmployeesQueries(iIndex)(0) & ") And (EmployeesBeneficiariesLKP.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesBeneficiariesLKP.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.IsDeduction=1) And (Concepts.ConceptID In (" & asEmployeesQueries(iIndex)(1) & "))" & sCondition & " Group By EmployeesBeneficiariesLKP.EmployeeID, EmployeesBeneficiariesLKP.BeneficiaryNumber", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudieron obtener los montos de las pensiones alimenticias."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & aPayrollComponent(N_FOR_DATE_PAYROLL) & "' As RecordDate, RecordID, EmployeeID, '124' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll Group By RecordID, EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo limpiar la tabla temporal."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							End If
						End If
					End If
				End If
				If bTimeout Then Exit For
			Next
		End If
	End If

	If Not bTimeout Then
		sErrorDescription = "No se pudieron obtener los montos de las pensiones alimenticias."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID=124) And (ConceptAmount<=0)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

		sErrorDescription = "No se pudieron obtener los montos de las pensiones alimenticias."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & aPayrollComponent(N_FOR_DATE_PAYROLL) & "' As RecordDate, '1' As RecordID1, RecordID As EmployeeID, '70' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount), '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=124) And (RecordID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & ")) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") Group By RecordID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	If Not bTimeout Then
Call DisplayTimeStamp("START: LEVEL 2, ConceptID=154 (RM. Retención de pago a terceros por mandato Judicial). " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		sQueryBegin = ""
		If InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Zones"
		If InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Areas"
		If InStr(1, sCondition, "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
		If InStr(1, sCondition, "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
		If InStr(1, sCondition, "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
		If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
		If aPayrollComponent(N_TYPE_ID_PAYROLL) <> 3 Then
			sErrorDescription = "No se pudieron obtener los montos de las pensiones alimenticias."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & aPayrollComponent(N_FOR_DATE_PAYROLL) & "' As RecordDate, EmployeesCreditorsLKP.EmployeeID As RecordID, EmployeesCreditorsLKP.CreditorNumber As EmployeeID, '155' As ConceptID, '1' As PayrollTypeID, ConceptAmount, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From EmployeesCreditorsLKP, CreditorsTypes, EmployeesChangesLKP, EmployeesHistoryList" & sQueryBegin & " Where (EmployeesCreditorsLKP.CreditorTypeID=CreditorsTypes.CreditorTypeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (CreditorsTypes.ConceptQttyID=1) And (EmployeesCreditorsLKP.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesCreditorsLKP.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")" & sCondition, "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If

		If lErrorNumber = 0 Then
			asEmployeesQueries = ""
			sErrorDescription = "No se pudo limpiar la tabla temporal."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CreditorTypeID, AppliesToID From CreditorsTypes Where (ConceptQttyID=2) And (AppliesToID Is Not Null)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				Do While Not oRecordset.EOF
					asEmployeesQueries = asEmployeesQueries & CStr(oRecordset.Fields("CreditorTypeID").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("AppliesToID").Value) & LIST_SEPARATOR
					oRecordset.MoveNext
				Loop
				oRecordset.Close
			End If
			asEmployeesQueries = Split(asEmployeesQueries, LIST_SEPARATOR)
			For iIndex = 0 To UBound(asEmployeesQueries) - 1
				asEmployeesQueries(iIndex) = Split(asEmployeesQueries(iIndex), SECOND_LIST_SEPARATOR)
				sErrorDescription = "No se pudo limpiar la tabla temporal."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudieron obtener los montos de las pensiones alimenticias."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '0' As RecordDate, EmployeesCreditorsLKP.EmployeeID As RecordID, EmployeesCreditorsLKP.CreditorNumber As EmployeeID, '155' As ConceptID, '1' As PayrollTypeID, Sum(Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptAmount * EmployeesCreditorsLKP.ConceptAmount / 100) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", Concepts, EmployeesCreditorsLKP, CreditorsTypes, EmployeesChangesLKP, EmployeesHistoryList" & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesCreditorsLKP.EmployeeID) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Concepts.ConceptID) And (EmployeesCreditorsLKP.CreditorTypeID=CreditorsTypes.CreditorTypeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And ((Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") Or (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordID=-" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")) And (CreditorsTypes.CreditorTypeID=" & asEmployeesQueries(iIndex)(0) & ") And (EmployeesCreditorsLKP.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesCreditorsLKP.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.IsDeduction=0) And (Concepts.ConceptID In (" & asEmployeesQueries(iIndex)(1) & "))" & sCondition & " Group By EmployeesCreditorsLKP.EmployeeID, EmployeesCreditorsLKP.CreditorNumber", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudieron obtener los montos de las pensiones alimenticias."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '1' As RecordDate, EmployeesCreditorsLKP.EmployeeID As RecordID, EmployeesCreditorsLKP.CreditorNumber As EmployeeID, '155' As ConceptID, '1' As PayrollTypeID, -Sum(Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptAmount * EmployeesCreditorsLKP.ConceptAmount / 100) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", Concepts, EmployeesCreditorsLKP, CreditorsTypes, EmployeesChangesLKP, EmployeesHistoryList" & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesCreditorsLKP.EmployeeID) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Concepts.ConceptID) And (EmployeesCreditorsLKP.CreditorTypeID=CreditorsTypes.CreditorTypeID) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And ((Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") Or (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".RecordID=-" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")) And (CreditorsTypes.CreditorTypeID=" & asEmployeesQueries(iIndex)(0) & ") And (EmployeesCreditorsLKP.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesCreditorsLKP.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.IsDeduction=1) And (Concepts.ConceptID In (" & asEmployeesQueries(iIndex)(1) & "))" & sCondition & " Group By EmployeesCreditorsLKP.EmployeeID, EmployeesCreditorsLKP.CreditorNumber", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudieron obtener los montos de las pensiones alimenticias."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & aPayrollComponent(N_FOR_DATE_PAYROLL) & "' As RecordDate, RecordID, EmployeeID, '155' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll Group By RecordID, EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo limpiar la tabla temporal."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							End If
						End If
					End If
				End If
				If bTimeout Then Exit For
			Next
		End If
	End If

	If Not bTimeout Then
		sErrorDescription = "No se pudieron obtener los montos de las pensiones alimenticias."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & aPayrollComponent(N_FOR_DATE_PAYROLL) & "' As RecordDate, '1' As RecordID1, RecordID As EmployeeID, '154' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount), '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=155) And (RecordID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & ")) And (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") Group By RecordID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	If Not bTimeout Then
		sErrorDescription = "No se pudieron obtener los montos de las pensiones alimenticias."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set RecordID=1 Where (RecordID=-" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	If Not bTimeout Then
		If bTruncate Then
Call DisplayTimeStamp("START: LEVEL 2, TRUNCATE DECIMALS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
			sTruncate = "154,155"
			If False Then
				sErrorDescription = "No se pudieron truncar los decimales de los montos."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=Round(ConceptAmount, 2) Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (" & sTruncate & "))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Else
				sErrorDescription = "No se pudo limpiar la tabla temporal."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PayrollInt", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudieron truncar los decimales de los montos."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=(ConceptAmount+0.005)*100 Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (" & sTruncate & "))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudieron truncar los decimales de los montos."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PayrollInt (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (" & sTruncate & "))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudo limpiar la tabla de montos."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (" & sTruncate & "))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudieron truncar los decimales de los montos."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID From PayrollInt", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudieron truncar los decimales de los montos."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=(ConceptAmount/100) Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (" & sTruncate & "))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	End If

	If Not bTimeout Then
		If True Or bRetroactive Then
			sErrorDescription = "No se pudo limpiar la tabla temporal."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudieron obtener los montos totales de las deducciones de los empleados."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudieron obtener los montos históricos de las deducciones de los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select -RecordID As RecordDate, ConceptRetention As RecordID1, EmployeeID, ConceptID, '1' As PayrollTypeID, -Sum(ConceptAmount) As ConceptAmount1, Sum(ConceptTaxes) As ConceptTaxes1, '0' As ConceptRetention1, '-1' As UserID From Payroll_" & Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000")) & " Where (RecordID=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select Distinct EmployeeID From Payroll)) And (ConceptID Not In (-2,-1,0)) Group By RecordID, ConceptRetention, EmployeeID, ConceptID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						sErrorDescription = "No se pudo limpiar la tabla temporal."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudieron obtener los montos totales de las deducciones de los empleados."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & aPayrollComponent(N_FOR_DATE_PAYROLL) & "' As RecordDate, RecordID, EmployeeID, ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, Sum(ConceptTaxes) As ConceptTaxes1, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll Group By EmployeeID, RecordID, ConceptID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudo limpiar la tabla temporal."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							End If
						End If
					End If
				End If
			End If

			If Not bTimeout Then
				If bTruncate Then
Call DisplayTimeStamp("START: LEVEL 2, TRUNCATE DECIMALS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
					If False Then
						sErrorDescription = "No se pudieron truncar los decimales de los montos."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=Round(ConceptAmount, 2) Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Else
						sErrorDescription = "No se pudo limpiar la tabla temporal."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PayrollInt", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudieron truncar los decimales de los montos."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=(ConceptAmount+0.005)*100 Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudieron truncar los decimales de los montos."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PayrollInt (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo limpiar la tabla de montos."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudieron truncar los decimales de los montos."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID From PayrollInt", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudieron truncar los decimales de los montos."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=(ConceptAmount/100) Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	End If

	Set oRecordset = Nothing
	DoCalculations = lErrorNumber
	Err.Clear
End Function
%>