<%
Function GetAbsencesHours(oADODBConnection, lEmployeeID, lAbsenceID, lStartDate, lEndDate, dHours, sErrorDescription)
'************************************************************
'Purpose: To get the number of absence hours registered for the employee
'Inputs:  oADODBConnection, lEmployeeID, lAbsenceID, lStartDate, lEndDate
'Outputs: dHours, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetAbsencesHours"
	Dim oRecordset
	Dim lErrorNumber

	dHours = 0
	Set oRecordset = Nothing
	GetAbsencesHours = lErrorNumber
	Err.Clear
End Function

Function GetConceptAmount(oADODBConnection, lPayrollID, lEmployeeID, dWorkingHours, lZoneID, lEconomicZoneID, sConceptID, bFromPayroll, bFromConceptValues, dAmount, sErrorDescription)
'************************************************************
'Purpose: To get the amount for the user concept
'Inputs:  oADODBConnection, lPayrollID, lEmployeeID, dWorkingHours, lZoneID, sConceptID, bFromPayroll, bFromConceptValues
'Outputs: dAmount, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetConceptAmount"
	Dim oRecordset
	Dim dTemp
	Dim dHours
	Dim dLimit
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener el monto del concepto."
	If bFromPayroll Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptAmount From Payroll_" & lPayrollID & " Where (EmployeeID=" & lEmployeeID & ") And (ConceptID In (" & sConceptID & "))", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	ElseIf bFromConceptValues Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From ConceptsValues Where (ConceptID In (" & sConceptID & ")) And (StartDate<=" & lPayrollID & ") And (EndDate>=" & lPayrollID & ")", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesConceptsLKP Where (EmployeeID=" & lEmployeeID & ") And (ConceptID In (" & sConceptID & ")) And (StartDate<=" & lPayrollID & ") And (EndDate>=" & lPayrollID & ")", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If
	dAmount = 0
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			If bFromPayroll Then
				dAmount = dAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
			Else
				Select Case CInt(oRecordset.Fields("ConceptQttyID").Value)
					Case 1 '$
						dAmount = dAmount + CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 2 '%
						Select Case CInt(oRecordset.Fields("ConceptTypeID").Value)
							Case 1 'Sobre el sueldo bruto
								dAmount = dAmount + CDbl(oRecordset.Fields("ConceptAmount").Value) / 100
							Case 2 'Sobre el sueldo neto
								dAmount = dAmount + CDbl(oRecordset.Fields("ConceptAmount").Value) / 100
							Case 3 'Sobre otro concepto
								lErrorNumber = GetConceptAmount(oADODBConnection, lPayrollID, lEmployeeID, dWorkingHours, lZoneID, lEconomicZoneID, CStr(oRecordset.Fields("AppliesToID").Value), True, False, dTemp, sErrorDescription)
								dAmount = dAmount + (dTemp * CDbl(oRecordset.Fields("ConceptAmount").Value) / 100)
						End Select
					Case 3 'DSM
						lErrorNumber = GetCurrencyValue(oADODBConnection, CInt(oRecordset.Fields("ConceptMinQttyID").Value), lZoneID, lPayrollID, dTemp, sErrorDescription)
						dAmount = dAmount + CDbl(dTemp) * CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 4 'Unidades
					Case 5 'Días de salario
						lErrorNumber = GetConceptAmount(oADODBConnection, lPayrollID, lEmployeeID, dWorkingHours, lZoneID, lEconomicZoneID, "1,3", True, True, dTemp, sErrorDescription)
						dAmount = dAmount + CDbl(oRecordset.Fields("ConceptAmount").Value) * dTemp / 30
					Case 6 'por hora
						lErrorNumber = GetConceptAmount(oADODBConnection, lPayrollID, lEmployeeID, dWorkingHours, lZoneID, lEconomicZoneID, "1,3", True, True, dTemp, sErrorDescription)
						dAmount = dAmount + CDbl(oRecordset.Fields("ConceptAmount").Value) * dTemp / (30 * dWorkingHours)
					Case 7 'Diferencia entre el puesto actual y el superior
					Case 8 'sobre las horas extras laboradas
						dHours = 0
						lErrorNumber = GetSpecialHours(oADODBConnection, lPayrollID, lEmployeeID, CInt(oRecordset.Fields("ConceptQttyID").Value), "And (Absences.AbsenceID=201)", dHours, sErrorDescription)
						dAmount = dAmount + CDbl(oRecordset.Fields("ConceptAmount").Value) * dHours
					Case 9 'sobre los domingos laborados
						dHours = 0
						lErrorNumber = GetSpecialHours(oADODBConnection, lPayrollID, lEmployeeID, CInt(oRecordset.Fields("ConceptQttyID").Value), "And (Absences.AbsenceID=202)", dHours, sErrorDescription)
						dAmount = dAmount + CDbl(oRecordset.Fields("ConceptAmount").Value) * dHours
					Case 10 'sobre el tiempo efectivo
					Case 11 'de acuerdo a la puntualidad
					Case 12 'Prima vacacional
					Case 13 'Días de salario burocrático
						lErrorNumber = GetCurrencyValue(oADODBConnection, CInt(oRecordset.Fields("ConceptQttyID").Value), lEconomicZoneID, lPayrollID, dTemp, sErrorDescription)
						dAmount = dAmount + CDbl(dTemp) * CDbl(oRecordset.Fields("ConceptAmount").Value)
					Case 14 'Monedas de oro
					Case 15 'Vales de despensa
					Case 16 'Evaluación de desempeño
				End Select
				If Not bFromConceptValues Then
					If CDbl(oRecordset.Fields("ConceptMin").Value) > 0 Then
						Select Case CInt(oRecordset.Fields("ConceptMinQttyID").Value)
							Case 1 '$
								If dAmount < CDbl(oRecordset.Fields("ConceptMin").Value) Then dAmount = CDbl(oRecordset.Fields("ConceptMin").Value)
							Case 3 'DSM
								lErrorNumber = GetCurrencyValue(oADODBConnection, CInt(oRecordset.Fields("ConceptMinQttyID").Value), lZoneID, lPayrollID, dLimit, sErrorDescription)
								dLimit = CDbl(dLimit) * CDbl(oRecordset.Fields("ConceptMin").Value)
								If dAmount < dLimit Then dAmount = dLimit
							Case 4 'Unidades
								If CInt(oRecordset.Fields("ConceptQttyID").Value) = 8 Then
									If dHours < CDbl(oRecordset.Fields("ConceptMin").Value) Then
										dAmount = dAmount - CDbl(oRecordset.Fields("ConceptAmount").Value) * dHours
										dHours = CDbl(oRecordset.Fields("ConceptMin").Value)
										dAmount = dAmount + CDbl(oRecordset.Fields("ConceptAmount").Value) * dHours
									End If
								Else
								End If
							Case 5 'Días de salario
								lErrorNumber = GetConceptAmount(oADODBConnection, lPayrollID, lEmployeeID, dWorkingHours, lZoneID, lEconomicZoneID, "1,3", True, True, dLimit, sErrorDescription)
								dLimit = CDbl(oRecordset.Fields("ConceptAmount").Value) * dLimit / 30
								If dAmount < dLimit Then dAmount = dLimit
							Case 13 'Días de salario burocrático
								lErrorNumber = GetCurrencyValue(oADODBConnection, CInt(oRecordset.Fields("ConceptMinQttyID").Value), lEconomicZoneID, lPayrollID, dLimit, sErrorDescription)
								dLimit = CDbl(dLimit) * CDbl(oRecordset.Fields("ConceptMin").Value)
								If dAmount < dLimit Then dAmount = dLimit
						End Select
					End If
					If CDbl(oRecordset.Fields("ConceptMax").Value) > 0 Then
						Select Case CInt(oRecordset.Fields("ConceptMaxQttyID").Value)
							Case 1 '$
								If dAmount > CDbl(oRecordset.Fields("ConceptMax").Value) Then dAmount = CDbl(oRecordset.Fields("ConceptMax").Value)
							Case 3 'DSM
								lErrorNumber = GetCurrencyValue(oADODBConnection, CInt(oRecordset.Fields("ConceptMaxQttyID").Value), lZoneID, lPayrollID, dLimit, sErrorDescription)
								dLimit = CDbl(dLimit) * CDbl(oRecordset.Fields("ConceptMax").Value)
								If dAmount > dLimit Then dAmount = dLimit
							Case 4 'Unidades
								If CInt(oRecordset.Fields("ConceptQttyID").Value) = 8 Then
									If dHours > CDbl(oRecordset.Fields("ConceptMax").Value) Then
										dAmount = dAmount - CDbl(oRecordset.Fields("ConceptAmount").Value) * dHours
										dHours = CDbl(oRecordset.Fields("ConceptMax").Value)
										dAmount = dAmount + CDbl(oRecordset.Fields("ConceptAmount").Value) * dHours
									End If
								Else
								End If
							Case 5 'Días de salario
								lErrorNumber = GetConceptAmount(oADODBConnection, lPayrollID, lEmployeeID, dWorkingHours, lZoneID, lEconomicZoneID, "1,3", True, True, dLimit, sErrorDescription)
								dLimit = dLimit + CDbl(oRecordset.Fields("ConceptAmount").Value) * dLimit / 30
								If dAmount > dLimit Then dAmount = dLimit
							Case 13 'Días de salario burocrático
								lErrorNumber = GetCurrencyValue(oADODBConnection, CInt(oRecordset.Fields("ConceptMaxQttyID").Value), lEconomicZoneID, lPayrollID, dLimit, sErrorDescription)
								dLimit = CDbl(dLimit) * CDbl(oRecordset.Fields("ConceptMax").Value)
								If dAmount > dLimit Then dAmount = dLimit
						End Select
					End If
				End If
			End If
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GetConceptAmount = lErrorNumber
	Err.Clear
End Function

Function GetConsecutiveID(oADODBConnection, lTypeID, lNewConsecutiveID, sErrorDescription)
'************************************************************
'Purpose: To get the next consecutive ID for the given record
'Inputs:  oADODBConnection, lTypeID
'Outputs: lNewConsecutiveID, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetConsecutiveID"
	Dim oRecordset
	Dim lErrorNumber

	lNewConsecutiveID = -1
	sErrorDescription = "No se pudo obtener el siguiente número consecutivo."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CurrentID From ConsecutiveIDs Where (IDType=" & lTypeID & ")", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			lNewConsecutiveID = CLng(oRecordset.Fields("CurrentID").Value) + 1
		Else
			Select Case lTypeID
				Case 1061
					lErrorNumber = GetNewIDFromTable(oADODBConnection, "Paperworks", "PaperworkID", "", 1, lNewConsecutiveID, sErrorDescription)
				Case Else
					lNewConsecutiveID = -1
			End Select
		End If
	End If

	Set oRecordset = Nothing
	GetConsecutiveID = lErrorNumber
	Err.Clear
End Function

Function GetCurrencyValue(oADODBConnection, lCurrencyID, lZoneTypeID, lPayrollID, sDSMValue, sErrorDescription)
'************************************************************
'Purpose: To get the value for the Minimum Wage for the given
'         zone and date.
'Inputs:  oADODBConnection, lCurrencyID, lZoneTypeID, lPayrollID
'Outputs: sDSMValue, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetCurrencyValue"
	Dim oRecordset
	Dim lErrorNumber

	Select Case lCurrencyID
		Case 13 'Días de Salario Burocrático
			sErrorDescription = "No se pudo obtener el valor del Salario Mínimo."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CurrencyValue From CurrenciesHistoryList Where (CurrencyID=" & lZoneTypeID + 2 & ") And (CurrencyDate=" & lPayrollID & ")", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If oRecordset.EOF Then
					sErrorDescription = "No se pudo obtener el valor del Salario Mínimo."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CurrencyValue From CurrenciesHistoryList Where (CurrencyID=" & lZoneTypeID + 2 & ") And (CurrencyDate<=" & lPayrollID & ") Order By CurrencyDate Desc", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If oRecordset.EOF Then
							sDSMValue = 0
						Else
							sDSMValue = CDbl(oRecordset.Fields("CurrencyValue").Value)
						End If
						oRecordset.Close
					End If
				Else
					sDSMValue = CDbl(oRecordset.Fields("CurrencyValue").Value)
				End If
				oRecordset.Close
			End If
		Case 14 'Monedas de oro
		Case 15 'Vales de despensa
		Case Else
			sErrorDescription = "No se pudo obtener el valor del Salario Mínimo."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CurrencyValue From CurrenciesHistoryList Where (CurrencyID=" & lZoneTypeID & ") And (CurrencyDate=" & lPayrollID & ")", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If oRecordset.EOF Then
					sErrorDescription = "No se pudo obtener el valor del Salario Mínimo."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CurrencyValue From CurrenciesHistoryList Where (CurrencyID=" & lZoneTypeID & ") And (CurrencyDate<=" & lPayrollID & ") Order By CurrencyDate Desc", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber = 0 Then
						If oRecordset.EOF Then
							sDSMValue = 0
						Else
							sDSMValue = CDbl(oRecordset.Fields("CurrencyValue").Value)
						End If
						oRecordset.Close
					End If
				Else
					sDSMValue = CDbl(oRecordset.Fields("CurrencyValue").Value)
				End If
				oRecordset.Close
			End If
	End Select

	Set oRecordset = Nothing
	GetCurrencyValue = lErrorNumber
	Err.Clear
End Function

Function GetLastPayroll(oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To get the status of the last payroll
'Inputs:  oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetLastPayroll"
	Dim oRecordset
	Dim lErrorNumber

	GetLastPayroll = -1
	sErrorDescription = "No se pudo obtener el número de la nómina."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PayrollID From Payrolls Order By PayrollDate Desc", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then GetLastPayroll = CLng(oRecordset.Fields("PayrollID").Value)
	End If

	Set oRecordset = Nothing
	Err.Clear
End Function

Function GetLastPayrollStatus(oADODBConnection, iPayrollID, iPayrollStatus, sErrorDescription)
'************************************************************
'Purpose: To get the status of the last payroll
'Inputs:  oADODBConnection
'Outputs: iPayrollID, iPayrollStatus, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetLastPayrollStatus"
	Dim oRecordset
	Dim lErrorNumber

	iPayrollID = -1
	iPayrollStatus = -1
	sErrorDescription = "No se pudo obtener el estatus de la nómina."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PayrollID, IsClosed From Payrolls Order By PayrollDate Desc", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			iPayrollID = CStr(oRecordset.Fields("PayrollID").Value)
			iPayrollStatus = CInt(oRecordset.Fields("IsClosed").Value)
		End If
	End If

	GetLastPayrollStatus = lErrorNumber
	Set oRecordset = Nothing
	Err.Clear
End Function

Function GetLastUserEntryToSystem(oADODBConnection, lUserID, sLastEntry, sErrorDescription)
'************************************************************
'Purpose: To get the last time the user entered to the system
'Inputs:  oADODBConnection, lUserID
'Outputs: sLastEntry, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetLastUserEntryToSystem"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la última vez que el usuario entró al sistema."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From SystemLogs Where (UserID=" & lUserID & ") Order By LogYear Desc, LogMonth Desc, LogDay Desc, LogHour Desc, LogMinute Desc, LogSecond Desc", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			sLastEntry = ""
		Else
			sLastEntry = DisplayDate(CInt(oRecordset("LogYear").Value), CInt(oRecordset("LogMonth").Value), CInt(oRecordset("LogDay").Value), CInt(oRecordset("LogHour").Value), CInt(oRecordset("LogMinute").Value), CInt(oRecordset("LogSecond").Value))
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GetLastUserEntryToSystem = lErrorNumber
	Err.Clear
End Function

Function GetNameFromTable(oADODBConnection, sTableName, sIDs, sTab, sSeparator, sNames, sErrorDescription)
'************************************************************
'Purpose: To get the name of the records given the IDs
'Inputs:  oADODBConnection, sTableName, sIDs, sTab, sSeparator
'Outputs: sNames, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetNameFromTable"
	Dim sKeyFieldName
	Dim sFieldName
	Dim sIDsForTable
	Dim iIndex
	Dim aTemp
	Dim sTemp
	Dim oRecordset
	Dim lErrorNumber

	sIDsForTable = sIDs
	If Len(sIDsForTable) = 0 Then sIDsForTable = "-2"
	Select Case sTableName
		Case "AreasURCTAUX"
			sTableName = "Areas"
			sKeyFieldName = "AreaID"
			sFieldName = "URCTAUX"
		Case "Absences"
			sKeyFieldName = "AbsenceID"
			sFieldName = "AbsenceShortName, AbsenceName"
		Case "Absences1"
			sTableName = "Absences"
			sKeyFieldName = "AbsenceID"
			sFieldName = "AbsenceShortName"
		Case "AbsenceTypes"
			sKeyFieldName = "AbsenceTypeID"
			sFieldName = "AbsenceTypeName"
		Case "AlimonyTypes"
			sKeyFieldName = "AlimonyTypeID"
			sFieldName = "AlimonyTypeName"
		Case "Antiquities"
			sKeyFieldName = "AntiquityID"
			sFieldName = "AntiquityName"
		Case "AreaLevelTypes"
			sKeyFieldName = "AreaLevelTypeID"
			sFieldName = "AreaLevelTypeName"
		Case "Areas"
			'If InStr(1, sIDs, ",", vbBinaryCompare) = 0 Then
				sKeyFieldName = "AreaID"
			'Else
			'	sKeyFieldName = "AreaPath"
			'End If
			sFieldName = "AreaCode, AreaName"
		Case "AreasFromCodes"
			sTableName = "Areas"
			sKeyFieldName = "AreaCode"
			sFieldName = "AreaID"
		Case "AuditOperationTypes"
			sKeyFieldName = "AuditOperationTypeID"
			sFieldName = "AuditOperationName"
		Case "AuditTypes"
			sKeyFieldName = "AuditTypeID"
			sFieldName = "AuditTypeName"
		Case "ParentAreaIDs"
			sTableName = "Areas"
			sKeyFieldName = "AreaID"
			sFieldName = "ParentID"
		Case "CodeAreas"
			sTableName = "Areas"
			sKeyFieldName = "AreaID"
			sFieldName = "AreaCode"
		Case "FullAreas"
			sTableName = "Areas"
			sKeyFieldName = "AreaID"
			sFieldName = "AreaCode, AreaName"
		Case "ShortAreas"
			sTableName = "Areas"
			sKeyFieldName = "AreaID"
			sFieldName = "AreaShortName"
		Case "SubAreas"
			sTableName = "Areas"
			sKeyFieldName = "AreaID"
			sFieldName = "AreaCode, AreaName"
		Case "AreaTypes"
			sKeyFieldName = "AreaTypeID"
			sFieldName = "AreaTypeName"
		Case "AttentionLevels"
			sKeyFieldName = "AttentionLevelID"
			sFieldName = "AttentionLevelName"
		Case "BankAccounts"
			sKeyFieldName = "AccountID"
			sFieldName = "AccountNumber"
		Case "EmployeeAccount"
			sTableName = "BankAccounts "
			sKeyFieldName = "(EndDate=30000000) And EmployeeID"
			sFieldName = "AccountNumber"
		Case "Banks"
			sKeyFieldName = "BankID"
			sFieldName = "BankName"
		Case "Branches"
			sKeyFieldName = "BranchID"
			sFieldName = "BranchShortName, BranchName"

		Case "Budgets"
			sKeyFieldName = "BudgetID"
			sFieldName = "BudgetShortName, BudgetName"
		Case "BudgetPath"
			sTableName = "Budgets "
			sKeyFieldName = "BudgetID"
			sFieldName = "BudgetPath"
		Case "BudgetsShortName"
			sTableName = "Budgets"
			sKeyFieldName = "BudgetID"
			sFieldName = "BudgetShortName"
		Case "BudgetsFunds"
			sKeyFieldName = "FundID"
			sFieldName = "FundName"
		Case "BudgetsDuties"
			sKeyFieldName = "DutyID"
			sFieldName = "DutyName"
		Case "BudgetsActiveDuties"
			sKeyFieldName = "ActiveDutyID"
			sFieldName = "ActiveDutyName"
		Case "BudgetsSpecificDuties"
			sKeyFieldName = "SpecificDutyID"
			sFieldName = "SpecificDutyName"
		Case "BudgetsPrograms"
			sKeyFieldName = "ProgramID"
			sFieldName = "ProgramName"
		Case "BudgetsConfineTypes"
			sKeyFieldName = "ConfineTypeID"
			sFieldName = "ConfineTypeName"
		Case "BudgetsActivities1"
			sKeyFieldName = "ActivityID"
			sFieldName = "ActivityName"
		Case "BudgetsActivities2"
			sKeyFieldName = "ActivityID"
			sFieldName = "ActivityName"
		Case "BudgetsProcesses"
			sKeyFieldName = "ProcessID"
			sFieldName = "ProcessName"
		Case "BudgetTypes"
			sKeyFieldName = "BudgetTypeID"
			sFieldName = "BudgetTypeName"

		Case "CashierOffices"
			sKeyFieldName = "CashierOfficeID"
			sFieldName = "CashierOfficeShortName, CashierOfficeName"
		Case "CenterSubtypes"
			sKeyFieldName = "CenterSubtypeID"
			sFieldName = "CenterSubtypeShortName, CenterSubtypeName"
		Case "CenterTypes"
			sKeyFieldName = "CenterTypeID"
			sFieldName = "CenterTypeShortName, CenterTypeName"
		Case "Companies"
			sKeyFieldName = "CompanyID"
			sFieldName = "CompanyShortName, CompanyName"
		Case "CompanyTypes"
			sKeyFieldName = "CompanyTypeID"
			sFieldName = "CompanyTypeName"
		Case "Concepts"
			sKeyFieldName = "(Concepts.EndDate=30000000) And ConceptID"
			sFieldName = "ConceptShortName, ConceptName"
		Case "ShortConcepts"
			sTableName = "Concepts"
			sKeyFieldName = "(Concepts.EndDate=30000000) And ConceptID"
			sFieldName = "ConceptShortName"
		Case "FullConcepts"
			sTableName = "Concepts"
			sKeyFieldName = "(Concepts.EndDate=30000000) And ConceptID"
			sFieldName = "ConceptShortName, ConceptName"
		Case "ConceptsNames"
			sTableName = "Concepts"
			sKeyFieldName = "(Concepts.EndDate=30000000) And ConceptID"
			sFieldName = "ConceptName"
		Case "ConceptTypes"
			sKeyFieldName = "ConceptTypeID"
			sFieldName = "ConceptTypeName"
		Case "ConfineTypes"
			sKeyFieldName = "ConfineTypeID"
			sFieldName = "ConfineTypeShortName, ConfineTypeName"
		Case "Countries"
			sKeyFieldName = "CountryID"
			sFieldName = "CountryName"
		Case "Nationalities"
			sTableName = "Countries"
			sKeyFieldName = "CountryID"
			sFieldName = "Nationality"
		Case "Credits"
			sKeyFieldName = "CreditID"
			sFieldName = "CreditID As CreditID2"
		Case "CreditsFiles"
			sKeyFieldName = "UploadedFileName"
			sFieldName = "UploadedFileName"
		Case "CreditTypes"
			sKeyFieldName = "CreditTypeID"
			sFieldName = "CreditTypeShortName, CreditTypeName"
		Case "Currencies"
			sKeyFieldName = "CurrencyID"
			sFieldName = "CurrencyName"
		Case "CurrenciesSymbols"
			sTableName = "Currencies"
			sKeyFieldName = "CurrencyID"
			sFieldName = "CurrencySymbol"
		Case "CurrenciesValues"
			sTableName = "CurrenciesHistoryList"
			sKeyFieldName = "(CurrencyDate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ") And CurrencyID"
			sFieldName = "CurrencyValue"
		Case "EconomicZones"
			sKeyFieldName = "EconomicZoneID"
			sFieldName = "EconomicZoneCode, EconomicZoneName"
		Case "CodeEconomicZones"
			sTableName = "EconomicZones"
			sKeyFieldName = "EconomicZoneID"
			sFieldName = "EconomicZoneCode"
		Case "EmployeeActivities"
			sKeyFieldName = "EmployeeActivityID"
			sFieldName = "EmployeeActivityName"
		Case "Employees"
			sKeyFieldName = "EmployeeID"
			sFieldName = "EmployeeName, EmployeeLastName, EmployeeLastName2"
		Case "EmployeesGenders"
			sTableName = "Employees"
			sKeyFieldName = "EmployeeID"
			sFieldName = "GenderID"
		Case "EmployeesIDFromAccessKey"
			sTableName = "Employees"
			sKeyFieldName = "EmployeeAccessKey"
			sFieldName = "EmployeeID"
		Case "EmployeesIDFromNumber"
			sTableName = "Employees"
			sKeyFieldName = "EmployeeNumber"
			sFieldName = "EmployeeID"
		Case "EmployeesIDFromRFC"
			sTableName = "Employees"
			sKeyFieldName = "RFC"
			sFieldName = "EmployeeID"
		Case "EmployeesNameFromNumber"
			sTableName = "Employees"
			sKeyFieldName = "EmployeeNumber"
			sFieldName = "EmployeeName, EmployeeLastName, EmployeeLastName2"
		Case "EmployeesNumber"
			sTableName = "Employees"
			sKeyFieldName = "EmployeeID"
			sFieldName = "EmployeeNumber"
		Case "EmployeeTypes"
			sKeyFieldName = "EmployeeTypeID"
			sFieldName = "EmployeeTypeShortName, EmployeeTypeName"
		Case "Genders"
			sKeyFieldName = "GenderID"
			sFieldName = "GenderName"
		Case "GeneratingAreas"
			sKeyFieldName = "GeneratingAreaID"
			sFieldName = "GeneratingAreaShortName, GeneratingAreaName"
		Case "GroupGradeLevels"
			sKeyFieldName = "GroupGradeLevelID"
			sFieldName = "GroupGradeLevelShortName,GroupGradeLevelName"
		Case "Handicaps"
			sKeyFieldName = "HandicapID"
			sFieldName = "HandicapName"
		Case "Holidays"
			sKeyFieldName = "(Holiday>" & sIDs & "0000) And (Holiday<" & sIDs & "9999) And Holiday Not "
			sFieldName = "Holiday"
		Case "Holiday"
			sTableName = "Holidays "
			sKeyFieldName = "Holiday"
			sFieldName = "Holiday"
		Case "HolidayDescription"
			sTableName = "Holidays "
			sKeyFieldName = "Holiday"
			sFieldName = "HolidayDescription"
		Case "Jobs"
			sKeyFieldName = "JobID"
			sFieldName = "JobNumber"
		Case "EmployeeIDsFromJobs"
			sTableName = "Employees"
			sKeyFieldName = "JobID"
			sFieldName = "EmployeeID"
		Case "JobTypes"
			sKeyFieldName = "JobTypeID"
			sFieldName = "JobTypeShortName, JobTypeName"
		Case "Journeys"
			sKeyFieldName = "JourneyID"
			sFieldName = "JourneyShortName, JourneyName"
		Case "ShortJourneys"
			sTableName = "Journeys"
			sKeyFieldName = "JourneyID"
			sFieldName = "JourneyShortName"
		Case "Justifications"
			sKeyFieldName = "JustificationID"
			sFieldName = "JustificationShortName, JustificationName"
		Case "Kardex5Origins"
			sKeyFieldName = "Kardex5OriginID"
			sFieldName = "Kardex5OriginName"
		Case "Kardex5Types"
			sKeyFieldName = "Kardex5TypeID"
			sFieldName = "Kardex5TypeName"
		Case "KardexChangeTypes"
			sKeyFieldName = "KardexChangeTypeID"
			sFieldName = "KardexChangeTypeName"
		Case "KardexOrigins"
			sKeyFieldName = "KardexOriginID"
			sFieldName = "KardexOriginName"
		Case "KardexRequirements"
			sKeyFieldName = "KardexRequirementID"
			sFieldName = "KardexRequirementName"
		Case "KardexTypes"
			sKeyFieldName = "KardexTypeID"
			sFieldName = "KardexTypeName"
		Case "Levels"
			sKeyFieldName = "LevelID"
			sFieldName = "LevelName"
		Case "LicenseSyndicateTypes"
			sKeyFieldName = "LicenseSyndicateTypeID"
			sFieldName = "LicenseSyndicateTypeName"
		Case "LicenseTypes"
			sKeyFieldName = "LicenseTypeID"
			sFieldName = "LicenseTypeName"
		Case "MaritalStatus"
			sKeyFieldName = "MaritalStatusID"
			sFieldName = "MaritalStatusName"
		Case "MedicalAreasTypes"
			sKeyFieldName = "MedicalAreasTypeID"
			sFieldName = "MedicalAreasTypeName"
		Case "OccupationTypes"
			sKeyFieldName = "OccupationTypeID"
			sFieldName = "OccupationTypeShortName, OccupationTypeName"
		Case "ShortOccupationTypes"
			sTableName = "OccupationTypes"
			sKeyFieldName = "OccupationTypeID"
			sFieldName = "OccupationTypeShortName"
		Case "PaperworkActions"
			sKeyFieldName = "PaperworkActionID"
			sFieldName = "PaperworkActionName"
		Case "PaperworkOwners"
			sTableName = "PaperworkOwners, Employees"
			sKeyFieldName = "(PaperworkOwners.EmployeeID=Employees.EmployeeID) And OwnerID"
			sFieldName = "OwnerName, EmployeeName, EmployeeLastName, EmployeeLastName2"
		Case "PaperworkSenders"
			sTableName = "PaperworkSenders"
			sKeyFieldName = "SenderID"
			sFieldName = "SenderID, SenderName, EmployeeName, PositionName"
		Case "PaperworkTypes"
			sKeyFieldName = "PaperworkTypeID"
			sFieldName = "PaperworkTypeName"
		Case "PaymentCenters"
			sKeyFieldName = "PaymentCenterID"
			sFieldName = "PaymentCenterShortName, PaymentCenterName"
		Case "PaymentsPayrollIDs"
			sTableName = "Payments"
			sKeyFieldName = "PaymentID"
			sFieldName = "PaymentDate"
		Case "PaymentsCancelationPayrollIDs"
			sTableName = "Payments"
			sKeyFieldName = "PaymentID"
			sFieldName = "CancelDate"
		Case "PaymentTypes"
			sKeyFieldName = "PaymentTypeID"
			sFieldName = "PaymentTypeName"
			sFieldName = "CancelDate"
		Case "Payrolls"
			sKeyFieldName = "PayrollID"
			sFieldName = "PayrollName"
		Case "ForPayrollID"
			sTableName = "Payrolls"
			sKeyFieldName = "PayrollID"
			sFieldName = "ForPayrollDate"
		Case "LastClosedPayrollID"
			sTableName = "Payrolls"
			sKeyFieldName = "(IsClosed=1) And (PayrollTypeID=1) And PayrollID Not"
			sFieldName = "Max(PayrollID)"
		Case "LastPayrollID"
			sTableName = "Payrolls"
			sKeyFieldName = "PayrollID Not"
			sFieldName = "Max(PayrollID)"
		Case "OpenPayrolls"
			sTableName = "Payrolls"
			sKeyFieldName = "(IsClosed<>1) And PayrollID Not"
			sFieldName = "PayrollName"
		Case "OpenPayrollIDs"
			sTableName = "Payrolls"
			sKeyFieldName = "(IsClosed<>1) And PayrollID Not"
			sFieldName = "PayrollID"
		Case "PayrollsTypes"
			sTableName = "Payrolls"
			sKeyFieldName = "PayrollID"
			sFieldName = "PayrollTypeID"
		Case "Periods"
			sKeyFieldName = "PeriodID"
			sFieldName = "PeriodName"
		Case "Positions"
			sIDs = split(sIDs, ",")
			sIDsForTable = CLng(sIDs(0))
			If UBound(sIDs) > 0 Then
				sKeyFieldName = " StartDate =" & CStr(sIDs(1)) & " And PositionID "
			Else 
				sKeyFieldName = " PositionID "
			End IF
			sFieldName = "PositionShortName, PositionName"
		Case "Priorities"
			sKeyFieldName = "PriorityID"
			sFieldName = "PriorityName"
		Case "ShortPositions"
			sTableName = "Positions"
			sKeyFieldName = "PositionID"
			sFieldName = "PositionShortName"
		Case "FullPositions"
			sTableName = "Positions"
			sKeyFieldName = "PositionID"
			sFieldName = "PositionShortName, PositionName"
		Case "PositionTypes"
			sKeyFieldName = "PositionTypeID"
			sFieldName = "PositionTypeShortName, PositionTypeName"
		Case "PositionTypes2"
			sKeyFieldName = "PositionTypeID"
			sFieldName = "PositionTypeName"
		Case "QttyValues"
			sKeyFieldName = "QttyID"
			sFieldName = "QttyValue"
		Case "QttyNames"
			sTableName = "QttyValues"
			sKeyFieldName = "QttyID"
			sFieldName = "QttyName"
		Case "Reasons"
			sKeyFieldName = "ReasonID"
			sFieldName = "ReasonShortName, ReasonName"
		Case "Requirements"
			sKeyFieldName = "RequirementID"
			sFieldName = "RequirementName"
		Case "RequirementsTypes"
			sKeyFieldName = "RequirementsTypeID"
			sFieldName = "RequirementsTypeName"
		Case "RiskLevels"
			sKeyFieldName = "RiskLevelID"
			sFieldName = "RiskLevelName"

		Case "SADE_Curso"
			sKeyFieldName = "ID_Curso"
			sFieldName = "Nombre_Curso"
		Case "SADE_Perfiles"
			sKeyFieldName = "ID_Perfil"
			sFieldName = "Nombre_Perfil"

		Case "Schoolarships"
			sKeyFieldName = "SchoolarshipID"
			sFieldName = "SchoolarshipName"
		Case "Services"
			sKeyFieldName = "ServiceID"
			sFieldName = "ServiceShortName, ServiceName"
		Case "Shifts"
			sKeyFieldName = "ShiftID"
			sFieldName = "ShiftShortName, ShiftName"
		Case "States"
			sKeyFieldName = "StateID"
			sFieldName = "StateCode, StateName"
		Case "CodeStates"
			sTableName = "States"
			sKeyFieldName = "StateID"
			sFieldName = "StateCode"
		Case "ShortStates"
			sTableName = "States"
			sKeyFieldName = "StateID"
			sFieldName = "StateShortName"
		Case "Status", "StatusAreas", "StatusBudgets", "StatusConceptsValues", "StatusEmployees", "StatusForms", "StatusJobs", "StatusLevels", "StatusPaperworks", "StatusPayments", "StatusPositions"
			sKeyFieldName = "StatusID"
			sFieldName = "StatusName"
		Case "StatusEmployeesActive"
			sTableName = "StatusEmployees"
			sKeyFieldName = "StatusID"
			sFieldName = "Active"
		Case "SubBranches"
			sKeyFieldName = "SubBranchID"
			sFieldName = "SubBranchShortName, SubBranchName"
		Case "SubjectTypes"
			sKeyFieldName = "SubjectTypeID"
			sFieldName = "SubjectTypeID, SubjectTypeName"
		Case "Syndicates"
			sKeyFieldName = "SyndicateID"
			sFieldName = "SyndicateShortName, SyndicateName"

		Case "TACO_AggregationTypes"
			sKeyFieldName = "AggregationTypeID"
			sFieldName = "AggregationTypeName"
		Case "TACO_Areas"
			sKeyFieldName = "AreaID"
			sFieldName = "AreaName"
		Case "TACO_Categories"
			sKeyFieldName = "CategoryID"
			sFieldName = "CategoryName"
		Case "TACO_Companies"
			sKeyFieldName = "CompanyID"
			sFieldName = "CompanyName"
		Case "TACO_FieldTypes"
			sKeyFieldName = "FieldTypeID"
			sFieldName = "FieldTypeName"
		Case "TACO_Labels"
			sKeyFieldName = "LabelID"
			sFieldName = "LabelName"
		Case "TACO_LabelsFullName"
			sTableName = TACO_PREFIX & "Labels"
			sKeyFieldName = "LabelID"
			sFieldName = "LabelArticle, LabelName"
		Case "TACO_LabelsFullNameForTask"
			sTableName = TACO_PREFIX & "Tasks, " & TACO_PREFIX & "Labels"
			aTemp = Split(sIDsForTable, ",", 2, vbBinaryCompare)
			sIDsForTable = aTemp(1)
			sKeyFieldName = "(" & TACO_PREFIX & "Tasks.LabelID=" & TACO_PREFIX & "Labels.LabelID) And (ProjectID=" & aTemp(0) & ") And TaskID"
			sFieldName = "LabelArticle, LabelName"
		Case "TACO_Projects"
			sKeyFieldName = "ProjectID"
			sFieldName = "ProjectName"
		Case "TACO_ProjectEasy"
			sTableName = TACO_PREFIX & "Projects"
			sKeyFieldName = "ProjectID"
			sFieldName = "EasyMode"
		Case "TACO_ProjectFile"
			sTableName = TACO_PREFIX & "Projects"
			sKeyFieldName = "ProjectID"
			sFieldName = "ProjectFile"
		Case "TACO_Status"
			sKeyFieldName = "StatusID"
			sFieldName = "StatusName"
		Case "TACO_Tasks"
			aTemp = Split(sIDsForTable, ",", 2, vbBinaryCompare)
			sIDsForTable = aTemp(1)
			sKeyFieldName = "(ProjectID=" & aTemp(0) & ") And TaskID"
			sFieldName = "TaskName"
		Case "TACO_TasksFullName"
			sTableName = TACO_PREFIX & "Tasks, " & TACO_PREFIX & "Labels"
			aTemp = Split(sIDsForTable, ",", 2, vbBinaryCompare)
			sIDsForTable = aTemp(1)
			sKeyFieldName = "(" & TACO_PREFIX & "Tasks.LabelID=" & TACO_PREFIX & "Labels.LabelID) And (ProjectID=" & aTemp(0) & ") And TaskID"
			sFieldName = "LabelArticle, LabelName, TaskName"
		Case "TACO_Variables"
			sKeyFieldName = "VariableID"
			sFieldName = "VariableName"

		Case "UserProfiles"
			sKeyFieldName = "ProfileID"
			sFieldName = "ProfileName"
		Case "Users"
			sKeyFieldName = "UserID"
			sFieldName = "UserName, UserLastName"
		Case "UserAccessKey"
			sTableName = "Users"
			sKeyFieldName = "UserAccessKey"
			sFieldName = "UserID"
		Case "UsersEmail"
			sTableName = "Users "
			sKeyFieldName = "UserID"
			sFieldName = "UserEmail"
		Case "WorkingCenters"
			sKeyFieldName = "WorkingCenterID"
			sFieldName = "WorkingCenterShortName, WorkingCenterName"
		Case "Zones"
			sTableName = "Zones "
			sKeyFieldName = "ZoneID"
			sFieldName = "ZoneName"
		Case "FullZones"
			sTableName = "Zones "
			If InStr(1, sIDs, ",", vbBinaryCompare) = 0 Then
				sKeyFieldName = "ZoneID"
			Else
				sKeyFieldName = "ZonePath"
			End If
			sFieldName = "ZoneCode, ZoneName"
		Case "ParentZones"
			sTableName = "Zones, Zones As ParentZones "
			sKeyFieldName = "(Zones.ParentID=ParentZones.ZoneID) And ParentZones.ZoneID"
			sFieldName = "ParentZones.ZoneName"
		Case "ParentZoneIDs"
			sTableName = "Zones"
			sKeyFieldName = "ZoneID"
			sFieldName = "ParentID"
		Case "ZoneTypes"
			sKeyFieldName = "ZoneTypeID"
			sFieldName = "ZoneTypeName"
	End Select

	sErrorDescription = "No se pudo obtener el campo de la tabla especificada."
	If StrComp(sTableName, "CreditsFiles", vbTextCompare) = 0 Then 
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select TOP 1" & sFieldName & " From Credits Where " & sKeyFieldName & " In (" & sIDsForTable & ") Order By " & sFieldName, "QueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & sFieldName & " From " & sTableName & " Where " & sKeyFieldName & " In (" & sIDsForTable & ") Order By " & sFieldName, "QueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If
	sNames = ""
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				sNames = sNames & sTab
				For iIndex = 0 To oRecordset.Fields.Count - 1
					sTemp = ""
					sTemp = CStr(oRecordset.Fields(iIndex).Value)
					Err.Clear
					If Len(sTemp) > 0 Then sNames = sNames & sTemp & " "
				Next
				sNames = Left(sNames, (Len(sNames) - Len(" ")))
				sNames = sNames & sSeparator
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			sNames = Left(sNames, (Len(sNames) - Len(sSeparator)))
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GetNameFromTable = lErrorNumber
	Err.Clear
End Function

Function GetNameFromTableByShortName(oADODBConnection, sTableName, sIDs, sTab, sSeparator, sNames, sErrorDescription)
'************************************************************
'Purpose: To get the name of the records given the IDs
'Inputs:  oADODBConnection, sTableName, sIDs, sTab, sSeparator
'Outputs: sNames, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetNameFromTableByShortName"
	Dim sKeyFieldName
	Dim sFieldName
	Dim sIDsForTable
	Dim iIndex
	Dim aTemp
	Dim sTemp
	Dim oRecordset
	Dim lErrorNumber

	sIDsForTable = sIDs
	If Len(sIDsForTable) = 0 Then sIDsForTable = "-2"
	Select Case sTableName
		Case "Absences"
			sKeyFieldName = "AbsenceShortName"
			sFieldName = "AbsenceShortName, AbsenceName"
		Case "CreditTypes"
			sKeyFieldName = "CreditTypeShortName"
			sFieldName = "CreditTypeShortName, CreditTypeName"
		Case "ExternalSpecialJourneys"
			sKeyFieldName = "RFC"
			sFieldName = "ExternalID"
	End Select

	sErrorDescription = "No se pudo obtener el campo de la tabla especificada."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select " & sFieldName & " From " & sTableName & " Where " & sKeyFieldName & " In ('" & sIDsForTable & "') Order By " & sFieldName, "QueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	sNames = ""
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				sNames = sNames & sTab
				For iIndex = 0 To oRecordset.Fields.Count - 1
					sTemp = ""
					sTemp = CStr(oRecordset.Fields(iIndex).Value)
					Err.Clear
					If Len(sTemp) > 0 Then sNames = sNames & sTemp & " "
				Next
				sNames = Left(sNames, (Len(sNames) - Len(" ")))
				sNames = sNames & sSeparator
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			sNames = Left(sNames, (Len(sNames) - Len(sSeparator)))
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GetNameFromTableByShortName = lErrorNumber
	Err.Clear
End Function

Function GetNewIDFromTable(oADODBConnection, sTableName, sIDField, sCondition, iDefaultValue, iNewID, sErrorDescription)
'************************************************************
'Purpose: To get a new ID from the specified table
'Inputs:  oADODBConnection, sTableName, sIDField, sCondition, iDefaultValue, iNewID, sErrorDescription
'Outputs: iNewID, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetNewIDFromTable"
	Dim oRecordset
	Dim lErrorNumber

	If Len(sCondition) > 0 Then
		sCondition = Trim(sCondition)
		If StrComp(sCondition, "Where ", vbTextCompare) <> 1 Then
			sCondition = "Where " & sCondition
		End If
	End If
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Max(" & sIDField & ") From " & sTableName & " " & sCondition, "QueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If IsNull(oRecordset.Fields(0).Value) Then
				iNewID = iDefaultValue
			Else
				iNewID = CLng(oRecordset.Fields(0).Value) + 1
			End If
		Else
			lErrorNumber = -1
			sErrorDescription = "No se pudo obtener un identificador para el nuevo registro."
			If Len(Err.description) > 0 Then
				sErrorDescription = sErrorDescription & "<BR /><B>Error del servidor Web: </B>" & Err.description
			End If
			Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "QueriesLib.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	GetNewIDFromTable = lErrorNumber
	Err.Clear
End Function

Function GetPayrollNumber(lPayrollID)
'************************************************************
'Purpose: To get the number of the payroll given the month and day
'Inputs:  lPayrollID
'Outputs: The number for the payroll in the year
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPayrollNumber"
	Dim iMonthDay

	iMonthDay = CInt(Right(CStr(lPayrollID), Len("0000")))
	GetPayrollNumber = ((Int(iMonthDay / 100) - 1) * 2) + 1
	If CInt(Right(lPayrollID, Len("00"))) > 15 Then GetPayrollNumber = GetPayrollNumber + 1

	Err.Clear
End Function

Function GetNumberForPayroll(lPayrollID)
'************************************************************
'Purpose: To transform a Payroll date from the YYYYMMDD format
'         to the YYYYQQ format
'Inputs:  lPayrollID
'Outputs: The number for the payroll
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetNumberForPayroll"
	Dim iYear
	Dim iMonth
	Dim iDay
	Dim iNumber

	iYear = CInt(Left(lPayrollID, Len("0000")))
	iMonth = CInt(Mid(lPayrollID, Len("00000"), Len("00")))
	iDay = CInt(Right(lPayrollID, Len("00")))
	iNumber = iMonth * 2
	If iDay <= 15 Then
		iNumber = iNumber - 1
	End If

	iMonth = CInt(iMonth + 0.1)
	GetNumberForPayroll = (iYear * 100) + iNumber

	Err.Clear
End Function

Function GetPayrollFromNumber(lPayrollNumber)
'************************************************************
'Purpose: To transform a Payroll date from the YYYYQQ format
'         to the YYYYMMDD format
'Inputs:  lPayrollNumber
'Outputs: The payroll in numeric format
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPayrollNumber"
	Dim iYear
	Dim iMonth
	Dim iDay

	iYear = CInt(Left(lPayrollNumber, Len("0000")))
	iMonth = CInt(Right(CStr(lPayrollNumber), Len("00"))) / 2
	If InStr(1, iMonth, ".", vbBinaryCompare) > 0 Then
		iDay = 15
	Else
		Select Case CInt(iMonth + 0.1)
			Case 1, 3, 5, 7, 8, 10, 12
				iDay = 31
			Case 2
				iDay = 28
				If (iYear Mod 4) = 0 Then iDay = 29
			Case Else
				iDay = 30
		End Select
	End If
	iMonth = CInt(iMonth + 0.1)
	GetPayrollFromNumber = (iYear * 10000) + (iMonth * 100) + iDay

	Err.Clear
End Function

Function GetPeriodsForPayroll(lPayrollID, lForPayrollDate, lStartDate)
'************************************************************
'Purpose: To get the IDs for the periods that apply to the
'         given payroll
'Inputs:  lPayrollID, lForPayrollDate, lStartDate
'Outputs: The IDs for the periods
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPeriodsForPayroll"
	Dim oRecordset
	Dim iDay
	Dim iMonth
	Dim iYear
	Dim iDiff
	Dim asDates
	Dim iIndex
	Dim lErrorNumber

	GetPeriodsForPayroll = "3,12,13,"
	iDay = CInt(Right(lForPayrollDate, 2))
	iMonth = CInt(Mid(lForPayrollDate, 5, 2))
	iYear = CInt(Left(lForPayrollDate, 4))
	If iDay > 15 Then GetPeriodsForPayroll = GetPeriodsForPayroll & "4,"
	If ((iMonth Mod 2) = 0) And (iDay > 15) Then GetPeriodsForPayroll = GetPeriodsForPayroll & "5,"
	If ((iMonth Mod 3) = 0) And (iDay > 15) Then GetPeriodsForPayroll = GetPeriodsForPayroll & "6,"
	If ((iMonth Mod 6) = 0) And (iDay > 15) Then GetPeriodsForPayroll = GetPeriodsForPayroll & "7,"
	If (iMonth = 12) And (iDay > 15) Then GetPeriodsForPayroll = GetPeriodsForPayroll & "8,"
	If lStartDate = -1 Then
		If ((iYear Mod 2) = 0) And (iMonth = 12) And (iDay > 15) Then GetPeriodsForPayroll = GetPeriodsForPayroll & "9,"
		If ((iYear Mod 5) = 0) And (iMonth = 12) And (iDay > 15) Then GetPeriodsForPayroll = GetPeriodsForPayroll & "10,"
	Else
		iDiff = Int((lStartDate - lForPayrollDate) / 10000)
		If ((iDiff Mod 2) = 0) And (iMonth = 12) And (iDay > 15) Then GetPeriodsForPayroll = GetPeriodsForPayroll & "9,"
		If ((iDiff Mod 5) = 0) And (iMonth = 12) And (iDay > 15) Then GetPeriodsForPayroll = GetPeriodsForPayroll & "10,"
	End If

	sErrorDescription = "No se pudieron obtener la periodicidad de los conceptos de pago."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PeriodID, PeriodDate, bSpecial From Periods Where (PeriodDate<>'-1') Order By PeriodID", "QueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			asDates = Split(CStr(oRecordset.Fields("PeriodDate").Value), ",")
			For iIndex = 0 To UBound(asDates)
				If InStr(1, CStr(lForPayrollDate), asDates(iIndex), vbBinaryCompare) > 0 Then GetPeriodsForPayroll = GetPeriodsForPayroll & CStr(oRecordset.Fields("PeriodID").Value) & ","
				If CInt(oRecordset.Fields("bSpecial").Value) = 1 Then
					If InStr(1, CStr(lPayrollID), asDates(iIndex), vbBinaryCompare) > 0 Then GetPeriodsForPayroll = GetPeriodsForPayroll & CStr(oRecordset.Fields("PeriodID").Value) & ","
				End If
			Next
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
	End If

	If Len(GetPeriodsForPayroll) > 0 Then GetPeriodsForPayroll = GetPeriodsForPayroll & "0"
	Set oRecordset = Nothing
	Err.Clear
End Function

Function GetSpecialHours(oADODBConnection, lPayrollID, lEmployeeID, iConceptQttyID, sCondition, dHours, sErrorDescription)
'************************************************************
'Purpose: To get the number of hours registered for the employee
'Inputs:  oADODBConnection, lPayrollID, lEmployeeID, iConceptQttyID, sCondition
'Outputs: dHours, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetSpecialHours"
	Dim oRecordset
	Dim lErrorNumber

	dHours = 0
	If Len(sCondition) > 0 Then
		sCondition = Trim(sCondition)
		If InStr(1, sCondition, "And ", vbBinaryCompare) <> 1 Then sCondition = "And " & sCondition
	End If
	sErrorDescription = "No se pudieron obtener las horas registradas para el empleado."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(AbsenceHours) As TotalHours From EmployeesAbsencesLKP, Absences Where (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And (EmployeeID=" & lEmployeeID & ") And (AppliedDate In (0," & lPayrollID & ")) And (Removed=0) " & sCondition, "QueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			dHours = CDbl(oRecordset.Fields("TotalHours").Value)
		End If
	End If

	Set oRecordset = Nothing
	GetSpecialHours = lErrorNumber
	Err.Clear
End Function

Function IsPayrollClosed(oADODBConnection, lPayrollID, sCondition, bPayrollIsClosed, sErrorDescription)
'************************************************************
'Purpose: To update the consecutive ID for the given record
'Inputs:  oADODBConnection, lTypeID, lConsecutiveID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "IsPayrollClosed"
	Dim oRecordset
	Dim lErrorNumber

	bPayrollIsClosed = False
	sErrorDescription = "No se pudo obtener la información de la nómina."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Payrolls Where (PayrollID=" & lPayrollID & ") And (IsClosed=1)", "QueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		bPayrollIsClosed = ((Not oRecordset.EOF) Or (InStr(1, sCondition, "Payments.", vbBinaryCompare) > 0))
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	IsPayrollClosed = lErrorNumber
	Err.Clear
End Function

Function UpdateConsecutiveID(oADODBConnection, lTypeID, lConsecutiveID, sErrorDescription)
'************************************************************
'Purpose: To update the consecutive ID for the given record
'Inputs:  oADODBConnection, lTypeID, lConsecutiveID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "UpdateConsecutiveID"
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener el siguiente número consecutivo."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update ConsecutiveIDs Set CurrentID=" & lConsecutiveID & "  Where (IDType=" & lTypeID & ")", "ReportsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

	UpdateConsecutiveID = lErrorNumber
	Err.Clear
End Function

Function VerifyExistenceOfRecordInDatabase(oADODBConnection, sTableName, sIDField, sLimitTypes, sValueField, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To verify if a record exist in determinated database table
'Inputs:  oADODBConnection, aEmployeeComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "VerifyExistenceOfRecordInDatabase"
	Dim lErrorNumber
	Dim sQuery
	Dim asField
	Dim asValue
	Dim asLimitTypes
	Dim iIndex
	Dim sQueryFieldCondition

	asField = Split(sIDField, ",", -1, vbBinaryCompare)
	asValue = Split(sValueField, ",", -1, vbBinaryCompare)
	asLimitTypes = Split(sLimitTypes, ",", -1, vbBinaryCompare)

	If (UBound(asField) = UBound(asValue)) And (UBound(asField) = UBound(asLimitTypes)) Then
		For iIndex = 0 To UBound(asField)
			Select Case CInt(asLimitTypes(iIndex))
				Case N_OPEN_MINIMUM
					sQueryFieldCondition = sQueryFieldCondition & " And (" & asField(iIndex) & "<='" & asValue(iIndex) & "')"
				Case N_OPEN_MAXIMUM
					sQueryFieldCondition = sQueryFieldCondition & " And (" & asField(iIndex) & ">='" & asValue(iIndex) & "')"
				Case N_CLOSED_MINIMUM
					sQueryFieldCondition = sQueryFieldCondition & " And (" & asField(iIndex) & "<'" & asValue(iIndex) & "')"
				Case N_CLOSED_MAXIMUM
					sQueryFieldCondition = sQueryFieldCondition & " And (" & asField(iIndex) & ">'" & asValue(iIndex) & "')"
				Case N_NONE
					sQueryFieldCondition = sQueryFieldCondition & " And (" & asField(iIndex) & "='" & asValue(iIndex) & "')"
			End Select
		Next
	Else
		lErrorNumber = -1
		sErrorDescription = "No se pudo verificar si el registro esta activo"
	End If

	sQueryFieldCondition = Right(sQueryFieldCondition, (Len(sQueryFieldCondition) - Len(" And")))
	sQuery = "Select * From " & sTableName & " Where " & sQueryFieldCondition

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "QueriesLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	If lErrorNumber = 0 Then
		VerifyExistenceOfRecordInDatabase = (Not oRecordset.EOF)
	Else
		sErrorDescription = "Error al verificar si el registro esta activo."
		VerifyExistenceOfRecordInDatabase = False
	End If
	Err.Clear
End Function
%>