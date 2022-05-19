<%
Const N_ID_PAYROLL = 0
Const S_NAME_PAYROLL = 1
Const N_DATE_PAYROLL = 2
Const S_CLC_PAYROLL = 3
Const N_TYPE_ID_PAYROLL = 4
Const N_FOR_DATE_PAYROLL = 5
Const N_IS_ACTIVE_1_PAYROLL = 6
Const N_IS_ACTIVE_2_PAYROLL = 7
Const N_IS_ACTIVE_3_PAYROLL = 8
Const N_IS_ACTIVE_4_PAYROLL = 9
Const N_IS_ACTIVE_5_PAYROLL = 10
Const N_IS_ACTIVE_6_PAYROLL = 11
Const N_IS_ACTIVE_7_PAYROLL = 12
Const N_IS_ACTIVE_8_PAYROLL = 13
Const N_IS_ACTIVE_9_PAYROLL = 14
Const N_IS_ACTIVE_10_PAYROLL = 15
Const N_IS_ACTIVE_11_PAYROLL = 16
Const N_IS_ACTIVE_12_PAYROLL = 17
Const N_CLOSED_PAYROLL = 18
Const S_QUERY_CONDITION_PAYROLL = 19
Const B_CHECK_FOR_DUPLICATED_PAYROLL = 20
Const B_IS_DUPLICATED_PAYROLL = 21
Const B_COMPONENT_INITIALIZED_PAYROLL = 22

Const N_PAYROLL_COMPONENT_SIZE = 22

Dim aPayrollComponent()
Redim aPayrollComponent(N_PAYROLL_COMPONENT_SIZE)

Function InitializePayrollComponent(oRequest, aPayrollComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Payroll Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aPayrollComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializePayrollComponent"
	Redim Preserve aPayrollComponent(N_PAYROLL_COMPONENT_SIZE)

	If IsEmpty(aPayrollComponent(N_ID_PAYROLL)) Then
		If Len(oRequest("PayrollID").Item) > 0 Then
			aPayrollComponent(N_ID_PAYROLL) = CLng(oRequest("PayrollID").Item)
		Else
			aPayrollComponent(N_ID_PAYROLL) = -1
		End If
	End If

	If IsEmpty(aPayrollComponent(S_NAME_PAYROLL)) Then
		If Len(oRequest("PayrollName").Item) > 0 Then
			aPayrollComponent(S_NAME_PAYROLL) = oRequest("PayrollName").Item
		Else
			aPayrollComponent(S_NAME_PAYROLL) = ""
		End If
	End If
	aPayrollComponent(S_NAME_PAYROLL) = Left(aPayrollComponent(S_NAME_PAYROLL), 255)

	If IsEmpty(aPayrollComponent(N_DATE_PAYROLL)) Then
		If Len(oRequest("PayrollYear").Item) > 0 Then
			aPayrollComponent(N_DATE_PAYROLL) = CLng(oRequest("PayrollYear").Item & Right(("0" & oRequest("PayrollMonth").Item), Len("00")) & Right(("0" & oRequest("PayrollDay").Item), Len("00")))
		ElseIf Len(oRequest("PayrollDate").Item) > 0 Then
			aPayrollComponent(N_DATE_PAYROLL) = CLng(oRequest("PayrollDate").Item)
		Else
			aPayrollComponent(N_DATE_PAYROLL) = 0
		End If
	End If
	If (aPayrollComponent(N_ID_PAYROLL) = -1) And (aPayrollComponent(N_DATE_PAYROLL) > 0) Then aPayrollComponent(N_ID_PAYROLL) = aPayrollComponent(N_DATE_PAYROLL)

	If IsEmpty(aPayrollComponent(S_CLC_PAYROLL)) Then
		If Len(oRequest("PayrollCLC").Item) > 0 Then
			aPayrollComponent(S_CLC_PAYROLL) = oRequest("PayrollCLC").Item
		Else
			aPayrollComponent(S_CLC_PAYROLL) = ""
		End If
	End If
	aPayrollComponent(S_CLC_PAYROLL) = Left(aPayrollComponent(S_CLC_PAYROLL), 20)

	If IsEmpty(aPayrollComponent(N_TYPE_ID_PAYROLL)) Then
		If Len(oRequest("PayrollTypeID").Item) > 0 Then
			aPayrollComponent(N_TYPE_ID_PAYROLL) = CLng(oRequest("PayrollTypeID").Item)
		Else
			aPayrollComponent(N_TYPE_ID_PAYROLL) = 1
		End If
	End If

	If IsEmpty(aPayrollComponent(N_FOR_DATE_PAYROLL)) Then
		If aPayrollComponent(N_TYPE_ID_PAYROLL) = 1 Then
			aPayrollComponent(N_FOR_DATE_PAYROLL) = aPayrollComponent(N_ID_PAYROLL)
		ElseIf Len(oRequest("ForPayrollDate").Item) > 0 Then
			aPayrollComponent(N_FOR_DATE_PAYROLL) = CLng(oRequest("ForPayrollDate").Item)
		Else
			aPayrollComponent(N_FOR_DATE_PAYROLL) = 0
		End If
	End If

	If IsEmpty(aPayrollComponent(N_CLOSED_PAYROLL)) Then
		If Len(oRequest("IsClosed").Item) > 0 Then
			aPayrollComponent(N_CLOSED_PAYROLL) = CInt(oRequest("IsClosed").Item)
		Else
			aPayrollComponent(N_CLOSED_PAYROLL) = 0
		End If
	End If

	If IsEmpty(aPayrollComponent(N_IS_ACTIVE_1_PAYROLL)) Then
		If Len(oRequest("IsActive_1").Item) > 0 Then
			aPayrollComponent(N_IS_ACTIVE_1_PAYROLL) = CInt(oRequest("IsActive_1").Item)
		Else
			aPayrollComponent(N_IS_ACTIVE_1_PAYROLL) = 1
		End If
	End If

	If IsEmpty(aPayrollComponent(N_IS_ACTIVE_2_PAYROLL)) Then
		If Len(oRequest("IsActive_2").Item) > 0 Then
			aPayrollComponent(N_IS_ACTIVE_2_PAYROLL) = CInt(oRequest("IsActive_2").Item)
		Else
			aPayrollComponent(N_IS_ACTIVE_2_PAYROLL) = 1
		End If
	End If

	If IsEmpty(aPayrollComponent(N_IS_ACTIVE_3_PAYROLL)) Then
		If Len(oRequest("IsActive_3").Item) > 0 Then
			aPayrollComponent(N_IS_ACTIVE_3_PAYROLL) = CInt(oRequest("IsActive_3").Item)
		Else
			aPayrollComponent(N_IS_ACTIVE_3_PAYROLL) = 1
		End If
	End If

	If IsEmpty(aPayrollComponent(N_IS_ACTIVE_4_PAYROLL)) Then
		If Len(oRequest("IsActive_4").Item) > 0 Then
			aPayrollComponent(N_IS_ACTIVE_4_PAYROLL) = CInt(oRequest("IsActive_4").Item)
		Else
			aPayrollComponent(N_IS_ACTIVE_4_PAYROLL) = 1
		End If
	End If

	If IsEmpty(aPayrollComponent(N_IS_ACTIVE_5_PAYROLL)) Then
		If Len(oRequest("IsActive_5").Item) > 0 Then
			aPayrollComponent(N_IS_ACTIVE_5_PAYROLL) = CInt(oRequest("IsActive_5").Item)
		Else
			aPayrollComponent(N_IS_ACTIVE_5_PAYROLL) = 1
		End If
	End If

	If IsEmpty(aPayrollComponent(N_IS_ACTIVE_6_PAYROLL)) Then
		If Len(oRequest("IsActive_6").Item) > 0 Then
			aPayrollComponent(N_IS_ACTIVE_6_PAYROLL) = CInt(oRequest("IsActive_6").Item)
		Else
			aPayrollComponent(N_IS_ACTIVE_6_PAYROLL) = 1
		End If
	End If

	If IsEmpty(aPayrollComponent(N_IS_ACTIVE_7_PAYROLL)) Then
		If Len(oRequest("IsActive_7").Item) > 0 Then
			aPayrollComponent(N_IS_ACTIVE_7_PAYROLL) = CInt(oRequest("IsActive_7").Item)
		Else
			aPayrollComponent(N_IS_ACTIVE_7_PAYROLL) = 1
		End If
	End If

	If IsEmpty(aPayrollComponent(N_IS_ACTIVE_8_PAYROLL)) Then
		If Len(oRequest("IsActive_8").Item) > 0 Then
			aPayrollComponent(N_IS_ACTIVE_8_PAYROLL) = CInt(oRequest("IsActive_8").Item)
		Else
			aPayrollComponent(N_IS_ACTIVE_8_PAYROLL) = 1
		End If
	End If

	If IsEmpty(aPayrollComponent(N_IS_ACTIVE_9_PAYROLL)) Then
		If Len(oRequest("IsActive_9").Item) > 0 Then
			aPayrollComponent(N_IS_ACTIVE_9_PAYROLL) = CInt(oRequest("IsActive_9").Item)
		Else
			aPayrollComponent(N_IS_ACTIVE_9_PAYROLL) = 1
		End If
	End If

	If IsEmpty(aPayrollComponent(N_IS_ACTIVE_10_PAYROLL)) Then
		If Len(oRequest("IsActive_10").Item) > 0 Then
			aPayrollComponent(N_IS_ACTIVE_10_PAYROLL) = CInt(oRequest("IsActive_10").Item)
		Else
			aPayrollComponent(N_IS_ACTIVE_10_PAYROLL) = 1
		End If
	End If

	If IsEmpty(aPayrollComponent(N_IS_ACTIVE_11_PAYROLL)) Then
		If Len(oRequest("IsActive_11").Item) > 0 Then
			aPayrollComponent(N_IS_ACTIVE_11_PAYROLL) = CInt(oRequest("IsActive_11").Item)
		Else
			aPayrollComponent(N_IS_ACTIVE_11_PAYROLL) = 1
		End If
	End If

	If IsEmpty(aPayrollComponent(N_IS_ACTIVE_12_PAYROLL)) Then
		If Len(oRequest("IsActive_12").Item) > 0 Then
			aPayrollComponent(N_IS_ACTIVE_12_PAYROLL) = CInt(oRequest("IsActive_12").Item)
		Else
			aPayrollComponent(N_IS_ACTIVE_12_PAYROLL) = 1
		End If
	End If

	aPayrollComponent(S_QUERY_CONDITION_PAYROLL) = ""
	aPayrollComponent(B_CHECK_FOR_DUPLICATED_PAYROLL) = True
	aPayrollComponent(B_IS_DUPLICATED_PAYROLL) = False

	aPayrollComponent(B_COMPONENT_INITIALIZED_PAYROLL) = True
	InitializePayrollComponent = Err.number
	Err.Clear
End Function

Function AddPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new payroll into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPayrollComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddPayroll"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPayrollComponent(B_COMPONENT_INITIALIZED_PAYROLL)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollComponent(oRequest, aPayrollComponent)
	End If

	If aPayrollComponent(N_ID_PAYROLL) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para la nueva nómina."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "Payrolls", "PayrollID", "", 1, aPayrollComponent(N_ID_PAYROLL), sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		If aPayrollComponent(N_TYPE_ID_PAYROLL) = 0 Then aPayrollComponent(N_ID_PAYROLL) = aPayrollComponent(N_ID_PAYROLL) & "0"
		If aPayrollComponent(B_CHECK_FOR_DUPLICATED_PAYROLL) Then
			lErrorNumber = CheckExistencyOfPayroll(oADODBConnection, False, aPayrollComponent, sErrorDescription)
			If aPayrollComponent(B_IS_DUPLICATED_PAYROLL) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe una nómina con el nombre '" & CleanStringForHTML(aPayrollComponent(S_NAME_PAYROLL)) & "' y/o para el " & DisplayDateFromSerialNumber(aPayrollComponent(N_DATE_PAYROLL), -1, -1, -1) & " registrada en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PayrollComponentConstants.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			End If
		End If

		If lErrorNumber = 0 Then
			If Not CheckPayrollInformationConsistency(aPayrollComponent, sErrorDescription) Then
				lErrorNumber = -1
			Else
				sErrorDescription = "No se pudo guardar la información de la nueva nómina."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payrolls (PayrollID, PayrollName, PayrollDate, PayrollNumber, PayrollCLC, PayrollTypeID, ForPayrollDate, IsActive_1, IsActive_2, IsActive_3, IsActive_4, IsActive_5, IsActive_6, IsActive_7, IsActive_8, IsActive_9, IsActive_10, IsActive_11, IsActive_12, IsClosed) Values (" & aPayrollComponent(N_ID_PAYROLL) & ", '" & Replace(aPayrollComponent(S_NAME_PAYROLL), "'", "") & "', " & aPayrollComponent(N_DATE_PAYROLL) & ", " & GetNumberForPayroll(aPayrollComponent(N_ID_PAYROLL)) & ", '" & Replace(aPayrollComponent(S_CLC_PAYROLL), "'", "") & "', " & aPayrollComponent(N_TYPE_ID_PAYROLL) & ", " & aPayrollComponent(N_FOR_DATE_PAYROLL) & ", 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, " & aPayrollComponent(N_CLOSED_PAYROLL) & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudo guardar la información de la nueva nómina."
					Select Case iConnectionType
						Case ORACLE
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "CREATE TABLE Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate int NOT NULL, RecordID int NOT NULL, EmployeeID INTEGER NOT NULL, ConceptID INTEGER NOT NULL, PayrollTypeID INTEGER NOT NULL, ConceptAmount decimal(18,2) NOT NULL, ConceptTaxes decimal(18,2) NOT NULL, ConceptRetention decimal(18,2) NOT NULL, UserID INTEGER NOT NULL, PRIMARY KEY (RecordDate, RecordID, EmployeeID, ConceptID))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						Case SQL_SERVER, SQL_SERVER_64_OLE, SQL_SERVER_64_DSNLess, SQL_SERVER_2008
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "CREATE TABLE Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate int NOT NULL, RecordID int NOT NULL, EmployeeID int NOT NULL, ConceptID int NOT NULL, PayrollTypeID int NOT NULL, ConceptAmount decimal(18,2) NOT NULL, ConceptTaxes decimal(18,2) NOT NULL, ConceptRetention decimal(18,2) NOT NULL, UserID int NOT NULL, PRIMARY KEY (RecordDate, RecordID, EmployeeID, ConceptID))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						Case Else
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "CREATE TABLE Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate int NOT NULL, RecordID int NOT NULL, EmployeeID int NOT NULL, ConceptID int NOT NULL, PayrollTypeID int NOT NULL, ConceptAmount float NOT NULL, ConceptTaxes float NOT NULL, ConceptRetention float NOT NULL, UserID int NOT NULL, PRIMARY KEY (RecordDate, RecordID, EmployeeID, ConceptID))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End Select
				End If
			End If
		End If
	End If

	AddPayroll = lErrorNumber
	Err.Clear
End Function

Function ActivateEmployeeTaxAdjustment(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To activate the tax adjustment for the given user
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ActivateEmployeeTaxAdjustment"
	Dim iIndex
	Dim asEmployeeIDs
	Dim lErrorNumber

	'asEmployeeIDs = oRequest("EmployeeID").Item
	asEmployeeIDs = Replace(Replace(Replace(oRequest("EmployeeIDs").Item, vbNewLine, ","), " ", ""), ",,", ",")
	Do While (InStr(1, asEmployeeIDs, ",,", vbBinaryCompare) > 0)
		asEmployeeIDs = Replace(asEmployeeIDs, ",,", ",")
	Loop
	sErrorDescription = "No se pudo actualizar la información del ajuste anual."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesForTaxAdjustment Set bTaxAdjustment=" & oRequest("TaxActivation").Item & " Where (EmployeeID In (" & asEmployeeIDs & ")) And (PayrollYear=" & oRequest("YearID").Item & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	asEmployeeIDs = Split(asEmployeeIDs, ",")
	For iIndex = 0 To UBound(asEmployeeIDs)
		If Len(asEmployeeIDs(iIndex)) > 0 Then
			sErrorDescription = "No se pudo actualizar la información del ajuste anual."
			lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into EmployeesForTaxAdjustment (EmployeeID, PayrollYear, bTaxAdjustment) Values (" & asEmployeeIDs(iIndex) & ", " & oRequest("YearID").Item & ", " & oRequest("TaxActivation").Item & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
		End If
	Next

	ActivateEmployeeTaxAdjustment = lErrorNumber
	Err.Clear
End Function

Function ApplyEmployeeTaxAdjustment(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To apply the given amount to the given user
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ApplyEmployeeTaxAdjustment"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudo actualizar la información del ajuste anual."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesForTaxAdjustment Where (EmployeeID=" & oRequest("EmployeeID").Item & ") And (PayrollYear=" & oRequest("YearID").Item & ") And (bTaxAdjustment=0)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	If lErrorNumber = 0 Then
		If oRecordset.EOF Then
			oRecordset.Close
			sErrorDescription = "No se pudo actualizar la información del ajuste anual."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & oRequest("YearID").Item & " Where (RecordDate=" & oRequest("YearID").Item & "9999) And (EmployeeID=" & oRequest("EmployeeID").Item & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo actualizar la información del ajuste anual."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & oRequest("YearID").Item & "1231 Where (RecordID=99) And (EmployeeID=" & oRequest("EmployeeID").Item & ") And (ConceptID=55)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo guardar la información del ajuste anual."
				lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into Payroll_" & oRequest("YearID").Item & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & oRequest("YearID").Item & "9999, 1, " & oRequest("EmployeeID").Item & ", 55, 1, " & oRequest("TaxAmount").Item & ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
			End If
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudo guardar la información del ajuste anual."
				lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into Payroll_" & oRequest("YearID").Item & "1231 (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & oRequest("YearID").Item & "1231, 99, " & oRequest("EmployeeID").Item & ", 55, 1, " & oRequest("TaxAmount").Item & ", 0, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
			End If
		Else
			lErrorNumber = -1
			sErrorDescription = "El empleado especificado está registrado en la lista de empleados que no desean el ajuste anual del impuesto sobre la renta."
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	ApplyEmployeeTaxAdjustment = lErrorNumber
	Err.Clear
End Function

Function BuildCondition(sCondition, sQueryBegin)
'************************************************************
'Purpose: To build the condition for the payroll process
'Outputs: sCondition, sQueryBegin
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildCondition"

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
'	sCondition = Replace(sCondition, "(Areas.", "(EmployeesHistoryList.")
	sCondition = Replace(sCondition, "(Jobs.", "(EmployeesHistoryList.")
'	sCondition = sCondition & " And (RecordID In (2,3,4))"
	If InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0 Then sCondition = " And (Areas.ZoneID=Zones.ZoneID)" & sCondition
	If InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0 Then sCondition = " And (EmployeesHistoryList.AreaID=Areas.AreaID)" & sCondition
	If InStr(1, sCondition, "(Jobs.", vbBinaryCompare) > 0 Then sCondition = " And (EmployeesHistoryList.JobID=Jobs.JobID)" & sCondition

	sQueryBegin = ""
	If InStr(1, sCondition, "(Employees.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Employees"
	If InStr(1, sCondition, "(Jobs.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Jobs"
	If InStr(1, sCondition, "(Zones.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Zones"
	If InStr(1, sCondition, "(Areas.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Areas"
	If InStr(1, sCondition, "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
	If InStr(1, sCondition, "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
	If InStr(1, sCondition, "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"

	BuildCondition = Err.number
	Err.Clear
End Function

Function GetPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about a payroll from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPayrollComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPayroll"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPayrollComponent(B_COMPONENT_INITIALIZED_PAYROLL)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollComponent(oRequest, aPayrollComponent)
	End If

	If aPayrollComponent(N_ID_PAYROLL) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador de la nómina para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PayrollComponentConstants.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información de la nómina."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Payrolls Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El nómina especificada no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PayrollComponentConstants.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aPayrollComponent(S_NAME_PAYROLL) = CStr(oRecordset.Fields("PayrollName").Value)
				aPayrollComponent(N_DATE_PAYROLL) = CLng(oRecordset.Fields("PayrollDate").Value)
				aPayrollComponent(S_CLC_PAYROLL) = CStr(oRecordset.Fields("PayrollCLC").Value)
				aPayrollComponent(N_TYPE_ID_PAYROLL) = CLng(oRecordset.Fields("PayrollTypeID").Value)
				aPayrollComponent(N_FOR_DATE_PAYROLL) = CLng(oRecordset.Fields("ForPayrollDate").Value)
				aPayrollComponent(N_IS_ACTIVE_1_PAYROLL) = CInt(oRecordset.Fields("IsActive_1").Value)
				aPayrollComponent(N_IS_ACTIVE_2_PAYROLL) = CInt(oRecordset.Fields("IsActive_2").Value)
				aPayrollComponent(N_IS_ACTIVE_3_PAYROLL) = CInt(oRecordset.Fields("IsActive_3").Value)
				aPayrollComponent(N_IS_ACTIVE_4_PAYROLL) = CInt(oRecordset.Fields("IsActive_4").Value)
				aPayrollComponent(N_IS_ACTIVE_5_PAYROLL) = CInt(oRecordset.Fields("IsActive_5").Value)
				aPayrollComponent(N_IS_ACTIVE_6_PAYROLL) = CInt(oRecordset.Fields("IsActive_6").Value)
				aPayrollComponent(N_IS_ACTIVE_7_PAYROLL) = CInt(oRecordset.Fields("IsActive_7").Value)
				aPayrollComponent(N_IS_ACTIVE_8_PAYROLL) = CInt(oRecordset.Fields("IsActive_8").Value)
				aPayrollComponent(N_IS_ACTIVE_9_PAYROLL) = CInt(oRecordset.Fields("IsActive_9").Value)
				aPayrollComponent(N_IS_ACTIVE_10_PAYROLL) = CInt(oRecordset.Fields("IsActive_10").Value)
				aPayrollComponent(N_IS_ACTIVE_11_PAYROLL) = CInt(oRecordset.Fields("IsActive_11").Value)
				aPayrollComponent(N_IS_ACTIVE_12_PAYROLL) = CInt(oRecordset.Fields("IsActive_12").Value)
				aPayrollComponent(N_CLOSED_PAYROLL) = CInt(oRecordset.Fields("IsClosed").Value)
				oRecordset.Close
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetPayroll = lErrorNumber
	Err.Clear
End Function

Function GetPayrolls(oRequest, oADODBConnection, aPayrollComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the payrolls from
'		  the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPayrollComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPayrolls"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPayrollComponent(B_COMPONENT_INITIALIZED_PAYROLL)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollComponent(oRequest, aPayrollComponent)
	End If

	aPayrollComponent(S_QUERY_CONDITION_PAYROLL) = Trim(aPayrollComponent(S_QUERY_CONDITION_PAYROLL))
	If Len(aPayrollComponent(S_QUERY_CONDITION_PAYROLL)) > 0 Then
		If InStr(1, aPayrollComponent(S_QUERY_CONDITION_PAYROLL), "And ", vbBinaryCompare) <> 1 Then aPayrollComponent(S_QUERY_CONDITION_PAYROLL) = "And " & aPayrollComponent(S_QUERY_CONDITION_PAYROLL)
	End If

	sErrorDescription = "No se pudo obtener la información de las nóminas."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Payrolls Where (PayrollID > -1) " & aPayrollComponent(S_QUERY_CONDITION_PAYROLL) & " Order By PayrollDate Desc, PayrollName", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

	GetPayrolls = lErrorNumber
	Err.Clear
End Function

Function ModifyPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing payroll in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPayrollComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyPayroll"
	Dim sOldPassword
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPayrollComponent(B_COMPONENT_INITIALIZED_PAYROLL)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollComponent(oRequest, aPayrollComponent)
	End If

	If aPayrollComponent(N_ID_PAYROLL) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador de la nómina a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PayrollComponentConstants.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If Not CheckPayrollInformationConsistency(aPayrollComponent, sErrorDescription) Then
			lErrorNumber = -1
		Else
			sErrorDescription = "No se pudo modificar la información de la nómina."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payrolls Set PayrollName='" & Replace(aPayrollComponent(S_NAME_PAYROLL), "'", "") & "', PayrollDate=" & aPayrollComponent(N_DATE_PAYROLL) & ", PayrollCLC='" & Replace(aPayrollComponent(S_CLC_PAYROLL), "'", "") & "', IsActive_1=" & aPayrollComponent(N_IS_ACTIVE_1_PAYROLL) & ", IsActive_2=" & aPayrollComponent(N_IS_ACTIVE_2_PAYROLL) & ", IsActive_3=" & aPayrollComponent(N_IS_ACTIVE_3_PAYROLL) & ", IsActive_4=" & aPayrollComponent(N_IS_ACTIVE_4_PAYROLL) & ", IsActive_5=" & aPayrollComponent(N_IS_ACTIVE_5_PAYROLL) & ", IsActive_6=" & aPayrollComponent(N_IS_ACTIVE_6_PAYROLL) & ", IsActive_7=" & aPayrollComponent(N_IS_ACTIVE_7_PAYROLL) & ", IsActive_8=" & aPayrollComponent(N_IS_ACTIVE_8_PAYROLL) & ", IsActive_9=" & aPayrollComponent(N_IS_ACTIVE_9_PAYROLL) & ", IsActive_10=" & aPayrollComponent(N_IS_ACTIVE_10_PAYROLL) & ", IsActive_11=" & aPayrollComponent(N_IS_ACTIVE_11_PAYROLL) & ", IsActive_12=" & aPayrollComponent(N_IS_ACTIVE_12_PAYROLL) & ", IsClosed=" & aPayrollComponent(N_CLOSED_PAYROLL) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	ModifyPayroll = lErrorNumber
	Err.Clear
End Function

Function RemovePayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
'************************************************************
'Purpose: To remove a payroll from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPayrollComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemovePayroll"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPayrollComponent(B_COMPONENT_INITIALIZED_PAYROLL)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollComponent(oRequest, aPayrollComponent)
	End If

	If aPayrollComponent(N_ID_PAYROLL) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó la nómina a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PayrollComponentConstants.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información de la nómina."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payrolls Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudo eliminar la contraseña de la nómina."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Drop Table Payroll_" & aPayrollComponent(N_ID_PAYROLL), "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	RemovePayroll = lErrorNumber
	Err.Clear
End Function

Function CalculatePayrol1(oRequest, oADODBConnection, iLevel, aPayrollComponent, sErrorDescription)
'************************************************************
'Purpose: To calculate the payroll and save it into the database
'Inputs:  oRequest, oADODBConnection, iLevel
'Outputs: aPayrollComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CalculatePayrol1"
	Const ROWS_PER_FILE = 10000
	Dim lPayrollDate
	Dim asEmployeesQueries
	Dim iCounter
	Dim iCounter2
	Dim iPayrollIndex
	Dim iIndex
	Dim jIndex
	Dim sFilePath
	Dim asFileContents
	Dim sDate
	Dim oStartDate
	Dim oEndDate
	Dim sQueryBegin
	Dim sQueryEnd
	Dim sCondition
	Dim sConceptCondition
	Dim lCurrentID
	Dim lCurrentID2
	Dim lConceptID
	Dim sCurrentID
	Dim adTotal
	Dim dAmount
	Dim dTaxAmount
	Dim dTemp
	Dim sTemp
	Dim lFirstPayroll
	Dim lLastPayroll
	Dim asPayrolls
	Dim lEmployeeTypeID
	Dim adTaxes
	Dim adAllowances
	Dim adTaxInvertions
	Dim bTruncate
	Dim sTruncate
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bTruncate = False
	sTruncate = ""
	bComponentInitialized = aPayrollComponent(B_COMPONENT_INITIALIZED_PAYROLL)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollComponent(oRequest, aPayrollComponent)
	End If

	If Not bTimeout Then
		lErrorNumber = DoCalculations2(aPayrollComponent, (aPayrollComponent(N_TYPE_ID_PAYROLL) = 4), False, sErrorDescription)
	End If

	CalculatePayrol1 = lErrorNumber
	Err.Clear
End Function


Function DoCalculations2(aPayrollComponent, bRetroactive, bAdjustment, sErrorDescription)
'************************************************************
'Purpose: To calculate the payroll and save it into the database
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DoCalculations2"
	Const ROWS_PER_FILE = 10000
	Const CONCEPTS_FOR_FACTOR = "1,3,12,13,14,38,49,89"
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
	Dim bOptimiza
	Dim mCount
	Dim lTotal
	Dim bCount

	Dim iPayrollIndex
	Dim asPayrolls


	bCount = True
	sCondition = ""
	bTruncate = False
	sTruncate = ""
	bOptimiza = True
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

	If bCount Then Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "Registros de ConceptID=0 en Payroll: " & vbTab & GetPayrollConcept1Count(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")

		If bOptimiza And ((iConnectionType = ORACLE) Or (iConnectionType = SQL_SERVER)) Then

			Call DisplayTimeStamp("START: LEVEL 2, CREATE FILES EmployeesConceptsLKP, QttyID=1. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
			Call AppendTextToFile(sFilePath & "_Rastreo.txt", "START: LEVEL 2, CREATE FILES EmployeesConceptsLKP, QttyID=1. " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
			If Not bTimeout Then
				iCounter = 0
				sQueryBegin = ""
				If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, sCondition, "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = ", EmployeesRevisions"
				sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select " & aPayrollComponent(N_FOR_DATE_PAYROLL) & " As RecordDate, 1 As RecordID, EmployeesConceptsLKP.EmployeeID, Concepts.ConceptID, 1 As PayrollTypeID, ConceptAmount, 0 As ConceptTaxes, 0 As ConceptRetention, " & aLoginComponent(N_USER_ID_LOGIN) & " As UserID From EmployeesConceptsLKP, Concepts, EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Jobs, Areas, Zones " & sQueryBegin & " Where (EmployeesConceptsLKP.ConceptID=Concepts.ConceptID) And (EmployeesConceptsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.JobID=Jobs.JobID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EmployeesConceptsLKP.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesConceptsLKP.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") " & Replace(sCondition, "(Positions.", "(EmployeesHistoryList.") & sConceptCondition & " And (ConceptQttyID=1) And (Concepts.PeriodID In (" & sPeriods & ")) And (EmployeesConceptsLKP.Active=1)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select " & aPayrollComponent(N_FOR_DATE_PAYROLL) & " As RecordDate, 1 As RecordID, EmployeesConceptsLKP.EmployeeID, Concepts.ConceptID, 1 As PayrollTypeID, ConceptAmount, 0 As ConceptTaxes, 0 As ConceptRetention, " & aLoginComponent(N_USER_ID_LOGIN) & " As UserID From EmployeesConceptsLKP, Concepts, EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, Jobs, Areas, Zones " & sQueryBegin & " Where (EmployeesConceptsLKP.ConceptID=Concepts.ConceptID) And (EmployeesConceptsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.JobID=Jobs.JobID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EmployeesConceptsLKP.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesConceptsLKP.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") " & Replace(sCondition, "(Positions.", "(EmployeesHistoryList.") & sConceptCondition & " And (ConceptQttyID=1) And (Concepts.PeriodID In (" & sPeriods & ")) And (EmployeesConceptsLKP.Active=1)" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
				If bCount Then Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "Registros de ConceptID=0 en Payroll: " & vbTab & GetPayrollConcept1Count(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")
			End If
		
			lErrorNumber = GatherTableStats(oRequest, oADODBConnection, "Payroll_" & aPayrollComponent(N_ID_PAYROLL), 25, "for all indexed columns size auto", True, sErrorDescription)
			Call AppendTextToFile(sFilePath & "_Rastreo.txt", "GatherTableStats(Payroll_20130131)" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
			If bCount Then Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "Registros de ConceptID=0 en Payroll: " & vbTab & GetPayrollConcept1Count(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")

			If Not bTimeout Then
				If bOptimiza And ((iConnectionType = ORACLE) Or (iConnectionType = SQL_SERVER)) Then
				'If False Then
					lTotal = 0
					Set oADODBCommand = Server.CreateObject("ADODB.Command")
					Set oADODBCommand.ActiveConnection = oADODBConnection
					oADODBCommand.commandtype=4
					oADODBCommand.commandtext = "SIAP.CreateAndRunFiles"
					Set param = oADODBCommand.Parameters
					param.append oADODBCommand.createparameter("iFileType", 3, 1)
					param.append oADODBCommand.createparameter("lPayrollID", 3, 1)
					param.append oADODBCommand.createparameter("lForPayrollID", 3, 1)
					param.append oADODBCommand.createparameter("iPayrollType", 3, 1)
					param.append oADODBCommand.createparameter("sConditions", vbString, 1)
					param.append oADODBCommand.createparameter("iInsertCounts", 3, 2)

					oADODBCommand("iFileType") = 1
					oADODBCommand("lPayrollID") = aPayrollComponent(N_ID_PAYROLL)
					oADODBCommand("lForPayrollID") = aPayrollComponent(N_FOR_DATE_PAYROLL)
					oADODBCommand("iPayrollType") = aPayrollComponent(N_TYPE_ID_PAYROLL)
					oADODBCommand("sConditions") = sCondition

					Call AppendTextToFile(sFilePath & "_Rastreo.txt", "CreateAndRunFiles(1). Antes" & vbTab & "aPayrollComponent(N_ID_PAYROLL)=" & aPayrollComponent(N_ID_PAYROLL) & vbTab & "aPayrollComponent(N_FOR_DATE_PAYROLL)=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "aPayrollComponent(N_TYPE_ID_PAYROLL)=" & aPayrollComponent(N_TYPE_ID_PAYROLL) & vbTab & "sCondition=" & sCondition & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
					oADODBCommand.Execute
					Call AppendTextToFile(sFilePath & "_Rastreo.txt", "CreateAndRunFiles(1). Después" & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")

					lTotal = oADODBCommand("iInsertCounts")
					Call AppendTextToFile(sFilePath & "_Rastreo.txt", "iInsertCounts" & lTotal & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
					If bCount Then Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "Registros de ConceptID=0 en Payroll: " & vbTab & GetPayrollConcept1Count(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")
					Set oADODBCommand = Nothing
					Set param = Nothing
					'lErrorNumber = GatherTableStats(oRequest, oADODBConnection, "Payroll_" & aPayrollComponent(N_ID_PAYROLL), 25, "for all indexed columns size auto", True, sErrorDescription)
					'Call AppendTextToFile(sFilePath & "_Rastreo.txt", "GatherTableStats(Payroll_20130131)" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
				Else
					lErrorNumber = CreateConceptsFile1(oRequest, oADODBConnection, 1, aPayrollComponent(N_FOR_DATE_PAYROLL), sErrorDescription)
					Call AppendTextToFile(sFilePath & "_Rastreo.txt", "CreateConceptsFile(oRequest, oADODBConnection, 1, aPayrollComponent(N_FOR_DATE_PAYROLL), sErrorDescription)" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
					If bCount Then Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")
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
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryLis1.EmployeeID From EmployeesChangesLKP, EmployeesHistoryLis1, StatusEmployees, Reasons " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryLis1.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryLis1.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryLis1.EmployeeDate<=EmployeesHistoryLis1.EndDate) And (EmployeesHistoryLis1.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryLis1.ReasonID=Reasons.ReasonID) And (EmployeesHistoryLis1.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) " & asEmployeesQueries(2) & " Order By EmployeesHistoryLis1.EmployeeID", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										If lErrorNumber = 0 Then
											'If Not oRecordset.EOF Then
											'	Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) " & asEmployeesQueries(2) & " Order By EmployeesHistoryList.EmployeeID" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
											'End If
											Do While Not oRecordset.EOF
												lErrorNumber = AppendTextToFile(sFilePath & "_Payroll1_" & Int(iCounter / ROWS_PER_FILE) & ".txt", asEmployeesQueries(0) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EmployeeID").Value), sErrorDescription)
												iCounter = iCounter + 1
												oRecordset.MoveNext
												'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
												If bTimeout Then Exit Do
											Loop
											'If Not oRecordset.EOF Then
											'	Call AppendTextToFile(sFilePath & "_Rastreo.txt", "iCounter: " & iCounter & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
											'End If
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
		Call AppendTextToFile(sFilePath & "_Rastreo.txt", "START: LEVEL 2, RUN FROM FILES, QttyID=1, " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
		If bCount Then Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")
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
			End If


			If False Then
				iCounter = 0
				Set oADODBCommand = Server.CreateObject("ADODB.Command")
				Set oADODBCommand.ActiveConnection = oADODBConnection
				oADODBCommand.commandtype=4
				oADODBCommand.commandtext = "SIAP.UpdateEmpChangesLKP"
				Set param = oADODBCommand.Parameters
				param.append oADODBCommand.createparameter("lPayrollID", 3, 1)
				param.append oADODBCommand.createparameter("iUpdateCounts", 3, 2)

				oADODBCommand("lPayrollID") = 20130131

				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "UpdateEmpChangesLKP. Antes" & vbTab & "aPayrollComponent(N_ID_PAYROLL)=" & aPayrollComponent(N_ID_PAYROLL) & vbTab & "aPayrollComponent(N_FOR_DATE_PAYROLL)=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "aPayrollComponent(N_TYPE_ID_PAYROLL)=" & aPayrollComponent(N_TYPE_ID_PAYROLL) & vbTab & "sCondition=" & sCondition & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
				oADODBCommand.Execute
				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "UpdateEmpChangesLKP. Después" & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")

				iCounter = oADODBCommand("iUpdateCounts")
				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "iUpdateCounts" & iCounter & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")
				Set oADODBCommand = Nothing
				Set param = Nothing
			End If


			If False Then
				If bCount Then Call AppendTextToFile(sFilePath & "_Rastros.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")
			End If

			If False Then

				If aPayrollComponent(N_TYPE_ID_PAYROLL) = 4 Then
					sErrorDescription = "No se pudieron obtener las nóminas para calcular los pagos retroactivos."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PayrollDate From Payrolls Where (PayrollDate>=" & lFirstPayroll & ") And (PayrollDate<=" & lLastPayroll & ") And (PayrollTypeID=1)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Call AppendTextToFile(sFilePath & "_Rastros.txt", "Select PayrollDate From Payrolls Where (PayrollDate>=" & lFirstPayroll & ") And (PayrollDate<=" & lLastPayroll & ") And (PayrollTypeID=1)" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
					If lErrorNumber = 0 Then
						asPayrolls = ""
						Do While Not oRecordset.EOF
							asPayrolls = asPayrolls & CStr(oRecordset.Fields("PayrollDate").Value) & ","
							oRecordset.MoveNext
							If Err.number <> 0 Then Exit Do
						Loop
						oRecordset.Close
					End If
				Else

	Call DisplayTimeStamp("START: LEVEL 2, RETROACTIVE. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
						Call BuildCondition(sCondition, "")

						sErrorDescription = "No se pudieron obtener los registros de la base de datos."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct EmployeesRevisions.StartPayrollID From EmployeesRevisions, Payrolls Where (EmployeesRevisions.StartPayrollID=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=1) And (EmployeesRevisions.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesRevisions.StartPayrollID<" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesRevisions.EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & ")) Order By EmployeesRevisions.StartPayrollID", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						Call AppendTextToFile(sFilePath & "_Rastros.txt", "Select Distinct EmployeesRevisions.StartPayrollID From EmployeesRevisions, Payrolls Where (EmployeesRevisions.StartPayrollID=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=1) And (EmployeesRevisions.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesRevisions.StartPayrollID<" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesRevisions.EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & ")) Order By EmployeesRevisions.StartPayrollID" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
						asPayrolls = ""
						If lErrorNumber = 0 Then
							Do While Not oRecordset.EOF
								asPayrolls = asPayrolls & CStr(oRecordset.Fields("StartPayrollID").Value) & ","
								oRecordset.MoveNext
								If Err.number <> 0 Then Exit Do
							Loop
							oRecordset.Close
						End If

					asPayrolls = asPayrolls & aPayrollComponent(N_FOR_DATE_PAYROLL) & ","
				End If
				If Len(asPayrolls) > 0 Then asPayrolls = Left(asPayrolls, (Len(asPayrolls) - Len(",")))

				If Len(asPayrolls) > 0 Then
					asPayrolls = Split(asPayrolls, ",")
					For iPayrollIndex = 0 To UBound(asPayrolls) - 1
						aPayrollComponent(N_FOR_DATE_PAYROLL) = CLng(asPayrolls(iPayrollIndex))
	Call LogErrorInXMLFile("123", "Vic: Inicia llamado a DoCalculations para revisión de " & aPayrollComponent(N_FOR_DATE_PAYROLL) & "aPayrollComponent(N_ID_PAYROLL)=" & aPayrollComponent(N_ID_PAYROLL) & vbTab & "aPayrollComponent(N_FOR_DATE_PAYROLL)=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "aPayrollComponent(N_TYPE_ID_PAYROLL)=" & aPayrollComponent(N_TYPE_ID_PAYROLL) & vbTab & "sCondition=" & sCondition, 0, "_", "_", 0)
	Call AppendTextToFile(sFilePath & "_Rastros.txt", "Vic: Inicia llamado a DoCalculations para " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "aPayrollComponent(N_ID_PAYROLL)=" & aPayrollComponent(N_ID_PAYROLL) & vbTab & "aPayrollComponent(N_FOR_DATE_PAYROLL)=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "aPayrollComponent(N_TYPE_ID_PAYROLL)=" & aPayrollComponent(N_TYPE_ID_PAYROLL) & vbTab & "sCondition=" & sCondition & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
						'lErrorNumber = DoCalculations1(aPayrollComponent, True, False, sErrorDescription)
	Call LogErrorInXMLFile("123", "Vic: Termina llamado a DoCalculations para " & aPayrollComponent(N_FOR_DATE_PAYROLL), 0, "_", "_", 0)
	Call AppendTextToFile(sFilePath & "_Rastros.txt", "Vic: Termina llamado a DoCalculations para revisión de " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
						If bTimeout Then Exit For
					Next
					aPayrollComponent(N_FOR_DATE_PAYROLL) = CLng(asPayrolls(iPayrollIndex))
				End If
				If Not bTimeout Then
	Call LogErrorInXMLFile("123", "Vic: Inicia llamado a DoCalculations para " & aPayrollComponent(N_FOR_DATE_PAYROLL), 0, "_", "_", 0)
	Call AppendTextToFile(sFilePath & "_Rastros.txt", "Vic: Inicia llamado 4 a DoCalculations para " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "aPayrollComponent(N_ID_PAYROLL)=" & aPayrollComponent(N_ID_PAYROLL) & vbTab & "aPayrollComponent(N_FOR_DATE_PAYROLL)=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "aPayrollComponent(N_TYPE_ID_PAYROLL)=" & aPayrollComponent(N_TYPE_ID_PAYROLL) & vbTab & "sCondition=" & sCondition & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
					'lErrorNumber = DoCalculations1(aPayrollComponent, (aPayrollComponent(N_TYPE_ID_PAYROLL) = 4), False, sErrorDescription)
	Call LogErrorInXMLFile("123", "Vic: Termina llamado a DoCalculations para " & aPayrollComponent(N_FOR_DATE_PAYROLL), 0, "_", "_", 0)
	Call AppendTextToFile(sFilePath & "_Rastros.txt", "Vic: Termina llamado 4 a DoCalculations para " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
				End If
			End If



			If False Then
				lTotal = 0
				Set oADODBCommand = Server.CreateObject("ADODB.Command")
				Set oADODBCommand.ActiveConnection = oADODBConnection
				oADODBCommand.commandtype=4
				oADODBCommand.commandtext = "SIAP.Prueba"
				Set param = oADODBCommand.Parameters
				param.append oADODBCommand.createparameter("NoEmp", 3, 1)
				param.append oADODBCommand.createparameter("Total", 3, 2)

				oADODBCommand("NoEmp") = 183945
				oADODBCommand.Execute
				lTotal = oADODBCommand("Total")
			End If


			If False Then
				'(lPayrollID NUMBER, lForPayrollID NUMBER, iPayrollType NUMBER, sConditions VARCHAR2 := '', iInsertCounts OUT NUMBER) 
				lTotal = 0
				Set oADODBCommand = Server.CreateObject("ADODB.Command")
				Set oADODBCommand.ActiveConnection = oADODBConnection
				oADODBCommand.commandtype=4
				oADODBCommand.commandtext = "SIAP.CreateAndRunCredits"
				Set param = oADODBCommand.Parameters
				param.append oADODBCommand.createparameter("lPayrollID", 3, 1)
				param.append oADODBCommand.createparameter("lForPayrollID", 3, 1)
				param.append oADODBCommand.createparameter("iPayrollType", 3, 1)
				param.append oADODBCommand.createparameter("sConditions", vbString, 1)
				param.append oADODBCommand.createparameter("iInsertCounts", 3, 2)

				oADODBCommand("lPayrollID") = aPayrollComponent(N_ID_PAYROLL)
				oADODBCommand("lForPayrollID") = aPayrollComponent(N_FOR_DATE_PAYROLL)
				oADODBCommand("iPayrollType") = aPayrollComponent(N_TYPE_ID_PAYROLL)
				oADODBCommand("sConditions") = sCondition

				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Antes" & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
				oADODBCommand.Execute
				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Después" & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")

				lErrorNumber = Err.number
				sErrorDescription = Err.description

				lTotal = oADODBCommand("iInsertCounts")
				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "iInsertCounts" & lTotal & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
			End If


			If False Then
				lTotal = 0
				Set oADODBCommand = Server.CreateObject("ADODB.Command")
				Set oADODBCommand.ActiveConnection = oADODBConnection
				oADODBCommand.commandtype=4
				oADODBCommand.commandtext = "SIAP.CreateAndRunConcepts"
				Set param = oADODBCommand.Parameters
				param.append oADODBCommand.createparameter("lPayrollIDC", 3, 1)
				param.append oADODBCommand.createparameter("lForPayrollIDC", 3, 1)
				param.append oADODBCommand.createparameter("iPayrollTypeC", 3, 1)
				param.append oADODBCommand.createparameter("sConditionsC", vbString, 1)
				param.append oADODBCommand.createparameter("iInsertCountsC", 3, 2)

				oADODBCommand("lPayrollIDC") = aPayrollComponent(N_ID_PAYROLL)
				oADODBCommand("lForPayrollIDC") = aPayrollComponent(N_FOR_DATE_PAYROLL)
				oADODBCommand("iPayrollTypeC") = aPayrollComponent(N_TYPE_ID_PAYROLL)
				oADODBCommand("sConditionsC") = sCondition

				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Antes" & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
				oADODBCommand.Execute
				lErrorNumber = Err.number
				sErrorDescription = Err.description
				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Después" & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")

				lTotal = oADODBCommand("iInsertCountsC")
				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "iInsertCounts=" & lTotal & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
			End If


			If False Then
				lTotal = 0
				Set oADODBCommand = Server.CreateObject("ADODB.Command")
				Set oADODBCommand.ActiveConnection = oADODBConnection
				oADODBCommand.commandtype=4
				oADODBCommand.commandtext = "SIAP.CreateAndRunFiles"
				Set param = oADODBCommand.Parameters
				param.append oADODBCommand.createparameter("iFileType", 3, 1)
				param.append oADODBCommand.createparameter("lPayrollID", 3, 1)
				param.append oADODBCommand.createparameter("lForPayrollID", 3, 1)
				param.append oADODBCommand.createparameter("iPayrollType", 3, 1)
				param.append oADODBCommand.createparameter("sConditions", vbString, 1)
				param.append oADODBCommand.createparameter("iInsertCounts", 3, 2)

				oADODBCommand("iFileType") = 5
				oADODBCommand("lPayrollID") = aPayrollComponent(N_ID_PAYROLL)
				oADODBCommand("lForPayrollID") = aPayrollComponent(N_FOR_DATE_PAYROLL)
				oADODBCommand("iPayrollType") = aPayrollComponent(N_TYPE_ID_PAYROLL)
				oADODBCommand("sConditions") = sCondition

				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Antes" & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
				oADODBCommand.Execute
				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Después" & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")

				lErrorNumber = Err.number
				sErrorDescription = Err.description

				lTotal = oADODBCommand("iInsertCounts")
				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "iInsertCounts" & lTotal & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
			End If



			If False Then
				lTotal = 0
				Set oADODBCommand = Server.CreateObject("ADODB.Command")
				Set oADODBCommand.ActiveConnection = oADODBConnection
				oADODBCommand.commandtype=4
				oADODBCommand.commandtext = "SIAP.UpdateAntiquities"
				Set param = oADODBCommand.Parameters
				param.append oADODBCommand.createparameter("lPayrollID", 3, 1)
				param.append oADODBCommand.createparameter("lForPayrollID", 3, 1)
				param.append oADODBCommand.createparameter("iPayrollType", 3, 1)
				param.append oADODBCommand.createparameter("sConditions", vbString, 1)
				param.append oADODBCommand.createparameter("iInsertCounts", 3, 2)
				param.append oADODBCommand.createparameter("sEmployeesFor44", vbString, 2, 3000)

				oADODBCommand("lPayrollID") = 20130131 'aPayrollComponent(N_ID_PAYROLL)
				oADODBCommand("lForPayrollID") = 20130131 'aPayrollComponent(N_FOR_DATE_PAYROLL)
				oADODBCommand("iPayrollType") = 1 'aPayrollComponent(N_TYPE_ID_PAYROLL)
				oADODBCommand("sConditions") = sCondition

				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "UpdateAntiquities. Antes" & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
				oADODBCommand.Execute
				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "UpdateAntiquities. Después" & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
				sEmployeesFor44 = oADODBCommand("sEmployeesFor44")
				sEmployeesFor44 = Trim(sEmployeesFor44)
				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "sEmployeesFor44: "& vbTab & sEmployeesFor44 & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
				lTotal = oADODBCommand("iInsertCounts")
				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "iInsertCounts" & lTotal & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
			End If


			If False Then 'Prueba antiguedad Procedimientos


						Call DisplayTimeStamp("START: LEVEL 2. Update Antiquities 1 y 2. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
						Call AppendTextToFile(sFilePath & "_Rastreo.txt", "START: LEVEL 2. Update Antiquities 1 y 2. " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
								If Not bTimeout Then
									sErrorDescription = "No se pudieron obtener las antigüedades de los empleados."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Antiquities Where (AntiquityID>-1)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Select * From Antiquities Where (AntiquityID>-1)" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
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
									Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate, EmployeesHistoryList.JobID, StatusEmployees.Active, Reasons.ActiveEmployeeID From EmployeesHistoryList, EmployeesChangesLKP, StatusEmployees, Reasons " & sQueryBegin & " Where (EmployeesHistoryList.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (StatusEmployees.Active=1) And (ActiveEmployeeID=1) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") " & sCondition & " Order By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeDate Desc" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
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
						Call AppendTextToFile(sFilePath & "_Rastreo.txt", "START: LEVEL 2. RUN FROM FILES, Update Employees.AntiquityID " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
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
									Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Select EmployeesAntiquitiesLKP.EmployeeID, AntiquityYears, AntiquityMonths, AntiquityDays From EmployeesAntiquitiesLKP, EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons " & sQueryBegin & " Where (EmployeesAntiquitiesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) " & sCondition & " Order By EmployeesAntiquitiesLKP.EmployeeID" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
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
						Call AppendTextToFile(sFilePath & "_Rastreo.txt", "START: LEVEL 2. RUN FROM FILES, Update Employees.Antiquity2ID- " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
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
									Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Select EmployeesAbsencesLKP.EmployeeID, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate From EmployeesAbsencesLKP, EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons " & sQueryBegin & " Where (EmployeesAbsencesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesAbsencesLKP.AbsenceID In (10,95)) And (EmployeesAbsencesLKP.OcurredDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") " & sCondition & " Order By EmployeesAbsencesLKP.EmployeeID" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
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
						Call AppendTextToFile(sFilePath & "_Rastreo.txt", "START: LEVEL 2. RUN FROM FILES, Update Employees.Antiquity2ID+ " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
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
									Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Select EmployeeID, Antiquity2ID, Antiquity3ID From Employees Where Antiquity2ID<0" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
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
						Call AppendTextToFile(sFilePath & "_Rastreo.txt", "START: LEVEL 2. RUN FROM FILES, Update Employees.Antiquity2ID " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
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

			End If			


		Else ' No optimiza

			iCounter = 0
			If Not bTimeout Then
				aiDays = Split("0,0", ",")
				For iIndex = 0 To UBound(aiDays)
					aiDays(iIndex) = 0
				Next


				If Not bTimeout Then
					sErrorDescription = "No se pudieron obtener las antigüedades de los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Antiquities Where (AntiquityID>-1)", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Select * From Antiquities Where (AntiquityID>-1)" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
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


				sErrorDescription = "No se pudieron obtener las antigüedades de los empleados."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate, EmployeesHistoryList.JobID, StatusEmployees.Active, Reasons.ActiveEmployeeID From EmployeesHistoryList, EmployeesChangesLKP, StatusEmployees, Reasons " & sQueryBegin & " Where (EmployeesHistoryList.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (StatusEmployees.Active=1) And (ActiveEmployeeID=1) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") AND (EmployeesHistoryList.EmployeeID IN(238,542)) " & sCondition & " Order By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeDate Desc", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate, EmployeesHistoryList.JobID, StatusEmployees.Active, Reasons.ActiveEmployeeID From EmployeesHistoryList, EmployeesChangesLKP, StatusEmployees, Reasons " & sQueryBegin & " Where (EmployeesHistoryList.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (StatusEmployees.Active=1) And (ActiveEmployeeID=1) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") AND (EmployeesHistoryList.EmployeeID IN(238,542)) " & sCondition & " Order By EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeDate Desc" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
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
		Call AppendTextToFile(sFilePath & "_Rastreo.txt", "START: LEVEL 2. RUN FROM FILES, Update Employees.AntiquityID " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
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

		End If

	Set oRecordset = Nothing
	DoCalculations2 = lErrorNumber
	Err.Clear
End Function


Function DoCalculations1(aPayrollComponent, bRetroactive, bAdjustment, sErrorDescription)
'************************************************************
'Purpose: To calculate the payroll and save it into the database
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DoCalculations1"
	Const ROWS_PER_FILE = 10000
	Const CONCEPTS_FOR_FACTOR = "1,3,12,13,14,38,49,89"
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
	Dim bOptimiza
	Dim mCount

	bTruncate = False
	sTruncate = ""
	bOptimiza = True
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


Call AppendTextToFile(sFilePath & "_Rastreo.txt", "START: LEVEL 2, CREATE FILES, QttyID=1. " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
		If Not bTimeout Then
			lErrorNumber = CreateConceptsFile(oRequest, oADODBConnection, 1, aPayrollComponent(N_FOR_DATE_PAYROLL), sErrorDescription)
			Call AppendTextToFile(sFilePath & "_Rastreo.txt", "CreateConceptsFile(oRequest, oADODBConnection, 1, aPayrollComponent(N_FOR_DATE_PAYROLL), sErrorDescription)" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
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
								Call AppendTextToFile(sFilePath & "_Rastreo.txt", "Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lPayID & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) " & asEmployeesQueries(2) & " Order By EmployeesHistoryList.EmployeeID" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
								mCount = 0
								If lErrorNumber = 0 Then
									Do While Not oRecordset.EOF
										lErrorNumber = AppendTextToFile(sFilePath & "_Payroll1_" & Int(iCounter / ROWS_PER_FILE) & ".txt", asEmployeesQueries(0) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EmployeeID").Value), sErrorDescription)
										iCounter = iCounter + 1
										mCount = mCount + 1
										oRecordset.MoveNext
										'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
										If bTimeout Then Exit Do
									Loop
									oRecordset.Close
								End If
								Call AppendTextToFile(sFilePath & "_Rastreo.txt", "iCounter: " & vbTab & iCounter & vbTab & "mCount: " & vbTab & mCount & "iIndex: " & iIndex+1 & mCount & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
							End If
						End If
						'If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
					Next
				End If

				If Not bTimeout Then
					If (lErrorNumber = 0) And (iCounter > 0) Then
'Call DisplayTimeStamp("START: LEVEL 2, RUN FROM FILES, QttyID=1, " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
Call AppendTextToFile(sFilePath & "_Rastreo.txt", "START: LEVEL 2, RUN FROM FILES, QttyID=1, " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")
						sQueryBegin = "Insert Into abPayroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ", 1, "
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
							'Call DeleteFile(sFilePath & "_Payroll1_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
						Next
					End If
				End If
			End If
		End If
Call AppendTextToFile(sFilePath & "_Rastreo.txt", "END: LEVEL 2, RUN FROM FILES, QttyID=1, " & iCounter & " RECORDS. " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "sErrorDescription=" & sErrorDescription & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastreo")

	Set oRecordset = Nothing
	DoCalculations1 = lErrorNumber
	Err.Clear
End Function


Function CalculatePayroll(oRequest, oADODBConnection, iLevel, aPayrollComponent, sErrorDescription)
'************************************************************
'Purpose: To calculate the payroll and save it into the database
'Inputs:  oRequest, oADODBConnection, iLevel
'Outputs: aPayrollComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CalculatePayroll"
	Const ROWS_PER_FILE = 10000
	Dim lPayrollDate
	Dim asEmployeesQueries
	Dim iCounter
	Dim iCounter2
	Dim iPayrollIndex
	Dim iIndex
	Dim jIndex
	Dim sFilePath
	Dim asFileContents
	Dim sDate
	Dim oStartDate
	Dim oEndDate
	Dim sQueryBegin
	Dim sQueryEnd
	Dim sCondition
	Dim sConceptCondition
	Dim lCurrentID
	Dim lCurrentID2
	Dim lConceptID
	Dim sCurrentID
	Dim adTotal
	Dim dAmount
	Dim dTaxAmount
	Dim dTemp
	Dim sTemp
	Dim lFirstPayroll
	Dim lLastPayroll
	Dim asPayrolls
	Dim lEmployeeTypeID
	Dim adTaxes
	Dim adAllowances
	Dim adTaxInvertions
	Dim bTruncate
	Dim sTruncate
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bTruncate = False
	sTruncate = ""
	bComponentInitialized = aPayrollComponent(B_COMPONENT_INITIALIZED_PAYROLL)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollComponent(oRequest, aPayrollComponent)
	End If

	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "alter session set sort_area_size=250000000", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	lErrorNumber = GetPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
	If lErrorNumber = 0 Then
		aPayrollComponent(N_FOR_DATE_PAYROLL) = CLng(aPayrollComponent(N_FOR_DATE_PAYROLL))
		lPayrollDate = aPayrollComponent(N_FOR_DATE_PAYROLL)
		oStartDate = Now()
		sFilePath = Server.MapPath("Export\Payroll_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & aPayrollComponent(N_ID_PAYROLL))

		If ((iLevel < 3) Or (iLevel=-1)) And (lErrorNumber = 0) Then
			Call BuildCondition(sCondition, sQueryBegin)
			Call AppendTextToFile(sFilePath & "_Rastros.txt", "sCondition " & sCondition & ", sQueryBegin " & sQueryBegin & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
			sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
			If Len(sCondition) = 0 Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (-" & aPayrollComponent(N_ID_PAYROLL) & "," & aPayrollComponent(N_ID_PAYROLL) & "))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Call AppendTextToFile(sFilePath & "_Rastros.txt", "Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (-" & aPayrollComponent(N_ID_PAYROLL) & "," & aPayrollComponent(N_ID_PAYROLL) & "))" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
				Response.Write "<!-- Query: Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (-" & aPayrollComponent(N_ID_PAYROLL) & "," & aPayrollComponent(N_ID_PAYROLL) & ")) -->" & vbNewLine
				If lErrorNumber = 0 Then
Call DisplayTimeStamp("START: LEVEL 1, INSERT RECORDS, EmployeesChanges")
					sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
					lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollDate, Max(EmployeeDate) As EmployeeDate1, " & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & " As FirstDate, " & aPayrollComponent(N_FOR_DATE_PAYROLL) & " As LastDate, 0 As Concepts40 From EmployeesHistoryList Where (EmployeeDate<=EndDate) And (EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EndDate>=" & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & ") And (Active=1) Group By EmployeeID", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
					Call AppendTextToFile(sFilePath & "_Rastros.txt", "Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollDate, Max(EmployeeDate) As EmployeeDate1, " & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & " As FirstDate, " & aPayrollComponent(N_FOR_DATE_PAYROLL) & " As LastDate, 0 As Concepts40 From EmployeesHistoryList Where (EmployeeDate<=EndDate) And (EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EndDate>=" & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & ") And (Active=1) Group By EmployeeID" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
					Response.Write "<!-- Query: Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollDate, Max(EmployeeDate) As EmployeeDate1, " & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & " As FirstDate, " & aPayrollComponent(N_FOR_DATE_PAYROLL) & " As LastDate, 0 As Concepts40 From EmployeesHistoryList Where (EmployeeDate<=EndDate) And (EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EndDate>=" & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & ") And (Active=1) Group By EmployeeID -->" & vbNewLine
				End If
			Else
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (-" & aPayrollComponent(N_ID_PAYROLL) & "," & aPayrollComponent(N_ID_PAYROLL) & ")) And (EmployeeID IN (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & "))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Call AppendTextToFile(sFilePath & "_Rastros.txt", "Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (-" & aPayrollComponent(N_ID_PAYROLL) & "," & aPayrollComponent(N_ID_PAYROLL) & ")) And (EmployeeID IN (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & "))" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
				Response.Write "<!-- Query: Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (-" & aPayrollComponent(N_ID_PAYROLL) & "," & aPayrollComponent(N_ID_PAYROLL) & ")) And (EmployeeID IN (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & ")) -->" & vbNewLine
				If lErrorNumber = 0 Then
Call DisplayTimeStamp("START: LEVEL 1, INSERT RECORDS, EmployeesChanges")
					sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
					lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollDate, Max(EmployeeDate) As EmployeeDate1, " & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & " As FirstDate, " & aPayrollComponent(N_FOR_DATE_PAYROLL) & " As LastDate, 0 As Concepts40 From EmployeesHistoryList" & sQueryBegin & " Where (EmployeeDate<=EndDate) And (EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EndDate>=" & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & ")" & sCondition & sConceptCondition & " And (Active=1) Group By EmployeeID", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
					Call AppendTextToFile(sFilePath & "_Rastros.txt", "Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollDate, Max(EmployeeDate) As EmployeeDate1, " & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & " As FirstDate, " & aPayrollComponent(N_FOR_DATE_PAYROLL) & " As LastDate, 0 As Concepts40 From EmployeesHistoryList" & sQueryBegin & " Where (EmployeeDate<=EndDate) And (EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EndDate>=" & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & ")" & sCondition & sConceptCondition & " And (Active=1) Group By EmployeeID" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
					Response.Write "<!-- Query: Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollDate, Max(EmployeeDate) As EmployeeDate1, " & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & " As FirstDate, " & aPayrollComponent(N_FOR_DATE_PAYROLL) & " As LastDate, 0 As Concepts40 From EmployeesHistoryList" & sQueryBegin & " Where (EmployeeDate<=EndDate) And (EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EndDate>=" & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & ")" & sCondition & sConceptCondition & " And (Active=1) Group By EmployeeID -->" & vbNewLine
				End If
			End If
			sErrorDescription = "No se pudieron eliminar los conceptos de pagos de la nómina."
			If (aPayrollComponent(N_TYPE_ID_PAYROLL) = 4) Then
				lFirstPayroll = CLng(oRequest("StartPayrollYear").Item & Right(("0" & oRequest("StartPayrollMonth").Item), Len("00")) & Right(("0" & oRequest("StartPayrollDay").Item), Len("00")))
				lLastPayroll = CLng(oRequest("EndPayrollYear").Item & Right(("0" & oRequest("EndPayrollMonth").Item), Len("00")) & Right(("0" & oRequest("EndPayrollDay").Item), Len("00")))
			End If
			If Len(sCondition) = 0 Then
				If (aPayrollComponent(N_TYPE_ID_PAYROLL) = 4) And (Len(oRequest("DeleteFromPayroll").Item) = 0) Then
					sCondition = " Where (RecordDate>=" & lFirstPayroll & ") And (RecordDate<=" & lLastPayroll & ")"
				End If
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & sCondition, "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				Call AppendTextToFile(sFilePath & "_Rastros.txt", "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & sCondition & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
				Response.Write vbNewLine & "<!-- Query: Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & sCondition & " -->" & vbNewLine
				sCondition = ""
			ElseIf (aPayrollComponent(N_TYPE_ID_PAYROLL) = 4) And (Len(oRequest("DeleteFromPayroll").Item) = 0) Then
				If iConnectionType <> ACCESS_DSN Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID = (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & ")) And (RecordDate>=" & lFirstPayroll & ") And (RecordDate<=" & lLastPayroll & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Call AppendTextToFile(sFilePath & "_Rastros.txt", "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID = (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & ")) And (RecordDate>=" & lFirstPayroll & ") And (RecordDate<=" & lLastPayroll & ")" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
					Response.Write vbNewLine & "<!-- Query: Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID = (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & ")) And (RecordDate>=" & lFirstPayroll & ") And (RecordDate<=" & lLastPayroll & ") -->" & vbNewLine
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & ")) And (RecordDate>=" & lFirstPayroll & ") And (RecordDate<=" & lLastPayroll & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Call AppendTextToFile(sFilePath & "_Rastros.txt", "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & ")) And (RecordDate>=" & lFirstPayroll & ") And (RecordDate<=" & lLastPayroll & ")" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
					Response.Write vbNewLine & "<!-- Query: Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & ")) And (RecordDate>=" & lFirstPayroll & ") And (RecordDate<=" & lLastPayroll & ") -->" & vbNewLine
				End If
			Else
				If iConnectionType <> ACCESS_DSN Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID = (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & "))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Call AppendTextToFile(sFilePath & "_Rastros.txt", "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID = (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & "))" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
					Response.Write vbNewLine & "<!-- Query: Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID = (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & ")) -->" & vbNewLine
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & "))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Call AppendTextToFile(sFilePath & "_Rastros.txt", "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & "))" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
					Response.Write vbNewLine & "<!-- Query: Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & ")) -->" & vbNewLine
				End If
			End If

'			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "exec dbms_stats.gather_table_stats('SIAP','Payroll_" & aPayrollComponent(N_ID_PAYROLL) & "',ESTIMATE_PERCENT=>25,METHOD_OPT=>'for all indexed columns size auto',CASCADE=>True)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
'			Call AppendTextToFile(sFilePath & "_Rastros.txt", "exec dbms_stats.gather_table_stats('SIAP','Payroll_" & aPayrollComponent(N_ID_PAYROLL) & "',ESTIMATE_PERCENT=>25,METHOD_OPT=>'for all indexed columns size auto',CASCADE=>True)" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
			
			If aPayrollComponent(N_TYPE_ID_PAYROLL) = 4 Then
				sErrorDescription = "No se pudieron obtener las nóminas para calcular los pagos retroactivos."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PayrollDate From Payrolls Where (PayrollDate>=" & lFirstPayroll & ") And (PayrollDate<=" & lLastPayroll & ") And (PayrollTypeID=1)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Call AppendTextToFile(sFilePath & "_Rastros.txt", "Select PayrollDate From Payrolls Where (PayrollDate>=" & lFirstPayroll & ") And (PayrollDate<=" & lLastPayroll & ") And (PayrollTypeID=1)" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
				If lErrorNumber = 0 Then
					asPayrolls = ""
					Do While Not oRecordset.EOF
						asPayrolls = asPayrolls & CStr(oRecordset.Fields("PayrollDate").Value) & ","
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
			Else
				If ((iLevel = 2) Or (iLevel=-1)) And (lErrorNumber = 0) Then
Call DisplayTimeStamp("START: LEVEL 2, RETROACTIVE. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
					Call BuildCondition(sCondition, "")

					sErrorDescription = "No se pudieron obtener los registros de la base de datos."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct EmployeesRevisions.StartPayrollID From EmployeesRevisions, Payrolls Where (EmployeesRevisions.StartPayrollID=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=1) And (EmployeesRevisions.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesRevisions.StartPayrollID<" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesRevisions.EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & ")) Order By EmployeesRevisions.StartPayrollID", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Call AppendTextToFile(sFilePath & "_Rastros.txt", "Select Distinct EmployeesRevisions.StartPayrollID From EmployeesRevisions, Payrolls Where (EmployeesRevisions.StartPayrollID=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=1) And (EmployeesRevisions.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesRevisions.StartPayrollID<" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesRevisions.EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & ")) Order By EmployeesRevisions.StartPayrollID" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
					asPayrolls = ""
					If lErrorNumber = 0 Then
						Do While Not oRecordset.EOF
							asPayrolls = asPayrolls & CStr(oRecordset.Fields("StartPayrollID").Value) & ","
							oRecordset.MoveNext
							If Err.number <> 0 Then Exit Do
						Loop
						oRecordset.Close
					End If
				End If
				asPayrolls = asPayrolls & aPayrollComponent(N_FOR_DATE_PAYROLL) & ","
			End If
			If Len(asPayrolls) > 0 Then asPayrolls = Left(asPayrolls, (Len(asPayrolls) - Len(",")))
		End If

		If ((iLevel = 2) Or (iLevel=-1)) And (lErrorNumber = 0) Then
			If Len(asPayrolls) > 0 Then
				asPayrolls = Split(asPayrolls, ",")
				For iPayrollIndex = 0 To UBound(asPayrolls) - 1
					aPayrollComponent(N_FOR_DATE_PAYROLL) = CLng(asPayrolls(iPayrollIndex))
Call LogErrorInXMLFile("123", "Vic: Inicia llamado a DoCalculations para " & aPayrollComponent(N_FOR_DATE_PAYROLL), 0, "_", "_", 0)
Call AppendTextToFile(sFilePath & "_Rastros.txt", "Vic: Inicia llamado a DoCalculations para revisión de " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "aPayrollComponent(N_ID_PAYROLL)=" & aPayrollComponent(N_ID_PAYROLL) & vbTab & "aPayrollComponent(N_FOR_DATE_PAYROLL)=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "aPayrollComponent(N_TYPE_ID_PAYROLL)=" & aPayrollComponent(N_TYPE_ID_PAYROLL) & vbTab & "sCondition=" & sCondition & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
					lErrorNumber = DoCalculations(aPayrollComponent, True, False, sErrorDescription)
Call LogErrorInXMLFile("123", "Vic: Termina llamado a DoCalculations para " & aPayrollComponent(N_FOR_DATE_PAYROLL), 0, "_", "_", 0)
Call AppendTextToFile(sFilePath & "_Rastros.txt", "Vic: Termina llamado a DoCalculations para revisión de " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
					If bTimeout Then Exit For
				Next
				aPayrollComponent(N_FOR_DATE_PAYROLL) = CLng(asPayrolls(iPayrollIndex))
			End If
			If Not bTimeout Then
Call LogErrorInXMLFile("123", "Vic: Inicia llamado a DoCalculations para " & aPayrollComponent(N_FOR_DATE_PAYROLL), 0, "_", "_", 0)
Call AppendTextToFile(sFilePath & "_Rastros.txt", "Vic: Inicia llamado 4 a DoCalculations para " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "aPayrollComponent(N_ID_PAYROLL)=" & aPayrollComponent(N_ID_PAYROLL) & vbTab & "aPayrollComponent(N_FOR_DATE_PAYROLL)=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "aPayrollComponent(N_TYPE_ID_PAYROLL)=" & aPayrollComponent(N_TYPE_ID_PAYROLL) & vbTab & "sCondition=" & sCondition & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
Call AppendTextToFile(sFilePath & "_Rastros.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "Registros de ConceptID=0 en Payroll: " & vbTab & GetPayrollConcept1Count(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")
				'If aPayrollComponent(N_ID_PAYROLL)<>20130131 Then
					lErrorNumber = DoCalculations(aPayrollComponent, (aPayrollComponent(N_TYPE_ID_PAYROLL) = 4), False, sErrorDescription)
				'End If
Call LogErrorInXMLFile("123", "Vic: Termina llamado a DoCalculations para " & aPayrollComponent(N_FOR_DATE_PAYROLL), 0, "_", "_", 0)
Call AppendTextToFile(sFilePath & "_Rastros.txt", "Vic: Termina llamado 4 a DoCalculations para " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
Call AppendTextToFile(sFilePath & "_Rastros.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "Registros de ConceptID=0 en Payroll: " & vbTab & GetPayrollConcept1Count(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")
			End If

			If Not bTimeout Then
				asPayrolls = ""
				sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct MissingDate From EmployeesAdjustmentsLKP Where (EmployeeID Not In (Select Distinct EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ")) And (PayrollDate In (0," & aPayrollComponent(N_ID_PAYROLL) & ")) And (Active=1) Order By MissingDate", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				If lErrorNumber = 0 Then
					If Not oRecordset.EOF Then
						Do While Not oRecordset.EOF
							asPayrolls = asPayrolls & CStr(oRecordset.Fields("MissingDate").Value) & ","
							oRecordset.MoveNext
						Loop
						oRecordset.Close
						If Len(asPayrolls) > 0 Then asPayrolls = Left(asPayrolls, (Len(asPayrolls) - Len(",")))
						asPayrolls = Split(asPayrolls, ",")

						For iPayrollIndex = 0 To UBound(asPayrolls)
							aPayrollComponent(N_FOR_DATE_PAYROLL) = CLng(asPayrolls(iPayrollIndex))
Call LogErrorInXMLFile("123", "Vic: Inicia llamado a DoCalculations para " & aPayrollComponent(N_FOR_DATE_PAYROLL), 0, "_", "_", 0)
Call AppendTextToFile(sFilePath & "_Rastros.txt", "Vic: Inicia llamado T a DoCalculations para " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
							lErrorNumber = DoCalculations(aPayrollComponent, True, True, sErrorDescription)
Call LogErrorInXMLFile("123", "Vic: Termina llamado a DoCalculations para " & aPayrollComponent(N_FOR_DATE_PAYROLL), 0, "_", "_", 0)
Call AppendTextToFile(sFilePath & "_Rastros.txt", "Vic: Termina llamado T a DoCalculations para " & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
							If bTimeout Then Exit For
						Next
					End If
				End If
			End If

			aPayrollComponent(N_FOR_DATE_PAYROLL) = CLng(lPayrollDate)

			Call BuildCondition(sCondition, "")
			Call AppendTextToFile(sFilePath & "_Rastros.txt", "sCondition: " & sCondition & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
			If Len(sCondition) = 0 Then
				If Not bTimeout Then
					sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Call AppendTextToFile(sFilePath & "_Rastros.txt", "Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ")" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
				End If
				If Not bTimeout Then
					sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
					lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollDate, Max(EmployeeDate) As EmployeeDate1, " & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & " As FirstDate, " & aPayrollComponent(N_FOR_DATE_PAYROLL) & " As LastDate, 0 As Concepts40 From EmployeesHistoryList Where (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select Distinct EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ")) Group By EmployeeID", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
					Call AppendTextToFile(sFilePath & "_Rastros.txt", "Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollDate, Max(EmployeeDate) As EmployeeDate1, " & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & " As FirstDate, " & aPayrollComponent(N_FOR_DATE_PAYROLL) & " As LastDate, 0 As Concepts40 From EmployeesHistoryList Where (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select Distinct EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ")) Group By EmployeeID" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
				End If
			Else
				If Not bTimeout Then
					sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeeID IN (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & "))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Call AppendTextToFile(sFilePath & "_Rastros.txt", "Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeeID IN (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & sConceptCondition & "))" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
				End If
				If Not bTimeout Then
					sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
					lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollDate, Max(EmployeeDate) As EmployeeDate1, " & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & " As FirstDate, " & aPayrollComponent(N_FOR_DATE_PAYROLL) & " As LastDate, 0 As Concepts40 From EmployeesHistoryList" & sQueryBegin & " Where (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select Distinct EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & "))" & sCondition & sConceptCondition & " Group By EmployeeID", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
					Call AppendTextToFile(sFilePath & "_Rastros.txt", "Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollDate, Max(EmployeeDate) As EmployeeDate1, " & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & " As FirstDate, " & aPayrollComponent(N_FOR_DATE_PAYROLL) & " As LastDate, 0 As Concepts40 From EmployeesHistoryList" & sQueryBegin & " Where (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select Distinct EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & "))" & sCondition & sConceptCondition & " Group By EmployeeID" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
				End If
			End If
			Call AppendTextToFile(sFilePath & "_Rastros.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "Registros de ConceptID=0 en Payroll: " & vbTab & GetPayrollConcept1Count(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")
			If Not bTimeout Then
				If True Then
					iCounter = 0
					Set oADODBCommand = Server.CreateObject("ADODB.Command")
					Set oADODBCommand.ActiveConnection = oADODBConnection
					oADODBCommand.commandtype=4
					oADODBCommand.commandtext = "SIAP.UpdateEmpChangesLKP"
					Set param = oADODBCommand.Parameters
					param.append oADODBCommand.createparameter("lPayrollID", 3, 1)
					param.append oADODBCommand.createparameter("iUpdateCounts", 3, 2)

					oADODBCommand("lPayrollID") = aPayrollComponent(N_ID_PAYROLL)

					Call AppendTextToFile(sFilePath & "_Rastros.txt", "UpdateEmpChangesLKP. Antes" & vbTab & "aPayrollComponent(N_ID_PAYROLL)=" & aPayrollComponent(N_ID_PAYROLL) & vbTab & "aPayrollComponent(N_FOR_DATE_PAYROLL)=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & vbTab & "aPayrollComponent(N_TYPE_ID_PAYROLL)=" & aPayrollComponent(N_TYPE_ID_PAYROLL) & vbTab & "sCondition=" & sCondition & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
					oADODBCommand.Execute
					Call AppendTextToFile(sFilePath & "_Rastros.txt", "UpdateEmpChangesLKP. Después" & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")

					iCounter = oADODBCommand("iUpdateCounts")
					Call AppendTextToFile(sFilePath & "_Rastros.txt", "iUpdateCounts" & iCounter & vbTab & "Err.number=" & Err.number & vbTab & "Err.description=" & Err.description & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
					Call AppendTextToFile(sFilePath & "_Rastros.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "Registros de ConceptID=0 en Payroll: " & vbTab & GetPayrollConcept1Count(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")
					Set oADODBCommand = Nothing
					Set param = Nothing
				Else
					iCounter = 0
					sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, Concepts40, Min(FirstDate) As MinFirstDate, Max(LastDate) As MaxLastDate From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate<0) Group By EmployeeID, Concepts40 Order By EmployeeID", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Call AppendTextToFile(sFilePath & "_Rastros.txt", "Select EmployeeID, Concepts40, Min(FirstDate) As MinFirstDate, Max(LastDate) As MaxLastDate From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate<0) Group By EmployeeID, Concepts40 Order By EmployeeID" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
					If lErrorNumber = 0 Then
						Do While Not oRecordset.EOF
							lErrorNumber = AppendTextToFile(sFilePath & "_EmployeesChangesLKP_" & Int(iCounter / ROWS_PER_FILE) & ".txt", CStr(oRecordset.Fields("EmployeeID").Value) & "," & CStr(oRecordset.Fields("MinFirstDate").Value) & "," & CStr(oRecordset.Fields("MaxLastDate").Value) & "," & CStr(oRecordset.Fields("Concepts40").Value), sErrorDescription)
							iCounter = iCounter + 1
							oRecordset.MoveNext
						Loop
						oRecordset.Close
					End If
					If (lErrorNumber = 0) And (iCounter > 0) Then
	Call DisplayTimeStamp("START: LEVEL 2, RUN FROM FILES, Update EmployeesChangesLKP " & iCounter & " RECORDS.")
	Call AppendTextToFile(sFilePath & "_Rastros.txt", "START: LEVEL 2, RUN FROM FILES, Update EmployeesChangesLKP " & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
						sQueryBegin = "Update EmployeesChangesLKP Set FirstDate=<FIRST_DATE />, LastDate=<LAST_DATE />, Concepts40=<CONCEPTS_40 /> Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeeID=<EMPLOYEE_ID />)"
						For jIndex = 0 To iCounter Step ROWS_PER_FILE
							asFileContents = GetFileContents(sFilePath & "_EmployeesChangesLKP_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
							If Len(asFileContents) > 0 Then
								asFileContents = Split(asFileContents, vbNewLine)
								For iIndex = 0 To UBound(asFileContents)
									If Len(asFileContents(iIndex)) > 0 Then
										asEmployeesQueries = Split(asFileContents(iIndex), ",")
										sErrorDescription = "No se pudo modificar la nómina del empleado."
										lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, Replace(Replace(Replace(Replace(sQueryBegin, "<FIRST_DATE />", asEmployeesQueries(1)), "<LAST_DATE />", asEmployeesQueries(2)), "<CONCEPTS_40 />", asEmployeesQueries(3)), "<EMPLOYEE_ID />", asEmployeesQueries(0)), "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
									End If
									If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
								Next
							End If
							Call DeleteFile(sFilePath & "_EmployeesChangesLKP_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
						Next
					End If
				End If
			End If

			Call AppendTextToFile(sFilePath & "_Rastros.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "Registros de ConceptID=0 en Payroll: " & vbTab & GetPayrollConcept1Count(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")
			If Not bTimeout Then
				sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate<0)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				Call AppendTextToFile(sFilePath & "_Rastros.txt", "Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate<0)" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")

				sErrorDescription = "No se pudieron eliminar los montos temporales."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID<=0)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				Call AppendTextToFile(sFilePath & "_Rastros.txt", "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (EmployeeID<=0)" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")

				sErrorDescription = "No se pudieron eliminar los montos temporales."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptAmount=0)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				Call AppendTextToFile(sFilePath & "_Rastros.txt", "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptAmount=0)" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")

				sErrorDescription = "No se pudieron obtener los montos totales de las percepciones de los empleados."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID In (-2,-1,0))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				Call AppendTextToFile(sFilePath & "_Rastros.txt", "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID In (-2,-1,0))" & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
			End If
			Call AppendTextToFile(sFilePath & "_Rastros.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "Registros de ConceptID=0 en Payroll: " & vbTab & GetPayrollConcept1Count(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")

Call DisplayTimeStamp("START: LEVEL 2, CREATE FILES, ConceptID=-1. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
Call AppendTextToFile(sFilePath & "_Rastros.txt", "START: LEVEL 2, CREATE FILES, ConceptID=-1." & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
Call AppendTextToFile(sFilePath & "_Rastros.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "Registros de ConceptID=0 en Payroll: " & vbTab & GetPayrollConcept1Count(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")
			If Not bTimeout Then
				Call BuildCondition("", sQueryBegin)
				sErrorDescription = "No se pudieron obtener los montos totales de las percepciones de los empleados."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & aPayrollComponent(N_FOR_DATE_PAYROLL) & "' As RecordDate, '1' As RecordID, EmployeeID, '-1' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", Concepts Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Concepts.ConceptID) And ((Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID<700000) Or (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID>=800000)) And (Concepts.IsDeduction=0) And (Concepts.ConceptID>0) And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") Group By EmployeeID", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			Call AppendTextToFile(sFilePath & "_Rastros.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "Registros de ConceptID=0 en Payroll: " & vbTab & GetPayrollConcept1Count(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")

Call DisplayTimeStamp("START: LEVEL 2, CREATE FILES, ConceptID=-2. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
Call AppendTextToFile(sFilePath & "_Rastros.txt", "START: LEVEL 2, CREATE FILES, ConceptID=-2." & vbTab & "lErrorNumber=" & lErrorNumber & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")) & vbTab & "bTimeout=" & bTimeout, "Registro de Rastros")
Call AppendTextToFile(sFilePath & "_Rastros.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "Registros de ConceptID=0 en Payroll: " & vbTab & GetPayrollConcept1Count(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")
			If Not bTimeout Then
				sErrorDescription = "No se pudieron obtener los montos totales de las deducciones de los empleados."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & aPayrollComponent(N_FOR_DATE_PAYROLL) & "' As RecordDate, '1' As RecordID, EmployeeID, '-2' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", Concepts Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Concepts.ConceptID) And ((Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID<700000) Or (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID>=800000)) And (Concepts.IsDeduction=1) And (Concepts.ConceptID>0) And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") Group By EmployeeID", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			End If
			Call AppendTextToFile(sFilePath & "_Rastros.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "Registros de ConceptID=0 en Payroll: " & vbTab & GetPayrollConcept1Count(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")

			If Not bTimeout Then
				If bTruncate Then
Call DisplayTimeStamp("START: LEVEL 2, TRUNCATE DECIMALS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
					sTruncate = "-2,-1"
					If False Then
						sErrorDescription = "No se pudieron truncar los decimales de los montos."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=Round(ConceptAmount, 2) Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (" & sTruncate & "))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Else
						sErrorDescription = "No se pudo limpiar la tabla temporal."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PayrollInt", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudieron truncar los decimales de los montos."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=(ConceptAmount+0.005)*100 Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (" & sTruncate & "))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudieron truncar los decimales de los montos."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PayrollInt (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (" & sTruncate & "))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo limpiar la tabla de montos."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (" & sTruncate & "))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudieron truncar los decimales de los montos."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID From PayrollInt", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudieron truncar los decimales de los montos."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=(ConceptAmount/100) Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (" & sTruncate & "))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If

Call DisplayTimeStamp("START: LEVEL 2, CREATE FILES, ConceptID=0. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
Call AppendTextToFile(sFilePath & "_Rastros.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "Registros de ConceptID=0 en Payroll: " & vbTab & GetPayrollConcept1Count(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")
			If Not bTimeout Then
				sErrorDescription = "No se pudo limpiar la tabla temporal."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudieron obtener los montos totales de las percepciones de los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '0' As RecordDate, '-1' As RecordID, EmployeeID, '-1' As ConceptID, '1' As PayrollTypeID, ConceptAmount, '0' As ConceptTaxes, '0' As ConceptRetention, '-1' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=-1)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudieron obtener los montos totales de las deducciones de los empleados."
					lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into Payroll (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '0' As RecordDate, '-2' As RecordID, EmployeeID, '-2' As ConceptID, '1' As PayrollTypeID, -ConceptAmount, '0' As ConceptTaxes, '0' As ConceptRetention, '-1' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=-2)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
				End If
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudieron obtener los montos totales de las percepciones y de las deducciones de los empleados."
					lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select '" & aPayrollComponent(N_FOR_DATE_PAYROLL) & "' As RecordDate, '1' As RecordID, EmployeeID, '0' As ConceptID, '1' As PayrollTypeID, Sum(ConceptAmount) As ConceptAmount1, '0' As ConceptTaxes, '0' As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll Group By EmployeeID", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
				End If
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudo limpiar la tabla temporal."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If

			If Not bTimeout Then
				If bTruncate Then
Call DisplayTimeStamp("START: LEVEL 2, TRUNCATE DECIMALS. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
					sTruncate = "0"
					If False Then
						sErrorDescription = "No se pudieron truncar los decimales de los montos."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=Round(ConceptAmount, 2) Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (" & sTruncate & "))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					Else
						sErrorDescription = "No se pudo limpiar la tabla temporal."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PayrollInt", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
						If lErrorNumber = 0 Then
							sErrorDescription = "No se pudieron truncar los decimales de los montos."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=(ConceptAmount+0.005)*100 Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (" & sTruncate & "))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudieron truncar los decimales de los montos."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PayrollInt (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (" & sTruncate & "))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudo limpiar la tabla de montos."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (" & sTruncate & "))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
									If lErrorNumber = 0 Then
										sErrorDescription = "No se pudieron truncar los decimales de los montos."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID From PayrollInt", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										If lErrorNumber = 0 Then
											sErrorDescription = "No se pudieron truncar los decimales de los montos."
											lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptAmount=(ConceptAmount/100) Where (RecordDate=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (ConceptID In (" & sTruncate & "))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If

		If ((iLevel = 3) Or (iLevel=-1)) And (lErrorNumber = 0) Then
Call DisplayTimeStamp("START: LEVEL 3, PREPARE Payments, PaymentsMessages, PaymentsRecords, PaymentsRecords2 TABLE")
			sErrorDescription = "No se pudieron preparar las tablas para los pagos."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payments Where (PaymentDate=" & aPayrollComponent(N_ID_PAYROLL) & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudieron preparar las tablas para los pagos."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesHistoryListForPayroll Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			End If
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudieron preparar las tablas para los pagos."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PaymentsRecords Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			End If
			If lErrorNumber = 0 Then
				sErrorDescription = "No se pudieron preparar las tablas para los pagos."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PaymentsRecords2 Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			End If

			sErrorDescription = "No se pudo limpiar la tabla temporal."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

			sErrorDescription = "No se pudo actualizar la tabla de respaldo del historial."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesHistoryListForPayroll Select EmployeesHistoryList.EmployeeID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollID, EmployeesHistoryList.EmployeeDate, EmployeeNumber, CompanyID, JobID, ServiceID, ZoneID, EmployeeTypeID, PositionTypeID, ClassificationID, GroupGradeLevelID, IntegrationID, JourneyID, ShiftID, WorkingHours, AreaID, PositionID, LevelID, StatusID, PaymentCenterID, RiskLevel, EmployeesHistoryList.Active, ReasonID, EmployeesHistoryList.ModifyDate, EmployeesHistoryList.PayrollDate, EmployeesHistoryList.UserID,BankAccounts.AccountNumber, BankAccounts.BankID From EmployeesHistoryList, EmployeesChangesLKP, BankAccounts Where (EmployeesHistoryList.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesHistoryList.EmployeeDate=EmployeesChangesLKP.EmployeeDate) And (EmployeesChangesLKP.PayrollID=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeID=BankAccounts.EmployeeID) And (BankAccounts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (BankAccounts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (BankAccounts.Active=1)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			If lErrorNumber = -2147217900 Then
				sErrorDescription = "Existen registros duplicados para los empleados, posiblemente en las cuentas bancarias. " & sErrorDescription
			ElseIf lErrorNumber = 0 Then
Call DisplayTimeStamp("START: LEVEL 3, CLOSE RETROACTIVE AMOUNTS")
				sErrorDescription = "No se pudieron actualizar los registros utilizados para el cálculo de los pagos retroactivos."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesHistoryList Set PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ", bProcessed=1 Where (bProcessed=0) And (EmployeesHistoryList.Active=1)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

Call DisplayTimeStamp("START: LEVEL 3, CLOSE RETROACTIVE AMOUNTS")
				sErrorDescription = "No se pudieron actualizar los registros utilizados para el cálculo de los pagos retroactivos."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAdjustmentsLKP Set PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & " Where (PayrollDate=0)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

Call DisplayTimeStamp("START: LEVEL 3, CLOSE RETROACTIVE AMOUNTS")
				sErrorDescription = "No se pudieron actualizar los registros utilizados para el cálculo de los pagos retroactivos."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesAbsencesLKP Set AppliedDate=" & aPayrollComponent(N_ID_PAYROLL) & " Where (AppliedDate=0)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)

				If iConnectionType <> ACCESS_DSN Then
Call DisplayTimeStamp("START: LEVEL 3, UPDATE Credits.DebtAmount")
					sErrorDescription = "No se pudieron actualizar los registros para los créditos."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Credits Set PaymentsCounter=PaymentsCounter+1, DebtAmount=DebtAmount-ConceptAmount From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", Credits Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=Credits.EmployeeID) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Credits.CreditTypeID) And (Credits.CreditTypeID>0) And ((Credits.PaymentsCounter<Credits.PaymentsNumber) Or (Credits.PaymentsNumber<1)) And (Credits.Active=1) And (Credits.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Credits.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				End If

Call DisplayTimeStamp("START: LEVEL 3, ADD PaymentMessages")
				lErrorNumber = InsertPaymentMessages(oADODBConnection, aPayrollComponent, sErrorDescription)

				If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 3, RUN FROM FILES, ADD PayrollsCLCs")
					sErrorDescription = "No se pudieron agregar los mensajes para las nóminas de los empleados."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PayrollsCLCs (PayrollID, EmployeeID, PayrollCode, PayrollCLC, FilterParameters) Select '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollID, EmployeeID, ' ' As PayrollCode, ' ' As PayrollCLC, ' ' As FilterParameters From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=0)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
Call AppendTextToFile(sFilePath & "_Rastros.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "Registros de ConceptID=0 en Payroll: " & vbTab & GetPayrollConcept1Count(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")

Call DisplayTimeStamp("START: LEVEL 3, UPDATE PAYROLL_YYYY")
				sErrorDescription = "No se pudo actualizar el acumulado anual."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & Left(aPayrollComponent(N_ID_PAYROLL), Len("0000")) & " Where (RecordDate=" & aPayrollComponent(N_ID_PAYROLL) & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If lErrorNumber = 0 Then
					sErrorDescription = "No se pudo actualizar el acumulado anual."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & Left(aPayrollComponent(N_ID_PAYROLL), Len("0000")) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select " & aPayrollComponent(N_ID_PAYROLL) & " As RecordDate, RecordDate As RecordID, EmployeeID, ConceptID, 1 As PayrollTypeID, ConceptAmount, 0 As ConceptTaxes, RecordID As ConceptRetention, '" & aLoginComponent(N_USER_ID_LOGIN) & "' As UserID From Payroll_" & aPayrollComponent(N_ID_PAYROLL), "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
		End If
Call AppendTextToFile(sFilePath & "_Rastros.txt", "Registros en Payroll: " & vbTab & GetPayrollCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), sErrorDescription) & vbTab & "Registros en Payroll Amount Cero: " & vbTab & GetPayrollConceptCount(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "Registros de ConceptID=0 en Payroll: " & vbTab & GetPayrollConcept1Count(oRequest, oADODBConnection, aPayrollComponent(N_ID_PAYROLL), 1, sErrorDescription) & vbTab & "TimeStamp=" & Year(Now()) & Right(("0" & Month(Now())), Len("00")) & Right(("0" & Day(Now())), Len("00")) & Right(("0" & Hour(Now())), Len("00")) & Right(("0" & Minute(Now())), Len("00")) & Right(("0" & Second(Now())), Len("00")), "Registro de Rastros")

Call DisplayTimeStamp("END")
		oEndDate = Now()
		If (lErrorNumber = 0) And (Not bTimeout) And (Not FileExists(Server.MapPath("Database\Stop.txt"), "")) And B_USE_SMTP Then
			If DateDiff("n", oStartDate, oEndDate) > 5 Then
				ReDim aEmailComponent(N_EMAIL_COMPONENT_SIZE)
				aEmailComponent(S_TO_EMAIL) = aLoginComponent(S_USER_E_MAIL_LOGIN)
				aEmailComponent(S_FROM_EMAIL) = aLoginComponent(S_USER_E_MAIL_LOGIN)
				Select Case iLevel
					Case 1
						aEmailComponent(S_SUBJECT_EMAIL) = "SIAP. La nómina ya fue creada."
						aEmailComponent(S_BODY_EMAIL) = GetFileContents(Server.MapPath("Template_PayrollReady.htm"), sErrorDescription)
					Case 3
						aEmailComponent(S_SUBJECT_EMAIL) = "SIAP. La nómina ya fue cerrada."
						aEmailComponent(S_BODY_EMAIL) = GetFileContents(Server.MapPath("Template_PayrollClosed.htm"), sErrorDescription)
					Case Else
						aEmailComponent(S_SUBJECT_EMAIL) = "SIAP. La nómina ya fue procesada."
						aEmailComponent(S_BODY_EMAIL) = GetFileContents(Server.MapPath("Template_PayrollCalculated.htm"), sErrorDescription)
				End Select
				If Len(aEmailComponent(S_BODY_EMAIL)) > 0 Then
					aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<SYSTEM_URL />", S_HTTP & SERVER_IP_FOR_LICENSE & SYSTEM_PORT & "/" & VIRTUAL_DIRECTORY_NAME)
					aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<USER_NAME />", CleanStringForHTML(aLoginComponent(S_USER_NAME_LOGIN) & " " & aLoginComponent(S_USER_LAST_NAME_LOGIN)))
					aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<PAYROLL_NAME />", CleanStringForHTML(aPayrollComponent(S_NAME_PAYROLL)))
					sDate = GetSerialNumberForDate(oStartDate)
					aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<START_DATE />", DisplayDateFromSerialNumber(Left(sDate, Len("00000000")), Mid(sDate, 9, 2), Mid(sDate, 11, 2), Mid(sDate, 13, 2)))
					sDate = GetSerialNumberForDate(oEndDate)
					aEmailComponent(S_BODY_EMAIL) = Replace(aEmailComponent(S_BODY_EMAIL), "<END_DATE />", DisplayDateFromSerialNumber(Left(sDate, Len("00000000")), Mid(sDate, 9, 2), Mid(sDate, 11, 2), Mid(sDate, 13, 2)))
					Call SendEmail(oRequest, aEmailComponent, "")
				End If
			End If
		End If
	End If
	If bTimeout Then
		Call LogErrorInXMLFile("999", "El cálculo de la nómina generó un error de Timeout debido a cuestiones de la base de datos o del procesador del serivdor", 000, "PayrollComponentConstants.asp", S_FUNCTION_NAME, N_SQL_ERROR_LEVEL)
		If B_USE_SMTP Then
			ReDim aEmailComponent(N_EMAIL_COMPONENT_SIZE)
			aEmailComponent(S_TO_EMAIL) = aLoginComponent(S_USER_E_MAIL_LOGIN)
			aEmailComponent(S_FROM_EMAIL) = aLoginComponent(S_USER_E_MAIL_LOGIN)
			aEmailComponent(S_SUBJECT_EMAIL) = "SIAP. El proceso de la nómina se truncó."
			aEmailComponent(S_BODY_EMAIL) = "El proceso de la nómina se truncó por Time out a las " & DisplayDateFromSerialNumber(Left(sDate, Len("00000000")), Mid(sDate, 9, 2), Mid(sDate, 11, 2), Mid(sDate, 13, 2))
			Call SendEmail(oRequest, aEmailComponent, "")
		End If
	End If
	If FileExists(Server.MapPath("Database\Stop.txt"), "") Then
		Call DeleteFile(Server.MapPath("Database\Stop.txt"), "")
		Call LogErrorInXMLFile("666", "El cálculo de la nómina fue denetido por el usuario " & aLoginComponent(S_ACCESS_KEY_LOGIN) & ".", 000, "PayrollComponentConstants.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
		If B_USE_SMTP Then
			ReDim aEmailComponent(N_EMAIL_COMPONENT_SIZE)
			aEmailComponent(S_TO_EMAIL) = aLoginComponent(S_USER_E_MAIL_LOGIN)
			aEmailComponent(S_FROM_EMAIL) = aLoginComponent(S_USER_E_MAIL_LOGIN)
			aEmailComponent(S_SUBJECT_EMAIL) = "SIAP. El proceso de la nómina se truncó."
			aEmailComponent(S_BODY_EMAIL) = "El proceso de la nómina se truncó, utilizando el archivo Stop.txt, a las " & DisplayDateFromSerialNumber(Left(sDate, Len("00000000")), Mid(sDate, 9, 2), Mid(sDate, 11, 2), Mid(sDate, 13, 2))
			Call SendEmail(oRequest, aEmailComponent, "")
		End If
	End If
	If (iLevel = 2) Or (iLevel=-1) Then
		Application.Contents("SIAP_CalculatePayroll") = ""
	End If

	Set oRecordset = Nothing
	CalculatePayroll = lErrorNumber
	Err.Clear
End Function

Function CalculateQttyID_8_9(oRequest, oADODBConnection, bCurrent, bRetroactive, sErrorDescription)
'************************************************************
'Purpose: To calculate the amount for extra hours and sundays
'Inputs:  oRequest, oADODBConnection, bCurrent, bRetroactive
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CalculateQttyID_8_9"
	Const ROWS_PER_FILE = 10000
	Dim lPayrollID
	Dim lForPayrollDate
	Dim lPayID
	Dim asEmployeesQueries
	Dim iCounter
	Dim iCounter2
	Dim iIndex
	Dim jIndex
	Dim kIndex
	Dim sPeriods
	Dim sFilePath
	Dim asFileContents
	Dim sQueryBegin
	Dim sQueryEnd
	Dim sCondition
	Dim lCurrentID
	Dim lCurrentID2
	Dim sCurrentID
	Dim adDSM
	Dim adTotal
	Dim dAmount
	Dim dTaxAmount
	Dim dTemp
	Dim oRecordset
	Dim lErrorNumber

	Call BuildCondition(sCondition, sQueryBegin)
	If bRetroactive And (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) Then sCondition = " And (EmployeesHistoryList.EmployeeID=EmployeesRevisions.EmployeeID) And (EmployeesRevisions.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesRevisions.StartPayrollID=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")" & sCondition
	lPayrollID = aPayrollComponent(N_ID_PAYROLL)
	lForPayrollDate = aPayrollComponent(N_FOR_DATE_PAYROLL)
	lPayID = lPayrollID
	If Not bCurrent Then
		lPayrollID = CLng(Left(CStr(aPayrollComponent(N_ID_PAYROLL)), Len("YYYYMM")) & "15")
		lForPayrollDate = CLng(Left(CStr(aPayrollComponent(N_FOR_DATE_PAYROLL)), Len("YYYYMM")) & "15")
		lPayID = Left(lPayrollID, Len("YYYY"))
	End If
	sFilePath = Server.MapPath("Export\Payroll_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & aPayrollComponent(N_ID_PAYROLL))

	sPeriods = ""
	sPeriods = GetPeriodsForPayroll(lPayrollID, lForPayrollDate, -1)

	adDSM = "0"
	sErrorDescription = "No se pudieron obtener los días de salario mínimo y los días de salario burocrático."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CurrencyValue From CurrenciesHistoryList Where (CurrencyDate=" & lForPayrollDate & ") And (CurrencyID In (1,2,3,4,5)) Order By CurrencyID", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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

Call DisplayTimeStamp("START: LEVEL 2, CREATE FILES, QttyID=8. " & lForPayrollDate)
	iCounter = 0
	If Not bTimeout Then
		If bCurrent Then lErrorNumber = CreateConceptsFile(oRequest, oADODBConnection, 8, lForPayrollDate, sErrorDescription)
		If lErrorNumber = 0 Then
			asFileContents = GetFileContents(PAYROLL_FILE8_PATH, sErrorDescription)
			If Len(asFileContents) > 0 Then
				asFileContents = Split(asFileContents, vbNewLine)
				For iIndex = 0 To UBound(asFileContents)
					If Len(asFileContents(iIndex)) > 0 Then
						asEmployeesQueries = Split(asFileContents(iIndex), LIST_SEPARATOR)
						If InStr(1, "," & sPeriods & ",", "," & asEmployeesQueries(2) & ",", vbbinaryCompare) > 0 Then
							sQueryBegin = ""
							asEmployeesQueries(9) = asEmployeesQueries(9) & sCondition
							If InStr(1, asEmployeesQueries(9), "(Jobs.", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", Jobs"
							If (InStr(1, asEmployeesQueries(9), "=Employees.", vbBinaryCompare) > 0) Or (InStr(1, asEmployeesQueries(9), "(Employees.", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", Employees"
							If InStr(1, asEmployeesQueries(9), "EmployeesChildrenLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesChildrenLKP"
							If InStr(1, asEmployeesQueries(9), "EmployeesRisksLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesRisksLKP"
							If InStr(1, asEmployeesQueries(9), "EmployeesSyndicatesLKP", vbBinaryCompare) > 0 Then sQueryBegin = sQueryBegin & ", EmployeesSyndicatesLKP"
							If (aPayrollComponent(N_TYPE_ID_PAYROLL) <> 4) And (InStr(1, asEmployeesQueries(9), "EmployeesRevisions", vbBinaryCompare) > 0) Then sQueryBegin = sQueryBegin & ", EmployeesRevisions"
							sErrorDescription = "No se pudieron obtener los empleados para registrar sus conceptos de pago en la nómina."
							If CInt(asEmployeesQueries(3)) = 8 Then
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, AbsenceHours, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate, ZoneTypeID From EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, EmployeesAbsencesLKP, Areas, Zones " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollDate=" & lForPayrollDate & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.EmployeeID=EmployeesAbsencesLKP.EmployeeID) And (EmployeesAbsencesLKP.AbsenceID In (201,909)) And (AppliedDate In (0," & lForPayrollDate & ")) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (EmployeesAbsencesLKP.Active=1) " & asEmployeesQueries(9) & " Order By EmployeesHistoryList.EmployeeID, OcurredDate", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							Else
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, AbsenceHours, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate, ZoneTypeID From EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, EmployeesAbsencesLKP, Areas, Zones " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & lPayrollID & ") And (EmployeesChangesLKP.PayrollDate=" & lForPayrollDate & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.ZoneID=Zones.ZoneID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.EmployeeID=EmployeesAbsencesLKP.EmployeeID) And (EmployeesAbsencesLKP.AbsenceID In (202,914)) And (AppliedDate In (0," & lForPayrollDate & ")) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (EmployeesAbsencesLKP.Active=1) " & asEmployeesQueries(9) & " Order By EmployeesHistoryList.EmployeeID, OcurredDate", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
							End If
							If lErrorNumber = 0 Then
								If Not oRecordset.EOF Then
									If Len(asEmployeesQueries(4)) = 0 Then
										lCurrentID = -2
										Do While Not oRecordset.EOF
											If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
												If lCurrentID <> -2 Then
													sCurrentID = sCurrentID & ";" & dAmount & "," & dTemp & ",,"
													lErrorNumber = AppendTextToFile(sFilePath & "_Payroll8_" & Int(iCounter / ROWS_PER_FILE) & ".txt", asEmployeesQueries(0) & SECOND_LIST_SEPARATOR & (dAmount * CDbl(asEmployeesQueries(1))) & SECOND_LIST_SEPARATOR & lCurrentID & SECOND_LIST_SEPARATOR & SECOND_LIST_SEPARATOR & dTaxAmount & SECOND_LIST_SEPARATOR & sCurrentID & SECOND_LIST_SEPARATOR & asEmployeesQueries(1), sErrorDescription)
													iCounter = iCounter + 1
												End If
												lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
												lCurrentID2 = 0
												sCurrentID = "0,0,,"
												dAmount = 0
												dTemp = 0
												dTaxAmount = adDSM(CInt(oRecordset.Fields("ZoneTypeID").Value))
											End If
											If CInt(asEmployeesQueries(3)) = 8 Then
												If lCurrentID2 <> GetWeekStartDate(CDbl(oRecordset.Fields("OcurredDate").Value)) Then
													If dTemp > 0 Then sCurrentID = sCurrentID & ";" & dAmount & "," & dTemp & ",,"
													dTemp = dTemp + 5
													lCurrentID2 = GetWeekStartDate(CDbl(oRecordset.Fields("OcurredDate").Value))
												End If
											Else
												'If lCurrentID2 <> GetPayrollStartDate(CDbl(oRecordset.Fields("OcurredDate").Value)) Then
												'	If dTemp > 0 Then sCurrentID = sCurrentID & ";" & dAmount & "," & dTemp & ",,"
													dTemp = dTemp + 1 'GetSundaysForPayroll(CDbl(oRecordset.Fields("OcurredDate").Value))
												'	lCurrentID2 = GetPayrollStartDate(CDbl(oRecordset.Fields("OcurredDate").Value))
												'End If
											End If
											dAmount = dAmount + CDbl(oRecordset.Fields("AbsenceHours").Value)
											oRecordset.MoveNext
											'If lErrorNumber <> 0 Then Exit Do
										Loop
										sCurrentID = sCurrentID & ";" & dAmount & "," & dTemp & ",,"
										lErrorNumber = AppendTextToFile(sFilePath & "_Payroll8_" & Int(iCounter / ROWS_PER_FILE) & ".txt", asEmployeesQueries(0) & SECOND_LIST_SEPARATOR & (dAmount * CDbl(asEmployeesQueries(1))) & SECOND_LIST_SEPARATOR & lCurrentID & SECOND_LIST_SEPARATOR & SECOND_LIST_SEPARATOR & dTaxAmount & SECOND_LIST_SEPARATOR & sCurrentID & SECOND_LIST_SEPARATOR & asEmployeesQueries(1), sErrorDescription)
										iCounter = iCounter + 1
									Else
										lCurrentID = -2
										Do While Not oRecordset.EOF
											If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
												If lCurrentID <> -2 Then
													sCurrentID = sCurrentID & ";" & dAmount & "," & dTemp & ",,"
													lErrorNumber = AppendTextToFile(sFilePath & "_Payroll8_" & Int(iCounter / ROWS_PER_FILE) & ".txt", asEmployeesQueries(0) & SECOND_LIST_SEPARATOR & (dAmount * CDbl(asEmployeesQueries(1)) / 100) & SECOND_LIST_SEPARATOR & lCurrentID & SECOND_LIST_SEPARATOR & asEmployeesQueries(4) & SECOND_LIST_SEPARATOR & dTaxAmount & SECOND_LIST_SEPARATOR & sCurrentID & SECOND_LIST_SEPARATOR & (CDbl(asEmployeesQueries(1)) / 100), sErrorDescription)
													iCounter = iCounter + 1
												End If
												lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
												lCurrentID2 = 0
												sCurrentID = "0,0,,"
												dAmount = 0
												dTemp = 0
												dTaxAmount = adDSM(CInt(oRecordset.Fields("ZoneTypeID").Value))
											End If
											If CInt(asEmployeesQueries(3)) = 8 Then
												If lCurrentID2 <> GetWeekStartDate(CDbl(oRecordset.Fields("OcurredDate").Value)) Then
													If dTemp > 0 Then sCurrentID = sCurrentID & ";" & dAmount & "," & dTemp & ",,"
													dTemp = dTemp + 5
													lCurrentID2 = GetWeekStartDate(CDbl(oRecordset.Fields("OcurredDate").Value))
												End If
											Else
												'If lCurrentID2 <> GetPayrollStartDate(CDbl(oRecordset.Fields("OcurredDate").Value)) Then
												'	If dTemp > 0 Then sCurrentID = sCurrentID & ";" & dAmount & "," & dTemp & ",,"
													dTemp = dTemp + 1 'GetSundaysForPayroll(CDbl(oRecordset.Fields("OcurredDate").Value))
												'	lCurrentID2 = GetPayrollStartDate(CDbl(oRecordset.Fields("OcurredDate").Value))
												'End If
											End If
											dAmount = dAmount + CDbl(oRecordset.Fields("AbsenceHours").Value)
											oRecordset.MoveNext
											'If lErrorNumber <> 0 Then Exit Do
										Loop
										sCurrentID = sCurrentID & ";" & dAmount & "," & dTemp & ",,"
										lErrorNumber = AppendTextToFile(sFilePath & "_Payroll8_" & Int(iCounter / ROWS_PER_FILE) & ".txt", asEmployeesQueries(0) & SECOND_LIST_SEPARATOR & (dAmount * CDbl(asEmployeesQueries(1)) / 100) & SECOND_LIST_SEPARATOR & lCurrentID & SECOND_LIST_SEPARATOR & asEmployeesQueries(4) & SECOND_LIST_SEPARATOR & dTaxAmount & SECOND_LIST_SEPARATOR & sCurrentID & SECOND_LIST_SEPARATOR & (CDbl(asEmployeesQueries(1)) / 100), sErrorDescription)
										iCounter = iCounter + 1
									End If
								End If
								oRecordset.Close
							End If
						End If
					End If
					'If lErrorNumber <> 0 Then Exit For
					If bTimeout Then Exit For
				Next
			End If

			If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 2, CREATE FILES, QttyID=9. " & lForPayrollDate)
				iCounter2 = 0
				sCurrentID = ""
				For jIndex = 0 To iCounter Step ROWS_PER_FILE
					asFileContents = GetFileContents(sFilePath & "_Payroll8_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
					If Len(asFileContents) > 0 Then
						asFileContents = Split(asFileContents, vbNewLine)
						For iIndex = 0 To UBound(asFileContents)
							If Len(asFileContents(iIndex)) > 0 Then
								asEmployeesQueries = Split(asFileContents(iIndex), SECOND_LIST_SEPARATOR)
								If StrComp(asEmployeesQueries(0) & asEmployeesQueries(2), sCurrentID, vbBinaryCompare) <> 0 Then
									If Len(asEmployeesQueries(3)) = 0 Then
										lErrorNumber = AppendTextToFile(sFilePath & "_Payroll9_" & Int(jIndex / ROWS_PER_FILE) & ".txt", asEmployeesQueries(0) & SECOND_LIST_SEPARATOR & asEmployeesQueries(1) & SECOND_LIST_SEPARATOR & asEmployeesQueries(2) & SECOND_LIST_SEPARATOR & asEmployeesQueries(4), sErrorDescription)
										iCounter2 = iCounter2 + 1
									Else
										sErrorDescription = "No se pudieron obtener los empleados para registrar sus conceptos de pago en la nómina."
										lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Sum(ConceptAmount) As TotalAmount, IsDeduction From Payroll_" & lPayID & ", Concepts Where (Payroll_" & lPayID & ".ConceptID=Concepts.ConceptID) And (RecordDate=" & lForPayrollDate & ") And (EmployeeID=" & asEmployeesQueries(2) & ") And (Concepts.ConceptID In (" & asEmployeesQueries(3) & ")) And (Concepts.StartDate<=" & lForPayrollDate & ") And (Concepts.EndDate>=" & lForPayrollDate & ") Group By IsDeduction", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
										If lErrorNumber = 0 Then
											dAmount = 0
											Do While Not oRecordset.EOF
												If CInt(oRecordset.Fields("IsDeduction").Value) = 0 Then
													dAmount = dAmount + CDbl(oRecordset.Fields("TotalAmount").Value)
												Else
													dAmount = dAmount - CDbl(oRecordset.Fields("TotalAmount").Value)
												End If
												oRecordset.MoveNext
											Loop
											oRecordset.Close
											adTotal = Split(asEmployeesQueries(5), ";")
											adTotal(0) = Split(adTotal(0), ",")
											adTotal(0)(3) = 0
											For kIndex = 1 To UBound(adTotal)
												adTotal(kIndex) = Split(adTotal(kIndex), ",")
												adTotal(kIndex)(2) = dAmount * (CDbl(adTotal(kIndex)(0)) - CDbl(adTotal(kIndex - 1)(0))) * CDbl(asEmployeesQueries(6))
												adTotal(kIndex)(3) = asEmployeesQueries(4) * (CDbl(adTotal(kIndex)(1)) - CDbl(adTotal(kIndex - 1)(1)))
												If (adTotal(kIndex)(2) - adTotal(kIndex)(3)) < 0 Then
													adTotal(kIndex)(3) = 0
												Else
													adTotal(kIndex)(3) = FormatNumber((adTotal(kIndex)(2) - adTotal(kIndex)(3)), 2, True, False, False)
												End If
												adTotal(0)(3) = adTotal(0)(3) + adTotal(kIndex)(3)
											Next
											lErrorNumber = AppendTextToFile(sFilePath & "_Payroll9_" & Int(jIndex / ROWS_PER_FILE) & ".txt", asEmployeesQueries(0) & SECOND_LIST_SEPARATOR & FormatNumber((dAmount * CDbl(asEmployeesQueries(1))), 2, True, False, False) & SECOND_LIST_SEPARATOR & asEmployeesQueries(2) & SECOND_LIST_SEPARATOR & adTotal(0)(3), sErrorDescription)
											iCounter2 = iCounter2 + 1
										End If
									End If
									sCurrentID = asEmployeesQueries(0) & asEmployeesQueries(2)
								End If
							End If
							'If lErrorNumber <> 0 Then Exit For
							If bTimeout Then Exit For
						Next
					End If
					Call DeleteFile(sFilePath & "_Payroll8_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
				Next
			End If

			If Not bTimeout Then
				If (lErrorNumber = 0) And (iCounter2 > 0) Then
Call DisplayTimeStamp("START: LEVEL 2, RUN FROM FILES, QttyID=8,9, " & iCounter2 & " RECORDS. " & lForPayrollDate)
					If bCurrent Then
						sQueryBegin = "Insert Into Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Values (" & lForPayrollDate & ", 1, <EMPLOYEE_ID />, <CONCEPT_ID />, 1, <CONCEPT_AMOUNT />, <CONCEPT_TAXES />, 0, " & aLoginComponent(N_USER_ID_LOGIN) & ")"
					Else
						sQueryBegin = "Update Payroll_" & Left(aPayrollComponent(N_ID_PAYROLL), Len("YYYY")) & " Set ConceptTaxes=<CONCEPT_TAXES /> Where (RecordDate=" & lForPayrollDate & ") And (EmployeeID=<EMPLOYEE_ID />) And (ConceptID=<CONCEPT_ID />)"
					End If
					For jIndex = 0 To iCounter2 Step ROWS_PER_FILE
						asFileContents = GetFileContents(sFilePath & "_Payroll9_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
						If Len(asFileContents) > 0 Then
							asFileContents = Split(asFileContents, vbNewLine)
							For iIndex = 0 To UBound(asFileContents)
								If Len(asFileContents(iIndex)) > 0 Then
									asEmployeesQueries = Split(asFileContents(iIndex), SECOND_LIST_SEPARATOR)
									sErrorDescription = "No se pudo agregar el concepto de pago y su monto a la nómina del empleado."
									lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, Replace(Replace(Replace(Replace(sQueryBegin, "<EMPLOYEE_ID />", asEmployeesQueries(2)), "<CONCEPT_ID />", asEmployeesQueries(0)), "<CONCEPT_AMOUNT />", asEmployeesQueries(1)), "<CONCEPT_TAXES />", asEmployeesQueries(3)), "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
								End If
								'If lErrorNumber <> 0 Then Exit For
								If bTimeout Then Exit For
							Next
						End If
						If iCounter2 > 0 Then Call DeleteFile(sFilePath & "_Payroll9_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
					Next
				End If
			End If
		End If
	End If

	If (Not bTimeout) And bCurrent Then
		sErrorDescription = "No se pudieron obtener los montos de los conceptos de pago para los empleados."
		If iConnectionType <> ACCESS_DSN Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=16) And (RecordDate=" & lForPayrollDate & ") And (EmployeeID = (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lForPayrollDate & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=16) And (RecordDate=" & lForPayrollDate & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lForPayrollDate & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If

	Set oRecordset = Nothing
	CalculateQttyID_8_9 = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfPayroll(oADODBConnection, bInSADE, aPayrollComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific payroll exists in the database
'Inputs:  oADODBConnection, bInSADE, aPayrollComponent
'Outputs: aPayrollComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfPayroll"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPayrollComponent(B_COMPONENT_INITIALIZED_PAYROLL)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollComponent(oRequest, aPayrollComponent)
	End If

	If (Len(aPayrollComponent(S_NAME_PAYROLL)) = 0) Or (aPayrollComponent(N_DATE_PAYROLL) = 0) Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el nombre y/o la fecha de la nómina para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PayrollComponentConstants.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia de la nómina en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Payrolls Where (PayrollName='" & Replace(aPayrollComponent(S_NAME_PAYROLL), "'", "") & "') And (PayrollDate=" & aPayrollComponent(N_DATE_PAYROLL) & ") And (PayrollID<>" & aPayrollComponent(N_ID_PAYROLL) & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			aPayrollComponent(B_IS_DUPLICATED_PAYROLL) = (Not oRecordset.EOF)
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	CheckExistencyOfPayroll = lErrorNumber
	Err.Clear
End Function

Function CheckPayrollInformationConsistency(aPayrollComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aPayrollComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckPayrollInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aPayrollComponent(N_ID_PAYROLL)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador de la nómina no es un valor numérico."
		bIsCorrect = False
	End If
	If Len(aPayrollComponent(S_NAME_PAYROLL)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El nombre de la nómina está vacío."
		bIsCorrect = False
	End If
	If Not IsNumeric(aPayrollComponent(N_DATE_PAYROLL)) Then aPayrollComponent(N_DATE_PAYROLL) = 0
	If aPayrollComponent(N_DATE_PAYROLL) <= 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La fecha de la nómina está vacía."
		bIsCorrect = False
	End If
	If Len(aPayrollComponent(S_CLC_PAYROLL)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- La CLC de la nómina está vacía."
		bIsCorrect = False
	End If
	If Not IsNumeric(aPayrollComponent(N_TYPE_ID_PAYROLL)) Then aPayrollComponent(N_TYPE_ID_PAYROLL) = 1
	If aPayrollComponent(N_TYPE_ID_PAYROLL) = 1 Then
		aPayrollComponent(N_FOR_DATE_PAYROLL) = aPayrollComponent(N_ID_PAYROLL)
	Else
		If Not IsNumeric(aPayrollComponent(N_FOR_DATE_PAYROLL)) Then aPayrollComponent(N_FOR_DATE_PAYROLL) = 0
	End If
	If Not IsNumeric(aPayrollComponent(N_CLOSED_PAYROLL)) Then aPayrollComponent(N_CLOSED_PAYROLL) = 0

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información de la nómina contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PayrollComponentConstants.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckPayrollInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function CreateConceptsFile1(oRequest, oADODBConnection, iFileType, lPayrollID, sErrorDescription)
'************************************************************
'Purpose: To create the text file used to calculate the payroll
'Inputs:  oRequest, oADODBConnection, iFileType, lPayrollID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CreateConceptsFile1"
	Dim oRecordset
	Dim sCondition
	Dim sConceptCondition
	Dim lErrorNumber

	If Len(oRequest("PayrollConceptID").Item) > 0 Then
		sConceptCondition = " And (Concepts.ConceptID In (" & oRequest("PayrollConceptID").Item & "))"
	End If
	If (iFileType = 1) Or (iFileType = -1) Then
		sErrorDescription = "No se pudieron obtener los conceptos de pagos y sus montos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptsValues.*, Concepts.PeriodID, Antiquities.StartYears, Antiquities.EndYears, Antiquities2.StartYears As StartYears2, Antiquities2.EndYears As EndYears2, Antiquities3.StartYears As StartYears3, Antiquities3.EndYears As EndYears3, Antiquities4.StartYears As StartYears4, Antiquities4.EndYears As EndYears4 From ConceptsValues, Concepts, Antiquities, Antiquities As Antiquities2, Antiquities As Antiquities3, Antiquities As Antiquities4 Where (ConceptsValues.ConceptID=Concepts.ConceptID) And (ConceptsValues.AntiquityID=Antiquities.AntiquityID) And (ConceptsValues.Antiquity2ID=Antiquities2.AntiquityID) And (ConceptsValues.Antiquity3ID=Antiquities3.AntiquityID) And (ConceptsValues.Antiquity4ID=Antiquities4.AntiquityID) And (ConceptsValues.StartDate<=" & lPayrollID & ") And (ConceptsValues.EndDate>=" & lPayrollID & ") And (Concepts.StartDate<=" & lPayrollID & ") And (Concepts.EndDate>=" & lPayrollID & ") And (ConceptQttyID=1) And (ConceptAmount>0)" & sConceptCondition & " Order By Concepts.OrderInList, Concepts.ConceptID, CompanyID Desc, EmployeeTypeID Desc, PositionTypeID Desc, EmployeeStatusID Desc, JobStatusID Desc, ClassificationID Desc, GroupGradeLevelID Desc, IntegrationID Desc, JourneyID Desc, WorkingHours Desc, AdditionalShift Desc, LevelID Desc, EconomicZoneID Desc, ServiceID Desc, ConceptsValues.AntiquityID Desc, ConceptsValues.Antiquity2ID Desc, ConceptsValues.Antiquity3ID Desc, ConceptsValues.Antiquity4ID Desc, ForRisk Desc, GenderID Desc, HasSyndicate Desc", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			'If FileExists(PAYROLL_FILE1_PATH, sErrorDescription) Then Call DeleteFile(PAYROLL_FILE1_PATH, "")
			Do While Not oRecordset.EOF
				sCondition = CStr(oRecordset.Fields("ConceptID").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptAmount").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("PeriodID").Value) & LIST_SEPARATOR
				If CLng(oRecordset.Fields("CompanyID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryLis1.CompanyID=" & CStr(oRecordset.Fields("CompanyID").Value) & ")"
				If CLng(oRecordset.Fields("EmployeeTypeID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryLis1.EmployeeTypeID=" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & ")"
				If CLng(oRecordset.Fields("PositionTypeID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryLis1.PositionTypeID=" & CStr(oRecordset.Fields("PositionTypeID").Value) & ")"
				If CLng(oRecordset.Fields("EmployeeStatusID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryLis1.StatusID=" & CStr(oRecordset.Fields("EmployeeStatusID").Value) & ")"
				If CLng(oRecordset.Fields("JobStatusID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryLis1.JobID=Jobs.JobID) And (Jobs.StatusID=" & CStr(oRecordset.Fields("JobStatusID").Value) & ")"
				If CLng(oRecordset.Fields("ClassificationID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryLis1.ClassificationID=" & CStr(oRecordset.Fields("ClassificationID").Value) & ")"
				If CLng(oRecordset.Fields("GroupGradeLevelID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryLis1.GroupGradeLevelID=" & CStr(oRecordset.Fields("GroupGradeLevelID").Value) & ")"
				If CLng(oRecordset.Fields("IntegrationID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryLis1.IntegrationID=" & CStr(oRecordset.Fields("IntegrationID").Value) & ")"
				If CLng(oRecordset.Fields("JourneyID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryLis1.JourneyID=" & CStr(oRecordset.Fields("JourneyID").Value) & ")"
				If CLng(oRecordset.Fields("WorkingHours").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryLis1.WorkingHours=" & CStr(oRecordset.Fields("WorkingHours").Value) & ")"
				If CLng(oRecordset.Fields("AdditionalShift").Value) > 0 Then sCondition = sCondition & " And ((Employees.StartHour3>0) Or (Employees.EndHour3>0))"
				If CLng(oRecordset.Fields("LevelID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryLis1.LevelID=" & CStr(oRecordset.Fields("LevelID").Value) & ")"
				If CLng(oRecordset.Fields("EconomicZoneID").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryLis1.AreaID=Areas.AreaID) And (Areas.EconomicZoneID=" & CStr(oRecordset.Fields("EconomicZoneID").Value) & ")"
				If CLng(oRecordset.Fields("ServiceID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryLis1.ServiceID=" & CStr(oRecordset.Fields("ServiceID").Value) & ")"
				If CLng(oRecordset.Fields("AntiquityID").Value) > -1 Then sCondition = sCondition & " And (Employees.AntiquityID=" & CStr(oRecordset.Fields("AntiquityID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity2ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity2ID=" & CStr(oRecordset.Fields("Antiquity2ID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity3ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity3ID=" & CStr(oRecordset.Fields("Antiquity3ID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity4ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity4ID=" & CStr(oRecordset.Fields("Antiquity4ID").Value) & ")"
				If CLng(oRecordset.Fields("ForRisk").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryLis1.EmployeeID=EmployeesRisksLKP.EmployeeID)"
				If CLng(oRecordset.Fields("GenderID").Value) > -1 Then sCondition = sCondition & " And (Employees.GenderID=" & CStr(oRecordset.Fields("GenderID").Value) & ")"
				If CLng(oRecordset.Fields("HasChildren").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryLis1.EmployeeID=EmployeesChildrenLKP.EmployeeID)"
				If CLng(oRecordset.Fields("HasSyndicate").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryLis1.EmployeeID=EmployeesSyndicatesLKP.EmployeeID)"
				If CLng(oRecordset.Fields("PositionID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryLis1.PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & ")"
''					If InStr(1, sCondition, "(EmployeesHistoryList.JobID=Jobs.JobID)", vbBinaryCompare) = 0 Then sCondition = sCondition & " And (EmployeesHistoryList.JobID=Jobs.JobID)"
''					sCondition = sCondition & " And (EmployeesHistoryList.PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & ")"
''				End If
				If InStr(1, sCondition, "(Employees.", vbBinaryCompare) > 0 Then sCondition = sCondition & " And (EmployeesHistoryLis1.EmployeeID=Employees.EmployeeID)"

				lErrorNumber = AppendTextToFile(PAYROLL_FILE1_PATH, sCondition, sErrorDescription)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			lErrorNumber = AppendTextToFile(PAYROLL_FILE1_PATH, "", sErrorDescription)
		End If
	End If

	CreateConceptsFile1 = lErrorNumber
	Err.Clear
End Function

Function CreateConceptsFile(oRequest, oADODBConnection, iFileType, lPayrollID, sErrorDescription)
'************************************************************
'Purpose: To create the text file used to calculate the payroll
'Inputs:  oRequest, oADODBConnection, iFileType, lPayrollID
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CreateConceptsFile"
	Dim oRecordset
	Dim sCondition
	Dim sConceptCondition
	Dim lErrorNumber

	If Len(oRequest("PayrollConceptID").Item) > 0 Then
		sConceptCondition = " And (Concepts.ConceptID In (" & oRequest("PayrollConceptID").Item & "))"
	End If
	If (iFileType = 1) Or (iFileType = -1) Then
		sErrorDescription = "No se pudieron obtener los conceptos de pagos y sus montos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptsValues.*, Concepts.PeriodID, Antiquities.StartYears, Antiquities.EndYears, Antiquities2.StartYears As StartYears2, Antiquities2.EndYears As EndYears2, Antiquities3.StartYears As StartYears3, Antiquities3.EndYears As EndYears3, Antiquities4.StartYears As StartYears4, Antiquities4.EndYears As EndYears4 From ConceptsValues, Concepts, Antiquities, Antiquities As Antiquities2, Antiquities As Antiquities3, Antiquities As Antiquities4 Where (ConceptsValues.ConceptID=Concepts.ConceptID) And (ConceptsValues.AntiquityID=Antiquities.AntiquityID) And (ConceptsValues.Antiquity2ID=Antiquities2.AntiquityID) And (ConceptsValues.Antiquity3ID=Antiquities3.AntiquityID) And (ConceptsValues.Antiquity4ID=Antiquities4.AntiquityID) And (ConceptsValues.StartDate<=" & lPayrollID & ") And (ConceptsValues.EndDate>=" & lPayrollID & ") And (Concepts.StartDate<=" & lPayrollID & ") And (Concepts.EndDate>=" & lPayrollID & ") And (ConceptQttyID=1) And (ConceptAmount>0)" & sConceptCondition & " Order By Concepts.OrderInList, Concepts.ConceptID, CompanyID Desc, EmployeeTypeID Desc, PositionTypeID Desc, EmployeeStatusID Desc, JobStatusID Desc, ClassificationID Desc, GroupGradeLevelID Desc, IntegrationID Desc, JourneyID Desc, WorkingHours Desc, AdditionalShift Desc, LevelID Desc, EconomicZoneID Desc, ServiceID Desc, ConceptsValues.AntiquityID Desc, ConceptsValues.Antiquity2ID Desc, ConceptsValues.Antiquity3ID Desc, ConceptsValues.Antiquity4ID Desc, ForRisk Desc, GenderID Desc, HasSyndicate Desc", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If FileExists(PAYROLL_FILE1_PATH, sErrorDescription) Then Call DeleteFile(PAYROLL_FILE1_PATH, "")
			Do While Not oRecordset.EOF
				sCondition = CStr(oRecordset.Fields("ConceptID").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptAmount").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("PeriodID").Value) & LIST_SEPARATOR
				If CLng(oRecordset.Fields("CompanyID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.CompanyID=" & CStr(oRecordset.Fields("CompanyID").Value) & ")"
				If CLng(oRecordset.Fields("EmployeeTypeID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeTypeID=" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & ")"
				If CLng(oRecordset.Fields("PositionTypeID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.PositionTypeID=" & CStr(oRecordset.Fields("PositionTypeID").Value) & ")"
				If CLng(oRecordset.Fields("EmployeeStatusID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.StatusID=" & CStr(oRecordset.Fields("EmployeeStatusID").Value) & ")"
				If CLng(oRecordset.Fields("JobStatusID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.JobID=Jobs.JobID) And (Jobs.StatusID=" & CStr(oRecordset.Fields("JobStatusID").Value) & ")"
				If CLng(oRecordset.Fields("ClassificationID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.ClassificationID=" & CStr(oRecordset.Fields("ClassificationID").Value) & ")"
				If CLng(oRecordset.Fields("GroupGradeLevelID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.GroupGradeLevelID=" & CStr(oRecordset.Fields("GroupGradeLevelID").Value) & ")"
				If CLng(oRecordset.Fields("IntegrationID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.IntegrationID=" & CStr(oRecordset.Fields("IntegrationID").Value) & ")"
				If CLng(oRecordset.Fields("JourneyID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.JourneyID=" & CStr(oRecordset.Fields("JourneyID").Value) & ")"
				If CLng(oRecordset.Fields("WorkingHours").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.WorkingHours=" & CStr(oRecordset.Fields("WorkingHours").Value) & ")"
				If CLng(oRecordset.Fields("AdditionalShift").Value) > 0 Then sCondition = sCondition & " And ((Employees.StartHour3>0) Or (Employees.EndHour3>0))"
				If CLng(oRecordset.Fields("LevelID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.LevelID=" & CStr(oRecordset.Fields("LevelID").Value) & ")"
				If CLng(oRecordset.Fields("EconomicZoneID").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.EconomicZoneID=" & CStr(oRecordset.Fields("EconomicZoneID").Value) & ")"
				If CLng(oRecordset.Fields("ServiceID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.ServiceID=" & CStr(oRecordset.Fields("ServiceID").Value) & ")"
				If CLng(oRecordset.Fields("AntiquityID").Value) > -1 Then sCondition = sCondition & " And (Employees.AntiquityID=" & CStr(oRecordset.Fields("AntiquityID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity2ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity2ID=" & CStr(oRecordset.Fields("Antiquity2ID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity3ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity3ID=" & CStr(oRecordset.Fields("Antiquity3ID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity4ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity4ID=" & CStr(oRecordset.Fields("Antiquity4ID").Value) & ")"
				If CLng(oRecordset.Fields("ForRisk").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=EmployeesRisksLKP.EmployeeID)"
				If CLng(oRecordset.Fields("GenderID").Value) > -1 Then sCondition = sCondition & " And (Employees.GenderID=" & CStr(oRecordset.Fields("GenderID").Value) & ")"
				If CLng(oRecordset.Fields("HasChildren").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=EmployeesChildrenLKP.EmployeeID)"
				If CLng(oRecordset.Fields("HasSyndicate").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=EmployeesSyndicatesLKP.EmployeeID)"
				If CLng(oRecordset.Fields("PositionID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & ")"
''					If InStr(1, sCondition, "(EmployeesHistoryList.JobID=Jobs.JobID)", vbBinaryCompare) = 0 Then sCondition = sCondition & " And (EmployeesHistoryList.JobID=Jobs.JobID)"
''					sCondition = sCondition & " And (EmployeesHistoryList.PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & ")"
''				End If
				If InStr(1, sCondition, "(Employees.", vbBinaryCompare) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID)"

				lErrorNumber = AppendTextToFile(PAYROLL_FILE1_PATH, sCondition, sErrorDescription)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			lErrorNumber = AppendTextToFile(PAYROLL_FILE1_PATH, "", sErrorDescription)
		End If
	End If

	If (iFileType = 2) Or (iFileType = -1) Then
		sErrorDescription = "No se pudieron obtener los conceptos de pagos y sus montos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptsValues.*, Concepts.PeriodID, Antiquities.StartYears, Antiquities.EndYears, Antiquities2.StartYears As StartYears2, Antiquities2.EndYears As EndYears2, Antiquities3.StartYears As StartYears3, Antiquities3.EndYears As EndYears3, Antiquities4.StartYears As StartYears4, Antiquities4.EndYears As EndYears4 From ConceptsValues, Concepts, Antiquities, Antiquities As Antiquities2, Antiquities As Antiquities3, Antiquities As Antiquities4 Where (ConceptsValues.ConceptID=Concepts.ConceptID) And (ConceptsValues.AntiquityID=Antiquities.AntiquityID) And (ConceptsValues.Antiquity2ID=Antiquities2.AntiquityID) And (ConceptsValues.Antiquity3ID=Antiquities3.AntiquityID) And (ConceptsValues.Antiquity4ID=Antiquities4.AntiquityID) And (ConceptsValues.StartDate<=" & lPayrollID & ") And (ConceptsValues.EndDate>=" & lPayrollID & ") And (Concepts.StartDate<=" & lPayrollID & ") And (Concepts.EndDate>=" & lPayrollID & ") And (ConceptQttyID=2) And (Concepts.ConceptID Not In (70)) " & sConceptCondition & " Order By Concepts.OrderInList, Concepts.ConceptID, CompanyID Desc, EmployeeTypeID Desc, PositionTypeID Desc, EmployeeStatusID Desc, JobStatusID Desc, ClassificationID Desc, GroupGradeLevelID Desc, IntegrationID Desc, JourneyID Desc, WorkingHours Desc, AdditionalShift Desc, LevelID Desc, EconomicZoneID Desc, ServiceID Desc, ConceptsValues.AntiquityID Desc, ConceptsValues.Antiquity2ID Desc, ConceptsValues.Antiquity3ID Desc, ConceptsValues.Antiquity4ID Desc, ForRisk Desc, GenderID Desc, HasSyndicate Desc", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If FileExists(PAYROLL_FILE2_PATH, sErrorDescription) Then Call DeleteFile(PAYROLL_FILE2_PATH, "")
			Do While Not oRecordset.EOF
				sCondition = CStr(oRecordset.Fields("ConceptID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptAmount").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("PeriodID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("AppliesToID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMin").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMinQttyID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMax").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMaxQttyID").Value) & LIST_SEPARATOR
				If CLng(oRecordset.Fields("CompanyID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.CompanyID=" & CStr(oRecordset.Fields("CompanyID").Value) & ")"
				If CLng(oRecordset.Fields("EmployeeTypeID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeTypeID=" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & ")"
				If CLng(oRecordset.Fields("PositionTypeID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.PositionTypeID=" & CStr(oRecordset.Fields("PositionTypeID").Value) & ")"
				If CLng(oRecordset.Fields("EmployeeStatusID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.StatusID=" & CStr(oRecordset.Fields("EmployeeStatusID").Value) & ")"
				If CLng(oRecordset.Fields("JobStatusID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.JobID=Jobs.JobID) And (Jobs.StatusID=" & CStr(oRecordset.Fields("JobStatusID").Value) & ")"
				If CLng(oRecordset.Fields("ClassificationID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.ClassificationID=" & CStr(oRecordset.Fields("ClassificationID").Value) & ")"
				If CLng(oRecordset.Fields("GroupGradeLevelID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.GroupGradeLevelID=" & CStr(oRecordset.Fields("GroupGradeLevelID").Value) & ")"
				If CLng(oRecordset.Fields("IntegrationID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.IntegrationID=" & CStr(oRecordset.Fields("IntegrationID").Value) & ")"
				If CLng(oRecordset.Fields("JourneyID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.JourneyID=" & CStr(oRecordset.Fields("JourneyID").Value) & ")"
				If CLng(oRecordset.Fields("WorkingHours").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.WorkingHours=" & CStr(oRecordset.Fields("WorkingHours").Value) & ")"
				If CLng(oRecordset.Fields("AdditionalShift").Value) > 0 Then sCondition = sCondition & " And ((Employees.StartHour3>0) Or (Employees.EndHour3>0))"
				If CLng(oRecordset.Fields("LevelID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.LevelID=" & CStr(oRecordset.Fields("LevelID").Value) & ")"
				If CLng(oRecordset.Fields("EconomicZoneID").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.EconomicZoneID=" & CStr(oRecordset.Fields("EconomicZoneID").Value) & ")"
				If CLng(oRecordset.Fields("ServiceID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.ServiceID=" & CStr(oRecordset.Fields("ServiceID").Value) & ")"
				If CLng(oRecordset.Fields("AntiquityID").Value) > -1 Then sCondition = sCondition & " And (Employees.AntiquityID=" & CStr(oRecordset.Fields("AntiquityID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity2ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity2ID=" & CStr(oRecordset.Fields("Antiquity2ID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity3ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity3ID=" & CStr(oRecordset.Fields("Antiquity3ID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity4ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity4ID=" & CStr(oRecordset.Fields("Antiquity4ID").Value) & ")"
				If CLng(oRecordset.Fields("ForRisk").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=EmployeesRisksLKP.EmployeeID)"
				If CLng(oRecordset.Fields("GenderID").Value) > -1 Then sCondition = sCondition & " And (Employees.GenderID=" & CStr(oRecordset.Fields("GenderID").Value) & ")"
				If CLng(oRecordset.Fields("HasChildren").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=EmployeesChildrenLKP.EmployeeID)"
				If CLng(oRecordset.Fields("HasSyndicate").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=EmployeesSyndicatesLKP.EmployeeID)"
				If CLng(oRecordset.Fields("PositionID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & ")"
''					If InStr(1, sCondition, "(EmployeesHistoryList.JobID=Jobs.JobID)", vbBinaryCompare) = 0 Then sCondition = sCondition & " And (EmployeesHistoryList.JobID=Jobs.JobID)"
''					sCondition = sCondition & " And (Jobs.PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & ")"
''				End If
				If InStr(1, sCondition, "(Employees.", vbBinaryCompare) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID)"

				lErrorNumber = AppendTextToFile(PAYROLL_FILE2_PATH, sCondition, sErrorDescription)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			lErrorNumber = AppendTextToFile(PAYROLL_FILE2_PATH, "", sErrorDescription)
		End If
	End If

	If (iFileType = 3) Or (iFileType = -1) Then
		sErrorDescription = "No se pudieron obtener los conceptos de pagos y sus montos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptsValues.*, Concepts.PeriodID, Antiquities.StartYears, Antiquities.EndYears, Antiquities2.StartYears As StartYears2, Antiquities2.EndYears As EndYears2, Antiquities3.StartYears As StartYears3, Antiquities3.EndYears As EndYears3, Antiquities4.StartYears As StartYears4, Antiquities4.EndYears As EndYears4 From ConceptsValues, Concepts, Antiquities, Antiquities As Antiquities2, Antiquities As Antiquities3, Antiquities As Antiquities4 Where (ConceptsValues.ConceptID=Concepts.ConceptID) And (ConceptsValues.AntiquityID=Antiquities.AntiquityID) And (ConceptsValues.Antiquity2ID=Antiquities2.AntiquityID) And (ConceptsValues.Antiquity3ID=Antiquities3.AntiquityID) And (ConceptsValues.Antiquity4ID=Antiquities4.AntiquityID) And (ConceptsValues.StartDate<=" & lPayrollID & ") And (ConceptsValues.EndDate>=" & lPayrollID & ") And (Concepts.StartDate<=" & lPayrollID & ") And (Concepts.EndDate>=" & lPayrollID & ") And (ConceptQttyID In (3,13)) " & sConceptCondition & " Order By Concepts.OrderInList, Concepts.ConceptID, CompanyID Desc, EmployeeTypeID Desc, PositionTypeID Desc, EmployeeStatusID Desc, JobStatusID Desc, ClassificationID Desc, GroupGradeLevelID Desc, IntegrationID Desc, JourneyID Desc, WorkingHours Desc, AdditionalShift Desc, LevelID Desc, EconomicZoneID Desc, ServiceID Desc, ConceptsValues.AntiquityID Desc, ConceptsValues.Antiquity2ID Desc, ConceptsValues.Antiquity3ID Desc, ConceptsValues.Antiquity4ID Desc, ForRisk Desc, GenderID Desc, HasSyndicate Desc", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If FileExists(PAYROLL_FILE3_PATH, sErrorDescription) Then Call DeleteFile(PAYROLL_FILE3_PATH, "")
			Do While Not oRecordset.EOF
				sCondition = CStr(oRecordset.Fields("ConceptID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptAmount").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("PeriodID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptQttyID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMin").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMinQttyID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMax").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMaxQttyID").Value) & LIST_SEPARATOR
				If CLng(oRecordset.Fields("CompanyID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.CompanyID=" & CStr(oRecordset.Fields("CompanyID").Value) & ")"
				If CLng(oRecordset.Fields("EmployeeTypeID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeTypeID=" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & ")"
				If CLng(oRecordset.Fields("PositionTypeID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.PositionTypeID=" & CStr(oRecordset.Fields("PositionTypeID").Value) & ")"
				If CLng(oRecordset.Fields("EmployeeStatusID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.StatusID=" & CStr(oRecordset.Fields("EmployeeStatusID").Value) & ")"
				If CLng(oRecordset.Fields("JobStatusID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.JobID=Jobs.JobID) And (Jobs.StatusID=" & CStr(oRecordset.Fields("JobStatusID").Value) & ")"
				If CLng(oRecordset.Fields("ClassificationID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.ClassificationID=" & CStr(oRecordset.Fields("ClassificationID").Value) & ")"
				If CLng(oRecordset.Fields("GroupGradeLevelID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.GroupGradeLevelID=" & CStr(oRecordset.Fields("GroupGradeLevelID").Value) & ")"
				If CLng(oRecordset.Fields("IntegrationID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.IntegrationID=" & CStr(oRecordset.Fields("IntegrationID").Value) & ")"
				If CLng(oRecordset.Fields("JourneyID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.JourneyID=" & CStr(oRecordset.Fields("JourneyID").Value) & ")"
				If CLng(oRecordset.Fields("WorkingHours").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.WorkingHours=" & CStr(oRecordset.Fields("WorkingHours").Value) & ")"
				If CLng(oRecordset.Fields("AdditionalShift").Value) > 0 Then sCondition = sCondition & " And ((Employees.StartHour3>0) Or (Employees.EndHour3>0))"
				If CLng(oRecordset.Fields("LevelID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.LevelID=" & CStr(oRecordset.Fields("LevelID").Value) & ")"
				If CLng(oRecordset.Fields("EconomicZoneID").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.EconomicZoneID=" & CStr(oRecordset.Fields("EconomicZoneID").Value) & ")"
				If CLng(oRecordset.Fields("ServiceID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.ServiceID=" & CStr(oRecordset.Fields("ServiceID").Value) & ")"
				If CLng(oRecordset.Fields("AntiquityID").Value) > -1 Then sCondition = sCondition & " And (Employees.AntiquityID=" & CStr(oRecordset.Fields("AntiquityID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity2ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity2ID=" & CStr(oRecordset.Fields("Antiquity2ID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity3ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity3ID=" & CStr(oRecordset.Fields("Antiquity3ID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity4ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity4ID=" & CStr(oRecordset.Fields("Antiquity4ID").Value) & ")"
				If CLng(oRecordset.Fields("ForRisk").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=EmployeesRisksLKP.EmployeeID)"
				If CLng(oRecordset.Fields("GenderID").Value) > -1 Then sCondition = sCondition & " And (Employees.GenderID=" & CStr(oRecordset.Fields("GenderID").Value) & ")"
				If CLng(oRecordset.Fields("HasChildren").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=EmployeesChildrenLKP.EmployeeID)"
				If CLng(oRecordset.Fields("HasSyndicate").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=EmployeesSyndicatesLKP.EmployeeID)"
				If CLng(oRecordset.Fields("PositionID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & ")"
''					If InStr(1, sCondition, "(EmployeesHistoryList.JobID=Jobs.JobID)", vbBinaryCompare) = 0 Then sCondition = sCondition & " And (EmployeesHistoryList.JobID=Jobs.JobID)"
''					sCondition = sCondition & " And (Jobs.PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & ")"
''				End If
				If InStr(1, sCondition, "(Employees.", vbBinaryCompare) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID)"

				lErrorNumber = AppendTextToFile(PAYROLL_FILE3_PATH, sCondition, sErrorDescription)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			lErrorNumber = AppendTextToFile(PAYROLL_FILE3_PATH, "", sErrorDescription)
		End If
	End If

	If (iFileType = 8) Or (iFileType = -1) Then
		sErrorDescription = "No se pudieron obtener los conceptos de pagos y sus montos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptsValues.*, Concepts.PeriodID, Antiquities.StartYears, Antiquities.EndYears, Antiquities2.StartYears As StartYears2, Antiquities2.EndYears As EndYears2, Antiquities3.StartYears As StartYears3, Antiquities3.EndYears As EndYears3, Antiquities4.StartYears As StartYears4, Antiquities4.EndYears As EndYears4 From ConceptsValues, Concepts, Antiquities, Antiquities As Antiquities2, Antiquities As Antiquities3, Antiquities As Antiquities4 Where (ConceptsValues.ConceptID=Concepts.ConceptID) And (ConceptsValues.AntiquityID=Antiquities.AntiquityID) And (ConceptsValues.Antiquity2ID=Antiquities2.AntiquityID) And (ConceptsValues.Antiquity3ID=Antiquities3.AntiquityID) And (ConceptsValues.Antiquity4ID=Antiquities4.AntiquityID) And (ConceptsValues.StartDate<=" & lPayrollID & ") And (ConceptsValues.EndDate>=" & lPayrollID & ") And (Concepts.StartDate<=" & lPayrollID & ") And (Concepts.EndDate>=" & lPayrollID & ") And (ConceptQttyID In (8,9)) " & sConceptCondition & " Order By Concepts.OrderInList, Concepts.ConceptID, CompanyID Desc, EmployeeTypeID Desc, PositionTypeID Desc, EmployeeStatusID Desc, JobStatusID Desc, ClassificationID Desc, GroupGradeLevelID Desc, IntegrationID Desc, JourneyID Desc, WorkingHours Desc, AdditionalShift Desc, LevelID Desc, EconomicZoneID Desc, ServiceID Desc, ConceptsValues.AntiquityID Desc, ConceptsValues.Antiquity2ID Desc, ConceptsValues.Antiquity3ID Desc, ConceptsValues.Antiquity4ID Desc, ForRisk Desc, GenderID Desc, HasSyndicate Desc", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If FileExists(PAYROLL_FILE8_PATH, sErrorDescription) Then Call DeleteFile(PAYROLL_FILE8_PATH, "")
			Do While Not oRecordset.EOF
				sCondition = CStr(oRecordset.Fields("ConceptID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptAmount").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("PeriodID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptQttyID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("AppliesToID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMin").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMinQttyID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMax").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMaxQttyID").Value) & LIST_SEPARATOR
				If CLng(oRecordset.Fields("CompanyID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.CompanyID=" & CStr(oRecordset.Fields("CompanyID").Value) & ")"
				If CLng(oRecordset.Fields("EmployeeTypeID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeTypeID=" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & ")"
				If CLng(oRecordset.Fields("PositionTypeID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.PositionTypeID=" & CStr(oRecordset.Fields("PositionTypeID").Value) & ")"
				If CLng(oRecordset.Fields("EmployeeStatusID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.StatusID=" & CStr(oRecordset.Fields("EmployeeStatusID").Value) & ")"
				If CLng(oRecordset.Fields("JobStatusID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.JobID=Jobs.JobID) And (Jobs.StatusID=" & CStr(oRecordset.Fields("JobStatusID").Value) & ")"
				If CLng(oRecordset.Fields("ClassificationID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.ClassificationID=" & CStr(oRecordset.Fields("ClassificationID").Value) & ")"
				If CLng(oRecordset.Fields("GroupGradeLevelID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.GroupGradeLevelID=" & CStr(oRecordset.Fields("GroupGradeLevelID").Value) & ")"
				If CLng(oRecordset.Fields("IntegrationID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.IntegrationID=" & CStr(oRecordset.Fields("IntegrationID").Value) & ")"
				If CLng(oRecordset.Fields("JourneyID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.JourneyID=" & CStr(oRecordset.Fields("JourneyID").Value) & ")"
				If CLng(oRecordset.Fields("WorkingHours").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.WorkingHours=" & CStr(oRecordset.Fields("WorkingHours").Value) & ")"
				If CLng(oRecordset.Fields("AdditionalShift").Value) > 0 Then sCondition = sCondition & " And ((Employees.StartHour3>0) Or (Employees.EndHour3>0))"
				If CLng(oRecordset.Fields("LevelID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.LevelID=" & CStr(oRecordset.Fields("LevelID").Value) & ")"
				If CLng(oRecordset.Fields("EconomicZoneID").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.EconomicZoneID=" & CStr(oRecordset.Fields("EconomicZoneID").Value) & ")"
				If CLng(oRecordset.Fields("ServiceID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.ServiceID=" & CStr(oRecordset.Fields("ServiceID").Value) & ")"
				If CLng(oRecordset.Fields("AntiquityID").Value) > -1 Then sCondition = sCondition & " And (Employees.AntiquityID=" & CStr(oRecordset.Fields("AntiquityID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity2ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity2ID=" & CStr(oRecordset.Fields("Antiquity2ID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity3ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity3ID=" & CStr(oRecordset.Fields("Antiquity3ID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity4ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity4ID=" & CStr(oRecordset.Fields("Antiquity4ID").Value) & ")"
				If CLng(oRecordset.Fields("ForRisk").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=EmployeesRisksLKP.EmployeeID)"
				If CLng(oRecordset.Fields("GenderID").Value) > -1 Then sCondition = sCondition & " And (Employees.GenderID=" & CStr(oRecordset.Fields("GenderID").Value) & ")"
				If CLng(oRecordset.Fields("HasChildren").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=EmployeesChildrenLKP.EmployeeID)"
				If CLng(oRecordset.Fields("HasSyndicate").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=EmployeesSyndicatesLKP.EmployeeID)"
				If CLng(oRecordset.Fields("PositionID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & ")"
''					If InStr(1, sCondition, "(EmployeesHistoryList.JobID=Jobs.JobID)", vbBinaryCompare) = 0 Then sCondition = sCondition & " And (EmployeesHistoryList.JobID=Jobs.JobID)"
''					sCondition = sCondition & " And (Jobs.PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & ")"
''				End If
				If InStr(1, sCondition, "(Employees.", vbBinaryCompare) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID)"

				lErrorNumber = AppendTextToFile(PAYROLL_FILE8_PATH, sCondition, sErrorDescription)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			lErrorNumber = AppendTextToFile(PAYROLL_FILE8_PATH, "", sErrorDescription)
		End If
	End If

	If (iFileType = 15) Or (iFileType = -1) Then
		sErrorDescription = "No se pudieron obtener los conceptos de pagos y sus montos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptsValues.*, Concepts.PeriodID, Antiquities.StartYears, Antiquities.EndYears, Antiquities2.StartYears As StartYears2, Antiquities2.EndYears As EndYears2, Antiquities3.StartYears As StartYears3, Antiquities3.EndYears As EndYears3, Antiquities4.StartYears As StartYears4, Antiquities4.EndYears As EndYears4 From ConceptsValues, Concepts, Antiquities, Antiquities As Antiquities2, Antiquities As Antiquities3, Antiquities As Antiquities4 Where (ConceptsValues.ConceptID=Concepts.ConceptID) And (ConceptsValues.AntiquityID=Antiquities.AntiquityID) And (ConceptsValues.Antiquity2ID=Antiquities2.AntiquityID) And (ConceptsValues.Antiquity3ID=Antiquities3.AntiquityID) And (ConceptsValues.Antiquity4ID=Antiquities4.AntiquityID) And (ConceptsValues.StartDate<=" & lPayrollID & ") And (ConceptsValues.EndDate>=" & lPayrollID & ") And (Concepts.StartDate<=" & lPayrollID & ") And (Concepts.EndDate>=" & lPayrollID & ") And (ConceptsValues.ConceptID In (" & SCHOOLARSHIP_CONCEPTS_FOR_PAYROLL & ")) " & sConceptCondition & " Order By Concepts.OrderInList, Concepts.ConceptID, CompanyID Desc, EmployeeTypeID Desc, PositionTypeID Desc, EmployeeStatusID Desc, JobStatusID Desc, ClassificationID Desc, GroupGradeLevelID Desc, IntegrationID Desc, JourneyID Desc, WorkingHours Desc, AdditionalShift Desc, LevelID Desc, EconomicZoneID Desc, ServiceID Desc, ConceptsValues.AntiquityID Desc, ConceptsValues.Antiquity2ID Desc, ConceptsValues.Antiquity3ID Desc, ConceptsValues.Antiquity4ID Desc, ForRisk Desc, GenderID Desc, SchoolarshipID Desc, HasSyndicate Desc", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If FileExists(PAYROLL_FILE15_PATH, sErrorDescription) Then Call DeleteFile(PAYROLL_FILE15_PATH, "")
			Do While Not oRecordset.EOF
				sCondition = CStr(oRecordset.Fields("ConceptID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptAmount").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("PeriodID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMin").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMinQttyID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMax").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMaxQttyID").Value) & LIST_SEPARATOR
				If CLng(oRecordset.Fields("CompanyID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.CompanyID=" & CStr(oRecordset.Fields("CompanyID").Value) & ")"
				If CLng(oRecordset.Fields("EmployeeTypeID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeTypeID=" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & ")"
				If CLng(oRecordset.Fields("PositionTypeID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.PositionTypeID=" & CStr(oRecordset.Fields("PositionTypeID").Value) & ")"
				If CLng(oRecordset.Fields("EmployeeStatusID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.StatusID=" & CStr(oRecordset.Fields("EmployeeStatusID").Value) & ")"
				If CLng(oRecordset.Fields("JobStatusID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.JobID=Jobs.JobID) And (Jobs.StatusID=" & CStr(oRecordset.Fields("JobStatusID").Value) & ")"
				If CLng(oRecordset.Fields("ClassificationID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.ClassificationID=" & CStr(oRecordset.Fields("ClassificationID").Value) & ")"
				If CLng(oRecordset.Fields("GroupGradeLevelID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.GroupGradeLevelID=" & CStr(oRecordset.Fields("GroupGradeLevelID").Value) & ")"
				If CLng(oRecordset.Fields("IntegrationID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.IntegrationID=" & CStr(oRecordset.Fields("IntegrationID").Value) & ")"
				If CLng(oRecordset.Fields("JourneyID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.JourneyID=" & CStr(oRecordset.Fields("JourneyID").Value) & ")"
				If CLng(oRecordset.Fields("WorkingHours").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.WorkingHours=" & CStr(oRecordset.Fields("WorkingHours").Value) & ")"
				If CLng(oRecordset.Fields("AdditionalShift").Value) > 0 Then sCondition = sCondition & " And ((Employees.StartHour3>0) Or (Employees.EndHour3>0))"
				If CLng(oRecordset.Fields("LevelID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.LevelID=" & CStr(oRecordset.Fields("LevelID").Value) & ")"
				If CLng(oRecordset.Fields("EconomicZoneID").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.EconomicZoneID=" & CStr(oRecordset.Fields("EconomicZoneID").Value) & ")"
				If CLng(oRecordset.Fields("ServiceID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.ServiceID=" & CStr(oRecordset.Fields("ServiceID").Value) & ")"
				If CLng(oRecordset.Fields("AntiquityID").Value) > -1 Then sCondition = sCondition & " And (Employees.AntiquityID=" & CStr(oRecordset.Fields("AntiquityID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity2ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity2ID=" & CStr(oRecordset.Fields("Antiquity2ID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity3ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity3ID=" & CStr(oRecordset.Fields("Antiquity3ID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity4ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity4ID=" & CStr(oRecordset.Fields("Antiquity4ID").Value) & ")"
				If CLng(oRecordset.Fields("ForRisk").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=EmployeesRisksLKP.EmployeeID)"
				If CLng(oRecordset.Fields("GenderID").Value) > -1 Then sCondition = sCondition & " And (Employees.GenderID=" & CStr(oRecordset.Fields("GenderID").Value) & ")"
				If CLng(oRecordset.Fields("SchoolarshipID").Value) > -1 Then sCondition = sCondition & " And (EmployeesChildrenLKP.LevelID=" & CStr(oRecordset.Fields("SchoolarshipID").Value) & ")"
				If CLng(oRecordset.Fields("HasSyndicate").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=EmployeesSyndicatesLKP.EmployeeID)"
				If CLng(oRecordset.Fields("PositionID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & ")"
''					If InStr(1, sCondition, "(EmployeesHistoryList.JobID=Jobs.JobID)", vbBinaryCompare) = 0 Then sCondition = sCondition & " And (EmployeesHistoryList.JobID=Jobs.JobID)"
''					sCondition = sCondition & " And (Jobs.PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & ")"
''				End If
				If InStr(1, sCondition, "(Employees.", vbBinaryCompare) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID)"

				lErrorNumber = AppendTextToFile(PAYROLL_FILE15_PATH, sCondition, sErrorDescription)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			lErrorNumber = AppendTextToFile(PAYROLL_FILE15_PATH, "", sErrorDescription)
		End If
	End If

	If (iFileType = 69) Or (iFileType = -1) Then
		sErrorDescription = "No se pudieron obtener los conceptos de pagos y sus montos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptsValues.*, Concepts.PeriodID, Antiquities.StartYears, Antiquities.EndYears, Antiquities2.StartYears As StartYears2, Antiquities2.EndYears As EndYears2, Antiquities3.StartYears As StartYears3, Antiquities3.EndYears As EndYears3, Antiquities4.StartYears As StartYears4, Antiquities4.EndYears As EndYears4 From ConceptsValues, Concepts, Antiquities, Antiquities As Antiquities2, Antiquities As Antiquities3, Antiquities As Antiquities4 Where (ConceptsValues.ConceptID=Concepts.ConceptID) And (ConceptsValues.AntiquityID=Antiquities.AntiquityID) And (ConceptsValues.Antiquity2ID=Antiquities2.AntiquityID) And (ConceptsValues.Antiquity3ID=Antiquities3.AntiquityID) And (ConceptsValues.Antiquity4ID=Antiquities4.AntiquityID) And (ConceptsValues.StartDate<=" & lPayrollID & ") And (ConceptsValues.EndDate>=" & lPayrollID & ") And (Concepts.StartDate<=" & lPayrollID & ") And (Concepts.EndDate>=" & lPayrollID & ") And (ConceptQttyID=2) And (Concepts.ConceptID In (70)) " & sConceptCondition & " Order By Concepts.OrderInList, Concepts.ConceptID, CompanyID Desc, EmployeeTypeID Desc, PositionTypeID Desc, EmployeeStatusID Desc, JobStatusID Desc, ClassificationID Desc, GroupGradeLevelID Desc, IntegrationID Desc, JourneyID Desc, WorkingHours Desc, AdditionalShift Desc, LevelID Desc, EconomicZoneID Desc, ServiceID Desc, ConceptsValues.AntiquityID Desc, ConceptsValues.Antiquity2ID Desc, ConceptsValues.Antiquity3ID Desc, ConceptsValues.Antiquity4ID Desc, ForRisk Desc, GenderID Desc, HasSyndicate Desc", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If FileExists(PAYROLL_FILE69_PATH, sErrorDescription) Then Call DeleteFile(PAYROLL_FILE69_PATH, "")
			Do While Not oRecordset.EOF
				sCondition = CStr(oRecordset.Fields("ConceptID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptAmount").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("PeriodID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("AppliesToID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMin").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMinQttyID").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMax").Value) & LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptMaxQttyID").Value) & LIST_SEPARATOR
				If CLng(oRecordset.Fields("CompanyID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.CompanyID=" & CStr(oRecordset.Fields("CompanyID").Value) & ")"
				If CLng(oRecordset.Fields("EmployeeTypeID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeTypeID=" & CStr(oRecordset.Fields("EmployeeTypeID").Value) & ")"
				If CLng(oRecordset.Fields("PositionTypeID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.PositionTypeID=" & CStr(oRecordset.Fields("PositionTypeID").Value) & ")"
				If CLng(oRecordset.Fields("EmployeeStatusID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.StatusID=" & CStr(oRecordset.Fields("EmployeeStatusID").Value) & ")"
				If CLng(oRecordset.Fields("JobStatusID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.JobID=Jobs.JobID) And (Jobs.StatusID=" & CStr(oRecordset.Fields("JobStatusID").Value) & ")"
				If CLng(oRecordset.Fields("ClassificationID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.ClassificationID=" & CStr(oRecordset.Fields("ClassificationID").Value) & ")"
				If CLng(oRecordset.Fields("GroupGradeLevelID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.GroupGradeLevelID=" & CStr(oRecordset.Fields("GroupGradeLevelID").Value) & ")"
				If CLng(oRecordset.Fields("IntegrationID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.IntegrationID=" & CStr(oRecordset.Fields("IntegrationID").Value) & ")"
				If CLng(oRecordset.Fields("JourneyID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.JourneyID=" & CStr(oRecordset.Fields("JourneyID").Value) & ")"
				If CLng(oRecordset.Fields("WorkingHours").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.WorkingHours=" & CStr(oRecordset.Fields("WorkingHours").Value) & ")"
				If CLng(oRecordset.Fields("AdditionalShift").Value) > 0 Then sCondition = sCondition & " And ((Employees.StartHour3>0) Or (Employees.EndHour3>0))"
				If CLng(oRecordset.Fields("LevelID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.LevelID=" & CStr(oRecordset.Fields("LevelID").Value) & ")"
				If CLng(oRecordset.Fields("EconomicZoneID").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.AreaID=Areas.AreaID) And (Areas.EconomicZoneID=" & CStr(oRecordset.Fields("EconomicZoneID").Value) & ")"
				If CLng(oRecordset.Fields("ServiceID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.ServiceID=" & CStr(oRecordset.Fields("ServiceID").Value) & ")"
				If CLng(oRecordset.Fields("AntiquityID").Value) > -1 Then sCondition = sCondition & " And (Employees.AntiquityID=" & CStr(oRecordset.Fields("AntiquityID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity2ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity2ID=" & CStr(oRecordset.Fields("Antiquity2ID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity3ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity3ID=" & CStr(oRecordset.Fields("Antiquity3ID").Value) & ")"
				If CLng(oRecordset.Fields("Antiquity4ID").Value) > -1 Then sCondition = sCondition & " And (Employees.Antiquity4ID=" & CStr(oRecordset.Fields("Antiquity4ID").Value) & ")"
				If CLng(oRecordset.Fields("ForRisk").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=EmployeesRisksLKP.EmployeeID)"
				If CLng(oRecordset.Fields("GenderID").Value) > -1 Then sCondition = sCondition & " And (Employees.GenderID=" & CStr(oRecordset.Fields("GenderID").Value) & ")"
				If CLng(oRecordset.Fields("HasChildren").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=EmployeesChildrenLKP.EmployeeID)"
				If CLng(oRecordset.Fields("HasSyndicate").Value) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=EmployeesSyndicatesLKP.EmployeeID)"
				If CLng(oRecordset.Fields("PositionID").Value) > -1 Then sCondition = sCondition & " And (EmployeesHistoryList.PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & ")"
''					If InStr(1, sCondition, "(EmployeesHistoryList.JobID=Jobs.JobID)", vbBinaryCompare) = 0 Then sCondition = sCondition & " And (EmployeesHistoryList.JobID=Jobs.JobID)"
''					sCondition = sCondition & " And (Jobs.PositionID=" & CStr(oRecordset.Fields("PositionID").Value) & ")"
''				End If
				If InStr(1, sCondition, "(Employees.", vbBinaryCompare) > 0 Then sCondition = sCondition & " And (EmployeesHistoryList.EmployeeID=Employees.EmployeeID)"

				lErrorNumber = AppendTextToFile(PAYROLL_FILE69_PATH, sCondition, sErrorDescription)
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			lErrorNumber = AppendTextToFile(PAYROLL_FILE69_PATH, "", sErrorDescription)
		End If
	End If

	CreateConceptsFile = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeeTaxActivationForm(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the information for tax activation using
'		  a HTML Form
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeTaxActivationForm"
	Dim lErrorNumber

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "var iTaxActivation = 0;" & vbNewLine

		Response.Write "function CheckPayrollFields(oForm) {" & vbNewLine
			Response.Write "if (oForm) {" & vbNewLine
'				Response.Write "if (oForm.EmployeeID.value.length == 0) {" & vbNewLine
'					Response.Write "alert('Favor de introducir el número del empleado y de validarlo.');" & vbNewLine
'					Response.Write "oForm.EmployeeNumber.focus();" & vbNewLine
'					Response.Write "return false;" & vbNewLine
'				Response.Write "}" & vbNewLine

				Response.Write "if (oForm.EmployeeIDs.value.length == 0) {" & vbNewLine
					Response.Write "alert('Favor de introducir el número del empleado.');" & vbNewLine
					Response.Write "oForm.EmployeeIDs.focus();" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "}" & vbNewLine

			Response.Write "return true;" & vbNewLine
		Response.Write "} // End of CheckPayrollFields" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine
	Response.Write "<FORM NAME=""PayrollFrm"" ID=""PayrollFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""GET"" onSubmit=""return CheckPayrollFields(this)"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & oRequest("SectionID").Item & """ />"
'		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE="""" />"

		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
'			Response.Write "<TR>"
'				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
'				Response.Write "<TD>"
'					Response.Write "<INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & oRequest("EmployeeID").Item & """ CLASS=""TextFields"" />"
'					Response.Write "<A HREF=""javascript: SearchRecord(document.PayrollFrm.EmployeeNumber.value, 'EmployeeNumber', 'SearchEmployeeNumberIFrame', 'PayrollFrm.EmployeeID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>"
'					Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""320"" HEIGHT=""22""></IFRAME>"
'				Response.Write "</TD>"
'			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Número del(os) empleado(s):<BR /></FONT>"
					Response.Write "<TEXTAREA NAME=""EmployeeIDs"" ID=""EmployeeIDsTxtArea"" ROWS=""5"" COLS=""60"" MAXLENGTH=""2000"" CLASS=""TextFields"">" & Replace(Replace(Replace(Replace(oRequest("EmployeeIDs").Item, vbNewLine & vbNewLine, ","), vbNewLine, ","), " ", ""), ",,", ",") & "</TEXTAREA><BR /><BR />"
				Response.Write "</TD>"
			Response.Write "</TR>"

			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2"">"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Año a procesar:&nbsp;</FONT>"
					Response.Write "<SELECT NAME=""YearID"" ID=""YearIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write "<OPTION VALUE=""" & Year(Date()) & """ SELECTED=""1"">" & Year(Date()) & "</OPTION>"
						If Month(Date()) < 3 Then Response.Write "<OPTION VALUE=""" & Year(Date()) - 1& """ SELECTED=""1"">" & Year(Date()) - 1 & "</OPTION>"
					Response.Write "</SELECT><BR /><BR />"
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2"">"
					Response.Write "<INPUT TYPE=""RADIO"" NAME=""TaxActivation"" ID=""TaxActivationRd"" VALUE=""0"" "
						If StrComp(oRequest("TaxActivation").Item, "1", vbBinaryCompare) <> 0 Then Response.Write " CHECKED=""1"""
					Response.Write " onClick=""iTaxActivation = 0;"" /> No se aplicará el recálculo anual.<BR />"
					Response.Write "<INPUT TYPE=""RADIO"" NAME=""TaxActivation"" ID=""TaxActivationRd"" VALUE=""1"" "
						If StrComp(oRequest("TaxActivation").Item, "1", vbBinaryCompare) = 0 Then Response.Write " CHECKED=""1"""
					Response.Write " onClick=""iTaxActivation = 1;"" /> Sí se aplicará el recálculo anual.<BR />"
				Response.Write "</FONT></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE><BR />"

		Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ActivateTax"" ID=""ActivateTaxBtn"" VALUE=""Actualizar Estatus"" CLASS=""Buttons"" />"
		Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
		Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Export"" ID=""ExportBtn"" VALUE=""Obtener Listado"" CLASS=""Buttons"" onClick=""var oExcel = OpenNewWindow('Export.asp?Excel=1&Action=EmployeesForTaxAdjustment&YearID=' + document.PayrollFrm.YearID.value + '&TaxActivation=' + iTaxActivation + '&AccessKey=" & aLoginComponent(S_ACCESS_KEY_LOGIN) & "&SIAP_SectionID=" & Request.Cookies("SIAP_SectionID") & "', '', 'ExportToExcel', 640, 480, 'yes', 'yes');"" />"
		Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
		Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?SectionID=" & oRequest("SectionID").Item & "'"" />"
	Response.Write "</FORM>"

	DisplayEmployeeTaxActivationForm = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeeTaxAdjustmentForm(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the information for tax adjustment using
'		  a HTML Form
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeTaxAdjustmentForm"
	Dim lErrorNumber
	Dim iIndex

	Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
		Response.Write "function CheckPayrollFields(oForm) {" & vbNewLine
			Response.Write "if (oForm) {" & vbNewLine
				Response.Write "if (oForm.EmployeeID.value.length == 0) {" & vbNewLine
					Response.Write "alert('Favor de introducir el número del empleado y de validarlo.');" & vbNewLine
					Response.Write "oForm.EmployeeNumber.focus();" & vbNewLine
					Response.Write "return false;" & vbNewLine
				Response.Write "}" & vbNewLine

				Response.Write "if (! CheckFloatValue(oForm.TaxAmount, 'el ajuste anual del impuesto sobre la renta', N_NO_RANK_FLAG, N_CLOSED_FLAG, 0, 0))" & vbNewLine
					Response.Write "return false;" & vbNewLine
			Response.Write "}" & vbNewLine

			Response.Write "return true;" & vbNewLine
		Response.Write "} // End of CheckPayrollFields" & vbNewLine
	Response.Write "//--></SCRIPT>" & vbNewLine
	Response.Write "<FORM NAME=""PayrollFrm"" ID=""PayrollFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""GET"" onSubmit=""return CheckPayrollFields(this)"">"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""SectionID"" ID=""SectionIDHdn"" VALUE=""" & oRequest("SectionID").Item & """ />"
		Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE="""" />"

		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
				Response.Write "<TD>"
					Response.Write "<INPUT TYPE=""TEXT"" NAME=""EmployeeNumber"" ID=""EmployeeNumberTxt"" SIZE=""6"" MAXLENGTH=""6"" VALUE=""" & oRequest("EmployeeID").Item & """ CLASS=""TextFields"" />"
					Response.Write "<A HREF=""javascript: SearchRecord(document.PayrollFrm.EmployeeNumber.value, 'EmployeeNumber', 'SearchEmployeeNumberIFrame', 'PayrollFrm.EmployeeID')""><IMG SRC=""Images/IcnSearch.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Validar el número de empleado"" BORDER=""0"" ALIGN=""ABSMIDDLE"" HSPACE=""10"" /></A>"
					Response.Write "<IFRAME SRC=""SearchRecord.asp"" NAME=""SearchEmployeeNumberIFrame"" FRAMEBORDER=""0"" WIDTH=""320"" HEIGHT=""22""></IFRAME>"
				Response.Write "</TD>"
			Response.Write "</TR>"

			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Año a procesar:&nbsp;</FONT></TD>"
				Response.Write "<TD><SELECT NAME=""YearID"" ID=""YearIDCmb"" SIZE=""1"" CLASS=""Lists"">"
					For iIndex = 2009 To Year(Date())
						Response.Write "<OPTION VALUE=""" & iIndex & """"
							If Len(oRequest("YearID").Item) > 0 Then
								If iIndex = CInt(oRequest("YearID").Item) Then Response.Write " SELECTED=""1"""
							Else
								If iIndex = Year(Date()) Then Response.Write " SELECTED=""1"""
							End If
						Response.Write ">" & iIndex & "</OPTION>"
					Next
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Ajuste a aplicar:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""TaxAmount"" ID=""TaxAmountTxt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & oRequest("TaxAmount").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
		Response.Write "</TABLE><BR />"

		Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""ApplyTax"" ID=""ApplyTaxBtn"" VALUE=""Aplicar Ajuste"" CLASS=""Buttons"" />"
		Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
		Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?SectionID=" & oRequest("SectionID").Item & "'"" />"
	Response.Write "</FORM>"

	DisplayEmployeeTaxAdjustmentForm = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeeTaxAdjustmentTable(oRequest, oADODBConnection, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the information for tax adjustment using
'		  a HTML Table
'Inputs:  oRequest, oADODBConnection, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeeTaxAdjustmentTable"
	Dim sCondition
	Dim oRecordset
	Dim iCounter
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sCondition = ""
	If Len(oRequest("YearID").Item) > 0 Then sCondition = sCondition & " And (PayrollYear=" & oRequest("YearID").Item & ")"
	If Len(oRequest("TaxActivation").Item) > 0 Then sCondition = sCondition & " And (bTaxAdjustment=" & oRequest("TaxActivation").Item & ")"
	sErrorDescription = "No se pudo obtener la información de los empleados para el recálculo anual del impuesto."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, AreaCode, AreaShortName, AreaName, EconomicZoneID, PositionShortName, PositionName From EmployeesForTaxAdjustment, Employees, Jobs, Areas, Positions Where (EmployeesForTaxAdjustment.EmployeeID=Employees.EmployeeID) And (Employees.JobID=Jobs.JobID) And (Jobs.AreaID=Areas.AreaID) And (Jobs.PositionID=Positions.PositionID) " & sCondition & " Order By EmployeeNumber", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>Empleados a los que "
				If CInt(oRequest("TaxActivation").Item) = 1 Then
					Response.Write "sí"
				Else
					Response.Write "no"
				End If
			Response.Write "se les recalculará el impuesto anual para el año " & oRequest("YearID").Item & "</B></FONT><BR /><BR />"
			Response.Write "<TABLE BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("Número de empleado,Apellido paterno,Apellido materno,Nombre,Adscripción,Centro de trabajo,Zona geográfica,<SPAN COLS=""2"" />Puesto", ",", -1, vbBinaryCompare)
				asCellWidths = Split("100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				iCounter = 0
				asCellWidths = Split(",,,,,,CENTER,,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sRowContents = ""
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
					If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
						sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName2").Value))
					Else
						sRowContents = sRowContents & " "
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("AreaName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EconomicZoneID").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionShortName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PositionName").Value))

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If
					oRecordset.MoveNext
					If (Err.number <> 0) Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen registros en la base de datos que cumplan con los criterios del filtro."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayEmployeeTaxAdjustmentTable = lErrorNumber
	Err.Clear
End Function

Function DisplayPayrollForm(oRequest, oADODBConnection, sAction, aPayrollComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about a payroll from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aPayrollComponent
'Outputs: aPayrollComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPayrollForm"
	Dim sNames
	Dim iTemp
	Dim lErrorNumber

	If aPayrollComponent(N_ID_PAYROLL) <> -1 Then
		lErrorNumber = GetPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckPayrollFields(oForm) {" & vbNewLine
				If Len(oRequest("Delete").Item) = 0 Then
					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "if (oForm.PayrollName.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir el nombre de la nómina.');" & vbNewLine
							Response.Write "oForm.PayrollName.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (oForm.PayrollCLC.value.length == 0) {" & vbNewLine
							Response.Write "oForm.PayrollCLC.value = '.';" & vbNewLine
'							Response.Write "alert('Favor de introducir la CLC de la nómina.');" & vbNewLine
'							Response.Write "oForm.PayrollCLC.focus();" & vbNewLine
'							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
					Response.Write "}" & vbNewLine
				End If
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckPayrollFields" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""PayrollFrm"" ID=""PayrollFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckPayrollFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""" & oRequest("Action").Item & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PayrollID"" ID=""PayrollIDHdn"" VALUE=""" & aPayrollComponent(N_ID_PAYROLL) & """ />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Nombre:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PayrollName"" ID=""PayrollNameTxt"" SIZE=""34"" MAXLENGTH=""255"" VALUE=""" & aPayrollComponent(S_NAME_PAYROLL) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">"
						Response.Write DisplayDateCombosUsingSerial(aPayrollComponent(N_DATE_PAYROLL), "Payroll", (Year(Date()) - 1), (Year(Date()) + 1), True, False)
					Response.Write "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">CLC:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""PayrollCLC"" ID=""PayrollCLCTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE=""" & aPayrollComponent(S_CLC_PAYROLL) & """ CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de nómina:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""PayrollTypeID"" ID=""PayrollTypeIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""if ((this.value == '1') || (this.value == '4')) {ShowDisplay(document.all['PayrollTypeDiv']);} else {ShowDisplay(document.all['PayrollTypeDiv']);}"">"
						If aPayrollComponent(N_TYPE_ID_PAYROLL) = -1 Then
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PayrollTypes", "PayrollTypeID", "PayrollTypeName", "(PayrollTypeID<>1)", "PayrollTypeID", aPayrollComponent(N_TYPE_ID_PAYROLL), "Ninguno;;;-1", sErrorDescription)
						Else
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PayrollTypes", "PayrollTypeID", "PayrollTypeName", "(PayrollTypeID=1)", "PayrollTypeID", aPayrollComponent(N_TYPE_ID_PAYROLL), "Ninguno;;;-1", sErrorDescription)
						End If
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""PayrollTypeDiv"" ID=""PayrollTypeDiv"""
					If aPayrollComponent(N_TYPE_ID_PAYROLL) = 1 Then Response.Write " STYLE=""display: none"""
				Response.Write ">"
					Response.Write "<TD COLSPAN=""2"">"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Nómina para el acumulado anual:<BR /></FONT>"
						Response.Write "<SELECT NAME=""ForPayrollDate"" ID=""ForPayrollDateCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Payrolls", "PayrollID", "PayrollDate, PayrollName", "(PayrollTypeID=1)", "PayrollID Desc", aPayrollComponent(N_FOR_DATE_PAYROLL), "Ninguna;;;-1", sErrorDescription)
						Response.Write "</SELECT>"
					Response.Write "</TD>"
				Response.Write "</TR>"
			Response.Write "</TABLE>"

			If False Then 'aPayrollComponent(N_ID_PAYROLL) <> -1 Then
				Response.Write "<BR /><FONT FACE=""Arial"" SIZE=""2"">"
					Response.Write "¿El nómina está cerrada?<BR />"
					Response.Write "<INPUT TYPE=""Radio"" NAME=""IsClosed"" ID=""IsClosedRd"" VALUE=""1"""
						If aPayrollComponent(N_CLOSED_PAYROLL) = 1 Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " />Sí&nbsp;&nbsp;&nbsp;<INPUT TYPE=""Radio"" NAME=""IsClosed"" ID=""IsClosedRd"" VALUE=""0"""
						If aPayrollComponent(N_CLOSED_PAYROLL) = 0 Then
							Response.Write " CHECKED=""1"""
						End If
					Response.Write " />No<BR />"
				Response.Write "</FONT>"
			Else
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IsClosed"" ID=""IsClosedHdn"" VALUE=""" & aPayrollComponent(N_CLOSED_PAYROLL) & """ />"
			End If
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IsActive_1"" ID=""IsActive_1Hdn"" VALUE=""" & aPayrollComponent(N_IS_ACTIVE_1_PAYROLL) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IsActive_2"" ID=""IsActive_2Hdn"" VALUE=""" & aPayrollComponent(N_IS_ACTIVE_2_PAYROLL) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IsActive_3"" ID=""IsActive_3Hdn"" VALUE=""" & aPayrollComponent(N_IS_ACTIVE_3_PAYROLL) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IsActive_4"" ID=""IsActive_4Hdn"" VALUE=""" & aPayrollComponent(N_IS_ACTIVE_4_PAYROLL) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IsActive_5"" ID=""IsActive_5Hdn"" VALUE=""" & aPayrollComponent(N_IS_ACTIVE_5_PAYROLL) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IsActive_6"" ID=""IsActive_6Hdn"" VALUE=""" & aPayrollComponent(N_IS_ACTIVE_6_PAYROLL) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IsActive_7"" ID=""IsActive_7Hdn"" VALUE=""" & aPayrollComponent(N_IS_ACTIVE_7_PAYROLL) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IsActive_8"" ID=""IsActive_8Hdn"" VALUE=""" & aPayrollComponent(N_IS_ACTIVE_8_PAYROLL) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IsActive_9"" ID=""IsActive_9Hdn"" VALUE=""" & aPayrollComponent(N_IS_ACTIVE_9_PAYROLL) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IsActive_10"" ID=""IsActive_10Hdn"" VALUE=""" & aPayrollComponent(N_IS_ACTIVE_10_PAYROLL) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IsActive_11"" ID=""IsActive_11Hdn"" VALUE=""" & aPayrollComponent(N_IS_ACTIVE_11_PAYROLL) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""IsActive_12"" ID=""IsActive_12Hdn"" VALUE=""" & aPayrollComponent(N_IS_ACTIVE_12_PAYROLL) & """ />"
			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""280"" HEIGHT=""1"" /><BR /><BR />"

			If aPayrollComponent(N_ID_PAYROLL) = -1 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemovePayrollWngDiv']); PayrollFrm.Remove.focus()"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='Main_ISSSTE.asp?SectionID=4'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemovePayrollWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
		If aPayrollComponent(N_ID_PAYROLL) = -1 Then
			iTemp = 15
			If Day(Date()) > 16 Then
				Select Case Month(Date())
					Case 2
						iTemp = 28
						If (Year(Date()) Mod 4) = 0 Then iTemp = 29
					Case 1, 3, 5, 7, 8, 10, 12
						iTemp = 31
					Case Else
						iTemp = 30
				End Select
			End If
			sNames = "1a"
			If iTemp > 15 Then sNames = "2a"

			Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
				Response.Write "SendURLValuesToForm('PayrollDay=" & iTemp & "&PayrollMonth=" & Right(("0" & Month(Date())), Len("00")) & "&PayrollYear=" & Year(Date()) & "&PayrollName=" & sNames & " quincena de " & asMonthNames_es(Month(Date())) & " de " & Year(Date()) & "', document.PayrollFrm);" & vbNewLine
				Response.Write "ChangeDaysListGivenTheMonthAndYear(" & Right(("0" & Month(Date())), Len("00")) & ", " & Year(Date()) & ", document.PayrollFrm.PayrollDay);" & vbNewLine
				Response.Write "SelectItemByText('" & iTemp & "', false, document.PayrollFrm.PayrollDay);" & vbNewLine
			Response.Write "//--></SCRIPT>" & vbNewLine
		End If
	End If

	DisplayPayrollForm = lErrorNumber
	Err.Clear
End Function

Function DisplayPayrollsStatusTable(oRequest, oADODBConnection, iColumn, aPayrollComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the payrolls from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, iColumn, aPayrollComponent
'Outputs: aPayrollComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPayrollsStatusTable"
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sBoldBegin
	Dim sBoldEnd
	Dim lErrorNumber

	aPayrollComponent(S_QUERY_CONDITION_PAYROLL) = "And (IsActive_" & iColumn & "<>0) And (IsClosed<>1) And (PayrollTypeID<>0)"
	lErrorNumber = GetPayrolls(oRequest, oADODBConnection, aPayrollComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("Nombre,Fecha,Registro Habilitado", ",", -1, vbBinaryCompare)
				asCellWidths = Split("220,180,80", ",", -1, vbBinaryCompare)

				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				asCellAlignments = Split(",,CENTER", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("PayrollID").Value), oRequest("PayrollID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PayrollName").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PayrollDate").Value), -1, -1, -1) & sBoldEnd
					If CLng(oRecordset.Fields("IsActive_" & iColumn).Value) = 1 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & "Sí"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & "No"
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					oRecordset.MoveNext
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen nóminas registradas en la base de datos."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayPayrollsStatusTable = lErrorNumber
	Err.Clear
End Function

Function DisplayPayrollsTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, aPayrollComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the payrolls from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, aPayrollComponent
'Outputs: aPayrollComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPayrollsTable"
	Dim oRecordset
	Dim iCounter
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim sBoldBegin
	Dim sBoldEnd
	Dim lErrorNumber

	lErrorNumber = GetPayrolls(oRequest, oADODBConnection, aPayrollComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER=""0"" CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bUseLinks And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
					asColumnsTitles = Split("&nbsp;,Nombre,Fecha,Cerrada,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,220,180,80,80", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("&nbsp;,Nombre,Fecha,Cerrada", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,220,180,80", ",", -1, vbBinaryCompare)
				End If
				If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
					lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				Else
					lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
				End If

				iCounter = 0
				asCellAlignments = Split(",,,CENTER,CENTER", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("PayrollID").Value), oRequest("PayrollID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					ElseIf CLng(oRecordset.Fields("IsClosed").Value) <> 1 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sRowContents = ""
					If CLng(oRecordset.Fields("IsClosed").Value) = 0 Then
						Select Case lIDColumn
							Case DISPLAY_RADIO_BUTTONS
								sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""PayrollID"" ID=""PayrollIDRd"" VALUE=""" & CStr(oRecordset.Fields("PayrollID").Value) & """ />"
							Case DISPLAY_CHECKBOXES
								sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""PayrollID"" ID=""PayrollIDChk"" VALUE=""" & CStr(oRecordset.Fields("PayrollID").Value) & """ />"
							Case Else
								sRowContents = sRowContents & "&nbsp;"
						End Select
					Else
						sRowContents = sRowContents & "&nbsp;"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("PayrollName").Value)) & sBoldEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PayrollDate").Value), -1, -1, -1) & sBoldEnd
					If CLng(oRecordset.Fields("IsClosed").Value) = 1 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & "Sí"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sBoldBegin & "No"
					End If
					If bUseLinks And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
						sRowContents = sRowContents & TABLE_SEPARATOR
							If CLng(oRecordset.Fields("IsClosed").Value) <> 1 Then
								If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
									sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=UpdatePayroll&PayrollID=" & CStr(oRecordset.Fields("PayrollID").Value) & "&Change=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
								End If

								If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS Then
									sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=RemovePayroll&PayrollID=" & CStr(oRecordset.Fields("PayrollID").Value) & "&Delete=1"">"
										sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
									sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
								End If
							End If
						sRowContents = sRowContents & "&nbsp;"
					End If

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					oRecordset.MoveNext
					iCounter = iCounter + 1
					If (Err.number <> 0) Or (iCounter >= 24) Then Exit Do
				Loop
			Response.Write "</TABLE>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen nóminas registradas en la base de datos."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayPayrollsTable = lErrorNumber
	Err.Clear
End Function

Function DisplayModifyPayrollMessage(iMessage, lPayrollID)
'************************************************************
'Purpose: To display the message to generate the pre-payroll
'		  or to close the payroll
'Inputs:  iMessage, lPayrollID
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayModifyPayrollMessage"
	Dim sFlags

	Response.Write "<FORM NAME=""ReportFrm"" ID=""ReportFrm"" ACTION=""" & GetASPFileName("") & """ METHOD=""POST"">"
	Response.Write "<IMG SRC=""Images/IcnInformationSmall.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
		Select Case iMessage
			Case 0
				Response.Write "<B>Deseo generar la prenómina</B><BR /><BR />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""ModifyPayroll"" />"
				Response.Write "<B>Si no desea generar la prenómina para todos los empleado</B>, especifique los criterios para acotar el universo:<BR />"

				'sFlags = L_NO_INSTRUCTIONS_FLAGS & "," & L_OPEN_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS
				sFlags = L_NO_INSTRUCTIONS_FLAGS & "," & L_OPEN_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS
				Call DisplayReportFilter(sFlags, sErrorDescription)

				Response.Write "<BR />&nbsp;&nbsp;&nbsp;<INPUT TYPE=""SUBMIT"" NAME=""CalculatePayroll"" ID=""CalculatePayrollBtn"" VALUE=""Generar Prenómina"" CLASS=""Buttons"" />"
			Case 1
				Response.Write "<B>Deseo cerrar una nómina</B><BR /><BR />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""ClosePayroll"" />"
				sFlags = L_NO_INSTRUCTIONS_FLAGS & "," & L_NO_DIV_FLAGS & "," & L_OPEN_PAYROLL_FLAGS
				Call DisplayReportFilter(sFlags, sErrorDescription)
				Response.Write "<BR />&nbsp;&nbsp;&nbsp;<INPUT TYPE=""SUBMIT"" NAME=""DoClose"" ID=""DoCloseBtn"" VALUE=""Cerrar Nómina"" CLASS=""Buttons"" />"
		End Select
	Response.Write "</FORM>"
	If (iMessage = 1) And (InStr(1, ",0,2,243,", "," & aLoginComponent(N_USER_ID_LOGIN) & ",", vbBinaryCompare) > 0) Then
		Response.Write "<BR /><IMG SRC=""Images/Transparent.gif"" WIDTH=""1"" HEIGHT=""100"" /><BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""400"" HEIGHT=""1"" /><BR /><BR />"
		Response.Write "<FORM NAME=""ProcessPayrollFrm"" ID=""ProcessPayrollFrm"" ACTION=""Payroll.asp"" METHOD=""POST"">"
		Response.Write "<IMG SRC=""Images/IcnInformationSmall.gif"" WIDTH=""16"" HEIGHT=""16"" ALIGN=""ABSMIDDLE"" HSPACE=""3"" />"
			Response.Write "<B>Deseo procesar una nómina migrada</B><BR /><BR />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""ClosePayroll"" />"
			sFlags = L_NO_INSTRUCTIONS_FLAGS & "," & L_NO_DIV_FLAGS & "," & L_PAYROLL_FLAGS
			Call DisplayReportFilter(sFlags, sErrorDescription)
			Response.Write "<INPUT TYPE=""RADIO"" NAME=""DoMessages"" ID=""DoMessagesRd"" VALUE=""1"" /> Recontruir EmployeesChangesLKP<BR />"
			Response.Write "<INPUT TYPE=""RADIO"" NAME=""DoMessages"" ID=""DoMessagesRd"" VALUE=""2"" /> Generar mensajes<BR />"
			Response.Write "<BR />&nbsp;&nbsp;&nbsp;<INPUT TYPE=""SUBMIT"" NAME=""DoClose"" ID=""DoCloseBtn"" VALUE=""Procesar Nómina"" CLASS=""Buttons"" />"
		Response.Write "</FORM>"
	End If

	DisplayModifyPayrollMessage = Err.number
	Err.Clear
End Function

Function InsertPaymentMessages(oADODBConnection, aPayrollComponent, sErrorDescription)
'************************************************************
'Purpose: To insert the payments messages for the given payroll into the database
'Inputs:  oRequest, oADODBConnection, iLevel
'Outputs: aPayrollComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InsertPaymentMessages"
	Const ROWS_PER_FILE = 10000
	Dim asEmployeesQueries
	Dim iCounter
	Dim iCounter2
	Dim iIndex
	Dim jIndex
	Dim sFilePath
	Dim asFileContents
	Dim sDate
	Dim sQueryBegin
	Dim sQueryEnd
	Dim sCondition
	Dim lCurrentID
	Dim lCurrentID2
	Dim lConceptID
	Dim sCurrentID
	Dim adTotal
	Dim dAmount
	Dim dTaxAmount
	Dim dTemp
	Dim sTemp
	Dim oRecordset
	Dim lErrorNumber

	sFilePath = Server.MapPath("Export\Payroll_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & aPayrollComponent(N_ID_PAYROLL))
	sErrorDescription = "No se pudieron preparar las tablas para los pagos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PaymentsMessages Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (bSpecial<>2)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		iCounter = 0
'CAMBIO DE ESTATUS
		sTemp = ""
		sErrorDescription = "No se pudieron agregar los mensajes para los pagos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, EmployeesHistoryList.EmployeeDate, EmployeesHistoryList.EndDate, StatusEmployees.StatusName From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And ((EmployeesHistoryList.Active=0) Or (StatusEmployees.Active=0) Or (Reasons.ActiveEmployeeID=2)) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesHistoryList.EmployeeID) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=0) Order By EmployeesHistoryList.EmployeeID", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				sTemp = ""
				sTemp = CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value)) & " del " & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EmployeeDate").Value), -1, -1, -1)
				If CLng(oRecordset.Fields("EndDate").Value) <> 30000000 Then sTemp = sTemp & " al " & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1, -1)
				lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", "-1" & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EmployeeID").Value) & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
				iCounter = iCounter + 1
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
		End If

'CONCEPTOS 37, 38 Y 39. ESTÍMULOS
		sTemp = ""
		dTemp = Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))
		Select Case Mid(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("00000"), Len("00"))
			Case "01"
				dTemp = (CInt(Left(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("0000"))) - 1) & "12"
			Case Else
				dTemp = dTemp & "0" & (CInt(Mid(aPayrollComponent(N_FOR_DATE_PAYROLL), Len("00000"), Len("00"))) - 1)
		End Select
		dTemp = Right(dTemp, Len("YYMM"))
		lCurrentID = -2
		sDate = Right(dTemp, Len("MM")) & "/" & Left(dTemp, Len("YY"))
		sErrorDescription = "No se pudieron agregar los mensajes para los pagos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, Concepts.ConceptID, ConceptShortName, ConceptName From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", Concepts Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Concepts.ConceptID) And (Concepts.ConceptID In (40, 41, 42)) Order By EmployeeID, ConceptShortName", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			sTemp = ""
			Do While Not oRecordset.EOF
				lConceptID = CLng(oRecordset.Fields("ConceptID").Value)
				If lCurrentID <> CStr(oRecordset.Fields("EmployeeID").Value) Then
					If lCurrentID <> -2 Then
						sTemp = sTemp & sDate
						lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", lConceptID & SECOND_LIST_SEPARATOR & lCurrentID & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
						iCounter = iCounter + 1
					End If
					sTemp = "Estímulos "
					lCurrentID = CStr(oRecordset.Fields("EmployeeID").Value)
				End If
				sTemp = sTemp & "<" & (CLng(oRecordset.Fields("ConceptID").Value) - 3) & "> "
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			sTemp = sTemp & sDate
			lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", lConceptID & SECOND_LIST_SEPARATOR & lCurrentID & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
			iCounter = iCounter + 1
		End If

'CONCEPTO 49. PREMIO TRABAJADOR DEL MES (ESTÍMULO)
		lCurrentID = -2
		sErrorDescription = "No se pudieron agregar los mensajes para los pagos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, Concepts.ConceptID, ConceptShortName, ConceptName From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", Concepts Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Concepts.ConceptID) And (Concepts.ConceptID In (50)) Order By EmployeeID, ConceptShortName", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			sTemp = ""
			Do While Not oRecordset.EOF
				lConceptID = CLng(oRecordset.Fields("ConceptID").Value)
				If lCurrentID <> CStr(oRecordset.Fields("EmployeeID").Value) Then
					If lCurrentID <> -2 Then
						sTemp = sTemp & sDate
						lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", lConceptID & SECOND_LIST_SEPARATOR & lCurrentID & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
						iCounter = iCounter + 1
					End If
					sTemp = "Estímulo "
					lCurrentID = CStr(oRecordset.Fields("EmployeeID").Value)
				End If
				sTemp = sTemp & CStr(oRecordset.Fields("ConceptShortName").Value) & " "
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			sTemp = sTemp & sDate
			lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", lConceptID & SECOND_LIST_SEPARATOR & lCurrentID & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
			iCounter = iCounter + 1
		End If

'CONCEPTO 40. ESTÍMULO MÉRITO RELEVANTE
		sTemp = ""
		sErrorDescription = "No se pudieron agregar los mensajes para los pagos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, RecordDate, ConceptTaxes, ConceptShortName, ConceptName From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", Concepts Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Concepts.ConceptID) And (ConceptTaxes>0) And (Concepts.ConceptID=43) Order By EmployeeID, RecordDate", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				sTemp = " <40> "
					If StrComp(Right(("000" & CLng(oRecordset.Fields("ConceptTaxes").Value)), Len("1")), "1", vbBinaryCompare) = 0 Then
						sTemp = sTemp & (CInt(Right(dTemp, Len("MM"))) - 2) & "/" & Left(dTemp, Len("YY")) & ", "
					End If
					If StrComp(Left(Right(("000" & CLng(oRecordset.Fields("ConceptTaxes").Value)), Len("11")), Len("1")), "1", vbBinaryCompare) = 0 Then
						sTemp = sTemp & (CInt(Right(dTemp, Len("MM"))) - 1) & "/" & Left(dTemp, Len("YY")) & ", "
					End If
					If CLng(oRecordset.Fields("ConceptTaxes").Value) >= 100 Then
						sTemp = sTemp & Right(dTemp, Len("MM")) & "/" & Left(dTemp, Len("YY")) & ", "
					End If
					sTemp = Left(sTemp, (Len(sTemp) - Len(", ")))
				lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", "43" & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EmployeeID").Value) & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
				iCounter = iCounter + 1
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
		End If

'CONCEPTO 9. HORAS EXTRAS
		sTemp = ""
		sErrorDescription = "No se pudieron agregar los mensajes para los pagos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesAbsencesLKP.EmployeeID, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate, AbsenceHours, ConceptShortName, ConceptName From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesAbsencesLKP, Concepts Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesAbsencesLKP.EmployeeID) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Concepts.ConceptID) And (EmployeesAbsencesLKP.AbsenceID In (201,909)) And (AppliedDate In (0," & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (EmployeesAbsencesLKP.Active=1) And (Concepts.ConceptID In (9)) Order By EmployeesAbsencesLKP.EmployeeID, EmployeesAbsencesLKP.OcurredDate, ConceptShortName", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		lCurrentID = -2
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
					If lCurrentID > -2 Then
						lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", "9" & SECOND_LIST_SEPARATOR & lCurrentID & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
						iCounter = iCounter + 1
					End If
					sTemp = "TIEMPO EXTRA: "
					lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
				End If
				sTemp = sTemp & "DÍA " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value)) & " CON " & CStr(oRecordset.Fields("AbsenceHours").Value) & " HRS. "
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			If Len(sTemp) > 0 Then
				lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", "9" & SECOND_LIST_SEPARATOR & lCurrentID & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
				iCounter = iCounter + 1
			End If
		End If
If False Then
'CONCEPTO 16. DOMINGOS
		sTemp = ""
		sErrorDescription = "No se pudieron agregar los mensajes para los pagos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesAbsencesLKP.EmployeeID, EmployeesAbsencesLKP.OcurredDate, EmployeesAbsencesLKP.EndDate, AbsenceHours, ConceptShortName, ConceptName From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesAbsencesLKP, Concepts Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesAbsencesLKP.EmployeeID) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Concepts.ConceptID) And (EmployeesAbsencesLKP.AbsenceID In (202,914)) And (AppliedDate In (0," & aPayrollComponent(N_FOR_DATE_PAYROLL) & ")) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (EmployeesAbsencesLKP.Active=1) And (Concepts.ConceptID In (16)) Order By EmployeesAbsencesLKP.EmployeeID, EmployeesAbsencesLKP.OcurredDate, ConceptShortName", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		lCurrentID = -2
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
					If lCurrentID > -2 Then
						lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", "16" & SECOND_LIST_SEPARATOR & lCurrentID & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
						iCounter = iCounter + 1
					End If
					sTemp = "DOMINGOS LABORADOS: "
					lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
				End If
				sTemp = sTemp & "DÍA " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value)) & ". "
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			If Len(sTemp) > 0 Then
				lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", "16" & SECOND_LIST_SEPARATOR & lCurrentID & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
				iCounter = iCounter + 1
			End If
		End If
End If
'CONCEPTO 50. INASISTENCIAS
		sTemp = ""
		sErrorDescription = "No se pudieron agregar los mensajes para los pagos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, Shifts.JourneyTypeID, Absences.AbsenceID, AbsenceShortName, OcurredDate From EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, EmployeesAbsencesLKP, Absences, Journeys, Shifts, JourneyTypes Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.EmployeeID=EmployeesAbsencesLKP.EmployeeID) And (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (Shifts.JourneyTypeID=JourneyTypes.JourneyTypeID) And (Journeys.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Journeys.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Shifts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Shifts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesAbsencesLKP.AbsenceID In (3,10,11,16,18,19,20,24,25,26,28,92,93,94)) And (AppliedDate In (0," & aPayrollComponent(N_ID_PAYROLL) & ")) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (JourneyTypes.JourneyFactor>0) And (EmployeesAbsencesLKP.Active=1) Order By EmployeesHistoryList.EmployeeID, AbsenceShortName, OcurredDate", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		lCurrentID = -2
		sCurrentID = ""
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
					If lCurrentID > -2 Then
						lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", "52" & SECOND_LIST_SEPARATOR & lCurrentID & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
						iCounter = iCounter + 1
					End If
					lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
					sCurrentID = GetPayrollNumber(CLng(oRecordset.Fields("OcurredDate").Value)) & "-" & Left(CStr(oRecordset.Fields("OcurredDate").Value), Len("0000")) & ". J" & CStr(oRecordset.Fields("JourneyTypeID").Value) & " C" & Right(CStr(oRecordset.Fields("AbsenceShortName").Value), Len("00"))
					sTemp = "QNA " & sCurrentID
				End If
				If StrComp(sCurrentID, (GetPayrollNumber(CLng(oRecordset.Fields("OcurredDate").Value)) & "-" & Left(CStr(oRecordset.Fields("OcurredDate").Value), Len("0000")) & ". J" & CStr(oRecordset.Fields("JourneyTypeID").Value) & " C" & Right(CStr(oRecordset.Fields("AbsenceShortName").Value), Len("00"))), vbBinaryCompare) <> 0 Then
					sCurrentID = GetPayrollNumber(CLng(oRecordset.Fields("OcurredDate").Value)) & "-" & Left(CStr(oRecordset.Fields("OcurredDate").Value), Len("0000")) & ". J" & CStr(oRecordset.Fields("JourneyTypeID").Value) & " C" & Right(CStr(oRecordset.Fields("AbsenceShortName").Value), Len("00"))
					sTemp = sTemp & ". QNA " & sCurrentID
				End If
				sTemp = sTemp & ". Día " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value))
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			If Len(sTemp) > 0 Then
				lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", "52" & SECOND_LIST_SEPARATOR & lCurrentID & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
				iCounter = iCounter + 1
			End If
		End If

'CONCEPTO 70. RETARDOS
		sTemp = ""
		sErrorDescription = "No se pudieron agregar los mensajes para los pagos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryList.EmployeeID, Shifts.JourneyTypeID, Absences.AbsenceID, AbsenceShortName, OcurredDate From EmployeesChangesLKP, EmployeesHistoryList, StatusEmployees, Reasons, EmployeesAbsencesLKP, Absences, Journeys, Shifts, JourneyTypes Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.StatusID=StatusEmployees.StatusID) And (EmployeesHistoryList.ReasonID=Reasons.ReasonID) And (EmployeesHistoryList.Active=1) And (StatusEmployees.Active=1) And (Reasons.ActiveEmployeeID<>2) And (EmployeesHistoryList.EmployeeID=EmployeesAbsencesLKP.EmployeeID) And (EmployeesAbsencesLKP.AbsenceID=Absences.AbsenceID) And (EmployeesHistoryList.JourneyID=Journeys.JourneyID) And (EmployeesHistoryList.ShiftID=Shifts.ShiftID) And (Shifts.JourneyTypeID=JourneyTypes.JourneyTypeID) And (Journeys.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Journeys.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Shifts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Shifts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesAbsencesLKP.AbsenceID In (1,2,4,5,21,23,27)) And (AppliedDate In (0," & aPayrollComponent(N_ID_PAYROLL) & ")) And (EmployeesAbsencesLKP.JustificationID=-1) And (EmployeesAbsencesLKP.Removed=0) And (JourneyTypes.JourneyFactor>0) And (EmployeesAbsencesLKP.Active=1) Order By EmployeesHistoryList.EmployeeID, AbsenceShortName, OcurredDate", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		lCurrentID = -2
		sCurrentID = ""
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				If lCurrentID <> CLng(oRecordset.Fields("EmployeeID").Value) Then
					If lCurrentID > -2 Then
						lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", "71" & SECOND_LIST_SEPARATOR & lCurrentID & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
						iCounter = iCounter + 1
					End If
					lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
					sCurrentID = GetPayrollNumber(CLng(oRecordset.Fields("OcurredDate").Value)) & "-" & Left(CStr(oRecordset.Fields("OcurredDate").Value), Len("0000")) & ". J" & CStr(oRecordset.Fields("JourneyTypeID").Value) & " C" & Right(CStr(oRecordset.Fields("AbsenceShortName").Value), Len("00"))
					sTemp = "QNA " & sCurrentID
				End If
				If StrComp(sCurrentID, (CStr(oRecordset.Fields("OcurredDate").Value) & ";" & CStr(oRecordset.Fields("AbsenceID").Value)), vbBinaryCompare) <> 0 Then
					sCurrentID = GetPayrollNumber(CLng(oRecordset.Fields("OcurredDate").Value)) & "-" & Left(CStr(oRecordset.Fields("OcurredDate").Value), Len("0000")) & ". J" & CStr(oRecordset.Fields("JourneyTypeID").Value) & " C" & Right(CStr(oRecordset.Fields("AbsenceShortName").Value), Len("00"))
					sTemp = sTemp & ". QNA " & sCurrentID
				End If
				sTemp = sTemp & ". Día " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("OcurredDate").Value))
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			If Len(sTemp) > 0 Then
				lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", "71" & SECOND_LIST_SEPARATOR & lCurrentID & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
				iCounter = iCounter + 1
			End If
		End If

'DESCUENTOS
		sTemp = ""
		iCounter2 = 0
		sErrorDescription = "No se pudieron agregar los mensajes para los pagos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Credits.EmployeeID, Credits.StartDate, Credits.EndDate, Credits.PaymentsCounter, Concepts.ConceptID, ConceptShortName, ConceptName From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", Credits, Concepts Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=Credits.EmployeeID) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=Credits.CreditTypeID) And (Credits.CreditTypeID=Concepts.ConceptID) And (Credits.CreditTypeID>0) And ((Credits.PaymentsCounter<Credits.PaymentsNumber) Or (Credits.PaymentsNumber<1)) And (Credits.Active=1) And (Credits.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Credits.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") Order By Credits.EmployeeID, ConceptShortName", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				lConceptID = CLng(oRecordset.Fields("ConceptID").Value)
				If (CLng(oRecordset.Fields("EndDate").Value) > 0) And (CLng(oRecordset.Fields("EndDate").Value) < 30000000) Then
					adTotal = Split("0,0,0,0,0,0", ",")
					Call GetAntiquityFromSerialDates(CLng(oRecordset.Fields("StartDate").Value), aPayrollComponent(N_ID_PAYROLL), adTotal(0), adTotal(1), adTotal(2))
					lCurrentID = (adTotal(0) * 24) + (adTotal(1) * 2)
					If adTotal(2) >= 15 Then lCurrentID = lCurrentID + 1
					Call GetAntiquityFromSerialDates(CLng(oRecordset.Fields("StartDate").Value), CLng(oRecordset.Fields("EndDate").Value), adTotal(3), adTotal(4), adTotal(5))
					lCurrentID2 = (adTotal(3) * 24) + (adTotal(4) * 2)
					If adTotal(5) >= 15 Then lCurrentID2 = lCurrentID2 + 1
					sTemp = ""
					sTemp = "Concepto " & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & ". Del " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)) & " al " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value)) & ". " & lCurrentID2 & " quincenas."
					lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", lConceptID & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EmployeeID").Value) & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
					lErrorNumber = AppendTextToFile(sFilePath & "_Messages2_" & Int(iCounter2 / ROWS_PER_FILE) & ".txt", CStr(oRecordset.Fields("EmployeeID").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("ConceptID").Value) & SECOND_LIST_SEPARATOR & lCurrentID & SECOND_LIST_SEPARATOR & lCurrentID2, sErrorDescription)
					iCounter = iCounter + 1
					iCounter2 = iCounter2 + 1
				Else
					sTemp = ""
					sTemp = "Concepto " & CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value)) & ". Del " & DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value))
					lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", lConceptID & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EmployeeID").Value) & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
					iCounter = iCounter + 1
				End If
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
		End If

'CONCEPTO 71. DEDUCCIÓN POR COBRO DE SUELDOS INDEBIDOS
		sTemp = ""
		sErrorDescription = "No se pudieron agregar los mensajes para los pagos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesConceptsLKP.EmployeeID, EmployeesConceptsLKP.StartDate, EmployeesConceptsLKP.EndDate, ConceptShortName, ConceptName From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesConceptsLKP, Concepts Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesConceptsLKP.EmployeeID) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=EmployeesConceptsLKP.ConceptID) And (EmployeesConceptsLKP.ConceptID=Concepts.ConceptID) And (EmployeesConceptsLKP.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesConceptsLKP.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.ConceptID In (72)) Order By EmployeesConceptsLKP.EmployeeID, ConceptShortName", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				adTotal = Split("0,0,0,0,0,0", ",")
				Call GetAntiquityFromSerialDates(CLng(oRecordset.Fields("StartDate").Value), aPayrollComponent(N_ID_PAYROLL), adTotal(0), adTotal(1), adTotal(2))
				lCurrentID = (adTotal(0) * 24) + (adTotal(1) * 2)
				If adTotal(2) >= 15 Then lCurrentID = lCurrentID + 1
				Call GetAntiquityFromSerialDates(CLng(oRecordset.Fields("StartDate").Value), CLng(oRecordset.Fields("EndDate").Value), adTotal(3), adTotal(4), adTotal(5))
				lCurrentID2 = (adTotal(3) * 24) + (adTotal(4) * 2)
				If adTotal(5) >= 15 Then lCurrentID2 = lCurrentID2 + 1
				sTemp = ""
				sTemp = CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value) & ". " & CStr(oRecordset.Fields("ConceptName").Value)) & " descuento " & lCurrentID & " de " & lCurrentID2 & "."
				lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", "72" & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EmployeeID").Value) & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
				iCounter = iCounter + 1
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
		End If

'CONCEPTO D2. RETENCIONES POR EXCESO DE LICENCIAS MÉDICAS
		sTemp = ""
		sErrorDescription = "No se pudieron agregar los mensajes para los pagos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesConceptsLKP.EmployeeID, EmployeesConceptsLKP.StartDate, EmployeesConceptsLKP.EndDate, ConceptShortName, ConceptName From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ", EmployeesConceptsLKP, Concepts Where (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".EmployeeID=EmployeesConceptsLKP.EmployeeID) And (Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ".ConceptID=EmployeesConceptsLKP.ConceptID) And (EmployeesConceptsLKP.ConceptID=Concepts.ConceptID) And (EmployeesConceptsLKP.StartDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesConceptsLKP.EndDate>=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (Concepts.ConceptID In (104)) Order By EmployeesConceptsLKP.EmployeeID, ConceptShortName", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				sTemp = ""
				sTemp = CleanStringForHTML(CStr(oRecordset.Fields("ConceptShortName").Value) & ". " & CStr(oRecordset.Fields("ConceptName").Value)) & ". El descuento aplica del " & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value), -1, -1 ,-1) & " al " & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("EndDate").Value), -1, -1 ,-1) & "."
				lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", "104" & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EmployeeID").Value) & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
				iCounter = iCounter + 1
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
		End If

'CONCEPTO AN IS SI SS
		sTemp = ""
		sErrorDescription = "No se pudieron agregar los mensajes para los pagos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID In (88,110,120,122)) Order By EmployeeID", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				sTemp = """SI"" SEGURO DE SEPARACIÓN INDIVIDUALIZADO APORTACIÓN PERSONAL DE MANDO, ""SS"" APORTACIÓN PATRONAL BRUTA<BR />""IS"" IMPUESTO SOBRE LA RENTA DEL ""SI"" SUBSIDIADO POR EL INSTITUTO,    ""AN"" APORTACIÓN NETA PATRONAL DEL ""SI"""
				lErrorNumber = AppendTextToFile(sFilePath & "_Messages_" & Int(iCounter / ROWS_PER_FILE) & ".txt", "88" & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EmployeeID").Value) & SECOND_LIST_SEPARATOR & sTemp, sErrorDescription)
				iCounter = iCounter + 1
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
		End If

		If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 3, RUN FROM FILES, ADD PaymentMessages")
			sQueryBegin = "Insert Into PaymentsMessages (RecordID, PayrollID, EmployeeID, CompanyID, AreaIDs, ZoneIDs, EmployeeTypeID, PositionID, BankID, ConceptID, bSpecial, Comments) Values ("
			sQueryEnd = ", -1, '-1', '-1', -1, -1, -1, "
			For jIndex = 0 To iCounter Step ROWS_PER_FILE
				asFileContents = GetFileContents(sFilePath & "_Messages_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
				If Len(asFileContents) > 0 Then
					asFileContents = Split(asFileContents, vbNewLine)
					lErrorNumber = GetNewIDFromTable(oADODBConnection, "PaymentsMessages", "RecordID", "", 1, lCurrentID, sErrorDescription)
					If lErrorNumber = 0 Then
						For iIndex = 0 To UBound(asFileContents)
							If Len(asFileContents(iIndex)) > 0 Then
								asEmployeesQueries = Split(asFileContents(iIndex), SECOND_LIST_SEPARATOR, 3, vbBinaryCompare)
								sErrorDescription = "No se pudo agregar el mensaje para la nómina del empleado."
								If CLng(asEmployeesQueries(0)) = 50 Then
									lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Update PaymentsMessages Set Comments=Comments+'  " & Replace(asEmployeesQueries(2), "Estímulo", "") & "' Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeeID=" & asEmployeesQueries(1) & ") And (ConceptID In (40,41,42))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
									lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, sQueryBegin & lCurrentID & ", " & aPayrollComponent(N_ID_PAYROLL) & ", " & asEmployeesQueries(1) & sQueryEnd & asEmployeesQueries(0) & ", 1, 'Estímulo" & asEmployeesQueries(2) & "')", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
								ElseIf CLng(asEmployeesQueries(0)) = 43 Then
									lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Update PaymentsMessages Set Comments=Comments+'  " & asEmployeesQueries(2) & "' Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeeID=" & asEmployeesQueries(1) & ") And (ConceptID In (40,41,42))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
									lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, sQueryBegin & lCurrentID & ", " & aPayrollComponent(N_ID_PAYROLL) & ", " & asEmployeesQueries(1) & sQueryEnd & asEmployeesQueries(0) & ", 1, 'Estímulo" & asEmployeesQueries(2) & "')", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
								Else
									lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, sQueryBegin & lCurrentID & ", " & aPayrollComponent(N_ID_PAYROLL) & ", " & asEmployeesQueries(1) & sQueryEnd & asEmployeesQueries(0) & ", 1, '" & asEmployeesQueries(2) & "')", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
								End If
								lCurrentID = lCurrentID + 1
							End If
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
						Next
					End If
				End If
				If iCounter > 0 Then Call DeleteFile(sFilePath & "_Messages_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
			Next
		End If

		sErrorDescription = "No se pudo agregar el mensaje para la nómina del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PaymentsMessages Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (ConceptID=43) And (EmployeeID In (Select EmployeeID From PaymentsMessages Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (ConceptID In (40,41,42))))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

		sErrorDescription = "No se pudo agregar el mensaje para la nómina del empleado."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From PaymentsMessages Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (ConceptID=50) And (EmployeeID In (Select EmployeeID From PaymentsMessages Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (ConceptID In (40,41,42))))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

		sErrorDescription = "No se pudieron agregar los mensajes para las nóminas de los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PaymentsMessages (RecordID, PayrollID, EmployeeID, CompanyID, AreaIDs, ZoneIDs, EmployeeTypeID, PositionID, BankID, ConceptID, bSpecial, Comments) Select Distinct EmployeeID+" & lCurrentID & " As RecordID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollID, EmployeeID, -1 As CompanyID, '-1' As AreaIDs, '-1' As ZoneIDs, -1 As EmployeeTypeID, -1 As PositionID, -1 As BankID, 0 As ConceptID, 3 As bSpecial, ' ' As Comments From EmployeesHistoryListForPayroll Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

		lErrorNumber = GetNewIDFromTable(oADODBConnection, "PaymentsMessages", "RecordID", "", 1, lCurrentID, sErrorDescription)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron agregar los mensajes para las nóminas de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PaymentsMessages (RecordID, PayrollID, EmployeeID, CompanyID, AreaIDs, ZoneIDs, EmployeeTypeID, PositionID, BankID, ConceptID, bSpecial, Comments) Select Distinct EmployeeID+" & lCurrentID & " As RecordID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollID, EmployeeID, -1 As CompanyID, '-1' As AreaIDs, '-1' As ZoneIDs, -1 As EmployeeTypeID, -1 As PositionID, -1 As BankID, 40 As ConceptID, 3 As bSpecial, ' ' As Comments From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (ConceptID=40) And (EmployeeID Not In (Select EmployeeID From PaymentsMessages Where (ConceptID=40)))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "PaymentsMessages", "RecordID", "", 1, lCurrentID, sErrorDescription)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron agregar los mensajes para las nóminas de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PaymentsMessages (RecordID, PayrollID, EmployeeID, CompanyID, AreaIDs, ZoneIDs, EmployeeTypeID, PositionID, BankID, ConceptID, bSpecial, Comments) Select Distinct EmployeeID+" & lCurrentID & " As RecordID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollID, EmployeeID, -1 As CompanyID, '-1' As AreaIDs, '-1' As ZoneIDs, -1 As EmployeeTypeID, -1 As PositionID, -1 As BankID, 41 As ConceptID, 3 As bSpecial, ' ' As Comments From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (ConceptID=41) And (EmployeeID Not In (Select EmployeeID From PaymentsMessages Where (ConceptID=41)))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "PaymentsMessages", "RecordID", "", 1, lCurrentID, sErrorDescription)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron agregar los mensajes para las nóminas de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PaymentsMessages (RecordID, PayrollID, EmployeeID, CompanyID, AreaIDs, ZoneIDs, EmployeeTypeID, PositionID, BankID, ConceptID, bSpecial, Comments) Select Distinct EmployeeID+" & lCurrentID & " As RecordID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollID, EmployeeID, -1 As CompanyID, '-1' As AreaIDs, '-1' As ZoneIDs, -1 As EmployeeTypeID, -1 As PositionID, -1 As BankID, 42 As ConceptID, 3 As bSpecial, ' ' As Comments From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (ConceptID=42) And (EmployeeID Not In (Select EmployeeID From PaymentsMessages Where (ConceptID=42)))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "PaymentsMessages", "RecordID", "", 1, lCurrentID, sErrorDescription)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron agregar los mensajes para las nóminas de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PaymentsMessages (RecordID, PayrollID, EmployeeID, CompanyID, AreaIDs, ZoneIDs, EmployeeTypeID, PositionID, BankID, ConceptID, bSpecial, Comments) Select Distinct EmployeeID+" & lCurrentID & " As RecordID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollID, EmployeeID, -1 As CompanyID, '-1' As AreaIDs, '-1' As ZoneIDs, -1 As EmployeeTypeID, -1 As PositionID, -1 As BankID, 43 As ConceptID, 3 As bSpecial, ' ' As Comments From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (ConceptID=43) And (EmployeeID Not In (Select EmployeeID From PaymentsMessages Where (ConceptID=43)))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "PaymentsMessages", "RecordID", "", 1, lCurrentID, sErrorDescription)
		If lErrorNumber = 0 Then
			sErrorDescription = "No se pudieron agregar los mensajes para las nóminas de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into PaymentsMessages (RecordID, PayrollID, EmployeeID, CompanyID, AreaIDs, ZoneIDs, EmployeeTypeID, PositionID, BankID, ConceptID, bSpecial, Comments) Select Distinct EmployeeID+" & lCurrentID & " As RecordID, " & aPayrollComponent(N_ID_PAYROLL) & " As PayrollID, EmployeeID, -1 As CompanyID, '-1' As AreaIDs, '-1' As ZoneIDs, -1 As EmployeeTypeID, -1 As PositionID, -1 As BankID, 50 As ConceptID, 3 As bSpecial, ' ' As Comments From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (RecordDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (ConceptID=50) And (EmployeeID Not In (Select EmployeeID From PaymentsMessages Where (ConceptID=50)))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If

		If (lErrorNumber = 0) And (iCounter2 > 0) Then
Call DisplayTimeStamp("START: LEVEL 3, RUN FROM FILES, ADD Payroll_YYYYMMDD.ConceptRetention")
			sQueryBegin = "Update Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Set ConceptRetention=<COUNTER_1 />.<COUNTER_2 />1 Where (EmployeeID=<EMPLOYEE_ID />) And (ConceptID=<CONCEPT_ID />)"
			For jIndex = 0 To iCounter2 Step ROWS_PER_FILE
				asFileContents = GetFileContents(sFilePath & "_Messages2_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
				If Len(asFileContents) > 0 Then
					asFileContents = Split(asFileContents, vbNewLine)
					For iIndex = 0 To UBound(asFileContents)
						If Len(asFileContents(iIndex)) > 0 Then
							asEmployeesQueries = Split(asFileContents(iIndex), SECOND_LIST_SEPARATOR, 4, vbBinaryCompare)
							sErrorDescription = "No se pudo agregar el mensaje para la nómina del empleado."
							lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, Replace(Replace(Replace(Replace(sQueryBegin, "<EMPLOYEE_ID />", asEmployeesQueries(0)), "<CONCEPT_ID />", asEmployeesQueries(1)), "<COUNTER_1 />", asEmployeesQueries(2)), "<COUNTER_2 />", asEmployeesQueries(3)), "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
						End If
						If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
					Next
				End If
				If iCounter2 > 0 Then Call DeleteFile(sFilePath & "_Messages2_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
			Next
		End If
	End If

	InsertPaymentMessages = Err.number
	Err.Clear
End Function

Function InsertEmployeesChangesLKP(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
'************************************************************
'Purpose: To calculate the payroll and save it into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPayrollComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InsertEmployeesChangesLKP"
	Const ROWS_PER_FILE = 10000
	Dim lPayrollDate
	Dim asEmployeesQueries
	Dim iCounter
	Dim iCounter2
	Dim iPayrollIndex
	Dim iIndex
	Dim jIndex
	Dim sFilePath
	Dim asFileContents
	Dim sDate
	Dim oStartDate
	Dim oEndDate
	Dim sQueryBegin
	Dim sQueryEnd
	Dim sCondition
	Dim dTemp
	Dim sTemp
	Dim lFirstPayroll
	Dim lLastPayroll
	Dim asPayrolls
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPayrollComponent(B_COMPONENT_INITIALIZED_PAYROLL)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollComponent(oRequest, aPayrollComponent)
	End If

	lErrorNumber = GetPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
	If lErrorNumber = 0 Then
		aPayrollComponent(N_FOR_DATE_PAYROLL) = CLng(aPayrollComponent(N_FOR_DATE_PAYROLL))
		lPayrollDate = aPayrollComponent(N_FOR_DATE_PAYROLL)
		oStartDate = Now()
		sFilePath = Server.MapPath("Export\Payroll_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & aPayrollComponent(N_ID_PAYROLL))

		sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (-" & aPayrollComponent(N_ID_PAYROLL) & "," & aPayrollComponent(N_ID_PAYROLL) & "))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write "<!-- Query: Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (-" & aPayrollComponent(N_ID_PAYROLL) & "," & aPayrollComponent(N_ID_PAYROLL) & ")) -->" & vbNewLine
		If lErrorNumber = 0 Then
Call DisplayTimeStamp("START: LEVEL 1, INSERT RECORDS, EmployeesChanges")
			sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
			lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollDate, Max(EmployeeDate) As EmployeeDate1, " & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & " As FirstDate, " & aPayrollComponent(N_FOR_DATE_PAYROLL) & " As LastDate, 0 As Concepts40 From EmployeesHistoryList Where (EmployeeDate<=EndDate) And (EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EndDate>=" & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & ") And (Active=1) Group By EmployeeID", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
			Response.Write "<!-- Query: Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollDate, Max(EmployeeDate) As EmployeeDate1, " & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & " As FirstDate, " & aPayrollComponent(N_FOR_DATE_PAYROLL) & " As LastDate, 0 As Concepts40 From EmployeesHistoryList Where (EmployeeDate<=EndDate) And (EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EndDate>=" & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & ") And (Active=1) Group By EmployeeID -->" & vbNewLine
		End If

		Call BuildCondition(sCondition, sQueryBegin)
		sErrorDescription = "No se pudieron eliminar los conceptos de pagos de la nómina."
		If (aPayrollComponent(N_TYPE_ID_PAYROLL) = 4) Then
			lFirstPayroll = CLng(oRequest("StartPayrollYear").Item & Right(("0" & oRequest("StartPayrollMonth").Item), Len("00")) & Right(("0" & oRequest("StartPayrollDay").Item), Len("00")))
			lLastPayroll = CLng(oRequest("EndPayrollYear").Item & Right(("0" & oRequest("EndPayrollMonth").Item), Len("00")) & Right(("0" & oRequest("EndPayrollDay").Item), Len("00")))
		End If

		If aPayrollComponent(N_TYPE_ID_PAYROLL) = 4 Then
			sErrorDescription = "No se pudieron obtener las nóminas para calcular los pagos retroactivos."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PayrollDate From Payrolls Where (PayrollDate>=" & lFirstPayroll & ") And (PayrollDate<=" & lLastPayroll & ") And (PayrollTypeID=1)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				asPayrolls = ""
				Do While Not oRecordset.EOF
					asPayrolls = asPayrolls & CStr(oRecordset.Fields("PayrollDate").Value) & ","
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				oRecordset.Close
			End If
		Else
			If ((iLevel = 2) Or (iLevel=-1)) And (lErrorNumber = 0) Then
Call DisplayTimeStamp("START: LEVEL 2, RETROACTIVE. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
				Call BuildCondition(sCondition, "")

				sErrorDescription = "No se pudieron obtener los registros de la base de datos."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct EmployeesRevisions.StartPayrollID From EmployeesRevisions, Payrolls Where (EmployeesRevisions.StartPayrollID=Payrolls.PayrollID) And (Payrolls.PayrollTypeID=1) And (EmployeesRevisions.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesRevisions.StartPayrollID<" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeesRevisions.EmployeeID In (Select Distinct EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList " & sQueryBegin & " Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) " & sCondition & ")) Order By EmployeesRevisions.StartPayrollID", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				asPayrolls = ""
				If lErrorNumber = 0 Then
					Do While Not oRecordset.EOF
						asPayrolls = asPayrolls & CStr(oRecordset.Fields("StartPayrollID").Value) & ","
						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
					oRecordset.Close
				End If
			End If
			asPayrolls = asPayrolls & aPayrollComponent(N_FOR_DATE_PAYROLL) & ","
		End If
		If Len(asPayrolls) > 0 Then asPayrolls = Left(asPayrolls, (Len(asPayrolls) - Len(",")))



		If Len(asPayrolls) > 0 Then
			asPayrolls = Split(asPayrolls, ",")
			For iPayrollIndex = 0 To UBound(asPayrolls) - 1
				aPayrollComponent(N_FOR_DATE_PAYROLL) = CLng(asPayrolls(iPayrollIndex))
Call LogErrorInXMLFile("123", "Vic: Inicia llamado a BuildEmployeesChangesLKP para " & aPayrollComponent(N_FOR_DATE_PAYROLL), 0, "_", "_", 0)
				lErrorNumber = BuildEmployeesChangesLKP(aPayrollComponent, True, False, sErrorDescription)
Call LogErrorInXMLFile("123", "Vic: Termina llamado a BuildEmployeesChangesLKP para " & aPayrollComponent(N_FOR_DATE_PAYROLL), 0, "_", "_", 0)
				If bTimeout Then Exit For
			Next
			aPayrollComponent(N_FOR_DATE_PAYROLL) = CLng(asPayrolls(iPayrollIndex))
		End If
		If Not bTimeout Then
Call LogErrorInXMLFile("123", "Vic: Inicia llamado a BuildEmployeesChangesLKP para " & aPayrollComponent(N_FOR_DATE_PAYROLL), 0, "_", "_", 0)
			lErrorNumber = BuildEmployeesChangesLKP(aPayrollComponent, (aPayrollComponent(N_TYPE_ID_PAYROLL) = 4), False, sErrorDescription)
Call LogErrorInXMLFile("123", "Vic: Termina llamado a BuildEmployeesChangesLKP para " & aPayrollComponent(N_FOR_DATE_PAYROLL), 0, "_", "_", 0)
		End If

		If Not bTimeout Then
			asPayrolls = ""
			sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Distinct MissingDate From EmployeesAdjustmentsLKP Where (EmployeeID Not In (Select Distinct EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ")) And (PayrollDate In (0," & aPayrollComponent(N_ID_PAYROLL) & ")) And (Active=1) Order By MissingDate", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					Do While Not oRecordset.EOF
						asPayrolls = asPayrolls & CStr(oRecordset.Fields("MissingDate").Value) & ","
						oRecordset.MoveNext
					Loop
					oRecordset.Close
					If Len(asPayrolls) > 0 Then asPayrolls = Left(asPayrolls, (Len(asPayrolls) - Len(",")))
					asPayrolls = Split(asPayrolls, ",")

					For iPayrollIndex = 0 To UBound(asPayrolls)
						aPayrollComponent(N_FOR_DATE_PAYROLL) = CLng(asPayrolls(iPayrollIndex))
Call LogErrorInXMLFile("123", "Vic: Inicia llamado a BuildEmployeesChangesLKP para " & aPayrollComponent(N_FOR_DATE_PAYROLL), 0, "_", "_", 0)
						lErrorNumber = BuildEmployeesChangesLKP(aPayrollComponent, True, True, sErrorDescription)
Call LogErrorInXMLFile("123", "Vic: Termina llamado a BuildEmployeesChangesLKP para " & aPayrollComponent(N_FOR_DATE_PAYROLL), 0, "_", "_", 0)
						If bTimeout Then Exit For
					Next
				End If
			End If
		End If

		aPayrollComponent(N_FOR_DATE_PAYROLL) = CLng(lPayrollDate)

		Call BuildCondition(sCondition, "")
		If Not bTimeout Then
			sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ")", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
		If Not bTimeout Then
			sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
			lErrorNumber = ExecuteInsertQuerySp(oADODBConnection, "Insert Into EmployeesChangesLKP (EmployeeID, PayrollID, PayrollDate, EmployeeDate, FirstDate, LastDate, Concepts40) Select EmployeeID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollID, '" & aPayrollComponent(N_ID_PAYROLL) & "' As PayrollDate, Max(EmployeeDate) As EmployeeDate1, " & GetPayrollStartDate(aPayrollComponent(N_FOR_DATE_PAYROLL)) & " As FirstDate, " & aPayrollComponent(N_FOR_DATE_PAYROLL) & " As LastDate, 0 As Concepts40 From EmployeesHistoryList Where (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeDate<=" & aPayrollComponent(N_FOR_DATE_PAYROLL) & ") And (EmployeeID In (Select Distinct EmployeeID From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & ")) Group By EmployeeID", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
		End If
		If Not bTimeout Then
			iCounter = 0
			sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, Concepts40, Min(FirstDate) As MinFirstDate, Max(LastDate) As MaxLastDate From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate<0) Group By EmployeeID, Concepts40 Order By EmployeeID", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				Do While Not oRecordset.EOF
					lErrorNumber = AppendTextToFile(sFilePath & "_EmployeesChangesLKP_" & Int(iCounter / ROWS_PER_FILE) & ".txt", CStr(oRecordset.Fields("EmployeeID").Value) & "," & CStr(oRecordset.Fields("MinFirstDate").Value) & "," & CStr(oRecordset.Fields("MaxLastDate").Value) & "," & CStr(oRecordset.Fields("Concepts40").Value), sErrorDescription)
					iCounter = iCounter + 1
					oRecordset.MoveNext
				Loop
				oRecordset.Close
			End If
			If (lErrorNumber = 0) And (iCounter > 0) Then
Call DisplayTimeStamp("START: LEVEL 2, RUN FROM FILES, Update EmployeesChangesLKP " & iCounter & " RECORDS.")
				sQueryBegin = "Update EmployeesChangesLKP Set FirstDate=<FIRST_DATE />, LastDate=<LAST_DATE />, Concepts40=<CONCEPTS_40 /> Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeeID=<EMPLOYEE_ID />)"
				For jIndex = 0 To iCounter Step ROWS_PER_FILE
					asFileContents = GetFileContents(sFilePath & "_EmployeesChangesLKP_" & Int(jIndex / ROWS_PER_FILE) & ".txt", sErrorDescription)
					If Len(asFileContents) > 0 Then
						asFileContents = Split(asFileContents, vbNewLine)
						For iIndex = 0 To UBound(asFileContents)
							If Len(asFileContents(iIndex)) > 0 Then
								asEmployeesQueries = Split(asFileContents(iIndex), ",")
								sErrorDescription = "No se pudo modificar la nómina del empleado."
								lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, Replace(Replace(Replace(Replace(sQueryBegin, "<FIRST_DATE />", asEmployeesQueries(1)), "<LAST_DATE />", asEmployeesQueries(2)), "<CONCEPTS_40 />", asEmployeesQueries(3)), "<EMPLOYEE_ID />", asEmployeesQueries(0)), "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription)
							End If
							If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit For
						Next
					End If
					Call DeleteFile(sFilePath & "_EmployeesChangesLKP_" & Int(jIndex / ROWS_PER_FILE) & ".txt", "")
				Next
			End If
		End If

		If Not bTimeout Then
			sErrorDescription = "No se pudieron obtener las últimas fechas de actualización de los empleados."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From EmployeesChangesLKP Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate<0)", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
		End If
	End If
	Application.Contents("SIAP_CalculatePayroll") = ""

	Set oRecordset = Nothing
	InsertEmployeesChangesLKP = lErrorNumber
	Err.Clear
End Function

Function BuildEmployeesChangesLKP(aPayrollComponent, bRetroactive, bAdjustment, sErrorDescription)
'************************************************************
'Purpose: To calculate the payroll and save it into the database
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "BuildEmployeesChangesLKP"
	Const ROWS_PER_FILE = 10000
	Dim lPayID
	Dim sFilePath
	Dim asFileContents
	Dim lStartDate
	Dim lEndDate
	Dim lTempStartDate
	Dim lTempEndDate
	Dim bCurrent
	Dim sQueryBegin
	Dim sQueryEnd
	Dim sCondition
	Dim sTable
	Dim lCurrentID
	Dim lCurrentID2
	Dim sCurrentID
	Dim dAmount
	Dim dTaxAmount
	Dim sEmployeeIDs
	Dim dTemp
	Dim sTemp
	Dim asSpecialConcepts
	Dim kIndex
	Dim oRecordset
	Dim lErrorNumber

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
									lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update EmployeesChangesLKP Set FirstDate=" & lTempStartDate & ", LastDate=" & lTempEndDate & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (-" & lPayID & "," & lPayID & ")) And (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
								End If
								lCurrentID = CLng(oRecordset.Fields("EmployeeID").Value)
								lTempStartDate = lStartDate
								lTempEndDate = lEndDate
							End If
							lTempStartDate = CLng(oRecordset.Fields("OcurredDate").Value)
							lTempEndDate = CLng(oRecordset.Fields("EndDate").Value)
							If lTempStartDate < lStartDate Then lTempStartDate = lStartDate
							If lTempEndDate > lEndDate Then lTempEndDate = lEndDate
							oRecordset.MoveNext
							'If Err.number <> 0 Then Exit Do
							If bTimeout Then Exit Do
						Loop
						If Not bTimeout Then
							If dAmount > 0 Then
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
								lErrorNumber = ExecuteUpdateQuerySp(oADODBConnection, "Update EmployeesChangesLKP Set FirstDate=" & lTempStartDate & ", LastDate=" & lTempEndDate & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate In (-" & lPayID & "," & lPayID & ")) And (EmployeeID=" & lCurrentID & ")", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription)
							End If
						End If
					End If
				End If
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
		End If
	End If

	If Not bAdjustment Then
Call DisplayTimeStamp("START: LEVEL 2, ESPECIALES. " & aPayrollComponent(N_FOR_DATE_PAYROLL))
		If Not bTimeout Then
			asSpecialConcepts = Split("40,41,42,43", ",")
			For kIndex = 0 To UBound(asSpecialConcepts)
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
																If dAmount > 3 Then sCurrentID = sCurrentID & "," & lCurrentID
																If dTemp > 3 Then sCurrentID = sCurrentID & "," & lCurrentID
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
													If dAmount > 3 Then sCurrentID = sCurrentID & "," & lCurrentID
													If dTemp > 3 Then sCurrentID = sCurrentID & "," & lCurrentID
												End If
												sErrorDescription = "No se pudo limpiar la tabla temporal."
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update EmployeesChangesLKP Set Concepts40=" & (dTaxAmount + 11) & " Where (PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (PayrollDate=" & lPayID & ") And (Concepts40<=" & dTaxAmount & ") And (EmployeeID In (" & sCurrentID & "))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

												sErrorDescription = "No se pudo limpiar la tabla temporal." 'Faltas y retardos
												lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll Where (EmployeeID In (" & sCurrentID & "))", "PayrollComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
												Response.Write "<!-- Query (" & asSpecialConcepts(kIndex) & "): Delete From Payroll Where (EmployeeID In (" & sCurrentID & ")) -->" & vbNewLine
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
			Next
		End If
	End If


	Set oRecordset = Nothing
	BuildEmployeesChangesLKP = lErrorNumber
	Err.Clear
End Function

Function GatherTableStats(oRequest, oADODBConnection, sTable, iPercent, sMethod, bCascade, sErrorDescription)
'************************************************************
'Purpose: To add a new payroll into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPayrollComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GatherTableStats"
	Dim lErrorNumber
	Dim bComponentInitialized

	Set oADODBCommand = Server.CreateObject("ADODB.Command")
	Set oADODBCommand.ActiveConnection = oADODBConnection
	oADODBCommand.commandtype=4
	oADODBCommand.commandtext = "SIAP.gatherTableStats"
	Set param = oADODBCommand.Parameters
	param.append oADODBCommand.createparameter("sTable", vbString, 1)
	param.append oADODBCommand.createparameter("iPercent", 3, 1)
	param.append oADODBCommand.createparameter("sMethod", vbString, 1)
	'param.append oADODBCommand.createparameter("bCascade", vbBoolean, 1)

	oADODBCommand("sTable") = sTable
	oADODBCommand("iPercent") = iPercent
	oADODBCommand("sMethod") = sMethod
	'oADODBCommand("bCascade") = bCascade

	oADODBCommand.Execute

	If Err.number < 0 Then
		lErrorNumber = -1
	Else
		lErrorNumber = 0
	End If

	Set oADODBCommand = Nothing
	Set param = Nothing
	GatherTableStats = lErrorNumber
	Err.Clear
End Function

Function GetPayrollCount(oRequest, oADODBConnection, lPayroll, sErrorDescription)
'************************************************************
'Purpose: To add a new payroll into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPayrollComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPayrollCount"
	Dim iPayrollCount

	Set oADODBCommand = Server.CreateObject("ADODB.Command")
	Set oADODBCommand.ActiveConnection = oADODBConnection
	oADODBCommand.commandtype=4
	oADODBCommand.commandtext = "SIAP.GetPayrollCount"
	Set param = oADODBCommand.Parameters
	param.append oADODBCommand.createparameter("lPayroll", 3, 1)
	param.append oADODBCommand.createparameter("iCount", 3, 2)

	oADODBCommand("lPayroll") = lPayroll

	oADODBCommand.Execute

	iPayrollCount = oADODBCommand("iCount")

	Set oADODBCommand = Nothing
	Set param = Nothing
	GetPayrollCount = iPayrollCount
	Err.Clear
End Function

Function GetPayrollConceptCount(oRequest, oADODBConnection, lPayroll, iConceptID, sErrorDescription)
'************************************************************
'Purpose: To add a new payroll into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPayrollComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPayrollConceptCount"
	Dim iPayrollCount

	Set oADODBCommand = Server.CreateObject("ADODB.Command")
	Set oADODBCommand.ActiveConnection = oADODBConnection
	oADODBCommand.commandtype=4
	oADODBCommand.commandtext = "SIAP.GetPayrollCptoCount"
	Set param = oADODBCommand.Parameters
	param.append oADODBCommand.createparameter("lPayroll", 3, 1)
	param.append oADODBCommand.createparameter("iConceptID", 3, 1)
	param.append oADODBCommand.createparameter("iCount", 3, 2)

	oADODBCommand("lPayroll") = lPayroll
	oADODBCommand("iConceptID") = iConceptID
	
	oADODBCommand.Execute

	iPayrollCount = oADODBCommand("iCount")

	Set oADODBCommand = Nothing
	Set param = Nothing
	GetPayrollConceptCount = iPayrollCount
	Err.Clear
End Function

Function GetPayrollConcept1Count(oRequest, oADODBConnection, lPayroll, iConceptID, sErrorDescription)
'************************************************************
'Purpose: To add a new payroll into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPayrollComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPayrollConcept1Count"
	Dim iPayrollCount

	Set oADODBCommand = Server.CreateObject("ADODB.Command")
	Set oADODBCommand.ActiveConnection = oADODBConnection
	oADODBCommand.commandtype=4
	oADODBCommand.commandtext = "SIAP.GetPayrollCpto1Count"
	Set param = oADODBCommand.Parameters
	param.append oADODBCommand.createparameter("lPayroll", 3, 1)
	param.append oADODBCommand.createparameter("iConceptID", 3, 1)
	param.append oADODBCommand.createparameter("iCount", 3, 2)

	oADODBCommand("lPayroll") = lPayroll
	oADODBCommand("iConceptID") = iConceptID
	
	oADODBCommand.Execute

	iPayrollCount = oADODBCommand("iCount")

	Set oADODBCommand = Nothing
	Set param = Nothing
	GetPayrollConcept1Count = iPayrollCount
	Err.Clear
End Function
%>