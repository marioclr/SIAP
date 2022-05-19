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
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "CREATE TABLE Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " (RecordDate int NOT NULL, RecordID int NOT NULL, EmployeeID INTEGER NOT NULL, ConceptID INTEGER NOT NULL, PayrollTypeID INTEGER NOT NULL, ConceptAmount FLOAT (8) NOT NULL, ConceptTaxes FLOAT (8) NOT NULL, ConceptRetention FLOAT (8) NOT NULL, UserID INTEGER NOT NULL, PRIMARY KEY (RecordDate, RecordID, EmployeeID, ConceptID))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
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

	bTruncate = True
	sTruncate = ""
	bComponentInitialized = aPayrollComponent(B_COMPONENT_INITIALIZED_PAYROLL)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePayrollComponent(oRequest, aPayrollComponent)
	End If

	lErrorNumber = GetPayroll(oRequest, oADODBConnection, aPayrollComponent, sErrorDescription)
	If lErrorNumber = 0 Then
		aPayrollComponent(N_FOR_DATE_PAYROLL) = CLng(aPayrollComponent(N_FOR_DATE_PAYROLL))
		lErrorNumber = DoCalculations(aPayrollComponent, False, False, sErrorDescription)
	End If
	Application.Contents("SIAP_CalculatePayroll") = ""

	Set oRecordset = Nothing
	CalculatePayroll = lErrorNumber
	Err.Clear
End Function

Function CalculateQttyID_8_9(oRequest, oADODBConnection, bCurrent, bRetroactive, sErrorDescription)
'************************************************************
'Purpose: To calculate the amount for extra hours and sundays
'Inputs:  oRequest, oADODBConnection, bMonthlyTaxes
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CalculateQttyID_8_9"
	Const ROWS_PER_FILE = 10000
	Const CONCEPTS_FOR_FACTOR = "1,3,13,14,38,49,89"
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
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & aPayrollComponent(N_ID_PAYROLL) & " Where (ConceptID=16) And (RecordDate=" & lForPayrollDate & ") And (EmployeeID In (Select EmployeesHistoryList.EmployeeID From EmployeesChangesLKP, EmployeesHistoryList Where (EmployeesChangesLKP.EmployeeID=EmployeesHistoryList.EmployeeID) And (EmployeesChangesLKP.EmployeeDate=EmployeesHistoryList.EmployeeDate) And (EmployeesChangesLKP.PayrollID=" & aPayrollComponent(N_ID_PAYROLL) & ") And (EmployeesChangesLKP.PayrollDate=" & lForPayrollDate & ") And (EmployeesHistoryList.EmployeeDate<=EmployeesHistoryList.EndDate) And (EmployeesHistoryList.EmployeeTypeID Not In (0,2,3,4))))", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
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
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ConceptsValues.*, Concepts.PeriodID, Antiquities.StartYears, Antiquities.EndYears, Antiquities2.StartYears As StartYears2, Antiquities2.EndYears As EndYears2, Antiquities3.StartYears As StartYears3, Antiquities3.EndYears As EndYears3, Antiquities4.StartYears As StartYears4, Antiquities4.EndYears As EndYears4 From ConceptsValues, Concepts, Antiquities, Antiquities As Antiquities2, Antiquities As Antiquities3, Antiquities As Antiquities4 Where (ConceptsValues.ConceptID=Concepts.ConceptID) And (ConceptsValues.AntiquityID=Antiquities.AntiquityID) And (ConceptsValues.Antiquity2ID=Antiquities2.AntiquityID) And (ConceptsValues.Antiquity3ID=Antiquities3.AntiquityID) And (ConceptsValues.Antiquity4ID=Antiquities4.AntiquityID) And (ConceptsValues.StartDate<=" & lPayrollID & ") And (ConceptsValues.EndDate>=" & lPayrollID & ") And (Concepts.StartDate<=" & lPayrollID & ") And (Concepts.EndDate>=" & lPayrollID & ") And (ConceptQttyID=1) " & sConceptCondition & " Order By Concepts.OrderInList, Concepts.ConceptID, CompanyID Desc, EmployeeTypeID Desc, PositionTypeID Desc, EmployeeStatusID Desc, JobStatusID Desc, ClassificationID Desc, GroupGradeLevelID Desc, IntegrationID Desc, JourneyID Desc, WorkingHours Desc, AdditionalShift Desc, LevelID Desc, EconomicZoneID Desc, ServiceID Desc, ConceptsValues.AntiquityID Desc, ConceptsValues.Antiquity2ID Desc, ConceptsValues.Antiquity3ID Desc, ConceptsValues.Antiquity4ID Desc, ForRisk Desc, GenderID Desc, HasSyndicate Desc", "PayrollComponentConstants.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
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

				sFlags = L_NO_INSTRUCTIONS_FLAGS & "," & L_OPEN_PAYROLL_FLAGS & "," & L_EMPLOYEE_NUMBER_FLAGS & "," & L_COMPANY_FLAGS & "," & L_EMPLOYEE_TYPE_FLAGS & "," & L_POSITION_TYPE_FLAGS & "," & L_CLASSIFICATION_FLAGS & "," & L_GROUP_GRADE_LEVEL_FLAGS & "," & L_INTEGRATION_FLAGS & "," & L_JOURNEY_FLAGS & "," & L_SHIFT_FLAGS & "," & L_LEVEL_FLAGS & "," & L_PAYMENT_CENTER_FLAGS & "," & L_ZONE_FLAGS & "," & L_AREA_FLAGS & "," & L_POSITION_FLAGS
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
	If InStr(1, ",0,2,243,", "," & aLoginComponent(N_USER_ID_LOGIN) & ",", vbBinaryCompare) > 0 Then
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
%>