<%
Const N_ID_PAYMENT = 0
Const S_CHECK_NUMBER_PAYMENT = 1
Const S_REPLACEMENT_NUMBER_PAYMENT = 2
Const N_EMPOYEE_ID_PAYMENT = 3
Const N_PAYMENT_TYPE_ID_PAYMENT = 4
Const N_DATE_PAYMENT = 5
Const N_REGISTERED_DATE_PAYMENT = 6
Const N_CHECK_DATE_PAYMENT = 7
Const N_ACCOUNT_ID_PAYMENT = 8
Const N_FROM_ACCOUNT_ID_PAYMENT = 9
Const D_CHECK_AMOUNT_PAYMENT = 10
Const N_CHECK_CURRENCY_ID_PAYMENT = 11
Const N_STATUS_ID_PAYMENT = 12
Const S_DESCRIPTION_PAYMENT = 13
Const B_IS_PAYMENT = 14
Const B_IS_UPDATED_PAYMENT = 15
Const N_REPLACEMENT_USER_ID_PAYMENT = 16
Const S_QUERY_CONDITION_PAYMENT = 17
Const S_SORT_COLUMN_PAYMENT = 18
Const B_SORT_DESCENDING_PAYMENT = 19
Const B_CHECK_FOR_DUPLICATED_PAYMENT = 20
Const B_IS_DUPLICATED_PAYMENT = 21
Const B_COMPONENT_INITIALIZED_PAYMENT = 22

Const N_PAYMENT_COMPONENT_SIZE = 22

Dim aPaymentComponent()
Redim aPaymentComponent(N_PAYMENT_COMPONENT_SIZE)

Function InitializePaymentComponent(oRequest, aPaymentComponent)
'************************************************************
'Purpose: To initialize the empty elements of the Payment Component
'         using the URL parameters or default values
'Inputs:  oRequest
'Outputs: aPaymentComponent
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "InitializePaymentComponent"
	Dim oItem
	Redim Preserve aPaymentComponent(N_PAYMENT_COMPONENT_SIZE)

	If IsEmpty(aPaymentComponent(N_ID_PAYMENT)) Then
		If Len(oRequest("PaymentID").Item) > 0 Then
			aPaymentComponent(N_ID_PAYMENT) = CLng(oRequest("PaymentID").Item)
		Else
			aPaymentComponent(N_ID_PAYMENT) = -1
		End If
	End If

	If IsEmpty(aPaymentComponent(S_CHECK_NUMBER_PAYMENT)) Then
		If Len(oRequest("CheckNumber").Item) > 0 Then
			aPaymentComponent(S_CHECK_NUMBER_PAYMENT) = oRequest("CheckNumber").Item
		Else
			aPaymentComponent(S_CHECK_NUMBER_PAYMENT) = ""
		End If
	End If
	aPaymentComponent(S_CHECK_NUMBER_PAYMENT) = Left(aPaymentComponent(S_CHECK_NUMBER_PAYMENT), 100)

	If IsEmpty(aPaymentComponent(S_REPLACEMENT_NUMBER_PAYMENT)) Then
		If Len(oRequest("ReplacementNumber").Item) > 0 Then
			aPaymentComponent(S_REPLACEMENT_NUMBER_PAYMENT) = oRequest("ReplacementNumber").Item
		Else
			aPaymentComponent(S_REPLACEMENT_NUMBER_PAYMENT) = ""
		End If
	End If
	aPaymentComponent(S_REPLACEMENT_NUMBER_PAYMENT) = Left(aPaymentComponent(S_REPLACEMENT_NUMBER_PAYMENT), 100)

	If IsEmpty(aPaymentComponent(N_EMPOYEE_ID_PAYMENT)) Then
		If Len(oRequest("EmployeeID").Item) > 0 Then
			aPaymentComponent(N_EMPOYEE_ID_PAYMENT) = CLng(oRequest("EmployeeID").Item)
		Else
			aPaymentComponent(N_EMPOYEE_ID_PAYMENT) = -1
		End If
	End If

	If IsEmpty(aPaymentComponent(N_PAYMENT_TYPE_ID_PAYMENT)) Then
		If Len(oRequest("PaymentTypeID").Item) > 0 Then
			aPaymentComponent(N_PAYMENT_TYPE_ID_PAYMENT) = CLng(oRequest("PaymentTypeID").Item)
		Else
			aPaymentComponent(N_PAYMENT_TYPE_ID_PAYMENT) = -1
		End If
	End If

	If IsEmpty(aPaymentComponent(N_DATE_PAYMENT)) Then
		If Len(oRequest("PaymentYear").Item) > 0 Then
			aPaymentComponent(N_DATE_PAYMENT) = CLng(oRequest("PaymentYear").Item & Right(("0" & oRequest("PaymentMonth").Item), Len("00")) & Right(("0" & oRequest("PaymentDay").Item), Len("00")))
		ElseIf Len(oRequest("PaymentDate").Item) > 0 Then
			aPaymentComponent(N_DATE_PAYMENT) = CLng(oRequest("PaymentDate").Item)
		Else
			aPaymentComponent(N_DATE_PAYMENT) = Left(GetSerialNumberForDate(""), Len("00000000"))
		End If
	End If

	If IsEmpty(aPaymentComponent(N_REGISTERED_DATE_PAYMENT)) Then
		If Len(oRequest("RegisteredYear").Item) > 0 Then
			aPaymentComponent(N_REGISTERED_DATE_PAYMENT) = CLng(oRequest("RegisteredYear").Item & Right(("0" & oRequest("RegisteredMonth").Item), Len("00")) & Right(("0" & oRequest("RegisteredDay").Item), Len("00")))
		ElseIf Len(oRequest("RegisteredDate").Item) > 0 Then
			aPaymentComponent(N_REGISTERED_DATE_PAYMENT) = CLng(oRequest("RegisteredDate").Item)
		Else
			aPaymentComponent(N_REGISTERED_DATE_PAYMENT) = Left(GetSerialNumberForDate(""), Len("00000000"))
		End If
	End If

	If IsEmpty(aPaymentComponent(N_CHECK_DATE_PAYMENT)) Then
		If Len(oRequest("CheckYear").Item) > 0 Then
			aPaymentComponent(N_CHECK_DATE_PAYMENT) = CLng(oRequest("CheckYear").Item & Right(("0" & oRequest("CheckMonth").Item), Len("00")) & Right(("0" & oRequest("CheckDay").Item), Len("00")))
		ElseIf Len(oRequest("CheckDate").Item) > 0 Then
			aPaymentComponent(N_CHECK_DATE_PAYMENT) = CLng(oRequest("CheckDate").Item)
		Else
			aPaymentComponent(N_CHECK_DATE_PAYMENT) = Left(GetSerialNumberForDate(""), Len("00000000"))
		End If
	End If

	If IsEmpty(aPaymentComponent(N_ACCOUNT_ID_PAYMENT)) Then
		If Len(oRequest("AccountID").Item) > 0 Then
			aPaymentComponent(N_ACCOUNT_ID_PAYMENT) = CLng(oRequest("AccountID").Item)
		Else
			aPaymentComponent(N_ACCOUNT_ID_PAYMENT) = -1
		End If
	End If

	If IsEmpty(aPaymentComponent(N_FROM_ACCOUNT_ID_PAYMENT)) Then
		If Len(oRequest("FromAccountID").Item) > 0 Then
			aPaymentComponent(N_FROM_ACCOUNT_ID_PAYMENT) = CLng(oRequest("FromAccountID").Item)
		Else
			aPaymentComponent(N_FROM_ACCOUNT_ID_PAYMENT) = -1
		End If
	End If

	If IsEmpty(aPaymentComponent(D_CHECK_AMOUNT_PAYMENT)) Then
		If Len(oRequest("CheckAmount").Item) > 0 Then
			aPaymentComponent(D_CHECK_AMOUNT_PAYMENT) = CDbl(oRequest("CheckAmount").Item)
		Else
			aPaymentComponent(D_CHECK_AMOUNT_PAYMENT) = 0
		End If
	End If

	If IsEmpty(aPaymentComponent(N_CHECK_CURRENCY_ID_PAYMENT)) Then
		If Len(oRequest("CheckCurrencyID").Item) > 0 Then
			aPaymentComponent(N_CHECK_CURRENCY_ID_PAYMENT) = CLng(oRequest("CheckCurrencyID").Item)
		Else
			aPaymentComponent(N_CHECK_CURRENCY_ID_PAYMENT) = 0
		End If
	End If

	If IsEmpty(aPaymentComponent(N_STATUS_ID_PAYMENT)) Then
		If Len(oRequest("StatusID").Item) > 0 Then
			aPaymentComponent(N_STATUS_ID_PAYMENT) = CLng(oRequest("StatusID").Item)
		Else
			aPaymentComponent(N_STATUS_ID_PAYMENT) = -1
		End If
	End If

	If IsEmpty(aPaymentComponent(S_DESCRIPTION_PAYMENT)) Then
		If Len(oRequest("Description").Item) > 0 Then
			aPaymentComponent(S_DESCRIPTION_PAYMENT) = oRequest("Description").Item
		Else
			aPaymentComponent(S_DESCRIPTION_PAYMENT) = ""
		End If
	End If
	aPaymentComponent(S_DESCRIPTION_PAYMENT) = Left(aPaymentComponent(S_DESCRIPTION_PAYMENT), 2000)

	If IsEmpty(aPaymentComponent(B_IS_PAYMENT)) Then
		If Len(oRequest("bIsPayment").Item) > 0 Then
			aPaymentComponent(B_IS_PAYMENT) = CInt(oRequest("bIsPayment").Item)
		Else
			aPaymentComponent(B_IS_PAYMENT) = 1
		End If
	End If

	If IsEmpty(aPaymentComponent(B_IS_UPDATED_PAYMENT)) Then
		If Len(oRequest("bIsUpdated").Item) > 0 Then
			aPaymentComponent(B_IS_UPDATED_PAYMENT) = CInt(oRequest("bIsUpdated").Item)
		Else
			aPaymentComponent(B_IS_UPDATED_PAYMENT) = 1
		End If
	End If

	aPaymentComponent(N_REPLACEMENT_USER_ID_PAYMENT) = -1
	If StrComp(aPaymentComponent(S_REPLACEMENT_NUMBER_PAYMENT), oRequest("PreviousReplacementNumber").Item, vbBinaryCompare) <> 0 Then
		aPaymentComponent(N_REPLACEMENT_USER_ID_PAYMENT) = aLoginComponent(N_USER_ID_LOGIN)
	End If

	aPaymentComponent(S_QUERY_CONDITION_PAYMENT) = ""
	If Len(oRequest("SortColumn").Item) > 0 Then
		aPaymentComponent(S_SORT_COLUMN_PAYMENT) = oRequest("SortColumn").Item
		Call SetOption(aOptionsComponent, PAYMENT_ORDER_OPTION, aPaymentComponent(S_SORT_COLUMN_PAYMENT), sErrorDescription)
	Else
		aPaymentComponent(S_SORT_COLUMN_PAYMENT) = GetOption(aOptionsComponent, PAYMENT_ORDER_OPTION)
	End If
	If Len(oRequest("Desc").Item) > 0 Then
		aPaymentComponent(B_SORT_DESCENDING_PAYMENT) = (StrComp(oRequest("Desc").Item, "1", vbBinaryCompare) = 0)
		Call SetOption(aOptionsComponent, PAYMENT_SORT_OPTION, oRequest("Desc").Item, sErrorDescription)
	Else
		aPaymentComponent(B_SORT_DESCENDING_PAYMENT) = (StrComp(GetOption(aOptionsComponent, PAYMENT_SORT_OPTION), "0", vbBinaryCompare) = 0)
	End If
	Call ModifyOptions(oRequest, oADODBConnection, aOptionsComponent, sErrorDescription)
	aPaymentComponent(B_CHECK_FOR_DUPLICATED_PAYMENT) = True
	aPaymentComponent(B_IS_DUPLICATED_PAYMENT) = False

	aPaymentComponent(B_COMPONENT_INITIALIZED_PAYMENT) = True
	InitializePaymentComponent = Err.number
	Err.Clear
End Function

Function AddPayment(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
'************************************************************
'Purpose: To add a new payment into the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPaymentComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddPayment"
	Dim sDate
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPaymentComponent(B_COMPONENT_INITIALIZED_PAYMENT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePaymentComponent(oRequest, aPaymentComponent)
	End If

	If aPaymentComponent(N_ID_PAYMENT) = -1 Then
		sErrorDescription = "No se pudo obtener un identificador para el nuevo cheque."
		lErrorNumber = GetNewIDFromTable(oADODBConnection, "Payments", "PaymentID", "", 1, aPaymentComponent(N_ID_PAYMENT), sErrorDescription)
	End If

	If lErrorNumber = 0 Then
		If aPaymentComponent(B_CHECK_FOR_DUPLICATED_PAYMENT) Then
			lErrorNumber = CheckExistencyOfPayment(aPaymentComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aPaymentComponent(B_IS_DUPLICATED_PAYMENT) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un cheque con el número " & aPaymentComponent(S_CHECK_NUMBER_PAYMENT) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PaymentComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				If Not CheckPaymentInformationConsistency(aPaymentComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					sErrorDescription = "No se pudo guardar la información del nuevo cheque."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payments (PaymentID, CheckNumber, ReplacementNumber, EmployeeID, PaymentTypeID, PaymentDate, RegisteredDate, CheckDate, CancelDate, AccountID, FromAccountID, CheckAmount, CheckCurrencyID, StatusID, Description, bIsPayment, bIsUpdated, LastUpdate, UserID, ReplacementUserID) Values (" & aPaymentComponent(N_ID_PAYMENT) & ", '" & Replace(aPaymentComponent(S_CHECK_NUMBER_PAYMENT), "'", "") & "', '" & Replace(aPaymentComponent(S_REPLACEMENT_NUMBER_PAYMENT), "'", "") & "', " & aPaymentComponent(N_EMPOYEE_ID_PAYMENT) & ", " & aPaymentComponent(N_PAYMENT_TYPE_ID_PAYMENT) & ", " & aPaymentComponent(N_DATE_PAYMENT) & ", " & aPaymentComponent(N_REGISTERED_DATE_PAYMENT) & ", " & aPaymentComponent(N_CHECK_DATE_PAYMENT) & ", 0, " & aPaymentComponent(N_ACCOUNT_ID_PAYMENT) & ", " & aPaymentComponent(N_FROM_ACCOUNT_ID_PAYMENT) & ", " & aPaymentComponent(D_CHECK_AMOUNT_PAYMENT) & ", " & aPaymentComponent(N_CHECK_CURRENCY_ID_PAYMENT) & ", " & aPaymentComponent(N_STATUS_ID_PAYMENT) & ", '" & Replace(aPaymentComponent(S_DESCRIPTION_PAYMENT), "'", "´") & "', " & aPaymentComponent(B_IS_PAYMENT) & ", " & aPaymentComponent(B_IS_UPDATED_PAYMENT) & ", " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", -1)", "PaymentComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
			End If
		End If
	End If

	AddPayment = lErrorNumber
	Err.Clear
End Function

Function GetPayment(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
'************************************************************
'Purpose: To get the information about an payment from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPaymentComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPayment"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPaymentComponent(B_COMPONENT_INITIALIZED_PAYMENT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePaymentComponent(oRequest, aPaymentComponent)
	End If

	If aPaymentComponent(N_ID_PAYMENT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del cheque para obtener su información."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PaymentComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo obtener la información del cheque."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Payments Where PaymentID=" & aPaymentComponent(N_ID_PAYMENT), "PaymentComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If oRecordset.EOF Then
				lErrorNumber = L_ERR_NO_RECORDS
				sErrorDescription = "El cheque especificado no se encuentra en el sistema."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PaymentComponent.asp", S_FUNCTION_NAME, N_MESSAGE_LEVEL)
			Else
				aPaymentComponent(S_CHECK_NUMBER_PAYMENT) = CStr(oRecordset.Fields("CheckNumber").Value)
				aPaymentComponent(S_REPLACEMENT_NUMBER_PAYMENT) = CStr(oRecordset.Fields("ReplacementNumber").Value)
				aPaymentComponent(N_EMPOYEE_ID_PAYMENT) = CLng(oRecordset.Fields("EmployeeID").Value)
				aPaymentComponent(N_PAYMENT_TYPE_ID_PAYMENT) = CLng(oRecordset.Fields("PaymentTypeID").Value)
				aPaymentComponent(N_DATE_PAYMENT) = CLng(oRecordset.Fields("PaymentDate").Value)
				aPaymentComponent(N_REGISTERED_DATE_PAYMENT) = CLng(oRecordset.Fields("RegisteredDate").Value)
				aPaymentComponent(N_CHECK_DATE_PAYMENT) = CLng(oRecordset.Fields("CheckDate").Value)
				aPaymentComponent(N_ACCOUNT_ID_PAYMENT) = CLng(oRecordset.Fields("AccountID").Value)
				aPaymentComponent(N_FROM_ACCOUNT_ID_PAYMENT) = CLng(oRecordset.Fields("FromAccountID").Value)
				aPaymentComponent(D_CHECK_AMOUNT_PAYMENT) = CDbl(oRecordset.Fields("CheckAmount").Value)
				aPaymentComponent(N_CHECK_CURRENCY_ID_PAYMENT) = CLng(oRecordset.Fields("CheckCurrencyID").Value)
				aPaymentComponent(N_STATUS_ID_PAYMENT) = CLng(oRecordset.Fields("StatusID").Value)
				aPaymentComponent(S_DESCRIPTION_PAYMENT) = CStr(oRecordset.Fields("Description").Value)
				aPaymentComponent(B_IS_PAYMENT) = CInt(oRecordset.Fields("bIsPayment").Value)
				aPaymentComponent(B_IS_UPDATED_PAYMENT) = CInt(oRecordset.Fields("bIsUpdated").Value)
			End If
			oRecordset.Close
		End If
	End If

	Set oRecordset = Nothing
	GetPayment = lErrorNumber
	Err.Clear
End Function

Function GetPayments(oRequest, oADODBConnection, aPaymentComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the payments from
'		  the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPaymentComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPayments"
	Dim sSort
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPaymentComponent(B_COMPONENT_INITIALIZED_PAYMENT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePaymentComponent(oRequest, aPaymentComponent)
	End If

	aPaymentComponent(S_QUERY_CONDITION_PAYMENT) = Trim(aPaymentComponent(S_QUERY_CONDITION_PAYMENT))
	If Len(aPaymentComponent(S_QUERY_CONDITION_PAYMENT)) > 0 Then
		If InStr(1, aPaymentComponent(S_QUERY_CONDITION_PAYMENT), "And ", vbBinaryCompare) <> 1 Then aPaymentComponent(S_QUERY_CONDITION_PAYMENT) = "And " & aPaymentComponent(S_QUERY_CONDITION_PAYMENT)
	End If
	sSort = aPaymentComponent(S_SORT_COLUMN_PAYMENT)
	If aPaymentComponent(B_SORT_DESCENDING_PAYMENT) Then sSort = Replace(sSort, ", ", " Desc, ") & " Desc"
	sErrorDescription = "No se pudo obtener la información de los cheques."
	If (Len(oRequest("Action").Item) > 0) And (StrComp(oRequest("Action").Item,"EmployeesMovements",vbBinaryCompare) = 0) Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "SELECT PaymentID,PaymentDate,CancelDate, CheckNumber, StatusName, StatusShortName, CheckAmount FROM Payments, StatusPayments WHERE (Payments.StatusID = StatusPayments.StatusID) " & aPaymentComponent(S_QUERY_CONDITION_PAYMENT) & " ORDER BY Payments.EmployeeID DESC", "PaymentComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payments.*, EmployeeName, EmployeeLastName, EmployeeLastName2, PaymentTypeName, BankAccounts.AccountNumber, FromBankAccounts.AccountNumber As FromAccountNumber, CurrencyName, CurrencySymbol, StatusName From Payments, Employees, PaymentTypes, BankAccounts, BankAccounts As FromBankAccounts, Currencies, StatusPayments Where (Payments.EmployeeID=Employees.EmployeeID) And (Payments.PaymentTypeID=PaymentTypes.PaymentTypeID) And (Payments.AccountID=BankAccounts.AccountID) And (Payments.AccountID=FromBankAccounts.AccountID) And (Payments.CheckCurrencyID=Currencies.CurrencyID) And (Payments.StatusID=StatusPayments.StatusID) " & aPaymentComponent(S_QUERY_CONDITION_PAYMENT) & " Order By " & sSort, "PaymentComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If

	GetPayments = lErrorNumber
	Err.Clear
End Function

Function ModifyPayment(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
'************************************************************
'Purpose: To modify an existing payment in the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPaymentComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyPayment"
	Dim sDate
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPaymentComponent(B_COMPONENT_INITIALIZED_PAYMENT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePaymentComponent(oRequest, aPaymentComponent)
	End If

	If aPaymentComponent(N_ID_PAYMENT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el identificador del cheque a modificar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PaymentComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		If aPaymentComponent(B_CHECK_FOR_DUPLICATED_PAYMENT) Then
			lErrorNumber = CheckExistencyOfPayment(aPaymentComponent, sErrorDescription)
		End If

		If lErrorNumber = 0 Then
			If aPaymentComponent(B_IS_DUPLICATED_PAYMENT) Then
				lErrorNumber = L_ERR_DUPLICATED_RECORD
				sErrorDescription = "Ya existe un cheque con el número " & aPaymentComponent(S_CHECK_NUMBER_PAYMENT) & "."
				Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PaymentComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
			Else
				If Not CheckPaymentInformationConsistency(aPaymentComponent, sErrorDescription) Then
					lErrorNumber = -1
				Else
					sErrorDescription = "No se pudo modificar la información del cheque."
					If aPaymentComponent(N_REPLACEMENT_USER_ID_PAYMENT) = -1 Then
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payments Set CheckNumber='" & Replace(aPaymentComponent(S_CHECK_NUMBER_PAYMENT), "'", "") & "', ReplacementNumber='" & Replace(aPaymentComponent(S_REPLACEMENT_NUMBER_PAYMENT), "'", "") & "', EmployeeID=" & aPaymentComponent(N_EMPOYEE_ID_PAYMENT) & ", PaymentTypeID=" & aPaymentComponent(N_PAYMENT_TYPE_ID_PAYMENT) & ", PaymentDate=" & aPaymentComponent(N_DATE_PAYMENT) & ", RegisteredDate=" & aPaymentComponent(N_REGISTERED_DATE_PAYMENT) & ", CheckDate=" & aPaymentComponent(N_CHECK_DATE_PAYMENT) & ", AccountID=" & aPaymentComponent(N_ACCOUNT_ID_PAYMENT) & ", FromAccountID=" & aPaymentComponent(N_FROM_ACCOUNT_ID_PAYMENT) & ", CheckAmount=" & aPaymentComponent(D_CHECK_AMOUNT_PAYMENT) & ", CheckCurrencyID=" & aPaymentComponent(N_CHECK_CURRENCY_ID_PAYMENT) & ", StatusID=" & aPaymentComponent(N_STATUS_ID_PAYMENT) & ", Description='" & Replace(aPaymentComponent(S_DESCRIPTION_PAYMENT), "'", "´") & "', bIsPayment=" & aPaymentComponent(B_IS_PAYMENT) & ", bIsUpdated=" & aPaymentComponent(B_IS_UPDATED_PAYMENT) & ", LastUpdate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", UserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (PaymentID=" & aPaymentComponent(N_ID_PAYMENT) & ")", "PaymentComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					Else
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payments Set CheckNumber='" & Replace(aPaymentComponent(S_CHECK_NUMBER_PAYMENT), "'", "") & "', ReplacementNumber='" & Replace(aPaymentComponent(S_REPLACEMENT_NUMBER_PAYMENT), "'", "") & "', EmployeeID=" & aPaymentComponent(N_EMPOYEE_ID_PAYMENT) & ", PaymentTypeID=" & aPaymentComponent(N_PAYMENT_TYPE_ID_PAYMENT) & ", PaymentDate=" & aPaymentComponent(N_DATE_PAYMENT) & ", RegisteredDate=" & aPaymentComponent(N_REGISTERED_DATE_PAYMENT) & ", CheckDate=" & aPaymentComponent(N_CHECK_DATE_PAYMENT) & ", AccountID=" & aPaymentComponent(N_ACCOUNT_ID_PAYMENT) & ", FromAccountID=" & aPaymentComponent(N_FROM_ACCOUNT_ID_PAYMENT) & ", CheckAmount=" & aPaymentComponent(D_CHECK_AMOUNT_PAYMENT) & ", CheckCurrencyID=" & aPaymentComponent(N_CHECK_CURRENCY_ID_PAYMENT) & ", StatusID=" & aPaymentComponent(N_STATUS_ID_PAYMENT) & ", Description='" & Replace(aPaymentComponent(S_DESCRIPTION_PAYMENT), "'", "´") & "', bIsPayment=" & aPaymentComponent(B_IS_PAYMENT) & ", bIsUpdated=" & aPaymentComponent(B_IS_UPDATED_PAYMENT) & ", LastUpdate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", ReplacementUserID=" & aPaymentComponent(N_REPLACEMENT_USER_ID_PAYMENT) & " Where (PaymentID=" & aPaymentComponent(N_ID_PAYMENT) & ")", "PaymentComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					End If
				End If
			End If
		End If
	End If

	ModifyPayment = lErrorNumber
	Err.Clear
End Function

Function RemovePayment(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
'************************************************************
'Purpose: To remove an payment from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aPaymentComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemovePayment"
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPaymentComponent(B_COMPONENT_INITIALIZED_PAYMENT)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePaymentComponent(oRequest, aPaymentComponent)
	End If

	If aPaymentComponent(N_ID_PAYMENT) = -1 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el cheque a eliminar."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PaymentComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo eliminar la información del cheque."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payments Where (PaymentID=" & aPaymentComponent(N_ID_PAYMENT) & ")", "PaymentComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	RemovePayment = lErrorNumber
	Err.Clear
End Function

Function CheckExistencyOfPayment(aPaymentComponent, sErrorDescription)
'************************************************************
'Purpose: To check if a specific payment exists in the database
'Inputs:  aPaymentComponent
'Outputs: aPaymentComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckExistencyOfPayment"
	Dim oRecordset
	Dim lErrorNumber
	Dim bComponentInitialized

	bComponentInitialized = aPaymentComponent(B_COMPONENT_INITIALIZED_JOB)
	If (IsEmpty(bComponentInitialized)) Or (Not bComponentInitialized) Then
		Call InitializePaymentComponent(oRequest, aPaymentComponent)
	End If

	If Len(aPaymentComponent(S_CHECK_NUMBER_PAYMENT)) = 0 Then
		lErrorNumber = -1
		sErrorDescription = "No se especificó el número del cheque para revisar su existencia en la base de datos."
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PaymentComponent.asp", S_FUNCTION_NAME, N_WARNING_LEVEL)
	Else
		sErrorDescription = "No se pudo revisar la existencia del cheque en la base de datos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Payments Where (PaymentID<>" & aPaymentComponent(N_ID_PAYMENT) & ") And (CheckNumber='" & Replace(aPaymentComponent(S_CHECK_NUMBER_PAYMENT), "'", "") & "')", "PaymentComponent.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				aPaymentComponent(B_IS_DUPLICATED_PAYMENT) = True
				aPaymentComponent(N_ID_PAYMENT) = -1
			End If
		End If
		oRecordset.Close
	End If

	Set oRecordset = Nothing
	CheckExistencyOfPayment = lErrorNumber
	Err.Clear
End Function

Function CheckPaymentInformationConsistency(aPaymentComponent, sErrorDescription)
'************************************************************
'Purpose: To check for errors in the information that is
'		  going to be added into the database
'Inputs:  aPaymentComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "CheckPaymentInformationConsistency"
	Dim bIsCorrect

	bIsCorrect = True

	If Not IsNumeric(aPaymentComponent(N_ID_PAYMENT)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El identificador del cheque no es un valor numérico."
		bIsCorrect = False
	End If
	If Len(aPaymentComponent(S_CHECK_NUMBER_PAYMENT)) = 0 Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- El número del cheque está vacío."
		bIsCorrect = False
	End If
	If Not IsNumeric(aPaymentComponent(N_EMPOYEE_ID_PAYMENT)) Then
		sErrorDescription = sErrorDescription & "<BR />&nbsp;- No se especificó el número de empleado."
		bIsCorrect = False
	End If
	If Not IsNumeric(aPaymentComponent(N_PAYMENT_TYPE_ID_PAYMENT)) Then aPaymentComponent(N_PAYMENT_TYPE_ID_PAYMENT) = -1
	If Not IsNumeric(aPaymentComponent(N_DATE_PAYMENT)) Then aPaymentComponent(N_DATE_PAYMENT) = Left(GetSerialNumberForDate(""), Len("00000000"))
	If Not IsNumeric(aPaymentComponent(N_REGISTERED_DATE_PAYMENT)) Then aPaymentComponent(N_REGISTERED_DATE_PAYMENT) = Left(GetSerialNumberForDate(""), Len("00000000"))
	If Not IsNumeric(aPaymentComponent(N_CHECK_DATE_PAYMENT)) Then aPaymentComponent(N_CHECK_DATE_PAYMENT) = Left(GetSerialNumberForDate(""), Len("00000000"))
	If Not IsNumeric(aPaymentComponent(N_ACCOUNT_ID_PAYMENT)) Then aPaymentComponent(N_ACCOUNT_ID_PAYMENT) = -1
	If Not IsNumeric(aPaymentComponent(N_FROM_ACCOUNT_ID_PAYMENT)) Then aPaymentComponent(N_FROM_ACCOUNT_ID_PAYMENT) = -1
	If Not IsNumeric(aPaymentComponent(D_CHECK_AMOUNT_PAYMENT)) Then aPaymentComponent(D_CHECK_AMOUNT_PAYMENT) = 0
	If Not IsNumeric(aPaymentComponent(N_CHECK_CURRENCY_ID_PAYMENT)) Then aPaymentComponent(N_CHECK_CURRENCY_ID_PAYMENT) = 0
	If Not IsNumeric(aPaymentComponent(N_STATUS_ID_PAYMENT)) Then aPaymentComponent(N_STATUS_ID_PAYMENT) = -1
	If Not IsNumeric(aPaymentComponent(B_IS_PAYMENT)) Then aPaymentComponent(B_IS_PAYMENT) = 1
	If Not IsNumeric(aPaymentComponent(B_IS_UPDATED_PAYMENT)) Then aPaymentComponent(B_IS_UPDATED_PAYMENT) = 1

	If Len(sErrorDescription) > 0 Then
		sErrorDescription = "La información del cheque contiene campos con valores erróneos: " & sErrorDescription
		Call LogErrorInXMLFile(lErrorNumber, sErrorDescription, 000, "PaymentComponent.asp", S_FUNCTION_NAME, N_ERROR_LEVEL)
	End If

	CheckPaymentInformationConsistency = bIsCorrect
	Err.Clear
End Function

Function DisplayPayment(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an payment from the
'		  database
'Inputs:  oRequest, oADODBConnection, sAction, aPaymentComponent
'Outputs: aPaymentComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPayment"
	Dim sNames
	Dim oRecordset
	Dim lErrorNumber

	If aPaymentComponent(N_ID_PAYMENT) <> -1 Then
		lErrorNumber = GetPayment(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><FONT FACE=""Arial"" SIZE=""2"">"
			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Número de cheque:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aPaymentComponent(S_CHECK_NUMBER_PAYMENT)) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "Employees", aPaymentComponent(N_EMPOYEE_ID_PAYMENT), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Empleado:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "PaymentTypes", aPaymentComponent(N_PAYMENT_TYPE_ID_PAYMENT), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Tipo de pago:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Fecha valor:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aPaymentComponent(N_DATE_PAYMENT), -1, -1, -1) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Fecha de registro:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aPaymentComponent(N_REGISTERED_DATE_PAYMENT), -1, -1, -1) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Fecha del cheque:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateFromSerialNumber(aPaymentComponent(N_CHECK_DATE_PAYMENT), -1, -1, -1) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "BankAccounts", aPaymentComponent(N_ACCOUNT_ID_PAYMENT), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Cuenta del beneficiario:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "BankAccounts", aPaymentComponent(N_FROM_ACCOUNT_ID_PAYMENT), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Cuenta girada:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "CurrenciesSymbols", aPaymentComponent(N_CHECK_CURRENCY_ID_PAYMENT), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Monto:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & FormatNumber(aPaymentComponent(D_CHECK_AMOUNT_PAYMENT), 2, True, False, True) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Call GetNameFromTable(oADODBConnection, "StatusPayments", aPaymentComponent(N_STATUS_ID_PAYMENT), "", "", sNames, sErrorDescription)
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Estatus:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(sNames) & "</FONT></TD>"
				Response.Write "</TR>"
				If Len(aPaymentComponent(S_REPLACEMENT_NUMBER_PAYMENT)) > 0 Then
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Este cheque fue reemplazado con:&nbsp;</B></FONT></TD>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & CleanStringForHTML(aPaymentComponent(S_REPLACEMENT_NUMBER_PAYMENT)) & "</FONT></TD>"
					Response.Write "</TR>"
				End If
				If Len(aPaymentComponent(S_DESCRIPTION_PAYMENT)) = 0 Then
					Response.Write "<TR><TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2""><B>Descripción:</B><BR />" & CleanStringForHTML(aPaymentComponent(S_DESCRIPTION_PAYMENT)) & "</FONT></TD></TR>"
				End If
			Response.Write "</TABLE>"
		Response.Write "</FONT></DIV>"
	End If

	DisplayPayment = lErrorNumber
	Err.Clear
End Function

Function DisplayPaymentForm(oRequest, oADODBConnection, sAction, aPaymentComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an payment from the
'		  database using a HTML Form
'Inputs:  oRequest, oADODBConnection, sAction, aPaymentComponent
'Outputs: aPaymentComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPaymentForm"
	Dim lErrorNumber

	If aPaymentComponent(N_ID_PAYMENT) <> -1 Then
		lErrorNumber = GetPayment(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
	End If
	If lErrorNumber = 0 Then
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "function CheckPaymentFields(oForm) {" & vbNewLine
				If Len(oRequest("Delete").Item) = 0 Then
					Response.Write "if (oForm) {" & vbNewLine
						Response.Write "if (oForm.CheckNumber.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir el número de cheque.');" & vbNewLine
							Response.Write "oForm.CheckNumber.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "if (oForm.EmployeeID.value.length == 0) {" & vbNewLine
							Response.Write "alert('Favor de introducir el número del empleado.');" & vbNewLine
							Response.Write "oForm.EmployeeID.focus();" & vbNewLine
							Response.Write "return false;" & vbNewLine
						Response.Write "}" & vbNewLine
						Response.Write "oForm.CheckAmount.value = oForm.CheckAmount.value.replace(/" & NUMERIC_SEPARATOR & "/gi, '');" & vbNewLine
						Response.Write "if (!CheckFloatValue(oForm.CheckAmount, 'el monto del cheque', N_MINIMUM_ONLY_FLAG, N_OPEN_FLAG, 0, 0))" & vbNewLine
							Response.Write "return false;" & vbNewLine
					Response.Write "}" & vbNewLine
				End If
				Response.Write "return true;" & vbNewLine
			Response.Write "} // End of CheckPaymentFields" & vbNewLine

			Response.Write "function ShowHideReplacement(sStatus) {" & vbNewLine
				Response.Write "if (sStatus == '2')" & vbNewLine
					Response.Write "ShowDisplay(document.all['ReplacementNumberDiv']);" & vbNewLine
				Response.Write "else" & vbNewLine
					Response.Write "HideDisplay(document.all['ReplacementNumberDiv']);" & vbNewLine
			Response.Write "} // End of ShowHideReplacement" & vbNewLine

		Response.Write "//--></SCRIPT>" & vbNewLine
		Response.Write "<FORM NAME=""PaymentFrm"" ID=""PaymentFrm"" ACTION=""" & sAction & """ METHOD=""POST"" onSubmit=""return CheckPaymentFields(this)"">"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Action"" ID=""ActionHdn"" VALUE=""Payments"" />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaymentID"" ID=""PaymentIDHdn"" VALUE=""" & aPaymentComponent(N_ID_PAYMENT) & """ />"
			Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PreviousReplacementNumber"" ID=""PreviousReplacementNumberHdn"" VALUE=""" & aPaymentComponent(S_REPLACEMENT_NUMBER_PAYMENT) & """ />"

			Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del cheque:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""CheckNumber"" ID=""CheckNumberTxt"" VALUE=""" & aPaymentComponent(S_CHECK_NUMBER_PAYMENT) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del empleado:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" VALUE=""" & aPaymentComponent(N_EMPOYEE_ID_PAYMENT) & """ SIZE=""10"" MAXLENGTH=""10"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Tipo de pago:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""PaymentTypeID"" ID=""PaymentTypeIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PaymentTypes", "PaymentTypeID", "PaymentTypeName", "(Active=1)", "PaymentTypeName", aPaymentComponent(N_PAYMENT_TYPE_ID_PAYMENT), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha valor:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aPaymentComponent(N_DATE_PAYMENT), "Payment", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de registro:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aPaymentComponent(N_REGISTERED_DATE_PAYMENT), "Registered", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha del cheque:&nbsp;</FONT></TD>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">" & DisplayDateCombosUsingSerial(aPaymentComponent(N_CHECK_DATE_PAYMENT), "Check", N_FORM_START_YEAR, Year(Date()), True, True) & "</FONT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cuenta del beneficiario:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""AccountID"" ID=""AccountIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BankAccounts", "AccountID", "AccountNumber", "(Active=1)", "AccountNumber", aPaymentComponent(N_ACCOUNT_ID_PAYMENT), "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Cuenta girada:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""FromAccountID"" ID=""FromAccountIDCmb"" SIZE=""1"" CLASS=""Lists"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BankAccounts", "AccountID", "AccountNumber", "(Active=1)", "AccountNumber", aPaymentComponent(N_FROM_ACCOUNT_ID_PAYMENT), "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Monto:&nbsp;</FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""CheckAmount"" ID=""CheckAmountTxt"" VALUE=""" & FormatNumber(aPaymentComponent(D_CHECK_AMOUNT_PAYMENT), 2, True, False, True) & """ SIZE=""20"" MAXLENGTH=""20"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				If B_ISSSTE Then
					Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CurrencyID"" ID=""CurrencyIDHdn"" VALUE=""0"" />"
				Else
					Response.Write "<TR>"
						Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Moneda:&nbsp;</FONT></TD>"
						Response.Write "<TD><SELECT NAME=""CurrencyID"" ID=""CurrencyIDCmb"" SIZE=""1"" CLASS=""Lists"">"
							Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Currencies", "CurrencyID", "CurrencyName", "(Active=1)", "CurrencyName", aPaymentComponent(N_CHECK_CURRENCY_ID_PAYMENT), "Ninguna;;;-1", sErrorDescription)
						Response.Write "</SELECT></TD>"
					Response.Write "</TR>"
				End If
				Response.Write "<TR>"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Estatus:&nbsp;</FONT></TD>"
					Response.Write "<TD><SELECT NAME=""StatusID"" ID=""StatusIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""ShowHideReplacement(this.value)"">"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusPayments", "StatusID", "StatusName", "(Active=1)", "StatusName", aPaymentComponent(N_STATUS_ID_PAYMENT), "Ninguno;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
				Response.Write "<TR NAME=""ReplacementNumberDiv"" ID=""ReplacementNumberDiv"""
					If aPaymentComponent(N_STATUS_ID_PAYMENT) <> 2 Then Response.Write " STYLE=""display: none"""
				Response.Write ">"
					Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>Reemplazado con:&nbsp;</B></FONT></TD>"
					Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""ReplacementNumber"" ID=""ReplacementNumberTxt"" VALUE=""" & aPaymentComponent(S_REPLACEMENT_NUMBER_PAYMENT) & """ SIZE=""30"" MAXLENGTH=""100"" CLASS=""TextFields"" /></TD>"
				Response.Write "</TR>"
				Response.Write "<TR><TD COLSPAN=""2"">"
						Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Descripción:<BR /></FONT>"
					Response.Write "<TEXTAREA NAME=""Description"" ID=""DescriptionTxtArea"" ROWS=""5"" COLS=""60"" MAXLENGTH=""2000"" CLASS=""TextFields"">" & aPaymentComponent(S_DESCRIPTION_PAYMENT) & "</TEXTAREA>"
				Response.Write "</TD></TR>"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""bIsPayment"" ID=""bIsPaymentHdn"" VALUE=""1"" />"
				Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""bIsUpdated"" ID=""bIsUpdatedHdn"" VALUE=""1"" />"
			Response.Write "</TABLE>"

			Response.Write "<BR /><IMG SRC=""Images/DotBlue.gif"" WIDTH=""500"" HEIGHT=""1"" /><BR /><BR />"

			If aPaymentComponent(N_ID_PAYMENT) = -1 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_ADD_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Add"" ID=""AddBtn"" VALUE=""Agregar"" CLASS=""Buttons"" />"
			ElseIf Len(oRequest("Delete").Item) > 0 Then
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS Then Response.Write "<INPUT TYPE=""BUTTON"" NAME=""RemoveWng"" NAME=""RemoveWngBtn"" VALUE=""Eliminar"" CLASS=""RedButtons"" onClick=""ShowDisplay(document.all['RemovePaymentWngDiv']); PaymentFrm.Remove.focus()"" />"
			Else
				If aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS Then Response.Write "<INPUT TYPE=""SUBMIT"" NAME=""Modify"" ID=""ModifyBtn"" VALUE=""Modificar"" CLASS=""Buttons"" />"
			End If
			Response.Write "<IMG SRC=""Images/Transparent.gif"" WIDTH=""100"" HEIGHT=""1"" />"
			Response.Write "<INPUT TYPE=""BUTTON"" NAME=""Cancel"" ID=""CancelBtn"" VALUE=""Cancelar"" CLASS=""Buttons"" onClick=""window.location.href='" & GetASPFileName("") & "?Action=Users'"" />"
			Response.Write "<BR /><BR />"
			Call DisplayWarningDiv("RemovePaymentWngDiv", "¿Está seguro que desea borrar el registro de la base de datos?")
		Response.Write "</FORM>"
	End If

	DisplayPaymentForm = lErrorNumber
	Err.Clear
End Function

Function DisplayPaymentAsHiddenFields(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about an payment using
'		  hidden form fields
'Inputs:  oRequest, oADODBConnection, aPaymentComponent
'Outputs: aPaymentComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPaymentAsHiddenFields"

	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaymentID"" ID=""PaymentIDHdn"" VALUE=""" & aPaymentComponent(N_ID_PAYMENT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CheckNumber"" ID=""CheckNumberHdn"" VALUE=""" & aPaymentComponent(S_CHECK_NUMBER_PAYMENT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""ReplacementNumber"" ID=""ReplacementNumberHdn"" VALUE=""" & aPaymentComponent(S_REPLACEMENT_NUMBER_PAYMENT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""EmployeeID"" ID=""EmployeeIDHdn"" VALUE=""" & aPaymentComponent(N_EMPOYEE_ID_PAYMENT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaymentTypeID"" ID=""PaymentTypeIDHdn"" VALUE=""" & aPaymentComponent(N_PAYMENT_TYPE_ID_PAYMENT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""PaymentDate"" ID=""PaymentDateHdn"" VALUE=""" & aPaymentComponent(N_DATE_PAYMENT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""RegisteredDate"" ID=""RegisteredDateHdn"" VALUE=""" & aPaymentComponent(N_REGISTERED_DATE_PAYMENT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CheckDate"" ID=""CheckDateHdn"" VALUE=""" & aPaymentComponent(N_CHECK_DATE_PAYMENT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""AccountID"" ID=""AccountIDHdn"" VALUE=""" & aPaymentComponent(N_ACCOUNT_ID_PAYMENT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""FromAccountID"" ID=""FromAccountIDHdn"" VALUE=""" & aPaymentComponent(N_FROM_ACCOUNT_ID_PAYMENT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CheckAmount"" ID=""CheckAmountHdn"" VALUE=""" & aPaymentComponent(D_CHECK_AMOUNT_PAYMENT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""CheckCurrencyID"" ID=""CheckCurrencyIDHdn"" VALUE=""" & aPaymentComponent(N_CHECK_CURRENCY_ID_PAYMENT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""StatusID"" ID=""StatusIDHdn"" VALUE=""" & aPaymentComponent(N_STATUS_ID_PAYMENT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""Description"" ID=""DescriptionHdn"" VALUE=""" & aPaymentComponent(S_DESCRIPTION_PAYMENT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""bIsPayment"" ID=""bIsPaymentHdn"" VALUE=""" & aPaymentComponent(B_IS_PAYMENT) & """ />"
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""bIsUpdated"" ID=""bIsUpdatedHdn"" VALUE=""" & aPaymentComponent(B_IS_UPDATED_PAYMENT) & """ />"

	DisplayPaymentAsHiddenFields = Err.number
	Err.Clear
End Function

Function DisplayPaymentsTable(oRequest, oADODBConnection, lIDColumn, bUseLinks, bForExport, aPaymentComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about all the payments from
'		  the database in a table
'Inputs:  oRequest, oADODBConnection, lIDColumn, bUseLinks, bForExport, aPaymentComponent
'Outputs: aPaymentComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPaymentsTable"
	Dim sRequest
	Dim sClosed
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
	Dim sAction
	Dim lErrorNumber

	sClosed = ","
	sErrorDescription = "No se pudieron obtener los registros de la base de datos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PayrollID From Payrolls Where (PayrollTypeID=0) And (IsClosed=1) Order By PayrollID", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			sClosed = sClosed & CStr(oRecordset.Fields("PayrollID").Value) & ","
			oRecordset.MoveNext
		Loop
		oRecordset.Close
	End If

	lErrorNumber = GetPayments(oRequest, oADODBConnection, aPaymentComponent, oRecordset, sErrorDescription)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetch(oRequest, CInt(oRequest("StartPage").Item), ROWS_REPORT, oRecordset)
			Response.Write "<DIV NAME=""ReportDiv"" ID=""ReportDiv""><TABLE BORDER="""
				If bForExport Then
					Response.Write "1"
				Else
					Response.Write "0"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				sRequest = RemoveParameterFromURLString(RemoveEmptyParametersFromURLString(oRequest), "ReinsuranceDisaster")
				If bUseLinks And Not bForExport And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
					sRowContents = "Acciones"
					asCellWidths = asCellWidths & "80"
				Else
					sRowContents = ""
					asCellWidths = asCellWidths & "20"
				End If
				asCellAlignments = asCellAlignments & "CENTER"

				sRowContents = sRowContents & TABLE_SEPARATOR & "No. Cheque"
				asCellWidths = asCellWidths & ",100"
				asCellAlignments = asCellAlignments & ","
				If False And Not bForExport Then
					sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
					sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?"
						If StrComp(aPaymentComponent(S_SORT_COLUMN_PAYMENT), "CheckNumber", vbBinaryCompare) = 0 Then
							If aPaymentComponent(B_SORT_DESCENDING_PAYMENT) Then
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "CheckNumber"), "Desc", "0") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedDesc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
							Else
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "CheckNumber"), "Desc", "1") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar descendentemente"" BORDER=""0"" />"
							End If
						Else
							sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "CheckNumber"), "Desc", "0") & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
						End If
					sRowContents = sRowContents & "</A>"
				End If
				If (Len(oRequest("Action").Item) > 0) And (StrComp(oRequest("Action").Item,"EmployeesMovements",vbBinaryCompare) <> 0) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & "Empleado"
					asCellWidths = asCellWidths & ",250"
					asCellAlignments = asCellAlignments & ","
				
					If False And Not bForExport Then
						sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
						sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?"
							If StrComp(aPaymentComponent(S_SORT_COLUMN_PAYMENT), "EmployeeLastName, EmployeeLastName2, EmployeeName", vbBinaryCompare) = 0 Then
								If aPaymentComponent(B_SORT_DESCENDING_PAYMENT) Then
									sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "EmployeeLastName, EmployeeLastName2, EmployeeName"), "Desc", "0") & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedDesc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
								Else
									sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "EmployeeLastName, EmployeeLastName2, EmployeeName"), "Desc", "1") & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar descendentemente"" BORDER=""0"" />"
								End If
							Else
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "EmployeeLastName, EmployeeLastName2, EmployeeName"), "Desc", "0") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
							End If
						sRowContents = sRowContents & "</A>"
					End If
				End If
				If (Len(oRequest("Action").Item) > 0) And (StrComp(oRequest("Action").Item,"EmployeesMovements",vbBinaryCompare) = 0) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & "Quincena de pago"
					
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & "Fecha del cheque"
				End If
				asCellWidths = asCellWidths & ",200"
				asCellAlignments = asCellAlignments & ","
				If False And Not bForExport Then
					sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
					sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?"
						If StrComp(aPaymentComponent(S_SORT_COLUMN_PAYMENT), "CheckDate", vbBinaryCompare) = 0 Then
							If aPaymentComponent(B_SORT_DESCENDING_PAYMENT) Then
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "CheckDate"), "Desc", "0") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedDesc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
							Else
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "CheckDate"), "Desc", "1") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar descendentemente"" BORDER=""0"" />"
							End If
						Else
							sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "CheckDate"), "Desc", "0") & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
						End If
					sRowContents = sRowContents & "</A>"
				End If
				If (Len(oRequest("Action").Item) > 0) And (StrComp(oRequest("Action").Item,"EmployeesMovements",vbBinaryCompare) = 0) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & "Monto"
					asCellWidths = asCellWidths & ",100"
					asCellAlignments = asCellAlignments & ",RIGHT"
					If False And Not bForExport Then
						sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
						sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?"
							If StrComp(aPaymentComponent(S_SORT_COLUMN_PAYMENT), "CheckAmount", vbBinaryCompare) = 0 Then
								If aPaymentComponent(B_SORT_DESCENDING_PAYMENT) Then
									sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "CheckAmount"), "Desc", "0") & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedDesc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
								Else
									sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "CheckAmount"), "Desc", "1") & """>"
									sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar descendentemente"" BORDER=""0"" />"
								End If
							Else
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "CheckAmount"), "Desc", "0") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
							End If
						sRowContents = sRowContents & "</A>"
					End If
				End If
				If (Len(oRequest("Action").Item) > 0) And (StrComp(oRequest("Action").Item,"EmployeesMovements",vbBinaryCompare) = 0) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & "Motivo de Cancelación"
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & "Estatus"
				End If
				
				asCellWidths = asCellWidths & ",200"
				asCellAlignments = asCellAlignments & ","
				If False And Not bForExport Then
					sRowContents = sRowContents & "&nbsp;&nbsp;&nbsp;"
					sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?"
						If StrComp(aPaymentComponent(S_SORT_COLUMN_PAYMENT), "StatusName", vbBinaryCompare) = 0 Then
							If aPaymentComponent(B_SORT_DESCENDING_PAYMENT) Then
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "StatusName"), "Desc", "0") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedDesc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
							Else
								sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "StatusName"), "Desc", "1") & """>"
								sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortedAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar descendentemente"" BORDER=""0"" />"
							End If
						Else
							sRowContents = sRowContents & ReplaceValueInURLString(ReplaceValueInURLString(sRequest, "SortColumn", "StatusName"), "Desc", "0") & """>"
							sRowContents = sRowContents & "<IMG SRC=""Images/ArrSortAsc.gif"" WIDTH=""8"" HEIGHT=""8"" ALT=""Ordenar ascendentemente"" BORDER=""0"" />"
						End If
					sRowContents = sRowContents & "</A>"
				End If
				If (Len(oRequest("Action").Item) > 0) And (StrComp(oRequest("Action").Item,"EmployeesMovements",vbBinaryCompare) = 0) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & "Fecha de Cancelación"
				Else
					sRowContents = sRowContents & TABLE_SEPARATOR & "Quincena de aplicación"
				End If
				asCellWidths = asCellWidths & ",100"
				asCellAlignments = asCellAlignments & ","
				
				If (Len(oRequest("Action").Item) > 0) And (StrComp(oRequest("Action").Item,"EmployeesMovements",vbBinaryCompare) <> 0) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & "¿Aplicado?"
					asCellWidths = asCellWidths & ",100"
					asCellAlignments = asCellAlignments & ",CENTER"
				End If
				asColumnsTitles = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				asCellWidths = Split(asCellWidths, ",", -1, vbBinaryCompare)
				asCellAlignments = Split(asCellAlignments, ",", -1, vbBinaryCompare)

				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				sAction = "ShowInfo"
				If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then sAction = "Change"
				iRecordCounter = 0
				Do While Not oRecordset.EOF
					sBoldBegin = ""
					sBoldEnd = ""
					If StrComp(CStr(oRecordset.Fields("PaymentID").Value), oRequest("PaymentID").Item, vbBinaryCompare) = 0 Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sFontBegin = ""
					sFontEnd = ""
					sRowContents = ""
					Select Case lIDColumn
						Case DISPLAY_RADIO_BUTTONS
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""PaymentID"" ID=""PaymentIDRd"" VALUE=""" & CStr(oRecordset.Fields("PaymentID").Value) & """ />"
						Case DISPLAY_CHECKBOXES
							sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""PaymentID"" ID=""PaymentIDChk"" VALUE=""" & CStr(oRecordset.Fields("PaymentID").Value) & """ />"
						Case Else
							If bUseLinks And Not bForExport And (((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS) Or ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS)) Then
								sRowContents = sRowContents & "&nbsp;"
									If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
										sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Payments&PaymentID=" & CStr(oRecordset.Fields("PaymentID").Value) & "&Tab=1&Change=1"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnModify.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Modificar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
									End If

									If B_DELETE And (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS Then
										sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Payments&PaymentID=" & CStr(oRecordset.Fields("PaymentID").Value) & "&Tab=1&Delete=1"">"
											sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Borrar"" BORDER=""0"" />"
										sRowContents = sRowContents & "</A>&nbsp;&nbsp;&nbsp;"
									End If

									If (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Then
										If CInt(oRecordset.Fields("Active").Value) = 0 Then 
											sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Payments&PaymentID=" & CStr(oRecordset.Fields("PaymentID").Value) & "&Tab=1&SetActive=1""><IMG SRC=""Images/BtnActive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Activar cheque"" BORDER=""0"" /></A>"
										Else
											sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=Payments&PaymentID=" & CStr(oRecordset.Fields("PaymentID").Value) & "&Tab=1&SetActive=0""><IMG SRC=""Images/BtnDeactive.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Desactivar cheque"" BORDER=""0"" /></A>"
										End If
									End If
								sRowContents = sRowContents & "&nbsp;"
							End If
					End Select
					If (Len(oRequest("Action").Item) > 0) And (StrComp(oRequest("Action").Item,"EmployeesMovements",vbBinaryCompare) <> 0) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & "<A"
						If Not bForExport Then sRowContents = sRowContents & " HREF=""Payments.asp?PaymentID=" & CStr(oRecordset.Fields("PaymentID").Value) & "&Tab=1&" & sAction & "=1"""
						sRowContents = sRowContents & ">" & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("CheckNumber").Value)) & "</A>" & sBoldEnd & sFontEnd
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & "<A"
							If Not bForExport Then sRowContents = sRowContents & " HREF=""Employees.asp?EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&Change=1&Tab=1"""
						If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
							sRowContents = sRowContents & ">" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value)) & sBoldEnd & sFontEnd & "</A>"
						Else
							sRowContents = sRowContents & ">" & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeLastName").Value) & ", " & CStr(oRecordset.Fields("EmployeeName").Value)) & sBoldEnd & sFontEnd & "</A>"
						End If
					Else
						sRowContents = sRowContents	& TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("CheckNumber").Value)) & sBoldEnd & sFontEnd
					End If
					If (Len(oRequest("Action").Item) > 0) And (StrComp(oRequest("Action").Item,"EmployeesMovements",vbBinaryCompare) = 0) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PaymentDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("CheckDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
					End If
					If (Len(oRequest("Action").Item) > 0) And (StrComp(oRequest("Action").Item,"EmployeesMovements",vbBinaryCompare) = 0) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & FormatNumber(CDbl(oRecordset.Fields("CheckAmount").Value), 2, True, False, True) & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("CurrencySymbol").Value)) & FormatNumber(CDbl(oRecordset.Fields("CheckAmount").Value), 2, True, False, True) & sBoldEnd & sFontEnd
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
						If CLng(oRecordset.Fields("CancelDate").Value) = 0 Then
							sRowContents = sRowContents & "<CENTER>---</CENTER>"
						Else
							sRowContents = sRowContents & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("CancelDate").Value), -1, -1 ,-1)
						End If
					sRowContents = sRowContents & sBoldEnd & sFontEnd
				If (Len(oRequest("Action").Item) > 0) And (StrComp(oRequest("Action").Item,"EmployeesMovements",vbBinaryCompare) <> 0) Then
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin
						If CLng(oRecordset.Fields("CancelDate").Value) = 0 Then
							sRowContents = sRowContents & "<CENTER>---</CENTER>"
						ElseIf InStr(1, sClosed, "," & CStr(oRecordset.Fields("CancelDate").Value) & ",", vbBinaryCompare) > 0 Then
							sRowContents = sRowContents & "Sí"
						Else
							sRowContents = sRowContents & "No"
						End If
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
					If (Not bForExport) And (iRecordCounter >= ROWS_REPORT) Then Exit Do
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE></DIV>" & vbNewLine
		Else
			lErrorNumber = L_ERR_NO_RECORDS
			sErrorDescription = "No existen cheques registrados en la base de datos."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayPaymentsTable = lErrorNumber
	Err.Clear
End Function
%>