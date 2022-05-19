<%
Function GetEmployeesBlockPayments(oRequest, oADODBConnection, aPaymentComponent, oRecordset, sErrorDescription)
'************************************************************
'Purpose: To get the information about all the Payments blocked for
'         the employee from the database
'Inputs:  oRequest, oADODBConnection
'Outputs: aAbsenceComponent, oRecordset, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetEmployeesBlockPayments"
	Dim sTables
	Dim sCondition
	Dim lErrorNumber
	Dim bComponentInitialized
	Dim sQuery

	If Len(oRequest("PayrollID").Item) > 0 Then
		If CLng(oRequest("PayrollID").Item) > 0 Then sCondition = sCondition & " And (EmployeesBlockPaymentsLKP.PayrollID=" & CLng(oRequest("PayrollID").Item) & ")"
	End If
	If Len(sCondition ) > 0 Then
		If InStr(1, sCondition , "And ", vbBinaryCompare) = 0 Then sCondition  = "And " & sCondition
	End If

	sQuery = "Select Employees.EmployeeID, Employees.EmployeeName + ' ' + Employees.EmployeeLastName + ' ' + Employees.EmployeeLastName2 As EmployeeFullName, PayrollID From Employees, EmployeesBlockPaymentsLKP Where (Employees.EmployeeID=EmployeesBlockPaymentsLKP.EmployeeID)"
	sQuery = sQuery & sCondition & " Order By PayrollID Desc, EmployeeID Desc"
	sErrorDescription = "No se pudo obtener la información de los registros."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, sQuery, "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	Response.Write "<INPUT TYPE=""HIDDEN"" NAME=""sQuery"" ID=""sQueryHdn"" VALUE=""" & sQuery & """ />"

	GetEmployeesBlockPayments = lErrorNumber
	Err.Clear
End Function

Function GetPaymentsURLValues(oRequest, sAction, bAction, sCondition)
'************************************************************
'Purpose: To initialize the global variables using the URL
'Inputs:  oRequest
'Outputs: sAction, bAction, sCondition
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "GetPaymentsURLValues"
	Dim oItem
	Dim aItem

	sAction = oRequest("Action").Item
	bAction = (Len(oRequest("Add").Item) > 0) Or (Len(oRequest("Modify").Item) > 0) Or (Len(oRequest("Remove").Item) > 0) Or (Len(oRequest("SetActive").Item) > 0)

	sCondition = ""
	If Len(oRequest("CheckNumber").Item) > 0 Then
		sCondition = sCondition & " And (CheckNumber Like '" & S_WILD_CHAR & Replace(oRequest("CheckNumber").Item, ", ", ",") & S_WILD_CHAR & "')"
	End If
	If Len(oRequest("EmployeeID").Item) > 0 Then
		sCondition = sCondition & " And (Payments.EmployeeID In ('" & Replace(Replace(oRequest("EmployeeID").Item, ", ", ","), ",", "','") & "'))"
	End If
	If Len(oRequest("PaymentTypeID").Item) > 0 Then
		sCondition = sCondition & " And (Payments.PaymentTypeID In (" & oRequest("PaymentTypeID").Item & "))"
	End If
	If (InStr(1, oRequest, "StartPayment", vbTextCompare) > 0) Or (InStr(1, oRequest, "EndPayment", vbTextCompare) > 0) Then Call GetStartAndEndDatesFromURL("StartPayment", "EndPayment", "PaymentDate", False, sCondition)
	If (InStr(1, oRequest, "StartRegistered", vbTextCompare) > 0) Or (InStr(1, oRequest, "EndRegistered", vbTextCompare) > 0) Then Call GetStartAndEndDatesFromURL("StartRegistered", "EndRegistered", "RegisteredDate", False, sCondition)
	If (InStr(1, oRequest, "StartCheck", vbTextCompare) > 0) Or (InStr(1, oRequest, "EndCheck", vbTextCompare) > 0) Then Call GetStartAndEndDatesFromURL("StartCheck", "EndCheck", "CheckDate", False, sCondition)
	If Len(oRequest("AccountID").Item) > 0 Then
		sCondition = sCondition & " And (Payments.AccountID In (" & oRequest("AccountID").Item & "))"
	End If
	If Len(oRequest("FromAccountID").Item) > 0 Then
		sCondition = sCondition & " And (Payments.FromAccountID In (" & oRequest("FromAccountID").Item & "))"
	End If
	If Len(oRequest("CheckCurrencyID").Item) > 0 Then
		sCondition = sCondition & " And (Payments.CheckCurrencyID In (" & oRequest("CheckCurrencyID").Item & "))"
	End If
	If Len(oRequest("StatusID").Item) > 0 Then
		sCondition = sCondition & " And (Payments.StatusID In (" & oRequest("StatusID").Item & "))"
	End If

	If Len(oRequest("DoSearch").Item) > 0 Then Response.Cookies("SIAP_SearchPath").Item = oRequest

	GetPaymentsURLValues = Err.number
	Err.Clear
End Function

Function DoPaymentCatalogsAction(oRequest, oADODBConnection, aCatalogComponent, bAction, sAction, sCondition, sErrorDescription)
'************************************************************
'Purpose: To initialize and add, modify, or remove entires in
'         the EmployeesKardex or EmployeesKardex2 tables
'Inputs:  oRequest, oADODBConnection, aCatalogComponent
'Outputs: aCatalogComponent, bAction, sAction, sCondition, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DoPaymentCatalogsAction"
	Dim sTemp
	Dim lErrorNumber

	'bSearchForm = (Len(oRequest("Search").Item) > 0)
	'bShowForm = ((Len(oRequest("New").Item) > 0) Or (Len(oRequest("Change").Item) > 0) Or (Len(oRequest("Delete").Item) > 0))
	If Len(oRequest("RecordID2").Item) > 0 Then aCatalogComponent(S_TABLE_NAME_CATALOG) = "PaymentsRecords2"
	Call InitializeCatalogs(oRequest)
	Call InitializeValuesForCatalogComponent(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
	If Len(oRequest("RecordID2").Item) > 0 Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) = oRequest("RecordID2").Item
	aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)) = CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG)))

	aCatalogComponent(S_QUERY_CONDITION_CATALOG) = ""
	If Len(oRequest("Add").Item) > 0 Then
		aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(3) = Replace(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(3), " ", "")
		If Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(3)) = 0 Then
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(3) = "0"
		End If
		If Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(18)) = 0 Then aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(18) = -1
		lErrorNumber = AddCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			Select Case sAction
				Case "PaymentsRecords"
					sCondition = " And (EmployeesHistoryListForPayroll.CompanyID In (" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2) & ")) And (EmployeesHistoryListForPayroll.EmployeeTypeID In (" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(5) & ")) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ")"
					If StrComp(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(4), "-1", vbBinaryCompare) <> 0 Then ' Entidad|ZoneIDs = Todas|No indicada
						sCondition = sCondition & " And (Zones1.ZonePath Like '" & S_WILD_CHAR & "," & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(4) & "," & S_WILD_CHAR & "')"
					End If
					If StrComp(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(3), "0", vbBinaryCompare) <> 0 Then ' Centros de pago|AreasIDs = 00|No aplica
						sCondition = sCondition & " And (PaymentCenters.AreaID In (" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(3) & "))"
					End If
					If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
						sCondition = sCondition & " And (EmployeesHistoryListForPayroll.PaymentCenterID In (" & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & "))"
					End If
					If (Len(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(18)) > 0) And (StrComp(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(18), "-1", vbBinaryCompare) <> 0) Then sCondition = sCondition & " And (EmployeesHistoryListForPayroll.EmployeeID In (" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(18) & "))"
					lErrorNumber = AddPayments(oRequest, oADODBConnection, aCatalogComponent, sCondition, sErrorDescription)

					If lErrorNumber = 0 Then
						If FileExists(Server.MapPath("Reports/Rep_1400_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & ".zip"), sErrorDescription) Then
							lErrorNumber = DeleteFile(Server.MapPath("Reports/Rep_1400_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & ".zip"), sErrorDescription)
						End If
					End If
				Case "Replacement"
					lErrorNumber = ModifyPayments(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
				Case "Reexpedition"
					lErrorNumber = IssuePayment(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
				Case "PrintPayments", "PaymentsMessages"
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1," & oRequest("PayrollID").Item & ",,-1,-1,-1,-1,-1,-1,-1,0,"
					aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
					aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1," & oRequest("PayrollID").Item & ",,-1,-1,-1,-1,-1,-1,-1,0,"
					aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
					bAction = False
			End Select
		End If
		If InStr(1, ",PrintPayments,PaymentsMessages,", "," & sAction & ",", vbBinaryCompare) = 0 Then bAction = True
	ElseIf Len(oRequest("Modify").Item) > 0 Then
		lErrorNumber = ModifyCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
		bAction = True
	ElseIf Len(oRequest("ChangeStatus").Item) > 0 Then
		lErrorNumber = ModifyPaymentsStatus(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
		bAction = True
	ElseIf Len(oRequest("UpdateStatus").Item) > 0 Then
		sTemp = ""
		If Len(oRequest("RecordToClose").Item) > 0 Then
			sTemp = Replace(oRequest("RecordToClose").Item, " ", "")
			sErrorDescription = "No se pudo actualizar el estatus de los registros."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PaymentsRecords Set bPrinted=2 Where (RecordID In (" & sTemp & "))", "PaymentsLib.asp", "_root", 000, sErrorDescription, Null)
		End If
		sTemp = ""
		If Len(oRequest("RecordToClose2").Item) > 0 Then
			sTemp = Replace(oRequest("RecordToClose2").Item, " ", "")
			sErrorDescription = "No se pudo actualizar el estatus de los registros."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update PaymentsRecords2 Set bPrinted=2 Where (RecordID In (" & sTemp & "))", "PaymentsLib.asp", "_root", 000, sErrorDescription, Null)
		End If
		bAction = True
	ElseIf Len(oRequest("BlockEmployees").Item) > 0 Then
		If StrComp(oRequest("RemoveBlockPayments").Item, 1, vbBinaryCompare) = 0 Then
			aPaymentComponent(N_EMPOYEE_ID_PAYMENT) = CLng(oRequest("EmployeeID").Item)
			lErrorNumber = RemoveEmployeeBlockPaymentsLKP(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
		Else
			'lErrorNumber = SetActiveForEmployeeAccount(oRequest, oADODBConnection, 0, sErrorDescription)
			aPaymentComponent(N_STATUS_ID_PAYMENT) = 4
			lErrorNumber = AddEmployeeBlockPaymentsLKP(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
		End If
		bAction = True
	ElseIf Len(oRequest("UnblockEmployees").Item) > 0 Then
		lErrorNumber = SetActiveForEmployeeAccount(oRequest, oADODBConnection, 1, sErrorDescription)
		bAction = True
	ElseIf Len(oRequest("UnblockEmployees").Item) > 0 Then
		lErrorNumber = SetActiveForEmployeeAccount(oRequest, oADODBConnection, 1, sErrorDescription)
		bAction = True
	ElseIf Len(oRequest("Remove").Item) > 0 Then
		lErrorNumber = GetCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
		If lErrorNumber = 0 Then
			Select Case sAction
				Case "PaymentsRecords", "Replacement", "RemovePaymentsRecords"
					If Len(oRequest("RecordID2").Item) = 0 Then
						sCondition = " (PaymentID>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12) & ") And (PaymentID<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13) & ")"
					Else
						sCondition = " (PaymentID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10) & ")"
					End If
					lErrorNumber = RemovePayments(oRequest, oADODBConnection, sCondition, sErrorDescription)
				Case "Reexpedition"
					sCondition = " (PaymentID>=xxx) And (PaymentID<=xxx)"
					lErrorNumber = RemovePayments(oRequest, oADODBConnection, sCondition, sErrorDescription)
			End Select
			If lErrorNumber = 0 Then
				lErrorNumber = RemoveCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) = -1
			End If
		End If
		If InStr(1, ",PrintPayments,PaymentsMessages,", "," & sAction & ",", vbBinaryCompare) = 0 Then
			bAction = True
		Else
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = "-1," & oRequest("PayrollID").Item & ",,-1,-1,-1,-1,-1,-1,0,"
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG) = Split(aCatalogComponent(AS_FIELDS_VALUES_CATALOG), ",")
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = "-1," & oRequest("PayrollID").Item & ",,-1,-1,-1,-1,-1,-1,0,"
			aCatalogComponent(AS_DEFAULT_VALUES_CATALOG) = Split(aCatalogComponent(AS_DEFAULT_VALUES_CATALOG), ",")
			bAction = False
		End If
	End If
	If (lErrorNumber = 0) And (Len(oRequest("ChangeStatus").Item) = 0) And (Len(oRequest("BlockEmployees").Item) = 0) And (Len(oRequest("UnblockEmployees").Item) = 0) Then
		If IsArray(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)) Then
			If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(aCatalogComponent(N_ID_CATALOG))) > -1 Then
				lErrorNumber = GetCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
			End If
		End If
	End If
	If Len(oRequest("DoSearch").Item) > 0 Then
		aCatalogComponent(S_QUERY_CONDITION_CATALOG) = ""
	End If
	aCatalogComponent(N_ACTIVE_CATALOG) = -1
	aCatalogComponent(S_URL_CATALOG) = "SectionID=" & oRequest("SectionID").Item

	DoPaymentCatalogsAction = lErrorNumber
	Err.Clear
End Function

Function DoPaymentsAction(oRequest, oADODBConnection, sAction, sErrorDescription)
'************************************************************
'Purpose: To add, change or delete the information of the
'         specified component
'Inputs:  oRequest, oADODBConnection, sAction
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DoPaymentsAction"
	Dim lErrorNumber

	If Len(oRequest("RemoveFile").Item) > 0 Then
		If FileExists(oRequest("FolderName").Item & "\" & oRequest("FileName").Item, sErrorDescription) Then
			lErrorNumber = DeleteFile(oRequest("FolderName").Item & "\" & oRequest("FileName").Item, sErrorDescription)
		End If
	ElseIf Len(oRequest("Add").Item) > 0 Then
		Select Case sAction
			Case "Payments"
				lErrorNumber = AddPayment(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
				If lErrorNumber <> 0 Then aPaymentComponent(N_ID_PAYMENT) = -1
		End Select
	ElseIf Len(oRequest("Modify").Item) > 0 Then
		Select Case sAction
			Case "Payments"
				lErrorNumber = ModifyPayment(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
		End Select
	ElseIf Len(oRequest("Remove").Item) > 0 Then
		Select Case sAction
			Case "Payments"
				If aPaymentComponent(N_ID_PAYMENT) > -1 Then
					lErrorNumber = GetPayment(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
					If lErrorNumber = 0 Then lErrorNumber = RemovePayment(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
				End If
		End Select
	End If

	DoPaymentsAction = lErrorNumber
	Err.Clear
End Function

Function AddPayments(oRequest, oADODBConnection, aCatalogComponent, sCondition, sErrorDescription)
'************************************************************
'Purpose: To add the employees' check payments
'Inputs:  oRequest, oADODBConnection, aCatalogComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddPayments"
	Dim oRecordset
	Dim sFileName
	Dim sContents
	Dim lPaymentID
	Dim lCheckNumber
	Dim iIndex
	Dim iStatus
	Dim sEmployeeIDs
	Dim aTemp
	Dim lRecordID
	Dim lErrorNumber

	lRecordID = aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)
	Application.Contents("SIAP_AddPayments") = Application.Contents("SIAP_AddPayments") & "," & lRecordID
	iIndex = 0
	aTemp = Split(Application.Contents("SIAP_AddPayments"), ",")
	Do While InStr(1, Application.Contents("SIAP_AddPayments"), ("," & lRecordID), vbBinaryCompare) <> 1
		iIndex = iIndex + 1
		If iIndex > 1000000 Then
			iIndex = 0
			sErrorDescription = "No se pudieron obtener los registros de la base de datos."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From PaymentsRecords Where (RecordID=" & aTemp(1) & ") And (FirstPaymentID=-1) And (LastPaymentID=-1)", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If oRecordset.EOF Then
					Application.Contents("SIAP_AddPayments") = Replace(Application.Contents("SIAP_AddPayments"), ("," & aTemp(1)), "")
				End If
				oRecordset.Close
			End If
		End If
	Loop

	sFileName = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Rep_471_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & Left(GetSerialNumberForDate(""), Len("00000000")) & ".txt")
	sErrorDescription = "No se pudo obtener el listado de empleados y de sus cheques."
	iStatus = -2
	Select Case CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14))
		Case 0 'Cheque
			sCondition = sCondition & " And (EmployeesHistoryListForPayroll.AccountNumber='.') And (EmployeesHistoryListForPayroll.BankID=" & oRequest("BankID").Item & ") And (ConceptID=0)"
		Case 1 'Depósito
			sCondition = sCondition & " And (EmployeesHistoryListForPayroll.AccountNumber<>'.') And (EmployeesHistoryListForPayroll.BankID=" & oRequest("BankID").Item & ") And (ConceptID=0)"
		Case 2 'Pensión alimenticia
			'sCondition = sCondition & " And (EmployeesHistoryListForPayroll.BankID=" & oRequest("BankID").Item & ") And (ConceptID=124)"
			sCondition = sCondition & " And (ConceptID=124)"
		Case 3 'Honorarios
			sCondition = sCondition & " And (EmployeesHistoryListForPayroll.EmployeeID>=600000) And (EmployeesHistoryListForPayroll.EmployeeID<700000) And (ConceptID=0)"
		Case 4 'Acreedores
			sCondition = sCondition & " And (ConceptID=155)"
	End Select
	Select Case CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14))
		Case 2 'Pensión alimenticia
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, BankAccounts.AccountID, EmployeesHistoryListForPayroll.EmployeeNumber, Sum(Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones As Zones1, Zones As Zones2, Zones, EmployeeTypes, BankAccounts Where (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.CompanyID=Companies.CompanyID) And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Companies.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesBeneficiariesLKP.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesBeneficiariesLKP.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " And (EmployeesBeneficiariesLKP.BeneficiaryNumber Not In (Select EmployeeID From Payments Where (PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & "))) Group By CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeesBeneficiariesLKP.BeneficiaryNumber, BankAccounts.AccountID, EmployeesHistoryListForPayroll.EmployeeNumber Order by CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryListForPayroll.EmployeeNumber", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, BankAccounts.AccountID, EmployeesHistoryListForPayroll.EmployeeNumber, Sum(Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptAmount) As TotalAmount From EmployeesBeneficiariesLKP, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones As Zones1, Zones As Zones2, Zones, EmployeeTypes, BankAccounts Where (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.CompanyID=Companies.CompanyID) And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (Companies.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesBeneficiariesLKP.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesBeneficiariesLKP.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " And (EmployeesBeneficiariesLKP.BeneficiaryNumber Not In (Select EmployeeID From Payments Where (PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & "))) Group By CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeesBeneficiariesLKP.BeneficiaryNumber, BankAccounts.AccountID, EmployeesHistoryListForPayroll.EmployeeNumber Order by CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryListForPayroll.EmployeeNumber -->" & vbNewLine
		Case 4 'Acreedores
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesCreditorsLKP.CreditorNumber As EmployeeID, -1 AccountID, EmployeesHistoryListForPayroll.EmployeeNumber, Sum(Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones As Zones1, Zones As Zones2, Zones, EmployeeTypes Where (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.CompanyID=Companies.CompanyID) And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (Companies.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesCreditorsLKP.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesCreditorsLKP.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " And (EmployeesCreditorsLKP.CreditorNumber Not In (Select EmployeeID From Payments Where (PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & "))) Group By CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryListForPayroll.EmployeeNumber Order by CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryListForPayroll.EmployeeNumber", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select EmployeesCreditorsLKP.CreditorNumber As EmployeeID, -1 AccountID, EmployeesHistoryListForPayroll.EmployeeNumber, Sum(Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptAmount) As TotalAmount From EmployeesCreditorsLKP, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones As Zones1, Zones As Zones2, Zones, EmployeeTypes Where (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.CompanyID=Companies.CompanyID) And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (Companies.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesCreditorsLKP.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesCreditorsLKP.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " And (EmployeesCreditorsLKP.CreditorNumber Not In (Select EmployeeID From Payments Where (PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & "))) Group By CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryListForPayroll.EmployeeNumber Order by CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryListForPayroll.EmployeeNumber -->" & vbNewLine
		Case Else 'Cheque, Depósito, Honorarios
			'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.AccountNumber, SUM(ConceptAmount) As TotalAmount From Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones As Zones1, Zones As Zones2, Zones, EmployeeTypes Where (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (Companies.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " And (EmployeesHistoryListForPayroll.EmployeeID Not In (Select EmployeeID From Payments Where (PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & "))) Group by CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.AccountNumber Order by CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.AccountNumber", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.AccountNumber, SUM(ConceptAmount) As TotalAmount From Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones As Zones1, Zones As Zones2, Zones, EmployeeTypes Where (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (Companies.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " And (EmployeesHistoryListForPayroll.EmployeeID Not In (Select EmployeeID From Payments Where (PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & "))) Group by CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.AccountNumber Order by CompanyShortName, PaymentCenters.AreaCode, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.AccountNumber", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.AccountNumber, SUM(ConceptAmount) As TotalAmount From Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones As Zones1, Zones As Zones2, Zones, EmployeeTypes Where (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Zones.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (Companies.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " And (EmployeesHistoryListForPayroll.EmployeeID Not In (Select EmployeeID From Payments Where (PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & "))) Group by CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.AccountNumber Order by CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeesHistoryListForPayroll.EmployeeID, EmployeesHistoryListForPayroll.AccountNumber -->" & vbNewLine
	End Select
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Do While Not oRecordset.EOF
				Select Case CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14))
					Case 2, 4
						lErrorNumber = AppendTextToFile(sFileName, (CleanStringforHTML(CStr(oRecordset.Fields("EmployeeID").Value)) & LIST_SEPARATOR & CleanStringforHTML(CStr(oRecordset.Fields("AccountID").Value)) & LIST_SEPARATOR & CleanStringforHTML(CStr(oRecordset.Fields("TotalAmount").Value))), sErrorDescription)
					Case Else
						lErrorNumber = AppendTextToFile(sFileName, (CleanStringforHTML(CStr(oRecordset.Fields("EmployeeID").Value)) & LIST_SEPARATOR & CleanStringforHTML(CStr(oRecordset.Fields("AccountNumber").Value)) & LIST_SEPARATOR & CleanStringforHTML(CStr(oRecordset.Fields("TotalAmount").Value))), sErrorDescription)
				End Select
				oRecordset.MoveNext
				If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
			Loop
			oRecordset.Close
			sErrorDescription = "No se pudo obtener un identificador para el nuevo cheque."
			lErrorNumber = GetNewIDFromTable(oADODBConnection, "Payments", "PaymentID", "", 1, lPaymentID, sErrorDescription)
			lCheckNumber = aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10)
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12) = lPaymentID
			If lErrorNumber = 0 Then
				sContents = GetFileContents(sFileName, sErrorDescription)
				sContents = Split(sContents, vbNewLine)
				sEmployeeIDs = ","
				For iIndex = 0 To UBound(sContents)
					If Len(sContents(iIndex)) > 0 Then
						sContents(iIndex) = Split(sContents(iIndex), LIST_SEPARATOR)
						If InStr(1, sEmployeeIDs, "," & sContents(iIndex)(0) & ",", vbBinaryCompare) = 0 Then
							'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payments (PaymentID, CheckNumber, ReplacementNumber, EmployeeID, PaymentTypeID, PaymentDate, RegisteredDate, CheckDate, CancelDate, AccountID, FromAccountID, CheckAmount, CheckCurrencyID, StatusID, Description, bIsPayment, bIsUpdated, LastUpdate, UserID, ReplacementUserID) Values (" & lPaymentID & ", '" & lCheckNumber & "', ' ', " & sContents(iIndex)(0) & ", 1, " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", " & aPaymentComponent(N_DATE_PAYMENT) & ", " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", 0, '" & sContents(iIndex)(1) & "', " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(7) & ", " & sContents(iIndex)(2) & ", 0, " & iStatus & ", ' ', 1, 0, " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", -1)", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							Select Case CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14))
								Case 2, 4 'Pensión alimenticia, Acreedores
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payments (PaymentID, CheckNumber, ReplacementNumber, EmployeeID, PaymentTypeID, PaymentDate, RegisteredDate, CheckDate, CancelDate, AccountID, FromAccountID, CheckAmount, CheckCurrencyID, StatusID, Description, bIsPayment, bIsUpdated, LastUpdate, UserID, ReplacementUserID) Values (" & lPaymentID & ", '" & lCheckNumber & "', ' ', " & sContents(iIndex)(0) & ", 1, " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", " & aPaymentComponent(N_DATE_PAYMENT) & ", " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", 0, '" & sContents(iIndex)(1) & "', " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(7) & ", " & sContents(iIndex)(2) & ", 0, " & iStatus & ", ' ', 1, 0, " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", -1)", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								Case Else
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payments (PaymentID, CheckNumber, ReplacementNumber, EmployeeID, PaymentTypeID, PaymentDate, RegisteredDate, CheckDate, CancelDate, AccountID, FromAccountID, CheckAmount, CheckCurrencyID, StatusID, Description, bIsPayment, bIsUpdated, LastUpdate, UserID, ReplacementUserID) Values (" & lPaymentID & ", '" & lCheckNumber & "', ' ', " & sContents(iIndex)(0) & ", 1, " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", " & aPaymentComponent(N_DATE_PAYMENT) & ", " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", 0, '-1', " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(7) & ", " & sContents(iIndex)(2) & ", 0, " & iStatus & ", ' ', 1, 0, " & Left(GetSerialNumberForDate(""), Len("00000000")) & ", " & aLoginComponent(N_USER_ID_LOGIN) & ", -1)", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							End Select
							If lErrorNumber = 0 Then
								sEmployeeIDs = sEmployeeIDs & sContents(iIndex)(0) & ","
								lPaymentID = lPaymentID + 1
								lCheckNumber = lCheckNumber + 1
							End If
						End If
					End If
				Next
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11) = lCheckNumber - 1
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13) = lPaymentID - 1
				lErrorNumber = ModifyCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
				Call DeleteFile(sFileName, "")
				Response.Redirect "Payments.asp?Action=PaymentsRecords&RecordID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & "&DisplayResults=1"
			End If
		Else
			Call RemoveCatalog(oRequest, oADODBConnection, aCatalogComponent, "")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) = -1
			lErrorNumber = -1
			sErrorDescription = "No se realizó la asignación de folios por alguna de estas dos razones:<OL><LI>No existen empleados que cumplan con los criterios de la búsqueda.</LI><LI>A los empleados que sí cumplieron con los criterios de la búsqueda ya se les asignó su número de folio.</LI></OL>"
		End If
	Else
		Call RemoveCatalog(oRequest, oADODBConnection, aCatalogComponent, "")
		aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) = -1
	End If

	Application.Contents("SIAP_AddPayments") = Replace(Application.Contents("SIAP_AddPayments"), ("," & lRecordID), "")
	Set oRecordset = Nothing
	AddPayments = lErrorNumber
	Err.Clear
End Function

Function AddEmployeeBlockPaymentsLKP(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
'************************************************************
'Purpose: To add payment block to employee for specific payroll
'Inputs:  oRequest, oADODBConnection, iActive
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "AddEmployeeBlockPaymentsLKP"
	Dim oItem
	Dim iIndex
	Dim sEmployeeIDs
	Dim lPayrollID
	Dim oRecordset
	Dim oRecordset1
	Dim lErrorNumber
	Dim sErrorsDescription

	sEmployeeIDs = Replace(oRequest("EmployeeNumbers").Item, vbNewLine, ",")
	Do While (InStr(1, sEmployeeIDs, ",,", vbBinaryCompare) > 0)
		sEmployeeIDs = Replace(sEmployeeIDs, ",,", ",")
		If Err.number <> 0 Then Exit Do
	Loop
	If InStr(1, sEmployeeIDs, ",", vbBinaryCompare) = 1 Then sEmployeeIDs = Right(sEmployeeIDs, (Len(sEmployeeIDs) - Len(",")))
	If InStrRev(sEmployeeIDs, ",") = Len(sEmployeeIDs) Then sEmployeeIDs = Left(sEmployeeIDs, (Len(sEmployeeIDs) - Len(",")))
	lPayrollID = CLng(oRequest("PayrollID").Item)

	If IsEmpty(lPayrollID) Then
		lErrorNumber = -1
		sErrorDescription = "No se especifico la quincena en la que se deben de agregar los bloqueos anticipados de los empleados."
	Else
		sEmployeeIDs = Split(sEmployeeIDs, "," , -1, vbBinaryCompare)
		For iIndex = 0 To UBound(sEmployeeIDs)
			aPaymentComponent(N_EMPOYEE_ID_PAYMENT) = CLng(sEmployeeIDs(iIndex))
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From EmployeesBlockPaymentsLKP Where (EmployeeID=" & aPaymentComponent(N_EMPOYEE_ID_PAYMENT) &") And (PayrollID=" & lPayrollID & ")", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset1)
			If lErrorNumber = 0 Then
				If Not oRecordset1.EOF Then
					sErrorsDescription = sErrorsDescription & "Ya existe un registro de bloqueo anticipado para el empleado" & aPaymentComponent(N_EMPOYEE_ID_PAYMENT) & ","
				Else
					sErrorDescription = "No se pudo agragar el registro para el bloqueo anticipado de pagos."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into EmployeesBlockPaymentsLKP (EmployeeID, PayrollID, StatusID) Values (" & aPaymentComponent(N_EMPOYEE_ID_PAYMENT) & ", " & lPayrollID & ", " & aPaymentComponent(N_STATUS_ID_PAYMENT) & ")", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
					If lErrorNumber <> 0 Then
						sErrorsDescription = sErrorsDescription & "No se pudo agregar el registro del bloqueo anticipado del empleado" & aPaymentComponent(N_EMPOYEE_ID_PAYMENT) & ","
					End If
				End If
			End If
		Next
	End If
	If Len(sErrorsDescription) Then
		lErrorNumber = -1
		sErrorDescription = Left(sErrorsDescription, (Len(sErrorsDescription) - Len(","))) & "."
	End If

	AddEmployeeBlockPaymentsLKP = lErrorNumber
	Err.Clear
End Function

Function IssuePayment(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the employee's check payment
'Inputs:  oRequest, oADODBConnection, aCatalogComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "IssuePayment"
	Dim oRecordset
	Dim sContents
	Dim sTemp
	Dim lPaymentID
	Dim iIndex
	Dim lErrorNumber

	sErrorDescription = "No se pudo obtener la información del empleado y de su cheque."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payments.* From EmployeesHistoryListForPayroll, Payments Where (Payments.EmployeeIDEmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Payments.PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Payments.CheckNumber='" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(5)  & "') And (Payments.StatusID=1) And (EmployeesHistoryListForPayroll.EmployeeID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2) & ")", "PaymentsLib.asp", "_root", 000, sErrorDescription, oRecordset)
	Response.Write vbNewLine & "<!-- Query: Select Payments.* From EmployeesHistoryListForPayroll, Payments Where (Payments.EmployeeIDEmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Payments.PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Payments.CheckNumber='" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(5)  & "') And (Payments.StatusID=1) And (EmployeesHistoryListForPayroll.EmployeeID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(2) & ") -->" & vbNewLine
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sContents = ""
			For iIndex = 0 To (oRecordset.Fields.Count - 1)
				sTemp = ""
				sTemp = CStr(oRecordset.Fields(iIndex).Value)
				Err.Clear
				sContents = sContents & sTemp & ", "
			Next
			sContents = Left(sContents, (Len(sContents) - Len(", ")))
			oRecordset.Close
			sErrorDescription = "No se pudo obtener un identificador para el nuevo cheque."
			lErrorNumber = GetNewIDFromTable(oADODBConnection, "Payments", "PaymentID", "", 1, lPaymentID, sErrorDescription)
			If lErrorNumber = 0 Then
				sContents = Split(sContents, ", ")
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payments Set ReplacementNumber='" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(6) & "', StatusID=3, LastUpdate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", ReplacementUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (PaymentID=" & sContents(0) & ")", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If lErrorNumber = 0 Then
					sContents(0) = lPaymentID
					sContents(1) = "'" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(6) & "'"
					sContents(2) = "' '"
					sContents(6) = Left(GetSerialNumberForDate(""), Len("00000000"))
					sContents(7) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(9)
					sContents(9) = aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(4)
					If CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(7)) <> 0 Then
						sContents(12) = "-2"
					Else
						sContents(12) = "-2"'"1"
					End If
					sContents(13) = "'" & sContents(13) & "'"
					sContents(16) = Left(GetSerialNumberForDate(""), Len("00000000"))
					sContents(17) = aLoginComponent(N_USER_ID_LOGIN)
					sContents(18) = -1
					sContents(iIndex)(7) = sContents(iIndex)(7) & ", 0"
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payments (PaymentID, CheckNumber, ReplacementNumber, EmployeeID, PaymentTypeID, PaymentDate, RegisteredDate, CheckDate, CancelDate, AccountID, FromAccountID, CheckAmount, CheckCurrencyID, StatusID, Description, bIsPayment, bIsUpdated, LastUpdate, UserID, ReplacementUserID) Values (" & Join(sContents, ", ") & ")", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
				aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10) = lPaymentID
				lErrorNumber = ModifyCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
				Response.Redirect "Payments.asp?Action=Reexpedition&PaymentID=" & lPaymentID & "&RecordID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & "&DisplayResults=1"
			End If
		Else
			Call RemoveCatalog(oRequest, oADODBConnection, aCatalogComponent, "")
			aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) = -1
			lErrorNumber = -1
			sErrorDescription = "No existen pagos en cheque que cumplan con los criterios de la búsqueda."
		End If
	Else
		Call RemoveCatalog(oRequest, oADODBConnection, aCatalogComponent, "")
		aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) = -1
	End If

	Set oRecordset = Nothing
	IssuePayment = lErrorNumber
	Err.Clear
End Function

Function ModifyPayments(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the employees' check payments
'Inputs:  oRequest, oADODBConnection, aCatalogComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyPayments"
	Dim oRecordset
	Dim sContents
	Dim sTemp
	Dim lPaymentID
	Dim lCheckNumber
	Dim iIndex
	Dim dIndex
	Dim bDone
	Dim aTemp
	Dim lRecordID
	Dim lErrorNumber

	lRecordID = aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0)
	Application.Contents("SIAP_AddPayments") = Application.Contents("SIAP_AddPayments") & "," & lRecordID
	iIndex = 0
	aTemp = Split(Application.Contents("SIAP_AddPayments"), ",")
	Do While InStr(1, Application.Contents("SIAP_AddPayments"), ("," & lRecordID), vbBinaryCompare) <> 1
		iIndex = iIndex + 1
		If iIndex > 1000000 Then
			iIndex = 0
			sErrorDescription = "No se pudieron obtener los registros de la base de datos."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From PaymentsRecords Where (RecordID=" & aTemp(1) & ") And (FirstPaymentID=-1) And (LastPaymentID=-1)", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			If lErrorNumber = 0 Then
				If oRecordset.EOF Then
					Application.Contents("SIAP_AddPayments") = Replace(Application.Contents("SIAP_AddPayments"), ("," & aTemp(1)), "")
				End If
				oRecordset.Close
			End If
		End If
	Loop

	bDone = False
	lCheckNumber = aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(8)
	sErrorDescription = "No se pudo obtener un identificador para el nuevo cheque."
	lErrorNumber = GetNewIDFromTable(oADODBConnection, "Payments", "PaymentID", "", 1, lPaymentID, sErrorDescription)
	aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12) = lPaymentID
	If lErrorNumber = 0 Then
		For dIndex = CDbl(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10)) To CDbl(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11))
			sErrorDescription = "No se pudo obtener el listado de empleados y de sus cheques."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Payments Where (PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (CheckNumber='" & Replace(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10), CDbl(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10)), dIndex) & "') Order By PaymentID Desc", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select * From Payments Where (PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (CheckNumber='" & Replace(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10), CDbl(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10)), dIndex) & "') Order By PaymentID Desc -->" & vbNewLine
			If lErrorNumber = 0 Then
				If Not oRecordset.EOF Then
					bDone = True
					sContents = ""
					For iIndex = 0 To (oRecordset.Fields.Count - 1)
						sTemp = ""
						sTemp = CStr(oRecordset.Fields(iIndex).Value)
						Err.Clear
						sContents = sContents & sTemp & ","
					Next
					oRecordset.Close
					If Len(sContents) > 0 Then sContents = Left(sContents, (Len(sContents) - Len(",")))
					If Len(sContents) > 0 Then
						sContents = Split(sContents, ",")
						sErrorDescription = "No se pudo guardar la información del registro."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payments Set ReplacementNumber='" & lCheckNumber & "', StatusID=2, LastUpdate=" & Left(GetSerialNumberForDate(""), Len("00000000")) & ", ReplacementUserID=" & aLoginComponent(N_USER_ID_LOGIN) & " Where (PaymentID=" & sContents(0) & ")", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						If lErrorNumber = 0 Then
							sContents(0) = lPaymentID
							sContents(1) = "'" & lCheckNumber & "'"
							sContents(2) = "' '"
							sContents(6) = Left(GetSerialNumberForDate(""), Len("00000000"))
							If CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14)) <> 0 Then
								sContents(12) = "-2"
							Else
								sContents(12) = "-2"
							End If
							sContents(14) = "'" & sContents(14) & "'"
							sContents(17) = Left(GetSerialNumberForDate(""), Len("00000000"))
							sContents(18) = aLoginComponent(N_USER_ID_LOGIN)
							sContents(19) = -1
							sContents(7) = sContents(7)
							sErrorDescription = "No se pudo guardar la información del registro."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payments (PaymentID, CheckNumber, ReplacementNumber, EmployeeID, PaymentTypeID, PaymentDate, RegisteredDate, CheckDate, CancelDate, AccountID, FromAccountID, CheckAmount, CheckCurrencyID, StatusID, Description, bIsPayment, bIsUpdated, LastUpdate, UserID, ReplacementUserID) Values (" & Join(sContents, ", ") & ")", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
							If lErrorNumber = 0 Then
								lPaymentID = lPaymentID + 1
								lCheckNumber = lCheckNumber + 1
							End If
						End If
					End If
				Else
					Exit For
					lErrorNumber = -1
				End If
			Else
				Exit For
				lErrorNumber = -1
			End If
		Next
	End If
	If Not bDone Then
		Call RemoveCatalog(oRequest, oADODBConnection, aCatalogComponent, "")
		aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) = -1
		sErrorDescription = "No existen pagos en cheque que cumplan con los criterios de la búsqueda."
	ElseIf lErrorNumber <> 0 Then
		aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11) = Replace(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10), CDbl(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10)), dIndex)
		Call ModifyCatalog(oRequest, oADODBConnection, aCatalogComponent, "")
		sErrorDescription = "No existen pagos en cheque que cumplan con los criterios de la búsqueda."
	Else
		aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(9) = lCheckNumber - 1
		aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13) = lPaymentID - 1
		lErrorNumber = ModifyCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
		Response.Redirect "Payments.asp?Action=Replacement&RecordID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & "&DisplayResults=1"
	End If

	Application.Contents("SIAP_AddPayments") = Replace(Application.Contents("SIAP_AddPayments"), ("," & lRecordID), "")
	Set oRecordset = Nothing
	ModifyPayments = lErrorNumber
	Err.Clear
End Function

Function ModifyPaymentsStatus(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To modify the employees' check payments
'Inputs:  oRequest, oADODBConnection, aCatalogComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ModifyPaymentsStatus"
	Dim oItem
	Dim lPaymentID
	Dim lPayrollID
	Dim lCancelationPayrollID
	Dim lTypeID
	Dim lErrorNumber

	For Each oItem In oRequest
		If InStr(1, oItem, "StatusID_", vbBinaryCompare) > 0 Then
			lPaymentID = CLng(Replace(oItem, "StatusID_", ""))
			Call GetNameFromTable(oADODBConnection, "PaymentsPayrollIDs", lPaymentID, "", "", lPayrollID, "")
			Call GetNameFromTable(oADODBConnection, "PaymentsCancelationPayrollIDs", lPaymentID, "", "", lCancelationPayrollID, "")
			Call GetNameFromTable(oADODBConnection, "PayrollsTypes", lCancelationPayrollID, "", "", lTypeID, "")
			sErrorDescription = "No se pudo actualizar el estatus del pago."
			If (Len(oRequest("Cancelled_" & lPaymentID).Item) > 0) And (CLng(oRequest(oItem).Item) <> 1) Then
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payments Set StatusID=" & oRequest(oItem).Item & ", Description='" & Replace(oRequest("Description").Item, "'", "´") & " ' Where (PaymentID=" & lPaymentID & ")", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
			Else
				If CLng(oRequest(oItem).Item) = 1 Then
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payments Set CancelDate=0, StatusID=" & oRequest(oItem).Item & ", Description='" & Replace(oRequest("Description").Item, "'", "´") & " ' Where (PaymentID=" & lPaymentID & ")", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				Else
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payments Set CancelDate=" & oRequest("CancelationPayrollID").Item & ", StatusID=" & oRequest(oItem).Item & ", Description='" & Replace(oRequest("Description").Item, "'", "´") & " ' Where (PaymentID=" & lPaymentID & ")", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				End If
				If lErrorNumber = 0 Then
					If CLng(oRequest(oItem).Item) = 1 Then
						If (StrComp(lTypeID, "0", vbBinaryCompare) = 0) And (CLng(lCancelationPayrollID) > 0) Then
							sErrorDescription = "No se pudo actualizar el estatus del pago."
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & lCancelationPayrollID & " Where (EmployeeID=" & oRequest("EmployeeNumber").Item & ") And (RecordID=" & lPayrollID & ")", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
					Else
						sErrorDescription = "No se pudo actualizar el estatus del pago."
						lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payroll_" & oRequest("CancelationPayrollID").Item & " Where (EmployeeID=" & oRequest("EmployeeNumber").Item & ") And (RecordID=" & lPayrollID & ")", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						If (lErrorNumber = 0) And (CLng(oRequest(oItem).Item) <> 1) Then
							sErrorDescription = "No se pudo actualizar el estatus del pago."
							'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & oRequest("CancelationPayrollID").Item & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, " & lPayrollID & " As RecordID, EmployeeID, ConceptID, PayrollTypeID, (ConceptAmount)*(-1), ConceptTaxes, RecordID As ConceptRetention, UserID From Payroll_" & lPayrollID & " Where (EmployeeID=" & oRequest("EmployeeNumber").Item & ")", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null) 'Revisar: MLima comento por no saber porque se hacen negativos
							lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into Payroll_" & oRequest("CancelationPayrollID").Item & " (RecordDate, RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, ConceptRetention, UserID) Select RecordDate, " & lPayrollID & " As RecordID, EmployeeID, ConceptID, PayrollTypeID, ConceptAmount, ConceptTaxes, RecordID As ConceptRetention, UserID From Payroll_" & lPayrollID & " Where (EmployeeID=" & oRequest("EmployeeNumber").Item & ")", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
						End If
					End If
				End If
			End If
		End If
	Next

	ModifyPaymentsStatus = lErrorNumber
	Err.Clear
End Function

Function RemoveEmployeeBlockPaymentsLKP(oRequest, oADODBConnection, aPaymentComponent, sErrorDescription)
'************************************************************
'Purpose: To remove payment block to employee for specific payroll
'Inputs:  oRequest, oADODBConnection, iActive
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemoveEmployeeBlockPaymentsLKP"
	Dim oItem
	Dim iIndex
	Dim sEmployeeIDs
	Dim lPayrollID
	Dim oRecordset
	Dim oRecordset1
	Dim lErrorNumber
	Dim sErrorsDescription

	lPayrollID = CLng(oRequest("PayrollID").Item)

	If IsEmpty(lPayrollID) Then
		lErrorNumber = -1
		sErrorDescription = "No se especifico la quincena en la que se debe eliminar el bloqueos anticipado de pago del empleado."
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete from EmployeesBlockPaymentsLKP Where (EmployeeID=" & aPaymentComponent(N_EMPOYEE_ID_PAYMENT) &") And (PayrollID=" & lPayrollID & ")", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset1)
		If lErrorNumber = 0 Then
			lErrorNumber = -1
			sErrorDescription = "No se pudo eliminar el bloqueos anticipado de pago del empleado " & aPaymentComponent(N_EMPOYEE_ID_PAYMENT)
		End If
	End If
	aPaymentComponent(N_EMPOYEE_ID_PAYMENT) = -1

	RemoveEmployeeBlockPaymentsLKP = lErrorNumber
	Err.Clear
End Function

Function RemovePayments(oRequest, oADODBConnection, sCondition, sErrorDescription)
'************************************************************
'Purpose: To remove the employees' check payments
'Inputs:  oRequest, oADODBConnection, sCondition
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "RemovePayments"
	Dim lErrorNumber

	If Len(oRequest("RecordID2").Item) > 0 Then
		sErrorDescription = "No se pudieron eliminar los cheques."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payments Set StatusID=1, ReplacementNumber=' ' Where (ReplacementNumber='" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(6) & "') And (PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (StatusID=3)", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	Else
		sErrorDescription = "No se pudieron eliminar los cheques."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payments Set StatusID=1, ReplacementNumber=' ' Where (ReplacementNumber In (Select CheckNumber From Payments Where " & sCondition & ")) And (PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (StatusID=2)", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If
	sErrorDescription = "No se pudieron eliminar los cheques."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From Payments Where " & sCondition & " And (StatusID In (-2,-1,1))", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)

	RemovePayments = lErrorNumber
	Err.Clear
End Function

Function SetActiveForEmployeeAccount(oRequest, oADODBConnection, iActive, sErrorDescription)
'************************************************************
'Purpose: To modify the active field for the BankAccounts table
'Inputs:  oRequest, oADODBConnection, iActive
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "SetActiveForEmployeeAccount"
	Dim sEmployeeIDs
	Dim lStartDate
	Dim lEndDate
	Dim lAccountID
	Dim sFilePath
	Dim sFileContents
	Dim asRows
	Dim asCells
	Dim iIndex
	Dim oRecordset
	Dim lErrorNumber

	sEmployeeIDs = Replace(oRequest("EmployeeNumbers").Item, vbNewLine, ",")
	Do While (InStr(1, sEmployeeIDs, ",,", vbBinaryCompare) > 0)
		sEmployeeIDs = Replace(sEmployeeIDs, ",,", ",")
		If Err.number <> 0 Then Exit Do
	Loop
	If InStr(1, sEmployeeIDs, ",", vbBinaryCompare) = 1 Then sEmployeeIDs = Right(sEmployeeIDs, (Len(sEmployeeIDs) - Len(",")))
	If InStrRev(sEmployeeIDs, ",") = Len(sEmployeeIDs) Then sEmployeeIDs = Left(sEmployeeIDs, (Len(sEmployeeIDs) - Len(",")))
	lStartDate = oRequest("StartPaymentYear").Item & oRequest("StartPaymentMonth").Item & oRequest("StartPaymentDay").Item
	lEndDate = oRequest("EndPaymentYear").Item & oRequest("EndPaymentMonth").Item & oRequest("EndPaymentDay").Item

	If iActive = 0 Then
		sErrorDescription = "No se pudieron modificar las cuentas bancarias."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From BankAccounts Where (EmployeeID In (" & sEmployeeIDs & ")) And (AccountNumber<>'.') And (EndDate=30000000)", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				sFilePath = Server.MapPath(REPORTS_PATH & "User_" & aLoginComponent(N_USER_ID_LOGIN) & "\Accounts_" & aLoginComponent(N_USER_ID_LOGIN) & "_" & GetSerialNumberForDate("") & ".txt")
				Do While Not oRecordset.EOF
					lErrorNumber = AppendTextToFile(sFilePath, CStr(oRecordset.Fields("EmployeeID").Value) & ", " & CStr(oRecordset.Fields("BankID").Value) & ", " & CStr(oRecordset.Fields("AccountNumber").Value), sErrorDescription)
					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
				oRecordset.Close
				sErrorDescription = "No se pudieron modificar las cuentas bancarias."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update BankAccounts Set EndDate=" & AddDaysToSerialDate(lStartDate, -1) & ", Active=0 Where (EmployeeID In (" & sEmployeeIDs & ")) And (AccountNumber<>'.') And (EndDate=30000000)", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
				If lErrorNumber = 0 Then
					sFileContents = GetFileContents(sFilePath, sErrorDescription)
					Call DeleteFile(sFilePath, "")
					asRows = Split(sFileContents, vbNewLine)
					For iIndex = 0 To UBound(asRows)
						If Len(asRows(iIndex)) > 0 Then
							lErrorNumber = GetNewIDFromTable(oADODBConnection, "BankAccounts", "AccountID", "", "-1", lAccountID, sErrorDescription)
							If lErrorNumber = 0 Then
								sErrorDescription = "No se pudieron modificar las cuentas bancarias."
								lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into BankAccounts (AccountID, EmployeeID, BankID, AccountNumber, StartDate, EndDate, Active) Values (" & lAccountID & "," & asRows(iIndex) & ", " & lStartDate & ", " & lEndDate & ", 0)", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudieron modificar las cuentas bancarias."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Insert Into BankAccounts (AccountID, EmployeeID, BankID, AccountNumber, StartDate, EndDate, Active) Values (" & lAccountID + 1 & "," & asRows(iIndex) & ", " & AddDaysToSerialDate(lEndDate, 1) & ", 30000000, 1)", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								End If
								asCells = Split(asRows(iIndex), ", ")
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudieron modificar las cuentas bancarios."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Delete From BankAccounts Where (EmployeeID=" & asCells(0) & ") And (StartDate>EndDate)", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								End If
								If lErrorNumber = 0 Then
									sErrorDescription = "No se pudieron modificar los depósitos bancarios."
									lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payments Set StatusID=4, Description='" & Replace(oRequest("Description").Item, "'", "´") & "' Where (EmployeeID=" & asCells(0) & ") And (PaymentDate>=" & lStartDate & ") And (PaymentDate<=" & lEndDate & ")", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
								End If
							End If
						End If
					Next
				End If
			Else
				oRecordset.Close
				lErrorNumber = -1
				sErrorDescription = "No existen cuentas bancarias que cumplan con los criterios de la búsqueda."
			End If
		End If
	Else
		sErrorDescription = "No se pudieron modificar los depósitos bancarios."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payments Set StatusID=1, Description='" & Replace(oRequest("Description").Item, "'", "´") & "' Where (EmployeeID=" & sEmployeeIDs & ") And (StatusID=4)", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
	End If

	SetActiveForEmployeeAccount = lErrorNumber
	Err.Clear
End Function

Function DisplayPaymentsSearchForm(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the search HTML form
'Inputs:  oRequest, oADODBConnection
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPaymentsSearchForm"

	Response.Write "<FORM NAME=""SearchFrm"" ID=""SearchFrm"" ACTION=""Payments.asp"" METHOD=""GET"">"
		Response.Write "<B>BÚSQUEDA DE CHEQUES</B><BR /><BR />"
		Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número del cheque:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""CheckNumber"" ID=""CheckNumberTxt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & oRequest("CheckNumber").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Número de empleado:&nbsp;</FONT></TD>"
				Response.Write "<TD><INPUT TYPE=""TEXT"" NAME=""EmployeeID"" ID=""EmployeeIDTxt"" SIZE=""10"" MAXLENGTH=""10"" VALUE=""" & oRequest("EmployeeID").Item & """ CLASS=""TextFields"" /></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Tipo de pago:&nbsp;</FONT></TD>"
				Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""PaymentTypeID"" ID=""PaymentTypeIDCmb"" SIZE=""5"" MULTIPLE=""3"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "PaymentTypes", "PaymentTypeID", "PaymentTypeName", "(Active=1)", "PaymentTypeName", oRequest("PaymentTypeID").Item, "Ninguno;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha valor:&nbsp;</FONT></TD>"
				Response.Write "<TD>"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartPaymentYear").Item), CInt(oRequest("StartPaymentMonth").Item), CInt(oRequest("StartPaymentDay").Item), "StartPaymentYear", "StartPaymentMonth", "StartPaymentDay", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndPaymentYear").Item), CInt(oRequest("EndPaymentMonth").Item), CInt(oRequest("EndPaymentDay").Item), "EndPaymentYear", "EndPaymentMonth", "EndPaymentDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha de registro:&nbsp;</FONT></TD>"
				Response.Write "<TD>"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartRegisteredYear").Item), CInt(oRequest("StartRegisteredMonth").Item), CInt(oRequest("StartRegisteredDay").Item), "StartRegisteredYear", "StartRegisteredMonth", "StartRegisteredDay", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndRegisteredYear").Item), CInt(oRequest("EndRegisteredMonth").Item), CInt(oRequest("EndRegisteredDay").Item), "EndRegisteredYear", "EndRegisteredMonth", "EndRegisteredDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			Response.Write "<TR>"
				Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Fecha del cheque:&nbsp;</FONT></TD>"
				Response.Write "<TD>"
					Response.Write "<FONT FACE=""Arial"" SIZE=""2"">Entre </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("StartCheckYear").Item), CInt(oRequest("StartCheckMonth").Item), CInt(oRequest("StartCheckDay").Item), "StartCheckYear", "StartCheckMonth", "StartCheckDay", N_FORM_START_YEAR, Year(Date()), True, True)
					Response.Write "<FONT FACE=""Arial"" SIZE=""2""> y el </FONT>"
					Response.Write DisplayDateCombos(CInt(oRequest("EndCheckYear").Item), CInt(oRequest("EndCheckMonth").Item), CInt(oRequest("EndCheckDay").Item), "EndCheckYear", "EndCheckMonth", "EndCheckDay", N_FORM_START_YEAR, Year(Date()), True, True)
				Response.Write "</TD>"
			Response.Write "</TR>"
			If Not B_ISSSTE Then
				Response.Write "<TR>"
					Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Cuenta del beneficiario:&nbsp;</FONT></TD>"
					Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""AccountID"" ID=""AccountIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BankAccounts", "AccountID", "AccountNumber", "(Active=1)", "AccountNumber", oRequest("AccountID").Item, "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
			End If
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Cuenta girada:&nbsp;</FONT></TD>"
				Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""FromAccountID"" ID=""FromAccountIDCmb"" SIZE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todas</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "BankAccounts", "AccountID", "AccountNumber", "(EmployeeID=-1) and (Active=1)", "AccountNumber", oRequest("AccountID").Item, "Ninguna;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			If Not B_ISSSTE Then
				Response.Write "<TR>"
					Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Moneda:&nbsp;</FONT></TD>"
					Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""CurrencyID"" ID=""CurrencyIDCmb"" SIZE=""5"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
						Response.Write "<OPTION VALUE="""">Todas</OPTION>"
						Response.Write GenerateListOptionsFromQuery(oADODBConnection, "Currencies", "CurrencyID", "CurrencyName", "(Active=1)", "CurrencyName", oRequest("CurrencyID").Item, "Ninguna;;;-1", sErrorDescription)
					Response.Write "</SELECT></TD>"
				Response.Write "</TR>"
			End If
			Response.Write "<TR>"
				Response.Write "<TD VALIGN=""TOP""><FONT FACE=""Arial"" SIZE=""2"">Estatus:&nbsp;</FONT></TD>"
				Response.Write "<TD VALIGN=""TOP""><SELECT NAME=""StatusID"" ID=""StatusIDCmb"" SIZE=""6"" MULTIPLE=""1"" CLASS=""Lists"" onChange=""if (this.options[0].selected) {UnselectAllItemsFromList(this); this.options[0].selected = true;}"">"
					Response.Write "<OPTION VALUE="""">Todos</OPTION>"
					Response.Write GenerateListOptionsFromQuery(oADODBConnection, "StatusPayments", "StatusID", "StatusName", "(Active=1)", "StatusName", oRequest("StatusID").Item, "Ninguna;;;-1", sErrorDescription)
				Response.Write "</SELECT></TD>"
			Response.Write "</TR>"
			Response.Write "<TR><TD COLSPAN=""2""><BR /><INPUT TYPE=""SUBMIT"" NAME=""DoSearch"" ID=""DoSearchBtn"" VALUE=""Buscar Cheques"" CLASS=""Buttons"" /></TD></TR>"
		Response.Write "</TABLE>"
	Response.Write "</FORM>"

	DisplayPaymentsSearchForm = Err.number
End Function

Function DisplayNewPaymentsTable(oRequest, oADODBConnection, bForExport, aCatalogComponent, sErrorDescription)
'************************************************************
'Purpose: To display the information about the new payments
'		  from the database in a table
'Inputs:  oRequest, oADODBConnection, bForExport, aCatalogComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayNewPaymentsTable"
	Dim sCondition
	Dim dTotal
	Dim oRecordset
	Dim iRecordCounter
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	If Len(oRequest("RecordID2").Item) > 0 Then aCatalogComponent(S_TABLE_NAME_CATALOG) = "PaymentsRecords2"
	Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
		Case "PaymentsRecords"
			If CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14)) = 2 Then 'Pensión alimenticia
				sCondition = " And (Payments.PaymentID>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12) & ") And (Payments.PaymentID<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13) & ")"
				sErrorDescription = "No se pudieron obtener los empleados que cumplen con los criterios de la búsqueda."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryListForPayroll.EmployeeNumber, Companies.CompanyShortName, Companies.CompanyName, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, Zones1.ZoneName As ZoneName1, PaymentCenters.AreaCode As PaymentCenterCode, PaymentCenters.AreaName As PaymentCenterName, EmployeeTypes.EmployeeTypeShortName, EmployeeTypes.EmployeeTypeName, CheckNumber, CheckAmount, StatusName From EmployeesBeneficiariesLKP, Payments, StatusPayments, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones As Zones1, Zones As Zones2, Zones As Zones3, EmployeeTypes, BankAccounts Where (Payments.StatusID=StatusPayments.StatusID) And (Payments.EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.CompanyID=Companies.CompanyID) And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones3.ZoneID) And (Zones3.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesBeneficiariesLKP.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Order by Companies.CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeesBeneficiariesLKP.BeneficiaryNumber", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesHistoryListForPayroll.EmployeeNumber, Companies.CompanyShortName, Companies.CompanyName, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, Zones1.ZoneName As ZoneName1, PaymentCenters.AreaCode As PaymentCenterCode, PaymentCenters.AreaName As PaymentCenterName, EmployeeTypes.EmployeeTypeShortName, EmployeeTypes.EmployeeTypeName, CheckNumber, CheckAmount, StatusName From EmployeesBeneficiariesLKP, Payments, StatusPayments, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones As Zones1, Zones As Zones2, Zones As Zones3, EmployeeTypes, BankAccounts Where (Payments.StatusID=StatusPayments.StatusID) And (Payments.EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.CompanyID=Companies.CompanyID) And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesBeneficiariesLKP.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones3.ZoneID) And (Zones3.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesBeneficiariesLKP.BeneficiaryNumber=BankAccounts.EmployeeID) And (EmployeesBeneficiariesLKP.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesBeneficiariesLKP.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Order by Companies.CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeesBeneficiariesLKP.BeneficiaryNumber -->" & vbNewLine
			ElseIf CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14)) = 4 Then 'Acreedores
				sCondition = " And (Payments.PaymentID>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12) & ") And (Payments.PaymentID<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13) & ")"
				sErrorDescription = "No se pudieron obtener los empleados que cumplen con los criterios de la búsqueda."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryListForPayroll.EmployeeNumber, Companies.CompanyShortName, Companies.CompanyName, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, Zones1.ZoneName As ZoneName1, PaymentCenters.AreaCode As PaymentCenterCode, PaymentCenters.AreaName As PaymentCenterName, EmployeeTypes.EmployeeTypeShortName, EmployeeTypes.EmployeeTypeName, CheckNumber, CheckAmount, StatusName From EmployeesCreditorsLKP, Payments, StatusPayments, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones As Zones1, Zones As Zones2, Zones As Zones3, EmployeeTypes, BankAccounts Where (Payments.StatusID=StatusPayments.StatusID) And (Payments.EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.CompanyID=Companies.CompanyID) And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones3.ZoneID) And (Zones3.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesCreditorsLKP.EmployeeID=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesCreditorsLKP.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Order by Companies.CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeesCreditorsLKP.CreditorNumber", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select EmployeesCreditorsLKP.CreditorNumber, EmployeesHistoryListForPayroll.EmployeeNumber, Companies.CompanyShortName, Companies.CompanyName, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, Zones1.ZoneName As ZoneName1, PaymentCenters.AreaCode As PaymentCenterCode, PaymentCenters.AreaName As PaymentCenterName, EmployeeTypes.EmployeeTypeShortName, EmployeeTypes.EmployeeTypeName, CheckNumber, CheckAmount, StatusName From EmployeesCreditorsLKP, Payments, StatusPayments, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones As Zones1, Zones As Zones2, Zones As Zones3, EmployeeTypes, BankAccounts Where (Payments.StatusID=StatusPayments.StatusID) And (Payments.EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.CompanyID=Companies.CompanyID) And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesCreditorsLKP.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones3.ZoneID) And (Zones3.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesCreditorsLKP.CreditorNumber=BankAccounts.EmployeeID) And (EmployeesCreditorsLKP.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesCreditorsLKP.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Order by Companies.CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeesCreditorsLKP.CreditorNumber -->" & vbNewLine
			ElseIf CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(8)) = -1 Then ' No es reexpedición
				sCondition = " And (Payments.PaymentID>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12) & ") And (Payments.PaymentID<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13) & ")"
				If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14)) <> 1 Then ' No es Depósito
					'sCondition = sCondition & " And (Payments.StatusID=-2)"
				End If
				sErrorDescription = "No se pudieron obtener los empleados que cumplen con los criterios de la búsqueda."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeNumber, Companies.CompanyShortName, Companies.CompanyName, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, Zones1.ZoneName As ZoneName1, PaymentCenters.AreaCode As PaymentCenterCode, PaymentCenters.AreaName As PaymentCenterName, EmployeeTypes.EmployeeTypeShortName, EmployeeTypes.EmployeeTypeName, CheckNumber, CheckAmount, StatusName From Payments, StatusPayments, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones As Zones1, Zones As Zones2, Zones As Zones3, EmployeeTypes Where (Payments.StatusID=StatusPayments.StatusID) And (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones3.ZoneID) And (Zones3.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Companies.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Order by Companies.CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeeNumber", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select EmployeeNumber, Companies.CompanyShortName, Companies.CompanyName, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, Zones1.ZoneName As ZoneName1, PaymentCenters.AreaCode As PaymentCenterCode, PaymentCenters.AreaName As PaymentCenterName, EmployeeTypes.EmployeeTypeShortName, EmployeeTypes.EmployeeTypeName, CheckNumber, CheckAmount, StatusName From Payments, StatusPayments, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones As Zones1, Zones As Zones2, Zones As Zones3, EmployeeTypes Where (Payments.StatusID=StatusPayments.StatusID) And (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones3.ZoneID) And (Zones3.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Companies.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Order by Companies.CompanyShortName, Areas1.AreaCode, PaymentCenters.AreaCode, EmployeeNumber -->" & vbNewLine
			Else
				sCondition = " And (Payments.PaymentID>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12) & ") And (Payments.PaymentID<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13) & ")"
				sErrorDescription = "No se pudieron obtener los empleados que cumplen con los criterios de la búsqueda."
				lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeNumber, Companies.CompanyShortName, Companies.CompanyName, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, Zones1.ZoneName As ZoneName1, PaymentCenters.AreaCode As PaymentCenterCode, PaymentCenters.AreaName As PaymentCenterName, EmployeeTypes.EmployeeTypeShortName, EmployeeTypes.EmployeeTypeName, Payments.CheckNumber, Payments.CheckAmount, Replacements.CheckNumber As ReplacedNumber, StatusName From Payments, StatusPayments, Payments As Replacements, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones As Zones1, Zones As Zones2, Zones As Zones3, EmployeeTypes Where (Payments.StatusID=StatusPayments.StatusID) And (Payments.PaymentDate=Replacements.PaymentDate) And (Payments.CheckNumber=Replacements.ReplacementNumber) And (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones3.ZoneID) And (Zones3.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (Companies.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Order by Companies.CompanyShortName, Areas1.AreaCode, Areas2.AreaCode, EmployeeNumber", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
				Response.Write vbNewLine & "<!-- Query: Select EmployeeNumber, Companies.CompanyShortName, Companies.CompanyName, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, Zones1.ZoneName As ZoneName1, PaymentCenters.AreaCode As PaymentCenterCode, PaymentCenters.AreaName As PaymentCenterName, EmployeeTypes.EmployeeTypeShortName, EmployeeTypes.EmployeeTypeName, Payments.CheckNumber, Payments.CheckAmount, Replacements.CheckNumber As ReplacedNumber, StatusName From Payments, StatusPayments, Payments As Replacements, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones As Zones1, Zones As Zones2, Zones As Zones3, EmployeeTypes, BankAccounts Where (Payments.StatusID=StatusPayments.StatusID) And (Payments.PaymentDate=Replacements.PaymentDate) And (Payments.CheckNumber=Replacements.ReplacementNumber) And (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones3.ZoneID) And (Zones3.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.EmployeeID=BankAccounts.EmployeeID) And (Companies.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Order by Companies.CompanyShortName, Areas1.AreaCode, Areas2.AreaCode, EmployeeNumber -->" & vbNewLine
			End If
		Case "PaymentsRecords2"
			sCondition = "And (Payments.PaymentID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11) & ")"
			sErrorDescription = "No se pudieron obtener los empleados que cumplen con los criterios de la búsqueda."
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeNumber, Companies.CompanyShortName, Companies.CompanyName, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, Zones1.ZoneName As ZoneName1, PaymentCenters.AreaCode As PaymentCenterCode, PaymentCenters.AreaName As PaymentCenterName, EmployeeTypes.EmployeeTypeShortName, EmployeeTypes.EmployeeTypeName, Payments.CheckNumber, Payments.CheckAmount, Replacements.CheckNumber As ReplacedNumber, StatusName From Payments, StatusPayments, Payments As Replacements, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones As Zones1, Zones As Zones2, Zones As Zones3, EmployeeTypes, BankAccounts Where (Payments.StatusID=StatusPayments.StatusID) And (Payments.PaymentDate=Replacements.PaymentDate) And (Payments.CheckNumber=Replacements.ReplacementNumber) And (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones3.ZoneID) And (Zones3.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.EmployeeID=BankAccounts.EmployeeID) And (Companies.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Order by Companies.CompanyShortName, Areas1.AreaCode, Areas2.AreaCode, EmployeeNumber", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select EmployeeNumber, Companies.CompanyShortName, Companies.CompanyName, Areas1.AreaCode As AreaCode1, Areas1.AreaName As AreaName1, Areas2.AreaCode As AreaCode2, Areas2.AreaName As AreaName2, Zones1.ZoneName As ZoneName1, PaymentCenters.AreaCode As PaymentCenterCode, PaymentCenters.AreaName As PaymentCenterName, EmployeeTypes.EmployeeTypeShortName, EmployeeTypes.EmployeeTypeName, Payments.CheckNumber, Payments.CheckAmount, Replacements.CheckNumber As ReplacedNumber, StatusName From Payments, StatusPayments, Payments As Replacements, EmployeesHistoryListForPayroll, Companies, Areas As Areas1, Areas As Areas2, Areas As PaymentCenters, Zones As Zones1, Zones As Zones2, Zones As Zones3, EmployeeTypes, BankAccounts Where (Payments.StatusID=StatusPayments.StatusID) And (Payments.PaymentDate=Replacements.PaymentDate) And (Payments.CheckNumber=Replacements.ReplacementNumber) And (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.CompanyID=Companies.CompanyID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (Areas2.ZoneID=Zones3.ZoneID) And (Zones3.ParentID=Zones2.ZoneID) And (Zones2.ParentID=Zones1.ZoneID) And (EmployeesHistoryListForPayroll.EmployeeTypeID=EmployeeTypes.EmployeeTypeID) And (EmployeesHistoryListForPayroll.EmployeeID=BankAccounts.EmployeeID) And (Companies.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Companies.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (PaymentCenters.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeeTypes.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (BankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Order by Companies.CompanyShortName, Areas1.AreaCode, Areas2.AreaCode, EmployeeNumber -->" & vbNewLine
		Case ""
	End Select
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			If Not bForExport Then Call DisplayIncrementalFetch(oRequest, CInt(oRequest("StartPage").Item), ROWS_REPORT, oRecordset)
			Response.Write "<FONT FACE=""Arial"" SIZE=""2"">"
				Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
					Case "PaymentsRecords"
						If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(8)) = -1 Then
							Response.Write "<B>Número de registros con folio:</B> " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11) - aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10) + 1 & "<BR />"
							Response.Write "<B>Rango de cheques generados:</B> " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(10) & " a " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11) & "<BR />"
							Response.Write "<B>Folio inicial siguiente:</B> " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(11) + 1 & "<BR />"

							asColumnsTitles = Split("Unidad administrativa,Centro de trabajo,Entidad,Número del empleado,Empresa,Tipo de tabulador,Número de cheque,Monto,Estatus del cheque", ",", -1, vbBinaryCompare)
							asCellWidths = Split("100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
							asCellAlignments = Split(",,,,,,,RIGHT,", ",", -1, vbBinaryCompare)
						Else
							Response.Write "<B>Total de cheques a reponer:</B> " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(9) - aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(8) + 1 & "<BR />"
							Response.Write "<B>Último cheque de la reposición:</B> " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(9) & "<BR />"
							Response.Write "<B>Inicio de la siguiente reposición:</B> " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(9) + 1 & "<BR />"

							asColumnsTitles = Split("Unidad administrativa,Centro de trabajo,Entidad,Número del empleado,Empresa,Tipo de tabulador,Número de cheque,Cheque repuesto,Monto,Estatus del cheque", ",", -1, vbBinaryCompare)
							asCellWidths = Split("100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
							asCellAlignments = Split(",,,,,,,,RIGHT,", ",", -1, vbBinaryCompare)
						End If
					Case "PaymentsRecords2"
						asColumnsTitles = Split("Unidad administrativa,Centro de trabajo,Número del empleado,Empresa,Tipo de tabulador,Número de cheque,Cheque repuesto,Monto,Estatus del cheque", ",", -1, vbBinaryCompare)
						asCellWidths = Split("100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
						asCellAlignments = Split(",,,,,,,RIGHT,", ",", -1, vbBinaryCompare)
				End Select
				Response.Write "<BR />"
			Response.Write "</FONT>"
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
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
				dTotal = 0.0
				Do While Not oRecordset.EOF
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("AreaCode1").Value) & ". " & CStr(oRecordset.Fields("AreaName1").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("PaymentCenterCode").Value) & ". " & CStr(oRecordset.Fields("PaymentCenterName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ZoneName1").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
						Select Case CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14))
							Case 2
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("BeneficiaryNumber").Value)) & "<BR />Empleado:&nbsp;"
							Case 4
								sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("CreditorNumber").Value)) & "<BR />Empleado:&nbsp;"
						End Select
					sRowContents = sRowContents & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CompanyShortName").Value) & ". " & CStr(oRecordset.Fields("CompanyName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeTypeShortName").Value) & ". " & CStr(oRecordset.Fields("EmployeeTypeName").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CheckNumber").Value))
					Select Case aCatalogComponent(S_TABLE_NAME_CATALOG)
						Case "PaymentsRecords"
							If CLng(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(8)) = -1 Then
							Else
								sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ReplacedNumber").Value))
							End If
						Case "PaymentsRecords2"
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ReplacedNumber").Value))
					End Select
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("CheckAmount").Value), 2, True, False, True)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StatusName").Value))
					dTotal = CDbl(dTotal) + CDbl(oRecordset.Fields("CheckAmount").Value)

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
				sRowContents = "<SPAN COLS=""6"" />" & TABLE_SEPARATOR & "<B>TOTAL</B>" & TABLE_SEPARATOR & "<B>" & FormatNumber(CDbl(dTotal), 2, True, False, True) & "</B>" & TABLE_SEPARATOR
				asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
				Else
					lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
				End If
			Response.Write "</TABLE>"
			oRecordset.Close
		Else
			lErrorNumber = -1
			sErrorDescription = "No existen empleados que cumplan con los criterios de la búsqueda."
		End If
	End If

	Set oRecordset = Nothing
	DisplayNewPaymentsTable = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeePaymentsTable(oRequest, oADODBConnection, iPaymentType, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the information about the employee's
'		  payments from the database in a table
'Inputs:  oRequest, oADODBConnection, iPaymentType, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeePaymentsTable"
	Dim sEmployeeIDs
	Dim sCondition
	Dim sClosed
	Dim oRecordset
	Dim sTemp
	Dim sComboHTML
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
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

	sCondition = ""
	'If Len(oRequest("CompanyID").Item) > 0 Then sCondition = sCondition & " And (EmployeesHistoryListForPayroll.CompanyID=" & oRequest("CompanyID").Item & ")"
	If Len(oRequest("EmployeeID").Item) > 0 Then
		sCondition = sCondition & " And (Payments.EmployeeID=" & oRequest("EmployeeID").Item & ")"
	ElseIf Len(oRequest("PayrollID").Item) > 0 Then
		If StrComp(oRequest("PayrollID").Item, "-1", vbBinaryCompare) <> 0 Then sCondition = sCondition & " And (EmployeesHistoryListForPayroll.PayrollID=" & oRequest("PayrollID").Item & ")"
	End If
	If Len(oRequest("EmployeeNumber").Item) > 0 Then sCondition = sCondition & " And (EmployeesHistoryListForPayroll.EmployeeNumber='" & Right(("000000" & oRequest("EmployeeNumber").Item), Len("000000")) & "')"
	If Len(oRequest("EmployeeNumbers").Item) > 0 Then
		sEmployeeIDs = Replace(oRequest("EmployeeNumbers").Item, vbNewLine, ",")
		Do While (InStr(1, sEmployeeIDs, ",,", vbBinaryCompare) > 0)
			sEmployeeIDs = Replace(sEmployeeIDs, ",,", ",")
			If Err.number <> 0 Then Exit Do
		Loop
		sCondition = sCondition & " And (Payments.EmployeeID In (" & sEmployeeIDs & "))"
	End If
	If Len(oRequest("PaymentStatusID").Item) > 0 Then sCondition = sCondition & " And (Payments.StatusID In (" & oRequest("PaymentStatusID").Item & "))"
	If Len(oRequest("CancelPayment").Item) > 0 Then sCondition = sCondition & " And (Payments.StatusID Not In (-2,-1,1,2,3))"
	Select Case iPaymentType
		Case -1
			If CLng(oRequest("EmployeeNumber").Item) >= 600000 And CLng(oRequest("EmployeeNumber").Item) <= 800000 Then
				sCondition = sCondition & " And (BankAccounts.AccountNumber='.') And (Payments.StatusID In (1,4))"
			Else
				sCondition = sCondition & " And (EmployeesHistoryListForPayroll.AccountNumber='.') And (Payments.StatusID In (1,4))"
			End If
		Case 0 'Cancelar pagos
			Select Case oRequest("PaymentType").Item
				Case "0"
					If CLng(oRequest("EmployeeNumber").Item) >= 600000 And CLng(oRequest("EmployeeNumber").Item) <= 800000 Then
						sCondition = sCondition & " And (BankAccounts.AccountNumber='.')"
					Else
						sCondition = sCondition & " And (EmployeesHistoryListForPayroll.AccountNumber='.')"
					End If
				Case "1"
					If CLng(oRequest("EmployeeNumber").Item) >= 600000 And CLng(oRequest("EmployeeNumber").Item) <= 800000 Then
						sCondition = sCondition & " And (BankAccounts.AccountNumber<>'.')"
					Else
						sCondition = sCondition & " And (EmployeesHistoryListForPayroll.AccountNumber<>'.')"
					End If
			End Select
		Case 1 'Bloquear depósitos
			If CLng(oRequest("EmployeeNumber").Item) >= 600000 And CLng(oRequest("EmployeeNumber").Item) <= 800000 Then
				sCondition = sCondition & " And (BankAccounts.AccountNumber<>'.') And (Payments.StatusID In (1,4))"
			Else
				sCondition = sCondition & " And (EmployeesHistoryListForPayroll.AccountNumber<>'.') And (Payments.StatusID In (1,4))"
			End If
	End Select
	sErrorDescription = "No se pudieron obtener los empleados que cumplen con los criterios de la búsqueda."
	If CLng(oRequest("EmployeeNumber").Item) >= 600000 And CLng(oRequest("EmployeeNumber").Item) <= 800000 Then
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payments.PaymentID, PaymentDate, CheckDate, CancelDate, EmployeeNumber, CheckNumber, LastUpdate, ReplacementNumber, CheckAmount, StatusPayments.StatusID, StatusPayments.StatusShortName, StatusPayments.StatusName From Payments, EmployeesHistoryListForPayroll, Areas As PaymentCenters, Companies, StatusPayments, BankAccounts Where (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payments.PaymentDate=EmployeesHistoryListForPayroll.PayrollID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.CompanyID=Companies.CompanyID) And (Payments.StatusID=StatusPayments.StatusID) And (Payments.EmployeeID=BankAccounts.EmployeeID) And (BankAccounts.StartDate<=EmployeesHistoryListForPayroll.PayrollID) And (Companies.StartDate<=EmployeesHistoryListForPayroll.PayrollID) And (Companies.EndDate>=EmployeesHistoryListForPayroll.PayrollID) And (BankAccounts.Active = 1) " & sCondition & " Order By PaymentDate Desc, CheckDate Desc", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select Payments.PaymentID, PaymentDate, CheckDate, CancelDate, EmployeeNumber, CheckNumber, LastUpdate, ReplacementNumber, CheckAmount, StatusPayments.StatusID, StatusPayments.StatusShortName, StatusPayments.StatusName From Payments, EmployeesHistoryListForPayroll, Areas As PaymentCenters, Companies, StatusPayments, BankAccounts Where (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payments.PaymentDate=EmployeesHistoryListForPayroll.PayrollID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.CompanyID=Companies.CompanyID) And (Payments.StatusID=StatusPayments.StatusID) And (Payments.EmployeeID=BankAccounts.EmployeeID) And (BankAccounts.StartDate<=EmployeesHistoryListForPayroll.PayrollID) And (BankAccounts.EndDate>=EmployeesHistoryListForPayroll.PayrollID) And (Companies.StartDate<=EmployeesHistoryListForPayroll.PayrollID) And (Companies.EndDate>=EmployeesHistoryListForPayroll.PayrollID) And (BankAccounts.Active = 1) " & sCondition & " Order By PaymentDate Desc, CheckDate Desc -->" & vbNewLine
	Else
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payments.PaymentID, PaymentDate, CheckDate, CancelDate, EmployeeNumber, CheckNumber, LastUpdate, ReplacementNumber, CheckAmount, StatusPayments.StatusID, StatusPayments.StatusShortName, StatusPayments.StatusName From Payments, EmployeesHistoryListForPayroll, Areas As PaymentCenters, Companies, StatusPayments Where (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payments.PaymentDate=EmployeesHistoryListForPayroll.PayrollID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.CompanyID=Companies.CompanyID) And (Payments.StatusID=StatusPayments.StatusID) And (Companies.StartDate<=EmployeesHistoryListForPayroll.PayrollID) And (Companies.EndDate>=EmployeesHistoryListForPayroll.PayrollID) And (BankAccounts.Active = 1) " & sCondition & " Order By PaymentDate Desc, CheckDate Desc", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select Payments.PaymentID, PaymentDate, CheckDate, CancelDate, EmployeeNumber, CheckNumber, LastUpdate, ReplacementNumber, CheckAmount, StatusPayments.StatusID, StatusPayments.StatusShortName, StatusPayments.StatusName From Payments, EmployeesHistoryListForPayroll, Areas As PaymentCenters, Companies, StatusPayments Where (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payments.PaymentDate=EmployeesHistoryListForPayroll.PayrollID) And (EmployeesHistoryListForPayroll.PaymentCenterID=PaymentCenters.AreaID) And (PaymentCenters.CompanyID=Companies.CompanyID) And (Payments.StatusID=StatusPayments.StatusID) And (Companies.StartDate<=EmployeesHistoryListForPayroll.PayrollID) And (Companies.EndDate>=EmployeesHistoryListForPayroll.PayrollID) And (BankAccounts.Active = 1) " & sCondition & " Order By PaymentDate Desc, CheckDate Desc -->" & vbNewLine
	End If
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			sComboHTML = "<SELECT NAME=""StatusID_<PAYMENT_ID />"" ID=""StatusID_<PAYMENT_ID />Cmb"" SIZE=""1"" CLASS=""Lists"">"
				If iPaymentType = 1 Then
					sComboHTML = sComboHTML & GenerateListOptionsFromQuery(oADODBConnection, "StatusPayments", "StatusID", "StatusShortName, StatusName", "(StatusID In (1,4))", "StatusShortName", "", "Ninguno;;;-1", sErrorDescription)
				Else
					sComboHTML = sComboHTML & GenerateListOptionsFromQuery(oADODBConnection, "StatusPayments", "StatusID", "StatusShortName, StatusName", "(StatusID Not In (-2,-1,2,3,4))", "StatusShortName", "", "Ninguno;;;-1", sErrorDescription)
				End If
			sComboHTML = sComboHTML & "</SELECT>"
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				asColumnsTitles = Split("Fecha de nómina,Fecha de pago,Empleado,No. de cheque/depósito,Fecha de reexpedición,No. de reexpedición,Monto,Quincena de cancelación,Motivo de cancelación", ",", -1, vbBinaryCompare)
				asCellWidths = Split("100,100,100,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
				If bForExport Then
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split(",,,,,,RIGHT,,,", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sRowContents = DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PaymentDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("CheckDate").Value), -1, -1, -1)
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("CheckNumber").Value))
					sTemp = ""
'					sTemp = CStr(oRecordset.Fields("ReplacementNumber").Value)
					sTemp = Replace(CStr(oRecordset.Fields("ReplacementNumber").Value)," ","",vbTextCompare)
					Err.Clear
					If Len(sTemp) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("LastUpdate").Value), -1, -1, -1)
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(sTemp)
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>---</CENTER>"
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>---</CENTER>"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & FormatNumber(CDbl(oRecordset.Fields("CheckAmount").Value), 2, True, False, True)
					If CLng(oRecordset.Fields("CancelDate").Value) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("CancelDate").Value), -1, -1, -1)
						sRowContents = sRowContents & "<INPUT TYPE=""HIDDEN"" NAME=""Cancelled_" & CStr(oRecordset.Fields("PaymentID").Value) & """ ID=""Cancelled_" & CStr(oRecordset.Fields("PaymentID").Value) & "Hdn"" VALUE=""1"" />"
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>---</CENTER>"
					End If
					If bForExport Or (StrComp(GetASPFileName(""), "Employees.asp", vbBinaryCompare) = 0) Then
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StatusShortName").Value) & ". " & CStr(oRecordset.Fields("StatusName").Value))
					Else
						If (Len(sTemp) > 0) Or (InStr(1, ",-1,2,3,", "," & CStr(oRecordset.Fields("StatusID").Value) & ",", vbBinaryCompare) > 0) Then
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StatusShortName").Value) & ". " & CStr(oRecordset.Fields("StatusName").Value))
						ElseIf InStr(1, sClosed, "," & CStr(oRecordset.Fields("CancelDate").Value) & ",", vbBinaryCompare) > 0 Then
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("StatusShortName").Value) & ". " & CStr(oRecordset.Fields("StatusName").Value))
						Else
							sRowContents = sRowContents & TABLE_SEPARATOR & Replace(Replace(sComboHTML, "<PAYMENT_ID />", CStr(oRecordset.Fields("PaymentID").Value)), "VALUE=""" & CStr(oRecordset.Fields("StatusID").Value) & """", "VALUE=""" & CStr(oRecordset.Fields("StatusID").Value) & """ SELECTED=""1""")
						End If
					End If
					'If CLng(oRecordset.Fields("CancelDate").Value) = 0 Then                    ' Columna de Cancelación aplicada eliminada
					'	sRowContents = sRowContents & TABLE_SEPARATOR & "<CENTER>---</CENTER>"
					'ElseIf InStr(1, sClosed, "," & CStr(oRecordset.Fields("CancelDate").Value) & ",", vbBinaryCompare) > 0 Then
					'	sRowContents = sRowContents & TABLE_SEPARATOR & "Sí"
					'Else
					'	sRowContents = sRowContents & TABLE_SEPARATOR & "No"
					'End If
					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If

					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE><BR />"
			oRecordset.Close
		Else
			lErrorNumber = -1
			sErrorDescription = "No existen pagos que cumplan con los criterios de la búsqueda."
		End If
	End If

	Set oRecordset = Nothing
	DisplayEmployeePaymentsTable = lErrorNumber
	Err.Clear
End Function

Function DisplayEmployeesBlockPaymentsTable(oRequest, oADODBConnection, bForExport, aPaymentComponent, sErrorDescription)
'************************************************************
'Purpose: To display the absences for the given absence for
'		  the employee from the database in a table
'Inputs:  oRequest, oADODBConnection, bForExport, aAbsenceComponent
'Outputs: aAbsenceComponent, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayEmployeesBlockPaymentsTable"
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
	Dim lDate
	Dim sConcept

	lDate = CLng(Left(GetSerialNumberForDate(""), Len("00000000")))
	lErrorNumber = GetEmployeesBlockPayments(oRequest, oADODBConnection, aPaymentComponent, oRecordset, sErrorDescription)
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
				If (Not bForExport) Then
					asColumnsTitles = Split("Acciones,N. Empleado,Nombre,Nómina bloqueada", ",", -1, vbBinaryCompare)
					asCellWidths = Split("200,100,400,200,", ",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,CENTER,CENTER,", ",", -1, vbBinaryCompare)
				Else
					asColumnsTitles = Split("N. Empleado,Nombre,Nómina bloqueada", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,400,200,", ",", -1, vbBinaryCompare)
					asCellAlignments = Split("CENTER,CENTER,", ",", -1, vbBinaryCompare)
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
					If (StrComp(CStr(oRecordset.Fields("EmployeeID").Value), oRequest("EmployeeID").Item, vbBinaryCompare) = 0) Then
						sBoldBegin = "<B>"
						sBoldEnd = "</B>"
					End If
					sFontBegin = ""
					sFontEnd = ""
					'If CInt(oRecordset.Fields("Removed").Value) = 1 Then
					'	sFontBegin = "<FONT COLOR=""#" & S_INACTIVE_TEXT_FOR_GUI & """>"
					'	sFontEnd = "</FONT>"
					'End If
					sRowContents = ""
					If (Not bForExport) And ((aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_MODIFY_PERMISSIONS) = N_MODIFY_PERMISSIONS Or (aLoginComponent(N_USER_PERMISSIONS_LOGIN) And N_REMOVE_PERMISSIONS) = N_REMOVE_PERMISSIONS) Then
						sRowContents = sRowContents & "&nbsp;"
						sRowContents = sRowContents & "<A HREF=""" & GetASPFileName("") & "?Action=BlockPayments&EmployeeID=" & CStr(oRecordset.Fields("EmployeeID").Value) & "&PayrollID=" & CStr(oRecordset.Fields("PayrollID").Value) & "&BlockEmployees=1&RemoveBlockPayments=1"">"
							sRowContents = sRowContents & "<IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar bloqueo"" BORDER=""0"" />"
						sRowContents = sRowContents & "</A>"
					End If
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeID").Value)) & sBoldEnd & sFontEnd
					sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & sBoldBegin & CleanStringForHTML(CStr(oRecordset.Fields("EmployeeFullName").Value)) & sBoldEnd & sFontEnd
					Call GetNameFromTable(oADODBConnection, "Payrolls", CStr(oRecordset.Fields("PayrollID").Value), "", "", sNames, sErrorDescription)
					If Len(sNames) > 0 Then
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & CleanStringForHTML(sNames) & sBoldEnd & sFontEnd
					Else
						sRowContents = sRowContents & TABLE_SEPARATOR & sFontBegin & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("AppliedDate").Value), -1, -1, -1) & sBoldEnd & sFontEnd
					End If
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
			sErrorDescription = "No se han registrado bloqueo anticipado de pagos."
		End If
	End If

	oRecordset.Close
	Set oRecordset = Nothing
	DisplayEmployeesBlockPaymentsTable = lErrorNumber
	Err.Clear
End Function

Function DisplayPaymentsMessages(oRequest, oADODBConnection, lPayrollID, bForExport, sErrorDescription)
'************************************************************
'Purpose: To display the messags for the payments
'		  records from the database in a table
'Inputs:  oRequest, oADODBConnection, lPayrollID, bForExport
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPaymentsMessages"
	Dim sNames
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber

	sErrorDescription = "No se pudieron obtener los registros de las asignaciones de folios para los pagos."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From PaymentsMessages Where (PayrollID=" & lPayrollID & ") And (bSpecial=0) Order By RecordID", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		If Not oRecordset.EOF Then
			Response.Write "<TABLE BORDER="""
				If Not bForExport Then
					Response.Write "0"
				Else
					Response.Write "1"
				End If
			Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
				If bForExport Then
					asColumnsTitles = Split("Mensaje,Condición", ",", -1, vbBinaryCompare)
					asCellWidths = Split("200,200", ",", -1, vbBinaryCompare)
					lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
				Else
					asColumnsTitles = Split("Mensaje,Condición,Acciones", ",", -1, vbBinaryCompare)
					asCellWidths = Split("200,200,100", ",", -1, vbBinaryCompare)
					If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
						lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					Else
						lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
					End If
				End If

				asCellAlignments = Split(",,CENTER", ",", -1, vbBinaryCompare)
				Do While Not oRecordset.EOF
					sRowContents = CleanStringForHTML(CStr(oRecordset.Fields("Comments").Value))
					sRowContents = sRowContents & TABLE_SEPARATOR
						If CLng(oRecordset.Fields("EmployeeID").Value) > 0 Then
							sRowContents = sRowContents & "<B>Empleado: </B>" & CleanStringForHTML(Right(("000000" & CStr(oRecordset.Fields("EmployeeID").Value)), Len("000000"))) & "<BR />"
						End If
						If CLng(oRecordset.Fields("CompanyID").Value) > -1 Then
							Call GetNameFromTable(oADODBConnection, "Companies", CLng(oRecordset.Fields("CompanyID").Value), "", ", ", sNames, sErrorDescription)
							sRowContents = sRowContents & "<B>Compañías: </B>" & CleanStringForHTML(sNames) & "<BR />"
						End If
						If StrComp(CStr(oRecordset.Fields("AreaIDs").Value), "-1", vbBinaryCompare) <> 0 Then
							Call GetNameFromTable(oADODBConnection, "Areas", Replace(CStr(oRecordset.Fields("AreaIDs").Value), " ", ""), "", ", ", sNames, sErrorDescription)
							sRowContents = sRowContents & "<B>Unidades administrativas: </B>" & CleanStringForHTML(sNames) & "<BR />"
						End If
						If StrComp(CStr(oRecordset.Fields("ZoneIDs").Value), "-1", vbBinaryCompare) <> 0 Then
							Call GetNameFromTable(oADODBConnection, "Zones", Replace(CStr(oRecordset.Fields("ZoneIDs").Value), " ", ""), "", ", ", sNames, sErrorDescription)
							sRowContents = sRowContents & "<B>Entidades: </B>" & CleanStringForHTML(sNames) & "<BR />"
						End If
						If CLng(oRecordset.Fields("EmployeeTypeID").Value) > -1 Then
							Call GetNameFromTable(oADODBConnection, "EmployeeTypes", CLng(oRecordset.Fields("EmployeeTypeID").Value), "", ", ", sNames, sErrorDescription)
							sRowContents = sRowContents & "<B>Tipos de tabulador: </B>" & CleanStringForHTML(sNames) & "<BR />"
						End If
						If CLng(oRecordset.Fields("PositionID").Value) > -1 Then
							Call GetNameFromTable(oADODBConnection, "Positions", CLng(oRecordset.Fields("PositionID").Value), "", ", ", sNames, sErrorDescription)
							sRowContents = sRowContents & "<B>Puestos: </B>" & CleanStringForHTML(sNames) & "<BR />"
						End If
						If CLng(oRecordset.Fields("BankID").Value) > -1 Then
							Call GetNameFromTable(oADODBConnection, "Banks", CLng(oRecordset.Fields("BankID").Value), "", ", ", sNames, sErrorDescription)
							sRowContents = sRowContents & "<B>Bancos: </B>" & CleanStringForHTML(sNames) & "<BR />"
						End If
						Select Case CLng(oRecordset.Fields("ConceptID").Value)
							Case 0
								sRowContents = sRowContents & "<B>Tipo de pago: </B>Cheques<BR />"
							Case 1
								sRowContents = sRowContents & "<B>Tipo de pago: </B>Depósitos<BR />"
							Case 2
								sRowContents = sRowContents & "<B>Tipo de pago: </B>Pensión alimenticia<BR />"
							Case 3
								sRowContents = sRowContents & "<B>Tipo de pago: </B>Honorarios<BR />"
							Case 4
								sRowContents = sRowContents & "<B>Tipo de pago: </B>Acreedores<BR />"
						End Select
					If Not bForExport Then sRowContents = sRowContents & TABLE_SEPARATOR & "<A HREF=""Payments.asp?Action=PrintPayments&Step=2&RecordID=" & CStr(oRecordset.Fields("RecordID").Value) & "&PayrollID=" & lPayrollID & "&Remove=1""><IMG SRC=""Images/BtnRemove.gif"" WIDTH=""10"" HEIGHT=""8"" ALT=""Eliminar mensaje"" BORDER=""0"" /></A>"

					asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
					Else
						lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
					End If

					oRecordset.MoveNext
					If Err.number <> 0 Then Exit Do
				Loop
			Response.Write "</TABLE><BR />"
			oRecordset.Close
		Else
			Response.Write "<FONT FACE=""Arial"" SIZE=""2""><B>No existen mensajes registrados para los pagos de esta quincena.</B></FONT>"
		End If
	End If

	Set oRecordset = Nothing
	DisplayPaymentsMessages = lErrorNumber
	Err.Clear
End Function

Function DisplayPaymentsMarginsSettings(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To display the information about the payments
'		  records from the database in a table
'Inputs:  oRequest, oADODBConnection, bForExport, sAction, lPayrollID
'Outputs: bAllPrinted, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPaymentsMarginsSettings"
	Dim asCompanies
	Dim asAreas
	Dim asEmployeeTypes
	Dim iIndex
	Dim sCondition
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim asAreaIDs
	Dim bDisplay
	Dim bEmpty
	Dim lErrorNumber

	If Len(oRequest("PosX1").Item) > 0 Then
		If IsNumeric(oRequest("PosX1").Item) Then Call SetOption(aOptionsComponent, CHECKS_LEFT_MARGIN1_OPTION, CLng(oRequest("PosX1").Item), sErrorDescription)
	End If
	If Len(oRequest("PosY1").Item) > 0 Then
		If IsNumeric(oRequest("PosY1").Item) Then Call SetOption(aOptionsComponent, CHECKS_TOP_MARGIN1_OPTION, CLng(oRequest("PosY1").Item), sErrorDescription)
	End If
	If Len(oRequest("PosX2").Item) > 0 Then
		If IsNumeric(oRequest("PosX2").Item) Then Call SetOption(aOptionsComponent, CHECKS_LEFT_MARGIN2_OPTION, CLng(oRequest("PosX2").Item), sErrorDescription)
	End If
	If Len(oRequest("PosY2").Item) > 0 Then
		If IsNumeric(oRequest("PosY2").Item) Then Call SetOption(aOptionsComponent, CHECKS_TOP_MARGIN2_OPTION, CLng(oRequest("PosY2").Item), sErrorDescription)
	End If
	Response.Write "<FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;Si se presenta un descuadre en la impresión del cheque, puede ajustar un desplazamiento extra que se aplicara a todo el documento, los cuales quedarán predefinidos para posteriores impresiones.&nbsp;</FONT>"
	Response.Write "<BR /><BR />"
	Response.Write "<FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;<B>Nota:&nbsp;</B>La posición base para la plantilla oficial es con valores 0 para todos los desplazamientos.&nbsp;</FONT>"
	Response.Write "<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0"">"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>&nbsp;&nbsp;&nbsp;Establezca el valor en mm. para&nbsp;&nbsp;&nbsp;</B></FONT></TD>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Desplazamiento izquierdo:&nbsp;</FONT>"
			Response.Write "<INPUT TYPE=""TEXT"" NAME=""PosX1"" ID=""PosX1Txt"" VALUE=""" & GetOption(aOptionsComponent, CHECKS_LEFT_MARGIN1_OPTION) & """ SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" /></TD>&nbsp;&nbsp;&nbsp;"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;Desplazamiento superior:&nbsp;</FONT>"
			Response.Write "<INPUT TYPE=""TEXT"" NAME=""PosY1"" ID=""PosY1Txt"" VALUE=""" & GetOption(aOptionsComponent, CHECKS_TOP_MARGIN1_OPTION) & """ SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" /></TD>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>&nbsp;&nbsp;&nbsp;de la sección superior del formato de impresion del cheque</B></FONT></TD>"
		Response.Write "</TR>"
		Response.Write "<TR>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>&nbsp;&nbsp;&nbsp;Establezca el valor en mm. para&nbsp;&nbsp;&nbsp;</B></FONT></TD>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">Desplazamiento izquierdo:&nbsp;</FONT>"
			Response.Write "<INPUT TYPE=""TEXT"" NAME=""PosX2"" ID=""PosX2Txt"" VALUE=""" & GetOption(aOptionsComponent, CHECKS_LEFT_MARGIN2_OPTION) & """ SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" /></TD>&nbsp;&nbsp;&nbsp;"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2"">&nbsp;&nbsp;&nbsp;Desplazamiento superior:&nbsp;</FONT>"
			Response.Write "<INPUT TYPE=""TEXT"" NAME=""PosY2"" ID=""PosY2Txt"" VALUE=""" & GetOption(aOptionsComponent, CHECKS_TOP_MARGIN2_OPTION) & """ SIZE=""6"" MAXLENGTH=""6"" CLASS=""TextFields"" /></TD>"
			Response.Write "<TD><FONT FACE=""Arial"" SIZE=""2""><B>&nbsp;&nbsp;&nbsp;de la sección desprendible del formato de impresion del cheque</B></FONT></TD>"
		Response.Write "</TR>"
	Response.Write "</TABLE>"
	Response.Write "<BR /><BR />"
	Set oRecordset = Nothing
	DisplayPaymentsMarginsSettings = lErrorNumber
	Err.Clear
End Function

Function DisplayPaymentRecordsTable(oRequest, oADODBConnection, sAction, lPayrollID, bForExport, bAllPrinted, sErrorDescription)
'************************************************************
'Purpose: To display the information about the payments
'		  records from the database in a table
'Inputs:  oRequest, oADODBConnection, bForExport, sAction, lPayrollID
'Outputs: bAllPrinted, sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "DisplayPaymentRecordsTable"
	Dim asCompanies
    Dim asZones
	Dim asAreas
	Dim asEmployeeTypes
	Dim iIndex
	Dim sCondition
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim asAreaIDs
    Dim asZonesIDs
	Dim bDisplay
	Dim bEmpty
	Dim asTemp
	Dim lErrorNumber

	bAllPrinted = True
	bEmpty = True
	asCompanies = ""
	sErrorDescription = "No se pudieron obtener los registros del catálogo."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CompanyID, CompanyShortName, CompanyName, StartDate, EndDate From Companies Where (CompanyID>-1) And (ParentID>-1) Order By CompanyID, StartDate", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			asCompanies = asCompanies & CStr(oRecordset.Fields("CompanyID").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("CompanyShortName").Value) & ". " & CStr(oRecordset.Fields("CompanyName").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("StartDate").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EndDate").Value) & LIST_SEPARATOR
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		oRecordset.Close
		asCompanies = Left(asCompanies, (Len(asCompanies) - Len(LIST_SEPARATOR)))
		asCompanies = Split(asCompanies, LIST_SEPARATOR)
		For iIndex = 0 To UBound(asCompanies)
			asCompanies(iIndex) = Split(asCompanies(iIndex), SECOND_LIST_SEPARATOR)
		Next
	End If

	asZones = ""
    sErrorDescription = "No se pudieron obtener los registros del catálogo."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select ZoneID, ZoneCode, ZoneName, StartDate, EndDate From Zones Where (ZoneID>-1) And (ParentID=-1) Order By ZoneID, StartDate", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			asZones = asZones & CStr(oRecordset.Fields("ZoneID").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("ZoneCode").Value) & ". " & CStr(oRecordset.Fields("ZoneName").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("StartDate").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EndDate").Value) & LIST_SEPARATOR
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		oRecordset.Close
		asZones = Left(asZones, (Len(asZones) - Len(LIST_SEPARATOR)))
		asZones = Split(asZones, LIST_SEPARATOR)
		For iIndex = 0 To UBound(asZones)
			asZones(iIndex) = Split(asZones(iIndex), SECOND_LIST_SEPARATOR)
		Next
	End If

	asAreas = ""
	sErrorDescription = "No se pudieron obtener los registros del catálogo."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select AreaID, AreaCode, AreaName, StartDate, EndDate From Areas Where (AreaID>-1) And (ParentID=-1) Order By AreaID, StartDate", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			asAreas = asAreas & CStr(oRecordset.Fields("AreaID").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("AreaCode").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("AreaName").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("StartDate").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EndDate").Value) & LIST_SEPARATOR
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		oRecordset.Close
		asAreas = Left(asAreas, (Len(asAreas) - Len(LIST_SEPARATOR)))
		asAreas = Split(asAreas, LIST_SEPARATOR)
		For iIndex = 0 To UBound(asAreas)
			asAreas(iIndex) = Split(asAreas(iIndex), SECOND_LIST_SEPARATOR)
		Next
	End If

	asEmployeeTypes = ""
	sErrorDescription = "No se pudieron obtener los registros del catálogo."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeTypeID, EmployeeTypeShortName, EmployeeTypeName, StartDate, EndDate From EmployeeTypes Where (EmployeeTypeID>-1) Order By EmployeeTypeID, StartDate", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Do While Not oRecordset.EOF
			asEmployeeTypes = asEmployeeTypes & CStr(oRecordset.Fields("EmployeeTypeID").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EmployeeTypeShortName").Value) & ". " & CStr(oRecordset.Fields("EmployeeTypeName").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("StartDate").Value) & SECOND_LIST_SEPARATOR & CStr(oRecordset.Fields("EndDate").Value) & LIST_SEPARATOR
			oRecordset.MoveNext
			If Err.number <> 0 Then Exit Do
		Loop
		oRecordset.Close
		asEmployeeTypes = Left(asEmployeeTypes, (Len(asEmployeeTypes) - Len(LIST_SEPARATOR)))
		asEmployeeTypes = Split(asEmployeeTypes, LIST_SEPARATOR)
		For iIndex = 0 To UBound(asEmployeeTypes)
			asEmployeeTypes(iIndex) = Split(asEmployeeTypes(iIndex), SECOND_LIST_SEPARATOR)
		Next
	End If

	If lErrorNumber = 0 Then
		If (StrComp(sAction, "PrintPayments", vbBinaryCompare) = 0) Or (StrComp(sAction, "RemovePaymentsRecords", vbBinaryCompare) = 0) Then
			sCondition = " And (PayrollID=" & lPayrollID & ")"
		Else
			sCondition = " And (bPrinted<>2) "
		End If
		Response.Write "<SCRIPT LANGUAGE=""JavaScript""><!--" & vbNewLine
			Response.Write "var aPaymentsRecords = new Array("
				Call GenerateJavaScriptArrayFromQuery(oADODBConnection, "PaymentsRecords", "RecordID", "FirstNumber, LastNumber, ReexpeditionNumber, EndNumber", "(RecordID>-1)" & sCondition, "RecordID", sErrorDescription)
			Response.Write "['-1', '-1', '-1', '-1', '-1']);" & vbNewLine

			Response.Write "function GetRecordNumbers(sRecordID) {" & vbNewLine
				Response.Write "var oForm = document.PrintFrm;" & vbNewLine

				Response.Write "if (oForm) {" & vbNewLine
					Response.Write "oForm.FilterFirstNumber.value = '';" & vbNewLine
					Response.Write "oForm.FilterLastNumber.value = '';" & vbNewLine
					Response.Write "for (var i=0; i<aPaymentsRecords.length; i++)" & vbNewLine
						Response.Write "if (aPaymentsRecords[i][0] == sRecordID) {" & vbNewLine
                        	Response.Write "if (aPaymentsRecords[i][3] == -1) {" & vbNewLine
								Response.Write "oForm.FilterFirstNumber.value = aPaymentsRecords[i][1];" & vbNewLine
								Response.Write "oForm.FilterLastNumber.value = aPaymentsRecords[i][2];" & vbNewLine
                        	Response.Write "} else {" & vbNewLine
								Response.Write "oForm.FilterFirstNumber.value = aPaymentsRecords[i][3];" & vbNewLine
								Response.Write "oForm.FilterLastNumber.value = aPaymentsRecords[i][4];" & vbNewLine
                        	Response.Write "}" & vbNewLine
						Response.Write "}" & vbNewLine
				Response.Write "}" & vbNewLine
			Response.Write "} // End of GetRecordNumbers" & vbNewLine
		Response.Write "//--></SCRIPT>" & vbNewLine

		sErrorDescription = "No se pudieron obtener los registros de las asignaciones de folios para los pagos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RecordID, PayrollID, CompanyIDs, ZoneIDs, AreaIDs, EmployeeTypeIDs, BankName, PaymentsRecords.EmployeeID, FirstNumber, LastNumber, ReexpeditionNumber, EndNumber, ConceptID, bPrinted From PaymentsRecords, Banks Where (PaymentsRecords.BankID=Banks.BankID) " & sCondition & " Order By RecordID", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select RecordID, PayrollID, CompanyIDs, ZoneIDs, AreaIDs, EmployeeTypeIDs, BankName, PaymentsRecords.EmployeeID, FirstNumber, LastNumber, ReexpeditionNumber, EndNumber, ConceptID, bPrinted From PaymentsRecords, Banks Where (PaymentsRecords.BankID=Banks.BankID) " & sCondition & " Order By RecordID -->" & vbNewLine
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				bEmpty = False
				Response.Write "<TABLE WIDTH=""1100"" BORDER="""
					If Not bForExport Then
						Response.Write "0"
					Else
						Response.Write "1"
					End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
					asColumnsTitles = Split("&nbsp;,Tipo,Fecha de pago,Compañía,Entidad,Área,Tipo de tabulador,Banco,No. empleado,Primer número de folio,Último número de folio,Inicio folio reposición,Fin folio reposición", ",", -1, vbBinaryCompare)
					asCellWidths = Split("100,100,100,100,200,200,100,100,100,100,100,100,100", ",", -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
					Else
						If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
							lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						Else
							lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						End If
					End If

					asCellAlignments = Split(",,,,,,,RIGHT,RIGHT", ",", -1, vbBinaryCompare)
					Do While Not oRecordset.EOF
						bDisplay = True
						If StrComp(aLoginComponent(N_PERMISSION_AREA_ID_LOGIN), "-1", vbBinaryCompare) <> 0 Then
							'If CStr(oRecordset.Fields("AreaIDs").Value) <> "0" Then
                                asAreaIDs = Split(Replace(CStr(oRecordset.Fields("ZoneIDs").Value), " ", ""), ",")
							    For iIndex = 0 To UBound(asAreaIDs)
								    If InStr(1, ("," & aLoginComponent(N_PERMISSION_AREA_ID_LOGIN) & ","), ("," & asAreaIDs(iIndex) & ","), vbBinaryCompare) = 0 Then
									    bDisplay = False
									    Exit For
									End If
								Next
							'End If
						End If
						sRowContents = ""
						If bDisplay Then
							If (StrComp(sAction, "PrintPayments", vbBinaryCompare) = 0) Then
								bAllPrinted = False
								sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""RecordID"" ID=""RecordIDRd"" VALUE=""" & CStr(oRecordset.Fields("RecordID").Value) & """ onClick=""iPrintCounter = 1; GetRecordNumbers(this.value);"" />&nbsp;"
								If FileExists(Server.MapPath("Reports/Rep_1400_" & CStr(oRecordset.Fields("RecordID").Value) & ".zip"), sErrorDescription) Then
									sRowContents = sRowContents & "<A HREF=""Reports/Rep_1400_" & CStr(oRecordset.Fields("RecordID").Value) & ".zip"" TARGET=""_blank""><IMG SRC=""Images/IcnFileZIP.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Revisar pagos previamente impresos"" BORDER=""0"" /></A>&nbsp;"
									If CInt(oRecordset.Fields("bPrinted").Value) = 1 Then
										sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""RecordToClose"" ID=""RecordToCloseChk"" VALUE=""" & CStr(oRecordset.Fields("RecordID").Value) & """ onClick=""if (this.checked) {iStatusCounter++;} else {iStatusCounter--;}"" />"
									Else
										sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEGITH=""10"" ALT=""Cheques impresos"" />"
									End If
								End If
							Else
								If (StrComp(sAction, "RemovePaymentsRecords", vbBinaryCompare) = 0) Then
									If (CInt(oRecordset.Fields("bPrinted").Value) <> 2) Then
										sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""RecordID"" ID=""RecordIDRd"" VALUE=""" & CStr(oRecordset.Fields("RecordID").Value) & """ onClick=""iPrintCounter = 1; GetRecordNumbers(this.value);"" />&nbsp;"
										bAllPrinted = False
									End If
								Else
									sRowContents = sRowContents & "&nbsp;"
								End If
							End If
							sRowContents = sRowContents & TABLE_SEPARATOR 
							Select Case CInt(oRecordset.Fields("ConceptID").Value)
								Case 0
									sRowContents = sRowContents & "Cheques"
								Case 1
									sRowContents = sRowContents & "Depósitos"
								Case 2
									sRowContents = sRowContents & "Pensión alimenticia"
								Case 3
									sRowContents = sRowContents & "Honorarios"
								Case 4
									sRowContents = sRowContents & "Acreedores"
							End Select
							sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PayrollID").Value), -1, -1, -1)
							sRowContents = sRowContents & TABLE_SEPARATOR
								For iIndex = 0 To UBound(asCompanies)
									If InStr(1, ("," & Replace(CStr(oRecordset.Fields("CompanyIDs").Value), " ", "") & ","), ("," & asCompanies(iIndex)(0) & ","), vbBinaryCompare) > 0 Then
										If (CLng(asCompanies(iIndex)(2)) <= CLng(oRecordset.Fields("PayrollID").Value)) And (CLng(asCompanies(iIndex)(3)) >= CLng(oRecordset.Fields("PayrollID").Value)) Then
											sRowContents = sRowContents & CleanStringForHTML(asCompanies(iIndex)(1)) & "<BR />"
										End If
									End If
								Next
							sRowContents = sRowContents & TABLE_SEPARATOR
								For iIndex = 0 To UBound(asZones)
									If InStr(1, ("," & Replace(CStr(oRecordset.Fields("ZoneIDs").Value), " ", "") & ","), ("," & asZones(iIndex)(0) & ","), vbBinaryCompare) > 0 Then
										If (CLng(asZones(iIndex)(2)) <= CLng(oRecordset.Fields("PayrollID").Value)) And (CLng(asZones(iIndex)(3)) >= CLng(oRecordset.Fields("PayrollID").Value)) Then
											sRowContents = sRowContents & CleanStringForHTML(asZones(iIndex)(1)) & "<BR />"
										End If
									End If
								Next
							sRowContents = sRowContents & TABLE_SEPARATOR & "<FONT TITLE="""
								For iIndex = 0 To UBound(asAreas)
									If InStr(1, ("," & Replace(CStr(oRecordset.Fields("AreaIDs").Value), " ", "") & ","), ("," & asAreas(iIndex)(0) & ","), vbBinaryCompare) > 0 Then
										If (CLng(asAreas(iIndex)(3)) <= CLng(oRecordset.Fields("PayrollID").Value)) And (CLng(asAreas(iIndex)(4)) >= CLng(oRecordset.Fields("PayrollID").Value)) Then
											sRowContents = sRowContents & CleanStringForHTML(asAreas(iIndex)(1) & ". " & asAreas(iIndex)(2)) & "&#13;"
										End If
									End If
								Next
							sRowContents = sRowContents & """>"
								For iIndex = 0 To UBound(asAreas)
									If InStr(1, ("," & Replace(CStr(oRecordset.Fields("AreaIDs").Value), " ", "") & ","), ("," & asAreas(iIndex)(0) & ","), vbBinaryCompare) > 0 Then
										If (CLng(asAreas(iIndex)(3)) <= CLng(oRecordset.Fields("PayrollID").Value)) And (CLng(asAreas(iIndex)(4)) >= CLng(oRecordset.Fields("PayrollID").Value)) Then
											sRowContents = sRowContents & CleanStringForHTML(asAreas(iIndex)(1)) & ", "
										End If
									End If
								Next
								sRowContents = Left(sRowContents, (Len(sRowContents) - Len(", ")))
							sRowContents = sRowContents & "</FONT>"
							sRowContents = sRowContents & TABLE_SEPARATOR
								For iIndex = 0 To UBound(asEmployeeTypes)
									If InStr(1, ("," & Replace(CStr(oRecordset.Fields("EmployeeTypeIDs").Value), " ", "") & ","), ("," & asEmployeeTypes(iIndex)(0) & ","), vbBinaryCompare) > 0 Then
										If (CLng(asEmployeeTypes(iIndex)(2)) <= CLng(oRecordset.Fields("PayrollID").Value)) And (CLng(asEmployeeTypes(iIndex)(3)) >= CLng(oRecordset.Fields("PayrollID").Value)) Then
											sRowContents = sRowContents & CleanStringForHTML(asEmployeeTypes(iIndex)(1)) & "<BR />"
										End If
									End If
								Next
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BankName").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR
								If CLng(oRecordset.Fields("EmployeeID").Value) > -1 Then
									sRowContents = sRowContents & Right(("000000" & CStr(oRecordset.Fields("EmployeeID").Value)), Len("000000"))
								Else
									sRowContents = sRowContents & "<CENTER>---</CENTER>"
								End If
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("FirstNumber").Value))
							sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("LastNumber").Value))
                            If CLng(oRecordset.Fields("ReexpeditionNumber").Value) = -1 Then
							    sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("NA")
                            Else
                                sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ReexpeditionNumber").Value))
                            End If
                            If CLng(oRecordset.Fields("EndNumber").Value) = -1 Then
							    sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML("NA")
                            Else
                                sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("EndNumber").Value))
                            End If

							asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
							If bForExport Then
								lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
							Else
								lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
							End If
						End If

						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
				Response.Write "</TABLE><BR />"
				oRecordset.Close
				Response.Write "&nbsp;&nbsp;&nbsp;<B>Primer número de folio: </B><INPUT TYPE=""TEXT"" NAME=""FilterFirstNumber"" ID=""FirstNumberTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE="""" CLASS=""TextFields"" /><BR />"
				Response.Write "&nbsp;&nbsp;&nbsp;<B>Último número de folio: </B><INPUT TYPE=""TEXT"" NAME=""FilterLastNumber"" ID=""LastNumberTxt"" SIZE=""20"" MAXLENGTH=""20"" VALUE="""" CLASS=""TextFields"" /><BR />"
				Response.Write "<BR />"
			End If
		End If
	End If

	If lErrorNumber = 0 Then
		If StrComp(sAction, "PrintPayments", vbBinaryCompare) = 0 Then
			sCondition = " And (PayrollID=" & lPayrollID & ") And (ConceptID=0) "
		Else
			sCondition = " And (bPrinted<>2) "
		End If
		sErrorDescription = "No se pudieron obtener los registros de las asignaciones de folios para los pagos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select RecordID, PayrollID, PaymentsRecords2.EmployeeID, BankName, AccountNumber, ReplacementNumber, bPrinted From PaymentsRecords2, Banks, BankAccounts Where (PaymentsRecords2.BankID=Banks.BankID) And (PaymentsRecords2.AccountID=BankAccounts.AccountID) " & sCondition & " Order By RecordID", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select RecordID, PayrollID, PaymentsRecords2.EmployeeID, BankName, AccountNumber, ReplacementNumber, bPrinted From PaymentsRecords2, Banks, BankAccounts Where (PaymentsRecords2.BankID=Banks.BankID) And (PaymentsRecords2.AccountID=BankAccounts.AccountID) " & sCondition & " Order By RecordID -->" & vbNewLine
		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				bEmpty = False
				Response.Write "<TABLE BORDER="""
					If Not bForExport Then
						Response.Write "0"
					Else
						Response.Write "1"
					End If
				Response.Write """ CELLSPACING=""0"" CELLPADDING=""0"">" & vbNewLine
					asColumnsTitles = Split("&nbsp;,Fecha de pago,No. Empleado,Banco,Cuenta,No. reexpedición", ",", -1, vbBinaryCompare)
					asCellWidths = Split("20,100,100,100,100,100", ",", -1, vbBinaryCompare)
					If bForExport Then
						lErrorNumber = DisplayTableHeaderPlainText(asColumnsTitles, True, sErrorDescription)
					Else
						If CInt(GetOption(aOptionsComponent, TABLE_STYLE_OPTION)) = 2 Then
							lErrorNumber = DisplayTableHeaderPlain(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						Else
							lErrorNumber = DisplayTableHeader3D(asColumnsTitles, asCellWidths, asTableColors, sErrorDescription)
						End If
					End If

					asCellAlignments = Split(",,,,,", ",", -1, vbBinaryCompare)
					Do While Not oRecordset.EOF
						sRowContents = ""
						If CInt(oRecordset.Fields("bPrinted").Value) <> 2 Then
							sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""RecordID2"" ID=""RecordID2Rd"" VALUE=""" & CStr(oRecordset.Fields("RecordID").Value) & """ onClick=""iPrintCounter = 1;"" />&nbsp;"
							bAllPrinted = False
						End If
						If FileExists(Server.MapPath("Reports/Rep_1400_" & CStr(oRecordset.Fields("RecordID").Value) & ".zip"), sErrorDescription) Then
							If StrComp(sAction, "PrintPayments", vbBinaryCompare) = 0 Then
								sRowContents = sRowContents & "<A HREF=""Reports/Rep_1400_" & CStr(oRecordset.Fields("RecordID").Value) & ".zip"" TARGET=""_blank""><IMG SRC=""Images/IcnFileZIP.gif"" WIDTH=""16"" HEIGHT=""16"" ALT=""Revisar pagos previamente impresos"" BORDER=""0"" /></A>"
								If CInt(oRecordset.Fields("bPrinted").Value) = 1 Then
									sRowContents = sRowContents & "<INPUT TYPE=""CHECKBOX"" NAME=""RecordToClose2"" ID=""RecordToClose2Chk"" VALUE=""" & CStr(oRecordset.Fields("RecordID").Value) & """ onClick=""if (this.checked) {iStatusCounter++;} else {iStatusCounter--;}"" />"
								Else
									sRowContents = sRowContents & "<IMG SRC=""Images/IcnCheck.gif"" WIDTH=""10"" HEGITH=""10"" ALT=""Cheques impresos"" />"
								End If
							Else
								sRowContents = sRowContents & "<INPUT TYPE=""RADIO"" NAME=""RecordID2"" ID=""RecordID2Rd"" VALUE=""" & CStr(oRecordset.Fields("RecordID").Value) & """ onClick=""iPrintCounter = 1;"" />"
								bAllPrinted = False
							End If
						Else
							sRowContents = sRowContents & "&nbsp;"
						End If
						sRowContents = sRowContents & TABLE_SEPARATOR & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PayrollID").Value), -1, -1, -1)
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(Right(("000000" & CStr(oRecordset.Fields("EmployeeID").Value)), Len("000000")))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("BankName").Value))
						asTemp = Split(CStr(oRecordset.Fields("AccountNumber").Value), LIST_SEPARATOR)
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(asTemp(0))
						sRowContents = sRowContents & TABLE_SEPARATOR & CleanStringForHTML(CStr(oRecordset.Fields("ReplacementNumber").Value))

						asRowContents = Split(sRowContents, TABLE_SEPARATOR, -1, vbBinaryCompare)
						If bForExport Then
							lErrorNumber = DisplayTableRowText(asRowContents, True, sErrorDescription)
						Else
							lErrorNumber = DisplayTableRow(asRowContents, asCellAlignments, asCellWidths, "", "", "", "", sErrorDescription)
						End If

						oRecordset.MoveNext
						If Err.number <> 0 Then Exit Do
					Loop
				Response.Write "</TABLE><BR />"
				oRecordset.Close
			End If
		End If
	End If
	If bEmpty Then
		lErrorNumber = -1
		sErrorDescription = "No existen cheques que cumplan con los criterios de la búsqueda."
	End If

	Set oRecordset = Nothing
	DisplayPaymentRecordsTable = lErrorNumber
	Err.Clear
End Function

Function PrintPayments1(oRequest, oADODBConnection, lRecordID, sErrorDescription)
'************************************************************
'Purpose: To send the payments information to printing formats
'Inputs:  oRequest, oADODBConnection, lRecordID, aCatalogComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "PrintPayments"
	Const CHECKS_PER_FILE = 2000
	Dim S_CREDITS_ID
	Dim lEmployeeCounter
	Dim lFileCounter
	Dim sPerceptions
	Dim sDeductions
	Dim lStartPayrollDate
	Dim lStartDate
	Dim lMinDate
	Dim lMaxDate
	Dim sDate
	Dim sFilePath
	Dim sImagesPath
	Dim sFileName
	Dim sDocumentName
	Dim lReportID
	Dim sCurrentNumber
	Dim bAlimony
	Dim bCreditor
	Dim sCondition
	Dim asMessages
	Dim asEmployeesMessages
	Dim sContents
	Dim asPath
	Dim iIndex
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim bFirstEmployee
	Dim lPerceptions
	Dim lDeductions
	Dim lConceptAmount
	Dim yPositionForConceptsP
	Dim yPositionForConceptsD
	Dim sSignatureName1
	Dim sSignatureName2
	Dim sPensionISSSTELogo
	Dim x1PositionDisplacement
	Dim y1PositionDisplacement
	Dim x2PositionDisplacement
	Dim y2PositionDisplacement
	Dim l01PositionDisplacement
	Dim l02PositionDisplacement
	Dim l03PositionDisplacement
	Dim l04PositionDisplacement
	Dim l05PositionDisplacement
	Dim l03PositionDisplacementR
	Dim l04PositionDisplacementR
	Dim l05PositionDisplacementR
	Dim lConceptsTPositionDisplacement
	Dim lConceptsPositionDisplacement
	Dim lSignaturesPositionDisplacement
	Dim lConceptsTPositionDisplacementR
	Dim lConceptsPositionDisplacementR
	Dim lSignaturesPositionDisplacementR
	Dim oEndDate

	sDate = GetSerialNumberForDate("")

	y1PositionDisplacement = -100
	If Len(oRequest("PosX1").Item) > 0 Then x1PositionDisplacement = CInt(CInt(oRequest("PosX1").Item) * (56.692913386))
	If Len(oRequest("PosY1").Item) > 0 Then y1PositionDisplacement = y1PositionDisplacement + CInt(CInt(oRequest("PosY1").Item) * (56.692913386))
	If Len(oRequest("PosX2").Item) > 0 Then x2PositionDisplacement = CInt(CInt(oRequest("PosX2").Item) * (56.692913386))
	If Len(oRequest("PosY2").Item) > 0 Then y2PositionDisplacement = y2PositionDisplacement + CInt(CInt(oRequest("PosY2").Item) * (56.692913386))

	l01PositionDisplacement = 0
	l02PositionDisplacement = 0
	l03PositionDisplacement = 0
	l04PositionDisplacement = 0
	l05PositionDisplacement = 0
	l03PositionDisplacementR = 0
	l04PositionDisplacementR = 0
	l05PositionDisplacementR = 0
	lConceptsTPositionDisplacement = 0
	lConceptsPositionDisplacement = 0
	lSignaturesPositionDisplacement = 0
	lConceptsTPositionDisplacementR = 0
	lConceptsPositionDisplacementR = 0
	lSignaturesPositionDisplacementR = 0

	sPensionISSSTELogo = "PensionISSSTE.jpg"
	sFilePath = Server.MapPath(REPORTS_PATH & "Rep_1400_" & sDate)
	sImagesPath = Server.MapPath(TEMPLATES_PATH) & "\Images"
	sErrorDescription = "Error al crear la carpeta en donde se almacenará el reporte"
	lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	If lErrorNumber = 0 Then
		sFileName = REPORTS_PATH & "Rep_1400_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & ".zip"
		sDocumentName = sFilePath & "\Rep_1400_" & sDate & "_<INDEX />.rtf"
		sErrorDescription = "No se pudieron obtener las nóminas de los empleados."
		If FileExists(Server.MapPath(sFileName), sErrorDescription) Then Call DeleteFile(Server.MapPath(sFileName), sErrorDescription)

'lErrorNumber = AppendTextToFile(sFilePath & ".txt", "Inicio de armado de mensajes", sErrorDescription) 'TRACE
		S_CREDITS_ID = ","
		sErrorDescription = "No se pudieron obtener los IDs de los créditos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CreditTypeID From CreditTypes Where (CreditTypeID>0) And (Active=1) Order By CreditTypeID", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				S_CREDITS_ID = S_CREDITS_ID & CStr(oRecordset.Fields("CreditTypeID").Value) & ","
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
		End If

		asMessages = ""
		sErrorDescription = "No se pudieron obtener los mensajes para los cheques."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From PaymentsMessages Where (PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (bSpecial=0) Order By RecordID", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select * From PaymentsMessages Where (PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (bSpecial=0) Order By RecordID -->" & vbNewLine
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				asMessages = asMessages & Replace(CStr(oRecordset.Fields("Comments").Value), "<BR />", ("\" & Chr(13) & Chr(10))) & LIST_SEPARATOR
				asMessages = asMessages & "," & Replace(CStr(oRecordset.Fields("EmployeeID").Value), " ", "") & "," & LIST_SEPARATOR
				asMessages = asMessages & "," & Replace(CStr(oRecordset.Fields("CompanyID").Value), " ", "") & "," & LIST_SEPARATOR
				asMessages = asMessages & "," & Replace(CStr(oRecordset.Fields("AreaIDs").Value), " ", "") & "," & LIST_SEPARATOR
				asMessages = asMessages & "," & Replace(CStr(oRecordset.Fields("ZoneIDs").Value), " ", "") & "," & LIST_SEPARATOR
				asMessages = asMessages & "," & Replace(CStr(oRecordset.Fields("EmployeeTypeID").Value), " ", "") & "," & LIST_SEPARATOR
				asMessages = asMessages & "," & Replace(CStr(oRecordset.Fields("PositionID").Value), " ", "") & "," & LIST_SEPARATOR
				asMessages = asMessages & "," & Replace(CStr(oRecordset.Fields("BankID").Value), " ", "") & "," & LIST_SEPARATOR
				asMessages = asMessages & Replace(CStr(oRecordset.Fields("ConceptID").Value), " ", "") & SECOND_LIST_SEPARATOR
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			asMessages = Split(asMessages, SECOND_LIST_SEPARATOR)
			For iIndex = 0 To UBound(asMessages)
				asMessages(iIndex) = Split(asMessages(iIndex), LIST_SEPARATOR)
			Next
		End If
'lErrorNumber = AppendTextToFile(sFilePath & ".txt", "Fin de armado de mensajes", sErrorDescription) 'TRACE

'lErrorNumber = AppendTextToFile(sFilePath & ".txt", "Inicio de armado de mensajes para usuarios", sErrorDescription) 'TRACE
		asEmployeesMessages = ""
		sErrorDescription = "No se pudieron obtener los mensajes para los cheques."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, Comments From PaymentsMessages Where (PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (bSpecial Not In (0,3)) And (EmployeeID In (Select EmployeeID From Payments Where (PaymentID>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12) & ") And (PaymentID<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13) & "))) Order By EmployeeID, RecordID", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select EmployeeID, Comments From PaymentsMessages Where (PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (bSpecial Not In (0,3)) And (EmployeeID In (Select EmployeeID From Payments Where (PaymentID>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12) & ") And (PaymentID<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13) & "))) Order By EmployeeID, RecordID -->" & vbNewLine
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				asEmployeesMessages = asEmployeesMessages & Replace(CStr(oRecordset.Fields("Comments").Value), "<BR />", ("\" & Chr(13) & Chr(10))) & TABLE_SEPARATOR
				asEmployeesMessages = asEmployeesMessages & CStr(oRecordset.Fields("EmployeeID").Value) & CATALOG_SEPARATOR
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			asEmployeesMessages = Split(asEmployeesMessages, CATALOG_SEPARATOR)
			For iIndex = 0 To UBound(asEmployeesMessages)
				asEmployeesMessages(iIndex) = Split(asEmployeesMessages(iIndex), TABLE_SEPARATOR)
			Next
		End If
'lErrorNumber = AppendTextToFile(sFilePath & ".txt", "Fin de armado de mensajes para usuarios", sErrorDescription) 'TRACE

'lErrorNumber = AppendTextToFile(sFilePath & ".txt", "Inicio de ejecución del query", sErrorDescription) 'TRACE
		bAlimony = (CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14)) = 2)
		bCreditor = (CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14)) = 4)
		lStartPayrollDate = GetPayrollStartDate(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1))
		sCondition = " And (Payments.PaymentID>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12) & ") And (Payments.PaymentID<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13) & ") And (Payments.StatusID In (-2,-1,1))"
		If Len(oRequest("FilterFirstNumber").Item) > 0 Then
			If IsNumeric(oRequest("FilterFirstNumber").Item) Then
				If (iConnectionType <> ACCESS_DSN) And (iConnectionType <> ORACLE) Then
					sCondition = sCondition & " And (Cast(Payments.CheckNumber As int)>=" & oRequest("FilterFirstNumber").Item & ")"
				Else
					If (iConnectionType = ORACLE) Then
						sCondition = sCondition & " And (Payments.CheckNumber>=" & oRequest("FilterFirstNumber").Item & ")"
					Else
						sCondition = sCondition & " And (Payments.CheckNumber>='" & oRequest("FilterFirstNumber").Item & "')"
					End If
				End If
			End If
		End If
		If Len(oRequest("FilterLastNumber").Item) > 0 Then
			If IsNumeric(oRequest("FilterLastNumber").Item) Then
				If (iConnectionType <> ACCESS_DSN) And (iConnectionType <> ORACLE) Then
					sCondition = sCondition & " And (Cast(Payments.CheckNumber As int)<=" & oRequest("FilterLastNumber").Item & ")"
				Else
					If (iConnectionType = ORACLE) Then
						sCondition = sCondition & " And (Payments.CheckNumber<=" & oRequest("FilterLastNumber").Item & ")"
					Else
						sCondition = sCondition & " And (Payments.CheckNumber<='" & oRequest("FilterLastNumber").Item & "')"
					End If
				End If
			End If
		End If

	If lErrorNumber = 0 Then
		lStartDate = Left(GetSerialNumberForDate(DateAdd("d", -15, DateAdd("m", -1, GetDateFromSerialNumber(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1))))), Len("00000000"))
		sErrorDescription = "No se pudieron cancelar los cheques de los empleados que tienen más de tres cheques consecutivos cancelados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payments Set StatusID=0, Description='Cancelación automática por tener los últimos 3 cheques cancelados.' Where (EmployeeID In (Select Payments.EmployeeID From Payments Where (Payments.PaymentDate>=" & lStartDate & ") And (Payments.PaymentDate<" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Payments.StatusID Not In (-2,-1,1)) Group By Payments.EmployeeID Having (Count(*)=3))) " & sCondition, "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If

	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudieron obtener los empleados que cumplen con los criterios de la búsqueda."
		If bAlimony Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaymentID, Payments.PaymentDate, CheckAmount, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeNumber, BeneficiaryName As EmployeeName, BeneficiaryLastName As EmployeeLastName, BeneficiaryLastName2 As EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID As AreaID1, Areas2.AreaCode As AreaCode2, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber As FromAccountNumber, ToBankAccounts.BankID As ToBankID, ToBankAccounts.AccountNumber As ToAccountNumber, Min(EmployeesChangesLKP.FirstDate) As MinDate, Max(EmployeesChangesLKP.LastDate) As MaxDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, '' As ConceptRetention, Sum(Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptAmount) As TotalAmount From Payments, EmployeesBeneficiariesLKP, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", Concepts, EmployeesChangesLKP, EmployeesHistoryListForPayroll, Employees, Positions, GroupGradeLevels, Levels, Areas As Areas1, Areas As Areas2, Zones, BankAccounts As FromBankAccounts, BankAccounts As ToBankAccounts Where (Payments.EmployeeID=Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptID=Concepts.ConceptID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Payments.FromAccountID=FromBankAccounts.AccountID) And (Payments.AccountID=ToBankAccounts.AccountID) And (Payments.PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".RecordDate=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesBeneficiariesLKP.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesBeneficiariesLKP.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Group By PaymentID, Payments.PaymentDate, CheckAmount, EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesBeneficiariesLKP.BeneficiaryNumber, BeneficiaryName, BeneficiaryLastName, BeneficiaryLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID, Areas2.AreaCode, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber, ToBankAccounts.BankID, ToBankAccounts.AccountNumber, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction Order by PaymentID, IsDeduction, ConceptShortName", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select PaymentID, Payments.PaymentDate, CheckAmount, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeNumber, BeneficiaryName As EmployeeName, BeneficiaryLastName As EmployeeLastName, BeneficiaryLastName2 As EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID As AreaID1, Areas2.AreaCode As AreaCode2, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber As FromAccountNumber, ToBankAccounts.BankID As ToBankID, ToBankAccounts.AccountNumber As ToAccountNumber, Min(EmployeesChangesLKP.FirstDate) As MinDate, Max(EmployeesChangesLKP.LastDate) As MaxDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, '' As ConceptRetention, Sum(Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptAmount) As TotalAmount From Payments, EmployeesBeneficiariesLKP, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", Concepts, EmployeesChangesLKP, EmployeesHistoryListForPayroll, Employees, Positions, GroupGradeLevels, Levels, Areas As Areas1, Areas As Areas2, Zones, BankAccounts As FromBankAccounts, BankAccounts As ToBankAccounts Where (Payments.EmployeeID=Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptID=Concepts.ConceptID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Payments.FromAccountID=FromBankAccounts.AccountID) And (Payments.AccountID=ToBankAccounts.AccountID) And (Payments.PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".RecordDate=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesBeneficiariesLKP.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesBeneficiariesLKP.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Group By PaymentID, Payments.PaymentDate, CheckAmount, EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesBeneficiariesLKP.BeneficiaryNumber, BeneficiaryName, BeneficiaryLastName, BeneficiaryLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID, Areas2.AreaCode, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber, ToBankAccounts.BankID, ToBankAccounts.AccountNumber, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction Order by PaymentID, IsDeduction, ConceptShortName -->" & vbNewLine
		ElseIf bCreditor Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaymentID, Payments.PaymentDate, CheckAmount, EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesCreditorsLKP.CreditorNumber As EmployeeNumber, CreditorName As EmployeeName, CreditorLastName As EmployeeLastName, CreditorLastName2 As EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID As AreaID1, Areas2.AreaCode As AreaCode2, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber As FromAccountNumber, ToBankAccounts.BankID As ToBankID, ToBankAccounts.AccountNumber As ToAccountNumber, Min(EmployeesChangesLKP.FirstDate) As MinDate, Max(EmployeesChangesLKP.LastDate) As MaxDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, '' As ConceptRetention, Sum(Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptAmount) As TotalAmount From Payments, EmployeesCreditorsLKP, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", Concepts, EmployeesChangesLKP, EmployeesHistoryListForPayroll, Employees, Positions, GroupGradeLevels, Levels, Areas As Areas1, Areas As Areas2, Zones, BankAccounts As FromBankAccounts, BankAccounts As ToBankAccounts Where (Payments.EmployeeID=Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptID=Concepts.ConceptID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Payments.FromAccountID=FromBankAccounts.AccountID) And (Payments.AccountID=ToBankAccounts.AccountID) And (Payments.PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".RecordDate=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesCreditorsLKP.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesCreditorsLKP.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Group By PaymentID, Payments.PaymentDate, CheckAmount, EmployeesCreditorsLKP.CreditorNumber, EmployeesCreditorsLKP.CreditorNumber, CreditorName, CreditorLastName, CreditorLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID, Areas2.AreaCode, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber, ToBankAccounts.BankID, ToBankAccounts.AccountNumber, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction Order by PaymentID, IsDeduction, ConceptShortName", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select PaymentID, Payments.PaymentDate, CheckAmount, EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesCreditorsLKP.CreditorNumber As EmployeeNumber, CreditorName As EmployeeName, CreditorLastName As EmployeeLastName, CreditorLastName2 As EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID As AreaID1, Areas2.AreaCode As AreaCode2, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber As FromAccountNumber, ToBankAccounts.BankID As ToBankID, ToBankAccounts.AccountNumber As ToAccountNumber, Min(EmployeesChangesLKP.FirstDate) As MinDate, Max(EmployeesChangesLKP.LastDate) As MaxDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, '' As ConceptRetention, Sum(Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptAmount) As TotalAmount From Payments, EmployeesCreditorsLKP, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", Concepts, EmployeesChangesLKP, EmployeesHistoryListForPayroll, Employees, Positions, GroupGradeLevels, Levels, Areas As Areas1, Areas As Areas2, Zones, BankAccounts As FromBankAccounts, BankAccounts As ToBankAccounts Where (Payments.EmployeeID=Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptID=Concepts.ConceptID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Payments.FromAccountID=FromBankAccounts.AccountID) And (Payments.AccountID=ToBankAccounts.AccountID) And (Payments.PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".RecordDate=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesCreditorsLKP.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesCreditorsLKP.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Group By PaymentID, Payments.PaymentDate, CheckAmount, EmployeesCreditorsLKP.CreditorNumber, EmployeesCreditorsLKP.CreditorNumber, CreditorName, CreditorLastName, CreditorLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID, Areas2.AreaCode, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber, ToBankAccounts.BankID, ToBankAccounts.AccountNumber, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction Order by PaymentID, IsDeduction, ConceptShortName -->" & vbNewLine
		Else
			'lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaymentID, Payments.PaymentDate, CheckAmount, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID As AreaID1, Areas2.AreaCode As AreaCode2, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber As FromAccountNumber, ToBankAccounts.BankID As ToBankID, ToBankAccounts.AccountNumber As ToAccountNumber, Min(EmployeesChangesLKP.FirstDate) As MinDate, Max(EmployeesChangesLKP.LastDate) As MaxDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Concepts.OrderInList, ConceptRetention, Sum(ConceptAmount) As TotalAmount From Payments, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", Concepts, EmployeesChangesLKP, EmployeesHistoryListForPayroll, Employees, Positions, GroupGradeLevels, Levels, Areas As Areas1, Areas As Areas2, Zones, BankAccounts As FromBankAccounts, BankAccounts As ToBankAccounts Where (Payments.EmployeeID=Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptID=Concepts.ConceptID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Payments.FromAccountID=FromBankAccounts.AccountID) And (Payments.AccountID=ToBankAccounts.AccountID) And (Payments.PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".RecordDate=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Concepts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Concepts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Group By PaymentID, Payments.PaymentDate, CheckAmount, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID, Areas2.AreaCode, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber, ToBankAccounts.BankID, ToBankAccounts.AccountNumber, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Concepts.OrderInList, ConceptRetention Order by PaymentID, IsDeduction, Concepts.OrderInList", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			'Response.Write vbNewLine & "<!-- Query: Select PaymentID, Payments.PaymentDate, CheckAmount, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID As AreaID1, Areas2.AreaCode As AreaCode2, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber As FromAccountNumber, ToBankAccounts.BankID As ToBankID, ToBankAccounts.AccountNumber As ToAccountNumber, Min(EmployeesChangesLKP.FirstDate) As MinDate, Max(EmployeesChangesLKP.LastDate) As MaxDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Concepts.OrderInList, ConceptRetention, Sum(ConceptAmount) As TotalAmount From Payments, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", Concepts, EmployeesChangesLKP, EmployeesHistoryListForPayroll, Employees, Positions, GroupGradeLevels, Levels, Areas As Areas1, Areas As Areas2, Zones, BankAccounts As FromBankAccounts, BankAccounts As ToBankAccounts Where (Payments.EmployeeID=Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptID=Concepts.ConceptID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Payments.FromAccountID=FromBankAccounts.AccountID) And (Payments.AccountID=ToBankAccounts.AccountID) And (Payments.PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".RecordDate=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Concepts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Concepts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Group By PaymentID, Payments.PaymentDate, CheckAmount, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID, Areas2.AreaCode, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber, ToBankAccounts.BankID, ToBankAccounts.AccountNumber, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Concepts.OrderInList, ConceptRetention Order by PaymentID, IsDeduction, Concepts.OrderInList -->" & vbNewLine
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payments.PaymentID, Payments.PaymentDate, CheckAmount, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesHistoryListForPayroll.BankID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID As AreaID1, Areas2.AreaCode As AreaCode2, ZonePath, Payments.CheckNumber, EmployeesHistoryListForPayroll.AccountNumber As ToAccountNumber, DatECHG.FirstDate MinDate, DatECHG.LastDate MaxDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Concepts.OrderInList, ConceptRetention, Sum(ConceptAmount) As TotalAmount From Payments, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", Concepts, EmployeesHistoryListForPayroll, (Select pymnt.PaymentID, Min(echg.FirstDate) FirstDate, Max(echg.LastDate) LastDate From Payments pymnt, EmployeesChangesLKP  echg Where (pymnt.PaymentDate = " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (echg.PayrollID = " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (pymnt.EmployeeID = echg.EmployeeID) And (pymnt.PaymentDate = echg.PayrollID) Group By pymnt.PaymentID) DatECHG, Employees, Positions, GroupGradeLevels, Levels, Areas As Areas1, Areas As Areas2, Zones Where (Payments.EmployeeID=Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptID=Concepts.ConceptID) And (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payments.PaymentID = DatECHG.PaymentID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Payments.PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Concepts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Concepts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Group By Payments.PaymentID, Payments.PaymentDate, CheckAmount, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, SocialSecurityNumber,BankID, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID, Areas2.AreaCode, ZonePath, Payments.CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, DatECHG.FirstDate, DatECHG.LastDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Concepts.OrderInList, ConceptRetention Order by Payments.PaymentID, IsDeduction, Concepts.OrderInList", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select Payments.PaymentID, Payments.PaymentDate, CheckAmount, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID As AreaID1, Areas2.AreaCode As AreaCode2, ZonePath, Payments.CheckNumber, EmployeesHistoryListForPayroll.AccountNumber As ToAccountNumber, DatECHG.FirstDate MinDate, DatECHG.LastDate MaxDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Concepts.OrderInList, ConceptRetention, Sum(ConceptAmount) As TotalAmount From Payments, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", Concepts, EmployeesHistoryListForPayroll, (Select pymnt.PaymentID, Min(echg.FirstDate) FirstDate, Max(echg.LastDate) LastDate From Payments pymnt, EmployeesChangesLKP  echg Where (pymnt.PaymentDate = " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (echg.PayrollID = " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (pymnt.EmployeeID = echg.EmployeeID) And (pymnt.PaymentDate = echg.PayrollID) Group By pymnt.PaymentID) DatECHG, Employees, Positions, GroupGradeLevels, Levels, Areas Areas1, Areas Areas2, Zones Where (Payments.EmployeeID=Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptID=Concepts.ConceptID) And (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payments.PaymentID = DatECHG.PaymentID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Payments.PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Concepts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Concepts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Group By Payments.PaymentID, Payments.PaymentDate, CheckAmount, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID, Areas2.AreaCode, ZonePath, Payments.CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, DatECHG.FirstDate, DatECHG.LastDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Concepts.OrderInList, ConceptRetention Order by Payments.PaymentID, IsDeduction, Concepts.OrderInList -->" & vbNewLine
		End If
	End If

'lErrorNumber = AppendTextToFile(sFilePath & ".txt", "Fin de ejecución del query", sErrorDescription) 'TRACE

		If lErrorNumber = 0 Then
			If Not oRecordset.EOF Then
				bFirstEmployee = True
				If lErrorNumber = 0 Then
					Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
					Response.Flush()

'lErrorNumber = AppendTextToFile(sFilePath & ".txt", "Inicio de preproceso", sErrorDescription) 'TRACE
					sCurrentNumber = CStr(oRecordset.Fields("CheckNumber").Value)
					lMinDate = 30000000
					lMaxDate = 0
					Do While Not oRecordset.EOF
						If lMinDate > oRecordset.Fields("MinDate").Value Then lMinDate = oRecordset.Fields("MinDate").Value
						If lMaxDate < oRecordset.Fields("MaxDate").Value Then lMaxDate = oRecordset.Fields("MaxDate").Value
						If StrComp(sCurrentNumber, CStr(oRecordset.Fields("CheckNumber").Value), vbBinaryCompare) <> 0 Then Exit Do
						oRecordset.MoveNext
						If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
					oRecordset.MoveFirst
'lErrorNumber = AppendTextToFile(sFilePath & ".txt", "Fin de preproceso", sErrorDescription) 'TRACE

					lEmployeeCounter = 0
					lFileCounter = 0
					sCurrentNumber = ""
					'lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\rtf1 \ansi \deff0 {\fonttbl {\f0\froman Times New Roman;}}\fs16", sErrorDescription)
					'lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\rtf1 \ansi \deff0 {\fonttbl {\f0\fswiss Arial;}}\fs18", sErrorDescription)
					lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\rtf1 \ansi \deff0 {\fonttbl {\f0\fmodern Tahoma;}}\fs18", sErrorDescription)
					If Not FileExists(sFilePath & "\" & sPensionISSSTELogo, sErrorDescription) Then
						lErrorNumber = CopyFile(sImagesPath & "\" & sPensionISSSTELogo, sFilePath & "\" & sPensionISSSTELogo, sErrorDescription)
					End If
					Do While Not oRecordset.EOF
						If StrComp(sCurrentNumber, CStr(oRecordset.Fields("CheckNumber").Value), vbBinaryCompare) <> 0 Then
'lErrorNumber = AppendTextToFile(sFilePath & ".txt", (sCurrentNumber & vbTab & lFileCounter), sErrorDescription) 'TRACE
							lEmployeeCounter = lEmployeeCounter + 1
							
							If (lEmployeeCounter Mod CHECKS_PER_FILE) = 0 Then
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx-50 \posy-50 \absw12500 \absh14399 \par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "}", sErrorDescription)
								lFileCounter = Int(lEmployeeCounter / CHECKS_PER_FILE)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\rtf1 \ansi \deff0 {\fonttbl {\f0\fmodern Tahoma;}}\fs18", sErrorDescription)
								bFirstEmployee = True
							End If
							yPositionForConceptsP=2700
							yPositionForConceptsD=2700
							If Not bFirstEmployee Then
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx-10 \posy-10 \absw12500 \absh14399 \par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\sbkpage{\*\atnid S A L T O  D E  S E C C I O N}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\sect\sectd{\*\atnid N U E V A  S E C C I O N}", sErrorDescription)
							End If
							
                            lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(500 + x1PositionDisplacement) & "\posy" & CInt(550 + l01PositionDisplacement + y1PositionDisplacement) & " \absw1247{\*\atnid NO.EMP.}", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(2200 + x1PositionDisplacement) & "\posy" & CInt(550 + l01PositionDisplacement + y1PositionDisplacement) & " \absw8000{\*\atnid NOMBRE}", sErrorDescription)
							        If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1)), sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1)), sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
                                    If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(8300 + x2PositionDisplacement)  & "\posy" & CInt(11700 + y2PositionDisplacement)  & " \absw2500{\*\atnid NO.CHEQUE}", sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(SizeText(CStr(oRecordset.Fields("CheckNumber").Value), "0", 10, 0)), sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        Else 
							            lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10000 + x1PositionDisplacement) & "\posy" & CInt(550 + l01PositionDisplacement + y1PositionDisplacement) & " \absw2500{\*\atnid NO.CHEQUE}", sErrorDescription)
							            lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(SizeText(CStr(oRecordset.Fields("CheckNumber").Value), "0", 10, 0)), sErrorDescription)
							            lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(500 + x1PositionDisplacement) & "\posy" & CInt(1050 + l02PositionDisplacement + y1PositionDisplacement) & " \absw2240{\*\atnid RFC}", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(4000 + x1PositionDisplacement) & "\posy" & CInt(1050 + l02PositionDisplacement + y1PositionDisplacement) & " \absw5075{\*\atnid CURP}", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(9000 + x1PositionDisplacement) & "\posy" & CInt(1050 + l02PositionDisplacement + y1PositionDisplacement) & " \absw3856{\*\atnid NSS}", sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(CStr(oRecordset.Fields("SocialSecurityNumber").Value)), sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(7740 + x1PositionDisplacement) & "\posy" & CInt(1050 + l02PositionDisplacement + y1PositionDisplacement) + l01PositionDisplacement & " \absw3856{\*\atnid NSS}", sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(CStr(oRecordset.Fields("SocialSecurityNumber").Value)), sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        End If
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' CLAVE PUESTO
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(500  + x1PositionDisplacement) & "\posy" & CInt(1500 + l03PositionDisplacement + y1PositionDisplacement) & " \absw2240{\*\atnid CLAVE_PUESTO}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(500  + x1PositionDisplacement) & "\posy" & CInt(1600 + l03PositionDisplacementR + y1PositionDisplacement) & " \absw2240{\*\atnid CLAVE_PUESTO}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), SizeText(CStr(oRecordset.Fields("PositionShortName").Value), " ", 7, 1), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' DESC_PUESTO
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(2900 + x1PositionDisplacement) & "\posy" & CInt(1500 + l03PositionDisplacement + y1PositionDisplacement) & " \absw3515{\*\atnid DESC_PUESTO}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(2300 + x1PositionDisplacement) & "\posy" & CInt(1600 + l03PositionDisplacementR + y1PositionDisplacement) & " \absw4115{\*\atnid DESC_PUESTO}", sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs14", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), SizeText(CStr(oRecordset.Fields("PositionName").Value), " ", 60, 1), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' NIV_SUB
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(6110 + x1PositionDisplacement) & "\posy" & CInt(1500 + l03PositionDisplacement + y1PositionDisplacement) & " \absw737{\*\atnid NIV_SUB}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(6210 + x1PositionDisplacement) & "\posy" & CInt(1600 + l03PositionDisplacementR + y1PositionDisplacement) & " \absw637{\*\atnid NIV_SUB}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(Left(CStr(oRecordset.Fields("LevelShortName").Value), Len("00")) & " / " & Right(CStr(oRecordset.Fields("LevelShortName").Value), Len("0"))), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' RG_MX
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(7010  + x1PositionDisplacement) & "\posy" & CInt(1500 + l03PositionDisplacement + y1PositionDisplacement) & " \absw567{\*\atnid RG_MX}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(7010  + x1PositionDisplacement) & "\posy" & CInt(1600 + l03PositionDisplacementR + y1PositionDisplacement) & " \absw567{\*\atnid RG_MX}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "2/0", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' CG_RP
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(7610- + x1PositionDisplacement) & "\posy" & CInt(1500 + l03PositionDisplacement + y1PositionDisplacement) & " \absw567{\*\atnid CG_RP}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(7610 + x1PositionDisplacement) & "\posy" & CInt(1600 + l03PositionDisplacementR + y1PositionDisplacement) & " \absw567{\*\atnid CG_RP}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "/X", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' CLAVE_PRESUPUESTAL
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(8600 + x1PositionDisplacement) & "\posy" & CInt(1500 + l03PositionDisplacement + y1PositionDisplacement) & " \absw1500{\*\atnid CLAVE_PRESUPUESTAL}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(8600 + x1PositionDisplacement) & "\posy" & CInt(1600 + l03PositionDisplacementR + y1PositionDisplacement) & " \absw1500{\*\atnid CLAVE_PRESUPUESTAL}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(CStr(oRecordset.Fields("AreaCode2").Value)), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' CLAVE_DISTRIBUCION
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10400 + x1PositionDisplacement) & "\posy" & CInt(1500 + l03PositionDisplacement + y1PositionDisplacement) & " \absw1500{\*\atnid CLAVE_DISTRIBUCION}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10400 + x1PositionDisplacement) & "\posy" & CInt(1600 + l03PositionDisplacementR + y1PositionDisplacement) & " \absw1500{\*\atnid CLAVE_DISTRIBUCION}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(CStr(oRecordset.Fields("AreaCode2").Value)), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' FECHA_INGRESO
						    		    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(500 + x1PositionDisplacement) & "\posy" & CInt(1950 + l04PositionDisplacement + y1PositionDisplacement) & " \absw1956{\*\atnid FECHA_INGRESO}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt( 500 + x1PositionDisplacement) & "\posy" & CInt(2100 + l04PositionDisplacementR + y1PositionDisplacement) & " \absw1956{\*\atnid FECHA_INGRESO}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' FECHA_NOMINA
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(2900 + x1PositionDisplacement) & "\posy" & CInt(1950 + l04PositionDisplacement + y1PositionDisplacement) & " \absw2256{\*\atnid FECHA_NOMINA}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(2900 + x1PositionDisplacement) & "\posy" & CInt(2100 + l04PositionDisplacementR + y1PositionDisplacement) & " \absw2256{\*\atnid FECHA_NOMINA}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), DisplayNumericDateFromSerialNumber(lPayrollID), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' PERIODO_PAGO
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(4500 + x1PositionDisplacement) & "\posy" & CInt(1950 + l04PositionDisplacement + y1PositionDisplacement) & " \absw3742{\*\atnid PERIODO_PAGO}", sErrorDescription)
							        Else
							            lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(4500 + x1PositionDisplacement) & "\posy" & CInt(2100 + l04PositionDisplacementR + y1PositionDisplacement) & " \absw3742{\*\atnid PERIODO_PAGO}", sErrorDescription)
							        End If
							        If lMinDate = 0 Then lMinDate = lPayrollID End If
							        If lMaxDate = 0 Then lMaxDate = lPayrollID End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), DisplayNumericDateFromSerialNumber(lMinDate) & " AL " & DisplayNumericDateFromSerialNumber(lMaxDate), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' NETO_A_PAGAR
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10200 + x1PositionDisplacement) & "\posy" & CInt(1950 + l04PositionDisplacement + y1PositionDisplacement) & " \absw1400{\*\atnid NETO_A_PAGAR}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10400 + x1PositionDisplacement) & "\posy" & CInt(2100 + l04PositionDisplacementR + y1PositionDisplacement) & " \absw1400{\*\atnid NETO_A_PAGAR}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\b", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), FormatNumber(CDbl(oRecordset.Fields("CheckAmount").Value), True), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							       

							If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(7100 + x2PositionDisplacement)& "\posy" & CInt(12000 + y2PositionDisplacement) & " \absw2156{\*\atnid FECHA_NOMINA}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PaymentDate").Value), -1, -1, -1), sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10200 + x2PositionDisplacement) & "\posy" & CInt(12850 + y2PositionDisplacement) & " \absw2000{\*\atnid MONEDA_NACIONAL}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), FormatNumber(CDbl(oRecordset.Fields("CheckAmount").Value), 2, True, False, True), sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(2212 + x2PositionDisplacement) & "\posy" & CInt(12800 + y2PositionDisplacement) & " \absw8500{\*\atnid NOMBRE_CHEQUE}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs18 \b", sErrorDescription)
								If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
									lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1)), sErrorDescription)
								Else
									lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1)), sErrorDescription)
								End If
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(1200 + x2PositionDisplacement)  & "\posy" & CInt(13250 + y2PositionDisplacement) & " \absw7950{\*\atnid MONEDA_NACIONAL_TEXTO}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), UCase(FormatNumberAsText(CDbl(oRecordset.Fields("CheckAmount").Value), True)), sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							End If
							If Not bAlimony And Not bCreditor Then
'								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(1400 + x1PositionDisplacement) & "\posy" & CInt(8000 + y1PositionDisplacement) & " \absw8878{\*\atnid LOGO_PENSION_ISSSTE}", sErrorDescription)
'								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\field\fldedit{\*\fldinst { INCLUDEPICTURE \\d", sErrorDescription)
'								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\~"&sPensionISSSTELogo, sErrorDescription)
'								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\\* MERGEFORMATINET }}{\fldrslt { }}}", sErrorDescription)
'								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							End If

							asPath = Split(CStr(oRecordset.Fields("ZonePath").Value), ",")
							If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then
								If bAlimony Or bCreditor Then
									sSignatureName1 = "0901.jpg"
									If Not FileExists(sFilePath & "\" & sSignatureName, sErrorDescription) Then
										lErrorNumber = CopyFile(sImagesPath & "\" & sSignatureName1, sFilePath & "\" & sSignatureName1, sErrorDescription)
									End If
									sSignatureName2 = "0902.jpg"
									If Not FileExists(sFilePath & "\" & sSignatureName, sErrorDescription) Then
										lErrorNumber = CopyFile(sImagesPath & "\" & sSignatureName2, sFilePath & "\" & sSignatureName2, sErrorDescription)
									End If
								ElseIf CLng(oRecordset.Fields("AreaID1").Value) <> 38 Then
									sSignatureName1 = Right(("00" & asPath(2)), Len("00")) & "01.jpg"
									If Not FileExists(sFilePath & "\" & sSignatureName1, sErrorDescription) Then
										lErrorNumber = CopyFile(sImagesPath & "\" & sSignatureName1, sFilePath & "\" & sSignatureName1, sErrorDescription)
									End If
									sSignatureName2 = Right(("00" & asPath(2)), Len("00")) & "02.jpg"
									If Not FileExists(sFilePath & "\" & sSignatureName2, sErrorDescription) Then
										lErrorNumber = CopyFile(sImagesPath & "\" & sSignatureName2, sFilePath & "\" & sSignatureName2, sErrorDescription)
									End If
								Else
									sSignatureName1 = "3801.jpg"
									If Not FileExists(sFilePath & "\" & sSignatureName, sErrorDescription) Then
										lErrorNumber = CopyFile(sImagesPath & "\" & sSignatureName1, sFilePath & "\" & sSignatureName1, sErrorDescription)
									End If
									sSignatureName2 = "3802.jpg"
									If Not FileExists(sFilePath & "\" & sSignatureName, sErrorDescription) Then
										lErrorNumber = CopyFile(sImagesPath & "\" & sSignatureName2, sFilePath & "\" & sSignatureName2, sErrorDescription)
									End If
								End If
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(6323 + x2PositionDisplacement) & "\posy" & CInt(13850 + lSignaturesPositionDisplacementR + y2PositionDisplacement) & " \absw2778{\*\atnid FIRMA1}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\field\fldedit{\*\fldinst { INCLUDEPICTURE \\d", sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\~"&sSignatureName1, sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\\* MERGEFORMATINET }}{\fldrslt { }}}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(9100 + x2PositionDisplacement) & "\posy" & CInt(13850 + lSignaturesPositionDisplacementR + y2PositionDisplacement) & " \absw2778{\*\atnid FIRMA2}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\field\fldedit{\*\fldinst { INCLUDEPICTURE \\d", sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\~"&sSignatureName2, sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\\* MERGEFORMATINET }}{\fldrslt { }}}", sErrorDescription)
								lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							End If

							sContents = ""
							For iIndex = 0 To UBound(asMessages) - 1
								If ((StrComp(asMessages(iIndex)(1), ",0,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(1), "," & CStr(oRecordset.Fields("EmployeeID").Value) & ",", vbBinaryCompare) > 0)) And _
								   ((StrComp(asMessages(iIndex)(2), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(2), "," & CStr(oRecordset.Fields("CompanyID").Value) & ",", vbBinaryCompare) > 0)) And _
								   ((StrComp(asMessages(iIndex)(3), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(3), "," & CStr(oRecordset.Fields("AreaID1").Value) & ",", vbBinaryCompare) > 0)) And _
								   ((StrComp(asMessages(iIndex)(4), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(4), "," & asPath(2) & ",", vbBinaryCompare) > 0)) And _
								   ((StrComp(asMessages(iIndex)(5), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(5), "," & CStr(oRecordset.Fields("EmployeeTypeID").Value) & ",", vbBinaryCompare) > 0)) And _
								   ((StrComp(asMessages(iIndex)(6), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(6), "," & CStr(oRecordset.Fields("PositionID").Value) & ",", vbBinaryCompare) > 0))Then 'And _
								   '((StrComp(asMessages(iIndex)(7), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(7), "," & CStr(oRecordset.Fields("ToBankID").Value) & ",", vbBinaryCompare) > 0)) Then
									Select Case CInt(asMessages(iIndex)(8))
										Case 0
											If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then sContents = sContents & asMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
										Case 1
											If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) <> 0 Then sContents = sContents & asMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
										Case 2
											If bAlimony Then sContents = sContents & asMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
										Case 4
											If bCreditor Then sContents = sContents & asMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
										Case Else
											sContents = sContents & asMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
									End Select
								End If
							Next
							For iIndex = 0 To UBound(asEmployeesMessages) - 1
								If StrComp(asEmployeesMessages(iIndex)(1), CStr(oRecordset.Fields("EmployeeID").Value), vbBinaryCompare) = 0 Then
									sContents = sContents & asEmployeesMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
								End If
							Next
							lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(900 + x1PositionDisplacement) & "\posy" & CInt(9200 + y1PositionDisplacement+1500) & " \absw9000{\*\atnid DESC_PUESTO}", sErrorDescription)
							lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), sContents, sErrorDescription)
							lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)

                            If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) <> 0 Then
                                lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(1500) & "\posy" & CInt(8100) & " \absw9000{\*\atnid LOGO_ISSSTE}", sErrorDescription)
						        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\field\fldedit{\*\fldinst { INCLUDEPICTURE \\d", sErrorDescription)
						        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\~"&sPensionISSSTELogo, sErrorDescription)
						        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\\* MERGEFORMATINET }}{\fldrslt { }}}", sErrorDescription)
						        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
                           End If
							sPerceptions = ""
							sDeductions = ""
							sCurrentNumber = CStr(oRecordset.Fields("CheckNumber").Value)
							bFirstEmployee = False
							lMinDate = 30000000
							lMaxDate = 0
						End If

                        Select Case CStr(oRecordset.Fields("ConceptShortName").Value)
							        Case "D"
								        lDeductions = CDbl(oRecordset.Fields("TotalAmount").Value)
								        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' TOTAL_DEDUCCIONES
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10524 + x1PositionDisplacement) & "\posy" & CInt(6600 + lConceptsTPositionDisplacement + y1PositionDisplacement) & " \absw1500{\*\atnid TOTAL_DEDUCCIONES}", sErrorDescription)
								        Else
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10599 + x1PositionDisplacement ) & "\posy" & CInt(6940 + lConceptsTPositionDisplacementR + y1PositionDisplacement) & " \absw1500{\*\atnid TOTAL_DEDUCCIONES}", sErrorDescription)
								        End If
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), FormatNumber(lDeductions, 2, True, False, True), sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
                                    Case "L"
							            '	'sContents = Replace(sContents, "<TOTAL />", FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True))
							            '	lConceptAmount = CDbl(oRecordset.Fields("TotalAmount").Value)
							            '	lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(800 + x1PositionDisplacement) & "\posy" & CInt(12814 + y1PositionDisplacement) & " \absw1701{\*\atnid NETO_A_PAGAR}", sErrorDescription)
							            '	lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), FormatNumber(lConceptAmount, 2, True, False, True), sErrorDescription)
							            '	lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        Case "P"
								        lPerceptions = CDbl(oRecordset.Fields("TotalAmount").Value)
								        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' TOTAL_PERCEPCIONES
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(4824 + x1PositionDisplacement) & "\posy" & CInt(6600 + lConceptsTPositionDisplacement + y1PositionDisplacement) & " \absw1500{\*\atnid TOTAL_PERCEPCIONES}", sErrorDescription)
								        Else
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(4924 + x1PositionDisplacement) & "\posy" & CInt(6940 + lConceptsTPositionDisplacementR + y1PositionDisplacement) & " \absw1500{\*\atnid TOTAL_PERCEPCIONES}", sErrorDescription)
						    	        End If
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), FormatNumber(lPerceptions, 2, True, False, True), sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
								        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' BASE_GRAVABLE
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(8400 + x1PositionDisplacement) & "\posy" & CInt(1950 + l04PositionDisplacement + y1PositionDisplacement) & " \absw1400{\*\atnid BASE_GRAVABLE}", sErrorDescription)
								        Else
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(8600 + x1PositionDisplacement) & "\posy" & CInt(2100 + l04PositionDisplacementR + y1PositionDisplacement) & " \absw1400{\*\atnid BASE_GRAVABLE}", sErrorDescription)
								        End If
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), FormatNumber(lPerceptions, 2, True, False, True), sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        Case Else
								        If CInt(oRecordset.Fields("IsDeduction").Value) = 0 Then
									        lConceptAmount = CDbl(oRecordset.Fields("TotalAmount").Value)
									        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' CONCEPTO1
										        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(500+ x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsP + lConceptsPositionDisplacement + y1PositionDisplacement-50) & " \absw737{\*\atnid CAVE.CPTO.} {\*\atnid CONCEPTO1}", sErrorDescription)
									        Else
										        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(500 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsP + 200 + lConceptsPositionDisplacementR + y1PositionDisplacement-50) & " \absw737{\*\atnid CAVE.CPTO.} {\*\atnid CONCEPTO1}", sErrorDescription)
									        End If
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs14", sErrorDescription)
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CStr(oRecordset.Fields("ConceptShortName").Value), sErrorDescription)
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
									        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' DESCRIPCION
										        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(1300 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsP + lConceptsPositionDisplacement + y1PositionDisplacement-50) & " \absw3515{\*\atnid DESCRIPCION}", sErrorDescription)
									        Else
										        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(1300 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsP + 200 + lConceptsPositionDisplacementR + y1PositionDisplacement-50) & " \absw3515{\*\atnid DESCRIPCION}", sErrorDescription)
									        End If
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs14", sErrorDescription)
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CStr(oRecordset.Fields("ConceptName").Value), sErrorDescription)
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
									        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' IMPORTE
										        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(4924 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsP + lConceptsPositionDisplacement + y1PositionDisplacement-50) & " \absw1214{\*\atnid IMPORTE}", sErrorDescription)
									        Else
										        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(4924 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsP + 200 + lConceptsPositionDisplacementR + y1PositionDisplacement-50) & " \absw1214{\*\atnid IMPORTE}", sErrorDescription)
									        End If
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs14", sErrorDescription)
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), FormatNumber(lConceptAmount, 2, True, False, True), sErrorDescription)
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
									        If Len(CStr(oRecordset.Fields("ConceptName").Value)) > 53 Then
										        yPositionForConceptsP = yPositionForConceptsP + 300
									        Else
										        yPositionForConceptsP = yPositionForConceptsP + 150
									        End If
								        Else
									        lConceptAmount = CDbl(oRecordset.Fields("TotalAmount").Value)
									        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' CONCEPTO1
										        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(5990 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsD + lConceptsPositionDisplacement + y1PositionDisplacement-50) & " \absw737{\*\atnid CAVE.CPTO.} {\*\atnid CONCEPTO1}", sErrorDescription)
									        Else
										        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(5990+ x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsD + 200 + lConceptsPositionDisplacementR + y1PositionDisplacement-50) & " \absw737{\*\atnid CAVE.CPTO.} {\*\atnid CONCEPTO1}", sErrorDescription)
									        End If
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs14", sErrorDescription)
					    			        	lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CStr(oRecordset.Fields("ConceptShortName").Value), sErrorDescription)
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
										If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' DESCRIPCION
											lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(6800 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsD + lConceptsPositionDisplacement + y1PositionDisplacement-50) & " \absw3515{\*\atnid DESCRIPCION}", sErrorDescription)
										Else
											lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(6800 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsD + 200 + lConceptsPositionDisplacementR + y1PositionDisplacement-50) & " \absw3515{\*\atnid DESCRIPCION}", sErrorDescription)
										End If
										lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs14", sErrorDescription)
										lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CStr(oRecordset.Fields("ConceptName").Value), sErrorDescription)
										lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)

										If InStr(1, S_CREDITS_ID, "," & CInt(oRecordset.Fields("ConceptID").Value) & ",", vbBinaryCompare) > 0 Then
											If CDbl(oRecordset.Fields("ConceptRetention").Value) > 0 Then
												If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' CONTADOR
													lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(9000 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsD + lConceptsPositionDisplacement + y1PositionDisplacement-50) & " \absw1214{\*\atnid CONTADOR}", sErrorDescription)
												Else
													lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(9000 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsD + 200 + lConceptsPositionDisplacementR + y1PositionDisplacement-50) & " \absw1214{\*\atnid CONTADOR}", sErrorDescription)
												End If
												lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs14", sErrorDescription)
												lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), Replace(Left(CStr(oRecordset.Fields("ConceptRetention").Value), (Len(CStr(oRecordset.Fields("ConceptRetention").Value)) - Len("1"))), ".", "/"), sErrorDescription)
												lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
											End If
										End If

										If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' IMPORTE
											lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10524 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsD + lConceptsPositionDisplacement + y1PositionDisplacement-50) & " \absw1214{\*\atnid IMPORTE}", sErrorDescription)
										Else
											lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10524 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsD + 200 + lConceptsPositionDisplacementR + y1PositionDisplacement-50) & " \absw1214{\*\atnid IMPORTE}", sErrorDescription)
										End If
										lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs14", sErrorDescription)
										lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), FormatNumber(lConceptAmount, 2, True, False, True), sErrorDescription)
										lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
										If Len(CStr(oRecordset.Fields("ConceptName").Value)) > 53 Then
											yPositionForConceptsD = yPositionForConceptsD + 300
										Else
											yPositionForConceptsD = yPositionForConceptsD + 150
										End If
								End If
						End Select

						If lMinDate > oRecordset.Fields("MinDate").Value Then lMinDate = oRecordset.Fields("MinDate").Value
						If lMaxDate < oRecordset.Fields("MaxDate").Value Then lMaxDate = oRecordset.Fields("MaxDate").Value
						oRecordset.MoveNext
						If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
'lErrorNumber = AppendTextToFile(sFilePath & ".txt", (sCurrentNumber & vbTab & lFileCounter), sErrorDescription) 'TRACE
					lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx-10 \posy-10 \absw12500 \absh14399 \par}", sErrorDescription)
					lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "}", sErrorDescription)
'lErrorNumber = AppendTextToFile(sFilePath & ".txt", "Inicio de armado de archivo ZIP", sErrorDescription) 'TRACE
					lErrorNumber = ZipFile(sFilePath, Server.MapPath(sFileName), sErrorDescription)
					oRecordset.Close
					sErrorDescription = "No se pudieron obtener los empleados que cumplen con los criterios de la búsqueda."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payments Set StatusID=1, Description='" & Replace(oRequest("Description").Item, "'", "´") & "' Where (PaymentID>-1) " & sCondition, "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					If lErrorNumber = 0 Then
						aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(15) = 1
						lErrorNumber = ModifyCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
					End If
					If lErrorNumber = 0 Then
						lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
					End If
					oEndDate = Now()
					If (lErrorNumber = 0) And B_USE_SMTP Then
						If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
					End If
				End If
			Else
				lErrorNumber = -1
				sErrorDescription = "No existen empleados que cumplan con los criterios de la búsqueda."
			End If
		End If
'lErrorNumber = DeleteFile(sFilePath & ".txt", sErrorDescription) 'TRACE
	End If

	Set oRecordset = Nothing
	PrintPayments = lErrorNumber
	Err.Clear
End Function

Function PrintPayments(oRequest, oADODBConnection, lRecordID, sErrorDescription)
'************************************************************
'Purpose: To send the payments information to printing formats
'Inputs:  oRequest, oADODBConnection, lRecordID, aCatalogComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "PrintPayments"
	Const CHECKS_PER_FILE = 2000
	Dim S_CREDITS_ID
	Dim lEmployeeCounter
	Dim lFileCounter
	Dim sPerceptions
	Dim sDeductions
	Dim lStartPayrollDate
	Dim lStartDate
	Dim lMinDate
	Dim lMaxDate
	Dim sDate
	Dim sFilePath
	Dim sImagesPath
	Dim sFileName
	Dim sDocumentName
	Dim lReportID
	Dim sCurrentNumber
	Dim bAlimony
	Dim bCreditor
	Dim sCondition
	Dim asMessages
	Dim asEmployeesMessages
	Dim sContents
	Dim asPath
	Dim iIndex
	Dim oRecordset
	Dim asColumnsTitles
	Dim asRowContents
	Dim sRowContents
	Dim asTableColors()
	Dim asCellWidths
	Dim asCellAlignments
	Dim lErrorNumber
	Dim bFirstEmployee
	Dim lPerceptions
	Dim lDeductions
	Dim lConceptAmount
	Dim yPositionForConceptsP
	Dim yPositionForConceptsD
	Dim sSignatureName1
	Dim sSignatureName2
	Dim sPensionISSSTELogo
	Dim x1PositionDisplacement
	Dim y1PositionDisplacement
	Dim x2PositionDisplacement
	Dim y2PositionDisplacement
	Dim l01PositionDisplacement
	Dim l02PositionDisplacement
	Dim l03PositionDisplacement
	Dim l04PositionDisplacement
	Dim l05PositionDisplacement
	Dim l03PositionDisplacementR
	Dim l04PositionDisplacementR
	Dim l05PositionDisplacementR
	Dim lConceptsTPositionDisplacement
	Dim lConceptsPositionDisplacement
	Dim lSignaturesPositionDisplacement
	Dim lConceptsTPositionDisplacementR
	Dim lConceptsPositionDisplacementR
	Dim lSignaturesPositionDisplacementR
	Dim oEndDate
    Dim bBackReport
    Dim bPage
    Dim lcurrentEmployeeID
    Dim sEmployeeId1
    Dim sEmployeeId2
    Dim sAccountEmployee1
    Dim sAccountEmployee2
    Dim bStatusReport
    Dim sEmployee1Information
    Dim sEmployee2information
    Dim bIsCheck

	sDate = GetSerialNumberForDate("")

	y1PositionDisplacement = -100
	If Len(oRequest("PosX1").Item) > 0 Then x1PositionDisplacement = CInt(CInt(oRequest("PosX1").Item) * (56.692913386))
	If Len(oRequest("PosY1").Item) > 0 Then y1PositionDisplacement = y1PositionDisplacement + CInt(CInt(oRequest("PosY1").Item) * (56.692913386))
	If Len(oRequest("PosX2").Item) > 0 Then x2PositionDisplacement = CInt(CInt(oRequest("PosX2").Item) * (56.692913386))
	If Len(oRequest("PosY2").Item) > 0 Then y2PositionDisplacement = y2PositionDisplacement + CInt(CInt(oRequest("PosY2").Item) * (56.692913386))

	l01PositionDisplacement = 0
	l02PositionDisplacement = 0
	l03PositionDisplacement = 0
	l04PositionDisplacement = 0
	l05PositionDisplacement = 0
	l03PositionDisplacementR = 0
	l04PositionDisplacementR = 0
	l05PositionDisplacementR = 0
	lConceptsTPositionDisplacement = 0
	lConceptsPositionDisplacement = 0
	lSignaturesPositionDisplacement = 0
	lConceptsTPositionDisplacementR = 0
	lConceptsPositionDisplacementR = 0
	lSignaturesPositionDisplacementR = 0

	sPensionISSSTELogo = "PensionISSSTE.jpg"
	sFilePath = Server.MapPath(REPORTS_PATH & "Rep_1400_" & sDate)
	sImagesPath = Server.MapPath(TEMPLATES_PATH) & "\Images"
	sErrorDescription = "Error al crear la carpeta en donde se almacenará el reporte"
	lErrorNumber = CreateFolder(sFilePath, sErrorDescription)
	If lErrorNumber = 0 Then
		sFileName = REPORTS_PATH & "Rep_1400_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(0) & ".zip"
		sDocumentName = sFilePath & "\Rep_1400_" & sDate & "_<INDEX />.rtf"
		sErrorDescription = "No se pudieron obtener las nóminas de los empleados."
		If FileExists(Server.MapPath(sFileName), sErrorDescription) Then Call DeleteFile(Server.MapPath(sFileName), sErrorDescription)

		S_CREDITS_ID = ","
		sErrorDescription = "No se pudieron obtener los IDs de los créditos."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select CreditTypeID From CreditTypes Where (CreditTypeID>0) And (Active=1) Order By CreditTypeID", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				S_CREDITS_ID = S_CREDITS_ID & CStr(oRecordset.Fields("CreditTypeID").Value) & ","
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
		End If

		asMessages = ""
		sErrorDescription = "No se pudieron obtener los mensajes para los cheques."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From PaymentsMessages Where (PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (bSpecial=0) Order By RecordID", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select * From PaymentsMessages Where (PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (bSpecial=0) Order By RecordID -->" & vbNewLine
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				asMessages = asMessages & Replace(CStr(oRecordset.Fields("Comments").Value), "<BR />", ("\" & Chr(13) & Chr(10))) & LIST_SEPARATOR
				asMessages = asMessages & "," & Replace(CStr(oRecordset.Fields("EmployeeID").Value), " ", "") & "," & LIST_SEPARATOR
				asMessages = asMessages & "," & Replace(CStr(oRecordset.Fields("CompanyID").Value), " ", "") & "," & LIST_SEPARATOR
				asMessages = asMessages & "," & Replace(CStr(oRecordset.Fields("AreaIDs").Value), " ", "") & "," & LIST_SEPARATOR
				asMessages = asMessages & "," & Replace(CStr(oRecordset.Fields("ZoneIDs").Value), " ", "") & "," & LIST_SEPARATOR
				asMessages = asMessages & "," & Replace(CStr(oRecordset.Fields("EmployeeTypeID").Value), " ", "") & "," & LIST_SEPARATOR
				asMessages = asMessages & "," & Replace(CStr(oRecordset.Fields("PositionID").Value), " ", "") & "," & LIST_SEPARATOR
				asMessages = asMessages & "," & Replace(CStr(oRecordset.Fields("BankID").Value), " ", "") & "," & LIST_SEPARATOR
				asMessages = asMessages & Replace(CStr(oRecordset.Fields("ConceptID").Value), " ", "") & SECOND_LIST_SEPARATOR
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			asMessages = Split(asMessages, SECOND_LIST_SEPARATOR)
			For iIndex = 0 To UBound(asMessages)
				asMessages(iIndex) = Split(asMessages(iIndex), LIST_SEPARATOR)
			Next
		End If

		asEmployeesMessages = ""
		sErrorDescription = "No se pudieron obtener los mensajes para los cheques."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select EmployeeID, Comments From PaymentsMessages Where (PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (bSpecial Not In (0,3)) And (EmployeeID In (Select EmployeeID From Payments Where (PaymentID>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12) & ") And (PaymentID<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13) & "))) Order By EmployeeID, RecordID", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
		Response.Write vbNewLine & "<!-- Query: Select EmployeeID, Comments From PaymentsMessages Where (PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (bSpecial Not In (0,3)) And (EmployeeID In (Select EmployeeID From Payments Where (PaymentID>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12) & ") And (PaymentID<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13) & "))) Order By EmployeeID, RecordID -->" & vbNewLine
		If lErrorNumber = 0 Then
			Do While Not oRecordset.EOF
				asEmployeesMessages = asEmployeesMessages & Replace(CStr(oRecordset.Fields("Comments").Value), "<BR />", ("\" & Chr(13) & Chr(10))) & TABLE_SEPARATOR
				asEmployeesMessages = asEmployeesMessages & CStr(oRecordset.Fields("EmployeeID").Value) & CATALOG_SEPARATOR
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
			oRecordset.Close
			asEmployeesMessages = Split(asEmployeesMessages, CATALOG_SEPARATOR)
			For iIndex = 0 To UBound(asEmployeesMessages)
				asEmployeesMessages(iIndex) = Split(asEmployeesMessages(iIndex), TABLE_SEPARATOR)
			Next
		End If

		bAlimony = (CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14)) = 2)
		bCreditor = (CInt(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(14)) = 4)
		lStartPayrollDate = GetPayrollStartDate(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1))
		sCondition = " And (Payments.PaymentID>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(12) & ") And (Payments.PaymentID<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(13) & ") And (Payments.StatusID In (-2,-1,1))"
		If Len(oRequest("FilterFirstNumber").Item) > 0 Then
			If IsNumeric(oRequest("FilterFirstNumber").Item) Then
				If (iConnectionType <> ACCESS_DSN) And (iConnectionType <> ORACLE) Then
					sCondition = sCondition & " And (Cast(Payments.CheckNumber As int)>=" & oRequest("FilterFirstNumber").Item & ")"
				Else
					If (iConnectionType = ORACLE) Then
						sCondition = sCondition & " And (Payments.CheckNumber>=" & oRequest("FilterFirstNumber").Item & ")"
					Else
						sCondition = sCondition & " And (Payments.CheckNumber>='" & oRequest("FilterFirstNumber").Item & "')"
					End If
				End If
			End If
		End If
		If Len(oRequest("FilterLastNumber").Item) > 0 Then
			If IsNumeric(oRequest("FilterLastNumber").Item) Then
				If (iConnectionType <> ACCESS_DSN) And (iConnectionType <> ORACLE) Then
					sCondition = sCondition & " And (Cast(Payments.CheckNumber As int)<=" & oRequest("FilterLastNumber").Item & ")"
				Else
					If (iConnectionType = ORACLE) Then
						sCondition = sCondition & " And (Payments.CheckNumber<=" & oRequest("FilterLastNumber").Item & ")"
					Else
						sCondition = sCondition & " And (Payments.CheckNumber<='" & oRequest("FilterLastNumber").Item & "')"
					End If
				End If
			End If
		End If

	If lErrorNumber = 0 Then
		lStartDate = Left(GetSerialNumberForDate(DateAdd("d", -15, DateAdd("m", -1, GetDateFromSerialNumber(aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1))))), Len("00000000"))
		sErrorDescription = "No se pudieron cancelar los cheques de los empleados que tienen más de tres cheques consecutivos cancelados."
		lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payments Set StatusID=0, Description='Cancelación automática por tener los últimos 3 cheques cancelados.' Where (EmployeeID In (Select Payments.EmployeeID From Payments Where (Payments.PaymentDate>=" & lStartDate & ") And (Payments.PaymentDate<" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Payments.StatusID Not In (-2,-1,1)) Group By Payments.EmployeeID Having (Count(*)=3))) " & sCondition, "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	End If

	If lErrorNumber = 0 Then
		sErrorDescription = "No se pudieron obtener los empleados que cumplen con los criterios de la búsqueda."
		If bAlimony Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaymentID, Payments.PaymentDate, CheckAmount, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeNumber, BeneficiaryName As EmployeeName, BeneficiaryLastName As EmployeeLastName, BeneficiaryLastName2 As EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID As AreaID1, Areas2.AreaCode As AreaCode2, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber As FromAccountNumber, ToBankAccounts.BankID As ToBankID, ToBankAccounts.AccountNumber As ToAccountNumber, Min(EmployeesChangesLKP.FirstDate) As MinDate, Max(EmployeesChangesLKP.LastDate) As MaxDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, '' As ConceptRetention, Sum(Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptAmount) As TotalAmount From Payments, EmployeesBeneficiariesLKP, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", Concepts, EmployeesChangesLKP, EmployeesHistoryListForPayroll, Employees, Positions, GroupGradeLevels, Levels, Areas As Areas1, Areas As Areas2, Zones, BankAccounts As FromBankAccounts, BankAccounts As ToBankAccounts Where (Payments.EmployeeID=Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptID=Concepts.ConceptID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Payments.FromAccountID=FromBankAccounts.AccountID) And (Payments.AccountID=ToBankAccounts.AccountID) And (Payments.PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".RecordDate=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesBeneficiariesLKP.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesBeneficiariesLKP.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Group By PaymentID, Payments.PaymentDate, CheckAmount, EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesBeneficiariesLKP.BeneficiaryNumber, BeneficiaryName, BeneficiaryLastName, BeneficiaryLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID, Areas2.AreaCode, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber, ToBankAccounts.BankID, ToBankAccounts.AccountNumber, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction Order by PaymentID, IsDeduction, ConceptShortName", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select PaymentID, Payments.PaymentDate, CheckAmount, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeID, EmployeesBeneficiariesLKP.BeneficiaryNumber As EmployeeNumber, BeneficiaryName As EmployeeName, BeneficiaryLastName As EmployeeLastName, BeneficiaryLastName2 As EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID As AreaID1, Areas2.AreaCode As AreaCode2, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber As FromAccountNumber, ToBankAccounts.BankID As ToBankID, ToBankAccounts.AccountNumber As ToAccountNumber, Min(EmployeesChangesLKP.FirstDate) As MinDate, Max(EmployeesChangesLKP.LastDate) As MaxDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, '' As ConceptRetention, Sum(Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptAmount) As TotalAmount From Payments, EmployeesBeneficiariesLKP, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", Concepts, EmployeesChangesLKP, EmployeesHistoryListForPayroll, Employees, Positions, GroupGradeLevels, Levels, Areas As Areas1, Areas As Areas2, Zones, BankAccounts As FromBankAccounts, BankAccounts As ToBankAccounts Where (Payments.EmployeeID=Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptID=Concepts.ConceptID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesBeneficiariesLKP.BeneficiaryNumber) And (EmployeesBeneficiariesLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesBeneficiariesLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Payments.FromAccountID=FromBankAccounts.AccountID) And (Payments.AccountID=ToBankAccounts.AccountID) And (Payments.PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".RecordDate=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesBeneficiariesLKP.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesBeneficiariesLKP.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Group By PaymentID, Payments.PaymentDate, CheckAmount, EmployeesBeneficiariesLKP.BeneficiaryNumber, EmployeesBeneficiariesLKP.BeneficiaryNumber, BeneficiaryName, BeneficiaryLastName, BeneficiaryLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID, Areas2.AreaCode, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber, ToBankAccounts.BankID, ToBankAccounts.AccountNumber, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction Order by PaymentID, IsDeduction, ConceptShortName -->" & vbNewLine
		ElseIf bCreditor Then
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select PaymentID, Payments.PaymentDate, CheckAmount, EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesCreditorsLKP.CreditorNumber As EmployeeNumber, CreditorName As EmployeeName, CreditorLastName As EmployeeLastName, CreditorLastName2 As EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID As AreaID1, Areas2.AreaCode As AreaCode2, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber As FromAccountNumber, ToBankAccounts.BankID As ToBankID, ToBankAccounts.AccountNumber As ToAccountNumber, Min(EmployeesChangesLKP.FirstDate) As MinDate, Max(EmployeesChangesLKP.LastDate) As MaxDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, '' As ConceptRetention, Sum(Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptAmount) As TotalAmount From Payments, EmployeesCreditorsLKP, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", Concepts, EmployeesChangesLKP, EmployeesHistoryListForPayroll, Employees, Positions, GroupGradeLevels, Levels, Areas As Areas1, Areas As Areas2, Zones, BankAccounts As FromBankAccounts, BankAccounts As ToBankAccounts Where (Payments.EmployeeID=Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptID=Concepts.ConceptID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Payments.FromAccountID=FromBankAccounts.AccountID) And (Payments.AccountID=ToBankAccounts.AccountID) And (Payments.PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".RecordDate=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesCreditorsLKP.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesCreditorsLKP.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Group By PaymentID, Payments.PaymentDate, CheckAmount, EmployeesCreditorsLKP.CreditorNumber, EmployeesCreditorsLKP.CreditorNumber, CreditorName, CreditorLastName, CreditorLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID, Areas2.AreaCode, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber, ToBankAccounts.BankID, ToBankAccounts.AccountNumber, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction Order by PaymentID, IsDeduction, ConceptShortName", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select PaymentID, Payments.PaymentDate, CheckAmount, EmployeesCreditorsLKP.CreditorNumber As EmployeeID, EmployeesCreditorsLKP.CreditorNumber As EmployeeNumber, CreditorName As EmployeeName, CreditorLastName As EmployeeLastName, CreditorLastName2 As EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID As AreaID1, Areas2.AreaCode As AreaCode2, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber As FromAccountNumber, ToBankAccounts.BankID As ToBankID, ToBankAccounts.AccountNumber As ToAccountNumber, Min(EmployeesChangesLKP.FirstDate) As MinDate, Max(EmployeesChangesLKP.LastDate) As MaxDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, '' As ConceptRetention, Sum(Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptAmount) As TotalAmount From Payments, EmployeesCreditorsLKP, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", Concepts, EmployeesChangesLKP, EmployeesHistoryListForPayroll, Employees, Positions, GroupGradeLevels, Levels, Areas As Areas1, Areas As Areas2, Zones, BankAccounts As FromBankAccounts, BankAccounts As ToBankAccounts Where (Payments.EmployeeID=Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptID=Concepts.ConceptID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesCreditorsLKP.CreditorNumber) And (EmployeesCreditorsLKP.EmployeeID=EmployeesChangesLKP.EmployeeID) And (EmployeesChangesLKP.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesCreditorsLKP.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Payments.FromAccountID=FromBankAccounts.AccountID) And (Payments.AccountID=ToBankAccounts.AccountID) And (Payments.PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".RecordDate=EmployeesChangesLKP.PayrollDate) And (EmployeesChangesLKP.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesCreditorsLKP.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesCreditorsLKP.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (FromBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (ToBankAccounts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Group By PaymentID, Payments.PaymentDate, CheckAmount, EmployeesCreditorsLKP.CreditorNumber, EmployeesCreditorsLKP.CreditorNumber, CreditorName, CreditorLastName, CreditorLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID, Areas2.AreaCode, ZonePath, Payments.CheckNumber, FromBankAccounts.AccountNumber, ToBankAccounts.BankID, ToBankAccounts.AccountNumber, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction Order by PaymentID, IsDeduction, ConceptShortName -->" & vbNewLine
		Else
			lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select Payments.PaymentID, Payments.PaymentDate, CheckAmount, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, EmployeesHistoryListForPayroll.BankID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID As AreaID1, Areas2.AreaCode As AreaCode2, ZonePath, Payments.CheckNumber, EmployeesHistoryListForPayroll.AccountNumber As ToAccountNumber, DatECHG.FirstDate MinDate, DatECHG.LastDate MaxDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Concepts.OrderInList, ConceptRetention, Sum(ConceptAmount) As TotalAmount From Payments, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", Concepts, EmployeesHistoryListForPayroll, (Select pymnt.PaymentID, Min(echg.FirstDate) FirstDate, Max(echg.LastDate) LastDate From Payments pymnt, EmployeesChangesLKP  echg Where (pymnt.PaymentDate = " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (echg.PayrollID = " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (pymnt.EmployeeID = echg.EmployeeID) And (pymnt.PaymentDate = echg.PayrollID) Group By pymnt.PaymentID) DatECHG, Employees, Positions, GroupGradeLevels, Levels, Areas As Areas1, Areas As Areas2, Zones Where (Payments.EmployeeID=Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptID=Concepts.ConceptID) And (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payments.PaymentID = DatECHG.PaymentID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Payments.PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Concepts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Concepts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Group By Payments.PaymentID, Payments.PaymentDate, CheckAmount, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, SocialSecurityNumber,BankID, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID, Areas2.AreaCode, ZonePath, Payments.CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, DatECHG.FirstDate, DatECHG.LastDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Concepts.OrderInList, ConceptRetention Order by Payments.PaymentID, IsDeduction, Concepts.OrderInList", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
			Response.Write vbNewLine & "<!-- Query: Select Payments.PaymentID, Payments.PaymentDate, CheckAmount, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID As AreaID1, Areas2.AreaCode As AreaCode2, ZonePath, Payments.CheckNumber, EmployeesHistoryListForPayroll.AccountNumber As ToAccountNumber, DatECHG.FirstDate MinDate, DatECHG.LastDate MaxDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Concepts.OrderInList, ConceptRetention, Sum(ConceptAmount) As TotalAmount From Payments, Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ", Concepts, EmployeesHistoryListForPayroll, (Select pymnt.PaymentID, Min(echg.FirstDate) FirstDate, Max(echg.LastDate) LastDate From Payments pymnt, EmployeesChangesLKP  echg Where (pymnt.PaymentDate = " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (echg.PayrollID = " & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (pymnt.EmployeeID = echg.EmployeeID) And (pymnt.PaymentDate = echg.PayrollID) Group By pymnt.PaymentID) DatECHG, Employees, Positions, GroupGradeLevels, Levels, Areas Areas1, Areas Areas2, Zones Where (Payments.EmployeeID=Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".ConceptID=Concepts.ConceptID) And (Payments.EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (Payments.PaymentID = DatECHG.PaymentID) And (Payroll_" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ".EmployeeID=EmployeesHistoryListForPayroll.EmployeeID) And (EmployeesHistoryListForPayroll.PayrollID=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (EmployeesHistoryListForPayroll.EmployeeID=Employees.EmployeeID) And (EmployeesHistoryListForPayroll.PositionID=Positions.PositionID) And (EmployeesHistoryListForPayroll.GroupGradeLevelID=GroupGradeLevels.GroupGradeLevelID) And (EmployeesHistoryListForPayroll.LevelID=Levels.LevelID) And (EmployeesHistoryListForPayroll.PaymentCenterID=Areas2.AreaID) And (Areas2.ParentID=Areas1.AreaID) And (Areas2.ZoneID=Zones.ZoneID) And (Payments.PaymentDate=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Positions.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (GroupGradeLevels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Levels.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas1.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Areas2.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Concepts.StartDate<=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") And (Concepts.EndDate>=" & aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(1) & ") " & sCondition & " Group By Payments.PaymentID, Payments.PaymentDate, CheckAmount, Employees.EmployeeID, EmployeesHistoryListForPayroll.EmployeeNumber, EmployeeName, EmployeeLastName, EmployeeLastName2, RFC, CURP, SocialSecurityNumber, Employees.StartDate, EmployeesHistoryListForPayroll.CompanyID, EmployeesHistoryListForPayroll.EmployeeTypeID, Positions.PositionID, PositionShortName, PositionName, GroupGradeLevelShortName, LevelShortName, Areas1.AreaID, Areas2.AreaCode, ZonePath, Payments.CheckNumber, EmployeesHistoryListForPayroll.AccountNumber, DatECHG.FirstDate, DatECHG.LastDate, Concepts.ConceptID, ConceptShortName, ConceptName, IsDeduction, Concepts.OrderInList, ConceptRetention Order by Payments.PaymentID, IsDeduction, Concepts.OrderInList -->" & vbNewLine
		End If
	End If
		If lErrorNumber = 0 Then
            bBackReport = true
            bPage = false
            bIsCheck=false
            sEmployee1Information = ""
            sEmployee2Information = ""
			If Not oRecordset.EOF Then
				bFirstEmployee = True
				If lErrorNumber = 0 Then
					Response.Write "<IFRAME SRC=""CheckFile.asp?FileName=" & Server.URLEncode(sFileName) & """ NAME=""CheckFileIFrame"" FRAMEBORDER=""0"" WIDTH=""100%"" HEIGHT=""140""></IFRAME>"
					Response.Flush()

					sCurrentNumber = CStr(oRecordset.Fields("CheckNumber").Value)
					lMinDate = 30000000
					lMaxDate = 0
                    
					Do While Not oRecordset.EOF
						If lMinDate > oRecordset.Fields("MinDate").Value Then lMinDate = oRecordset.Fields("MinDate").Value
						If lMaxDate < oRecordset.Fields("MaxDate").Value Then lMaxDate = oRecordset.Fields("MaxDate").Value
						If StrComp(sCurrentNumber, CStr(oRecordset.Fields("CheckNumber").Value), vbBinaryCompare) <> 0 Then Exit Do
						oRecordset.MoveNext29
						If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do
					Loop
					oRecordset.MoveFirst
					lEmployeeCounter = 0
					lFileCounter = 0
                    Dim band
                    Dim employeeNumber
                    Dim asPathEmployee1
                    Dim asPathEmployee2
					sCurrentNumber = ""
					lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\rtf1 \ansi \deff0 {\fonttbl {\f0\fmodern Tahoma;}}\fs18", sErrorDescription)
					If Not FileExists(sFilePath & "\" & sPensionISSSTELogo, sErrorDescription) Then
						lErrorNumber = CopyFile(sImagesPath & "\" & sPensionISSSTELogo, sFilePath & "\" & sPensionISSSTELogo, sErrorDescription)
					End If
                    lcurrentEmployeeID=oRecordset.Fields("EmployeeNumber").Value 
                    band=false
                    bStatusReport = true
                    If Not oRecordset.EOF Then 
					    Do While bStatusReport
                          employeeNumber = CStr(oRecordset.Fields("EmployeeNumber").Value)
                            If StrComp(lcurrentEmployeeID, CStr(oRecordset.Fields("EmployeeNumber").Value), vbBinaryCompare) <> 0  Then 'Indica cuando se cambia de empleado
                               band = true
                            End If
                            If (band And bPage And Not bIsCheck) OR (band And oRecordset.EOF And Not bIsCheck) OR (band And bIsCheck)  Then 'Escribir parte posterior de los mensajes
                                
                                If  Len(sEmployee2Information) =0 Then
                                  sEmployee2information="******" 
                                  asPathEmployee2=",,,," 
                                End If
                                sEmployee1Information = Split(sEmployee1Information, "*")  
                                sEmployee2Information = Split(sEmployee2Information,"*")
                                
                                asPathEmployee1 = Split(CStr(sEmployee1Information(11)), ",")
                                asPathEmployee2 = Split(CStr(sEmployee2Information(11)), ",")
                   
                                If Not bIsCheck  Then
                                    
                                    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx-10 \posy-10 \absw12500 \absh14399 \par}", sErrorDescription)
					                lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\sbkpage{\*\atnid S A L T O  D E  S E C C I O N}", sErrorDescription)
					                lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\sect\sectd{\*\atnid N U E V A  S E C C I O N}", sErrorDescription)

                                    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(1500) & "\posy" & CInt(600) & " \absw9000{\*\atnid LOGO_ISSSTE}", sErrorDescription)
						            lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\field\fldedit{\*\fldinst { INCLUDEPICTURE \\d", sErrorDescription)
						            lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\~"&sPensionISSSTELogo, sErrorDescription)
						            lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\\* MERGEFORMATINET }}{\fldrslt { }}}", sErrorDescription)
						            lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
                                    If Not StrComp(CStr(sEmployee2Information(0)), "", vbBinaryCompare) = 0 then
                                        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(1500) & "\posy" & CInt(8500) & " \absw9000{\*\atnid LOGO_ISSSTE}", sErrorDescription)
						                lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\field\fldedit{\*\fldinst { INCLUDEPICTURE \\d", sErrorDescription)
						                lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\~"&sPensionISSSTELogo, sErrorDescription)
						                lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\\* MERGEFORMATINET }}{\fldrslt { }}}", sErrorDescription)
						                lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
                                    End If
                                End If
                               
                            'Mensajes Empleado 1
                            sContents = ""
							For iIndex = 0 To UBound(asMessages) - 1
								If ((StrComp(asMessages(iIndex)(1), ",0,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(1), "," & CStr(sEmployee1Information(0)) & ",", vbBinaryCompare) > 0)) And _
								   ((StrComp(asMessages(iIndex)(2), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(2), "," & CStr(sEmployee1Information(7)) & ",", vbBinaryCompare) > 0)) And _
								   ((StrComp(asMessages(iIndex)(3), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(3), "," & CStr(sEmployee1Information(8)) & ",", vbBinaryCompare) > 0)) And _
								   ((StrComp(asMessages(iIndex)(4), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(4), "," & asPathEmployee1(2) & ",", vbBinaryCompare) > 0)) And _
								   ((StrComp(asMessages(iIndex)(5), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(5), "," & CStr(sEmployee1Information(9)) & ",", vbBinaryCompare) > 0)) And _
								   ((StrComp(asMessages(iIndex)(6), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(6), "," & CStr(sEmployee1Information(10)) & ",", vbBinaryCompare) > 0)) Then 'And _
								   '((StrComp(asMessages(iIndex)(7), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(7), "," & CStr(oRecordset.Fields("ToBankID").Value) & ",", vbBinaryCompare) > 0)) Then
									Select Case CInt(asMessages(iIndex)(8))
										    Case 0
											    If StrComp(sEmployee1Information(1), ".", vbBinaryCompare) = 0 OR StrComp(sAccountEmployee2, ".", vbBinaryCompare) = 0 Then sContents = sContents & asMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
										    Case 1
											    If StrComp(sEmployee1Information(1), ".", vbBinaryCompare) <> 0 OR StrComp(sAccountEmployee2, ".", vbBinaryCompare) = 0 Then sContents = sContents & asMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
										    Case 2
											    If bAlimony Then sContents = sContents & asMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
										    Case 4
											    If bCreditor Then sContents = sContents & asMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
										    Case Else
											    sContents = sContents & asMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
									    End Select
								End If
							Next
							For iIndex = 0 To UBound(asEmployeesMessages) - 1
								If StrComp(asEmployeesMessages(iIndex)(1), CStr(sEmployee1Information(0)), vbBinaryCompare) = 0 Then
									sContents = sContents & asEmployeesMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
								End If
							Next

                                If bIsCheck Then                                    
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(900 + x1PositionDisplacement) & "\posy" & CInt(9200 + y1PositionDisplacement) & " \absw9000{\*\atnid DESC_PUESTO}", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), sContents, sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
                                Else
                                    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(900 + x1PositionDisplacement) & "\posy" & CInt(3700 ) & " \absw9000{\*\atnid DESC_PUESTO}", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), sContents, sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
                                     
                                End IF
                               
							    'Mensajes del empleado 2 
                                If Not StrComp(sEmployee2Information(0), "", vbBinaryCompare) = 0 then
                                
                                    sContents = ""
							        For iIndex = 0 To UBound(asMessages) - 1
								        If ((StrComp(asMessages(iIndex)(1), ",0,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(1), "," & CStr(sEmployee2Information(0)) & ",", vbBinaryCompare) > 0)) And _
								   ((StrComp(asMessages(iIndex)(2), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(2), "," & CStr(sEmployee2Information(7)) & ",", vbBinaryCompare) > 0)) And _
								   ((StrComp(asMessages(iIndex)(3), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(3), "," & CStr(sEmployee2Information(8)) & ",", vbBinaryCompare) > 0)) And _
								   ((StrComp(asMessages(iIndex)(4), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(4), "," & asPathEmployee2(2) & ",", vbBinaryCompare) > 0)) And _
								   ((StrComp(asMessages(iIndex)(5), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(5), "," & CStr(sEmployee2Information(9)) & ",", vbBinaryCompare) > 0)) And _
								   ((StrComp(asMessages(iIndex)(6), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(6), "," & CStr(sEmployee2Information(10)) & ",", vbBinaryCompare) > 0))Then 'And _
								   '((StrComp(asMessages(iIndex)(7), ",-1,", vbBinaryCompare) = 0) Or (InStr(1, asMessages(iIndex)(7), "," & CStr(oRecordset.Fields("ToBankID").Value) & ",", vbBinaryCompare) > 0)) Then
									        Select Case CInt(asMessages(iIndex)(8))
										            Case 0
											            If StrComp(sEmployee2Information(1), ".", vbBinaryCompare) = 0 OR StrComp(sAccountEmployee2, ".", vbBinaryCompare) = 0 Then sContents = sContents & asMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
										            Case 1
											            If StrComp(sEmployee2Information(1), ".", vbBinaryCompare) <> 0 OR StrComp(sAccountEmployee2, ".", vbBinaryCompare) = 0 Then sContents = sContents & asMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
										            Case 2
											            If bAlimony Then sContents = sContents & asMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
										            Case 4
											            If bCreditor Then sContents = sContents & asMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
										            Case Else
											            sContents = sContents & asMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
									            End Select
								        End If
							        Next
							        For iIndex = 0 To UBound(asEmployeesMessages) - 1
								        If StrComp(asEmployeesMessages(iIndex)(1), CStr(sEmployee2Information(0)), vbBinaryCompare) = 0 Then
									        sContents = sContents & asEmployeesMessages(iIndex)(0) & "\" & Chr(13) & Chr(10)
								        End If
							        Next                                 
                                    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(900 + x1PositionDisplacement) & "\posy" & CInt(3500 + y1PositionDisplacement) & " \absw9000{\*\atnid DESC_PUESTO}", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), sContents, sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
                                End If
                                                               
                                asPath = Split(CStr(oRecordset.Fields("ZonePath").Value), ",")
							    If bIsCheck Then
								    If bAlimony Or bCreditor Then
									    sSignatureName1 = "0901.jpg"
									    If Not FileExists(sFilePath & "\" & sSignatureName, sErrorDescription) Then
										    lErrorNumber = CopyFile(sImagesPath & "\" & sSignatureName1, sFilePath & "\" & sSignatureName1, sErrorDescription)
									    End If
									    sSignatureName2 = "0902.jpg"
									    If Not FileExists(sFilePath & "\" & sSignatureName, sErrorDescription) Then
										    lErrorNumber = CopyFile(sImagesPath & "\" & sSignatureName2, sFilePath & "\" & sSignatureName2, sErrorDescription)
									    End If
								    ElseIf CLng(oRecordset.Fields("AreaID1").Value) <> 38 Then
									    sSignatureName1 = Right(("00" & asPath(2)), Len("00")) & "01.jpg"
									    If Not FileExists(sFilePath & "\" & sSignatureName1, sErrorDescription) Then
										    lErrorNumber = CopyFile(sImagesPath & "\" & sSignatureName1, sFilePath & "\" & sSignatureName1, sErrorDescription)
									    End If
									    sSignatureName2 = Right(("00" & asPath(2)), Len("00")) & "02.jpg"
									    If Not FileExists(sFilePath & "\" & sSignatureName2, sErrorDescription) Then
										    lErrorNumber = CopyFile(sImagesPath & "\" & sSignatureName2, sFilePath & "\" & sSignatureName2, sErrorDescription)
									    End If
								    Else
									    sSignatureName1 = "3801.jpg"
									    If Not FileExists(sFilePath & "\" & sSignatureName, sErrorDescription) Then
										    lErrorNumber = CopyFile(sImagesPath & "\" & sSignatureName1, sFilePath & "\" & sSignatureName1, sErrorDescription)
									    End If
									    sSignatureName2 = "3802.jpg"
									    If Not FileExists(sFilePath & "\" & sSignatureName, sErrorDescription) Then
										    lErrorNumber = CopyFile(sImagesPath & "\" & sSignatureName2, sFilePath & "\" & sSignatureName2, sErrorDescription)
									    End If
								    End If
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(6323 + x2PositionDisplacement) & "\posy" & CInt(13850 + lSignaturesPositionDisplacementR + y2PositionDisplacement) & " \absw2778{\*\atnid FIRMA1}", sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\field\fldedit{\*\fldinst { INCLUDEPICTURE \\d", sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\~"&sSignatureName1, sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\\* MERGEFORMATINET }}{\fldrslt { }}}", sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(9100 + x2PositionDisplacement) & "\posy" & CInt(13850 + lSignaturesPositionDisplacementR + y2PositionDisplacement) & " \absw2778{\*\atnid FIRMA2}", sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\field\fldedit{\*\fldinst { INCLUDEPICTURE \\d", sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\~"&sSignatureName2, sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\\* MERGEFORMATINET }}{\fldrslt { }}}", sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
                                    
                                    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(8300 + x2PositionDisplacement) & "\posy" & CInt(11700 + y2PositionDisplacement) & " \absw2500{\*\atnid NO.CHEQUE}", sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(SizeText(CStr(sEmployee1Information(3)), "0", 10, 0)), sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
                                   

                                    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(7100 + x2PositionDisplacement) & "\posy" & CInt(12000 + y2PositionDisplacement) & " \absw2156{\*\atnid FECHA_NOMINA}", sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), sEmployee1Information(5), sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10200 + x2PositionDisplacement) & "\posy" & CInt(12850 + y2PositionDisplacement) & " \absw2000{\*\atnid MONEDA_NACIONAL}", sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), sEmployee1Information(4), sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(2212 + x2PositionDisplacement) & "\posy" & CInt(12800 + y2PositionDisplacement) & " \absw8500{\*\atnid NOMBRE_CHEQUE}", sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs18 \b", sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), sEmployee1Information(2), sErrorDescription)
								    
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(1200 + x2PositionDisplacement) & "\posy" & CInt(13250 + y2PositionDisplacement) & " \absw7950{\*\atnid MONEDA_NACIONAL_TEXTO}", sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter),sEmployee1Information(6), sErrorDescription)
								    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
        
                                End If

                                If oRecordset.EOF  Then
                                    bStatusReport = false                                  
                                End If 

                                band = false     
                                bPage = false  
                                sEmployeeId1 = 0
                                sEmployeeId2 = 0 
                                sAccountEmployee1 = ""
                                sAccountEmployee2 = ""
                                sContents=""
                                sEmployee1Information=""
                                sEmployee2Information=""
                                asPathEmployee1=""
                                asPathEmployee2=""
                                bIsCheck = false
                                
                            Else
                                 If StrComp(sCurrentNumber, CStr(oRecordset.Fields("CheckNumber").Value), vbBinaryCompare) <> 0 Then
							        lEmployeeCounter = lEmployeeCounter + 1
                                    lcurrentEmployeeID=oRecordset.Fields("EmployeeNumber").Value
                                    If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then
                                        bIsCheck = true
                                        band = false
                                    End If
                                    
                                    If StrComp(sEmployee1Information, "", vbBinaryCompare) = 0 Then 'Información del empleado 1
                                        If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
                                            sEmployee1Information = CStr(oRecordset.Fields("EmployeeNumber").Value) & "*" & CStr(oRecordset.Fields("ToAccountNumber").Value) & "*" & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value) & "*" & CStr(oRecordset.Fields("CheckNumber").Value) & "*" & FormatNumber(CDbl(oRecordset.Fields("CheckAmount").Value)) & "*" & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PaymentDate").Value),-1,-1,-1) & "*" & UCase(FormatNumberAsText(CDbl(oRecordset.Fields("CheckAmount").Value),true)) &"*"& oRecordset.Fields("CompanyID").Value &"*"& oRecordset.Fields("AreaID1").Value &"*"& oRecordset.Fields("EmployeeTypeID").Value &"*"& oRecordset.Fields("PositionID").Value &"*"& oRecordset.Fields("ZonePath").Value
								        Else
                                            sEmployee1Information = oRecordset.Fields("EmployeeNumber").Value & "*" &oRecordset.Fields("ToAccountNumber").Value & "*" & oRecordset.Fields("EmployeeLastName").Value & " " & oRecordset.Fields("EmployeeName").Value & "*" & CStr(oRecordset.Fields("CheckNumber").Value) & "*" &  FormatNumber(CDbl(oRecordset.Fields("CheckAmount").Value)) & "*" & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PaymentDate")),-1,-1,-1) & "*" & UCase(FormatNumberAsText(CDbl(oRecordset.Fields("CheckAmount").Value),true))&"*"& oRecordset.Fields("CompanyID").Value &"*"& oRecordset.Fields("AreaID1").Value &"*"& oRecordset.Fields("EmployeeTypeID").Value &"*"& oRecordset.Fields("PositionID").Value &"*"& oRecordset.Fields("ZonePath").Value
							            End If
                                       
                                    ElseIf StrComp(sEmployee2Information, "", vbBinaryCompare) = 0 Then 'Información del empleado 1
                                         If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
                                            sEmployee2Information = CStr(oRecordset.Fields("EmployeeNumber").Value) & "*" & CStr(oRecordset.Fields("ToAccountNumber").Value) & "*" & CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value) & "*" & CStr(oRecordset.Fields("CheckNumber").Value) & "*" & FormatNumber(CDbl(oRecordset.Fields("CheckAmount").Value)) & "*" & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PaymentDate").Value),-1,-1,-1) & "*" & UCase(FormatNumberAsText(CDbl(oRecordset.Fields("CheckAmount").Value),true)) &"*"& oRecordset.Fields("CompanyID").Value &"*"& oRecordset.Fields("AreaID1").Value &"*"& oRecordset.Fields("EmployeeTypeID").Value &"*"& oRecordset.Fields("PositionID").Value  &"*"& oRecordset.Fields("ZonePath").Value
								        Else
                                             sEmployee2Information = oRecordset.Fields("EmployeeNumber").Value & "*" &oRecordset.Fields("ToAccountNumber").Value & "*" & oRecordset.Fields("EmployeeLastName").Value & " " & oRecordset.Fields("EmployeeName").Value & "*" & CStr(oRecordset.Fields("CheckNumber").Value) & "*" &  FormatNumber(CDbl(oRecordset.Fields("CheckAmount").Value)) & "*" & DisplayDateFromSerialNumber(CLng(oRecordset.Fields("PaymentDate")),-1,-1,-1) & "*" & UCase(FormatNumberAsText(CDbl(oRecordset.Fields("CheckAmount").Value),true))&"*"& oRecordset.Fields("CompanyID").Value &"*"& oRecordset.Fields("AreaID1").Value &"*" & oRecordset.Fields("EmployeeTypeID").Value &"*"& oRecordset.Fields("PositionID").Value &"*"& oRecordset.Fields("ZonePath").Value
							            End If

                                    End If
                                    
							        If (lEmployeeCounter Mod CHECKS_PER_FILE) = 0 Then
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx-50 \posy-50 \absw12500 \absh14399 \par}", sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "}", sErrorDescription)
								        lFileCounter = Int(lEmployeeCounter / CHECKS_PER_FILE)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\rtf1 \ansi \deff0 {\fonttbl {\f0\fmodern Tahoma;}}\fs18", sErrorDescription)
								        
							        End If
							        yPositionForConceptsP=2700
							        yPositionForConceptsD=2700
                                    If (lEmployeeCounter Mod 2)=0 And Not bFirstEmployee And Not bIsCheck Then
                                         y1PositionDisplacement = y1PositionDisplacement+7948
                                         bPage = true
                                         band=false
                                         
                                    ElseIf Not bFirstEmployee Then
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx-10 \posy-10 \absw12500 \absh14399 \par}", sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\sbkpage{\*\atnid S A L T O  D E  S E C C I O N}", sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\sect\sectd{\*\atnid N U E V A  S E C C I O N}", sErrorDescription)
                                        If not bIsCheck Then
                                            y1PositionDisplacement = y1PositionDisplacement-7948
                                        End If
                                    
                                    End If

                                    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(500 + x1PositionDisplacement) & "\posy" & CInt(550 + l01PositionDisplacement + y1PositionDisplacement) & " \absw1247{\*\atnid NO.EMP.}", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(CStr(oRecordset.Fields("EmployeeNumber").Value)), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(2200 + x1PositionDisplacement) & "\posy" & CInt(550 + l01PositionDisplacement + y1PositionDisplacement) & " \absw8000{\*\atnid NOMBRE}", sErrorDescription)
							        If Not IsNull(oRecordset.Fields("EmployeeLastName2").Value) Then
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeLastName2").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1)), sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(SizeText(CStr(oRecordset.Fields("EmployeeLastName").Value) & " " & CStr(oRecordset.Fields("EmployeeName").Value), " ", 70, 1)), sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
                                    If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(8300 + x2PositionDisplacement) & "\posy" & CInt(550 + l01PositionDisplacement + y1PositionDisplacement)  & " \absw2500{\*\atnid NO.CHEQUE}", sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(SizeText(CStr(oRecordset.Fields("CheckNumber").Value), "0", 10, 0)), sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        Else 
							            lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10000 + x1PositionDisplacement) & "\posy" & CInt(550 + l01PositionDisplacement + y1PositionDisplacement) & " \absw2500{\*\atnid NO.CHEQUE}", sErrorDescription)
							            lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(SizeText(CStr(oRecordset.Fields("CheckNumber").Value), "0", 10, 0)), sErrorDescription)
							            lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(500 + x1PositionDisplacement) & "\posy" & CInt(1050 + l02PositionDisplacement + y1PositionDisplacement) & " \absw2240{\*\atnid RFC}", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(CStr(oRecordset.Fields("RFC").Value)), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(4000 + x1PositionDisplacement) & "\posy" & CInt(1050 + l02PositionDisplacement + y1PositionDisplacement) & " \absw5075{\*\atnid CURP}", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(CStr(oRecordset.Fields("CURP").Value)), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(9000 + x1PositionDisplacement) & "\posy" & CInt(1050 + l02PositionDisplacement + y1PositionDisplacement) & " \absw3856{\*\atnid NSS}", sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(CStr(oRecordset.Fields("SocialSecurityNumber").Value)), sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(7740 + x1PositionDisplacement) & "\posy" & CInt(1050 + l02PositionDisplacement + y1PositionDisplacement) + l01PositionDisplacement & " \absw3856{\*\atnid NSS}", sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(CStr(oRecordset.Fields("SocialSecurityNumber").Value)), sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        End If
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' CLAVE PUESTO
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(500  + x1PositionDisplacement) & "\posy" & CInt(1500 + l03PositionDisplacement + y1PositionDisplacement) & " \absw2240{\*\atnid CLAVE_PUESTO}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(500  + x1PositionDisplacement) & "\posy" & CInt(1600 + l03PositionDisplacementR + y1PositionDisplacement) & " \absw2240{\*\atnid CLAVE_PUESTO}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), SizeText(CStr(oRecordset.Fields("PositionShortName").Value), " ", 7, 1), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' DESC_PUESTO
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(2900 + x1PositionDisplacement) & "\posy" & CInt(1500 + l03PositionDisplacement + y1PositionDisplacement) & " \absw3515{\*\atnid DESC_PUESTO}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(2300 + x1PositionDisplacement) & "\posy" & CInt(1600 + l03PositionDisplacementR + y1PositionDisplacement) & " \absw4115{\*\atnid DESC_PUESTO}", sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs14", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), SizeText(CStr(oRecordset.Fields("PositionName").Value), " ", 60, 1), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' NIV_SUB
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(6110 + x1PositionDisplacement) & "\posy" & CInt(1500 + l03PositionDisplacement + y1PositionDisplacement) & " \absw737{\*\atnid NIV_SUB}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(6210 + x1PositionDisplacement) & "\posy" & CInt(1600 + l03PositionDisplacementR + y1PositionDisplacement) & " \absw637{\*\atnid NIV_SUB}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(Left(CStr(oRecordset.Fields("LevelShortName").Value), Len("00")) & " / " & Right(CStr(oRecordset.Fields("LevelShortName").Value), Len("0"))), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' RG_MX
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(7010  + x1PositionDisplacement) & "\posy" & CInt(1500 + l03PositionDisplacement + y1PositionDisplacement) & " \absw567{\*\atnid RG_MX}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(7010  + x1PositionDisplacement) & "\posy" & CInt(1600 + l03PositionDisplacementR + y1PositionDisplacement) & " \absw567{\*\atnid RG_MX}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "2/0", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' CG_RP
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(7610- + x1PositionDisplacement) & "\posy" & CInt(1500 + l03PositionDisplacement + y1PositionDisplacement) & " \absw567{\*\atnid CG_RP}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(7610 + x1PositionDisplacement) & "\posy" & CInt(1600 + l03PositionDisplacementR + y1PositionDisplacement) & " \absw567{\*\atnid CG_RP}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "/X", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' CLAVE_PRESUPUESTAL
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(8600 + x1PositionDisplacement) & "\posy" & CInt(1500 + l03PositionDisplacement + y1PositionDisplacement) & " \absw1500{\*\atnid CLAVE_PRESUPUESTAL}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(8600 + x1PositionDisplacement) & "\posy" & CInt(1600 + l03PositionDisplacementR + y1PositionDisplacement) & " \absw1500{\*\atnid CLAVE_PRESUPUESTAL}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(CStr(oRecordset.Fields("AreaCode2").Value)), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' CLAVE_DISTRIBUCION
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10400 + x1PositionDisplacement) & "\posy" & CInt(1500 + l03PositionDisplacement + y1PositionDisplacement) & " \absw1500{\*\atnid CLAVE_DISTRIBUCION}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10400 + x1PositionDisplacement) & "\posy" & CInt(1600 + l03PositionDisplacementR + y1PositionDisplacement) & " \absw1500{\*\atnid CLAVE_DISTRIBUCION}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CleanStringForHTML(CStr(oRecordset.Fields("AreaCode2").Value)), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' FECHA_INGRESO
						    		    lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(500 + x1PositionDisplacement) & "\posy" & CInt(1950 + l04PositionDisplacement + y1PositionDisplacement) & " \absw1956{\*\atnid FECHA_INGRESO}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt( 500 + x1PositionDisplacement) & "\posy" & CInt(2100 + l04PositionDisplacementR + y1PositionDisplacement) & " \absw1956{\*\atnid FECHA_INGRESO}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), DisplayNumericDateFromSerialNumber(CLng(oRecordset.Fields("StartDate").Value)), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' FECHA_NOMINA
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(2900 + x1PositionDisplacement) & "\posy" & CInt(1950 + l04PositionDisplacement + y1PositionDisplacement) & " \absw2256{\*\atnid FECHA_NOMINA}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(2900 + x1PositionDisplacement) & "\posy" & CInt(2100 + l04PositionDisplacementR + y1PositionDisplacement) & " \absw2256{\*\atnid FECHA_NOMINA}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), DisplayNumericDateFromSerialNumber(lPayrollID), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' PERIODO_PAGO
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(4500 + x1PositionDisplacement) & "\posy" & CInt(1950 + l04PositionDisplacement + y1PositionDisplacement) & " \absw3742{\*\atnid PERIODO_PAGO}", sErrorDescription)
							        Else
							            lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(4500 + x1PositionDisplacement) & "\posy" & CInt(2100 + l04PositionDisplacementR + y1PositionDisplacement) & " \absw3742{\*\atnid PERIODO_PAGO}", sErrorDescription)
							        End If
							        If lMinDate = 0 Then lMinDate = lPayrollID End If
							        If lMaxDate = 0 Then lMaxDate = lPayrollID End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), DisplayNumericDateFromSerialNumber(lMinDate) & " AL " & DisplayNumericDateFromSerialNumber(lMaxDate), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' NETO_A_PAGAR
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10200 + x1PositionDisplacement) & "\posy" & CInt(1950 + l04PositionDisplacement + y1PositionDisplacement) & " \absw1400{\*\atnid NETO_A_PAGAR}", sErrorDescription)
							        Else
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10400 + x1PositionDisplacement) & "\posy" & CInt(2100 + l04PositionDisplacementR + y1PositionDisplacement) & " \absw1400{\*\atnid NETO_A_PAGAR}", sErrorDescription)
							        End If
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\b", sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), FormatNumber(CDbl(oRecordset.Fields("CheckAmount").Value), True), sErrorDescription)
							        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							       
							        sPerceptions = ""
							        sDeductions = ""
							        sCurrentNumber = CStr(oRecordset.Fields("CheckNumber").Value)
							        bFirstEmployee = False
							        lMinDate = 30000000
							        lMaxDate = 0
						        End If
                               
						        Select Case CStr(oRecordset.Fields("ConceptShortName").Value)
							        Case "D"
								        lDeductions = CDbl(oRecordset.Fields("TotalAmount").Value)
								        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' TOTAL_DEDUCCIONES
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10524 + x1PositionDisplacement) & "\posy" & CInt(6600 + lConceptsTPositionDisplacement + y1PositionDisplacement) & " \absw1500{\*\atnid TOTAL_DEDUCCIONES}", sErrorDescription)
								        Else
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10599 + x1PositionDisplacement ) & "\posy" & CInt(6940 + lConceptsTPositionDisplacementR + y1PositionDisplacement) & " \absw1500{\*\atnid TOTAL_DEDUCCIONES}", sErrorDescription)
								        End If
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), FormatNumber(lDeductions, 2, True, False, True), sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
                                    Case "L"
							            '	'sContents = Replace(sContents, "<TOTAL />", FormatNumber(CDbl(oRecordset.Fields("ConceptAmount").Value), 2, True, False, True))
							            '	lConceptAmount = CDbl(oRecordset.Fields("TotalAmount").Value)
							            '	lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(800 + x1PositionDisplacement) & "\posy" & CInt(12814 + y1PositionDisplacement) & " \absw1701{\*\atnid NETO_A_PAGAR}", sErrorDescription)
							            '	lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), FormatNumber(lConceptAmount, 2, True, False, True), sErrorDescription)
							            '	lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        Case "P"
								        lPerceptions = CDbl(oRecordset.Fields("TotalAmount").Value)
								        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' TOTAL_PERCEPCIONES
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(4824 + x1PositionDisplacement) & "\posy" & CInt(6600 + lConceptsTPositionDisplacement + y1PositionDisplacement) & " \absw1500{\*\atnid TOTAL_PERCEPCIONES}", sErrorDescription)
								        Else
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(4924 + x1PositionDisplacement) & "\posy" & CInt(6940 + lConceptsTPositionDisplacementR + y1PositionDisplacement) & " \absw1500{\*\atnid TOTAL_PERCEPCIONES}", sErrorDescription)
						    	        End If
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), FormatNumber(lPerceptions, 2, True, False, True), sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
								        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' BASE_GRAVABLE
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(8400 + x1PositionDisplacement) & "\posy" & CInt(1950 + l04PositionDisplacement + y1PositionDisplacement) & " \absw1400{\*\atnid BASE_GRAVABLE}", sErrorDescription)
								        Else
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(8600 + x1PositionDisplacement) & "\posy" & CInt(2100 + l04PositionDisplacementR + y1PositionDisplacement) & " \absw1400{\*\atnid BASE_GRAVABLE}", sErrorDescription)
								        End If
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), FormatNumber(lPerceptions, 2, True, False, True), sErrorDescription)
								        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
							        Case Else
								        If CInt(oRecordset.Fields("IsDeduction").Value) = 0 Then
									        lConceptAmount = CDbl(oRecordset.Fields("TotalAmount").Value)
									        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' CONCEPTO1
										        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(500+ x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsP + lConceptsPositionDisplacement + y1PositionDisplacement-50) & " \absw737{\*\atnid CAVE.CPTO.} {\*\atnid CONCEPTO1}", sErrorDescription)
									        Else
										        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(500 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsP + 200 + lConceptsPositionDisplacementR + y1PositionDisplacement-50) & " \absw737{\*\atnid CAVE.CPTO.} {\*\atnid CONCEPTO1}", sErrorDescription)
									        End If
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs14", sErrorDescription)
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CStr(oRecordset.Fields("ConceptShortName").Value), sErrorDescription)
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
									        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' DESCRIPCION
										        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(1300 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsP + lConceptsPositionDisplacement + y1PositionDisplacement-50) & " \absw3515{\*\atnid DESCRIPCION}", sErrorDescription)
									        Else
										        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(1300 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsP + 200 + lConceptsPositionDisplacementR + y1PositionDisplacement-50) & " \absw3515{\*\atnid DESCRIPCION}", sErrorDescription)
									        End If
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs14", sErrorDescription)
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CStr(oRecordset.Fields("ConceptName").Value), sErrorDescription)
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
									        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' IMPORTE
										        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(4924 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsP + lConceptsPositionDisplacement + y1PositionDisplacement-50) & " \absw1214{\*\atnid IMPORTE}", sErrorDescription)
									        Else
										        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(4924 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsP + 200 + lConceptsPositionDisplacementR + y1PositionDisplacement-50) & " \absw1214{\*\atnid IMPORTE}", sErrorDescription)
									        End If
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs14", sErrorDescription)
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), FormatNumber(lConceptAmount, 2, True, False, True), sErrorDescription)
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
									        If Len(CStr(oRecordset.Fields("ConceptName").Value)) > 53 Then
										        yPositionForConceptsP = yPositionForConceptsP + 300
									        Else
										        yPositionForConceptsP = yPositionForConceptsP + 150
									        End If
								        Else
									        lConceptAmount = CDbl(oRecordset.Fields("TotalAmount").Value)
									        If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' CONCEPTO1
										        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(5990 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsD + lConceptsPositionDisplacement + y1PositionDisplacement-50) & " \absw737{\*\atnid CAVE.CPTO.} {\*\atnid CONCEPTO1}", sErrorDescription)
									        Else
										        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(5990+ x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsD + 200 + lConceptsPositionDisplacementR + y1PositionDisplacement-50) & " \absw737{\*\atnid CAVE.CPTO.} {\*\atnid CONCEPTO1}", sErrorDescription)
									        End If
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs14", sErrorDescription)
					    			        	lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CStr(oRecordset.Fields("ConceptShortName").Value), sErrorDescription)
									        lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
										If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' DESCRIPCION
											lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(6800 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsD + lConceptsPositionDisplacement + y1PositionDisplacement-50) & " \absw3515{\*\atnid DESCRIPCION}", sErrorDescription)
										Else
											lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(6800 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsD + 200 + lConceptsPositionDisplacementR + y1PositionDisplacement-50) & " \absw3515{\*\atnid DESCRIPCION}", sErrorDescription)
										End If
										lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs14", sErrorDescription)
										lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), CStr(oRecordset.Fields("ConceptName").Value), sErrorDescription)
										lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)

										If InStr(1, S_CREDITS_ID, "," & CInt(oRecordset.Fields("ConceptID").Value) & ",", vbBinaryCompare) > 0 Then
											If CDbl(oRecordset.Fields("ConceptRetention").Value) > 0 Then
												If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' CONTADOR
													lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(9000 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsD + lConceptsPositionDisplacement + y1PositionDisplacement-50) & " \absw1214{\*\atnid CONTADOR}", sErrorDescription)
												Else
													lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(9000 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsD + 200 + lConceptsPositionDisplacementR + y1PositionDisplacement-50) & " \absw1214{\*\atnid CONTADOR}", sErrorDescription)
												End If
												lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs14", sErrorDescription)
												lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), Replace(Left(CStr(oRecordset.Fields("ConceptRetention").Value), (Len(CStr(oRecordset.Fields("ConceptRetention").Value)) - Len("1"))), ".", "/"), sErrorDescription)
												lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
											End If
										End If

										If StrComp(oRecordset.Fields("ToAccountNumber").Value, ".", vbBinaryCompare) = 0 Then ' IMPORTE
											lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10524 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsD + lConceptsPositionDisplacement + y1PositionDisplacement-50) & " \absw1214{\*\atnid IMPORTE}", sErrorDescription)
										Else
											lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "{\pard \pvpg\phpg \posx" & CInt(10524 + x1PositionDisplacement) & "\posy" & CInt(yPositionForConceptsD + 200 + lConceptsPositionDisplacementR + y1PositionDisplacement-50) & " \absw1214{\*\atnid IMPORTE}", sErrorDescription)
										End If
										lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\fs14", sErrorDescription)
										lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), FormatNumber(lConceptAmount, 2, True, False, True), sErrorDescription)
										lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "\par}", sErrorDescription)
										If Len(CStr(oRecordset.Fields("ConceptName").Value)) > 53 Then
											yPositionForConceptsD = yPositionForConceptsD + 300
										Else
											yPositionForConceptsD = yPositionForConceptsD + 150
										End If
								End If
						End Select
                                                   
						        If lMinDate > oRecordset.Fields("MinDate").Value Then lMinDate = oRecordset.Fields("MinDate").Value
						        If lMaxDate < oRecordset.Fields("MaxDate").Value Then lMaxDate = oRecordset.Fields("MaxDate").Value
						        oRecordset.MoveNext 
                                employeeNumber = ""                     
						        If (Err.number <> 0) Or (lErrorNumber <> 0) Then Exit Do                             
                            End If                         
                        Loop
                    End If

					lErrorNumber = AppendTextToFile(Replace(sDocumentName, "<INDEX />", lFileCounter), "}", sErrorDescription)
					lErrorNumber = ZipFile(sFilePath, Server.MapPath(sFileName), sErrorDescription)
					oRecordset.Close
					sErrorDescription = "No se pudieron obtener los empleados que cumplen con los criterios de la búsqueda."
					lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Update Payments Set StatusID=1, Description='" & Replace(oRequest("Description").Item, "'", "´") & "' Where (PaymentID>-1) " & sCondition, "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, Null)
					If lErrorNumber = 0 Then
						aCatalogComponent(AS_FIELDS_VALUES_CATALOG)(15) = 1
						lErrorNumber = ModifyCatalog(oRequest, oADODBConnection, aCatalogComponent, sErrorDescription)
					End If
					If lErrorNumber = 0 Then
						lErrorNumber = DeleteFolder(sFilePath, sErrorDescription)
					End If
					oEndDate = Now()
					If (lErrorNumber = 0) And B_USE_SMTP Then
						If DateDiff("n", oStartDate, oEndDate) > 5 Then lErrorNumber = SendReportAlert(sFileName, CLng(Left(sDate, (Len("00000000")))), sErrorDescription)
					End If
				End If
			Else
				lErrorNumber = -1
				sErrorDescription = "No existen empleados que cumplan con los criterios de la búsqueda."
			End If
		End If
	End If

	Set oRecordset = Nothing
	PrintPayments = lErrorNumber
	Err.Clear
End Function

Function ShowSignatures(oRequest, oADODBConnection, sErrorDescription)
'************************************************************
'Purpose: To show the signatures for check payments
'Inputs:  oRequest, oADODBConnection, aCatalogComponent
'Outputs: sErrorDescription
'************************************************************
	On Error Resume Next
	Const S_FUNCTION_NAME = "ShowSignatures"
	Dim oRecordset
	Dim lErrorNumber

	sErrorDescription = "No se pudieron obtener los registros del catálogo."
	lErrorNumber = ExecuteSQLQuery(oADODBConnection, "Select * From Areas Where (ParentID=-1) And (AreaID>0) Order By AreaShortName", "PaymentsLib.asp", S_FUNCTION_NAME, 000, sErrorDescription, oRecordset)
	If lErrorNumber = 0 Then
		Response.Write "<!-- FIRMAS -->" & vbNewLine
		Response.Write "<TABLE BORDER=""1"" CELLPADDING=""0"" CELLSPACING=""0"">" & vbNewLine
			Do While Not oRecordset.EOF
				Response.Write "<TR><TD COLSPAN=""2""><FONT FACE=""Arial"" SIZE=""2""><B>" & CleanStringForHTML(CStr(oRecordset.Fields("AreaShortName").Value) & ". " & oRecordset.Fields("AreaName").Value) & "</B></FONT></TD></TR>" & vbNewLine
				Response.Write "<TR>"
					Response.Write "<TD ALIGN=""CENTER""><IMG SRC=""Templates/Images/" & Right(("00" & CStr(oRecordset.Fields("AreaID").Value)), Len("00")) & "01.jpg"" WIDTH=""135"" HEIGHT=""51"" /></TD>"
					Response.Write "<TD ALIGN=""CENTER""><IMG SRC=""Templates/Images/" & Right(("00" & CStr(oRecordset.Fields("AreaID").Value)), Len("00")) & "02.jpg"" WIDTH=""135"" HEIGHT=""51"" /></TD>"
				Response.Write "</TR>" & vbNewLine
				Response.Write "<TR>"
					Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">IZQUIERDA</FONT></TD>"
					Response.Write "<TD ALIGN=""CENTER""><FONT FACE=""Arial"" SIZE=""2"">DERECHA</FONT></TD>"
				Response.Write "</TR>" & vbNewLine
				oRecordset.MoveNext
				If Err.number <> 0 Then Exit Do
			Loop
		Response.Write "</TABLE>" & vbNewLine
		Response.Write "<!-- FIRMAS -->" & vbNewLine
	End If

	Set oRecordset = Nothing
	ShowSignatures = lErrorNumber
	Err.Clear
End Function
%>